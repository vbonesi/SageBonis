# -*- coding: utf-8 -*-
"""
sync_botoes.py - Garante os botoes da Trilha Completa DENTRO da planilha.

O .ods ja tinha menu e barra "SageBonis" (ver sync_menu.py), mas na propria
planilha so existiam 3 botoes, todos da Trilha Simples: "Importar arquivos
.dats" e "Exportar arquivos .dats" na aba Geral, e "Verificar Cores" na aba
Cores. As 7 macros da Completa (e as duas parciais da Simples) so eram
acessiveis por um menu que o usuario precisa saber que existe.

Esta ferramenta injeta:
  - um PAINEL na aba Geral (coluna H), com um botao por macro; e
  - um botao CONTEXTUAL na aba de config de cada recurso (IEDs -> Gerar IED,
    TrocaId -> Trocar ID, ...), ancorado logo a direita do cabecalho -- onde o
    usuario acabou de digitar. Nao da pra coloca-lo acima da linha 1: as macros
    leem o cabecalho de A1, entao a linha 1 tem que continuar sendo cabecalho.

Ao contrario do sync_macro.py/sync_menu.py, esta precisa do LibreOffice: botao
de formulario nao vive num XML separado como o menu, e sim dentro do
content.xml, no meio da estrutura da aba (forms + draw:control ancorado a
celula). Montar isso na mao seria fragil; aqui o proprio LibreOffice escreve.

E idempotente: todo botao que ela cria leva o prefixo "btnSB_" no nome, e cada
execucao remove os antigos com esse prefixo antes de recriar. Os 3 botoes
originais (Importar/Exportar/Cores) nao tem o prefixo e nunca sao tocados.

Uso:
    python sync_botoes.py status --ods completa/SageBonis.ods   # mostra o que falta
    python sync_botoes.py sync   --ods completa/SageBonis.ods   # cria/atualiza

Opcionais:
    --porta N       porta UNO (padrao 2077)
    --no-backup     nao cria .bak antes de gravar (sync)
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time

PREFIXO = "btnSB_"

# --- Painel da aba Geral -------------------------------------------------
NOME_ABA_PAINEL = "Geral"
COLUNA_PAINEL = 7          # H -- livre: o conteudo da Geral vai ate a coluna F
LINHA_TITULO_COMPLETA = 0  # H1
LINHA_PRIMEIRO_BOTAO = 1   # H2, um botao a cada 2 linhas
PASSO_LINHA = 2
LARGURA_BOTAO = 5000       # 1/100 mm  (5 cm)
ALTURA_BOTAO = 900         # 0,9 cm
# Rótulo do status das macros da Completa, logo abaixo do painel. A célula do valor
# (H23) é a CELULA_STATUS_COMPLETA do ImportadorSAGE.py -- se mudar lá, muda aqui.
LINHA_ROTULO_STATUS = 21   # H22

# (funcao, rotulo) -- mesma ordem e rotulos do menu SageBonis (sync_menu.py)
ITENS_COMPLETA = [
    ("verificar_base", "Verificar Base"),
    ("unificar_pontos", "Unificar Pontos"),
    ("extrair_pontos", "Extrair Pontos"),
    ("gerar_ied", "Gerar IED"),
    ("trocar_id_global", "Trocar ID"),
    ("estatistica_base", "Estatística"),
    ("gerir_includes", "Gerir Includes"),
]
# As parciais da Simples tambem so existiam no menu, apesar de a lista de
# entidades que elas usam (coluna C) estar bem ali na Geral.
ITENS_PARCIAIS = [
    ("importar_parcial", "Importar Parcial"),
    ("exportar_parcial", "Exportar Parcial"),
]

# --- Botoes contextuais (aba de config -> macro que a consome) -----------
BOTOES_CONTEXTUAIS = [
    ("VerificacaoRefs", "verificar_base", "Verificar Base"),
    ("PontoDigital", "unificar_pontos", "Unificar Pontos"),
    ("PontoAnalogico", "unificar_pontos", "Unificar Pontos"),
    ("ComandoAvulso", "unificar_pontos", "Unificar Pontos"),
    ("CanaisDistribuicao", "unificar_pontos", "Unificar Pontos"),
    ("DistribuicaoPontos", "unificar_pontos", "Unificar Pontos"),
    ("IEDs", "gerar_ied", "Gerar IED"),
    ("TrocaId", "trocar_id_global", "Trocar ID"),
    ("SubstituirIncludes", "gerir_includes", "Gerir Includes"),
]
COLUNAS_APOS_CABECALHO = 2  # folga entre a ultima coluna de cabecalho e o botao


def _href(funcao):
    return ("vnd.sun.star.script:ImportadorSAGE.py$%s?language=Python&location=document"
            % funcao)


def _verificar_lock(ods_path):
    """Recusa operar se o LibreOffice tem o .ods aberto (arquivo de lock
    '.~lock.<nome>#' ao lado dele) -- mesma protecao do sync_macro.py: salvar por
    cima da sessao aberta perderia a alteracao em silencio."""
    pasta, nome = os.path.split(os.path.abspath(ods_path))
    lock_path = os.path.join(pasta, f".~lock.{nome}#")
    if os.path.exists(lock_path):
        sys.exit(
            f"ERRO: {ods_path} parece estar aberto no LibreOffice ({lock_path} existe). "
            "Feche o documento antes de continuar."
        )


class SessaoLibreOffice:
    """Sobe um soffice headless proprio (profile isolado), abre o .ods indicado e
    fecha tudo ao sair. Escreve no arquivo de verdade -- nao numa copia --, por
    isso o cuidado com o lock antes de comecar."""

    def __init__(self, ods_path, porta=2077, timeout_boot=20, timeout_conexao=40):
        self.ods_path = os.path.abspath(ods_path)
        self.porta = porta
        self.timeout_boot = timeout_boot
        self.timeout_conexao = timeout_conexao
        self.profile_dir = None
        self.processo = None
        self.doc = None

    def __enter__(self):
        self.profile_dir = tempfile.mkdtemp(prefix="sagebonis_botoes_profile_")
        self.processo = subprocess.Popen(
            ["soffice", "--headless", "--norestore", "--nologo",
             "--accept=socket,host=localhost,port=%d;urp;" % self.porta,
             "-env:UserInstallation=file://%s" % self.profile_dir],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        time.sleep(self.timeout_boot)
        self._conectar()
        from com.sun.star.beans import PropertyValue
        p = PropertyValue()
        p.Name = "Hidden"
        p.Value = True
        self.doc = self.desktop.loadComponentFromURL(
            "file://" + self.ods_path, "_blank", 0, (p,))
        if self.doc is None:
            raise RuntimeError("Não foi possível abrir %s no LibreOffice." % self.ods_path)
        return self

    def _conectar(self):
        import uno
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        inicio = time.time()
        ultimo_erro = None
        while time.time() - inicio < self.timeout_conexao:
            try:
                self.ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext"
                    % self.porta)
                self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.ctx)
                return
            except Exception as e:  # noqa: PERF203 -- retry ate o soffice subir
                ultimo_erro = e
                time.sleep(0.5)
        raise RuntimeError("Não conectou ao soffice na porta %d: %s" % (self.porta, ultimo_erro))

    def salvar(self):
        self.doc.store()

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.doc is not None:
                self.doc.close(False)
        except Exception:
            pass
        try:
            if self.processo is not None:
                self.processo.terminate()
                self.processo.wait(timeout=20)
        except Exception:
            pass
        if self.profile_dir:
            shutil.rmtree(self.profile_dir, ignore_errors=True)
        return False


def _ultima_coluna_cabecalho(sheet):
    """Índice da última coluna preenchida na linha 1 da aba (-1 se vazia)."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    fim = cursor.getRangeAddress().EndColumn
    if fim < 0:
        return -1
    linha = sheet.getCellRangeByPosition(0, 0, fim, 0).getDataArray()[0]
    ultima = -1
    for i, valor in enumerate(linha):
        if str(valor).strip():
            ultima = i
    return ultima


def _shapes_gerados(sheet):
    """Shapes de botão criados por esta ferramenta (prefixo no nome do controle)."""
    dp = sheet.DrawPage
    encontrados = []
    for i in range(dp.Count):
        shape = dp.getByIndex(i)
        try:
            nome = shape.Control.Name
        except Exception:
            continue
        if nome.startswith(PREFIXO):
            encontrados.append((nome, shape))
    return encontrados


def _remover_gerados(sheet):
    dp = sheet.DrawPage
    removidos = 0
    for _nome, shape in _shapes_gerados(sheet):
        dp.remove(shape)
        removidos += 1
    return removidos


def _inserir_botao(doc, sheet, funcao, rotulo, coluna, linha, sufixo=""):
    """Cria o botão ancorado na célula (coluna, linha) e liga no script da macro."""
    from com.sun.star.awt import Size, Point
    from com.sun.star.script import ScriptEventDescriptor

    nome = PREFIXO + funcao + sufixo
    modelo = doc.createInstance("com.sun.star.form.component.CommandButton")
    modelo.Name = nome
    modelo.Label = rotulo
    # Botão de formulário some do documento se ficar "focável" em modo de edição;
    # o padrão já serve, só garantimos que não vira submit de formulário.
    modelo.ButtonType = 0  # com.sun.star.form.FormButtonType.PUSH

    forma = doc.createInstance("com.sun.star.drawing.ControlShape")
    forma.setSize(Size(LARGURA_BOTAO, ALTURA_BOTAO))
    forma.setPosition(Point(0, 0))
    forma.Control = modelo
    sheet.DrawPage.add(forma)
    forma.Anchor = sheet.getCellByPosition(coluna, linha)

    form, idx = _localizar_no_formulario(sheet, nome)
    ev = ScriptEventDescriptor()
    ev.ListenerType = "XActionListener"
    ev.EventMethod = "actionPerformed"
    ev.ScriptType = "Script"
    ev.ScriptCode = _href(funcao)
    form.registerScriptEvent(idx, ev)
    return nome


def _localizar_no_formulario(sheet, nome_controle):
    """(formulário, índice) do controle recém-adicionado -- é onde o evento é
    registrado (XEventAttacherManager é do formulário, não do shape)."""
    forms = sheet.DrawPage.Forms
    for f in range(forms.Count):
        form = forms.getByIndex(f)
        for i in range(form.Count):
            if form.getByIndex(i).Name == nome_controle:
                return form, i
    raise RuntimeError("Controle %s não encontrado em nenhum formulário da aba." % nome_controle)


def _aplicar_painel(doc):
    """Painel da aba Geral: título + botões da Completa + botões das parciais."""
    sheet = doc.Sheets.getByName(NOME_ABA_PAINEL)
    criados = []
    linha = LINHA_PRIMEIRO_BOTAO
    sheet.getCellByPosition(COLUNA_PAINEL, LINHA_TITULO_COMPLETA).setString("Trilha Completa")
    for funcao, rotulo in ITENS_COMPLETA:
        criados.append(_inserir_botao(doc, sheet, funcao, rotulo, COLUNA_PAINEL, linha))
        linha += PASSO_LINHA
    linha += 1
    sheet.getCellByPosition(COLUNA_PAINEL, linha).setString("Parcial (aba ativa ou lista)")
    linha += 1
    for funcao, rotulo in ITENS_PARCIAIS:
        criados.append(_inserir_botao(doc, sheet, funcao, rotulo, COLUNA_PAINEL, linha))
        linha += PASSO_LINHA
    sheet.getCellByPosition(COLUNA_PAINEL, LINHA_ROTULO_STATUS).setString(
        "Status da Trilha Completa")
    return criados


def _aplicar_contextuais(doc):
    """Um botão na aba de config de cada recurso, logo à direita do cabeçalho."""
    criados = []
    for nome_aba, funcao, rotulo in BOTOES_CONTEXTUAIS:
        if not doc.Sheets.hasByName(nome_aba):
            continue  # aba ainda não criada nesta planilha -- nada a fazer
        sheet = doc.Sheets.getByName(nome_aba)
        coluna = _ultima_coluna_cabecalho(sheet) + 1 + COLUNAS_APOS_CABECALHO
        criados.append(_inserir_botao(doc, sheet, funcao, rotulo, coluna, 0,
                                      sufixo="_" + nome_aba.lower()))
    return criados


def _abas_com_botao(doc):
    """{aba: [nomes dos botões gerados]} -- usado pelo 'status'."""
    saida = {}
    for i in range(doc.Sheets.Count):
        sheet = doc.Sheets.getByIndex(i)
        nomes = [nome for nome, _shape in _shapes_gerados(sheet)]
        if nomes:
            saida[sheet.Name] = sorted(nomes)
    return saida


def _esperados(abas_existentes):
    """Nomes que o sync criaria nesta planilha. Botão contextual de aba que ainda
    não existe não conta como faltando -- a aba nasce quando a macro roda."""
    nomes = [PREFIXO + f for f, _ in ITENS_COMPLETA + ITENS_PARCIAIS]
    nomes += [PREFIXO + f + "_" + aba.lower()
              for aba, f, _ in BOTOES_CONTEXTUAIS if aba in abas_existentes]
    return sorted(nomes)


def acao_status(ods_path, porta):
    with SessaoLibreOffice(ods_path, porta=porta) as sessao:
        atuais = _abas_com_botao(sessao.doc)
        abas_existentes = {sessao.doc.Sheets.getByIndex(i).Name
                           for i in range(sessao.doc.Sheets.Count)}
    todos = sorted(n for nomes in atuais.values() for n in nomes)
    esperados = _esperados(abas_existentes)
    faltando = [n for n in esperados if n not in todos]
    sobrando = [n for n in todos if n not in esperados]
    for aba, nomes in sorted(atuais.items()):
        print(f"{aba}: {', '.join(nomes)}")
    if not atuais:
        print("Nenhum botão gerado por esta ferramenta ainda.")
    if faltando:
        print("FALTAM: %s" % ", ".join(faltando))
    if sobrando:
        print("SOBRANDO (serão recriados no sync): %s" % ", ".join(sobrando))
    if not faltando and not sobrando:
        print("OK: todos os botões esperados já estão na planilha.")


def acao_sync(ods_path, porta, backup=True):
    if backup:
        destino = ods_path + ".bak"
        shutil.copy2(ods_path, destino)
        print("Backup criado: %s" % destino)
    with SessaoLibreOffice(ods_path, porta=porta) as sessao:
        doc = sessao.doc
        removidos = sum(_remover_gerados(doc.Sheets.getByIndex(i))
                        for i in range(doc.Sheets.Count))
        criados = _aplicar_painel(doc) + _aplicar_contextuais(doc)
        sessao.salvar()
    print("Botões antigos removidos: %d" % removidos)
    print("Botões criados: %d (%s)" % (len(criados), ", ".join(criados)))
    print("OK: %s atualizado." % ods_path)
    print("Obs.: o .ods foi reescrito pelo LibreOffice -- rode "
          "'python sync_macro.py status' para conferir a macro embutida.")


def main():
    p = argparse.ArgumentParser(
        description="Cria/atualiza os botões das macros dentro do SageBonis.ods")
    p.add_argument("acao", choices=["status", "sync"])
    p.add_argument("--ods", default="completa/SageBonis.ods")
    p.add_argument("--porta", type=int, default=2077)
    p.add_argument("--no-backup", action="store_true")
    args = p.parse_args()

    if not os.path.exists(args.ods):
        sys.exit("ERRO: %s não encontrado." % args.ods)
    _verificar_lock(args.ods)

    if args.acao == "status":
        acao_status(args.ods, args.porta)
    else:
        acao_sync(args.ods, args.porta, backup=not args.no_backup)


if __name__ == "__main__":
    main()
