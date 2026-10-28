# -*- coding: utf-8 -*-

import os
import re
import time
import traceback

# ===============================================================
# ================ MACRO SAGE - VERSÃO 0.9.3 ====================
# ===============================================================
# Este script é utilizado como macro no LibreOffice Calc para importar e exportar arquivos .dat
# do Sistema Aberto de Gerenciamento de Energia (SAGE).
#
# A partir desta versão, o script lê dinamicamente as configurações de
# ordenação, cores e validação das abas "MaisUsadas" e "EntidadesValoresAtributos".
#
# Adicionado a aba cores para facilitar a aplicação de temas de cores.
# Correção de importação de comentarios
#
# Alterações da versão 0.9.2 por Felipe Santos:
# - Compatibilidade com LibreOffice Flatpak usando macro embutida no documento.
# - Correção do nome da aba Geral e busca de abas tolerante a maiúsculas/minúsculas.
# - Tratamento mais seguro para erros de configuração.
# - A planilha SageBonis.ods agora pode carregar este script a partir do próprio documento,
#   usando URLs de macro com location=document.
# - Foram adicionados menu e barra de ferramentas SageBonis dentro do arquivo ODS.
#
# Alterações da versão 0.9.3:
# - Reincorporada a sanitização para latin-1 (_sanitizar_para_latin1) na exportação,
#   evitando UnicodeEncodeError quando há caracteres Unicode (aspas curvas, travessões,
#   etc.) que não existem no ISO-8859-1 esperado pelo SAGE.
# - Resultado da reconciliação entre a macro embutida no .ods (0.9.2) e o histórico do repo.
#
# Desenvolvido para rodar com a planilha SageBonis.ods
# Duvidas/Bugs/Sugestões - (11) 95456-4510 - Victor Bonesi - https://github.com/vbonesi/SageBonis

# ===============================================================
# ==================== CONFIGURAÇÃO GERAL =======================
# ===============================================================

# --- Nomes de Abas ---
# Importante: o LibreOffice diferencia maiúsculas/minúsculas em getByName().
# A planilha possui a aba "Geral" com G maiúsculo; usar "geral" causava falha
# nas rotinas de importação/exportação antes mesmo de ler os caminhos.
NOME_ABA_GERAL = "Geral"
NOME_ABA_MAIS_USADAS = "MaisUsadas"
NOME_ABA_VALIDACAO = "EntidadeAtributoValor"
NOME_ABA_OPMSK = "opmsk"
NOME_ABA_CORES = "Cores"

# --- Lista de Abas a Ignorar ---
FOLHAS_IGNORADAS = [NOME_ABA_GERAL, NOME_ABA_MAIS_USADAS, NOME_ABA_VALIDACAO, NOME_ABA_OPMSK, NOME_ABA_CORES]

# --- Posições das Células na Aba "Geral" ---
CELULA_CAMINHO_IMPORTACAO = (0, 3)  # A4
CELULA_STATUS_IMPORTACAO = (1, 3)   # B4
CELULA_CAMINHO_EXPORTACAO = (0, 6)  # A7
CELULA_STATUS_EXPORTACAO = (1, 6)   # B7
# Lista de entidades da importação/exportação parcial na aba Geral: coluna C, a partir
# da linha 14 (coordenadas UNO são base zero) e até o fim da área usada da aba. Antes
# era um range fixo de 130 linhas (C14:C144): limite arbitrário e silencioso -- a
# entidade listada depois dele simplesmente não era processada (auditoria #12).
COLUNA_ENTIDADES_PARCIAL = 2
PRIMEIRA_LINHA_ENTIDADES_PARCIAL = 13

# --- Códigos de Controle (Coluna "Gera") ---
CODIGO_BLOCO_ATIVO = 'x'
CODIGO_BLOCO_COMENTADO = 'c'
CODIGO_COMENTARIO_SIMPLES = 'n'
CODIGO_IGNORAR_LINHA = 'q'
CODIGO_INCLUDE = 'i'
CODIGO_INCLUDE_COMENTADO = 'u'

# --- Cabeçalhos Padrão ---
CABEÇALHO_COLUNA_ORIGEM = "Origem"
CABEÇALHO_COLUNA_CONTROLE = "Gera"
CABEÇALHO_COLUNA_DADOS = "Comentario/Include"

# --- NOVO: Cores para Linhas Alternadas (Efeito Zebra) ---
# Cores em formato numérico (Decimal de Hex BGR: Blue-Green-Red)
COR_LINHA_PAR = 16777215   # Branco (0xFFFFFF)
COR_LINHA_IMPAR = 15790320  # Cinza muito claro (0xF0F0F0)

# --- Constantes Técnicas ---
FLAGS_LIMPAR_TUDO = 1048575
LIMITE_CARACTERES_VALIDACAO = 250 # Manteremos para a próxima etapa

# --- Codificação dos Arquivos DAT do SAGE ---
ENCODING_EXPORTACAO_SAGE = 'latin-1'  # ISO-8859-1 (padrão esperado pelo SAGE)
ENCODINGS_IMPORTACAO_SAGE = ('utf-8', 'latin-1')  # utf-8 primeiro: latin-1 decodifica
# qualquer sequência de bytes sem erro, então na ordem antiga um .dat em utf-8 nunca
# caía no fallback e importava como mojibake (achado de auditoria).

# Mapeamento de caracteres Unicode comuns que não existem no ISO-8859-1
_UNICODE_PARA_LATIN1 = str.maketrans({
    '–': '-',    # en dash
    '—': '-',    # em dash
    '‘': "'",    # aspas simples esquerda
    '’': "'",    # aspas simples direita
    '“': '"',    # aspas duplas esquerda
    '”': '"',    # aspas duplas direita
    '…': '...',  # reticências
    '•': '-',    # bullet
})

def _sanitizar_para_latin1(texto):
    """Retorna (texto_sanitizado, n_substituicoes). n_substituicoes conta caracteres
    sem equivalente em latin-1 (fora do mapeamento acima) que o encode(errors=
    'replace') vira '?' silenciosamente -- todo char com ord()>255 não cabe em
    latin-1, que mapeia só 0-255 direto pra byte."""
    texto = texto.translate(_UNICODE_PARA_LATIN1)
    n_substituicoes = sum(1 for c in texto if ord(c) > 255)
    return texto.encode('latin-1', errors='replace').decode('latin-1'), n_substituicoes


# --- Expressões Regulares ---
REGEX_INCLUDE = re.compile(r'^\s*#\s*include\s+(.*)', re.IGNORECASE)
REGEX_INCLUDE_COMENTADO = re.compile(r'^\s*;\s*#\s*include\s+(.*)', re.IGNORECASE)
REGEX_INICIO_BLOCO_COMENTADO = re.compile(r'^\s*;\s*([A-Z0-9_]+)\s*$', re.IGNORECASE)

# --- Debug/Diagnóstico de Importação ---
DEBUG_IMPORTACAO = False
LOG_IMPORTACAO_RESUMO = True
LOG_IMPORTACAO_AVISOS = True
WATCHDOG_MAX_ITERACOES_SEM_PROGRESSO = 1000

# O que o parser teve que preservar de um jeito diferente do original durante UMA
# importação (zerado em _executar_importacao, lido pelas macros pra compor o status):
# linhas fora do formato mantidas como comentário, atributos repetidos num bloco e
# arquivos vazios marcados. Antes isso eram avisos no console, que ninguém lê.
ACUMULADO_IMPORTACAO = {'preservadas': 0, 'repetidos': 0, 'arquivo_vazio': 0,
                        'abas_esvaziadas': [], 'ignoradas': []}


def _indice_de_abas(doc):
    """{nome_em_minusculas: nome_real} de todas as abas do documento."""
    sheets = doc.getSheets()
    return {sheets.getByIndex(i).getName().lower(): sheets.getByIndex(i).getName()
            for i in range(sheets.getCount())}


def _get_sheet(doc, sheet_name):
    """Aba pelo nome, tolerante a maiúsculas/minúsculas.

    O LibreOffice trata nome de aba como case-sensitive em getByName(), mas não deixa
    duas abas diferirem só na caixa -- então uma busca por nome normalizado é sempre
    não-ambígua. Sem isso, "Geral" x "geral" derruba a macro inteira.

    Monta o índice normalizado de uma vez em vez de varrer o container a cada
    tentativa (reescrita alinhada com o núcleo compartilhado import/export).
    """
    sheets = doc.getSheets()
    if sheets.hasByName(sheet_name):
        return sheets.getByName(sheet_name)
    real = _indice_de_abas(doc).get(sheet_name.strip().lower())
    if real is None:
        raise KeyError("Aba '%s' não encontrada." % sheet_name)
    return sheets.getByName(real)


def _aba_existe_ci(doc, nome):
    """True se existe uma aba com esse nome (ignorando maiúsc/minúsc)."""
    return (doc.getSheets().hasByName(nome)
            or nome.strip().lower() in _indice_de_abas(doc))


# Arquivo de dados do SAGE: 'x.dat' e as variantes 'x.dat.novo', 'pdd.dat.rgv'...
_RE_ARQUIVO_DAT = re.compile(r'^(.+?)\.dat(?:\.[^.]+)?$', re.IGNORECASE)


def _chave_do_arquivo(nome):
    """Entidade que dá nome ao arquivo (minúscula): 'pds.dat' -> 'pds',
    'pdd.dat.rgv' -> 'pdd'. É onde caem o cabeçalho e os #include que abrem o arquivo
    antes do primeiro bloco."""
    base = os.path.basename(nome)
    m = _RE_ARQUIVO_DAT.match(base)
    chave = (m.group(1) if m else os.path.splitext(base)[0]).strip().lower()
    # É nome de aba: sem os caracteres que o LibreOffice proíbe e sem espaço nas pontas.
    # O espaço importa: existem bases com 'grcmp .dat' (espaço no nome), e _get_sheet ignora
    # espaço nas pontas -- a "aba" 'grcmp ' caía em cima da 'grcmp' e a SOBRESCREVIA
    # (todos os grcmp.dat da base sumiam na exportação).
    return re.sub(r"[\[\]*?:/\\]", "_", chave).strip("'") or "_"


def _log_importacao(level, message, force=False):
    """Logger simples e opcional para diagnóstico da importação."""
    if force:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level == 'DEBUG' and DEBUG_IMPORTACAO:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level == 'INFO' and LOG_IMPORTACAO_RESUMO:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level in ['WARN', 'ERROR'] and LOG_IMPORTACAO_AVISOS:
        print(f"[IMPORTACAO:{level}] {message}")


# Só espaço "de verdade" é borda de valor: str.strip() sem argumento também tira o espaço
# não-quebrável (U+00A0), que existe em Latin-1 e que o carregador do SAGE MANTÉM no valor
# ('...1.Grau\u00a0' aparecia assim na listagem do carregador e a importação o comia).
_ESPACOS_ASCII = ' \t\r\n\f\v'


def _classificar_linha_dat(raw_line, entidades_validas):
    """Classifica a linha do arquivo DAT para manter o parser determinístico."""
    original_line = raw_line.strip('\r\n')
    stripped_line = original_line.strip(_ESPACOS_ASCII)

    if not stripped_line:
        return {'type': 'blank', 'original': original_line, 'stripped': stripped_line}

    include_comentado_match = REGEX_INCLUDE_COMENTADO.match(original_line)
    if include_comentado_match:
        return {
            'type': 'include_commented',
            'original': original_line,
            'stripped': stripped_line,
            'value': include_comentado_match.group(1).strip(_ESPACOS_ASCII)
        }

    include_match = REGEX_INCLUDE.match(original_line)
    if include_match:
        return {
            'type': 'include',
            'original': original_line,
            'stripped': stripped_line,
            'value': include_match.group(1).strip(_ESPACOS_ASCII)
        }

    if stripped_line.upper() in entidades_validas:
        return {
            'type': 'entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': stripped_line.upper()
        }

    commented_block_match = REGEX_INICIO_BLOCO_COMENTADO.match(original_line)
    if commented_block_match and commented_block_match.group(1).upper() in entidades_validas:
        return {
            'type': 'commented_entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': commented_block_match.group(1).upper()
        }

    if stripped_line.startswith(';'):
        return {
            'type': 'comment',
            'original': original_line,
            'stripped': stripped_line,
            'value': original_line.lstrip(';').lstrip()
        }

    if '=' in stripped_line:
        key, value = stripped_line.split('=', 1)
        return {
            'type': 'attribute',
            'original': original_line,
            'stripped': stripped_line,
            'key': key.strip(_ESPACOS_ASCII),
            'value': value.strip(_ESPACOS_ASCII)
        }

    return {'type': 'invalid', 'original': original_line, 'stripped': stripped_line}


def _classificar_linha_bloco_comentado(raw_line, entidades_validas):
    """
    Classifica linhas quando estamos dentro de um bloco comentado.
    O conteúdo útil está sempre após o primeiro ';'.
    """
    original_line = raw_line.strip('\r\n')
    stripped_line = original_line.strip(_ESPACOS_ASCII)

    if not stripped_line:
        return {'type': 'blank', 'original': original_line, 'stripped': stripped_line}

    if not stripped_line.startswith(';'):
        return {'type': 'block_end', 'original': original_line, 'stripped': stripped_line}

    inner_line = stripped_line[1:].strip(_ESPACOS_ASCII)
    if not inner_line:
        return {'type': 'comment', 'original': original_line, 'stripped': stripped_line, 'value': ''}

    if inner_line.upper() in entidades_validas:
        return {
            'type': 'commented_entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': inner_line.upper()
        }

    if REGEX_INCLUDE_COMENTADO.match(stripped_line):
        return {'type': 'block_end', 'original': original_line, 'stripped': stripped_line}

    if '=' in inner_line:
        key, value = inner_line.split('=', 1)
        return {
            'type': 'attribute',
            'original': original_line,
            'stripped': stripped_line,
            'key': key.strip(_ESPACOS_ASCII),
            'value': value.strip(_ESPACOS_ASCII)
        }

    return {'type': 'comment', 'original': original_line, 'stripped': stripped_line, 'value': inner_line}


def _proximo_estrutural(lines, i, entidades_validas):
    """Índice e tipo da primeira linha depois de `i` que decide o destino de um registro
    ativo: um atributo vivo, um novo nome de entidade, um #include ou uma linha fora do
    formato. Comentários (inclusive `;ENTIDADE` e `;#include`) e linhas em branco não
    contam. Devolve (None, None) se o arquivo acaba antes."""
    for j in range(i + 1, len(lines)):
        tipo = _classificar_linha_dat(lines[j], entidades_validas)['type']
        if tipo in ('blank', 'comment', 'commented_entity_start', 'include_commented'):
            continue
        return j, tipo
    return None, None


def _iniciar_bloco(entidade_nome, tipo_bloco, relative_path, comentarios_iniciais=None):
    bloco = {
        'type': tipo_bloco,
        'identifier': entidade_nome,
        'attributes': {},
        'comments': [],
        'origem': relative_path
    }
    if comentarios_iniciais:
        bloco['comments'].extend(comentarios_iniciais)
    return bloco


def _finalizar_bloco(current_block, all_data, relative_path, stats, line_no):
    if not current_block:
        return

    ponto = {
        'type': current_block['type'],
        'identifier': current_block['identifier'],
        'attributes': current_block['attributes'],
        'origem': relative_path
    }

    if current_block['comments']:
        ponto['comment'] = "\n".join(current_block['comments'])

    if not current_block['attributes']:
        stats['warnings'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} bloco {current_block['identifier']} finalizado sem atributos.",
            force=True
        )

    if current_block['attributes'] and 'ID' not in current_block['attributes']:
        stats['warnings'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} bloco {current_block['identifier']} finalizado sem ID.",
            force=True
        )

    if current_block['attributes'] or current_block['comments']:
        chave = current_block['identifier'].lower()
        all_data.setdefault(chave, []).append(ponto)
        stats['entities_imported'] += 1


def _descarregar_comentarios_pendentes(pending_comments, all_data, chave_entidade, relative_path):
    """Grava comentários soltos (fora de bloco) como linhas 'n' na entidade corrente,
    na ordem em que apareceram, e devolve a lista já vazia.

    Antes eles eram descartados: comentário antes de um #include (cabeçalho de arquivo,
    comum nas bases reais) ou no fim do arquivo sumia na importação e não voltava na
    exportação -- round-trip destrutivo silencioso (auditoria #4). O tipo 'n' já era
    suportado de ponta a ponta (write_to_sheet e exportação), só não era produzido pela
    importação. Um ponto por linha, porque a exportação de 'n' prefixa um único ';'.
    """
    for texto in pending_comments:
        all_data.setdefault(chave_entidade, []).append({
            'type': CODIGO_COMENTARIO_SIMPLES,
            'data': texto,
            'origem': relative_path
        })
    return []

# ===============================================================
# =================== CLASSE DE CONFIGURAÇÃO ====================
# ===============================================================

class SageConfig:
    """Carrega e armazena todas as configurações das abas auxiliares."""
    def __init__(self, doc):
        self.doc = doc
        self.ordem_entidades = []
        self.cores_entidades = {}
        self.ordem_atributos = {}
        self.regras_validacao = {} # Mantido por segurança, mas não será preenchido
        
        self._carregar_configuracoes()

    def _carregar_configuracoes(self):
        """Método principal para chamar os carregadores."""
        self._carregar_mais_usadas()
        # A LINHA ABAIXO FOI REMOVIDA:
        # self._carregar_validacao()

    def _carregar_mais_usadas(self):
        """Lê a aba 'MaisUsadas' para obter ordem, cores e atributos prioritários."""
        try:
            sheet = _get_sheet(self.doc, NOME_ABA_MAIS_USADAS)
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            data_range = cursor.getRangeAddress()
            data = sheet.getCellRangeByPosition(0, 0, data_range.EndColumn, data_range.EndRow).getDataArray()

            if not data or len(data) < 2: return

            for row_idx, row_data in enumerate(data[1:], 1):
                if not row_data or not row_data[0]: continue
                entidade_nome = str(row_data[0]).lower().strip()
                if not entidade_nome: continue

                self.ordem_entidades.append(entidade_nome)
                cell = sheet.getCellByPosition(0, row_idx)
                self.cores_entidades[entidade_nome] = cell.CellBackColor
                atributos = [str(attr).upper() for attr in row_data[1:] if attr]
                if atributos:
                    self.ordem_atributos[entidade_nome] = atributos
        except Exception as e:
            # Config é acessório: sem ela a importação ainda roda, só perde a ordem
            # das entidades, as cores por aba e a prioridade de colunas. Por isso
            # avisa e segue, em vez de derrubar a macro inteira.
            print("AVISO: a aba '%s' não pôde ser lida (%s: %s). A importação segue sem "
                  "ordem/cores configuradas." % (NOME_ABA_MAIS_USADAS, type(e).__name__, e))

    # A FUNÇÃO _carregar_validacao FOI COMPLETAMENTE REMOVIDA DESTA CLASSE

# ===============================================================
# ================= FUNÇÕES DE IMPORTAÇÃO =======================
# ===============================================================

def _reportar_falha_inesperada(geral_sheet, celula_status, acao, excecao, consequencia):
    """Escreve no status o erro que escapou de toda validação e imprime o traceback.

    Sem isso, uma exceção no meio da rotina deixava o status preso em "Processando..."
    e a planilha (ou a pasta de saída) parcialmente reescrita, sem nada indicando que
    o que está na tela não é confiável (auditoria #9)."""
    print(traceback.format_exc())
    mensagem = f"ERRO: falha inesperada durante a {acao} ({type(excecao).__name__}: {excecao}). {consequencia}"
    try:
        geral_sheet.getCellByPosition(*celula_status).setString(mensagem)
    except Exception:
        print(mensagem)


def _ler_entidades_parciais(geral_sheet):
    """Nomes (minúsculos) listados na coluna de entidades da aba Geral, da primeira
    linha da lista até o fim da área usada da aba. Linhas em branco são puladas, então
    um buraco no meio da lista não corta o resto."""
    cursor = geral_sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    ultima_linha = cursor.getRangeAddress().EndRow
    if ultima_linha < PRIMEIRA_LINHA_ENTIDADES_PARCIAL:
        return []
    dados = geral_sheet.getCellRangeByPosition(
        COLUNA_ENTIDADES_PARCIAL, PRIMEIRA_LINHA_ENTIDADES_PARCIAL,
        COLUNA_ENTIDADES_PARCIAL, ultima_linha).getDataArray()
    return [str(row[0]).strip().lower() for row in dados if row and str(row[0]).strip()]


def _sufixo_importacao():
    """Sufixo do status da importação com o que ela teve que tratar de um jeito diferente
    do original -- string vazia quando não há o que avisar. Nada disso é perda (a linha
    continua na planilha), mas é o tipo de coisa que o usuário precisa saber que existe."""
    partes = []
    n = ACUMULADO_IMPORTACAO['arquivo_vazio']
    if n:
        partes.append(f"{n} arquivo(s) vazio(s) marcado(s) com um comentário vazio para a "
                      "exportação recriá-lo (um #include aponta para ele)")
    n = ACUMULADO_IMPORTACAO['preservadas']
    if n:
        partes.append(f"{n} linha(s) fora do formato mantida(s) como comentário")
    n = ACUMULADO_IMPORTACAO['repetidos']
    if n:
        partes.append(f"{n} atributo(s) repetido(s) num bloco (vale o último; o anterior virou comentário)")
    ignoradas = ACUMULADO_IMPORTACAO['ignoradas']
    if ignoradas:
        partes.append(f"{len(ignoradas)} entidade(s) NÃO importada(s) porque o nome coincide com uma aba "
                      f"de configuração/relatório: {', '.join(ignoradas[:5])}")
    esvaziadas = ACUMULADO_IMPORTACAO['abas_esvaziadas']
    if esvaziadas:
        nomes = ", ".join(esvaziadas[:5]) + ("..." if len(esvaziadas) > 5 else "")
        partes.append(f"{len(esvaziadas)} aba(s) com dados de uma base anterior e que não existem "
                      f"nesta pasta foram esvaziadas: {nomes}")
    return f" ({'; '.join(partes)})" if partes else ""


def importar_dats(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # Mantém a variável inicializada para que o bloco except não tente escrever
    # status em uma aba que falhou ao ser localizada.
    geral_sheet = None
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
        return

    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Processando importação total...")
    try:
        _executar_importacao(doc, folder_path, lista_entidades=None, modo_importacao='REPLACE')
    except Exception as e:
        _reportar_falha_inesperada(
            geral_sheet, CELULA_STATUS_IMPORTACAO, "importação total", e,
            "A planilha pode ter ficado parcialmente atualizada -- feche sem salvar antes de tentar de novo.")
        return
    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(
        "Importação total concluída com sucesso!" + _sufixo_importacao())


def importar_parcial(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    # A aba Geral pode não ser encontrada se o arquivo for alterado manualmente;
    # por isso o tratamento de erro precisa funcionar mesmo sem geral_sheet.
    geral_sheet = None
    
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
        return

    entidades_a_importar = []
    modo = 'REPLACE' 
    if active_sheet_name.lower() == NOME_ABA_GERAL.lower():
        entidades_a_importar = _ler_entidades_parciais(geral_sheet)
        if not entidades_a_importar:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("AVISO: Nenhuma entidade listada para importação parcial.")
            return
    else:
        entidades_a_importar.append(active_sheet_name.lower())
        modo = 'UPDATE'

    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"Processando importação de: {', '.join(entidades_a_importar)}...")
    try:
        _executar_importacao(doc, folder_path, lista_entidades=entidades_a_importar, modo_importacao=modo)
    except Exception as e:
        _reportar_falha_inesperada(
            geral_sheet, CELULA_STATUS_IMPORTACAO, "importação parcial", e,
            "A planilha pode ter ficado parcialmente atualizada -- feche sem salvar antes de tentar de novo.")
        return
    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(
        "Importação parcial concluída com sucesso!" + _sufixo_importacao())


def _ler_base(base_folder_path, lista_entidades=None):
    """Lê todos os arquivos da base (lógica PURA, sem planilha) e devolve
    {entidade: [linhas]} -- é o que a importação escreve nas abas. Separada de
    _executar_importacao pra que os testes de round-trip exercitem a leitura real."""
    all_data = {}
    for chave_acumulada in ACUMULADO_IMPORTACAO:
        ACUMULADO_IMPORTACAO[chave_acumulada] = [] if isinstance(ACUMULADO_IMPORTACAO[chave_acumulada], list) else 0
    # Todo nome de entidade que existe na base, não só os da pasta do arquivo: o carregador
    # do SAGE reconhece uma entidade em qualquer arquivo, e um 'lsc.dat' numa pasta sem
    # 'cnf.dat' pode muito bem ter blocos CNF. Limitar à pasta fazia esses blocos virarem
    # atributos órfãos.
    entidades_validas_set = set()
    for _, _, files in os.walk(base_folder_path):
        entidades_validas_set.update(
            _chave_do_arquivo(f).upper() for f in files if _RE_ARQUIVO_DAT.match(f))

    arquivos_lidos = set()

    def ler_arquivo(full_path):
        relative_path = os.path.relpath(full_path, base_folder_path)
        arquivos_lidos.add(os.path.normcase(relative_path))
        parse_dat_file(full_path, relative_path, all_data, entidades_validas_set)

    for root, _, files in os.walk(base_folder_path):
        for file_name in files:
            if not file_name.lower().endswith('.dat'):
                continue
            if lista_entidades is not None and _chave_do_arquivo(file_name) not in lista_entidades:
                continue
            ler_arquivo(os.path.join(root, file_name))

    # Arquivos que um #include ATIVO aponta e que não terminam em '.dat' (ocr.dat.novo,
    # pdd.dat.rgv...) também são a base: o carregador os lê. Ficavam de fora da planilha e
    # portanto da exportação, e o #include passava a apontar para um arquivo que não
    # existe (medido no carregador do SAGE: a carga aborta). Os includes não aninham,
    # mas o laço trata o caso de um arquivo desses incluir outro.
    while True:
        novos = []
        for linhas in list(all_data.values()):
            for linha in linhas:
                if linha.get('type') != CODIGO_INCLUDE:
                    continue
                alvo = os.path.normpath(os.path.join(
                    base_folder_path, str(linha.get('data', '')).strip().strip('"\'').replace('\\', '/')))
                rel = os.path.normcase(os.path.relpath(alvo, base_folder_path))
                if (rel in arquivos_lidos or rel.startswith('..') or alvo in novos
                        or alvo.lower().endswith('.dat') or not os.path.isfile(alvo)):
                    continue
                if lista_entidades is not None and _chave_do_arquivo(alvo) not in lista_entidades:
                    continue
                novos.append(alvo)
        if not novos:
            break
        for alvo in novos:
            ler_arquivo(alvo)
    return all_data


def _esvaziar_abas_ausentes(doc, entidades_importadas):
    """Importação TOTAL: a planilha passa a refletir a pasta. Aba de entidade que sobrou de
    uma base anterior e que não existe nesta pasta é esvaziada (mantém o cabeçalho).

    Antes ela ficava com os dados velhos, e a exportação total os gravava na pasta de
    saída: arquivos de uma base dentro da outra, sem aviso nenhum. Só mexe em aba com o
    formato de entidade (Origem / Gera / Comentario-Include na linha 1)."""
    nao_entidade = _abas_nao_entidade()
    esvaziadas = []
    for sheet in list(doc.getSheets()):
        nome = sheet.getName()
        if nome.lower() in nao_entidade or nome.lower() in entidades_importadas:
            continue
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        if addr.EndRow < 1:
            continue
        cabecalho = [str(c) for c in sheet.getCellRangeByPosition(
            0, 0, addr.EndColumn, 0).getDataArray()[0]]
        if not {CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS} <= set(cabecalho):
            continue
        sheet.getCellRangeByPosition(0, 1, addr.EndColumn, addr.EndRow).clearContents(FLAGS_LIMPAR_TUDO)
        esvaziadas.append(nome)
    ACUMULADO_IMPORTACAO['abas_esvaziadas'] = esvaziadas


def _executar_importacao(doc, base_folder_path, lista_entidades, modo_importacao):
    """
    Função interna que executa a importação, agora usando as configurações carregadas.
    """
    # ALTERAÇÃO: Carrega as configurações da planilha
    config = SageConfig(doc)
    all_data = _ler_base(base_folder_path, lista_entidades)
    prioridade_entidades = {entidade: idx for idx, entidade in enumerate(config.ordem_entidades)}

    # ALTERAÇÃO: Ordena as entidades a serem escritas com base na configuração
    entidades_importadas = all_data.keys()
    abas_ordenadas = sorted(
        entidades_importadas,
        key=lambda e: prioridade_entidades.get(e, float('inf'))
    )

    # Lógica de escrita na planilha
    abas_a_escrever = lista_entidades if lista_entidades is not None else abas_ordenadas
    nao_entidade = _abas_nao_entidade()
    for entidade_nome in abas_a_escrever:
        pontos = all_data.get(entidade_nome)
        if pontos:
            if entidade_nome in nao_entidade:
                # Um .dat com o nome de uma aba de config ('geral.dat', 'cores.dat'...) iria
                # escrever em cima dela: apagaria o caminho da pasta da aba Geral, ou
                # estouraria na aba protegida no meio da importação.
                ACUMULADO_IMPORTACAO['ignoradas'].append(entidade_nome)
                continue
            # Passa o objeto de configuração para a função de escrita
            write_to_sheet(doc, entidade_nome, pontos, modo_importacao, config)

    if lista_entidades is None:
        _esvaziar_abas_ausentes(doc, set(all_data) - set(ACUMULADO_IMPORTACAO['ignoradas']))


def _mover_aba_para_o_fim(doc, nome):
    """Põe a aba por último, em qualquer versão do LibreOffice.

    O destino do moveByName não quer dizer a mesma coisa em todo lugar: no 26.8,
    mover pro índice do último realmente põe no fim; no 24.2 (o do Ubuntu da CI) a
    aba para UMA posição antes do fim. Por isso aqui a gente move, confere e, se
    faltou, traz pra trás quem ficou depois dela -- movimento pra trás não tem essa
    ambiguidade. A ordem é conveniência (é ela que faz a planilha sair na ordem da
    MaisUsadas): se o LibreOffice não colaborar, desiste em silêncio.
    """
    abas = doc.getSheets()
    try:
        nomes = list(abas.ElementNames)
        if nomes[-1] == nome:
            return
        abas.moveByName(nome, len(nomes) - 1)
        nomes = list(abas.ElementNames)
        while nomes[-1] != nome:
            i = nomes.index(nome)
            abas.moveByName(nomes[i + 1], i)
            novos = list(abas.ElementNames)
            if novos == nomes:  # não saiu do lugar: insistir seria laço infinito
                return
            nomes = novos
    except Exception:
        pass  # nunca vale perder a importação por causa da ordem das abas


def _montar_matriz_aba(pontos_importados, ordem_atributos_aba):
    """Monta (cabecalhos, matriz) de uma aba a partir das linhas lidas (lógica PURA).
    `ordem_atributos_aba` é a prioridade de colunas da aba MaisUsadas; o resto das
    colunas segue a ordem em que o atributo apareceu pela primeira vez."""
    # Ordem de primeira aparição (dict preserva inserção), não a de um set: o conjunto
    # embaralha a cada execução (hash randomizado de str), e todo atributo fora da
    # MaisUsadas ia pra uma coluna diferente a cada importação -- o .dat exportado
    # mudava de ordem sem ninguém ter mexido em nada, e o diff no Git virava ruído.
    todos_atributos = list(dict.fromkeys(
        attr for p in pontos_importados if 'attributes' in p for attr in p['attributes']))
    prioridade_atributos = {attr: idx for idx, attr in enumerate(ordem_atributos_aba)}
    atributos_ordenados = sorted(
        todos_atributos,
        key=lambda a: prioridade_atributos.get(a, float('inf'))
    )
    cabecalhos = [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS] + atributos_ordenados
    header_to_col = {header: idx for idx, header in enumerate(cabecalhos)}
    
    # --- Preenchimento dos Dados (agora em lote para reduzir chamadas UNO) ---
    data_matrix = [cabecalhos]
    for ponto in pontos_importados:
        row_data = [''] * len(cabecalhos)
        row_data[0] = ponto.get('origem', '')
        row_data[1] = ponto['type']
        if ponto['type'] in [CODIGO_COMENTARIO_SIMPLES, CODIGO_INCLUDE, CODIGO_INCLUDE_COMENTADO]:
            row_data[2] = ponto.get('data', '')
        elif ponto['type'] in [CODIGO_BLOCO_ATIVO, CODIGO_BLOCO_COMENTADO]:
            row_data[2] = ponto.get('comment', '')
        if 'attributes' in ponto:
            for attr_key, attr_value in ponto['attributes'].items():
                col_idx = header_to_col.get(attr_key)
                if col_idx is not None:
                    row_data[col_idx] = attr_value
        data_matrix.append(row_data)
    return cabecalhos, data_matrix


def write_to_sheet(doc, sheet_name, pontos_importados, modo, config):
    """
    Versão limpa e otimizada. Escreve os dados e aplica formatação visual básica,
    incluindo o efeito zebrado nas linhas importadas + 20 linhas extras.
    """
    # --- Preparação da aba ---
    # Aba que já existe é LIMPA e reaproveitada, nunca removida. Antes o modo REPLACE
    # fazia removeByName + insertByName: entre um e outro a aba não existia, e uma falha
    # ali (ou em qualquer ponto seguinte da importação) deixava a planilha sem ela, sem
    # rollback nenhum -- o risco de destrutividade apontado na auditoria. Limpar dá o
    # mesmo resultado sem essa janela.
    if _aba_existe_ci(doc, sheet_name):
        sheet = _get_sheet(doc, sheet_name)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        sheet.getCellRangeByPosition(0, 0, addr.EndColumn, addr.EndRow).clearContents(FLAGS_LIMPAR_TUDO)
        if modo != 'UPDATE':
            # Recriar jogava a aba pro fim, e é assim que a ordem final acaba sendo a da
            # aba MaisUsadas (a importação escreve nessa ordem). Mover preserva esse
            # efeito sem destruir nada.
            _mover_aba_para_o_fim(doc, sheet.Name)
    else:
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(sheet_name, new_sheet)
        sheet = _get_sheet(doc, sheet_name)

    # --- Aplicação de Cores de Aba e Ordenação de Colunas (sem alterações) ---
    cor_aba = config.cores_entidades.get(sheet_name.lower())
    if cor_aba is not None and cor_aba != -1:
        sheet.TabColor = cor_aba
    cabecalhos, data_matrix = _montar_matriz_aba(
        pontos_importados, config.ordem_atributos.get(sheet_name.lower(), []))

    if data_matrix:
        num_rows = len(data_matrix) - 1
        num_cols = len(cabecalhos) - 1
        target_range = sheet.getCellRangeByPosition(0, 0, num_cols, num_rows)
        target_range.setDataArray(tuple(tuple(str(cell) for cell in row) for row in data_matrix))

    # --- PACOTE DE POLIMENTO VISUAL SIMPLIFICADO ---
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    last_col = cursor.getRangeAddress().EndColumn
    last_row = cursor.getRangeAddress().EndRow
    
    # Formatação do Cabeçalho
    if last_row >= 0:
        header_range = sheet.getCellRangeByPosition(0, 0, last_col, 0)
        header_range.HoriJustify = 2 # CENTER
        if cor_aba is not None and cor_aba != -1:
            header_range.CellBackColor = cor_aba

    # Alinhamento da Coluna "Gera"
    try:
        gera_col_idx = cabecalhos.index(CABEÇALHO_COLUNA_CONTROLE)
        if last_row > 0:
            gera_col_range = sheet.getCellRangeByPosition(gera_col_idx, 1, gera_col_idx, last_row)
            gera_col_range.HoriJustify = 2
    except ValueError: pass 

    # Largura Ótima das Colunas
    columns = sheet.getColumns()
    for i in range(last_col + 1):
        columns.getByIndex(i).OptimalWidth = True

    # --- LÓGICA OTIMIZADA: PINTURA DE LINHAS ALTERNADAS ---
    # Aplica a formatação nas linhas de dados + 20 linhas extras.
    num_linhas_formatar = last_row + 21
    if last_row > 0:
        for r in range(1, num_linhas_formatar):
            cor_a_aplicar = COR_LINHA_IMPAR if r % 2 != 0 else COR_LINHA_PAR
            row_range = sheet.getCellRangeByPosition(0, r, last_col, r)
            row_range.CellBackColor = cor_a_aplicar

    # O BLOCO DE CÓDIGO PARA VALIDAÇÃO DE DADOS FOI COMPLETAMENTE REMOVIDO

# ===============================================================
# =================== LÓGICA DE PARSING =========================
# ===============================================================
def parse_dat_file(file_path, relative_path, all_data, entidades_validas):
    start_time = time.perf_counter()
    lines = None
    for encoding in ENCODINGS_IMPORTACAO_SAGE:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue
        except IOError as e:
            print(f"Erro ao ler o arquivo {file_path}: {e}")
            return

    # latin-1 (último da lista) decodifica qualquer byte, então `lines` nunca fica None
    # aqui: a falha de leitura real (IOError) já retornou acima.

    i = 0
    chave_do_arquivo = _chave_do_arquivo(file_path)
    current_entidade_chave = chave_do_arquivo
    # Conta de linhas já guardadas, pra saber no fim se este arquivo rendeu alguma.
    linhas_antes = sum(len(v) for v in all_data.values())
    pending_comments = []
    current_block = None
    stats = {
        'lines_total': len(lines),
        'entities_imported': 0,
        'comments': 0,
        'ignored_lines': 0,
        'invalid_lines': 0,
        'warnings': 0,
        'preservadas': 0,
        'repetidos': 0,
        'arquivo_vazio': 0
    }
    last_i = -1
    iteracoes_sem_progresso = 0
    # Cache da olhada adiante (_proximo_estrutural): até que linha o último resultado vale.
    estrutural_ate = -1
    estrutural_tipo = None

    _log_importacao('DEBUG', f"Iniciando parse de {relative_path} com {len(lines)} linhas.")

    while i < len(lines):
        if i == last_i:
            iteracoes_sem_progresso += 1
            if iteracoes_sem_progresso >= WATCHDOG_MAX_ITERACOES_SEM_PROGRESSO:
                raise RuntimeError(
                    f"Watchdog de importação disparado em {relative_path} na linha {i + 1}: iterações sem avanço."
                )
        else:
            iteracoes_sem_progresso = 0
            last_i = i

        raw_line = lines[i]
        line_no = i + 1

        if current_block and current_block['type'] == CODIGO_BLOCO_COMENTADO:
            line_info = _classificar_linha_bloco_comentado(raw_line, entidades_validas)
        else:
            line_info = _classificar_linha_dat(raw_line, entidades_validas)

        _log_importacao(
            'DEBUG',
            f"{relative_path}:{line_no} bloco={current_block['identifier'] if current_block else '-'} tipo={line_info['type']}"
        )

        if current_block:
            # O carregador do SAGE trata toda linha que começa com ';' como comentário --
            # até ';ENTIDADE' e ';#include' -- e só encerra o registro num NOVO nome de
            # entidade. Atributo solto depois desses comentários continua o MESMO registro
            # (provado no carregador do SAGE: tag + ';E2M' + atributos = 1 registro). Encerrar o bloco
            # no ';ENTIDADE' partia o registro em dois: um bloco sem atributos, que o
            # carregador rejeita ("obrigatorio nao preenchido"), e atributos órfãos. Então,
            # num bloco ativo, o comentário só encerra o bloco se NÃO houver atributo vivo
            # depois dele antes da próxima entidade/include.
            if (current_block['type'] == CODIGO_BLOCO_ATIVO
                    and line_info['type'] in ('commented_entity_start', 'include_commented')):
                if i >= estrutural_ate:
                    estrutural_ate, estrutural_tipo = _proximo_estrutural(lines, i, entidades_validas)
                    if estrutural_ate is None:
                        estrutural_ate = len(lines)
                if estrutural_tipo == 'attribute':
                    current_block['comments'].append(line_info['stripped'].lstrip(';').lstrip())
                    stats['comments'] += 1
                    i += 1
                    continue

            if line_info['type'] in ['entity_start', 'commented_entity_start', 'include', 'include_commented', 'block_end']:
                _finalizar_bloco(current_block, all_data, relative_path, stats, line_no)
                current_block = None
                continue

            if line_info['type'] == 'blank':
                stats['ignored_lines'] += 1
                i += 1
                continue

            if line_info['type'] == 'comment':
                current_block['comments'].append(line_info.get('value', ''))
                stats['comments'] += 1
                i += 1
                continue

            if line_info['type'] == 'attribute':
                chave_attr = line_info['key']
                if chave_attr in current_block['attributes']:
                    # Atributo repetido no mesmo bloco: o dict guarda só o último, e o
                    # anterior sumia sem aviso. Fica como comentário do bloco (o SAGE
                    # rejeita atributo repetido em bloco ativo, e em bloco comentado a
                    # alternativa que o autor deixou de lado vale ser lembrada).
                    current_block['comments'].append(
                        f"{chave_attr} = {current_block['attributes'][chave_attr]}")
                    stats['repetidos'] += 1
                current_block['attributes'][chave_attr] = line_info['value']
                i += 1
                continue

            # Linha fora do formato dentro do bloco: vira comentário do bloco em vez de
            # sumir (antes só um aviso no console, que ninguém vê).
            current_block['comments'].append(line_info['stripped'])
            stats['preservadas'] += 1
            stats['invalid_lines'] += 1
            stats['warnings'] += 1
            _log_importacao(
                'WARN',
                f"{relative_path}:{line_no} linha inválida dentro do bloco {current_block['identifier']} "
                f"(preservada como comentário): {line_info['original']}",
                force=True
            )
            i += 1
            continue

        if line_info['type'] == 'blank':
            stats['ignored_lines'] += 1
            i += 1
            continue

        if line_info['type'] == 'include_commented':
            pending_comments = _descarregar_comentarios_pendentes(
                pending_comments, all_data, current_entidade_chave, relative_path)
            ponto = {'type': CODIGO_INCLUDE_COMENTADO, 'data': line_info['value'], 'origem': relative_path}
            all_data.setdefault(current_entidade_chave, []).append(ponto)
            i += 1
            continue

        if line_info['type'] == 'include':
            pending_comments = _descarregar_comentarios_pendentes(
                pending_comments, all_data, current_entidade_chave, relative_path)
            ponto = {'type': CODIGO_INCLUDE, 'data': line_info['value'], 'origem': relative_path}
            all_data.setdefault(current_entidade_chave, []).append(ponto)
            i += 1
            continue

        if line_info['type'] == 'entity_start':
            entidade_nome = line_info['entity']
            current_entidade_chave = entidade_nome.lower()
            current_block = _iniciar_bloco(
                entidade_nome,
                CODIGO_BLOCO_ATIVO,
                relative_path,
                comentarios_iniciais=pending_comments
            )
            pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'commented_entity_start':
            entidade_nome = line_info['entity']
            current_entidade_chave = entidade_nome.lower()
            current_block = _iniciar_bloco(
                entidade_nome,
                CODIGO_BLOCO_COMENTADO,
                relative_path,
                comentarios_iniciais=pending_comments
            )
            pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'comment':
            pending_comments.append(line_info.get('value', ''))
            stats['comments'] += 1
            i += 1
            continue

        # Atributo sem bloco, ou linha que não é nada do formato (um rodapé do tipo "Arquivo gerado
        # automaticamente por ...", sem ';', é o caso real mais comum em arquivos gerados por
        # ferramenta): o carregador do SAGE tolera, então a importação também. Mas
        # tolerar não é jogar fora -- a linha segue como comentário, no lugar em que
        # estava, e o round-trip a devolve com ';' na frente.
        pending_comments.append(line_info['stripped'])
        stats['preservadas'] += 1
        stats['warnings'] += 1
        stats['invalid_lines'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} "
            f"{'atributo fora de bloco' if line_info['type'] == 'attribute' else 'linha não reconhecida'} "
            f"(preservado como comentário): {line_info['original']}",
            force=True
        )
        i += 1

    if current_block:
        _finalizar_bloco(current_block, all_data, relative_path, stats, len(lines))

    # Comentários soltos no fim do arquivo (sem bloco aberto depois deles) também
    # precisam ser gravados, senão somem no round-trip (auditoria #4).
    pending_comments = _descarregar_comentarios_pendentes(
        pending_comments, all_data, current_entidade_chave, relative_path)

    # Arquivo que não rendeu linha nenhuma (só linhas em branco) ainda faz parte da base:
    # o pai dele o inclui com #include, e o carregador do SAGE ABORTA a carga inteira se o
    # arquivo incluído não existe ("Arquivo ... nao encontrado"). Sem marca na planilha a
    # exportação não o recriava -- medido no carregador do SAGE: com
    # arquivos vazios, o round-trip não carregava. A marca é um comentário vazio (Gera = n), que a
    # exportação devolve como ';' e o SAGE lê sem problema.
    if sum(len(v) for v in all_data.values()) == linhas_antes:
        all_data.setdefault(chave_do_arquivo, []).append({
            'type': CODIGO_COMENTARIO_SIMPLES,
            'data': '',
            'origem': relative_path
        })
        stats['arquivo_vazio'] = 1

    for chave_stat in ('preservadas', 'repetidos', 'arquivo_vazio'):
        ACUMULADO_IMPORTACAO[chave_stat] += stats[chave_stat]

    elapsed = time.perf_counter() - start_time
    _log_importacao(
        'INFO',
        (
            f"Arquivo {relative_path} processado em {elapsed:.3f}s. "
            f"linhas={stats['lines_total']} entidades={stats['entities_imported']} "
            f"comentarios={stats['comments']} ignoradas={stats['ignored_lines']} "
            f"invalidas={stats['invalid_lines']} avisos={stats['warnings']}"
        )
    )

# ===============================================================
# ================= FUNÇÕES DE EXPORTAÇÃO =======================
# ===============================================================

def _abas_nao_entidade():
    """Conjunto (lower) de toda aba que não é entidade SAGE -- aqui, as de
    configuração (FOLHAS_IGNORADAS). Montado uma vez por chamada, em vez de
    refazer a lista a cada aba testada; nome e uso alinhados com o núcleo
    compartilhado de import/export."""
    return {ign.lower() for ign in FOLHAS_IGNORADAS}


def exportar_dats(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # Mesma proteção das rotinas de importação: evita erro secundário no except
    # quando a aba Geral não existe ou foi renomeada.
    geral_sheet = None
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()
        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
        return

    geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("Processando exportação total...")
    abas_a_exportar = _abas_de_entidade(doc)
    try:
        erros, total_substituicoes, total_sem_origem = _exportar_folhas(abas_a_exportar, export_folder)
    except Exception as e:
        _reportar_falha_inesperada(
            geral_sheet, CELULA_STATUS_EXPORTACAO, "exportação", e,
            "Parte dos .dat pode já ter sido gravada; a versão anterior de cada um está no .bak ao lado.")
        return

    if erros:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: {'; '.join(erros)}")
    else:
        aviso = _avisos_exportacao(total_substituicoes, total_sem_origem)
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"Exportação total concluída com sucesso!{aviso}")


def exportar_parcial(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    # Exportação parcial também depende da aba Geral para ler pasta/lista.
    # Inicializar com None permite emitir diagnóstico seguro em caso de falha.
    geral_sheet = None
    
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()
        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
        return

    abas_a_exportar = []
    if active_sheet_name.lower() == NOME_ABA_GERAL.lower():
        nomes_entidades = _ler_entidades_parciais(geral_sheet)
        if not nomes_entidades:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("AVISO: Nenhuma entidade listada para exportação parcial.")
            return
        for nome in nomes_entidades:
            try:
                abas_a_exportar.append(_get_sheet(doc, nome))
            except KeyError:
                pass
    else:
        # Garante que a aba ativa não seja uma aba ignorada
        if active_sheet_name.lower() not in _abas_nao_entidade():
            abas_a_exportar.append(active_sheet)

    geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"Processando exportação de: {', '.join(s.getName() for s in abas_a_exportar)}...")
    try:
        # As outras abas entram só como contexto: um .dat que várias abas alimentam sai
        # inteiro, não só com as linhas da aba exportada (ver _exportar_folhas).
        erros, total_substituicoes, total_sem_origem = _exportar_folhas(
            abas_a_exportar, export_folder, contexto=_abas_de_entidade(doc))
    except Exception as e:
        _reportar_falha_inesperada(
            geral_sheet, CELULA_STATUS_EXPORTACAO, "exportação", e,
            "Parte dos .dat pode já ter sido gravada; a versão anterior de cada um está no .bak ao lado.")
        return

    if erros:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: {'; '.join(erros)}")
    else:
        aviso = _avisos_exportacao(total_substituicoes, total_sem_origem)
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"Exportação parcial concluída com sucesso!{aviso}")


def _avisos_exportacao(total_substituicoes, total_sem_origem):
    """Sufixo de aviso do status da exportação -- string vazia quando não há o que avisar.

    Agrega o que a exportação silenciava: linhas ignoradas por não terem "Origem"
    (auditoria #8) e caracteres sem equivalente em Latin-1 trocados por '?'."""
    partes = []
    if total_sem_origem:
        partes.append(f"{total_sem_origem} linha(s) sem Origem ignorada(s)")
    if total_substituicoes:
        partes.append(f"{total_substituicoes} caractere(s) sem equivalente em Latin-1 substituído(s) por '?'")
    return f" ({'; '.join(partes)})" if partes else ""


def _abas_de_entidade(doc):
    """As abas que são entidade (as que a exportação total percorre), na ordem da planilha."""
    return [s for s in doc.getSheets() if s.getName().lower() not in _abas_nao_entidade()]


def _coletar_folha(sheet):
    """Lê uma aba e devolve (erro_ou_None, {origem: [blocos de texto]}, n_linhas_sem_origem).

    Só lê: quem grava é _gravar_arquivos. A separação existe porque UM .dat pode ser
    alimentado por mais de uma aba (um 'pdf_1.dat' tem PDF e, se começa por um #include,
    uma linha na aba 'pdf_1' -- a aba do nome do arquivo --, além dos blocos na 'pdf').
    Gravar aba por aba fazia a segunda sobrescrever a primeira, e a primeira ia parar no
    .bak: parte do conteúdo sumia do .dat (medido em bases reais).
    """
    sheet_name = sheet.getName()
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    data_range = cursor.getRangeAddress()
    data_array = sheet.getCellRangeByPosition(0, 0, data_range.EndColumn, data_range.EndRow).getDataArray()

    if not data_array or len(data_array) < 2: return None, {}, 0

    headers = data_array[0]
    try:
        origem_col_idx = headers.index(CABEÇALHO_COLUNA_ORIGEM)
        gera_col_idx = headers.index(CABEÇALHO_COLUNA_CONTROLE)
        dados_col_idx = headers.index(CABEÇALHO_COLUNA_DADOS)
    except ValueError:
        return f"Aba '{sheet_name}' não possui as colunas 'Origem', 'Gera' ou 'Dados'.", {}, 0

    dados_agrupados_por_arquivo = {}
    linhas_sem_origem = 0

    for row_data in data_array[1:]:
        if len(row_data) <= max(origem_col_idx, gera_col_idx, dados_col_idx): continue
        origem_path = str(row_data[origem_col_idx])
        control_code = str(row_data[gera_col_idx]).lower()
        if not control_code or control_code == CODIGO_IGNORAR_LINHA: continue
        if not origem_path:
            # Sem "Origem" não há destino: ponto criado à mão na planilha nunca chegava
            # ao .dat e nada avisava (auditoria #8). Continua não exportando -- não dá
            # pra adivinhar o arquivo --, mas agora aparece no status.
            linhas_sem_origem += 1
            continue
        dados_agrupados_por_arquivo.setdefault(origem_path, [])
        bloco_final = None
        dado_principal = str(row_data[dados_col_idx])
        if control_code == CODIGO_INCLUDE and dado_principal:
            bloco_final = f'#include {dado_principal}'
        elif control_code == CODIGO_INCLUDE_COMENTADO and dado_principal:
            bloco_final = f';#include {dado_principal}'
        elif control_code == CODIGO_COMENTARIO_SIMPLES:
            bloco_final = f';{dado_principal}'
        elif control_code in [CODIGO_BLOCO_ATIVO, CODIGO_BLOCO_COMENTADO]:
            comment_lines = [line for line in dado_principal.splitlines()]
            attribute_lines = []
            for col_idx, header in enumerate(headers):
                if header in [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS]:
                    continue
                value = str(row_data[col_idx]) if len(row_data) > col_idx else ""
                if value:
                    attribute_lines.append(f"\t{header} = {value}")

            if comment_lines or attribute_lines:
                if control_code == CODIGO_BLOCO_COMENTADO:
                    point_lines = [f";{sheet_name.upper()}"]
                    point_lines.extend([f";{line}" for line in comment_lines])
                    point_lines.extend([f";{line}" for line in attribute_lines])
                else:
                    point_lines = [sheet_name.upper()]
                    point_lines.extend([f";{line}" for line in comment_lines])
                    point_lines.extend(attribute_lines)
                bloco_final = "\n".join(point_lines)
        if bloco_final is not None:
            dados_agrupados_por_arquivo[origem_path].append(bloco_final)

    return None, dados_agrupados_por_arquivo, linhas_sem_origem


def _gravar_arquivos(export_folder, arquivos):
    """Grava {origem: [blocos]} na pasta de saída, um arquivo por vez, com .bak do
    anterior. Devolve (lista_de_erros, n_substituicoes) -- n_substituicoes é a soma de
    caracteres sem equivalente em latin-1 trocados por '?' ao longo de todos os arquivos.
    Um arquivo que falha não impede a gravação dos demais."""
    total_substituicoes = 0
    erros = []
    for relative_path, file_content_list in arquivos.items():
        try:
            full_output_path = os.path.join(export_folder, relative_path)

            # --- INÍCIO DA LÓGICA DE BACKUP ---
            if os.path.exists(full_output_path):
                backup_path = full_output_path + ".bak"
                # Remove um backup antigo, se existir, para evitar erros no rename
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                os.rename(full_output_path, backup_path)
            # --- FIM DA LÓGICA DE BACKUP ---

            os.makedirs(os.path.dirname(full_output_path), exist_ok=True)
            conteudo, n_subst = _sanitizar_para_latin1("\n\n".join(file_content_list) + "\n")
            total_substituicoes += n_subst
            with open(full_output_path, 'w', encoding=ENCODING_EXPORTACAO_SAGE) as f:
                f.write(conteudo)
        except IOError as e:
            erros.append(f"Falha ao escrever {relative_path}: {e}")
    return erros, total_substituicoes


def _exportar_folhas(folhas, export_folder, contexto=None):
    """Exporta as abas `folhas` e devolve (lista_de_erros, n_substituicoes, n_sem_origem).

    Cada arquivo é gravado UMA vez, com as linhas de todas as abas que apontam pra ele.
    `contexto` (padrão: as próprias folhas) é o conjunto de abas consultado pra montar
    cada arquivo: na exportação parcial, passa-se todas as abas de entidade, pra que um
    .dat compartilhado não saia truncado só com as linhas da aba exportada.

    Ordem dentro do arquivo: primeiro as linhas da aba que leva o nome do arquivo (é onde
    caem o cabeçalho e os #include que abrem o arquivo), depois as das demais abas na
    ordem da planilha. A ordem relativa entre entidades diferentes no mesmo arquivo não
    existe na planilha e o SAGE não depende dela."""
    contexto = list(contexto) if contexto is not None else list(folhas)
    nomes_contexto = {s.getName() for s in contexto}
    todas = list(contexto) + [s for s in folhas if s.getName() not in nomes_contexto]

    erros = []
    por_aba = []  # (ordem, nome_da_aba, {origem: blocos})
    sem_origem = 0
    selecionadas = {s.getName() for s in folhas}
    for ordem, sheet in enumerate(todas):
        erro, arquivos, n_sem = _coletar_folha(sheet)
        if sheet.getName() in selecionadas:
            sem_origem += n_sem
            if erro:
                erros.append(erro)
        por_aba.append((ordem, sheet.getName(), arquivos))

    origens = []
    for ordem, nome, arquivos in por_aba:
        if nome in selecionadas:
            origens.extend(o for o in arquivos if o not in origens)

    a_gravar = {}
    for origem in origens:
        stem = _chave_do_arquivo(origem)
        partes = [(0 if nome.lower() == stem else 1, ordem, arquivos[origem])
                  for ordem, nome, arquivos in por_aba if origem in arquivos]
        a_gravar[origem] = [b for _, _, blocos in sorted(partes, key=lambda t: t[:2]) for b in blocos]

    erros_gravacao, substituicoes = _gravar_arquivos(export_folder, a_gravar)
    erros.extend(erros_gravacao)
    return erros, substituicoes, sem_origem


def _exportar_folha(sheet, export_folder):
    """Exporta uma única aba (compatibilidade: a exportação usa _exportar_folhas).
    Retorna (erro_ou_None, n_substituicoes, n_linhas_sem_origem)."""
    erros, substituicoes, sem_origem = _exportar_folhas([sheet], export_folder)
    return (erros[0] if erros else None), substituicoes, sem_origem

# ===============================================================
# ================= FUNÇÃO DE CORES DO TEMA =====================
# ===============================================================

# --- CONFIGURAÇÃO DA ABA DE CORES ---
# Verifique o nome da aba onde a tabela de cores se encontra.
# Se for "TEMA ESCURO" use-o, caso contrário ajuste.
NOME_ABA_TEMA_CORES = "Cores" 
# Colunas:
COL_R_DEC = 5  # Coluna F (índice 5)
COL_G_DEC = 6  # Coluna G (índice 6)
COL_B_DEC = 7  # Coluna H (índice 7)
COL_COR_AMOSTRA = 11 # Coluna L (índice 11)


def rgb_to_bgr_decimal(r, g, b):
    """
    Converte os valores RGB (Red, Green, Blue) de 0-255 para
    o formato BGR Decimal (Blue-Green-Red), que é o padrão
    de cor numérica para CellBackColor no LibreOffice.
    Fórmula: (B * 256^2) + (G * 256^1) + (R * 256^0)
    """
    try:
        # Garante que os valores são inteiros e no intervalo 0-255.
        r = int(r) if 0 <= r <= 255 else 0
        g = int(g) if 0 <= g <= 255 else 0
        b = int(b) if 0 <= b <= 255 else 0
        
        # Se os três valores forem zero (preto), retorna a cor
        if r == 0 and g == 0 and b == 0:
             return 0 # Preto é o valor 0
        
        # Se houver algum valor diferente de zero, calcula o BGR.
        # B é o mais significativo (<< 16), G é o intermediário (<< 8), R é o menos significativo.
        bgr_decimal = (r * 65536) + (g * 256) + b
        return bgr_decimal
    except:
        # Retorna -1 para sinalizar que a célula deve ficar sem cor
        return -1


def atualizar_amostras_cores(*args):
    """
    Lê os valores RGB Decimais (Colunas H, I, J) da aba do tema
    e aplica a cor de fundo (CellBackColor) na coluna de amostra (L).
    """
    try:
        doc = XSCRIPTCONTEXT.getDocument() # type: ignore
        sheets = doc.getSheets()

        if not sheets.hasByName(NOME_ABA_TEMA_CORES):
            print(f"ERRO: A aba de tema '{NOME_ABA_TEMA_CORES}' não foi encontrada.")
            return

        sheet = _get_sheet(doc, NOME_ABA_TEMA_CORES)
        
        # 1. Determina a última linha preenchida para otimizar a leitura
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_row = cursor.getRangeAddress().EndRow
        
        # 2. Leitura otimizada de um bloco de dados (Colunas R DEC a B DEC)
        # Lemos de R DEC (H) até B DEC (J) da linha 1 até a última.
        # Os índices iniciais são: COL_R_DEC (7) e LINHA 1.
        data_range = sheet.getCellRangeByPosition(COL_R_DEC, 1, COL_B_DEC, last_row)
        data = data_range.getDataArray()
        
        # 3. Itera sobre os dados lidos
        for row_idx, row_data in enumerate(data):
            # Índices relativos ao bloco de dados lido: 0=R, 1=G, 2=B
            r_dec = row_data[0]
            g_dec = row_data[1]
            b_dec = row_data[2]
            
            # Converte para BGR Decimal
            bgr_cor = rgb_to_bgr_decimal(r_dec, g_dec, b_dec)
            
            # A linha de destino é o índice da linha atual (row_idx) + 1 (cabeçalho)
            target_row = row_idx + 1 
            
            amostra_cell = sheet.getCellByPosition(COL_COR_AMOSTRA, target_row)
            
            if bgr_cor == -1:
                # Se a conversão falhar ou os valores não existirem, remove a cor
                amostra_cell.CellBackColor = -1 # O valor -1 no LibreOffice remove o preenchimento
            else:
                # Aplica a cor
                amostra_cell.CellBackColor = bgr_cor
                
        print("Amostras de cores do tema atualizadas com sucesso!")

    except Exception as e:
        print(f"ERRO ao aplicar as cores do tema: {e}")

# ===============================================================
# ================= EXPOSIÇÃO PARA LIBREOFFICE ==================
# ===============================================================
g_exportedScripts = importar_dats, exportar_dats, importar_parcial, exportar_parcial, atualizar_amostras_cores
