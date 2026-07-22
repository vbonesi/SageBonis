# -*- coding: utf-8 -*-
"""
sync_menu.py - Garante que o menu "SageBonis" e a toolbar customizada dentro do
.ods incluam as funcoes da Trilha Completa (verificar_base, unificar_pontos,
extrair_pontos, trocar_id_global, estatistica_base, gerir_includes).

Um arquivo .ods e um ZIP; o menu/toolbar customizados vivem em
Configurations2/menubar/menubar.xml e Configurations2/toolbar/sagebonis.xml. Ate
agora só tinham os 5 itens da Trilha Simples (importar/exportar/cores) -- as
funcoes da Completa só eram acessiveis rodando a macro manualmente (Ferramentas >
Macros). Esta ferramenta adiciona os itens que faltarem, com um separador antes do
grupo novo, preservando a estrutura exigida pelo ODF (mimetype primeiro, sem
compressao) -- mesmo cuidado do sync_macro.py. E idempotente: rodar de novo depois
que os itens já existem não duplica nada.

So faz sentido no .ods da Trilha Completa -- a Simples nao tem essas funcoes.

Uso:
    python sync_menu.py status --ods completa/SageBonis.ods   # mostra o que falta
    python sync_menu.py sync   --ods completa/SageBonis.ods   # adiciona o que falta
"""

import argparse
import os
import re
import shutil
import sys
import zipfile

CAMINHO_MENUBAR = "Configurations2/menubar/menubar.xml"
CAMINHO_TOOLBAR = "Configurations2/toolbar/sagebonis.xml"

# (funcao, rotulo) -- ordem de exibicao do grupo "Completa". O grupo "Simples" ja
# existe nos dois arquivos originais e nao e tocado.
ITENS_COMPLETA = [
    ("verificar_base", "Verificar Base"),
    ("unificar_pontos", "Unificar Pontos"),
    ("extrair_pontos", "Extrair Pontos"),
    ("trocar_id_global", "Trocar ID"),
    ("estatistica_base", "Estatística"),
    ("gerir_includes", "Gerir Includes"),
]


def _href(funcao):
    # &amp; (nao &) porque isso vai dentro de um atributo XML.
    return "vnd.sun.star.script:ImportadorSAGE.py$%s?language=Python&amp;location=document" % funcao


def _itens_faltantes(conteudo, itens):
    return [(f, r) for f, r in itens if _href(f) not in conteudo]


def _atualizar_menubar(conteudo):
    faltantes = _itens_faltantes(conteudo, ITENS_COMPLETA)
    if not faltantes:
        return conteudo, faltantes
    padrao = re.compile(
        r'(<menu:menu menu:id="private:resource/menubar/sagebonis"[^>]*>\s*<menu:menupopup>)'
        r'(.*?)'
        r'(</menu:menupopup>\s*</menu:menu>)',
        re.DOTALL,
    )
    m = padrao.search(conteudo)
    if not m:
        sys.exit("ERRO: bloco do menu 'sagebonis' nao encontrado em menubar.xml -- nada alterado.")
    separador = "      <menu:menuseparator/>\n" if m.group(2).strip() else ""
    itens_xml = "".join(
        '      <menu:menuitem menu:id="%s" menu:label="%s"/>\n' % (_href(f), r)
        for f, r in faltantes
    )
    meio_novo = m.group(2) + separador + itens_xml
    return conteudo[:m.start()] + m.group(1) + meio_novo + m.group(3) + conteudo[m.end():], faltantes


def _atualizar_toolbar(conteudo):
    faltantes = _itens_faltantes(conteudo, ITENS_COMPLETA)
    if not faltantes:
        return conteudo, faltantes
    fechamento = "</toolbar:toolbar>"
    if fechamento not in conteudo:
        sys.exit("ERRO: '</toolbar:toolbar>' nao encontrado em sagebonis.xml -- nada alterado.")
    separador = " <toolbar:toolbarseparator/>\n" if "<toolbar:toolbaritem" in conteudo else ""
    itens_xml = "".join(
        ' <toolbar:toolbaritem xlink:href="%s" toolbar:text="%s" toolbar:style="text"/>\n'
        % (_href(f), r)
        for f, r in faltantes
    )
    return conteudo.replace(fechamento, separador + itens_xml + fechamento), faltantes


def _ler(ods_path, caminho_interno):
    with zipfile.ZipFile(ods_path, "r") as z:
        if caminho_interno not in z.namelist():
            sys.exit(f"ERRO: '{caminho_interno}' nao existe dentro de {ods_path}.")
        return z.read(caminho_interno).decode("utf-8")


def cmd_status(ods_path):
    menubar = _ler(ods_path, CAMINHO_MENUBAR)
    toolbar = _ler(ods_path, CAMINHO_TOOLBAR)
    falta_menu = _itens_faltantes(menubar, ITENS_COMPLETA)
    falta_barra = _itens_faltantes(toolbar, ITENS_COMPLETA)
    if not falta_menu and not falta_barra:
        print(f"OK: menu e toolbar de {ods_path} ja tem todos os itens da Completa.")
        return
    if falta_menu:
        print("Faltam no menu:", ", ".join(r for _, r in falta_menu))
    if falta_barra:
        print("Faltam na toolbar:", ", ".join(r for _, r in falta_barra))


def cmd_sync(ods_path, fazer_backup=True):
    menubar = _ler(ods_path, CAMINHO_MENUBAR)
    toolbar = _ler(ods_path, CAMINHO_TOOLBAR)
    novo_menubar, faltantes_menu = _atualizar_menubar(menubar)
    novo_toolbar, faltantes_barra = _atualizar_toolbar(toolbar)

    if not faltantes_menu and not faltantes_barra:
        print(f"OK: nada a fazer, {ods_path} ja tem todos os itens da Completa.")
        return

    if fazer_backup:
        bak = ods_path + ".bak"
        shutil.copy2(ods_path, bak)
        print(f"Backup criado: {bak}")

    substituicoes = {CAMINHO_MENUBAR: novo_menubar.encode("utf-8"),
                     CAMINHO_TOOLBAR: novo_toolbar.encode("utf-8")}

    tmp = ods_path + ".tmp"
    with zipfile.ZipFile(ods_path, "r") as zin:
        infos = zin.infolist()
        mimetype = [i for i in infos if i.filename == "mimetype"]
        resto = [i for i in infos if i.filename != "mimetype"]
        ordenadas = mimetype + resto

        with zipfile.ZipFile(tmp, "w") as zout:
            for info in ordenadas:
                dados = substituicoes.get(info.filename, None)
                if dados is None:
                    dados = zin.read(info.filename)
                compress = zipfile.ZIP_STORED if info.filename == "mimetype" else (
                    info.compress_type or zipfile.ZIP_DEFLATED
                )
                novo_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                novo_info.compress_type = compress
                novo_info.external_attr = info.external_attr
                novo_info.internal_attr = info.internal_attr
                novo_info.create_system = info.create_system
                zout.writestr(novo_info, dados)

    os.replace(tmp, ods_path)
    if faltantes_menu:
        print("Adicionado ao menu:", ", ".join(r for _, r in faltantes_menu))
    if faltantes_barra:
        print("Adicionado a toolbar:", ", ".join(r for _, r in faltantes_barra))
    print(f"OK: {ods_path} atualizado.")


def main():
    p = argparse.ArgumentParser(description="Sincroniza o menu/toolbar SageBonis com as funcoes disponiveis")
    p.add_argument("acao", choices=["status", "sync"])
    p.add_argument("--ods", default="completa/SageBonis.ods")
    p.add_argument("--no-backup", action="store_true")
    a = p.parse_args()

    if a.acao == "status":
        cmd_status(a.ods)
    elif a.acao == "sync":
        cmd_sync(a.ods, fazer_backup=not a.no_backup)


if __name__ == "__main__":
    main()
