# -*- coding: utf-8 -*-
"""Confere, via UNO real, que os botões da planilha estão ligados às macras certas.

O que este teste protege: macro nova entrando na trilha sem botão na planilha (ou
com o botão apontando pro script errado). A ligação vive num ScriptEventDescriptor
registrado no formulário da aba -- é isso que o clique dispara --, então é isso que
o teste lê de volta, e não só a existência do desenho do botão.

Os botões são criados por sync_botoes.py; este teste não os cria, só verifica o que
a planilha distribuída traz.

Roda com: python completa/tests/teste_uno_botoes.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from uno_harness import TesteUno  # noqa: E402
import sync_botoes  # noqa: E402

falhas = []


def check(nome, cond):
    print("[%s] %s" % ("OK" if cond else "FALHOU", nome))
    if not cond:
        falhas.append(nome)


def eventos_por_controle(sheet):
    """{nome do controle: ScriptCode do evento actionPerformed} de uma aba."""
    saida = {}
    forms = sheet.DrawPage.Forms
    for f in range(forms.Count):
        form = forms.getByIndex(f)
        for i in range(form.Count):
            nome = form.getByIndex(i).Name
            for ev in form.getScriptEvents(i):
                if ev.EventMethod == "actionPerformed":
                    saida[nome] = ev.ScriptCode
    return saida


with TesteUno(porta=2600) as t:
    doc = t.doc
    eventos = {}
    for i in range(doc.Sheets.Count):
        sheet = doc.Sheets.getByIndex(i)
        for nome, script in eventos_por_controle(sheet).items():
            eventos[nome] = (sheet.Name, script)

    # --- Botões originais da Trilha Simples: continuam lá e continuam ligados ---
    for nome_controle, funcao in (("Importar", "importar_dats"),
                                  ("Exportar", "exportar_dats"),
                                  ("Cores", "atualizar_amostras_cores")):
        aba_script = eventos.get(nome_controle)
        check("botão original '%s' continua ligado a %s" % (nome_controle, funcao),
              aba_script is not None and funcao in aba_script[1])

    # --- Painel da aba Geral: uma macro da Completa por botão ---
    for funcao, _rotulo in sync_botoes.ITENS_COMPLETA + sync_botoes.ITENS_PARCIAIS:
        nome = sync_botoes.PREFIXO + funcao
        aba_script = eventos.get(nome)
        check("painel da Geral tem botão de %s" % funcao,
              aba_script is not None and aba_script[0] == "Geral"
              and aba_script[1] == sync_botoes._href(funcao))

    # --- Botão contextual na aba de config de cada recurso ---
    for nome_aba, funcao, _rotulo in sync_botoes.BOTOES_CONTEXTUAIS:
        if not doc.Sheets.hasByName(nome_aba):
            continue
        nome = sync_botoes.PREFIXO + funcao + "_" + nome_aba.lower()
        aba_script = eventos.get(nome)
        check("aba %s tem botão contextual de %s" % (nome_aba, funcao),
              aba_script is not None and aba_script[0] == nome_aba
              and aba_script[1] == sync_botoes._href(funcao))

    # --- Toda macro exposta no menu também tem botão (o menu não pode andar sozinho) ---
    funcoes_com_botao = {s.split("$", 1)[1].split("?", 1)[0]
                         for _aba, s in eventos.values() if "$" in s}
    menu_sem_botao = sorted({f for f, _r in sync_botoes.ITENS_COMPLETA} - funcoes_com_botao)
    check("nenhuma macro da Completa ficou só no menu (%s)" % (menu_sem_botao or "nenhuma"),
          not menu_sem_botao)

    # --- O botão contextual não pode cobrir dado: fica à direita do cabeçalho ---
    for nome_aba, funcao, _rotulo in sync_botoes.BOTOES_CONTEXTUAIS:
        if not doc.Sheets.hasByName(nome_aba):
            continue
        sheet = doc.Sheets.getByName(nome_aba)
        ultima = sync_botoes._ultima_coluna_cabecalho(sheet)
        nome = sync_botoes.PREFIXO + funcao + "_" + nome_aba.lower()
        coluna_ancora = None
        dp = sheet.DrawPage
        for i in range(dp.Count):
            shape = dp.getByIndex(i)
            try:
                if shape.Control.Name != nome:
                    continue
            except Exception:
                continue
            coluna_ancora = shape.Anchor.CellAddress.Column
        check("botão de %s fica depois da última coluna de cabeçalho (col %s > %s)"
              % (nome_aba, coluna_ancora, ultima),
              coluna_ancora is not None and coluna_ancora > ultima)

print()
if falhas:
    print("%d checagem(ns) FALHOU/FALHARAM: %s" % (len(falhas), falhas))
    raise SystemExit(1)
print("Todas as checagens UNO de botões passaram.")
