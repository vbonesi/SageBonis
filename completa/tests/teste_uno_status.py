# -*- coding: utf-8 -*-
"""Confere, via UNO real, que toda macro da Trilha Completa reporta o que fez.

Antes, só importar/exportar davam retorno (Geral!B4/B7). As macras da Completa
terminavam em silêncio: "não fez nada" era indistinguível de "fez e não achou nada"
-- o caso que mais confunde é justamente o de nenhuma linha ativa na aba de config.

Cada macro escreve uma linha em Geral!H23 (CELULA_STATUS_COMPLETA). Este teste roda
as 7 na mesma planilha descartável e exige que o status mude e diga de quem é.

Roda com: python completa/tests/teste_uno_status.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

CELULA_STATUS = (7, 22)  # H23 -- espelha CELULA_STATUS_COMPLETA da macro

falhas = []


def check(nome, cond, detalhe=""):
    print("[%s] %s%s" % ("OK" if cond else "FALHOU", nome, "" if cond else " -> " + detalhe))
    if not cond:
        falhas.append(nome)


MACROS = [
    "verificar_base",
    "estatistica_base",
    "gerir_includes",
    "extrair_pontos",
    "gerar_ied",
    "unificar_pontos",
    "trocar_id_global",
]

with TesteUno(porta=2700) as t:
    anterior = None
    for macro in MACROS:
        t.definir_celula("Geral", CELULA_STATUS[0], CELULA_STATUS[1], "")
        t.chamar_macro(macro)
        status = t.ler_celula("Geral", *CELULA_STATUS)
        check("%s escreve status na Geral" % macro, bool(status.strip()), repr(status))
        check("%s se identifica no status" % macro, status.startswith(macro + ":"), repr(status))
        anterior = status

    # O rótulo ao lado do status também é escrito pela macro (planilha antiga não tem).
    rotulo = t.ler_celula("Geral", CELULA_STATUS[0], CELULA_STATUS[1] - 1)
    check("rótulo do status é criado na Geral", "Trilha Completa" in rotulo, repr(rotulo))

    # Caso que motivou o recurso: aba de config sem nenhuma linha ativa. A macro não
    # tem o que fazer -- e isso precisa aparecer, não virar silêncio.
    t.definir_celula("Geral", CELULA_STATUS[0], CELULA_STATUS[1], "")
    t.chamar_macro("trocar_id_global")
    status_vazio = t.ler_celula("Geral", *CELULA_STATUS)
    check("macro sem nada a fazer explica o porquê",
          "nada alterado" in status_vazio and "Ativa" in status_vazio, repr(status_vazio))

    # E o caminho oposto: com trabalho de verdade, o status conta quanto foi feito.
    linha = t.proxima_linha_livre("IEDs")
    t.escrever_linha("IEDs", linha, {"ID": "STATUS_T", "Protocolo": "104",
                                     "Direcao": "Aquisicao", "Gera": "x"})
    t.chamar_macro("gerar_ied")
    status_ied = t.ler_celula("Geral", *CELULA_STATUS)
    check("gerar_ied conta IEDs e linhas geradas",
          "IED(s) gerado(s)" in status_ied and "LSC" in status_ied.upper(), repr(status_ied))

print()
if falhas:
    print("%d checagem(ns) FALHOU/FALHARAM: %s" % (len(falhas), falhas))
    raise SystemExit(1)
print("Todas as checagens UNO de status passaram.")
