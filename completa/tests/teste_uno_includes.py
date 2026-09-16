# -*- coding: utf-8 -*-
"""Teste ponta a ponta, via LibreOffice/UNO, da macro gerir_includes."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

falhas = []


def check(nome, cond):
    print("[%s] %s" % ("OK" if cond else "FALHOU", nome))
    if not cond:
        falhas.append(nome)


ORIGEM = "__teste_uno_includes__.dat"
ANTIGO = "__TESTE_INCLUDE_ANTIGO__"
NOVO = "__TESTE_INCLUDE_NOVO__"

with TesteUno(porta=2300) as t:
    # A primeira execução deve criar/garantir as duas abas administrativas.
    t.chamar_macro("gerir_includes")
    check("cria/garante aba SubstituirIncludes", t.sheet_existe("SubstituirIncludes"))
    check("cria/garante aba RelatorioIncludes", t.sheet_existe("RelatorioIncludes"))

    linha_pds = t.proxima_linha_livre("PDS")
    t.escrever_linha("PDS", linha_pds, {
        "Origem": ORIGEM, "Gera": "i", "Comentario/Include": ANTIGO + "/ativo.dat",
    })
    t.escrever_linha("PDS", linha_pds + 1, {
        "Origem": ORIGEM, "Gera": "u", "Comentario/Include": ANTIGO + "/comentado.dat",
    })
    t.escrever_linha("PDS", linha_pds + 2, {
        "Origem": ORIGEM, "Gera": "x", "Comentario/Include": ANTIGO + "/nao_alterar",
        "ID": "__TESTE_PONTO_NORMAL__",
    })

    linha_regra = t.proxima_linha_livre("SubstituirIncludes")
    t.escrever_linha("SubstituirIncludes", linha_regra, {
        "Buscar": ANTIGO, "Substituir": NOVO, "Ativa": "x",
    })
    t.escrever_linha("SubstituirIncludes", linha_regra + 1, {
        "Buscar": NOVO, "Substituir": "__REGRA_INATIVA__", "Ativa": "n",
    })

    t.chamar_macro("gerir_includes")
    linhas = [r for r in t.ler_aba("PDS") if r.get("Origem") == ORIGEM]
    includes = [r for r in linhas if r.get("Gera", "").strip().lower() in ("i", "u")]
    ponto = next((r for r in linhas if r.get("ID") == "__TESTE_PONTO_NORMAL__"), None)
    check("substitui includes ativo e comentado", len(includes) == 2 and all(
        r.get("Comentario/Include", "").startswith(NOVO) for r in includes))
    check("regra inativa não é aplicada", all(
        "__REGRA_INATIVA__" not in r.get("Comentario/Include", "") for r in includes))
    check("não altera linha normal Gera=x", ponto is not None and
          ponto.get("Comentario/Include") == ANTIGO + "/nao_alterar")

    relatorio = [r.get("Include", "") for r in t.ler_aba("RelatorioIncludes")]
    check("relatório lista os dois paths atualizados",
          sum(NOVO in texto for texto in relatorio) == 2)

    # Rodar novamente não deve duplicar nem transformar de novo os dados do teste.
    t.chamar_macro("gerir_includes")
    relatorio_2 = [r.get("Include", "") for r in t.ler_aba("RelatorioIncludes")]
    check("segunda execução é idempotente", sum(NOVO in texto for texto in relatorio_2) == 2 and
          not any("__REGRA_INATIVA__" in texto for texto in relatorio_2))

print()
if falhas:
    print("%d checagem(ns) FALHOU/FALHARAM: %s" % (len(falhas), falhas))
    raise SystemExit(1)
print("Todas as checagens UNO de includes passaram.")
