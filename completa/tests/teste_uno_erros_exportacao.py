# -*- coding: utf-8 -*-
"""Testa diagnósticos de erro da importação/exportação em LibreOffice/UNO real."""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

falhas = []


def check(nome, cond):
    print("[%s] %s" % ("OK" if cond else "FALHOU", nome))
    if not cond:
        falhas.append(nome)


pasta_inexistente = tempfile.mkdtemp(prefix="sagebonis_inexistente_")
shutil.rmtree(pasta_inexistente)
pasta_valida = tempfile.mkdtemp(prefix="sagebonis_erros_exportacao_")
pasta_import_invalida = tempfile.mkdtemp(prefix="sagebonis_falha_import_")

try:
    with TesteUno(porta=2400) as t:
        t.definir_celula("Geral", 0, 6, pasta_inexistente)
        t.chamar_macro("exportar_dats")
        status_total = t.ler_celula("Geral", 1, 6)
        check("exportação total rejeita pasta inexistente",
              "ERRO" in status_total.upper() and "PASTA VÁLIDA" in status_total.upper())

        t.ativar_aba("PDS")
        t.chamar_macro("exportar_parcial")
        status_parcial = t.ler_celula("Geral", 1, 6)
        check("exportação parcial rejeita pasta inexistente",
              "ERRO" in status_parcial.upper() and "PASTA VÁLIDA" in status_parcial.upper())

        # Uma aba tabular desconhecida e sem os cabeçalhos mínimos deve produzir
        # diagnóstico, sem derrubar o LibreOffice nem escrever saída parcial nela.
        nome_invalida = "AbaInvalidaTeste"
        t.doc.Sheets.insertNewByName(nome_invalida, t.doc.Sheets.getCount())
        t.definir_celula(nome_invalida, 0, 0, "CabecalhoIncorreto")
        t.definir_celula(nome_invalida, 0, 1, "dado")
        t.definir_celula("Geral", 0, 6, pasta_valida)
        t.chamar_macro("exportar_dats")
        status_aba = t.ler_celula("Geral", 1, 6)
        check("exportação total reporta aba sem colunas obrigatórias",
              "ERRO" in status_aba.upper() and nome_invalida in status_aba)

        # Linha sem "Origem": não tem destino, então não exporta -- mas precisa
        # aparecer no status em vez de sumir calada (auditoria #8).
        linha_sem_origem = t.proxima_linha_livre("PDS")
        t.escrever_linha("PDS", linha_sem_origem, {
            "Gera": "n", "Comentario/Include": "linha sem origem",
        })
        t.definir_celula("Geral", 0, 6, pasta_valida)
        t.ativar_aba("PDS")
        t.chamar_macro("exportar_parcial")
        status_sem_origem = t.ler_celula("Geral", 1, 6)
        check("linha sem Origem é contada no status da exportação",
              "sucesso" in status_sem_origem.lower() and
              "sem Origem ignorada" in status_sem_origem)

        # Cria um arquivo no lugar em que a macro precisaria criar uma pasta.
        # Isso força FileExistsError/OSError sem depender de permissões Unix.
        bloqueio = os.path.join(pasta_valida, "bloqueio")
        with open(bloqueio, "w", encoding="ascii") as f:
            f.write("impede a criacao da subpasta")
        linha = t.proxima_linha_livre("PDS")
        t.escrever_linha("PDS", linha, {
            "Origem": "bloqueio/pds.dat", "Gera": "n",
            "Comentario/Include": "forçar falha de escrita",
        })
        t.ativar_aba("PDS")
        t.chamar_macro("exportar_parcial")
        status_escrita = t.ler_celula("Geral", 1, 6)
        check("falha física de escrita é capturada e identifica o arquivo",
              "ERRO" in status_escrita.upper() and
              "bloqueio/pds.dat" in status_escrita and
              "Falha ao escrever" in status_escrita)

        # Falha inesperada no meio da importação precisa virar status de ERRO em vez de
        # deixar "Processando..." na tela com a planilha meio reescrita (auditoria #9).
        # Um .dat cujo nome não pode virar nome de aba no LibreOffice (':') força uma
        # exceção UNO real na escrita, já depois do parse. Fica por último de propósito:
        # a importação total reescreve as abas usadas pelas checagens acima.
        with open(os.path.join(pasta_import_invalida, "a:b.dat"), "w", encoding="latin-1") as f:
            f.write("A:B\n\tID = X1\n\tNOME = Teste\n")
        t.definir_celula("Geral", 0, 3, pasta_import_invalida)
        t.chamar_macro("importar_dats")
        status_falha = t.ler_celula("Geral", 1, 3)
        check("falha inesperada na importação vira status de ERRO",
              status_falha.startswith("ERRO: falha inesperada") and
              "feche sem salvar" in status_falha)
finally:
    shutil.rmtree(pasta_valida, ignore_errors=True)
    shutil.rmtree(pasta_import_invalida, ignore_errors=True)

print()
if falhas:
    print("%d checagem(ns) FALHOU/FALHARAM: %s" % (len(falhas), falhas))
    raise SystemExit(1)
print("Todas as checagens UNO de erros de importação/exportação passaram.")
