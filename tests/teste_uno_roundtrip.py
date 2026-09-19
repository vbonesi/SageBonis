# -*- coding: utf-8 -*-
"""Round-trip import/export via LibreOffice/UNO real: gera uma base .dat sintética
pequena (autocontida, sem depender de caminho externo), importa numa cópia descartável
do SageBonis.ods, exporta de volta e confere o resultado byte a byte.

É o guarda-corpo do que a ferramenta promete: o que entra tem que sair igual -- acento
em Latin-1, include ativo e comentado, bloco comentado, comentário do bloco, linhas
ignoradas (Gera=q e vazio), backup .bak na reexportação, e exportação parcial pela aba
ativa e pela lista da aba Geral.

Nasceu como teste de paridade entre as duas trilhas do projeto; depois que a variante
avançada virou ferramenta própria (19/09/2026), virou o teste de round-trip desta aqui.
Nunca escreve nos arquivos rastreados.

Requer 'soffice' no PATH. Roda com:
    python tests/teste_uno_roundtrip.py
"""
import filecmp
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ODS = os.path.join(_RAIZ, "SageBonis.ods")
PY = os.path.join(_RAIZ, "ImportadorSAGE.py")

falhas = []


def check(nome, cond):
    status = "OK" if cond else "FALHOU"
    print(f"[{status}] {nome}")
    if not cond:
        falhas.append(nome)


def escrever_dat(pasta, nome_arquivo, conteudo):
    with open(os.path.join(pasta, nome_arquivo), "w", encoding="latin-1") as f:
        f.write(conteudo)


def gerar_fixture(pasta):
    """Base .dat sintética mínima: 1 LSC + 1 CNF + PDS ativo/comentado,
    includes ativo/comentado (e acento, para testar Latin-1) + 2 PDF."""
    escrever_dat(pasta, "lsc.dat", """
LSC
   ID= TESTE1
 NOME= Canal de teste
  GSD= TST
  MAP= GERAL
NSRV1= localhost
NSRV2= localhost
  TCV= CNVM
 TIPO= AA
  TTP= CX104

""")
    escrever_dat(pasta, "cnf.dat", """
CNF
    ID= TESTE1
   LSC= TESTE1
CONFIG= PlPr= 1 LiPr= 1 PlRe= 0 LiRe= 0

""")
    escrever_dat(pasta, "pds.dat", """#include sub/ativo.dat
;#include sub/inativo.dat

; comentário preservado no bloco ativo
PDS
   ID= TESTE:DJ:POSICAO
 NOME= Disjuntor Posição Testé
  TAC= TESTE1

PDS
   ID= TESTE:SC:POSICAO
 NOME= Seccionadora Posicao
  TAC= TESTE1

;PDS
;   ID= TESTE:DJ:RESERVA
; NOME= Disjuntor reserva
;  TAC= TESTE1

""")
    escrever_dat(pasta, "pdf.dat", """
PDF
    ID= TESTE_IED.CTRL-XCBR1$ST$Pos
   PNT= TESTE:DJ:POSICAO
 TPPNT= PDS
 KCONV= DPS0

PDF
    ID= TESTE_IED.CTRL-XCBR2$ST$Pos
   PNT= TESTE:SC:POSICAO
 TPPNT= PDS
 KCONV= DPS0

""")


def rodar_import_export(porta, py_origem, pasta_entrada, pasta_saida, pasta_parcial,
                        pasta_parcial_lista, pasta_parcial_vazia, pasta_parcial_longa):
    with TesteUno(porta=porta, ods_origem=ODS, py_origem=py_origem) as t:
        t.definir_celula("Geral", 0, 3, pasta_entrada)  # A4
        t.chamar_macro("importar_dats")
        status_import = t.ler_celula("Geral", 1, 3)  # B4

        # Exercita codigos que nao surgem naturalmente do parser de .dat.
        linha = t.proxima_linha_livre("PDS")
        t.escrever_linha("PDS", linha, {
            "Origem": "pds.dat", "Gera": "n",
            "Comentario/Include": "Comentário solto — emoji 😀",
        })
        t.escrever_linha("PDS", linha + 1, {
            "Origem": "pds.dat", "Gera": "q", "ID": "NAO_EXPORTAR_Q",
        })
        t.escrever_linha("PDS", linha + 2, {
            "Origem": "pds.dat", "Gera": "", "ID": "NAO_EXPORTAR_VAZIO",
        })

        t.definir_celula("Geral", 0, 6, pasta_saida)  # A7
        t.chamar_macro("exportar_dats")
        status_export = t.ler_celula("Geral", 1, 6)  # B7
        # Segunda exportacao do mesmo conteudo deve criar backups dos .dat.
        t.chamar_macro("exportar_dats")
        status_reexport = t.ler_celula("Geral", 1, 6)  # B7

        t.definir_celula("Geral", 0, 6, pasta_parcial)  # A7
        t.ativar_aba("PDS")
        t.chamar_macro("exportar_parcial")
        status_parcial = t.ler_celula("Geral", 1, 6)  # B7

        # Pela aba Geral, a coluna C a partir da linha 14 define a lista; nomes
        # inexistentes sao ignorados.
        t.ativar_aba("Geral")
        for linha_geral in range(13, 144):
            t.definir_celula("Geral", 2, linha_geral, "")
        t.definir_celula("Geral", 2, 13, "PDS")
        t.definir_celula("Geral", 2, 14, "PDF")
        t.definir_celula("Geral", 2, 15, "NAO_EXISTE")
        t.definir_celula("Geral", 0, 6, pasta_parcial_lista)
        t.chamar_macro("exportar_parcial")
        status_parcial_lista = t.ler_celula("Geral", 1, 6)

        for linha_geral in range(13, 144):
            t.definir_celula("Geral", 2, linha_geral, "")
        t.definir_celula("Geral", 0, 6, pasta_parcial_vazia)
        t.chamar_macro("exportar_parcial")
        status_parcial_vazia = t.ler_celula("Geral", 1, 6)

        # Lista longa: a entidade fica na linha 201, bem depois do antigo limite fixo
        # de 130 linhas (C14:C144), que ignorava em silencio tudo dali pra baixo
        # (auditoria #12). Fica por ultimo de proposito: celula escrita fora da area
        # usada original do .ods nao volta a ficar vazia por setString("") neste
        # LibreOffice, entao ela contaminaria os cenarios seguintes.
        t.definir_celula("Geral", 2, 200, "PDF")
        t.definir_celula("Geral", 0, 6, pasta_parcial_longa)
        t.chamar_macro("exportar_parcial")
        status_parcial_longa = t.ler_celula("Geral", 1, 6)
    return (status_import, status_export, status_reexport, status_parcial,
            status_parcial_lista, status_parcial_vazia, status_parcial_longa)


pasta_entrada = tempfile.mkdtemp(prefix="sagebonis_entrada_")
pasta_saida = tempfile.mkdtemp(prefix="sagebonis_saida_")
pasta_parcial = tempfile.mkdtemp(prefix="sagebonis_parcial_")
pasta_lista = tempfile.mkdtemp(prefix="sagebonis_lista_")
pasta_vazia = tempfile.mkdtemp(prefix="sagebonis_vazia_")
pasta_longa = tempfile.mkdtemp(prefix="sagebonis_longa_")

try:
    gerar_fixture(pasta_entrada)

    (status_import, status_export, status_reexport, status_parcial,
     status_lista, status_vazia, status_longa) = rodar_import_export(
        2200, PY, pasta_entrada, pasta_saida, pasta_parcial,
        pasta_lista, pasta_vazia, pasta_longa)
    check("importação sem erro", "ERRO" not in status_import.upper())
    check("exportação sem erro", "ERRO" not in status_export.upper())
    check("reexportação sem erro", "ERRO" not in status_reexport.upper())
    check("exportação parcial sem erro", "ERRO" not in status_parcial.upper())

    arquivos = sorted(n for n in os.listdir(pasta_saida) if n.endswith(".dat"))
    check(f"exportou os .dat esperados ({arquivos})",
          arquivos == ["cnf.dat", "lsc.dat", "pdf.dat", "pds.dat"])

    caminho_pds = os.path.join(pasta_saida, "pds.dat")
    with open(caminho_pds, "rb") as f:
        pds_bytes = f.read()
    pds_texto = pds_bytes.decode("latin-1")
    check("round-trip: include ativo preservado", "#include sub/ativo.dat" in pds_texto)
    check("round-trip: include comentado preservado", ";#include sub/inativo.dat" in pds_texto)
    check("round-trip: bloco comentado preservado",
          ";PDS" in pds_texto and ";\tID = TESTE:DJ:RESERVA" in pds_texto)
    check("round-trip: comentário do bloco ativo preservado",
          ";comentário preservado no bloco ativo" in pds_texto)
    check("round-trip: saída contém o acento em Latin-1",
          "Posição Testé".encode("latin-1") in pds_bytes)
    check("round-trip: saída não codificou o acento em UTF-8",
          "Posição Testé".encode("utf-8") not in pds_bytes)
    check("exportação: comentário simples e Unicode foram sanitizados",
          ";Comentário solto - emoji ?" in pds_texto)
    check("exportação: linha Gera=q foi ignorada", "NAO_EXPORTAR_Q" not in pds_texto)
    check("exportação: linha Gera vazio foi ignorada", "NAO_EXPORTAR_VAZIO" not in pds_texto)
    caminho_backup = caminho_pds + ".bak"
    check("reexportação: backup .bak foi criado", os.path.isfile(caminho_backup))
    check("reexportação: backup preserva o conteúdo anterior",
          os.path.isfile(caminho_backup) and filecmp.cmp(caminho_pds, caminho_backup, shallow=False))

    check("exportação parcial: aba PDS gera somente pds.dat",
          sorted(os.listdir(pasta_parcial)) == ["pds.dat"])
    check("exportação parcial: conteúdo equivale ao pds.dat da exportação total",
          filecmp.cmp(caminho_pds, os.path.join(pasta_parcial, "pds.dat"), shallow=False))
    check("exportação parcial por lista: concluída sem erro",
          "ERRO" not in status_lista.upper())
    check("exportação parcial por lista: gera somente PDS e PDF",
          sorted(os.listdir(pasta_lista)) == ["pdf.dat", "pds.dat"])
    check("exportação parcial por lista: nome inexistente é ignorado",
          "nao_existe.dat" not in os.listdir(pasta_lista))
    check("exportação parcial: entidade listada além da linha 144 entra",
          sorted(os.listdir(pasta_longa)) == ["pdf.dat"])
    check("exportação parcial por lista longa: concluída sem erro",
          "ERRO" not in status_longa.upper())
    check("exportação parcial por lista vazia: emite aviso", "AVISO" in status_vazia.upper())
    check("exportação parcial por lista vazia: não cria arquivos", not os.listdir(pasta_vazia))

finally:
    for pasta in (pasta_entrada, pasta_saida, pasta_parcial, pasta_lista,
                  pasta_vazia, pasta_longa):
        shutil.rmtree(pasta, ignore_errors=True)

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    sys.exit(1)
print("Todas as checagens de round-trip import/export passaram.")
