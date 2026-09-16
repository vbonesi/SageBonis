# -*- coding: utf-8 -*-
"""Verifica o critério de maturidade "paridade de import/export com a Trilha
Simples" (PLANEJAMENTO.md): confirma que completa/ImportadorSAGE.py ainda
produz o MESMO resultado de .dat que a raiz (ImportadorSAGE.py da Trilha
Simples) para o núcleo compartilhado de importar_dats/exportar_dats.

Método: gera uma base .dat sintética pequena e autocontida (não depende de
nenhum caminho externo), importa ela em 2 cópias descartáveis do MESMO
SageBonis.ods em branco da raiz -- uma rodando o ImportadorSAGE.py da raiz
(Simples) sem alterações, outra com o ImportadorSAGE.py da Completa injetado
-- exporta as duas de volta pra pastas separadas, e faz diff byte-a-byte dos
.dat resultantes. Nunca escreve nos arquivos rastreados.

Requer 'soffice' no PATH. Roda com:
    python completa/tests/teste_paridade_import_export.py
"""
import filecmp
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

_RAIZ_COMPLETA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_RAIZ_PROJETO = os.path.dirname(_RAIZ_COMPLETA)
ODS_SIMPLES = os.path.join(_RAIZ_PROJETO, "SageBonis.ods")
PY_SIMPLES = os.path.join(_RAIZ_PROJETO, "ImportadorSAGE.py")
PY_COMPLETA = os.path.join(_RAIZ_COMPLETA, "ImportadorSAGE.py")

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


def rodar_import_export(porta, py_origem, pasta_entrada, pasta_saida, pasta_parcial):
    with TesteUno(porta=porta, ods_origem=ODS_SIMPLES, py_origem=py_origem) as t:
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
    return status_import, status_export, status_reexport, status_parcial


pasta_entrada = tempfile.mkdtemp(prefix="sagebonis_paridade_entrada_")
pasta_saida_simples = tempfile.mkdtemp(prefix="sagebonis_paridade_simples_")
pasta_saida_completa = tempfile.mkdtemp(prefix="sagebonis_paridade_completa_")
pasta_parcial_simples = tempfile.mkdtemp(prefix="sagebonis_parcial_simples_")
pasta_parcial_completa = tempfile.mkdtemp(prefix="sagebonis_parcial_completa_")

try:
    gerar_fixture(pasta_entrada)

    status_import_s, status_export_s, status_reexport_s, status_parcial_s = rodar_import_export(
        2200, PY_SIMPLES, pasta_entrada, pasta_saida_simples, pasta_parcial_simples)
    check("Simples: importação sem erro", "ERRO" not in status_import_s.upper())
    check("Simples: exportação sem erro", "ERRO" not in status_export_s.upper())
    check("Simples: reexportação sem erro", "ERRO" not in status_reexport_s.upper())
    check("Simples: exportação parcial sem erro", "ERRO" not in status_parcial_s.upper())

    status_import_c, status_export_c, status_reexport_c, status_parcial_c = rodar_import_export(
        2201, PY_COMPLETA, pasta_entrada, pasta_saida_completa, pasta_parcial_completa)
    check("Completa: importação sem erro", "ERRO" not in status_import_c.upper())
    check("Completa: exportação sem erro", "ERRO" not in status_export_c.upper())
    check("Completa: reexportação sem erro", "ERRO" not in status_reexport_c.upper())
    check("Completa: exportação parcial sem erro", "ERRO" not in status_parcial_c.upper())

    arquivos_simples = sorted(n for n in os.listdir(pasta_saida_simples) if n.endswith(".dat"))
    arquivos_completa = sorted(n for n in os.listdir(pasta_saida_completa) if n.endswith(".dat"))
    check(f"mesmo conjunto de arquivos .dat exportados ({arquivos_simples})",
          arquivos_simples == arquivos_completa)

    _, mismatches, erros = filecmp.cmpfiles(
        pasta_saida_simples, pasta_saida_completa, arquivos_simples, shallow=False)
    check("nenhum arquivo .dat difere byte-a-byte entre Simples e Completa",
          not mismatches and not erros)

    # Paridade sozinha poderia esconder um erro comum às duas variantes. Estes
    # checks confirmam também o conteúdo funcional esperado do round-trip.
    caminho_pds = os.path.join(pasta_saida_completa, "pds.dat")
    with open(caminho_pds, "rb") as f:
        pds_bytes = f.read()
    pds_texto = pds_bytes.decode("latin-1")
    check("round-trip: include ativo preservado", "#include sub/ativo.dat" in pds_texto)
    check("round-trip: include comentado preservado", ";#include sub/inativo.dat" in pds_texto)
    check("round-trip: bloco comentado preservado", ";PDS" in pds_texto and ";\tID = TESTE:DJ:RESERVA" in pds_texto)
    check("round-trip: comentário do bloco ativo preservado", ";comentário preservado no bloco ativo" in pds_texto)
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

    arquivos_parciais_s = sorted(os.listdir(pasta_parcial_simples))
    arquivos_parciais_c = sorted(os.listdir(pasta_parcial_completa))
    check("exportação parcial: aba PDS gera somente pds.dat",
          arquivos_parciais_s == ["pds.dat"] and arquivos_parciais_c == ["pds.dat"])
    check("exportação parcial: Simples e Completa geram o mesmo pds.dat",
          filecmp.cmp(os.path.join(pasta_parcial_simples, "pds.dat"),
                      os.path.join(pasta_parcial_completa, "pds.dat"), shallow=False))
    check("exportação parcial: conteúdo equivale ao pds.dat da exportação total",
          filecmp.cmp(caminho_pds, os.path.join(pasta_parcial_completa, "pds.dat"), shallow=False))
    if mismatches:
        for nome in mismatches:
            print(f"  DIFF em {nome}:")
            with open(os.path.join(pasta_saida_simples, nome), encoding="latin-1") as f1, \
                 open(os.path.join(pasta_saida_completa, nome), encoding="latin-1") as f2:
                print("    Simples: ", f1.read())
                print("    Completa:", f2.read())

finally:
    shutil.rmtree(pasta_entrada, ignore_errors=True)
    shutil.rmtree(pasta_saida_simples, ignore_errors=True)
    shutil.rmtree(pasta_saida_completa, ignore_errors=True)
    shutil.rmtree(pasta_parcial_simples, ignore_errors=True)
    shutil.rmtree(pasta_parcial_completa, ignore_errors=True)

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    sys.exit(1)
print("Todas as checagens de paridade import/export passaram.")
