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
    """Base .dat sintética mínima: 1 LSC + 1 CNF + 2 PDS (1 com acento, pra
    testar o round-trip Latin-1) + 2 PDF referenciando os PDS."""
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
    escrever_dat(pasta, "pds.dat", """
PDS
   ID= TESTE:DJ:POSICAO
 NOME= Disjuntor Posição Testé
  TAC= TESTE1

PDS
   ID= TESTE:SC:POSICAO
 NOME= Seccionadora Posicao
  TAC= TESTE1

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


def rodar_import_export(porta, py_origem, pasta_entrada, pasta_saida):
    with TesteUno(porta=porta, ods_origem=ODS_SIMPLES, py_origem=py_origem) as t:
        t.definir_celula("Geral", 0, 3, pasta_entrada)  # A4
        t.chamar_macro("importar_dats")
        status_import = t.ler_celula("Geral", 1, 3)  # B4

        t.definir_celula("Geral", 0, 6, pasta_saida)  # A7
        t.chamar_macro("exportar_dats")
        status_export = t.ler_celula("Geral", 1, 6)  # B7
    return status_import, status_export


pasta_entrada = tempfile.mkdtemp(prefix="sagebonis_paridade_entrada_")
pasta_saida_simples = tempfile.mkdtemp(prefix="sagebonis_paridade_simples_")
pasta_saida_completa = tempfile.mkdtemp(prefix="sagebonis_paridade_completa_")

try:
    gerar_fixture(pasta_entrada)

    status_import_s, status_export_s = rodar_import_export(
        2200, PY_SIMPLES, pasta_entrada, pasta_saida_simples)
    check("Simples: importação sem erro", "ERRO" not in status_import_s.upper())
    check("Simples: exportação sem erro", "ERRO" not in status_export_s.upper())

    status_import_c, status_export_c = rodar_import_export(
        2201, PY_COMPLETA, pasta_entrada, pasta_saida_completa)
    check("Completa: importação sem erro", "ERRO" not in status_import_c.upper())
    check("Completa: exportação sem erro", "ERRO" not in status_export_c.upper())

    arquivos_simples = sorted(os.listdir(pasta_saida_simples))
    arquivos_completa = sorted(os.listdir(pasta_saida_completa))
    check(f"mesmo conjunto de arquivos .dat exportados ({arquivos_simples})",
          arquivos_simples == arquivos_completa)

    _, mismatches, erros = filecmp.cmpfiles(
        pasta_saida_simples, pasta_saida_completa, arquivos_simples, shallow=False)
    check("nenhum arquivo .dat difere byte-a-byte entre Simples e Completa",
          not mismatches and not erros)
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

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    sys.exit(1)
print("Todas as checagens de paridade import/export passaram.")
