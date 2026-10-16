# -*- coding: utf-8 -*-
"""Round-trip importar -> exportar no LibreOffice REAL, com as macros embutidas no .ods.

Cobre o que só o LibreOffice real mostra (a lógica pura está em smoke_test_roundtrip.py):

  * a importação TOTAL esvazia as abas de entidade que sobraram de uma base anterior (o .ods
    distribuído traz dados de demonstração; sem isso a exportação gravava arquivos de uma
    base dentro da outra);
  * 'grcmp .dat' (espaço no nome) NÃO sobrescreve a aba 'grcmp' (o _get_sheet ignora o
    espaço nas pontas);
  * valores que o Calc poderia interpretar (fórmula, número, data, texto longo) saem
    exatamente como entraram;
  * a exportação PARCIAL de uma aba não trunca o .dat que outra aba também alimenta;
  * exportar duas vezes dá os mesmos bytes (colunas em ordem determinística);
  * o status da importação conta ao usuário o que foi tratado de um jeito diferente.
"""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gramatica_sage import ler_latin1, registros_do_carregador  # noqa: E402
from uno_harness import TesteUno  # noqa: E402

falhas = []


def check(nome, cond):
    print("[%s] %s" % ("OK" if cond else "FALHOU", nome))
    if not cond:
        falhas.append(nome)


def escrever(pasta, rel, texto):
    caminho = os.path.join(pasta, rel)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "wb") as f:
        f.write(texto.encode("latin-1"))


# Valores que o Calc tem motivo para interpretar (fórmula, número, data, texto longo...).
VALORES = {
    "V01": "=SUM(1+1)", "V02": "+5", "V03": "-7", "V04": "@x", "V05": "0012", "V06": "1E5",
    "V07": "TRUE", "V08": "1,5", "V09": "12/05/2024", "V10": "'aspas", "V11": "00:30",
    "V12": "1.0", "V13": "007.50", "V14": "=== X ===", "V15": "#REF!", "V16": "x" * 40000,
    "V17": "2024-01-02", "V18": "1/2", "V19": "0x1F", "V20": "a  b", "V21": "=cmd|calc",
    "V22": "100%", "V23": "1e-3", "V24": "--", "V25": "=", "V26": "com nbsp",
}

pasta_base = tempfile.mkdtemp(prefix="sagebonis_rtuno_base_")
pasta_saida = tempfile.mkdtemp(prefix="sagebonis_rtuno_saida_")
pasta_saida2 = tempfile.mkdtemp(prefix="sagebonis_rtuno_saida2_")
pasta_parcial = tempfile.mkdtemp(prefix="sagebonis_rtuno_parcial_")

try:
    linhas_pds = ["PDS"] + ["\t%s = %s" % (k, v) for k, v in VALORES.items()]
    escrever(pasta_base, "pds.dat", "#include A/pds.dat\n#include vazio/pds.dat\n")
    escrever(pasta_base, "A/pds.dat", "\n".join(linhas_pds) + "\n\nPDS\n\tID = P2\n")
    escrever(pasta_base, "vazio/pds.dat", "\n")
    # um .dat alimentado por DUAS abas: a do nome do arquivo (pdf_1: o #include) e a da entidade (pdf)
    escrever(pasta_base, "pdf_1.dat", "#include B/pdf.dat\nPDF\n\tID = F1\n\tPNT = P2\n")
    escrever(pasta_base, "B/pdf.dat", "PDF\n\tID = F2\n\tPNT = P2\n")
    # espaço no nome do arquivo
    escrever(pasta_base, "grcmp.dat", "GRCMP\n\tGRUPO = G1\n\tPNT = P2\n")
    escrever(pasta_base, "Z/grcmp .dat", "\n")
    escrever(pasta_base, "ocr.dat", "#include SAGE/ocr.dat.novo\n")
    # um .dat com o nome de uma aba de configuração não pode escrever em cima dela
    escrever(pasta_base, "geral.dat", "GERAL\n\tID = NAO_PODE\n")
    escrever(pasta_base, "SAGE/ocr.dat.novo", "OCR\n\tID = OCR1\n\tTEXTO = ok\n")

    with TesteUno(porta=2500) as t:
        # uma aba de entidade com dados de uma base anterior (a planilha distribuída vem sem nenhuma)
        t.doc.Sheets.insertNewByName("aba_velha", t.doc.Sheets.getCount())
        t.definir_celula("aba_velha", 0, 0, "Origem")
        t.definir_celula("aba_velha", 1, 0, "Gera")
        t.definir_celula("aba_velha", 2, 0, "Comentario/Include")
        t.definir_celula("aba_velha", 0, 1, "velha.dat")
        t.definir_celula("aba_velha", 1, 1, "n")
        t.definir_celula("aba_velha", 2, 1, "lixo de uma base anterior")
        t.definir_celula("Geral", 0, 3, pasta_base)
        t.definir_celula("Geral", 0, 6, pasta_saida)
        t.chamar_macro("importar_dats")
        status_imp = t.ler_celula("Geral", 1, 3)
        check("importação total conclui", "concluída com sucesso" in status_imp)
        check("'geral.dat' (nome de aba de config) é ignorado e avisado no status",
              "geral" in status_imp and "NÃO importada" in status_imp)
        check("o status conta os arquivos vazios marcados", "arquivo(s) vazio(s)" in status_imp)
        check("o status conta as abas esvaziadas", "esvaziadas" in status_imp)
        check("aba de base anterior foi esvaziada (sem lixo na exportação)",
              t.contar_linhas("aba_velha") == 0)
        check("'Geral' (config) NÃO é tocada pelo esvaziamento",
              t.ler_celula("Geral", 0, 3) == pasta_base)
        check("não nasce aba 'grcmp ' (com espaço): o arquivo 'grcmp .dat' cai na 'grcmp'",
              not any(n != n.strip() for n in t.doc.Sheets.ElementNames))
        check("o bloco do grcmp.dat da raiz continua na aba grcmp (não foi sobrescrito)",
              any(r.get("GRUPO") == "G1" for r in t.ler_aba("grcmp")))

        t.chamar_macro("exportar_dats")
        check("exportação total conclui", "concluída com sucesso" in t.ler_celula("Geral", 1, 6))

        # ----- o que saiu, na visão do carregador -----
        texto_pds = ler_latin1(os.path.join(pasta_saida, "A", "pds.dat"))
        regs = dict((r[0], r[1]) for r in registros_do_carregador(texto_pds) if r[0] == "PDS" and r[1] and dict(r[1]).get("V01"))
        atributos = dict(regs.get("PDS", ()))
        check("os 26 valores 'perigosos' saem exatamente como entraram",
              all(atributos.get(k) == v for k, v in VALORES.items()))
        if not all(atributos.get(k) == v for k, v in VALORES.items()):
            for k, v in VALORES.items():
                if atributos.get(k) != v:
                    print("   DIFERE", k, repr(v[:30]), "->", repr(str(atributos.get(k))[:30]))
        check("arquivo vazio incluído é recriado",
              os.path.isfile(os.path.join(pasta_saida, "vazio", "pds.dat")))
        check("'grcmp .dat' (espaço no nome) é recriado",
              os.path.isfile(os.path.join(pasta_saida, "Z", "grcmp .dat")))
        check("ocr.dat.novo (incluído, não termina em .dat) é recriado",
              os.path.isfile(os.path.join(pasta_saida, "SAGE", "ocr.dat.novo")))
        regs_pdf1 = registros_do_carregador(ler_latin1(os.path.join(pasta_saida, "pdf_1.dat")))
        check("arquivo de duas abas (pdf_1.dat) sai inteiro: #include + bloco",
              ("#include", "B/pdf.dat") in regs_pdf1 and ("PDF", (("ID", "F1"), ("PNT", "P2"))) in regs_pdf1)
        check("nenhum .bak criado numa exportação para pasta nova",
              not [f for _, _, fs in os.walk(pasta_saida) for f in fs if f.endswith(".bak")])
        esperados = {"pds.dat", "A/pds.dat", "vazio/pds.dat", "pdf_1.dat", "B/pdf.dat", "grcmp.dat",
                     "Z/grcmp .dat", "ocr.dat", "SAGE/ocr.dat.novo"}
        gravados = {os.path.relpath(os.path.join(a, f), pasta_saida)
                    for a, _, fs in os.walk(pasta_saida) for f in fs}
        check("nenhum arquivo da base de demonstração vaza para a saída", gravados == esperados)
        if gravados != esperados:
            print("   a mais:", sorted(gravados - esperados)[:5], "a menos:", sorted(esperados - gravados)[:5])

        # ----- determinismo: exportar de novo (e reimportar) dá os mesmos bytes -----
        t.definir_celula("Geral", 0, 6, pasta_saida2)
        t.chamar_macro("exportar_dats")
        mudou = []
        for a, _, fs in os.walk(pasta_saida):
            for f in fs:
                rel = os.path.relpath(os.path.join(a, f), pasta_saida)
                p2 = os.path.join(pasta_saida2, rel)
                if not os.path.exists(p2) or open(os.path.join(a, f), "rb").read() != open(p2, "rb").read():
                    mudou.append(rel)
        check("exportar duas vezes dá os mesmos bytes", not mudou)

        # ----- exportação parcial de UMA aba não trunca o arquivo que outra também alimenta -----
        t.definir_celula("Geral", 0, 6, pasta_parcial)
        t.ativar_aba("pdf")
        t.chamar_macro("exportar_parcial")
        regs_parcial = registros_do_carregador(ler_latin1(os.path.join(pasta_parcial, "pdf_1.dat")))
        check("exportação parcial da aba 'pdf' mantém o #include que vem da aba 'pdf_1'",
              ("#include", "B/pdf.dat") in regs_parcial and ("PDF", (("ID", "F1"), ("PNT", "P2"))) in regs_parcial)
        check("exportação parcial grava só os arquivos da aba exportada",
              not os.path.exists(os.path.join(pasta_parcial, "A", "pds.dat")))

        # ----- importação PARCIAL (aba ativa): reimporta só a entidade, sem esvaziar as outras -----
        linhas_pdf_antes = t.contar_linhas("pdf")
        t.ativar_aba("pds")
        t.chamar_macro("importar_parcial")
        status_parcial_imp = t.ler_celula("Geral", 1, 3)
        check("importação parcial da aba ativa conclui", "Importação parcial concluída" in status_parcial_imp)
        check("importação parcial segue o #include de arquivo fora de .dat só da entidade pedida",
              t.contar_linhas("pdf") == linhas_pdf_antes and t.contar_linhas("ocr") >= 1)
        check("importação parcial não esvazia as outras abas", "esvaziadas" not in status_parcial_imp)
        check("importação parcial da 'pds' devolve o valor 'perigoso' intacto",
              any(r.get("V01") == "=SUM(1+1)" for r in t.ler_aba("pds")))
finally:
    for p in (pasta_base, pasta_saida, pasta_saida2, pasta_parcial):
        shutil.rmtree(p, ignore_errors=True)

print()
if falhas:
    print("%d checagem(ns) FALHOU/FALHARAM: %s" % (len(falhas), falhas))
    raise SystemExit(1)
print("Todas as checagens UNO de round-trip passaram.")
