# -*- coding: utf-8 -*-
"""Smoke test do ROUND-TRIP importação -> exportação (lógica pura, sem LibreOffice).

Cada caso é um defeito que o round-trip tinha de verdade, medido no carregador do SAGE
(`AtualizaBD fria fonte`):

  1. .dat VAZIO referenciado por #include não era recriado: o carregador ABORTA a carga
     inteira se o arquivo incluído não existe.
  2. Duas abas alimentando o MESMO .dat (a do nome do arquivo + a da entidade dos blocos)
     gravavam uma por cima da outra: parte do conteúdo ia parar no .bak.
  3. Linha fora do formato (um rodapé sem ';', p.ex. "Arquivo gerado automaticamente por ...")
     sumia sem aviso.
  4. Registro com ';ENTIDADE' (comentário) no meio: o carregador o trata como UM registro; o
     parser o partia em um bloco sem atributos (rejeitado) + atributos órfãos.
  5. #include de arquivo que não termina em '.dat' (ocr.dat.novo, pdd.dat.rgv) nunca era
     importado, e o #include passava a apontar para um arquivo inexistente.
  6. Atributo repetido no bloco: o dict ficava só com o último e o anterior sumia.
  7. Espaço não-quebrável (válido em Latin-1) virava espaço comum.
  8. Ordem de colunas dependia da ordem de um set (muda a cada execução).
  9. Arquivo 'grcmp .dat' (espaço no nome) virava a aba 'grcmp ' que sobrescrevia a 'grcmp'.
 10. Bloco de uma entidade num arquivo de pasta sem o .dat dessa entidade virava órfão.

A régua é a GRAMÁTICA DO CARREGADOR (tests/gramatica_sage.py): dois textos que dão os mesmos
registros são a mesma base para o SAGE, mesmo que o arquivo mude de formato.

Roda com: python tests/smoke_test_roundtrip.py
"""
import importlib.util
import os
import re
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gramatica_sage import ler_latin1, registros_do_carregador  # noqa: E402

RAIZ_PROJETO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

falhas = []


def check(nome, cond):
    print(f"[{'OK' if cond else 'FALHOU'}] {nome}")
    if not cond:
        falhas.append(nome)


def carregar():
    # IMPORTADORSAGE_PY permite apontar para uma cópia alterada da macro (teste de mutação: cada
    # defeito corrigido tem que fazer ao menos uma checagem daqui falhar quando reintroduzido)
    caminho = os.environ.get("IMPORTADORSAGE_PY") or os.path.join(RAIZ_PROJETO, "ImportadorSAGE.py")
    spec = importlib.util.spec_from_file_location("mod_roundtrip", caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod._log_importacao = lambda *args, **kwargs: None  # os avisos (force=True) poluem a saída
    return mod


# ---------------------------------------------------------------------------
# "Planilha" mínima: o que _exportar_folhas lê (getName/createCursor/getDataArray)
# ---------------------------------------------------------------------------
class _Faixa:
    def __init__(self, linhas):
        self.linhas = linhas

    def getDataArray(self):
        return tuple(tuple(r) for r in self.linhas)


class _Cursor:
    def __init__(self, n_linhas, n_colunas):
        self.n_linhas, self.n_colunas = n_linhas, n_colunas

    def gotoEndOfUsedArea(self, _):
        pass

    def getRangeAddress(self):
        class _Endereco:
            pass
        e = _Endereco()
        e.EndRow, e.EndColumn = self.n_linhas - 1, self.n_colunas - 1
        return e


class FolhaFalsa:
    def __init__(self, nome, matriz):
        self.nome, self.matriz = nome, matriz

    def getName(self):
        return self.nome

    def createCursor(self):
        return _Cursor(len(self.matriz), len(self.matriz[0]))

    def getCellRangeByPosition(self, c0, r0, c1, r1):
        return _Faixa([linha[c0:c1 + 1] for linha in self.matriz[r0:r1 + 1]])


def round_trip(mod, origem, destino):
    """importação (_ler_base) -> planilha (_montar_matriz_aba) -> exportação (_exportar_folhas)."""
    dados = mod._ler_base(origem)
    folhas = [FolhaFalsa(ent, mod._montar_matriz_aba(pontos, [])[1]) for ent, pontos in dados.items()]
    erros, substituicoes, sem_origem = mod._exportar_folhas(folhas, destino)
    return dados, erros


def escrever(pasta, rel, texto, binario=None):
    caminho = os.path.join(pasta, rel)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "wb") as f:
        f.write(binario if binario is not None else texto.encode("latin-1"))


def main():
    mod = carregar()
    base = tempfile.mkdtemp(prefix="sagebonis_rt_")
    try:
        origem, destino = os.path.join(base, "orig"), os.path.join(base, "saida")
        os.makedirs(destino)

        # --- a base de teste ----------------------------------------------------------
        escrever(origem, "pds.dat",
                 "#include A/pds.dat\n#include vazio/pds.dat\n#include B/pdf_1.dat\n"
                 ";#include comentado/pds.dat\n")
        escrever(origem, "A/pds.dat",
                 ";cabeçalho do arquivo\nPDS\n\tID = P1\n\tNOME = Posição Testé\n"
                 "\tOBSRV = com nbsp\n\tNOME = Nome repetido\n\n"
                 "PDS\n\tID = P2\n\tTAC = T1\n\tNOME = termina com nbsp\u00a0\n"
                 "\nArquivo gerado automaticamente por gerador-exemplo versao 1.0.")
        escrever(origem, "vazio/pds.dat", "\n\n")
        # (2) arquivo com cabeçalho/#include + blocos de OUTRA entidade que não a do nome
        escrever(origem, "B/pdf_1.dat", "#include B/outro/pdf.dat\nPDF\n\tID = F1\n\tPNT = P1\n")
        escrever(origem, "B/outro/pdf.dat", "PDF\n\tID = F2\n\tPNT = P2\n")
        # (4) tag ativa + ';ENTIDADE' comentado + atributos soltos = UM registro para o carregador
        escrever(origem, "e2m.dat",
                 "E2M\n; Planilha: X\n;E2M\n;\tIDPTO = ANTIGO\nIDPTO = P1\nMAP = TRIPS\nTIPO = PDS\n\n"
                 "E2M\n\tIDPTO = P2\n\tMAP = IMPED\n\tTIPO = PDS\n")
        # (5) include de arquivo que não termina em .dat; e um backup que NÃO é incluído
        escrever(origem, "ocr.dat", "#include SAGE/ocr.dat.novo\nRodape fora de bloco, sem ponto e virgula\n")
        escrever(origem, "SAGE/ocr.dat.novo", "OCR\n\tID = OCR1\n\tTEXTO = ok\n")
        escrever(origem, "SAGE/ocr.dat.bak", "OCR\n\tID = OCR_VELHO\n")
        # (9) espaço no nome do arquivo
        escrever(origem, "grcmp.dat", "GRCMP\n\tGRUPO = G1\n\tPNT = P1\n")
        escrever(origem, "X/grcmp .dat", "GRCMP\n\tGRUPO = G2\n\tPNT = P2\n")
        escrever(origem, "Z/grcmp .dat", "\n")  # vazio: a linha-marca cai na aba do nome do arquivo
        # (10) bloco CNF num arquivo de pasta que não tem cnf.dat
        escrever(origem, "cnf.dat", "CNF\n\tID = C1\n")
        escrever(origem, "Y/lsc.dat", "LSC\n\tID = L1\n\nCNF\n\tID = C2\n\tLSC = L1\n")
        # UTF-8 (a macro importa utf-8 e exporta latin-1)
        escrever(origem, "tac.dat", None, binario="TAC\n\tID = T1\n\tNOME = Ligação\n".encode("utf-8"))

        dados, erros = round_trip(mod, origem, destino)
        check("round-trip termina sem erro", not erros)

        # ===== (1) arquivo vazio =======================================================
        check("(1) arquivo vazio referenciado por #include é recriado",
              os.path.isfile(os.path.join(destino, "vazio", "pds.dat")))
        check("(1) e o carregador o lê como arquivo sem registros",
              registros_do_carregador(ler_latin1(os.path.join(destino, "vazio", "pds.dat"))) == [])
        check("(1) o contador avisa o usuário (2 arquivos vazios na base de teste)", mod.ACUMULADO_IMPORTACAO["arquivo_vazio"] == 2)

        # ===== (2) duas abas, um arquivo ===============================================
        check("(2) nenhum .bak criado pela exportação",
              not [f for _, _, fs in os.walk(destino) for f in fs if f.endswith(".bak")])
        regs_pdf1 = registros_do_carregador(ler_latin1(os.path.join(destino, "B", "pdf_1.dat")))
        check("(2) o arquivo das duas abas sai inteiro: #include + bloco",
              ("#include", "B/outro/pdf.dat") in regs_pdf1
              and ("PDF", (("ID", "F1"), ("PNT", "P1"))) in regs_pdf1)

        # ===== (3) linha fora do formato ===============================================
        texto_a = ler_latin1(os.path.join(destino, "A", "pds.dat"))
        check("(3) rodapé sem ';' é preservado (como comentário)",
              "Arquivo gerado automaticamente por gerador-exemplo" in texto_a)
        check("(3) e vira comentário (o carregador não o lê como registro/atributo)",
              all(l.startswith(";") for l in texto_a.splitlines()
                  if "Arquivo gerado automaticamente" in l))
        check("(3) o contador avisa o usuário", mod.ACUMULADO_IMPORTACAO["preservadas"] >= 2)
        texto_ocr = ler_latin1(os.path.join(destino, "ocr.dat"))
        check("(3) linha fora do formato FORA de bloco também é preservada, como comentário",
              ";Rodape fora de bloco, sem ponto e virgula" in texto_ocr)

        # ===== (4) registro com ';ENTIDADE' no meio ====================================
        regs_e2m = [r for r in registros_do_carregador(ler_latin1(os.path.join(destino, "e2m.dat")))
                    if r[0] != "#include"]
        check("(4) o registro com ';E2M' no meio continua UM registro",
              len(regs_e2m) == 2)
        check("(4) com os atributos vivos (IDPTO/MAP/TIPO), sem bloco vazio",
              ("E2M", (("IDPTO", "P1"), ("MAP", "TRIPS"), ("TIPO", "PDS"))) in regs_e2m)
        check("(4) e o carregador não vê nenhum registro sem atributo",
              all(r[1] for r in regs_e2m))

        # ===== (5) include que não termina em .dat =====================================
        check("(5) ocr.dat.novo (incluído) foi importado e exportado no mesmo caminho",
              os.path.isfile(os.path.join(destino, "SAGE", "ocr.dat.novo")))
        check("(5) o conteúdo está lá",
              ("OCR", (("ID", "OCR1"), ("TEXTO", "ok"))) in registros_do_carregador(
                  ler_latin1(os.path.join(destino, "SAGE", "ocr.dat.novo"))))
        check("(5) .dat.bak (não incluído) NÃO é importado",
              not os.path.exists(os.path.join(destino, "SAGE", "ocr.dat.bak")))

        # ===== (6) atributo repetido ===================================================
        regs_a = registros_do_carregador(texto_a)
        check("(6) atributo repetido: vale o último no registro",
              ("PDS", (("ID", "P1"), ("NOME", "Nome repetido"), ("OBSRV", "com nbsp"))) in regs_a)
        check("(6) o valor anterior não some: fica como comentário",
              "NOME = Posição Testé" in texto_a)
        check("(6) o contador avisa o usuário", mod.ACUMULADO_IMPORTACAO["repetidos"] >= 1)

        # ===== (7) espaço não-quebrável ================================================
        check("(7) espaço não-quebrável é preservado (existe em Latin-1)", "com nbsp" in texto_a)
        check("(7) inclusive na borda do valor (o carregador o mantém; str.strip() o comia)",
              "NOME = termina com nbsp\u00a0\n" in texto_a)

        # ===== (9) arquivo com espaço no nome ==========================================
        check("(9) 'grcmp .dat' (espaço no nome) sai no mesmo caminho",
              os.path.isfile(os.path.join(destino, "X", "grcmp .dat")))
        check("(9) e não derruba o grcmp.dat da raiz",
              ("GRCMP", (("GRUPO", "G1"), ("PNT", "P1"))) in registros_do_carregador(
                  ler_latin1(os.path.join(destino, "grcmp.dat"))))
        check("(9) o arquivo vazio com espaço no nome também sai",
              os.path.isfile(os.path.join(destino, "Z", "grcmp .dat")))
        check("(9) nenhuma aba com espaço no nome", all(k == k.strip() for k in dados))

        # ===== (10) entidade em arquivo de pasta sem o .dat dela ======================
        regs_lsc = registros_do_carregador(ler_latin1(os.path.join(destino, "Y", "lsc.dat")))
        check("(10) bloco CNF no lsc.dat de pasta sem cnf.dat é bloco, não órfão",
              ("CNF", (("ID", "C2"), ("LSC", "L1"))) in regs_lsc)

        # ===== encoding ================================================================
        bruto = open(os.path.join(destino, "tac.dat"), "rb").read()
        check("arquivo em UTF-8 sai em Latin-1 (acento = 1 byte)", b"Liga\xe7\xe3o" in bruto)

        # ===== (8) ordem de colunas determinística =====================================
        cab1 = mod._montar_matriz_aba(dados["pds"], [])[0]
        cab2 = mod._montar_matriz_aba(list(reversed(dados["pds"])), ["TAC"])[0]
        check("(8) colunas: ordem de primeira aparição (determinística)",
              cab1[3:] == ["ID", "NOME", "OBSRV", "TAC"])
        check("(8) colunas: a prioridade da MaisUsadas continua mandando",
              cab2[3] == "TAC")

        # ===== propriedades gerais =====================================================
        # Todo registro que o carregador lia na original está na saída (e vice-versa).
        def todos_registros(pasta):
            res = []
            for a, _, fs in os.walk(pasta):
                for f in fs:
                    if re.search(r"\.dat(\.[^.]+)?$", f) and not f.endswith(".bak") and "OCR_VELHO" not in ler_latin1(os.path.join(a, f)):
                        rel = os.path.relpath(os.path.join(a, f), pasta)
                        res.extend((rel, r) for r in registros_do_carregador(ler_latin1(os.path.join(a, f))))
            return sorted(res, key=repr)

        # a original em UTF-8 vira Latin-1 de propósito: compara decodificando cada lado do seu jeito
        def normaliza(rs):
            return [(rel, r) for rel, r in rs if rel != "tac.dat"]

        orig_regs, saida_regs = normaliza(todos_registros(origem)), normaliza(todos_registros(destino))
        so_orig = [x for x in orig_regs if x not in saida_regs]
        so_saida = [x for x in saida_regs if x not in orig_regs]
        check("todos os registros e includes da original estão na saída (visão do carregador)",
              not so_orig and not so_saida)
        if so_orig or so_saida:
            print("   só na original:", so_orig[:4])
            print("   só na saída:", so_saida[:4])

        # idempotência: importar a saída e exportar de novo dá os mesmos bytes
        destino2 = os.path.join(base, "saida2")
        os.makedirs(destino2)
        round_trip(mod, destino, destino2)
        mudou = []
        for a, _, fs in os.walk(destino):
            for f in fs:
                rel = os.path.relpath(os.path.join(a, f), destino)
                p2 = os.path.join(destino2, rel)
                if not os.path.exists(p2) or open(os.path.join(a, f), "rb").read() != open(p2, "rb").read():
                    mudou.append(rel)
        check("o round-trip é estável: exportar a saída de novo não muda nenhum byte", not mudou)
        if mudou:
            print("   mudaram:", mudou[:5])
    finally:
        shutil.rmtree(base, ignore_errors=True)

    if falhas:
        print(f"\n{len(falhas)} checagem(ns) falharam.")
        raise SystemExit(1)
    print("Todas as checagens do smoke test passaram.")


if __name__ == "__main__":
    main()
