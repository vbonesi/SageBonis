# -*- coding: utf-8 -*-
"""Smoke test do parser de .dat (lógica pura, sem UNO) -- o componente mais
arriscado do projeto e o único que não tinha teste nenhum (auditoria #7).

Roda as MESMAS checagens contra as duas trilhas (raiz e completa/), porque
parse_dat_file/_classificar_linha_dat/_finalizar_bloco são idênticos nas duas e
precisam continuar assim.

Cobre também as regressões dos achados já corrigidos:
  - #4: entidade com dígito em bloco comentado (';E2M') tem que virar bloco
        comentado, não comentário solto (que seria perdido).
  - #5: .dat gravado em utf-8 tem que importar sem mojibake (utf-8 é tentado
        antes do latin-1).

Roda com: python completa/tests/smoke_test_parser.py
"""
import importlib.util
import os
import shutil
import tempfile

RAIZ_PROJETO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TRILHAS = [
    ("simples", os.path.join(RAIZ_PROJETO, "ImportadorSAGE.py")),
    ("completa", os.path.join(RAIZ_PROJETO, "completa", "ImportadorSAGE.py")),
]

falhas = []


def check(nome, cond):
    status = "OK" if cond else "FALHOU"
    print(f"[{status}] {nome}")
    if not cond:
        falhas.append(nome)


def carregar(caminho):
    spec = importlib.util.spec_from_file_location("mod_parser", caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    # O parser loga cada arquivo processado; silencia pra saída do teste ficar legível.
    mod.LOG_IMPORTACAO_RESUMO = False
    return mod


def escrever(pasta, nome, texto, encoding="latin-1", binario=None):
    caminho = os.path.join(pasta, nome)
    if binario is not None:
        with open(caminho, "wb") as f:
            f.write(binario)
    else:
        with open(caminho, "w", encoding=encoding) as f:
            f.write(texto)
    return caminho


def parsear(mod, caminho, nome_relativo, entidades_validas):
    all_data = {}
    mod.parse_dat_file(caminho, nome_relativo, all_data, entidades_validas)
    return all_data


def rodar(trilha, caminho_modulo):
    mod = carregar(caminho_modulo)
    pasta = tempfile.mkdtemp(prefix="sagebonis_parser_")
    p = lambda nome: f"{trilha}: {nome}"
    try:
        # ----------------------------------------------------------
        # 1. Bloco ativo + comentário de cabeçalho + acento latin-1
        # ----------------------------------------------------------
        arq = escrever(pasta, "pds.dat",
                       ";Comentario do ponto\n"
                       "PDS\n"
                       "\tID = P1\n"
                       "\tNOME = Posição Testé\n")
        dados = parsear(mod, arq, "pds.dat", {"PDS", "PDF", "E2M"})
        pds = dados.get("pds", [])
        check(p("bloco ativo: 1 ponto em pds"), len(pds) == 1)
        check(p("bloco ativo: tipo 'x'"), pds and pds[0]["type"] == mod.CODIGO_BLOCO_ATIVO)
        check(p("bloco ativo: atributos ID/NOME"),
              pds and pds[0]["attributes"] == {"ID": "P1", "NOME": "Posição Testé"})
        check(p("bloco ativo: comentário anterior vira comment do bloco"),
              pds and pds[0].get("comment") == "Comentario do ponto")
        check(p("latin-1: acento preservado sem mojibake"),
              pds and pds[0]["attributes"].get("NOME") == "Posição Testé")

        # ----------------------------------------------------------
        # 2. Bloco comentado com dígito no nome (auditoria #4)
        # ----------------------------------------------------------
        arq = escrever(pasta, "e2m.dat",
                       ";E2M\n"
                       ";\tID = E1\n"
                       ";\tNOME = Comentado\n")
        dados = parsear(mod, arq, "e2m.dat", {"E2M", "GR2ACT", "PDS"})
        e2m = dados.get("e2m", [])
        check(p("#4 regressão: ';E2M' vira bloco comentado (não comentário solto)"),
              len(e2m) == 1 and e2m[0]["type"] == mod.CODIGO_BLOCO_COMENTADO)
        check(p("#4 regressão: atributos do bloco comentado preservados"),
              e2m and e2m[0]["attributes"] == {"ID": "E1", "NOME": "Comentado"})

        arq = escrever(pasta, "gr2act.dat", ";GR2ACT\n;\tID = G1\n")
        dados = parsear(mod, arq, "gr2act.dat", {"GR2ACT"})
        check(p("#4 regressão: ';GR2ACT' (dígito no meio) também vira bloco"),
              len(dados.get("gr2act", [])) == 1)

        # Entidade desconhecida comentada continua sendo comentário, não bloco.
        arq = escrever(pasta, "pdf.dat", "PDF\n\tID = F0\n;NAOEHENTIDADE\n")
        dados = parsear(mod, arq, "pdf.dat", {"PDF"})
        check(p("#4: ';NAOEHENTIDADE' não cria entidade nova"),
              list(dados.keys()) == ["pdf"])

        # ----------------------------------------------------------
        # 3. Arquivo em utf-8 (auditoria #5)
        # ----------------------------------------------------------
        arq = escrever(pasta, "pdf_utf8.dat", "PDF\n\tID = F1\n\tNOME = Tensão média\n",
                       encoding="utf-8")
        dados = parsear(mod, arq, "pdf_utf8.dat", {"PDF"})
        pdf = dados.get("pdf", [])
        check(p("#5 regressão: .dat utf-8 importa sem mojibake"),
              pdf and pdf[0]["attributes"].get("NOME") == "Tensão média")

        # O mesmo texto em latin-1 continua correto (a ordem utf-8 -> latin-1 não
        # pode quebrar o formato nativo do SAGE).
        arq = escrever(pasta, "pdf_latin1.dat", "PDF\n\tID = F2\n\tNOME = Tensão média\n")
        dados = parsear(mod, arq, "pdf_latin1.dat", {"PDF"})
        pdf = dados.get("pdf", [])
        check(p("#5 regressão: .dat latin-1 continua correto"),
              pdf and pdf[0]["attributes"].get("NOME") == "Tensão média")

        # ----------------------------------------------------------
        # 4. CRLF (arquivo editado no Windows)
        # ----------------------------------------------------------
        arq = escrever(pasta, "pas.dat", None,
                       binario="PAS\r\n\tID = A1\r\n\tNOME = Com CRLF\r\n".encode("latin-1"))
        dados = parsear(mod, arq, "pas.dat", {"PAS"})
        pas = dados.get("pas", [])
        check(p("CRLF: bloco importado"), len(pas) == 1)
        check(p("CRLF: nenhum '\\r' sobra nos valores"),
              pas and all("\r" not in v for v in pas[0]["attributes"].values()))

        # ----------------------------------------------------------
        # 5. Includes ativo e comentado
        # ----------------------------------------------------------
        arq = escrever(pasta, "cnf.dat",
                       "#include sub/ativo.dat\n"
                       ";#include sub/inativo.dat\n")
        dados = parsear(mod, arq, "cnf.dat", {"CNF"})
        cnf = dados.get("cnf", [])
        check(p("include: 2 linhas (ativo + comentado)"), len(cnf) == 2)
        check(p("include ativo: tipo 'i' com path"),
              len(cnf) == 2 and cnf[0]["type"] == mod.CODIGO_INCLUDE
              and cnf[0]["data"] == "sub/ativo.dat")
        check(p("include comentado: tipo 'u' com path"),
              len(cnf) == 2 and cnf[1]["type"] == mod.CODIGO_INCLUDE_COMENTADO
              and cnf[1]["data"] == "sub/inativo.dat")

        # ----------------------------------------------------------
        # 6. Origem: toda linha carrega o caminho relativo (usado na exportação)
        # ----------------------------------------------------------
        check(p("origem: caminho relativo gravado em cada ponto"),
              cnf and all(linha["origem"] == "cnf.dat" for linha in cnf))
    finally:
        shutil.rmtree(pasta, ignore_errors=True)


def main():
    for trilha, caminho in TRILHAS:
        print(f"\n--- trilha {trilha} ({os.path.relpath(caminho, RAIZ_PROJETO)}) ---")
        rodar(trilha, caminho)
    print()
    if falhas:
        print(f"FALHOU: {len(falhas)} checagem(ns): {falhas}")
        raise SystemExit(1)
    print("Todas as checagens do smoke test do parser passaram.")


if __name__ == "__main__":
    main()
