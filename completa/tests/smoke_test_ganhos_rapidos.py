# -*- coding: utf-8 -*-
"""Smoke test em memória dos ganhos rápidos (lógica pura, sem UNO).
Roda com: python completa/tests/smoke_test_ganhos_rapidos.py"""
import importlib.util
import os

CAMINHO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ImportadorSAGE.py")
spec = importlib.util.spec_from_file_location("mod", CAMINHO)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

falhas = []


def check(nome, cond):
    status = "OK" if cond else "FALHOU"
    print(f"[{status}] {nome}")
    if not cond:
        falhas.append(nome)


# ------------------------------------------------------------------
# 1. Mapa de referencias por destino (inverso de REGRAS_REFS_PADRAO)
# ------------------------------------------------------------------
mapa = mod._construir_mapa_referencias_por_destino()
check("mapa: PDS tem PDD.PDS como referencia", ("PDD", "PDS") in mapa.get("PDS", []))
check("mapa: PDS tem PDF.PNT como referencia (multi-destino PDS|PDD)", ("PDF", "PNT") in mapa.get("PDS", []))
check("mapa: PDD tambem tem PDF.PNT (o outro lado do multi-destino)", ("PDF", "PNT") in mapa.get("PDD", []))
check("mapa: TAC tem PDS.TAC como referencia", ("PDS", "TAC") in mapa.get("TAC", []))

# ------------------------------------------------------------------
# 2. Troca de ID simples com propagacao
# ------------------------------------------------------------------
PDS_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC"]
PDD_H = ["Origem", "Gera", "Comentario/Include", "ID", "PDS", "TDD"]
TAC_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME"]

entidades = mod._preparar_entidades_mutaveis({
    "pds": (PDS_H, [["", "x", "", "ID_ANTIGO", "Disjuntor", "TAC1"]]),
    "pdd": (PDD_H, [["", "x", "", "ID_ANTIGO_D", "ID_ANTIGO", "TDD1"]]),
    "tac": (TAC_H, [["", "x", "", "TAC1", "Tac 1"]]),
})
tocadas, relatorio = mod._trocar_id_em_entidades(entidades, "ID_ANTIGO", "ID_NOVO", mapa)
check("troca: entidade origem (pds) tocada", "pds" in tocadas)
check("troca: entidade referenciadora (pdd) tocada", "pdd" in tocadas)
check("troca: PDS.ID atualizado", entidades["pds"][1][0][mod._idx_coluna(PDS_H, "ID")] == "ID_NOVO")
check("troca: PDD.PDS (referencia) atualizado", entidades["pdd"][1][0][mod._idx_coluna(PDD_H, "PDS")] == "ID_NOVO")
check("troca: PDD.ID (proprio ID, sem relacao) NAO alterado",
      entidades["pdd"][1][0][mod._idx_coluna(PDD_H, "ID")] == "ID_ANTIGO_D")

# ------------------------------------------------------------------
# 3. ID nao encontrado / ambiguo
# ------------------------------------------------------------------
tocadas2, relatorio2 = mod._trocar_id_em_entidades(entidades, "NAO_EXISTE", "X", mapa)
check("troca: ID inexistente nao toca nada", len(tocadas2) == 0)
check("troca: ID inexistente reporta claramente", "não encontrado" in relatorio2[0])

entidades_ambiguo = mod._preparar_entidades_mutaveis({
    "pds": (PDS_H, [["", "x", "", "DUP", "A", ""]]),
    "pas": (["Origem", "Gera", "Comentario/Include", "ID", "NOME"], [["", "x", "", "DUP", "B"]]),
})
tocadas3, relatorio3 = mod._trocar_id_em_entidades(entidades_ambiguo, "DUP", "X", mapa)
check("troca: ID ambiguo (2 entidades) nao toca nada", len(tocadas3) == 0)
check("troca: ID ambiguo reporta como tal", "ambíguo" in relatorio3[0])

# ------------------------------------------------------------------
# 4. Troca em lote encadeada (A->B, depois B->C na mesma rodada)
# ------------------------------------------------------------------
entidades_cadeia = mod._preparar_entidades_mutaveis({"pds": (PDS_H, [["", "x", "", "A", "N", ""]])})
mod._trocar_id_em_entidades(entidades_cadeia, "A", "B", mapa)
mod._trocar_id_em_entidades(entidades_cadeia, "B", "C", mapa)
check("troca encadeada: resultado final e C",
      entidades_cadeia["pds"][1][0][mod._idx_coluna(PDS_H, "ID")] == "C")

# ------------------------------------------------------------------
# 5. Estatistica
# ------------------------------------------------------------------
entidades_stats = {
    "pds": (PDS_H, [["", "x", "", "A", "", ""], ["", "n", "", "", "", ""], ["", "x", "", "B", "", ""]]),
    "tac": (TAC_H, [["", "x", "", "T1", ""]]),
}
stats = mod._calcular_estatisticas(entidades_stats)
d = {nome: (total, ativas) for nome, total, ativas in stats}
check("estatistica: pds tem 3 linhas totais, 2 ativas", d["pds"] == (3, 2))
check("estatistica: tac tem 1/1", d["tac"] == (1, 1))

# ------------------------------------------------------------------
# 6. Gestao de includes
# ------------------------------------------------------------------
entidades_inc = mod._preparar_entidades_mutaveis({
    "pds": (PDS_H, [
        ["", "x", "comentario do ponto ativo com old_dir dentro", "P1", "Ponto ativo", ""],
        ["", "i", "old_dir/sub1.dat", "", "", ""],
        ["", "u", "old_dir/sub2.dat", "", "", ""],
        ["", "n", "comentario simples", "", "", ""],
    ]),
})
listagem = mod._listar_includes(entidades_inc)
check("includes: lista so as 2 linhas de include (i/u)", len(listagem) == 2)
check("includes: paths corretos", {p for _, _, p in listagem} == {"old_dir/sub1.dat", "old_dir/sub2.dat"})

tocadas_inc, n_inc = mod._substituir_em_includes(entidades_inc, "old_dir", "new_dir")
check("includes: 2 substituicoes feitas", n_inc == 2)
listagem2 = mod._listar_includes(entidades_inc)
check("includes: paths atualizados", {p for _, _, p in listagem2} == {"new_dir/sub1.dat", "new_dir/sub2.dat"})
col_dados = mod._idx_coluna(PDS_H, "Comentario/Include")
check("includes: NAO mexeu no comentario da linha ativa (Gera=x)",
      "old_dir" in entidades_inc["pds"][1][0][col_dados])

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
