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
        ["", " I ", "old_dir/sub3.dat", "", "", ""],
        ["", "n", "comentario simples", "", "", ""],
    ]),
    # Entidades sem as colunas padrao devem ser ignoradas com seguranca.
    "config": (["Chave", "Valor"], [["old_dir", "nao alterar"]]),
})
listagem = mod._listar_includes(entidades_inc)
check("includes: lista so as 3 linhas de include (i/u, sem diferenciar caixa/espacos)", len(listagem) == 3)
check("includes: paths corretos", {p for _, _, p in listagem} == {
    "old_dir/sub1.dat", "old_dir/sub2.dat", "old_dir/sub3.dat"})
check("includes: informa as linhas reais da planilha (cabecalho = linha 1)",
      [linha for _, linha, _ in listagem] == [3, 4, 5])

tocadas_inc, n_inc, mudancas_inc = mod._substituir_em_includes(entidades_inc, "old_dir", "new_dir")
check("includes: 3 substituicoes feitas", n_inc == 3)
check("includes: reporta somente a entidade realmente tocada", tocadas_inc == {"pds"})
listagem2 = mod._listar_includes(entidades_inc)
check("includes: paths atualizados", {p for _, _, p in listagem2} == {
    "new_dir/sub1.dat", "new_dir/sub2.dat", "new_dir/sub3.dat"})
col_dados = mod._idx_coluna(PDS_H, "Comentario/Include")
check("includes: NAO mexeu no comentario da linha ativa (Gera=x)",
      "old_dir" in entidades_inc["pds"][1][0][col_dados])
tocadas_repeticao, n_repeticao, _mud = mod._substituir_em_includes(entidades_inc, "old_dir", "new_dir")
check("includes: repetir a mesma regra e idempotente", not tocadas_repeticao and n_repeticao == 0)
antes_vazio = list(_listar for _listar in mod._listar_includes(entidades_inc))
tocadas_vazio, n_vazio, _mud_vazio = mod._substituir_em_includes(entidades_inc, "", "nao_deve_entrar")
check("includes: busca vazia nao altera nada", not tocadas_vazio and n_vazio == 0)
check("includes: busca vazia preserva todos os paths", mod._listar_includes(entidades_inc) == antes_vazio)
check("includes: relatorio traz antes -> depois de cada linha alterada",
      sorted(mudancas_inc) == [("pds", 3, "old_dir/sub1.dat", "new_dir/sub1.dat"),
                               ("pds", 4, "old_dir/sub2.dat", "new_dir/sub2.dat"),
                               ("pds", 5, "old_dir/sub3.dat", "new_dir/sub3.dat")])

# Fronteira de token: "jdm" nao pode casar dentro de "ajdm"/"jdmx" (auditoria #11).
entidades_token = mod._preparar_entidades_mutaveis({
    "pds": (PDS_H, [
        ["", "i", "jdm/pds.dat", "", "", ""],        # token isolado -> troca
        ["", "i", "base/ajdm.dat", "", "", ""],      # sufixo de outra palavra -> nao
        ["", "i", "base/jdmx.dat", "", "", ""],      # prefixo de outra palavra -> nao
        ["", "i", "base/pds_jdm.dat", "", "", ""],   # separado por '_' -> troca
        ["", "u", "sub/jdm.dat", "", "", ""],        # include comentado -> troca
    ]),
})
tocadas_token, n_token, mudancas_token = mod._substituir_em_includes(entidades_token, "jdm", "itb")
check("includes: substitui so o token de path, nao a substring", n_token == 3)
check("includes: paths vizinhos ficam intactos",
      {p for _, _, p in mod._listar_includes(entidades_token)} == {
          "itb/pds.dat", "base/ajdm.dat", "base/jdmx.dat", "base/pds_itb.dat", "sub/itb.dat"})
check("includes: mudancas listam so as linhas realmente alteradas",
      [linha for _, linha, _, _ in sorted(mudancas_token)] == [2, 5, 6])

_toc_sep, n_sep, _mud_sep = mod._substituir_em_includes(
    mod._preparar_entidades_mutaveis({"pds": (PDS_H, [["", "i", "a/jdm/b.dat", "", "", ""]])}),
    "/jdm/", "/itb/")
check("includes: busca ja delimitada por separador continua casando", n_sep == 1)

# ------------------------------------------------------------------
# 7. Sanitizacao da exportacao Latin-1
# ------------------------------------------------------------------
texto_latin1, substituicoes_latin1 = mod._sanitizar_para_latin1("Ação, café e posição")
check("latin-1: acentos suportados sao preservados", texto_latin1 == "Ação, café e posição")
check("latin-1: acentos suportados nao contam como substituicao", substituicoes_latin1 == 0)

texto_unicode, substituicoes_unicode = mod._sanitizar_para_latin1(
    "Traço – aspas “teste” reticências… bullet • espaço\u00a0não-quebrável emoji 😀")
check("latin-1: pontuacao Unicode conhecida e normalizada",
      texto_unicode == 'Traço - aspas "teste" reticências... bullet - espaço não-quebrável emoji ?')
check("latin-1: somente caractere sem mapeamento conta como substituicao",
      substituicoes_unicode == 1)
check("latin-1: resultado sanitizado sempre pode ser codificado",
      texto_unicode.encode("latin-1").decode("latin-1") == texto_unicode)

# ------------------------------------------------------------------
# 8. Abas internas: uma lista so (auditoria #1)
# ------------------------------------------------------------------
# Guarda contra a divergencia que ja aconteceu duas vezes: aba de config/relatorio
# nova criada pela Completa sem entrar na lista -> a exportacao total tenta exporta-la,
# nao acha as 3 colunas padrao e termina inteira em "ERRO".
nao_entidade = mod._abas_nao_entidade()
constantes_aba = {nome: valor for nome, valor in vars(mod).items()
                  if nome.startswith("NOME_ABA_") and isinstance(valor, str)}
fora_da_lista = sorted(f"{nome}={valor}" for nome, valor in constantes_aba.items()
                       if valor.lower() not in nao_entidade)
check("abas: toda NOME_ABA_* da macro esta na lista de nao-entidade (%s constantes)"
      % len(constantes_aba), not fora_da_lista)
if fora_da_lista:
    print("      faltando:", ", ".join(fora_da_lista))
check("abas: aba de entidade de verdade nao entra na lista",
      "pds" not in nao_entidade and "nv2" not in nao_entidade)

# ------------------------------------------------------------------
# 9. Avisos do status da exportacao (auditoria #8)
# ------------------------------------------------------------------
check("avisos: sem nada a avisar, sufixo vazio", mod._avisos_exportacao(0, 0) == "")
check("avisos: so linhas sem Origem",
      mod._avisos_exportacao(0, 3) == " (3 linha(s) sem Origem ignorada(s))")
aviso_completo = mod._avisos_exportacao(2, 3)
check("avisos: os dois avisos aparecem juntos",
      "3 linha(s) sem Origem" in aviso_completo and "2 caractere(s)" in aviso_completo)

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
