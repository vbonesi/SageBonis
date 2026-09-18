# -*- coding: utf-8 -*-
"""Smoke test do verificador de base (lógica pura, sem UNO) -- auditoria #6.

O verificador era a outra peça crítica sem teste automatizado: nada exercitava
_rodar_checagens, e a única ferramenta existente (completa/testar_verificacao.py)
depende de um .ods real salvo na mão, então não serve de rede de regressão.

Cobre também a reconciliação da aba VerificacaoRefs com REGRAS_REFS_PADRAO e o
grafo de troca de ID, que eram duas fontes de verdade divergentes.

Roda com: python completa/tests/smoke_test_verificador.py
"""
import importlib.util
import os

CAMINHO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ImportadorSAGE.py")
spec = importlib.util.spec_from_file_location("mod_verificador", CAMINHO)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

falhas = []


def check(nome, cond):
    status = "OK" if cond else "FALHOU"
    print(f"[{status}] {nome}")
    if not cond:
        falhas.append(nome)


def descrs(analise, sev=None):
    return [a["descr"] for a in analise.achados if sev is None or a["sev"] == sev]


H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC"]
H_TAC = ["Origem", "Gera", "Comentario/Include", "ID", "NOME"]

# ------------------------------------------------------------------
# 1. _check_ids: ID vazio e ID duplicado em ponto ativo
# ------------------------------------------------------------------
entidades = {
    "pds": (H, [
        ["pds.dat", "x", "", "P1", "Disjuntor", "TAC1"],
        ["pds.dat", "x", "", "", "Sem ID", "TAC1"],        # ERRO: ativo sem ID
        ["pds.dat", "x", "", "P1", "Duplicado", "TAC1"],   # ERRO: ID duplicado
        ["pds.dat", "c", "", "", "Comentado sem ID", ""],  # inativo: não conta
    ]),
    "tac": (H_TAC, [["tac.dat", "x", "", "TAC1", "Tac 1"]]),
}
analise = mod._rodar_checagens(entidades, [])
check("ids: ponto ativo sem ID vira ERRO", "Ponto ativo sem ID" in descrs(analise, mod.SEV_ERRO))
check("ids: ID duplicado vira ERRO",
      any(d.startswith("ID duplicado") for d in descrs(analise, mod.SEV_ERRO)))
check("ids: linha inativa (Gera='c') não gera achado", analise.erros == 2)
check("ids: base sem outros problemas não gera aviso", analise.avisos == 0)

# ------------------------------------------------------------------
# 2. _check_tamanho_id: limite conhecido por entidade (AVISO, não ERRO)
# ------------------------------------------------------------------
longo = "L" * (mod.LIMITES_TAMANHO_ID["LSC"] + 1)
analise = mod._rodar_checagens({"lsc": (H_TAC, [["lsc.dat", "x", "", longo, "Canal"]])}, [])
check("tamanho: ID acima do limite da entidade vira AVISO", analise.avisos == 1 and analise.erros == 0)
analise = mod._rodar_checagens({"lsc": (H_TAC, [["lsc.dat", "x", "", "L" * 8, "Canal"]])}, [])
check("tamanho: ID no limite exato não gera achado", analise.achados == [])

# ------------------------------------------------------------------
# 3. _check_integridade_referencial: FK simples, multi-destino e regra inaplicável
# ------------------------------------------------------------------
entidades = {
    "pds": (H, [
        ["pds.dat", "x", "", "P1", "Ok", "TAC1"],
        ["pds.dat", "x", "", "P2", "FK quebrada", "TAC_INEXISTENTE"],
        ["pds.dat", "c", "", "P3", "Inativo", "TAC_INEXISTENTE"],  # inativo: não checa
    ]),
    "tac": (H_TAC, [["tac.dat", "x", "", "TAC1", "Tac 1"]]),
}
regra_tac = [("PDS", "TAC", ("TAC",), "ID")]
analise = mod._rodar_checagens(entidades, regra_tac)
check("fk: referência inexistente vira ERRO", analise.erros == 1)
check("fk: só a linha ativa é checada",
      [a["valor"] for a in analise.achados if a["sev"] == mod.SEV_ERRO] == ["TAC_INEXISTENTE"])

# Multi-destino: PDF.PNT pode apontar para PDS ou PDD -- a união vale.
entidades_multi = {
    "pdf": (["Origem", "Gera", "Comentario/Include", "ID", "PNT"], [
        ["pdf.dat", "x", "", "F1", "P1"],    # existe em PDS
        ["pdf.dat", "x", "", "F2", "D1"],    # existe em PDD
        ["pdf.dat", "x", "", "F3", "XX"],    # não existe em nenhum
    ]),
    "pds": (H_TAC, [["pds.dat", "x", "", "P1", "Ponto"]]),
    "pdd": (H_TAC, [["pdd.dat", "x", "", "D1", "Ponto duplo"]]),
}
analise = mod._rodar_checagens(entidades_multi, [("PDF", "PNT", ("PDS", "PDD"), "ID")])
check("fk multi-destino: união dos destinos aceita os dois lados", analise.erros == 1)
check("fk multi-destino: só o valor inexistente é acusado",
      analise.achados[0]["valor"] == "XX")

# Regra cujo destino não existe na base vira AVISO (não trava a verificação).
analise = mod._rodar_checagens({"pds": (H, [["pds.dat", "x", "", "P1", "Ok", "TAC1"]])},
                               [("PDS", "TAC", ("TAC",), "ID")])
check("fk: destino ausente da base vira AVISO, não ERRO",
      analise.erros == 0 and analise.avisos == 1)

# Sem regras (estado de fábrica: todas inativas), nenhuma checagem de FK roda.
analise = mod._rodar_checagens(entidades, [])
check("fk: sem regras ativas não há achado de referência",
      not any("Referencia" in d for d in descrs(analise)))

# ------------------------------------------------------------------
# 4. Reconciliação da aba VerificacaoRefs (auditoria #6)
# ------------------------------------------------------------------
faltantes_do_zero = mod._regras_refs_faltantes([])
check("refs: aba vazia recebe todas as regras padrão",
      len(faltantes_do_zero) == len(mod.REGRAS_REFS_PADRAO))
check("refs: regras novas nascem inativas",
      all(linha[4] == "N" for linha in faltantes_do_zero))

ja_na_aba = [[ent_o, attr_o, ent_d, attr_d, "N"]
             for ent_o, attr_o, ent_d, attr_d in mod.REGRAS_REFS_PADRAO]
check("refs: aba já em dia não recebe nada", mod._regras_refs_faltantes(ja_na_aba) == [])

# A regra que o usuário ativou/editou não pode voltar duplicada.
editada = [list(linha) for linha in ja_na_aba]
editada[0][4] = "S"
editada[1][2] = editada[1][2].lower()            # destino em minúsculas
editada[2][0] = " %s " % editada[2][0]           # espaços em volta
check("refs: regra ativada/editada pelo usuário não duplica",
      mod._regras_refs_faltantes(editada) == [])

parcial = ja_na_aba[:-3]
check("refs: só as regras ausentes são acrescentadas",
      len(mod._regras_refs_faltantes(parcial)) == 3)

# ------------------------------------------------------------------
# 5. Grafo de troca de ID x regras da aba (auditoria #6)
# ------------------------------------------------------------------
mapa_padrao = mod._construir_mapa_referencias_por_destino()
check("mapa: sem regras extra continua igual ao padrão",
      ("PDD", "PDS") in mapa_padrao.get("PDS", []))

mapa_com_extra = mod._construir_mapa_referencias_por_destino(
    [("MINHAENT", "MEUPDS", ("PDS",), "ID")])
check("mapa: regra própria da aba entra no grafo da troca de ID",
      ("MINHAENT", "MEUPDS") in mapa_com_extra.get("PDS", []))
check("mapa: regra da aba que repete a padrão não duplica o par",
      mod._construir_mapa_referencias_por_destino(
          [("PDD", "PDS", ("PDS",), "ID")]).get("PDS", []).count(("PDD", "PDS")) == 1)

if falhas:
    print(f"\nFALHOU: {len(falhas)} checagem(ns): {falhas}")
    raise SystemExit(1)
print("\nTodas as checagens do smoke test do verificador passaram.")
