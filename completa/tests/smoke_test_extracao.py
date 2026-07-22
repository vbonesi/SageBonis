# -*- coding: utf-8 -*-
"""Smoke test em memória da extração reversa (lógica pura, sem UNO).
Roda com: python completa/tests/smoke_test_extracao.py"""
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


def ent(headers, linhas):
    return (headers, linhas)


PDS_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC", "OCR"]
PDF_H = ["Origem", "Gera", "Comentario/Include", "ID", "NV2", "PNT", "TPPNT", "KCONV", "DESC1"]
CGS_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC", "PAC", "PINT", "TIPOE", "TPCTL"]
CGF_H = ["Origem", "Gera", "Comentario/Include", "ID", "NV2", "CGS", "KCONV"]
PAS_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC", "OCR"]
PAF_H = ["Origem", "Gera", "Comentario/Include", "ID", "NV2", "PNT", "TPPNT", "KCONV1", "KCONV2", "KCONV3", "DESC1"]
PDD_H = ["Origem", "Gera", "Comentario/Include", "ID", "PDS", "TDD", "ORDEM"]

# ------------------------------------------------------------------
# 1. Digital simples (1 PDS + 1 PDF)
# ------------------------------------------------------------------
entidades = {
    "pds": ent(PDS_H, [["04C1\\pds.dat", "x", "", "DEM:230:LIN01:DJ:POS", "Disjuntor Posicao", "TAC1", "OCR1"]]),
    "pdf": ent(PDF_H, [["04C1\\pdf.dat", "x", "", "IED1.CTRL-XCBR$ST$Pos", "IED1.CTRL_ADAQ",
                        "DEM:230:LIN01:DJ:POS", "PDS", "DPS0", "Disjuntor Posicao"]]),
}
linhas_pd = mod._extrair_ponto_digital(entidades)
check("simples: 1 linha extraida", len(linhas_pd) == 1)
check("simples: Comando=N (sem CGS)", linhas_pd[0]["Comando"] == "N")
check("simples: ID_Fisico do PDF", linhas_pd[0]["ID_Fisico"] == "IED1.CTRL-XCBR$ST$Pos")
check("simples: TAC/OCR vieram do PDS", linhas_pd[0]["TAC"] == "TAC1" and linhas_pd[0]["OCR"] == "OCR1")

# ------------------------------------------------------------------
# 2. Digital com comando associado (CGS.ID == PDS.ID)
# ------------------------------------------------------------------
entidades2 = dict(entidades)
entidades2["cgs"] = ent(CGS_H, [["", "x", "", "DEM:230:LIN01:DJ:POS", "Disjuntor Posicao", "TAC1", "", "", "", ""]])
entidades2["cgf"] = ent(CGF_H, [["", "x", "", "IED1.CTRL-CSWI$CO$Pos", "IED1.CTRL_CSIM",
                                 "DEM:230:LIN01:DJ:POS", "SBOw TERM"]])
linhas_pd2 = mod._extrair_ponto_digital(entidades2)
check("comando: Comando=S", linhas_pd2[0]["Comando"] == "S")
check("comando: ID_Fisico_Comando preenchido", linhas_pd2[0]["ID_Fisico_Comando"] == "IED1.CTRL-CSWI$CO$Pos")

# ------------------------------------------------------------------
# 3. Digital redundante (2 PDF apontando pro mesmo PDS)
# ------------------------------------------------------------------
entidades3 = {
    "pds": ent(PDS_H, [["", "x", "", "DEM:230:LIN01:DJ:POS", "Disjuntor", "TAC1", "OCR1"]]),
    "pdf": ent(PDF_H, [
        ["", "x", "", "IED_P.CTRL-XCBR$ST$Pos", "IED_P.CTRL_ADAQ", "DEM:230:LIN01:DJ:POS", "PDS", "DPS0", "P"],
        ["", "x", "", "IED_D.CTRL-XCBR$ST$Pos", "IED_D.CTRL_ADAQ", "DEM:230:LIN01:DJ:POS", "PDS", "DPS0", "D"],
    ]),
}
linhas_pd3 = mod._extrair_ponto_digital(entidades3)
check("redundante: 2 linhas extraidas (mesmo ID_Logico)", len(linhas_pd3) == 2)
check("redundante: ID_Fisico distintos",
      {l["ID_Fisico"] for l in linhas_pd3} == {"IED_P.CTRL-XCBR$ST$Pos", "IED_D.CTRL-XCBR$ST$Pos"})

# ------------------------------------------------------------------
# 4. PDS sem PDF correspondente (tipicamente ponto calculado) -- NAO extrai (regressao:
# versao anterior fabricava ID_Fisico=ID_Logico, o que gerava PDF fantasma ao rodar
# unificar_pontos() de novo -- achado em teste real via UNO)
# ------------------------------------------------------------------
entidades4 = {"pds": ent(PDS_H, [["", "x", "", "DEM:ORFAO", "Ponto sem PDF", "TAC1", ""]])}
linhas_pd4 = mod._extrair_ponto_digital(entidades4)
check("orfao (sem PDF): nao extrai nenhuma linha", len(linhas_pd4) == 0)

entidades4b = {"pas": ent(PAS_H, [["", "x", "", "DEM:ORFAO_ANA", "Ponto sem PAF", "TAC1", ""]])}
linhas_pa4b = mod._extrair_ponto_analogico(entidades4b)
check("orfao analogico (sem PAF): nao extrai nenhuma linha", len(linhas_pa4b) == 0)

# ------------------------------------------------------------------
# 5. Comando avulso (CGS sem PDS/PAS correspondente, tipo COM_SAGE)
# ------------------------------------------------------------------
entidades5 = {
    "cgs": ent(CGS_H, [
        ["", "x", "", "DEM:SAGE:RESET", "Reset", "TAC_LOCAL", "COM_SAGE", "", "AFIC", "CSAC"],
        ["", "x", "", "DEM:SAGE:RESYNC", "Resync", "TAC_LOCAL", "COM_SAGE", "", "AFIC", "CSAC"],
    ]),
    "cgf": ent(CGF_H, [
        ["", "x", "", "TAC1.CTRL-CSWI$CO$Rst", "TAC1.CTRL_CSIM", "DEM:SAGE:RESET", "SBOw"],
        ["", "x", "", "TAC1.CTRL-CSWI$CO$Sync", "TAC1.CTRL_CSIM", "DEM:SAGE:RESYNC", "SBOw"],
    ]),
}
linhas_avulsos = mod._extrair_comandos_avulsos(entidades5, ids_logicos_com_ponto=set())
check("avulso: 2 comandos extraidos", len(linhas_avulsos) == 2)
check("avulso: mesmo TAC/PAC compartilhado",
      all(l["TAC"] == "TAC_LOCAL" and l["PAC"] == "COM_SAGE" for l in linhas_avulsos))
# se o CGS.ID JA aparece como ponto (ids_logicos_com_ponto), nao deve virar avulso
linhas_avulsos_filtrado = mod._extrair_comandos_avulsos(entidades5, ids_logicos_com_ponto={"DEM:SAGE:RESET"})
check("avulso: exclui CGS que ja tem ponto proprio", len(linhas_avulsos_filtrado) == 1)

# ------------------------------------------------------------------
# 6. Canais/distribuicao: inferencia de Metodo (sufixo/prefixo)
# ------------------------------------------------------------------
check("infere sufixo", mod._inferir_metodo("ABC", "ABC_COR") == ("Sufixo", "_COR"))
check("infere prefixo", mod._inferir_metodo("ABC", "N104_ABC") == ("Prefixo", "N104_"))
check("sem transformacao -> None,None", mod._inferir_metodo("ABC", "ABC") == (None, None))
check("nao reconhecido -> Explicito sem valor1 (achado real conv_iccp104)",
      mod._inferir_metodo("ABC", "XYZ123") == ("Explicito", None))

entidades6 = {
    "pdd": ent(PDD_H, [
        ["", "x", "", "DEM:230:LIN01:DJ:POS_COR", "DEM:230:LIN01:DJ:POS", "COR_TDD", "1"],
        ["", "x", "", "N104_DEM:230:LIN01:DJ:POS", "DEM:230:LIN01:DJ:POS", "N104_TDD", "1"],
        ["", "x", "", "EGRD_totalmente_diferente", "DEM:230:LIN01:SC1:POS", "IND_TDD", "1"],
    ]),
}
canais, dist = mod._extrair_canais_e_distribuicao(entidades6)
check("canais: 3 canais distintos (3 TDD)", len(canais) == 3)
canal_cor = next(c for c in canais if c["TDD"] == "COR_TDD")
check("canais: metodo sufixo inferido pro COR_TDD", canal_cor["Metodo"] == "Sufixo" and canal_cor["Valor1"] == "_COR")
canal_n104 = next(c for c in canais if c["TDD"] == "N104_TDD")
check("canais: metodo prefixo inferido pro N104_TDD", canal_n104["Metodo"] == "Prefixo" and canal_n104["Valor1"] == "N104_")
canal_ind = next(c for c in canais if c["TDD"] == "IND_TDD")
check("canais: metodo explicito inferido pro IND_TDD (ID nao relacionado)", canal_ind["Metodo"] == "Explicito")
check("distribuicao: 3 linhas de ligacao", len(dist) == 3)
check("distribuicao: liga ao ID_Logico original (nao ao transformado)",
      {d["ID_Logico"] for d in dist} == {"DEM:230:LIN01:DJ:POS", "DEM:230:LIN01:SC1:POS"})
linha_ind = next(d for d in dist if d["Canal"] == "IND_TDD")
check("distribuicao: IDExplicito preenchido com o ID observado",
      linha_ind["IDExplicito"] == "EGRD_totalmente_diferente")

# ------------------------------------------------------------------
# 6b. _gerar_distribuicao com Metodo=Explicito (forward)
# ------------------------------------------------------------------
canais_explicito = {"ind": {"TDD": "IND_TDD", "Metodo": "Explicito", "Valor1": "", "Valor2": ""}}
dist_explicito = {"DEM:PONTO": [("IND", "EGRD_ID_MANUAL")]}
pdd_explicito = mod._gerar_distribuicao("DEM:PONTO", "PDS", canais_explicito, dist_explicito)
check("explicito: usa o IDExplicito como ID final", len(pdd_explicito) == 1
      and pdd_explicito[0]["ID"] == "EGRD_ID_MANUAL")
check("explicito: PDD.PDS preserva o ID logico original", pdd_explicito[0]["PDS"] == "DEM:PONTO")

dist_explicito_vazio = {"DEM:PONTO": [("IND", "")]}
pdd_explicito_vazio = mod._gerar_distribuicao("DEM:PONTO", "PDS", canais_explicito, dist_explicito_vazio)
check("explicito sem IDExplicito preenchido: nao gera PDD (nunca ID vazio)", len(pdd_explicito_vazio) == 0)

# ------------------------------------------------------------------
# 7. Idempotencia do upsert com chave composta (DistribuicaoPontos)
# ------------------------------------------------------------------
headers_dp = ["ID_Logico", "Canal", "IDExplicito", "Ativo"]
h1, l1 = mod._mesclar_linhas_upsert(headers_dp, [], dist, colunas_chave=("ID_Logico", "Canal"))
check("upsert composto: 1a rodada cria 3 linhas", len(l1) == 3)
h2, l2 = mod._mesclar_linhas_upsert(h1, l1, dist, colunas_chave=("ID_Logico", "Canal"))
check("upsert composto: 2a rodada nao duplica", len(l2) == 3)
# adiciona 1 ligacao nova (canal diferente) -- deve virar 4
dist_extra = dist + [{"ID_Logico": "DEM:230:LIN01:DJ:POS", "Canal": "OUTRO_TDD", "Ativo": "S"}]
h3, l3 = mod._mesclar_linhas_upsert(h1, l1, dist_extra, colunas_chave=("ID_Logico", "Canal"))
check("upsert composto: nova combinacao ID_Logico+Canal adiciona linha", len(l3) == 4)

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
