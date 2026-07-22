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
CGS_H = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC", "PAC", "PINT", "TIPOE", "TPCTL",
         "LMI1C", "LMI2C", "LMS1C", "LMS2C"]
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
# 2. Digital com comando associado (CGS.PAC == PDS.ID -- o FK real; CGS.ID
# aqui tambem bate por coincidencia, ver 2b pro caso em que NAO bate)
# ------------------------------------------------------------------
entidades2 = dict(entidades)
entidades2["cgs"] = ent(CGS_H, [["", "x", "", "DEM:230:LIN01:DJ:POS", "Disjuntor Posicao", "TAC1",
                                  "DEM:230:LIN01:DJ:POS", "", "", ""]])
entidades2["cgf"] = ent(CGF_H, [["", "x", "", "IED1.CTRL-CSWI$CO$Pos", "IED1.CTRL_CSIM",
                                 "DEM:230:LIN01:DJ:POS", "SBOw TERM"]])
linhas_pd2 = mod._extrair_ponto_digital(entidades2)
check("comando: Comando=S", linhas_pd2[0]["Comando"] == "S")

# ------------------------------------------------------------------
# 2b. Digital com comando associado, CGS.ID DIFERENTE do ID do ponto (achado
# real: base jdm/CHESF tem CGS.ID="JDM:REGU:STPC" com PAC="JDM:REGU-STPS" --
# a identidade do comando e' independente da do ponto; so' o PAC liga os dois,
# exatamente como o grafo de FK do verificador ja documentava: CGS.PAC->PDS|PAS)
# ------------------------------------------------------------------
entidades2b = dict(entidades)
entidades2b["cgs"] = ent(CGS_H, [["", "x", "", "CMD:INDEPENDENTE:123", "Comando com ID proprio", "TAC1",
                                   "DEM:230:LIN01:DJ:POS", "", "", ""]])
entidades2b["cgf"] = ent(CGF_H, [["", "x", "", "IED1.CTRL-CSWI$CO$Pos", "IED1.CTRL_CSIM",
                                  "CMD:INDEPENDENTE:123", "SBOw TERM"]])
linhas_pd2b = mod._extrair_ponto_digital(entidades2b)
check("comando com ID proprio: Comando=S mesmo com CGS.ID != PDS.ID (via PAC)",
      linhas_pd2b[0]["Comando"] == "S")
check("comando com ID proprio: ID_Fisico_Comando ainda resolvido (CGF casado pelo ID real do CGS)",
      linhas_pd2b[0]["ID_Fisico_Comando"] == "IED1.CTRL-CSWI$CO$Pos")
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
# 4c. Analogico simples (1 PAS + 1 PAF, sem comando)
# ------------------------------------------------------------------
entidades4c = {
    "pas": ent(PAS_H, [["", "x", "", "DEM:UG01:POT_ATIVA", "Potencia Ativa", "TAC1", "OCR1"]]),
    "paf": ent(PAF_H, [["", "x", "", "IED1.MEAS-MMXU$MX$TotW", "IED1.MEAS_ADAQ",
                        "DEM:UG01:POT_ATIVA", "MV", "1", "0", "0", "Potencia Ativa"]]),
}
linhas_pa4c = mod._extrair_ponto_analogico(entidades4c)
check("analogico simples: 1 linha extraida", len(linhas_pa4c) == 1)
check("analogico simples: Comando=N (sem CGS)", linhas_pa4c[0]["Comando"] == "N")
check("analogico simples: ID_Fisico do PAF", linhas_pa4c[0]["ID_Fisico"] == "IED1.MEAS-MMXU$MX$TotW")

# ------------------------------------------------------------------
# 4d. Analogico com comando (setpoint, CGS.PAC == PAS.ID -- o FK real) --
# reverso de _gerar_fan_out_analogico; achado real CGS.TIPO=PAS com limites
# LMI1C/LMS1C (tucurui/jdm, ver PLANEJAMENTO.md). jdm tambem confirmou que
# CGS.ID pode ser totalmente independente do ID do ponto (ver 4e).
# ------------------------------------------------------------------
entidades4d = dict(entidades4c)
entidades4d["cgs"] = ent(CGS_H, [["", "x", "", "DEM:UG01:POT_ATIVA", "Potencia Ativa", "TAC1",
                                   "DEM:UG01:POT_ATIVA", "", "", "", "0", "0", "100", "100"]])
entidades4d["cgf"] = ent(CGF_H, [["", "x", "", "IED1.CTRL-ATCC$CO$SetMag", "IED1.CTRL_CSIM",
                                  "DEM:UG01:POT_ATIVA", "APC"]])
linhas_pa4d = mod._extrair_ponto_analogico(entidades4d)
check("analogico com comando: Comando=S", linhas_pa4d[0]["Comando"] == "S")
check("analogico com comando: ID_Fisico_Comando preenchido",
      linhas_pa4d[0]["ID_Fisico_Comando"] == "IED1.CTRL-ATCC$CO$SetMag")
check("analogico com comando: KCONV_Comando preenchido", linhas_pa4d[0]["KCONV_Comando"] == "APC")
check("analogico com comando: limites LMI1C/LMS1C carregados",
      linhas_pa4d[0]["LMI1C"] == "0" and linhas_pa4d[0]["LMS1C"] == "100")
check("analogico com comando: limites LMI2C/LMS2C carregados",
      linhas_pa4d[0]["LMI2C"] == "0" and linhas_pa4d[0]["LMS2C"] == "100")

# ------------------------------------------------------------------
# 4e. Analogico com comando, CGS.ID DIFERENTE do ID do ponto -- reproduz
# literalmente o achado real da base jdm/CHESF: CGS.ID="JDM:REGU:STPC",
# PAC="JDM:REGU-STPS", LMI1C=680/LMS1C=715 (setpoint numerico de tensao via
# UTR). So' o PAC (nao o ID) liga o comando ao ponto.
# ------------------------------------------------------------------
entidades4e = {
    "pas": ent(PAS_H, [["", "x", "", "JDM:REGU-STPS", "Valor da Tensao de Regulacao via UTR-JDM",
                        "JDM", ""]]),
    "paf": ent(PAF_H, [["", "x", "", "JDM_ADNP_1_AANL_103", "JDM_ADNP_1_AANL",
                        "JDM:REGU-STPS", "PAS", ".045", "0", "BIP", ""]]),
    "cgs": ent(CGS_H, [["", "x", "", "JDM:REGU:STPC", "Valor da Tensao de Regulacao via UTR-JDM", "JDM",
                        "JDM:REGU-STPS", "JDM:UTR-069:90:PAUT", "STPT", "CSAC", "680", "0", "715", "0"]]),
    "cgf": ent(CGF_H, [["", "x", "", "JDM_CDNP_2_CSTP_0", "JDM_CDNP_2_CSTP",
                        "JDM:REGU:STPC", "ON"]]),
}
linhas_pa4e = mod._extrair_ponto_analogico(entidades4e)
check("achado real jdm: Comando=S mesmo com CGS.ID != PAS.ID (via PAC)", linhas_pa4e[0]["Comando"] == "S")
check("achado real jdm: limites LMI1C=680/LMS1C=715 carregados (setpoint numerico de tensao)",
      linhas_pa4e[0]["LMI1C"] == "680" and linhas_pa4e[0]["LMS1C"] == "715")
check("achado real jdm: ID_Fisico_Comando resolvido (CGF casado pelo ID real do CGS)",
      linhas_pa4e[0]["ID_Fisico_Comando"] == "JDM_CDNP_2_CSTP_0")

# ------------------------------------------------------------------
# 5. Comando avulso (CGS sem PDS/PAS correspondente, tipo COM_SAGE)
# ------------------------------------------------------------------
entidades5 = {
    "cgs": ent(CGS_H, [
        ["", "x", "", "DEM:SAGE:RESET", "Reset", "TAC_LOCAL", "COM_SAGE", "", "AFIC", "CSAC"],
        ["", "x", "", "DEM:SAGE:RESYNC", "Resync", "TAC_LOCAL", "COM_SAGE", "", "AFIC", "CSAC"],
        # 3o CGS, PAC != "COM_SAGE" -- so' pra testar a exclusao abaixo (ver 5c).
        ["", "x", "", "DEM:SAGE:TRIP", "Trip", "TAC_LOCAL", "DEM:REAL:PONTO_JA_EXTRAIDO", "", "AFIC", "CSAC"],
    ]),
    "cgf": ent(CGF_H, [
        ["", "x", "", "TAC1.CTRL-CSWI$CO$Rst", "TAC1.CTRL_CSIM", "DEM:SAGE:RESET", "SBOw"],
        ["", "x", "", "TAC1.CTRL-CSWI$CO$Sync", "TAC1.CTRL_CSIM", "DEM:SAGE:RESYNC", "SBOw"],
        ["", "x", "", "TAC1.CTRL-CSWI$CO$Trip", "TAC1.CTRL_CSIM", "DEM:SAGE:TRIP", "SBOw"],
    ]),
}
linhas_avulsos = mod._extrair_comandos_avulsos(entidades5, ids_logicos_com_ponto=set())
check("avulso: 3 comandos extraidos (nenhum ponto real extraido ainda)", len(linhas_avulsos) == 3)
check("avulso: RESET/RESYNC compartilham o mesmo PAC (ponto generico)",
      all(l["PAC"] == "COM_SAGE" for l in linhas_avulsos if l["ID"] in ("DEM:SAGE:RESET", "DEM:SAGE:RESYNC"))
      and all(l["TAC"] == "TAC_LOCAL" for l in linhas_avulsos))
# 5c. PAC (nao ID) e' o que decide exclusao -- se o PAC de um CGS ja aparece
# como ponto extraido (ids_logicos_com_ponto), esse CGS NAO deve virar avulso
# (ja foi contabilizado como Comando=S do proprio ponto); os outros 2
# (PAC="COM_SAGE", nunca bate com um ID de ponto de verdade) continuam avulso.
linhas_avulsos_filtrado = mod._extrair_comandos_avulsos(
    entidades5, ids_logicos_com_ponto={"DEM:REAL:PONTO_JA_EXTRAIDO"})
check("avulso: exclui so' o CGS cujo PAC ja tem ponto proprio (via PAC, nao ID)",
      len(linhas_avulsos_filtrado) == 2
      and {l["ID"] for l in linhas_avulsos_filtrado} == {"DEM:SAGE:RESET", "DEM:SAGE:RESYNC"})

# ------------------------------------------------------------------
# 5b. Comando avulso ANALOGICO (com limites LMI1C/LMS1C, achado real ur_mir/tucurui)
# ------------------------------------------------------------------
entidades5b = {
    "cgs": ent(CGS_H, [
        ["", "x", "", "DEM:SAGE:SETPOINT_DUMMY", "Setpoint Dummy", "TAC_LOCAL", "COM_SAGE_ANA",
         "", "", "", "0", "0", "100", "100"],
    ]),
    "cgf": ent(CGF_H, [
        ["", "x", "", "TAC1.CTRL-ATCC$CO$SetMag", "TAC1.CTRL_CSIM", "DEM:SAGE:SETPOINT_DUMMY", "APC"],
    ]),
}
linhas_avulsos_ana = mod._extrair_comandos_avulsos(entidades5b, ids_logicos_com_ponto=set())
check("avulso analogico: 1 comando extraido", len(linhas_avulsos_ana) == 1)
check("avulso analogico: limites LMI1C/LMS1C carregados",
      linhas_avulsos_ana[0]["LMI1C"] == "0" and linhas_avulsos_ana[0]["LMS1C"] == "100")

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
