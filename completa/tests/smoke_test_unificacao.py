# -*- coding: utf-8 -*-
"""Smoke test em memória da Unificação de Pontos (lógica pura, sem UNO).
Roda com: python completa/tests/smoke_test_unificacao.py"""
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


def by_id(linhas, id_valor):
    return next((l for l in linhas if l.get("ID") == id_valor), None)


# ------------------------------------------------------------------
# 1. Digital simples (1 origem, sem comando, sem distribuicao)
# ------------------------------------------------------------------
headers = mod.CABECALHOS_PONTO_DIGITAL


def linha_pd(**kw):
    """Linha de PontoDigital por nome de coluna -- imune a coluna nova no meio do
    cabeçalho (fixture posicional já quebrou quando NV2_Comando entrou)."""
    d = {h: "" for h in headers}
    d.update(kw)
    return [d[h] for h in headers]


linhas = [
    linha_pd(ID_Logico="DEM:230:LIN01:DJ:POS", ID_Fisico="IED1.CTRL-XCBR$ST$Pos",
             NOME="Disjuntor Posicao", NV2="IED1.CTRL_ADAQ", KCONV="DPS0", TAC="TAC1",
             OCR="OCR1", Comando="N", Gera="x"),
]
saida = mod._gerar_fan_out_digital(linhas, headers, {}, {})
check("1 origem: 1 PDF", len(saida["pdf"]) == 1)
check("1 origem: 1 PDS", len(saida["pds"]) == 1)
check("1 origem: PDS.TPFIL=NLFL", saida["pds"][0]["TPFIL"] == "NLFL")
check("1 origem: PDF.PNT=ID_Logico", saida["pdf"][0]["PNT"] == "DEM:230:LIN01:DJ:POS")
check("1 origem: sem RFC", len(saida["rfc"]) == 0)
check("1 origem: sem CGF/CGS (Comando=N)", len(saida["cgf"]) == 0 and len(saida["cgs"]) == 0)
check("1 origem: sem PDD (sem distribuicao)", len(saida["pdd"]) == 0)

# ------------------------------------------------------------------
# 2. Digital com comando associado
# ------------------------------------------------------------------
linhas_cmd = [
    linha_pd(ID_Logico="DEM:230:LIN01:DJ:POS", ID_Fisico="IED1.CTRL-XCBR$ST$Pos",
             NOME="Disjuntor Posicao", NV2="IED1.CTRL_ADAQ", KCONV="DPS0", TAC="TAC1",
             OCR="OCR1", Comando="S", ID_Fisico_Comando="IED1.CTRL-CSWI$CO$Pos",
             KCONV_Comando="SBOw TERM", Gera="x"),
]
saida2 = mod._gerar_fan_out_digital(linhas_cmd, headers, {}, {})
check("comando: 1 CGS", len(saida2["cgs"]) == 1)
check("comando: 1 CGF", len(saida2["cgf"]) == 1)
check("comando: CGS.ID == PDS.ID (mesmo ID)", saida2["cgs"][0]["ID"] == saida2["pds"][0]["ID"])
check("comando: CGF.CGS == ID logico", saida2["cgf"][0]["CGS"] == "DEM:230:LIN01:DJ:POS")
check("comando: CGS.PAC == self", saida2["cgs"][0]["PAC"] == saida2["cgs"][0]["ID"])
check("comando: sem NV2_Comando nem NV2 conhecido, cai no NV2 de leitura (comportamento antigo)",
      saida2["cgf"][0]["NV2"] == "IED1.CTRL_ADAQ")

# O NV2 do comando é outro grupo (NV1 de comando), não o de leitura: numa base DNP3 real
# o ponto lê em EX1_ADNP_1_ASIM e comanda em EX1_CDNP_2_CDUP. Copiar o de leitura pro CGF
# jogava todo comando no grupo errado -- e o round-trip extrair->unificar fazia isso calado.
linhas_cmd_dnp = [
    linha_pd(ID_Logico="EX1:UTR-069:90:RGAT", ID_Fisico="EX1_ADNP_1_ASIM_510",
             NOME="Regulador automatico", NV2="EX1_ADNP_1_ASIM", KCONV="SQI", TAC="EX1",
             Comando="S", ID_Fisico_Comando="EX1_CDNP_2_CDUP_5102", KCONV_Comando="SQD",
             Gera="x"),
]
saida_cmd_inferido = mod._gerar_fan_out_digital(
    linhas_cmd_dnp, headers, {}, {}, {"EX1_ADNP_1_ASIM", "EX1_CDNP_2_CDUP"})
check("comando: NV2 inferido do proprio ID fisico do comando",
      saida_cmd_inferido["cgf"][0]["NV2"] == "EX1_CDNP_2_CDUP")
check("comando: PDF continua no NV2 de leitura",
      saida_cmd_inferido["pdf"][0]["NV2"] == "EX1_ADNP_1_ASIM")

saida_cmd_desconhecido = mod._gerar_fan_out_digital(
    linhas_cmd_dnp, headers, {}, {}, {"EX1_ADNP_1_ASIM"})
check("comando: NV2 inferido so vale se existir na base (senao mantem o de leitura)",
      saida_cmd_desconhecido["cgf"][0]["NV2"] == "EX1_ADNP_1_ASIM")

linhas_cmd_explicito = [
    linha_pd(ID_Logico="EX1:UTR-069:90:PAUT", ID_Fisico="EX1_ADNP_1_ASIM_514",
             NOME="Paralelismo automatico", NV2="EX1_ADNP_1_ASIM", TAC="EX1", Comando="S",
             ID_Fisico_Comando="EX1_CDNP_2_CDUP_5103", NV2_Comando="MEU_NV2_PROPRIO",
             Gera="x"),
]
saida_cmd_explicito = mod._gerar_fan_out_digital(
    linhas_cmd_explicito, headers, {}, {}, {"EX1_CDNP_2_CDUP"})
check("comando: coluna NV2_Comando ganha da inferencia",
      saida_cmd_explicito["cgf"][0]["NV2"] == "MEU_NV2_PROPRIO")

# 61850 e OIDs de SNMP nao seguem "NV2 + _ + endereco" -- a inferencia tem que se calar.
saida_cmd_61850 = mod._gerar_fan_out_digital(
    linhas_cmd, headers, {}, {}, {"IED1.CTRL_ADAQ", "IED1.CTRL_CDAQ"})
check("comando: ID fisico sem sufixo numerico (61850) nao gera palpite",
      saida_cmd_61850["cgf"][0]["NV2"] == "IED1.CTRL_ADAQ")

# ------------------------------------------------------------------
# 3. Digital redundante (2 origens fisicas, mesmo ID_Logico)
# ------------------------------------------------------------------
linhas_red = [
    linha_pd(ID_Logico="DEM:230:LIN01:DJ:POS", ID_Fisico="IED_P.CTRL-XCBR$ST$Pos",
             NOME="Disjuntor Posicao P", NV2="IED_P.CTRL_ADAQ", KCONV="DPS0", TAC="TAC1",
             OCR="OCR1", Comando="N", Gera="x"),
    linha_pd(ID_Logico="DEM:230:LIN01:DJ:POS", ID_Fisico="IED_D.CTRL-XCBR$ST$Pos",
             NOME="Disjuntor Posicao D", NV2="IED_D.CTRL_ADAQ", KCONV="DPS0", TAC="TAC1",
             OCR="OCR1", Comando="N", Gera="x"),
]
saida3 = mod._gerar_fan_out_digital(linhas_red, headers, {}, {})
check("redundante: 2 PDF", len(saida3["pdf"]) == 2)
check("redundante: 1 unico PDS", len(saida3["pds"]) == 1)
check("redundante: PDS.TPFIL=FIL5", saida3["pds"][0]["TPFIL"] == "FIL5")
check("redundante: 2 RFC em ordem 1,2", [r["ORDEM"] for r in saida3["rfc"]] == ["1", "2"])
check("redundante: RFC.PARC aponta os 2 PDF distintos",
      {r["PARC"] for r in saida3["rfc"]} == {"IED_P.CTRL-XCBR$ST$Pos", "IED_D.CTRL-XCBR$ST$Pos"})
check("redundante: RFC.TIPOP=EDC (digital)", all(r["TIPOP"] == "EDC" for r in saida3["rfc"]))

# ------------------------------------------------------------------
# 3b. Origem sem ID_Fisico (ex.: extraida de ponto calculado, sem PDF) -- regressao:
# achado em teste real via UNO que isso gerava um PDF fantasma; agora deve gerar
# SO o PDS, sem PDF/RFC algum pra essa origem.
# ------------------------------------------------------------------
linhas_sem_fisico = [
    linha_pd(ID_Logico="DEM:CALC:PONTO", NOME="Ponto calculado", TAC="TAC_CALC",
             OCR="OCR1", Comando="N", Gera="x"),
]
saida3b = mod._gerar_fan_out_digital(linhas_sem_fisico, headers, {}, {})
check("sem ID_Fisico: nenhum PDF gerado", len(saida3b["pdf"]) == 0)
check("sem ID_Fisico: nenhum RFC gerado", len(saida3b["rfc"]) == 0)
check("sem ID_Fisico: PDS ainda e gerado (TPFIL=NLFL)",
      len(saida3b["pds"]) == 1 and saida3b["pds"][0]["TPFIL"] == "NLFL")

# mistura: 1 origem com fisico + 1 sem -- deve contar como NAO redundante (so 1 fisica)
linhas_mistas = [
    linha_pd(ID_Logico="DEM:MISTO", ID_Fisico="IED1.CTRL-XCBR$ST$Pos", NOME="Com fisico",
             NV2="NV2A", KCONV="DPS0", TAC="TAC1", OCR="OCR1", Comando="N", Gera="x"),
    linha_pd(ID_Logico="DEM:MISTO", NOME="Sem fisico (linha manual incompleta)",
             TAC="TAC1", OCR="OCR1", Comando="N", Gera="x"),
]
saida3c = mod._gerar_fan_out_digital(linhas_mistas, headers, {}, {})
check("origens mistas: so 1 PDF (a com ID_Fisico)", len(saida3c["pdf"]) == 1)
check("origens mistas: NAO redundante (so 1 fisica de verdade) -> TPFIL=NLFL",
      saida3c["pds"][0]["TPFIL"] == "NLFL")
check("origens mistas: sem RFC (nao e redundante)", len(saida3c["rfc"]) == 0)

# ------------------------------------------------------------------
# 4. Distribuicao (Metodo Sufixo, Prefixo, Substituir)
# ------------------------------------------------------------------
check("metodo prefixo", mod._aplicar_metodo("ABC", "Prefixo", "PRE_", "") == "PRE_ABC")
check("metodo sufixo", mod._aplicar_metodo("ABC", "Sufixo", "_SUF", "") == "ABC_SUF")
check("metodo substituir", mod._aplicar_metodo("ABC_ORIG", "Substituir", "ORIG", "NOVO") == "ABC_NOVO")
check("metodo desconhecido nao transforma", mod._aplicar_metodo("ABC", "", "x", "y") == "ABC")

canais = {"cor": {"TDD": "COR_TDD", "Metodo": "Sufixo", "Valor1": "_COR", "Valor2": ""},
          "n104": {"TDD": "N104_TDD", "Metodo": "Prefixo", "Valor1": "N104_", "Valor2": ""}}
distribuicoes = {"DEM:230:LIN01:DJ:POS": [("COR", ""), ("N104", "")]}
saida4 = mod._gerar_fan_out_digital(linhas, headers, canais, distribuicoes)
check("distribuicao: 2 PDD (2 canais ativos)", len(saida4["pdd"]) == 2)
ids_pdd = {p["ID"] for p in saida4["pdd"]}
check("distribuicao: sufixo aplicado", "DEM:230:LIN01:DJ:POS_COR" in ids_pdd)
check("distribuicao: prefixo aplicado", "N104_DEM:230:LIN01:DJ:POS" in ids_pdd)
check("distribuicao: PDD.PDS preserva o ID logico original",
      all(p["PDS"] == "DEM:230:LIN01:DJ:POS" for p in saida4["pdd"]))
check("distribuicao: canal inativo/inexistente e ignorado",
      len(mod._gerar_fan_out_digital(linhas, headers, {}, distribuicoes)["pdd"]) == 0)

# ------------------------------------------------------------------
# 5. Analogico simples + redundante (2a fonte de medicao) + comando (setpoint)
# ------------------------------------------------------------------
headers_ana = mod.CABECALHOS_PONTO_ANALOGICO


def linha_ana(**kw):
    d = {h: "" for h in headers_ana}
    d.update(kw)
    return [d[h] for h in headers_ana]


linhas_ana = [
    linha_ana(ID_Logico="DEM:230:LIN01:MED:IA", ID_Fisico="IED1.MEAS-MMXU$MX$A",
              NOME="Corrente Fase A", NV2="IED1.MEAS_AAAQ", KCONV1="1", KCONV2="0",
              KCONV3="MV0", TAC="TAC1", OCR="OCR2", Gera="x"),
    linha_ana(ID_Logico="DEM:230:LIN01:MED:IA", ID_Fisico="IEDX.MEAS-MMXU$MX$A",
              NOME="Corrente Fase A (2a fonte)", NV2="IEDX.MEAS_AAAQ", KCONV1="1", KCONV2="0",
              KCONV3="MV0", TAC="TAC1", OCR="OCR2", Gera="x"),
]
saida5 = mod._gerar_fan_out_analogico(linhas_ana, headers_ana, {}, {})
check("analogico redundante: 2 PAF", len(saida5["paf"]) == 2)
check("analogico redundante: 1 PAS com TPFIL=FIL5", len(saida5["pas"]) == 1 and saida5["pas"][0]["TPFIL"] == "FIL5")
check("analogico redundante: RFC.TIPOP=VAC", all(r["TIPOP"] == "VAC" for r in saida5["rfc"]))
check("analogico sem Comando=S: nenhum CGS/CGF gerado",
      len(saida5["cgs"]) == 0 and len(saida5["cgf"]) == 0)

# comando analogico (setpoint) -- achado real: CGS.TIPO=PAS em 6 bases
# independentes (ver PLANEJAMENTO.md); LMI1C/LMI2C/LMS1C/LMS2C = limites
linhas_ana_cmd = [
    linha_ana(ID_Logico="DEM:UG01:SETPOINT", ID_Fisico="IED1.REGU-STPS$MX$Val",
              NOME="Setpoint Geracao UG01", NV2="IED1.REGU_AAAQ", TAC="TAC1",
              Comando="S", ID_Fisico_Comando="IED1.REGU-STPS$CO$Val",
              KCONV_Comando="MV0", LMI1C="-999999", LMI2C="0", LMS1C="999999", LMS2C="100", Gera="x"),
]
saida5b = mod._gerar_fan_out_analogico(linhas_ana_cmd, headers_ana, {}, {})
check("analogico com comando: 1 CGS, 1 CGF", len(saida5b["cgs"]) == 1 and len(saida5b["cgf"]) == 1)
check("analogico com comando: CGS.PAC == self", saida5b["cgs"][0]["PAC"] == saida5b["cgs"][0]["ID"])
check("analogico com comando: CGS tem os 4 limites",
      saida5b["cgs"][0]["LMI1C"] == "-999999" and saida5b["cgs"][0]["LMI2C"] == "0"
      and saida5b["cgs"][0]["LMS1C"] == "999999" and saida5b["cgs"][0]["LMS2C"] == "100")
check("analogico com comando: CGF.ID == ID_Fisico_Comando",
      saida5b["cgf"][0]["ID"] == "IED1.REGU-STPS$CO$Val")

# ------------------------------------------------------------------
# 6. Comando avulso (COM_SAGE generico, varios comandos no mesmo TAC/PAC) --
# digital e analogico (achado real: PAC=MC_DUMMY_SAGE_ANA em ur_mir)
# ------------------------------------------------------------------
headers_cmd = mod.CABECALHOS_COMANDO_AVULSO


def linha_cmd(**kw):
    d = {h: "" for h in headers_cmd}
    d.update(kw)
    return [d[h] for h in headers_cmd]


linhas_avulsos = [
    linha_cmd(ID="DEM:SAGE:RESET", ID_Fisico="TAC1.CTRL-CSWI$CO$Rst", NOME="Reset alarmes SAGE",
              NV2="TAC1.CTRL_CSIM", KCONV="SBOw", TAC="TAC_LOCAL", PAC="COM_SAGE",
              TIPOE="AFIC", TPCTL="CSAC", Gera="x"),
    linha_cmd(ID="DEM:SAGE:RESYNC", ID_Fisico="TAC1.CTRL-CSWI$CO$Sync", NOME="Resync SAGE",
              NV2="TAC1.CTRL_CSIM", KCONV="SBOw", TAC="TAC_LOCAL", PAC="COM_SAGE",
              TIPOE="AFIC", TPCTL="CSAC", Gera="x"),
]
saida6 = mod._gerar_comandos_avulsos(linhas_avulsos, headers_cmd)
check("avulso: 2 CGS distintos", len({c["ID"] for c in saida6["cgs"]}) == 2)
check("avulso: 2 CGF distintos", len({c["ID"] for c in saida6["cgf"]}) == 2)

linhas_avulsos_ana = [
    linha_cmd(ID="DEM:SAGE:SETPOINT_DUMMY", ID_Fisico="TAC1.CTRL-STPS$CO$Val",
              NOME="Setpoint dummy p/ ancorar transportador", NV2="TAC1.CTRL_CSIM",
              KCONV="MV0", TAC="TAC_LOCAL", PAC="COM_SAGE_ANA",
              LMI1C="0", LMI2C="0", LMS1C="100", LMS2C="100", Gera="x"),
]
saida6b = mod._gerar_comandos_avulsos(linhas_avulsos_ana, headers_cmd)
check("avulso analogico: CGS carrega os limites", saida6b["cgs"][0]["LMI1C"] == "0"
      and saida6b["cgs"][0]["LMS1C"] == "100")
check("avulso: mesmo TAC/PAC compartilhado (ponto generico)",
      all(c["TAC"] == "TAC_LOCAL" and c["PAC"] == "COM_SAGE" for c in saida6["cgs"]))

# ------------------------------------------------------------------
# 7. Upsert idempotente (regenerar nao duplica, preserva coluna extra)
# ------------------------------------------------------------------
headers_pds = ["Origem", "Gera", "Comentario/Include", "ID", "NOME", "TAC", "OBSRV"]
linhas_pds_existentes = [
    ["04C1\\pds.dat", "x", "", "DEM:230:LIN01:DJ:POS", "Nome antigo", "TAC_VELHO", "comentario manual do usuario"],
]
novas = [{"ID": "DEM:230:LIN01:DJ:POS", "NOME": "Disjuntor Posicao", "TAC": "TAC1", "TPFIL": "NLFL"}]
h_final, l_final = mod._mesclar_linhas_upsert(headers_pds, linhas_pds_existentes, novas)
check("upsert: nao duplica linha (1 so)", len(l_final) == 1)
col_obsrv = mod._idx_coluna(h_final, "OBSRV")
check("upsert: preserva coluna extra nao tocada (OBSRV)", l_final[0][col_obsrv] == "comentario manual do usuario")
col_nome = mod._idx_coluna(h_final, "NOME")
check("upsert: atualiza campo tocado (NOME)", l_final[0][col_nome] == "Disjuntor Posicao")
col_tpfil = mod._idx_coluna(h_final, "TPFIL")
check("upsert: adiciona coluna nova (TPFIL) quando nao existia", col_tpfil >= 0 and l_final[0][col_tpfil] == "NLFL")

# roda de novo (regenerar) -- ainda 1 linha so
h_final2, l_final2 = mod._mesclar_linhas_upsert(h_final, l_final, novas)
check("upsert: rodar 2x seguidas ainda da 1 linha", len(l_final2) == 1)

# ponto novo (ID nao existia) -- deve ser ADICIONADO com Gera=x e Origem=UnificacaoPontos
novas_com_extra = novas + [{"ID": "DEM:230:LIN01:SC1:POS", "NOME": "Seccionadora 1"}]
h_final3, l_final3 = mod._mesclar_linhas_upsert(headers_pds, linhas_pds_existentes, novas_com_extra)
check("upsert: adiciona linha nova quando ID nao existia", len(l_final3) == 2)
nova_linha = by_id(
    [dict(zip(h_final3, r)) for r in l_final3], "DEM:230:LIN01:SC1:POS")
check("upsert: linha nova marcada Gera=x", nova_linha is not None and nova_linha.get("Gera") == "x")
check("upsert: linha nova marcada Origem=UnificacaoPontos",
      nova_linha is not None and nova_linha.get("Origem") == mod.ORIGEM_GERADO)

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
