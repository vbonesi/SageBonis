# -*- coding: utf-8 -*-
"""Smoke test em memória do assistente de IED/protocolo (lógica pura, sem UNO).
Roda com: python completa/tests/smoke_test_ied.py"""
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


headers = mod.CABECALHOS_IEDS


def linha_ied(**kw):
    """Monta uma linha (list) da aba IEDs a partir de kwargs, na ordem de headers."""
    d = {h: "" for h in headers}
    d.update(kw)
    return [d[h] for h in headers]


# ------------------------------------------------------------------
# 1. Protocolo desconhecido / linha incompleta -> nao gera nada
# ------------------------------------------------------------------
saida0 = mod._gerar_infra_ied(linha_ied(ID="X", Protocolo="999", Direcao="Aquisicao"), headers)
check("protocolo desconhecido: nada gerado", all(len(v) == 0 for v in saida0.values()))
saida0b = mod._gerar_infra_ied(linha_ied(ID="", Protocolo="104", Direcao="Aquisicao"), headers)
check("sem ID: nada gerado", all(len(v) == 0 for v in saida0b.values()))

# ------------------------------------------------------------------
# 2. Aquisicao 104, nao redundante
# ------------------------------------------------------------------
linha_aq = linha_ied(ID="GRD", Protocolo="104", Direcao="Aquisicao", Nome="Ligacao 104 GRD",
                      GSD="GT_SCD_1", PlPr="7", LiPr="7", PlRe="8", LiRe="8", Gera="x")
saida_aq = mod._gerar_infra_ied(linha_aq, headers)

check("aquis: 1 LSC", len(saida_aq["lsc"]) == 1)
check("aquis: LSC.TCV=CNVM/TTP=CX104/TIPO=AA",
      saida_aq["lsc"][0]["TCV"] == "CNVM" and saida_aq["lsc"][0]["TTP"] == "CX104"
      and saida_aq["lsc"][0]["TIPO"] == "AA")
check("aquis: 1 CNF com IGNERS/SINCR/INVAL", "IGNERS=" in saida_aq["cnf"][0]["CONFIG"]
      and "INVAL=" in saida_aq["cnf"][0]["CONFIG"])
check("aquis: CNF.CONFIG tem PlPr=7...LiRe=8", "PlPr= 7" in saida_aq["cnf"][0]["CONFIG"]
      and "LiRe= 8" in saida_aq["cnf"][0]["CONFIG"])
check("aquis: 1 CXU com defaults (AQANL=1000, INTGR=180000 p/ aquisicao)",
      saida_aq["cxu"][0]["AQANL"] == "1000" and saida_aq["cxu"][0]["INTGR"] == "180000")
check("aquis: nao redundante -> 1 UTR (so PRI)", len(saida_aq["utr"]) == 1
      and saida_aq["utr"][0]["ID"] == "GRD_PRI")
check("aquis: ENU sempre em par PRI/REV mesmo sem redundancia", len(saida_aq["enu"]) == 2)
check("aquis: 1 TAC, nenhum TDD", len(saida_aq["tac"]) == 1 and len(saida_aq["tdd"]) == 0)
check("aquis: TAC.LSC=GRD", saida_aq["tac"][0]["LSC"] == "GRD")
check("aquis: 2 NV1 (leitura A104 + comando C104)", len(saida_aq["nv1"]) == 2)
tns1_aq = {n["TN1"] for n in saida_aq["nv1"]}
check("aquis: TN1 = A104 e C104", tns1_aq == {"A104", "C104"})
check("aquis: 4 NV2 (3 leitura + 1 comando)", len(saida_aq["nv2"]) == 4)
tns2_aq = {n["TN2"] for n in saida_aq["nv2"]}
check("aquis: TN2 = ASIM/ADUP/APFL/CDUP", tns2_aq == {"ASIM", "ADUP", "APFL", "CDUP"})
nv2_cdup = next(n for n in saida_aq["nv2"] if n["TN2"] == "CDUP")
check("aquis: NV2 CDUP tem TPPNT=CGF", nv2_cdup["TPPNT"] == "CGF")

# ------------------------------------------------------------------
# 3. Aquisicao 104, redundante
# ------------------------------------------------------------------
linha_aq_red = linha_ied(ID="GRD2", Protocolo="104", Direcao="Aquisicao",
                          PlPr="7", LiPr="7", PlRe="8", LiRe="8", Redundante="S", Gera="x")
saida_aq_red = mod._gerar_infra_ied(linha_aq_red, headers)
check("aquis redundante: 2 UTR (PRI+REV)", len(saida_aq_red["utr"]) == 2)
check("aquis redundante: IDs GRD2_PRI e GRD2_REV",
      {u["ID"] for u in saida_aq_red["utr"]} == {"GRD2_PRI", "GRD2_REV"})

# ------------------------------------------------------------------
# 4. Distribuicao 104
# ------------------------------------------------------------------
linha_dist = linha_ied(ID="COT", Protocolo="104", Direcao="Distribuicao",
                        PlPr="5", LiPr="5", PlRe="6", LiRe="6", Gera="x")
saida_dist = mod._gerar_infra_ied(linha_dist, headers)
check("dist: LSC.TIPO=DD", saida_dist["lsc"][0]["TIPO"] == "DD")
check("dist: CNF.CONFIG SEM IGNERS/SINCR/INVAL", "IGNERS=" not in saida_dist["cnf"][0]["CONFIG"])
check("dist: CNF.CONFIG com PlPr=5...LiRe=6", "PlPr= 5" in saida_dist["cnf"][0]["CONFIG"]
      and "LiRe= 6" in saida_dist["cnf"][0]["CONFIG"])
check("dist: CXU.INTGR usa default de distribuicao (50)", saida_dist["cxu"][0]["INTGR"] == "50")
check("dist: 2 TDD (DIG+ANA), nenhum TAC", len(saida_dist["tdd"]) == 2 and len(saida_dist["tac"]) == 0)
check("dist: TDD IDs corretos", {t["ID"] for t in saida_dist["tdd"]} == {"COT_DIG", "COT_ANA"})
tns1_dist = {n["TN1"] for n in saida_dist["nv1"]}
check("dist: TN1 = D104 e O104", tns1_dist == {"D104", "O104"})

# ------------------------------------------------------------------
# 5. Override manual vence o default
# ------------------------------------------------------------------
linha_override = linha_ied(ID="X1", Protocolo="104", Direcao="Aquisicao",
                            AQANL="9999", MAP="MEUMAP", Gera="x")
saida_override = mod._gerar_infra_ied(linha_override, headers)
check("override: AQANL customizado vence o default", saida_override["cxu"][0]["AQANL"] == "9999")
check("override: MAP customizado vence o default", saida_override["lsc"][0]["MAP"] == "MEUMAP")
check("override: NSRV1 continua no default (localhost)", saida_override["lsc"][0]["NSRV1"] == "localhost")

# ------------------------------------------------------------------
# 6. _prefixo_tn1 direto
# ------------------------------------------------------------------
check("prefixo aquis leitura", mod._prefixo_tn1("104", "Aquisicao", "leitura") == "A104")
check("prefixo aquis comando", mod._prefixo_tn1("104", "Aquisicao", "comando") == "C104")
check("prefixo dist leitura", mod._prefixo_tn1("104", "Distribuicao", "leitura") == "D104")
check("prefixo dist comando", mod._prefixo_tn1("104", "Distribuicao", "comando") == "O104")
check("prefixo com sufixo != nome do protocolo (caso DNP3)",
      mod._prefixo_tn1("DNP", "Aquisicao", "leitura") == "ADNP")

# ------------------------------------------------------------------
# 7. Protocolo 101 (confirmado contra base real do usuario, SE Miracema/neoenergia)
# ------------------------------------------------------------------
linha_101_aq = linha_ied(ID="NEOA", Protocolo="101", Direcao="Aquisicao",
                          PlPr="2", LiPr="1", PlRe="2", LiRe="3", Gera="x")
saida_101_aq = mod._gerar_infra_ied(linha_101_aq, headers)
check("101 aquis: LSC.TCV=CNVG/TTP=IEC2S", saida_101_aq["lsc"][0]["TCV"] == "CNVG"
      and saida_101_aq["lsc"][0]["TTP"] == "IEC2S")
tns1_101 = {n["TN1"] for n in saida_101_aq["nv1"]}
check("101 aquis: TN1 = A101 e C101", tns1_101 == {"A101", "C101"})

linha_101_dist = linha_ied(ID="NEOD", Protocolo="101", Direcao="Distribuicao",
                            PlPr="2", LiPr="2", PlRe="2", LiRe="4", Gera="x")
saida_101_dist = mod._gerar_infra_ied(linha_101_dist, headers)
check("101 dist: LSC.TIPO=DD", saida_101_dist["lsc"][0]["TIPO"] == "DD")
check("101 dist: CNF.CONFIG sem IGNERS (mesmo formato do 104)",
      "IGNERS=" not in saida_101_dist["cnf"][0]["CONFIG"])
tns1_101d = {n["TN1"] for n in saida_101_dist["nv1"]}
check("101 dist: TN1 = D101 e O101", tns1_101d == {"D101", "O101"})

# ------------------------------------------------------------------
# 8. Protocolo DNP3 (confirmado contra base real, ctl_dnp_mdb/DJ9E539 -- so
# aquisicao; distribuicao extrapolada por consistencia, ver PARAMS_PROTOCOLO)
# ------------------------------------------------------------------
linha_dnp_aq = linha_ied(ID="DJ9E539", Protocolo="DNP3", Direcao="Aquisicao",
                          PlPr="4", LiPr="10", PlRe="0", LiRe="0", Gera="x")
saida_dnp_aq = mod._gerar_infra_ied(linha_dnp_aq, headers)
check("DNP3 aquis: LSC.TCV=CNVH/TTP=IEC3S", saida_dnp_aq["lsc"][0]["TCV"] == "CNVH"
      and saida_dnp_aq["lsc"][0]["TTP"] == "IEC3S")
check("DNP3 aquis: CNF.CONFIG usa TZBR/DnpLvl (nao IGNERS/SINCR/INVAL)",
      "TZBR=" in saida_dnp_aq["cnf"][0]["CONFIG"] and "DnpLvl=" in saida_dnp_aq["cnf"][0]["CONFIG"]
      and "IGNERS=" not in saida_dnp_aq["cnf"][0]["CONFIG"])
check("DNP3 aquis: PlPr/LiPr/PlRe/LiRe vem ANTES de TZBR/DnpLvl (ordem real observada)",
      saida_dnp_aq["cnf"][0]["CONFIG"].index("PlPr=") < saida_dnp_aq["cnf"][0]["CONFIG"].index("TZBR="))
tns1_dnp = {n["TN1"] for n in saida_dnp_aq["nv1"]}
check("DNP3 aquis: TN1 = ADNP e CDNP (sufixo 'DNP', nao 'DNP3')", tns1_dnp == {"ADNP", "CDNP"})
tns2_dnp = {n["TN2"] for n in saida_dnp_aq["nv2"]}
check("DNP3 aquis: TN2 analogico = AANL (nao APFL)", "AANL" in tns2_dnp and "APFL" not in tns2_dnp)

# Distribuicao confirmada contra base real (ctl/COGTXA21, Drive/Projetos/
# _scada/DNP3-MDB.zip) -- 3 diferencas do 104/101, corrigidas depois de terem
# sido extrapoladas incorretamente na 1a versao desta entrada.
linha_dnp_dist = linha_ied(ID="COGTXA21", Protocolo="DNP3", Direcao="Distribuicao",
                            PlPr="2", LiPr="5", PlRe="2", LiRe="6", Gera="x")
saida_dnp_dist = mod._gerar_infra_ied(linha_dnp_dist, headers)
check("DNP3 dist: LSC.TTP=UDPF3 (NAO IEC3S da aquisicao, mesmo TCV=CNVH)",
      saida_dnp_dist["lsc"][0]["TTP"] == "UDPF3" and saida_dnp_dist["lsc"][0]["TCV"] == "CNVH")
check("DNP3 dist: CNF.CONFIG TEM TZBR/DnpLvl tambem (achado real, diferente do 104/101)",
      "TZBR=" in saida_dnp_dist["cnf"][0]["CONFIG"] and "DnpLvl=" in saida_dnp_dist["cnf"][0]["CONFIG"])
check("DNP3 dist: 1 TDD so (sem split _DIG/_ANA como 104/101)",
      len(saida_dnp_dist["tdd"]) == 1 and saida_dnp_dist["tdd"][0]["ID"] == "COGTXA21")
tns1_dnpd = {n["TN1"] for n in saida_dnp_dist["nv1"]}
check("DNP3 dist: TN1 = DDNP e ODNP", tns1_dnpd == {"DDNP", "ODNP"})
tns2_dnpd = {n["TN2"] for n in saida_dnp_dist["nv2"]}
check("DNP3 dist: comando roteia CDUP E CSIM juntos (aquisicao so tem CDUP)",
      {"CDUP", "CSIM"} <= tns2_dnpd)

# ------------------------------------------------------------------
# 9. Protocolo MODBUS (confirmado contra base real, mdb_alat_calc/MDB1 -- so
# aquisicao, sem comando; distribuicao extrapolada por consistencia, sem base
# real disponivel, mesma ressalva do DNP3)
# ------------------------------------------------------------------
linha_mdb_aq = linha_ied(ID="MDB1", Protocolo="MODBUS", Direcao="Aquisicao", Nome="Aquisicao MDB1 ModBus",
                         INS="ARA2", PlPr="1", LiPr="1", PlRe="0", LiRe="0", Gera="x")
saida_mdb_aq = mod._gerar_infra_ied(linha_mdb_aq, headers)
check("MODBUS aquis: LSC.TCV=CNVJ/TTP=TMBUS", saida_mdb_aq["lsc"][0]["TCV"] == "CNVJ"
      and saida_mdb_aq["lsc"][0]["TTP"] == "TMBUS")
check("MODBUS aquis: CNF.CONFIG usa PROTO (nao IGNERS/TZBR)",
      "PROTO=" in saida_mdb_aq["cnf"][0]["CONFIG"] and "IGNERS=" not in saida_mdb_aq["cnf"][0]["CONFIG"]
      and "TZBR=" not in saida_mdb_aq["cnf"][0]["CONFIG"])
check("MODBUS aquis: PROTO default BIN quando nao informado",
      "PROTO= BIN" in saida_mdb_aq["cnf"][0]["CONFIG"])
check("MODBUS aquis: PlPr/LiPr/PlRe/LiRe vem ANTES de PROTO (mesma ordem do DNP3)",
      saida_mdb_aq["cnf"][0]["CONFIG"].index("PlPr=") < saida_mdb_aq["cnf"][0]["CONFIG"].index("PROTO="))
tns1_mdb = {n["TN1"] for n in saida_mdb_aq["nv1"]}
check("MODBUS aquis: TN1 = AMDB e CMDB (sufixo 'MDB', nao 'MODBUS')", tns1_mdb == {"AMDB", "CMDB"})
tns2_mdb = {n["TN2"] for n in saida_mdb_aq["nv2"]}
check("MODBUS aquis: TN2 leitura = ALAT/AANL/ASTP (sem ASIM/ADUP)",
      tns2_mdb == {"ALAT", "AANL", "ASTP", "CDUP"})
nv2_alat = next(n for n in saida_mdb_aq["nv2"] if n["TN2"] == "ALAT")
check("MODBUS aquis: NV2 ALAT tem TPPNT=PDF (digital)", nv2_alat["TPPNT"] == "PDF")
check("MODBUS aquis: 2 TAC (ASAC + AFIL)", len(saida_mdb_aq["tac"]) == 2)
tac_ids_mdb = {t["ID"] for t in saida_mdb_aq["tac"]}
check("MODBUS aquis: TAC IDs = MDB1 e MDB1_FIL", tac_ids_mdb == {"MDB1", "MDB1_FIL"})
tac_tpaqs_mdb = {t["TPAQS"] for t in saida_mdb_aq["tac"]}
check("MODBUS aquis: TAC TPAQS = ASAC e AFIL", tac_tpaqs_mdb == {"ASAC", "AFIL"})
check("MODBUS aquis: INS propagado pros 2 TAC", all(t["INS"] == "ARA2" for t in saida_mdb_aq["tac"]))

linha_mdb_dist = linha_ied(ID="DMDB", Protocolo="MODBUS", Direcao="Distribuicao",
                           PlPr="1", LiPr="1", PlRe="0", LiRe="0", Gera="x")
saida_mdb_dist = mod._gerar_infra_ied(linha_mdb_dist, headers)
check("MODBUS dist: CNF.CONFIG sem PROTO (so na aquisicao)", "PROTO=" not in saida_mdb_dist["cnf"][0]["CONFIG"])
check("MODBUS dist: 2 TDD, nenhum TAC", len(saida_mdb_dist["tdd"]) == 2 and len(saida_mdb_dist["tac"]) == 0)
tns1_mdbd = {n["TN1"] for n in saida_mdb_dist["nv1"]}
check("MODBUS dist: TN1 = DMDB e OMDB", tns1_mdbd == {"DMDB", "OMDB"})

# ------------------------------------------------------------------
# 10. Regressao: 104/101/DNP3 continuam intactos apos generalizar
# grupos_leitura/grupos_comando/tacs (nao regrediram pro refactor do MODBUS)
# ------------------------------------------------------------------
check("104 aquis ainda com TAC simples (sem sufixo)", saida_aq["tac"][0]["ID"] == "GRD")
check("104 aquis ainda com TN2 leitura = ASIM/ADUP/APFL (nao ALAT/AANL/ASTP)",
      tns2_aq == {"ASIM", "ADUP", "APFL", "CDUP"})
check("DNP3 aquis ainda com TN2 analogico = AANL (regressao pos-refactor)",
      "AANL" in tns2_dnp and "ALAT" not in tns2_dnp)

# ------------------------------------------------------------------
# 11. Protocolo 61850 (confirmado contra base real, par/CTEEP -- 12 IEDs reais,
# 100% consistentes). Modelo BEM diferente: 1 linha = IED completo (aquisicao E
# distribuicao juntos), sem CXU/UTR/ENU, TN1 fixo NLN1.
# ------------------------------------------------------------------
linha_61850 = linha_ied(ID="UPCP_TR3", Protocolo="61850", Direcao="", Nome="UPCP_TR3 - 138",
                        GSD="PAR", INS="PAR", Gera="x")
saida_61850 = mod._gerar_infra_ied(linha_61850, headers)
check("61850: 1 LSC", len(saida_61850["lsc"]) == 1)
check("61850: LSC.TCV=CNVO/TTP=MMST/TIPO=AD (sempre AD, nao AA/DD)",
      saida_61850["lsc"][0]["TCV"] == "CNVO" and saida_61850["lsc"][0]["TTP"] == "MMST"
      and saida_61850["lsc"][0]["TIPO"] == "AD")
check("61850: LSC.VERBD=SCL_AUTO", saida_61850["lsc"][0]["VERBD"] == "SCL_AUTO")
check("61850: CNF.CONFIG usa OPMSK (default 228521), nao PlPr/LiPr",
      "OPMSK= 228521" in saida_61850["cnf"][0]["CONFIG"] and "PlPr=" not in saida_61850["cnf"][0]["CONFIG"])
check("61850: CNF.CONFIG usa GOOSE (default 0)", "GOOSE= 0" in saida_61850["cnf"][0]["CONFIG"])
check("61850: 1 TAC e 1 TDD (sempre os dois, mesmo ID do LSC)",
      len(saida_61850["tac"]) == 1 and len(saida_61850["tdd"]) == 1
      and saida_61850["tac"][0]["ID"] == "UPCP_TR3" and saida_61850["tdd"][0]["ID"] == "UPCP_TR3")
check("61850: TAC.INS propagado", saida_61850["tac"][0]["INS"] == "PAR")
check("61850: SEM CXU/UTR/ENU (nao usa essa camada)",
      len(saida_61850["cxu"]) == 0 and len(saida_61850["utr"]) == 0 and len(saida_61850["enu"]) == 0)
check("61850: 1 MUL, ID=CNF=ID_IED (achado real, diferente do sufixo _AQ do ICCP)",
      len(saida_61850["mul"]) == 1 and saida_61850["mul"][0]["ID"] == "UPCP_TR3"
      and saida_61850["mul"][0]["CNF"] == "UPCP_TR3" and saida_61850["mul"][0]["GSD"] == "PAR")
check("61850: SEMPRE 2 ENM (nao condicional a Redundante, confirmado 90/90 na base real)",
      len(saida_61850["enm"]) == 2
      and {e["ID"] for e in saida_61850["enm"]} == {"UPCP_TR31", "UPCP_TR32"}
      and all(e["MUL"] == "UPCP_TR3" for e in saida_61850["enm"]))
check("61850: 1 NV1 so, TN1=NLN1 fixo", len(saida_61850["nv1"]) == 1
      and saida_61850["nv1"][0]["TN1"] == "NLN1")
tns2_61850 = {n["TN2"] for n in saida_61850["nv2"]}
check("61850: TN2 = ADAQ/AAAQ/CSIM (nao ASIM/ADUP/CDUP dos outros protocolos)",
      tns2_61850 == {"ADAQ", "AAAQ", "CSIM"})
nv2_adaq = next(n for n in saida_61850["nv2"] if n["TN2"] == "ADAQ")
check("61850: NV2 ADAQ tem TPPNT=PDF", nv2_adaq["TPPNT"] == "PDF")
nv2_csim = next(n for n in saida_61850["nv2"] if n["TN2"] == "CSIM")
check("61850: NV2 CSIM tem TPPNT=CGF", nv2_csim["TPPNT"] == "CGF")

# Override manual do OPMSK vence o default (ex.: IED virtual redundante, bit 12)
linha_61850_virtual = linha_ied(ID="UCD1_CNF", Protocolo="61850", OPMSK="8010", Gera="x")
saida_61850_virtual = mod._gerar_infra_ied(linha_61850_virtual, headers)
check("61850: override de OPMSK vence o default (caso IED virtual, bit 12)",
      "OPMSK= 8010" in saida_61850_virtual["cnf"][0]["CONFIG"])

# ------------------------------------------------------------------
# 12. Protocolo SNMP (confirmado contra 2 bases reais independentes, par/CTEEP
# e jdm/CHESF, 100% consistentes entre si). Cabe no caminho padrao (LSC.TIPO
# segue Direcao, tem CXU/UTR/ENU), so com CNF.CONFIG/TN1/ENUTR/sem-comando
# diferentes.
# ------------------------------------------------------------------
linha_snmp = linha_ied(ID="SWA1", Protocolo="SNMP", Direcao="Aquisicao", Nome="LIGACAO SNMP em SWA1",
                       GSD="PAR", INS="PAR", HOST="172.30.45.71", Gera="x")
saida_snmp = mod._gerar_infra_ied(linha_snmp, headers)
check("SNMP: LSC.TCV=CNVI/TTP=TSNMP/TIPO=AA (segue Direcao, igual aos 4 classicos)",
      saida_snmp["lsc"][0]["TCV"] == "CNVI" and saida_snmp["lsc"][0]["TTP"] == "TSNMP"
      and saida_snmp["lsc"][0]["TIPO"] == "AA")
check("SNMP: CNF.CONFIG = VERSAO/HOST/COMMUNITY (nao PlPr/LiPr)",
      "VERSAO= 2c" in saida_snmp["cnf"][0]["CONFIG"] and "HOST= 172.30.45.71" in saida_snmp["cnf"][0]["CONFIG"]
      and "COMMUNITY= public" in saida_snmp["cnf"][0]["CONFIG"] and "PlPr=" not in saida_snmp["cnf"][0]["CONFIG"])
check("SNMP: 1 TAC (aquisicao), nenhum TDD", len(saida_snmp["tac"]) == 1 and len(saida_snmp["tdd"]) == 0)
check("SNMP: TEM CXU/UTR/ENU (ao contrario do 61850)",
      len(saida_snmp["cxu"]) == 1 and len(saida_snmp["utr"]) >= 1 and len(saida_snmp["enu"]) == 2)
check("SNMP: UTR nao-redundante -> 1 so (PRI)", len(saida_snmp["utr"]) == 1
      and saida_snmp["utr"][0]["ORDEM"] == "PRI")
check("SNMP: ENUTR=1 no PRI (nao '9' dos outros protocolos)", saida_snmp["utr"][0]["ENUTR"] == "1")
check("SNMP: 1 NV1 so (sem grupo de comando)", len(saida_snmp["nv1"]) == 1)
check("SNMP: TN1 fixo = SNM1 (sem prefixo A/C/D/O)", saida_snmp["nv1"][0]["TN1"] == "SNM1")
tns2_snmp = {n["TN2"] for n in saida_snmp["nv2"]}
check("SNMP: TN2 = ASIM so (1 grupo, digital)", tns2_snmp == {"ASIM"})
check("SNMP: NV2 ASIM tem TPPNT=PDF", saida_snmp["nv2"][0]["TPPNT"] == "PDF")

linha_snmp_red = linha_ied(ID="SWB1", Protocolo="SNMP", Direcao="Aquisicao",
                           HOST="172.30.45.72", Redundante="S", Gera="x")
saida_snmp_red = mod._gerar_infra_ied(linha_snmp_red, headers)
check("SNMP redundante: 2 UTR (PRI+REV), ENUTR 1 e 0", len(saida_snmp_red["utr"]) == 2
      and {u["ENUTR"] for u in saida_snmp_red["utr"]} == {"1", "0"})

# ------------------------------------------------------------------
# 13. Protocolo ICCP (confirmado contra o manual oficial SAGE Anx15, CEPEL --
# nao contra base real, o SkillSAGE nao tem nenhuma). Modelo BEM diferente:
# sem CXU/UTR/ENU/TAC/TDD, usa MUL+ENM; 1 unico NV1 com ate 8 tipos de NV2
# (aquisicao E distribuicao no mesmo canal).
# ------------------------------------------------------------------
linha_iccp = linha_ied(ID="A2B", Protocolo="ICCP", Nome="Ligacao ICCP Centro A-B",
                       GSD="GT1", VERBD="A2B_2024", Gera="x")
saida_iccp = mod._gerar_infra_ied(linha_iccp, headers)
check("ICCP: 1 LSC", len(saida_iccp["lsc"]) == 1)
check("ICCP: LSC.TCV=CNVN/TTP=MMST/TIPO=AD (bidirecional, igual ao 61850)",
      saida_iccp["lsc"][0]["TCV"] == "CNVN" and saida_iccp["lsc"][0]["TTP"] == "MMST"
      and saida_iccp["lsc"][0]["TIPO"] == "AD")
check("ICCP: LSC.VERBD = Acordo Bilateral informado", saida_iccp["lsc"][0]["VERBD"] == "A2B_2024")
check("ICCP: LSC.NSERV1 default derivado do ID quando nao informado",
      saida_iccp["lsc"][0]["NSERV1"] == "A2B_SRV1")
check("ICCP: LSC.NSERV2 vazio quando nao redundante", saida_iccp["lsc"][0]["NSERV2"] == "")
check("ICCP: CNF.CONFIG tem ApTitle (default compartilhado com 61850)",
      "ApTitle= 1 1 10 / 1 1 10" in saida_iccp["cnf"][0]["CONFIG"])
check("ICCP: CNF.CONFIG tem OPMSK=0 (default proprio do ICCP, NAO 228521 do 61850)",
      "OPMSK= 0" in saida_iccp["cnf"][0]["CONFIG"] and "OPMSK= 228521" not in saida_iccp["cnf"][0]["CONFIG"])
check("ICCP: CNF.CONFIG tem T2V=0 e BLC3=0 (defaults do manual)",
      "T2V= 0" in saida_iccp["cnf"][0]["CONFIG"] and "BLC3= 0" in saida_iccp["cnf"][0]["CONFIG"])
check("ICCP: CNF.CONFIG NAO tem PlPr/LiPr nem VERSAO/HOST (nao e 60870 nem SNMP)",
      "PlPr=" not in saida_iccp["cnf"][0]["CONFIG"] and "VERSAO=" not in saida_iccp["cnf"][0]["CONFIG"])
check("ICCP: SEM CXU/UTR/ENU/TAC/TDD (usa MUL/ENM no lugar)",
      len(saida_iccp["cxu"]) == 0 and len(saida_iccp["utr"]) == 0 and len(saida_iccp["enu"]) == 0
      and len(saida_iccp["tac"]) == 0 and len(saida_iccp["tdd"]) == 0)
check("ICCP: 1 MUL, ID = <ID>_AQ", len(saida_iccp["mul"]) == 1 and saida_iccp["mul"][0]["ID"] == "A2B_AQ")
check("ICCP: MUL.CNF aponta pro CNF do canal", saida_iccp["mul"][0]["CNF"] == "A2B")
check("ICCP: 1 ENM so (nao redundante), ID = NSERV1", len(saida_iccp["enm"]) == 1
      and saida_iccp["enm"][0]["ID"] == "A2B_SRV1" and saida_iccp["enm"][0]["MUL"] == "A2B_AQ")
check("ICCP: 1 NV1 so, TN1=NLN1", len(saida_iccp["nv1"]) == 1 and saida_iccp["nv1"][0]["TN1"] == "NLN1")
tns2_iccp = {n["TN2"] for n in saida_iccp["nv2"]}
check("ICCP: 8 tipos de NV2 (aquisicao E distribuicao juntos)", tns2_iccp == {
    "ADAQ", "AAAQ", "ATTA", "CSIM", "DDAQ", "DAAQ", "DTTA", "CDUP"})
nv2_ddaq = next(n for n in saida_iccp["nv2"] if n["TN2"] == "DDAQ")
check("ICCP: NV2 DDAQ (distribuicao digital) tem TPPNT=PDF (mesmo tipo fisico da aquisicao)",
      nv2_ddaq["TPPNT"] == "PDF")
nv2_cdup = next(n for n in saida_iccp["nv2"] if n["TN2"] == "CDUP")
check("ICCP: NV2 CDUP (roteamento de controle) tem TPPNT=CGF", nv2_cdup["TPPNT"] == "CGF")

linha_iccp_red = linha_ied(ID="C2D", Protocolo="ICCP", Redundante="S", Gera="x")
saida_iccp_red = mod._gerar_infra_ied(linha_iccp_red, headers)
check("ICCP redundante: 2 ENM (servidor principal + reserva)", len(saida_iccp_red["enm"]) == 2)
check("ICCP redundante: LSC.NSERV2 preenchido", saida_iccp_red["lsc"][0]["NSERV2"] == "C2D_SRV2")

# ------------------------------------------------------------------
# 14. Extracao reversa de IEDs (_extrair_ieds) -- unitarios diretos
# ------------------------------------------------------------------
check("_protocolo_e_direcao_por_lsc: 104 aquis (TIPO=AA)",
      mod._protocolo_e_direcao_por_lsc("CNVM", "CX104", "AA") == ("104", "Aquisicao"))
check("_protocolo_e_direcao_por_lsc: 104 dist (TIPO=DD)",
      mod._protocolo_e_direcao_por_lsc("CNVM", "CX104", "DD") == ("104", "Distribuicao"))
check("_protocolo_e_direcao_por_lsc: DNP3 dist via TTP=UDPF3 (nao precisa olhar TIPO)",
      mod._protocolo_e_direcao_por_lsc("CNVH", "UDPF3", "DD") == ("DNP3", "Distribuicao"))
check("_protocolo_e_direcao_por_lsc: DNP3 aquis via TTP=IEC3S (TIPO desempata)",
      mod._protocolo_e_direcao_por_lsc("CNVH", "IEC3S", "AA") == ("DNP3", "Aquisicao"))
check("_protocolo_e_direcao_por_lsc: 61850 sempre Direcao vazia (bidirecional)",
      mod._protocolo_e_direcao_por_lsc("CNVO", "MMST", "AD") == ("61850", ""))
check("_protocolo_e_direcao_por_lsc: ICCP sempre Direcao vazia (bidirecional)",
      mod._protocolo_e_direcao_por_lsc("CNVN", "MMST", "AD") == ("ICCP", ""))
check("_protocolo_e_direcao_por_lsc: TCV/TTP desconhecido -> None,None (protocolo nao modelado)",
      mod._protocolo_e_direcao_por_lsc("CNVX", "XXXX", "AA") == (None, None))

check("_parsear_config_cnf: valor multi-token (ApTitle, achado real 61850/ICCP)",
      mod._parsear_config_cnf("ApTitle= 1 1 10 / 1 1 10 AeQ= 1", ["ApTitle", "AeQ"])
      == {"ApTitle": "1 1 10 / 1 1 10", "AeQ": "1"})
check("_parsear_config_cnf: campo ausente no texto nao aparece no resultado",
      "SINCR" not in mod._parsear_config_cnf("PlPr= 7 LiPr= 7", ["PlPr", "LiPr", "SINCR"]))
check("_parsear_config_cnf: campo com valor vazio no meio (caso ICCP sem IDIG/IANL/IDIS)",
      mod._parsear_config_cnf("IDIG=  IANL=  IDIS=  TOUT= 10", ["IDIG", "IANL", "IDIS", "TOUT"])
      == {"IDIG": "", "IANL": "", "IDIS": "", "TOUT": "10"})


# ------------------------------------------------------------------
# 15. Extracao reversa de IEDs (_extrair_ieds) -- round-trip forward->reverso
# em cima de TODAS as saidas de _gerar_infra_ied* ja construidas acima (secoes
# 2 a 13): gera pra frente, empacota a saida como se fosse entidade ja
# importada (headers = uniao das chaves + "Gera"=x sempre), roda o reverso e
# confere que os campos originais da linha de IEDs sao reconstruidos.
# ------------------------------------------------------------------
def _dicts_para_entidade(lista_dicts):
    """[{...}, ...] (saida de _gerar_infra_ied*) -> (headers, linhas) como se
    fosse uma aba de entidade ja importada -- 'Gera'=x sempre (sem isso,
    _linhas_ativas_como_dicts descartaria tudo por falta da coluna de controle)."""
    if not lista_dicts:
        return (None, None)
    campos = sorted({chave for d in lista_dicts for chave in d.keys()})
    colunas = campos + ["Gera"]
    linhas = [[d.get(c, "") for c in campos] + ["x"] for d in lista_dicts]
    return (colunas, linhas)


def _entidades_de_saida(saida):
    """Empacota a saida de _gerar_infra_ied(_61850/_iccp) no formato 'entidades'
    esperado por _extrair_ieds: {nome: (headers, linhas)}."""
    return {nome: _dicts_para_entidade(linhas) for nome, linhas in saida.items()}


# --- 104 ---
ied_aq = by_id(mod._extrair_ieds(_entidades_de_saida(saida_aq)), "GRD")
check("round-trip 104 aquis: Protocolo/Direcao", ied_aq["Protocolo"] == "104" and ied_aq["Direcao"] == "Aquisicao")
check("round-trip 104 aquis: Nome/GSD preservados",
      ied_aq["Nome"] == "Ligacao 104 GRD" and ied_aq["GSD"] == "GT_SCD_1")
check("round-trip 104 aquis: PlPr/LiPr/PlRe/LiRe do CNF.CONFIG",
      (ied_aq["PlPr"], ied_aq["LiPr"], ied_aq["PlRe"], ied_aq["LiRe"]) == ("7", "7", "8", "8"))
check("round-trip 104 aquis: IGNERS/SINCR/INVAL tambem parseados (extras so na aquisicao)",
      (ied_aq["IGNERS"], ied_aq["SINCR"], ied_aq["INVAL"]) == ("0", "0", "103"))
check("round-trip 104 aquis: AQANL/INTGR do CXU", ied_aq["AQANL"] == "1000" and ied_aq["INTGR"] == "180000")
check("round-trip 104 aquis: NTENT/RESPT do UTR, TDESC/TRANS/VLUTR do ENU",
      ied_aq["NTENT"] == "4" and ied_aq["RESPT"] == "1500"
      and ied_aq["TDESC"] == "15" and ied_aq["TRANS"] == "12" and ied_aq["VLUTR"] == "0")
check("round-trip 104 aquis: nao redundante -> Redundante vazio", ied_aq["Redundante"] == "")

ied_aq_red = by_id(mod._extrair_ieds(_entidades_de_saida(saida_aq_red)), "GRD2")
check("round-trip 104 aquis redundante: Redundante=S inferido de 2 UTR (PRI+REV)",
      ied_aq_red["Redundante"] == "S")

ied_dist = by_id(mod._extrair_ieds(_entidades_de_saida(saida_dist)), "COT")
check("round-trip 104 dist: Direcao=Distribuicao (via TIPO=DD)", ied_dist["Direcao"] == "Distribuicao")
check("round-trip 104 dist: PlPr/LiPr/PlRe/LiRe do CNF.CONFIG",
      (ied_dist["PlPr"], ied_dist["LiPr"], ied_dist["PlRe"], ied_dist["LiRe"]) == ("5", "5", "6", "6"))
check("round-trip 104 dist: SEM IGNERS/SINCR/INVAL (nao existem no CONFIG de distribuicao)",
      "IGNERS" not in ied_dist and "SINCR" not in ied_dist and "INVAL" not in ied_dist)
check("round-trip 104 dist: INTGR usa o valor de distribuicao (50)", ied_dist["INTGR"] == "50")
check("round-trip 104 dist: sem TAC -> INS nao aparece", "INS" not in ied_dist)

# --- 101 ---
ied_101_aq = by_id(mod._extrair_ieds(_entidades_de_saida(saida_101_aq)), "NEOA")
check("round-trip 101 aquis: Protocolo/Direcao/PlPr..LiRe",
      ied_101_aq["Protocolo"] == "101" and ied_101_aq["Direcao"] == "Aquisicao"
      and (ied_101_aq["PlPr"], ied_101_aq["LiPr"], ied_101_aq["PlRe"], ied_101_aq["LiRe"])
      == ("2", "1", "2", "3"))
ied_101_dist = by_id(mod._extrair_ieds(_entidades_de_saida(saida_101_dist)), "NEOD")
check("round-trip 101 dist: Direcao=Distribuicao, sem IGNERS",
      ied_101_dist["Direcao"] == "Distribuicao" and "IGNERS" not in ied_101_dist)

# --- DNP3 (o caso mais intrincado: TTP muda por direcao, extras tambem na distribuicao) ---
ied_dnp_aq = by_id(mod._extrair_ieds(_entidades_de_saida(saida_dnp_aq)), "DJ9E539")
check("round-trip DNP3 aquis: Protocolo/Direcao (via TTP=IEC3S)",
      ied_dnp_aq["Protocolo"] == "DNP3" and ied_dnp_aq["Direcao"] == "Aquisicao")
check("round-trip DNP3 aquis: PlPr/LiPr/PlRe/LiRe + TZBR/DnpLvl",
      (ied_dnp_aq["PlPr"], ied_dnp_aq["LiPr"], ied_dnp_aq["PlRe"], ied_dnp_aq["LiRe"]) == ("4", "10", "0", "0")
      and ied_dnp_aq["TZBR"] == "0" and ied_dnp_aq["DnpLvl"] == "2")

ied_dnp_dist = by_id(mod._extrair_ieds(_entidades_de_saida(saida_dnp_dist)), "COGTXA21")
check("round-trip DNP3 dist: Protocolo/Direcao reconhecidos via TTP=UDPF3 (nao TIPO)",
      ied_dnp_dist["Protocolo"] == "DNP3" and ied_dnp_dist["Direcao"] == "Distribuicao")
check("round-trip DNP3 dist: PlPr/LiPr/PlRe/LiRe + TZBR/DnpLvl TAMBEM presentes (achado real)",
      (ied_dnp_dist["PlPr"], ied_dnp_dist["LiPr"], ied_dnp_dist["PlRe"], ied_dnp_dist["LiRe"])
      == ("2", "5", "2", "6")
      and ied_dnp_dist["TZBR"] == "0" and ied_dnp_dist["DnpLvl"] == "2")

# --- MODBUS (INS propagado via TAC dual; sem PROTO na distribuicao) ---
ied_mdb_aq = by_id(mod._extrair_ieds(_entidades_de_saida(saida_mdb_aq)), "MDB1")
check("round-trip MODBUS aquis: PROTO do CNF.CONFIG", ied_mdb_aq["PROTO"] == "BIN")
check("round-trip MODBUS aquis: INS recuperado do TAC (dos 2, mesmo valor)", ied_mdb_aq["INS"] == "ARA2")
ied_mdb_dist = by_id(mod._extrair_ieds(_entidades_de_saida(saida_mdb_dist)), "DMDB")
check("round-trip MODBUS dist: sem PROTO (nao existe no CONFIG de distribuicao)", "PROTO" not in ied_mdb_dist)

# --- 61850 (o parser precisa acertar ApTitle/PS multi-token) ---
ied_61850 = by_id(mod._extrair_ieds(_entidades_de_saida(saida_61850)), "UPCP_TR3")
check("round-trip 61850: Protocolo/Direcao (bidirecional, Direcao vazia)",
      ied_61850["Protocolo"] == "61850" and ied_61850["Direcao"] == "")
check("round-trip 61850: Nome/GSD/INS preservados",
      ied_61850["Nome"] == "UPCP_TR3 - 138" and ied_61850["GSD"] == "PAR" and ied_61850["INS"] == "PAR")
check("round-trip 61850: ApTitle multi-token reconstruido inteiro",
      ied_61850["ApTitle"] == "1 1 10 / 1 1 10")
check("round-trip 61850: PS multi-token reconstruido inteiro (outro campo c/ '/')",
      ied_61850["PS"] == "1 / 1")
check("round-trip 61850: OPMSK/GOOSE (ultimo campo do CONFIG) corretos",
      ied_61850["OPMSK"] == "228521" and ied_61850["GOOSE"] == "0")
check("round-trip 61850: Redundante nao se aplica (chave ausente, sem efeito no forward)",
      "Redundante" not in ied_61850)

ied_61850_virtual = by_id(mod._extrair_ieds(_entidades_de_saida(saida_61850_virtual)), "UCD1_CNF")
check("round-trip 61850 virtual: override de OPMSK preservado (bit 12)",
      ied_61850_virtual["OPMSK"] == "8010")

# --- SNMP (cnf_campos totalmente customizado, sem PlPr/LiPr) ---
ied_snmp = by_id(mod._extrair_ieds(_entidades_de_saida(saida_snmp)), "SWA1")
check("round-trip SNMP: VERSAO/HOST/COMMUNITY do CNF.CONFIG customizado",
      ied_snmp["VERSAO"] == "2c" and ied_snmp["HOST"] == "172.30.45.71" and ied_snmp["COMMUNITY"] == "public")
check("round-trip SNMP: sem PlPr (cnf_campos substitui a base inteira)", "PlPr" not in ied_snmp)
check("round-trip SNMP: INS do TAC, Direcao=Aquisicao (via TIPO=AA)",
      ied_snmp["INS"] == "PAR" and ied_snmp["Direcao"] == "Aquisicao")
check("round-trip SNMP: nao redundante -> Redundante vazio", ied_snmp["Redundante"] == "")

ied_snmp_red = by_id(mod._extrair_ieds(_entidades_de_saida(saida_snmp_red)), "SWB1")
check("round-trip SNMP redundante: Redundante=S inferido de 2 UTR", ied_snmp_red["Redundante"] == "S")

# --- ICCP (MUL/ENM em vez de UTR pra Redundante; OPMSK default proprio) ---
ied_iccp = by_id(mod._extrair_ieds(_entidades_de_saida(saida_iccp)), "A2B")
check("round-trip ICCP: Protocolo/Direcao (bidirecional, Direcao vazia)",
      ied_iccp["Protocolo"] == "ICCP" and ied_iccp["Direcao"] == "")
check("round-trip ICCP: VERBD/NSERV1 preservados, NSERV2 vazio (nao redundante)",
      ied_iccp["VERBD"] == "A2B_2024" and ied_iccp["NSERV1"] == "A2B_SRV1" and ied_iccp["NSERV2"] == "")
check("round-trip ICCP: ApTitle multi-token (default compartilhado com 61850)",
      ied_iccp["ApTitle"] == "1 1 10 / 1 1 10")
check("round-trip ICCP: OPMSK=0 (default proprio do ICCP, NAO 228521 do 61850)", ied_iccp["OPMSK"] == "0")
check("round-trip ICCP: T2V/BLC3 do CONFIG (defaults do manual)",
      ied_iccp["T2V"] == "0" and ied_iccp["BLC3"] == "0")
check("round-trip ICCP: nao redundante -> Redundante vazio (so 1 ENM)", ied_iccp["Redundante"] == "")

ied_iccp_red = by_id(mod._extrair_ieds(_entidades_de_saida(saida_iccp_red)), "C2D")
check("round-trip ICCP redundante: Redundante=S inferido de 2 ENM", ied_iccp_red["Redundante"] == "S")

# --- CXU/CNF/LSC com IDs INDEPENDENTES (achado real, base jdm/CHESF: LSC.ID=
# "JDM", CNF.ID="L_JDM_CNF", CXU.ID="CA_JDM_CXU" -- nenhum dos 3 igual ao outro,
# diferente do forward daqui, que sempre usa o mesmo id_ied nos 3). A travessia
# tem que seguir LSC->CNF (por LSC.ID)->UTR (por UTR.CNF)->CXU/ENU (por
# UTR.CXU) -- nunca assumir que CXU/UTR/ENU repetem o ID do LSC. ---
entidades_ids_independentes = {
    "lsc": _dicts_para_entidade([{"ID": "JDM", "TCV": "CNVH", "TTP": "IEC3S", "TIPO": "AA",
                                   "NOME": "Ligacao Aquisicao DNP3 SE Jardim", "GSD": "JDM"}]),
    "cnf": _dicts_para_entidade([{"ID": "L_JDM_CNF", "LSC": "JDM",
                                   "CONFIG": "PlPr= 1 LiPr= 1 PlRe= 2 LiRe= 1 TZBR= 0 DnpLvl= 1"}]),
    "cxu": _dicts_para_entidade([{"ID": "CA_JDM_CXU", "GSD": "JDM", "ORDEM": "1",
                                   "AQANL": "200", "AQPOL": "600", "AQTOT": "60000", "INTGR": "60000",
                                   "NFAIL": "5", "SFAIL": "200", "FAILP": "0", "FAILR": "0"}]),
    "utr": _dicts_para_entidade([
        {"ID": "JDM_UTRP", "CNF": "L_JDM_CNF", "CXU": "CA_JDM_CXU", "ENUTR": "7",
         "NTENT": "3", "RESPT": "3000", "ORDEM": "PRI"},
        {"ID": "JDM_UTRR", "CNF": "L_JDM_CNF", "CXU": "CA_JDM_CXU", "ENUTR": "7",
         "NTENT": "3", "RESPT": "3000", "ORDEM": "REV"},
    ]),
    "enu": _dicts_para_entidade([
        {"ID": "JDM_CXU_ENUP", "CXU": "CA_JDM_CXU", "ORDEM": "PRI", "TDESC": "0", "TRANS": "0", "VLUTR": "1"},
        {"ID": "JDM_CXU_ENUR", "CXU": "CA_JDM_CXU", "ORDEM": "REV", "TDESC": "0", "TRANS": "0", "VLUTR": "17"},
    ]),
    "tac": _dicts_para_entidade([{"ID": "JDM", "NOME": "Terminal Aquisicao/Controle de JDM",
                                   "INS": "JDM", "TPAQS": "ASAC", "LSC": "JDM"}]),
}
ied_ids_indep = by_id(mod._extrair_ieds(entidades_ids_independentes), "JDM")
check("extracao real (IDs independentes): Protocolo/Direcao reconhecidos mesmo com CXU.ID != LSC.ID",
      ied_ids_indep is not None and ied_ids_indep["Protocolo"] == "DNP3" and ied_ids_indep["Direcao"] == "Aquisicao")
check("extracao real (IDs independentes): CNF.CONFIG parseado (PlPr/TZBR/DnpLvl)",
      ied_ids_indep["PlPr"] == "1" and ied_ids_indep["TZBR"] == "0" and ied_ids_indep["DnpLvl"] == "1")
check("extracao real (IDs independentes): AQANL/INTGR do CXU encontrados via UTR.CNF->UTR.CXU",
      ied_ids_indep["AQANL"] == "200" and ied_ids_indep["INTGR"] == "60000")
check("extracao real (IDs independentes): NTENT/RESPT do UTR, Redundante=S (2 UTR)",
      ied_ids_indep["NTENT"] == "3" and ied_ids_indep["RESPT"] == "3000" and ied_ids_indep["Redundante"] == "S")
check("extracao real (IDs independentes): TDESC/TRANS/VLUTR do ENU encontrados via CXU.ID real",
      ied_ids_indep["TDESC"] == "0" and ied_ids_indep["TRANS"] == "0" and ied_ids_indep["VLUTR"] == "1")
check("extracao real (IDs independentes): INS do TAC (via TAC.LSC, sempre foi por FK de verdade)",
      ied_ids_indep["INS"] == "JDM")

# --- protocolo desconhecido: LSC solto (nunca gerado por _gerar_infra_ied, mas
# pode existir numa base real de protocolo ainda nao modelado) -- ignorado ---
entidades_desconhecido = _dicts_para_entidade(
    [{"ID": "X103", "TCV": "CNVZ", "TTP": "ZZZ", "TIPO": "AA", "NOME": "Protocolo nao modelado"}])
check("round-trip: TCV/TTP nao reconhecido -> nao extrai nenhuma linha",
      mod._extrair_ieds({"lsc": entidades_desconhecido}) == [])

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
