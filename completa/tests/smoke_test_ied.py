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

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    raise SystemExit(1)
print("Todas as checagens do smoke test passaram.")
