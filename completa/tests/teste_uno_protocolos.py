# -*- coding: utf-8 -*-
"""Teste ponta-a-ponta via UNO real (LibreOffice headless) do assistente de
IED/protocolo, para todos os 7 protocolos entregues, contra uma CÓPIA
DESCARTÁVEL da planilha real (nunca escreve no arquivo rastreado -- ver
uno_harness.py).

A profundidade de cada campo/default já é coberta pelo smoke test em memória
(smoke_test_ied.py, ~108 checks). Este teste foca no que só o UNO real prova:
a aba IEDs/entidades sendo criadas corretamente pelo LibreOffice de verdade,
o upsert não sendo destrutivo com dados REAIS pré-existentes (mul/enm de
61850, com centenas de linhas), e a integração com unificar_pontos.

Requer 'soffice' no PATH. Roda com:
    python completa/tests/teste_uno_protocolos.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_harness import TesteUno  # noqa: E402

falhas = []


def check(nome, cond):
    status = "OK" if cond else "FALHOU"
    print(f"[{status}] {nome}")
    if not cond:
        falhas.append(nome)


with TesteUno(porta=2100) as t:
    # mul/enm já existem na planilha real com dados de 61850 (achado durante o
    # desenvolvimento do ICCP) -- guarda a contagem ANTES pra confirmar depois
    # que o upsert só ADICIONA, nunca mexe no que já existe.
    mul_antes = t.contar_linhas("MUL")
    enm_antes = t.contar_linhas("ENM")

    t.chamar_macro("gerar_ied")
    headers_ieds = t.headers_de("IEDs")
    colunas_por_protocolo = {
        "104/101/DNP3/MODBUS": ("PlPr", "LiPr", "PlRe", "LiRe"),
        "61850/ICCP (MMS)": ("ApTitle", "AeQ", "PS", "SS", "TS", "OPMSK"),
        "MODBUS": ("PROTO",),
        "SNMP": ("VERSAO", "HOST", "COMMUNITY"),
        "ICCP": ("VERBD", "NSERV1", "NSERV2", "IDIG", "IANL", "IDIS", "T2V", "BLC3"),
    }
    for grupo, campos in colunas_por_protocolo.items():
        check(f"aba IEDs tem as colunas de {grupo}", all(c in headers_ieds for c in campos))

    linha0 = t.proxima_linha_livre("IEDs")
    # (id, protocolo, direcao, campos extras) -- 1 caso por protocolo entregue
    casos = [
        ("GRD104T", "104", "Aquisicao", {"PlPr": "7", "LiPr": "7", "PlRe": "8", "LiRe": "8"}),
        ("NEO101T", "101", "Aquisicao", {"PlPr": "2", "LiPr": "1", "PlRe": "2", "LiRe": "3"}),
        ("DJ9DNPT", "DNP3", "Aquisicao", {"PlPr": "4", "LiPr": "10", "PlRe": "0", "LiRe": "0"}),
        ("COGTDNPT", "DNP3", "Distribuicao", {"PlPr": "2", "LiPr": "5", "PlRe": "2", "LiRe": "6"}),
        ("MDB1T", "MODBUS", "Aquisicao", {"PlPr": "1", "LiPr": "1", "PlRe": "0", "LiRe": "0"}),
        ("TR3T", "61850", "", {"GSD": "PAR"}),
        ("SWA1T", "SNMP", "Aquisicao", {"HOST": "172.30.45.71"}),
        ("A2BT", "ICCP", "", {"VERBD": "TESTE_CONSOLIDADO"}),
    ]
    for i, (id_ied, protocolo, direcao, extras) in enumerate(casos):
        valores = {"ID": id_ied, "Protocolo": protocolo, "Direcao": direcao, "Gera": "x"}
        valores.update(extras)
        t.escrever_linha("IEDs", linha0 + i, valores)
    t.chamar_macro("gerar_ied")

    lsc = t.ler_aba("LSC")
    cxu = t.ler_aba("CXU")
    tac = t.ler_aba("TAC")
    tdd = t.ler_aba("TDD")
    mul = t.ler_aba("MUL")
    enm = t.ler_aba("ENM")

    for id_ied, protocolo, _, _ in casos:
        existe = any(l.get("ID") == id_ied for l in lsc)
        check(f"UNO real: LSC criado ({protocolo}, ID={id_ied})", existe)

    # Distinções estruturais entre protocolos (confirmadas no smoke test em
    # memória, aqui só confirmamos que sobrevivem à passagem pelo LibreOffice real)
    check("UNO real: 61850 e ICCP NAO usam CXU (MMS bidirecional sem essa camada)",
          not any(c.get("ID") in ("TR3T", "A2BT") for c in cxu))
    check("UNO real: 104/101/DNP3/MODBUS/SNMP usam CXU normalmente",
          all(any(c.get("ID") == id_ied for c in cxu)
              for id_ied in ("GRD104T", "NEO101T", "DJ9DNPT", "MDB1T", "SWA1T")))
    check("UNO real: 61850 tem TAC E TDD (mesmo ID)",
          any(x.get("ID") == "TR3T" for x in tac) and any(x.get("ID") == "TR3T" for x in tdd))
    check("UNO real: ICCP NAO tem TAC nem TDD (usa MUL/ENM)",
          not any(x.get("ID") == "A2BT" for x in tac) and not any(x.get("ID") == "A2BT" for x in tdd))
    check("UNO real: ICCP criou MUL (A2BT_AQ) e ENM", any(m.get("ID") == "A2BT_AQ" for m in mul)
          and any(e.get("MUL") == "A2BT_AQ" for e in enm))
    check("UNO real: upsert em MUL/ENM só ADICIONOU (dados reais de 61850 pré-existentes intactos)",
          t.contar_linhas("MUL") == mul_antes + 1 and t.contar_linhas("ENM") == enm_antes + 1)

    # Integração ponta-a-ponta: NV2 do 104 (ASIM) -> ponto novo em PontoDigital -> PDF
    nv1 = t.ler_aba("NV1")
    nv2 = t.ler_aba("NV2")
    nv1_grd = next(n for n in nv1 if n.get("CNF") == "GRD104T" and n.get("TN1") == "A104")
    nv2_asim = next(n for n in nv2 if n.get("NV1") == nv1_grd["ID"] and n.get("TN2") == "ASIM")

    t.chamar_macro("unificar_pontos")
    linha_pd = t.proxima_linha_livre("PontoDigital")
    t.escrever_linha("PontoDigital", linha_pd, {
        "ID_Logico": "TESTE_CONSOL_PT001", "ID_Fisico": "TESTE_CONSOL_PT001_FIS",
        "NOME": "Ponto teste consolidado", "NV2": nv2_asim["ID"], "Gera": "x",
    })
    t.chamar_macro("unificar_pontos")
    pdf = t.ler_aba("PDF")
    check("integração: PDF criado a partir do NV2 do 104 (via unificar_pontos)",
          any(l.get("ID") == "TESTE_CONSOL_PT001_FIS" and l.get("NV2") == nv2_asim["ID"] for l in pdf))

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    sys.exit(1)
print("Todas as checagens do teste UNO real (protocolos) passaram.")
