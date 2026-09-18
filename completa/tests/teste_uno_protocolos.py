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

    # Aba de config recém-criada: o upsert precisa respeitar o cabeçalho canônico que o
    # _garantir_aba_config acabou de escrever. Antes ele adotava o cabeçalho mínimo
    # (Origem/Gera/Comentario/Include), escrevia mais estreito por cima e deixava as
    # colunas antigas sobrando -- a aba ficava com cabeçalho duplicado e as macros
    # passavam a ler uma 2ª coluna "Gera" sempre vazia (nenhuma linha era processada).
    if t.sheet_existe("IEDs"):
        t.doc.Sheets.removeByName("IEDs")
    t.chamar_macro("extrair_pontos")
    headers_zerado = t.headers_de("IEDs")
    repetidos = sorted({h for h in headers_zerado if headers_zerado.count(h) > 1})
    check(f"aba IEDs criada do zero não tem cabeçalho duplicado ({repetidos or 'nenhum'})",
          not repetidos)

    # Linha de base da extração reversa: quantos IEDs a planilha distribuída já
    # reconstrói sozinha, por protocolo. Os casos do teste entram depois, +1 cada.
    ieds_base = t.ler_aba("IEDs")
    n_61850_base = sum(1 for l in ieds_base if l.get("Protocolo") == "61850")
    n_snmp_base = sum(1 for l in ieds_base if l.get("Protocolo") == "SNMP")

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
    check("UNO real: 61850 TAMBÉM cria MUL/ENM (achado real, ID=CNF, sempre 2 ENM)",
          any(m.get("ID") == "TR3T" and m.get("CNF") == "TR3T" for m in mul)
          and {e.get("ID") for e in enm if e.get("MUL") == "TR3T"} == {"TR3T1", "TR3T2"})
    check("UNO real: upsert em MUL/ENM só ADICIONOU (dados reais pré-existentes intactos) "
          "-- +2 MUL (ICCP + 61850) e +3 ENM (1 do ICCP + 2 do 61850)",
          t.contar_linhas("MUL") == mul_antes + 2 and t.contar_linhas("ENM") == enm_antes + 3)

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

    # Extração reversa de IEDs (espelho de gerar_ied, ver _extrair_ieds): roda
    # extrair_pontos (mesmo entry point que também reconstrói PontoDigital/etc a
    # partir da base real inteira -- ~2900 PDF pré-existentes) e confirma que as
    # 8 linhas de IEDs recém-criadas acima são reconstruídas corretamente, entre
    # elas os 106 LSC/CNF de 61850 já reais da planilha (que não vieram de
    # gerar_ied nesta rodada) -- prova que o parser de CNF.CONFIG e o
    # reconhecimento de TCV/TTP aguentam dado de produção de verdade, não só
    # os casos sintéticos do smoke test em memória.
    # Round-trip do comando: o CGF vive num NV2 de COMANDO (EX1_CDNP_2_CDUP), não no de
    # leitura do ponto (EX1_ADNP_1_ASIM). A geração copiava o NV2 de leitura, então
    # extrair_pontos -> unificar_pontos movia todo comando da base pro grupo errado, sem
    # avisar. Guarda o mapa ID->NV2 de antes pra comparar depois do round-trip completo.
    cgf_antes = {l["ID"]: l.get("NV2", "") for l in t.ler_aba("CGF") if l.get("ID")}

    t.chamar_macro("extrair_pontos")
    ieds_apos = t.ler_aba("IEDs")

    def by_id_ieds(id_ied):
        return next((l for l in ieds_apos if l.get("ID") == id_ied), None)

    ied_grd = by_id_ieds("GRD104T")
    check("extração reversa (UNO real): 104 aquis reconstrói Protocolo/Direcao/PlPr..LiRe",
          ied_grd is not None and ied_grd.get("Protocolo") == "104" and ied_grd.get("Direcao") == "Aquisicao"
          and ied_grd.get("PlPr") == "7" and ied_grd.get("LiRe") == "8")
    ied_cogtdnpt = by_id_ieds("COGTDNPT")
    check("extração reversa (UNO real): DNP3 distribuição reconhecida via TTP=UDPF3 (não TIPO)",
          ied_cogtdnpt is not None and ied_cogtdnpt.get("Protocolo") == "DNP3"
          and ied_cogtdnpt.get("Direcao") == "Distribuicao")
    ied_mdb1t = by_id_ieds("MDB1T")
    check("extração reversa (UNO real): MODBUS reconstrói PROTO=BIN",
          ied_mdb1t is not None and ied_mdb1t.get("PROTO") == "BIN")
    ied_tr3t = by_id_ieds("TR3T")
    check("extração reversa (UNO real): 61850 reconstrói ApTitle multi-token + GSD/INS",
          ied_tr3t is not None and ied_tr3t.get("ApTitle") == "1 1 10 / 1 1 10" and ied_tr3t.get("GSD") == "PAR")
    ied_a2bt = by_id_ieds("A2BT")
    check("extração reversa (UNO real): ICCP reconstrói VERBD e OPMSK=0 próprio (não 228521 do 61850)",
          ied_a2bt is not None and ied_a2bt.get("VERBD") == "TESTE_CONSOLIDADO" and ied_a2bt.get("OPMSK") == "0")
    check("extração reversa (UNO real): upsert casou por ID, não duplicou nenhuma das 8 linhas escritas",
          sum(1 for l in ieds_apos if l.get("ID") in {c[0] for c in casos}) == len(casos))
    # A extração reversa não olha só as linhas criadas aqui: reconstrói também os LSC
    # que já vieram da base importada na planilha (contados no início) -- +1 de cada dos
    # casos de teste acima. Comparar com a linha de base, e não com um número fixo,
    # deixa o teste independente do tamanho da base que a planilha distribui.
    check(f"extração reversa (UNO real): manteve os {n_61850_base} LSC de 61850 da base (+1 do teste)",
          sum(1 for l in ieds_apos if l.get("Protocolo") == "61850") == n_61850_base + 1)
    check(f"extração reversa (UNO real): manteve os {n_snmp_base} LSC de SNMP da base (+1 do teste)",
          sum(1 for l in ieds_apos if l.get("Protocolo") == "SNMP") == n_snmp_base + 1)

    # Fecha o round-trip: com as abas de config reconstruídas pela extração, gerar de
    # novo não pode mexer no NV2 de nenhum comando que já existia na base.
    t.chamar_macro("unificar_pontos")
    cgf_depois = {l["ID"]: l.get("NV2", "") for l in t.ler_aba("CGF") if l.get("ID")}
    mudaram = sorted(id_cgf for id_cgf, nv2 in cgf_antes.items()
                     if id_cgf in cgf_depois and cgf_depois[id_cgf] != nv2)
    check("round-trip extrair->unificar preserva o NV2 de comando dos %d CGF da base (%s)"
          % (len(cgf_antes), mudaram[:3] or "nenhum mudou"), not mudaram)
    check("round-trip: nenhum CGF da base sumiu",
          all(id_cgf in cgf_depois for id_cgf in cgf_antes))

print()
if falhas:
    print(f"{len(falhas)} checagem(ns) FALHOU/FALHARAM: {falhas}")
    sys.exit(1)
print("Todas as checagens do teste UNO real (protocolos) passaram.")
