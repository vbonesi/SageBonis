# -*- coding: utf-8 -*-
"""
testar_verificacao.py — Roda o verificador de base (verificar_base) FORA do
LibreOffice, lendo os dados direto de um arquivo .ods.

Por que existe: a lógica das checagens em ImportadorSAGE.py é pura (opera sobre
mapas em memória, não sobre objetos UNO). Este script lê o .ods, monta o mesmo
mapa de entidades e as mesmas regras da aba "VerificacaoRefs", e chama EXATAMENTE
as funções do verificador. Assim dá para testar/depurar a lista de relacionamentos
sem abrir o LibreOffice — basta salvar o .ods com a base importada.

Uso:
    python completa/testar_verificacao.py [caminho/para/arquivo.ods]
    (padrão: completa/SageBonis.ods)

Importante: salve o .ods no LibreOffice ANTES de rodar (o script lê o arquivo em
disco, não a sessão aberta). Para testar relacionamentos, ative regras na aba
"VerificacaoRefs" (coluna Ativa = S) e salve.
"""

import importlib.util
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

_NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_NS = {
    "table": _NS_TABLE,
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

# Acima deste limite, repetições (de colunas ou linhas) são tratadas como o
# "preenchimento" que o ODF usa para chegar ao fim da planilha — e ignoradas.
_CAP_REPETICAO = 1024


def _carregar_macro_modulo():
    """Importa o ImportadorSAGE.py ao lado deste script como módulo."""
    aqui = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(aqui, "ImportadorSAGE.py")
    spec = importlib.util.spec_from_file_location("importador_sage_completa", caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _texto_celula(cell):
    """Concatena os parágrafos text:p de uma célula (linhas separadas por \\n)."""
    ps = cell.findall("text:p", _NS)
    if not ps:
        return ""
    return "\n".join("".join(p.itertext()) for p in ps)


def parse_ods(path):
    """Lê content.xml de um .ods -> {nome_aba: [linha0(=cabeçalho), linha1, ...]}.
    Cada linha é uma lista de strings. Expande repetições e apara vazios finais."""
    with zipfile.ZipFile(path) as z:
        conteudo = z.read("content.xml")
    root = ET.fromstring(conteudo)

    resultado = {}
    for table in root.iter("{%s}table" % _NS_TABLE):
        nome = table.get("{%s}name" % _NS_TABLE)
        if nome is None:
            continue
        linhas = []
        for row in table.findall("table:table-row", _NS):
            rrep = int(row.get("{%s}number-rows-repeated" % _NS_TABLE, "1"))
            celulas = []
            for cell in list(row):
                tag = cell.tag.split("}")[-1]
                if tag not in ("table-cell", "covered-table-cell"):
                    continue
                crep = int(cell.get("{%s}number-columns-repeated" % _NS_TABLE, "1"))
                val = _texto_celula(cell)
                if val == "" and crep > _CAP_REPETICAO:
                    crep = 0  # preenchimento de colunas até o fim — ignora
                celulas.extend([val] * crep)
            while celulas and celulas[-1] == "":
                celulas.pop()
            if not celulas and rrep > _CAP_REPETICAO:
                rrep = 0  # preenchimento de linhas até o fim — ignora
            for _ in range(rrep):
                linhas.append(list(celulas))
        while linhas and not any(c.strip() for c in linhas[-1]):
            linhas.pop()
        resultado[nome] = linhas
    return resultado


def montar_entidades(mod, abas):
    """{nome: (headers, linhas)} para as abas que são entidades (têm coluna 'Gera')."""
    ignoradas = set(n.lower() for n in mod.FOLHAS_IGNORADAS)
    ignoradas.update({mod.NOME_ABA_ANALISE.lower(), mod.NOME_ABA_VERIFICACAO_REFS.lower()})
    entidades = {}
    for nome, linhas in abas.items():
        if nome.lower() in ignoradas or len(linhas) < 2:
            continue
        headers = [str(h) for h in linhas[0]]
        if mod._idx_coluna(headers, mod.CABEÇALHO_COLUNA_CONTROLE) < 0:
            continue
        entidades[nome] = (headers, linhas[1:])
    return entidades


def montar_regras(mod, abas):
    """Lê regras ATIVAS da aba VerificacaoRefs do .ods (mesma semântica da macro)."""
    aba = None
    for nome, linhas in abas.items():
        if nome.lower() == mod.NOME_ABA_VERIFICACAO_REFS.lower():
            aba = linhas
            break
    if not aba or len(aba) < 2:
        return [], 0
    regras, total = [], 0
    for row in aba[1:]:
        if len(row) < 5:
            continue
        ent_o, attr_o = str(row[0]).strip(), str(row[1]).strip()
        ent_d = mod._parse_entidades_destino(row[2])
        attr_d = str(row[3]).strip()
        if not (ent_o and attr_o and ent_d and attr_d):
            continue
        total += 1
        if str(row[4]).strip().lower() in mod._VALORES_ATIVO:
            regras.append((ent_o, attr_o, ent_d, attr_d))
    return regras, total


def montar_dominios(mod, abas):
    """Lê a aba EntidadeAtributoValor do .ods (mesma semântica de mod._carregar_dominios)."""
    aba = None
    for nome, linhas in abas.items():
        if nome.lower() == mod.NOME_ABA_VALIDACAO.lower():
            aba = linhas
            break
    dominios = {}
    if aba and len(aba) >= 2:
        for row in aba[1:]:
            if len(row) < 3:
                continue
            entidade = str(row[0]).strip().lower()
            atributo = str(row[1]).strip().upper()
            if not entidade or not atributo:
                continue
            valores = {str(v).strip() for v in row[2:] if str(v).strip()}
            if valores:
                dominios[(entidade, atributo)] = valores
    mod._mesclar_dominios_suplementares(dominios)
    return dominios


def imprimir_relatorio(mod, analise):
    cab = mod.CABECALHOS_ANALISE
    achados = sorted(analise.achados, key=lambda a: mod._ORDEM_SEV.get(a["sev"], 9))
    linhas = [[a["sev"], a["entidade"], str(a["linha"]), a["atributo"], a["valor"], a["descr"]]
              for a in achados]
    if not linhas:
        linhas = [["OK", "-", "-", "-", "-", "Nenhum problema encontrado."]]
    larguras = [len(h) for h in cab]
    for row in linhas:
        for i, c in enumerate(row):
            larguras[i] = max(larguras[i], len(c))
    fmt = "  ".join("{:<%d}" % w for w in larguras)
    print(fmt.format(*cab))
    print("  ".join("-" * w for w in larguras))
    for row in linhas:
        print(fmt.format(*row))
    print("\nResumo: %d ERRO(s), %d AVISO(s)." % (analise.erros, analise.avisos))


def main():
    caminho = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "SageBonis.ods")
    if not os.path.isfile(caminho):
        sys.exit("ERRO: arquivo nao encontrado: %s" % caminho)

    mod = _carregar_macro_modulo()
    abas = parse_ods(caminho)
    entidades = montar_entidades(mod, abas)
    regras, total_regras = montar_regras(mod, abas)
    dominios = montar_dominios(mod, abas)

    print("Arquivo: %s" % caminho)
    print("Entidades encontradas (%d): %s" % (len(entidades), ", ".join(sorted(entidades))))
    print("Regras de relacionamento: %d ativa(s) de %d." % (len(regras), total_regras))
    if regras:
        for ent_o, attr_o, ent_d, attr_d in regras:
            print("   ativa -> %s.%s deve existir em %s.%s" % (ent_o, attr_o, "/".join(ent_d), attr_d))
    print("Dominios carregados: %d par(es) entidade.atributo." % len(dominios))
    print()

    analise = mod._rodar_checagens(entidades, regras, dominios)
    imprimir_relatorio(mod, analise)


if __name__ == "__main__":
    main()
