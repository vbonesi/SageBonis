# -*- coding: utf-8 -*-

import os
import re

# ===============================================================
# ================ MACRO SAGE - VERSÃO 0.2 ======================
# ===============================================================
# Este script é utilizado como macro no LibreOffice Calc para importar e exportar arquivos .dat
# do Sistema Aberto de Gerenciamento de Energia (SAGE).
#
# Desenvolvido para rodar com a planilha SageBonis.ods
# Duvidas/Bugs/Sugestões - (11) 95456-4510 - Victor Bonesi - https://github.com/vbonesi/SageBonis

# ===============================================================
# ==================== CONFIGURAÇÃO GERAL =======================
# ===============================================================
# Nesta seção, todas as configurações e "valores mágicos" do script
# são definidos como constantes para facilitar a manutenção.

# --- Nomes de Abas ---
NOME_ABA_GERAL = "geral"

# --- Posições das Células na Aba "geral" ---
CELULA_CAMINHO_IMPORTACAO = (0, 3)  # A4
CELULA_STATUS_IMPORTACAO = (1, 3)   # B4
CELULA_CAMINHO_EXPORTACAO = (0, 6)  # A7
CELULA_STATUS_EXPORTACAO = (1, 6)   # B7

# --- Intervalo para Funções Parciais (Col Ini, Lin Ini, Col Fim, Lin Fim) ---
RANGE_ENTIDADES_PARCIAL = (2, 13, 2, 143) # C15:C145

# --- Códigos de Controle (Coluna "Gera") ---
CODIGO_BLOCO_ATIVO = 'x'
CODIGO_BLOCO_COMENTADO = 'c'
CODIGO_COMENTARIO_SIMPLES = 'n'
CODIGO_IGNORAR_LINHA = 'q'
CODIGO_INCLUDE = 'i'
CODIGO_INCLUDE_COMENTADO = 'u'

# --- Cabeçalhos Padrão ---
CABEÇALHO_COLUNA_ORIGEM = "Origem"
CABEÇALHO_COLUNA_CONTROLE = "Gera"
CABEÇALHO_COLUNA_DADOS = "Aux" # Coluna para caminhos de include e comentários

# --- Constantes Técnicas ---
FLAGS_LIMPAR_TUDO = 1048575

# --- Expressões Regulares ---
# Regex para encontrar #include e capturar o caminho (sem aspas).
REGEX_INCLUDE = re.compile(r'^\s*#\s*include\s+(.*)', re.IGNORECASE)
# Regex para encontrar ;#include e capturar o caminho (sem aspas).
REGEX_INCLUDE_COMENTADO = re.compile(r'^\s*;\s*#\s*include\s+(.*)', re.IGNORECASE)

# ===============================================================
# ================= FUNÇÕES DE IMPORTAÇÃO =======================
# ===============================================================

def importar_dats(*args):
    """(IMPORTAÇÃO TOTAL) Varre a pasta de origem e todas as subpastas, importando todos os arquivos .dat."""
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    try:
        geral_sheet = doc.getSheets().getByName(NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}") # type: ignore
        return

    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Processando importação total...")
    _executar_importacao(doc, folder_path, lista_entidades=None, modo_importacao='REPLACE')
    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Importação total concluída com sucesso!")


def importar_parcial(*args):
    """(IMPORTAÇÃO PARCIAL) Importa apenas as entidades selecionadas, varrendo todas as subpastas."""
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    
    try:
        geral_sheet = doc.getSheets().getByName(NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}") # type: ignore
        return

    entidades_a_importar = []
    modo = 'REPLACE' 
    if active_sheet_name.lower() == NOME_ABA_GERAL.lower():
        range_entidades = geral_sheet.getCellRangeByPosition(*RANGE_ENTIDADES_PARCIAL)
        dados_entidades = range_entidades.getDataArray()
        entidades_a_importar = [row[0].lower() for row in dados_entidades if row and row[0]]
        if not entidades_a_importar:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("AVISO: Nenhuma entidade listada para importação parcial.")
            return
    else:
        entidades_a_importar.append(active_sheet_name.lower())
        modo = 'UPDATE'

    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"Processando importação de: {', '.join(entidades_a_importar)}...")
    _executar_importacao(doc, folder_path, lista_entidades=entidades_a_importar, modo_importacao=modo)
    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Importação parcial concluída com sucesso!")


def _executar_importacao(doc, base_folder_path, lista_entidades, modo_importacao):
    """
    Função interna que executa a importação recursiva.
    
    Argumentos:
        base_folder_path (str): A pasta raiz para iniciar a varredura.
        lista_entidades (list|None): Lista de entidades para filtrar. Se None, importa todas.
        modo_importacao (str): 'REPLACE' ou 'UPDATE'.
    """
    all_data = {}
    
    # Percorre recursivamente a pasta de importação
    for root, _, files in os.walk(base_folder_path):
        for file_name in files:
            if not file_name.lower().endswith('.dat'):
                continue

            entidade_nome = os.path.splitext(file_name)[0].lower()

            # Se for uma importação parcial, filtra as entidades
            if lista_entidades is not None and entidade_nome not in lista_entidades:
                continue

            full_path = os.path.join(root, file_name)
            # Calcula o caminho relativo para a coluna 'Origem'
            relative_path = os.path.relpath(full_path, base_folder_path)
            
            # Obtém todas as entidades válidas para o parser (pode ser todas ou a lista filtrada)
            entidades_validas_set = {os.path.splitext(f)[0].upper() for f in os.listdir(root) if f.lower().endswith('.dat')}
            
            parse_dat_file(full_path, relative_path, all_data, entidades_validas_set)

    # Lógica de escrita na planilha (limpar/substituir abas)
    abas_a_escrever = lista_entidades if lista_entidades is not None else all_data.keys()
    for entidade_nome in abas_a_escrever:
        pontos = all_data.get(entidade_nome)
        if pontos:
            write_to_sheet(doc, entidade_nome, pontos, modo_importacao)

# [O restante do código, como as funções de exportação, permanecem iguais]
# ...

def write_to_sheet(doc, sheet_name, pontos_importados, modo):
    """
    Escreve os dados em uma aba, garantindo que a coluna 'Dados' exista
    para receber caminhos de include e comentários.
    """
    if modo == 'UPDATE' and doc.getSheets().hasByName(sheet_name):
        sheet = doc.getSheets().getByName(sheet_name)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        range_to_clear = sheet.getCellRangeByPosition(0, 0, cursor.getRangeAddress().EndColumn, cursor.getRangeAddress().EndRow)
        range_to_clear.clearContents(FLAGS_LIMPAR_TUDO)
    else:
        if doc.getSheets().hasByName(sheet_name):
            doc.getSheets().removeByName(sheet_name)
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(sheet_name, new_sheet)
        sheet = doc.getSheets().getByName(sheet_name)

    todos_atributos = set()
    for ponto in pontos_importados:
        if 'attributes' in ponto:
            todos_atributos.update(ponto['attributes'].keys())

    # Monta os cabeçalhos. A coluna 'Dados' agora é fixa.
    cabecalhos = [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS] + sorted(list(todos_atributos))
    
    for i, header_text in enumerate(cabecalhos):
        sheet.getCellByPosition(i, 0).setString(header_text)

    # Preenche as linhas da planilha.
    for row_idx, ponto in enumerate(pontos_importados, start=1):
        sheet.getCellByPosition(0, row_idx).setString(ponto.get('origem', '')) # Coluna Origem
        sheet.getCellByPosition(1, row_idx).setString(ponto['type'])          # Coluna Gera

        # Para includes e comentários, o 'data' vai para a coluna 'Dados' (índice 2).
        if ponto['type'] in [CODIGO_COMENTARIO_SIMPLES, CODIGO_INCLUDE, CODIGO_INCLUDE_COMENTADO]:
            sheet.getCellByPosition(2, row_idx).setString(ponto.get('data', ''))
        
        # Para blocos, os atributos vão para as colunas correspondentes.
        elif 'attributes' in ponto:
            for attr_key, attr_value in ponto['attributes'].items():
                try:
                    col_idx = cabecalhos.index(attr_key)
                    sheet.getCellByPosition(col_idx, row_idx).setString(attr_value)
                except ValueError:
                    pass

# ===============================================================
# =================== LÓGICA INTERNA (NÃO ALTERADA) ==================
# ===============================================================
# As funções abaixo não precisaram de alteração.
# Inclua-as no seu script final.

def parse_dat_file(file_path, relative_path, all_data, entidades_validas):
    """
    Analisa um único arquivo .dat, extraindo informações, incluindo
    includes ativos e comentados, e armazenando a origem.
    """
    try:
        with open(file_path, 'r', encoding='latin-1', errors='ignore') as f:
            lines = f.readlines()
    except IOError as e:
        print(f"Erro ao ler o arquivo {file_path}: {e}")
        return

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        original_line = lines[i].strip('\r\n')
        i += 1

        if not line:
            continue

        entidade_chave = os.path.splitext(os.path.basename(file_path))[0].lower()

        # --- NOVA LÓGICA DE PARSING ---
        # Prioridade 1: Verificar se é um include comentado (;#include)
        include_comentado_match = REGEX_INCLUDE_COMENTADO.match(original_line)
        if include_comentado_match:
            caminho_do_include = include_comentado_match.group(1).strip()
            ponto = {
                'type': CODIGO_INCLUDE_COMENTADO,
                'data': caminho_do_include, # Salva SÓ o caminho
                'origem': relative_path
            }
            all_data.setdefault(entidade_chave, []).append(ponto)
            continue

        # Prioridade 2: Verificar se é um include ativo (#include)
        include_match = REGEX_INCLUDE.match(original_line)
        if include_match:
            caminho_do_include = include_match.group(1).strip()
            ponto = {
                'type': CODIGO_INCLUDE,
                'data': caminho_do_include, # Salva SÓ o caminho
                'origem': relative_path
            }
            all_data.setdefault(entidade_chave, []).append(ponto)
            continue

        # Prioridade 3: Tratar blocos comentados (;PDS) e comentários simples (;)
        if line.startswith(';'):
            match = re.match(r'^; *([A-Z_]+)', line, re.IGNORECASE)
            if match and match.group(1).upper() in entidades_validas:
                # É um bloco de entidade comentado
                # (A lógica aqui permanece a mesma de antes)
                ponto = {
                    'type': CODIGO_BLOCO_COMENTADO, 'identifier': match.group(1).upper(),
                    'attributes': {}, 'origem': relative_path
                }
                while i < len(lines) and lines[i].strip().startswith(';'):
                    attr_line = lines[i].strip()[1:].strip()
                    if '=' in attr_line: key, value = attr_line.split('=', 1); ponto['attributes'][key.strip()] = value.strip()
                    i += 1
                entidade_ponto = ponto['identifier'].lower()
                all_data.setdefault(entidade_ponto, []).append(ponto)
            else:
                # É um comentário simples
                all_data.setdefault(entidade_chave, []).append({
                    'type': CODIGO_COMENTARIO_SIMPLES, 'data': original_line, 'origem': relative_path
                })
            continue

        # Prioridade 4: Tratar blocos ativos (PDS)
        if line.upper() in entidades_validas:
            # (A lógica aqui permanece a mesma de antes)
            ponto = {
                'type': CODIGO_BLOCO_ATIVO, 'identifier': line.upper(),
                'attributes': {}, 'origem': relative_path
            }
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line or next_line.upper() in entidades_validas or next_line.startswith(';') or REGEX_INCLUDE.match(lines[i]) or REGEX_INCLUDE_COMENTADO.match(lines[i]):
                    break
                if '=' in next_line: key, value = next_line.split('=', 1); ponto['attributes'][key.strip()] = value.strip()
                i += 1
            if ponto['attributes']:
                entidade_ponto = ponto['identifier'].lower()
                all_data.setdefault(entidade_ponto, []).append(ponto)
            continue

# --- Funções de Exportação (sem alterações) ---

def exportar_dats(*args):
    """(EXPORTAÇÃO TOTAL) Exporta todas as abas (exceto 'geral') para arquivos .dat."""
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    try:
        geral_sheet = doc.getSheets().getByName(NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()

        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}") # type: ignore
        return

    geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("Processando exportação total...")
    abas_a_exportar = [s for s in doc.getSheets() if s.getName().lower() != NOME_ABA_GERAL.lower()]
    erros = [_exportar_folha(sheet, export_folder) for sheet in abas_a_exportar]
    erros = [e for e in erros if e] # Filtra apenas as mensagens de erro
    
    if erros:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: {'; '.join(erros)}")
    else:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("Exportação total concluída com sucesso!")

def exportar_parcial(*args):
    """(EXPORTAÇÃO PARCIAL) Exporta apenas as entidades selecionadas."""
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    
    try:
        geral_sheet = doc.getSheets().getByName(NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()

        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}") # type: ignore
        return

    abas_a_exportar = []
    if active_sheet_name.lower() == NOME_ABA_GERAL.lower():
        range_entidades = geral_sheet.getCellRangeByPosition(*RANGE_ENTIDADES_PARCIAL)
        dados_entidades = range_entidades.getDataArray()
        nomes_entidades = [row[0].lower() for row in dados_entidades if row and row[0]]
        if not nomes_entidades:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("AVISO: Nenhuma entidade listada para exportação parcial.")
            return
        for nome in nomes_entidades:
            if doc.getSheets().hasByName(nome):
                abas_a_exportar.append(doc.getSheets().getByName(nome))
    else:
        abas_a_exportar.append(active_sheet)

    geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"Processando exportação de: {', '.join(s.getName() for s in abas_a_exportar)}...")
    erros = [_exportar_folha(sheet, export_folder) for sheet in abas_a_exportar]
    erros = [e for e in erros if e]

    if erros:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: {'; '.join(erros)}")
    else:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("Exportação parcial concluída com sucesso!")

def _exportar_folha(sheet, export_folder):
    """
    Exporta uma única aba, remontando os includes (ativos e comentados)
    a partir da coluna 'Dados'.
    """
    sheet_name = sheet.getName()
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    data_range = cursor.getRangeAddress()
    data_array = sheet.getCellRangeByPosition(0, 0, data_range.EndColumn, data_range.EndRow).getDataArray()

    if not data_array or len(data_array) < 2: return

    headers = data_array[0]
    try:
        origem_col_idx = headers.index(CABEÇALHO_COLUNA_ORIGEM)
        gera_col_idx = headers.index(CABEÇALHO_COLUNA_CONTROLE)
        dados_col_idx = headers.index(CABEÇALHO_COLUNA_DADOS)
    except ValueError:
        return f"Aba '{sheet_name}' não possui as colunas 'Origem', 'Gera' ou 'Dados'."

    dados_agrupados_por_arquivo = {}

    for row_data in data_array[1:]:
        if len(row_data) <= max(origem_col_idx, gera_col_idx, dados_col_idx): continue
        
        origem_path = str(row_data[origem_col_idx])
        control_code = str(row_data[gera_col_idx]).lower()

        if not origem_path or not control_code or control_code == CODIGO_IGNORAR_LINHA:
            continue

        dados_agrupados_por_arquivo.setdefault(origem_path, [])
        bloco_final = None
        
        # --- NOVA LÓGICA DE EXPORTAÇÃO ---
        dado_principal = str(row_data[dados_col_idx])

        if control_code == CODIGO_INCLUDE and dado_principal:
            # CORREÇÃO: Remove as aspas da exportação
            bloco_final = f'#include {dado_principal}'

        elif control_code == CODIGO_INCLUDE_COMENTADO and dado_principal:
            # CORREÇÃO: Remove as aspas da exportação
            bloco_final = f';#include {dado_principal}'

        elif control_code == CODIGO_COMENTARIO_SIMPLES:
            bloco_final = f';{dado_principal}'

        elif control_code in [CODIGO_BLOCO_ATIVO, CODIGO_BLOCO_COMENTADO]:
            point_lines = [sheet_name.upper()]
            for col_idx, header in enumerate(headers):
                if header in [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS]:
                    continue
                value = str(row_data[col_idx]) if len(row_data) > col_idx else ""
                if value: point_lines.append(f"\t{header} = {value}")
            
            if len(point_lines) > 1:
                bloco_texto = "\n".join(point_lines)
                if control_code == CODIGO_BLOCO_COMENTADO:
                    bloco_final = "\n".join([f";{line}" for line in bloco_texto.split('\n')])
                else:
                    bloco_final = bloco_texto

        if bloco_final is not None:
            dados_agrupados_por_arquivo[origem_path].append(bloco_final)

    # A parte que escreve os arquivos permanece a mesma.
    for relative_path, file_content_list in dados_agrupados_por_arquivo.items():
        try:
            full_output_path = os.path.join(export_folder, relative_path)
            os.makedirs(os.path.dirname(full_output_path), exist_ok=True)
            with open(full_output_path, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(file_content_list)); f.write("\n")
        except IOError as e:
            return f"Falha ao escrever {relative_path}: {e}"
            
    return None

# ===============================================================
# ================= EXPOSIÇÃO PARA LIBREOFFICE ==================
# ===============================================================
g_exportedScripts = importar_dats, exportar_dats, importar_parcial, exportar_parcial