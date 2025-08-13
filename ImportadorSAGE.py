import os
import re

# ===============================================================
# ================ MACRO SAGE - VERSÃO 0.1 ======================
# ===============================================================
# Este script é utilizado como macro no LibreOffice Calc para importar e exportar arquivos .dat
# do Sistema Aberto de Gerenciamento de Energia (SAGE), um sistema SCADA. Ele lê arquivos .dat
# contendo configurações de base de dados, organiza os dados em abas da planilha e permite exportar
# novamente para o formato .dat.
# Desenvolvido para rodar com a planilha SageBonis.ods
# Duvidas/Bugs/Sugestões - (11) 95456-4510 - Victor Bonesi

# ===============================================================
# ==================== FUNÇÃO DE IMPORTAÇÃO =====================
# ===============================================================

def importar_dats():
    """
    Função principal de importação dos arquivos .dat para o LibreOffice Calc.
    - Lê o caminho da pasta informado na aba 'geral'.
    - Procura arquivos .dat na pasta.
    - Para cada arquivo, faz o parsing dos dados (inclusive includes recursivos).
    - Cria uma aba para cada entidade encontrada, preenchendo os dados importados.
    - Atualiza o status na aba 'geral'.
    """
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    try:
        geral_sheet = doc.getSheets().getByName("geral")
        path_cell = geral_sheet.getCellByPosition(0, 3)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(0, 3).setString(f"ERRO: Caminho invalido!")
            return
        geral_sheet.getCellByPosition(1, 3).setString("Processando...")
    except Exception:
        try:
            new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
            doc.getSheets().insertByName("geral", new_sheet)
            geral_sheet = doc.getSheets().getByName("geral")
            geral_sheet.getCellByPosition(1, 3).setString("ERRO: Caminho em Branco!")
        except Exception:
            pass
        return

    # Lista todos os arquivos .dat na pasta informada
    dat_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.dat')]
    if not dat_files:
        geral_sheet.getCellByPosition(1, 3).setString("ERRO: Nenhum arquivo .dat encontrado.")
        return

    entidades_validas = {f.split('.')[0].upper() for f in dat_files}
    all_data = {}

    # Para cada arquivo .dat, faz o parsing recursivo (considerando includes)
    for file_name in dat_files:
        file_path = os.path.join(folder_path, file_name)
        file_key = os.path.basename(file_path).lower().split('.')[0]
        all_data.setdefault(file_key, [])
        parse_dat_recursive(file_path, all_data, entidades_validas, set())
    
    # Limpa abas antigas (exceto 'geral')
    sheet_names = doc.getSheets().getElementNames()
    for name in sheet_names:
        if name.lower() != 'geral':
            try:
                doc.getSheets().removeByName(name)
            except Exception:
                pass # Ignora erros ao remover abas que talvez não existam mais

    # Cria as novas abas
    for entidade, dados in sorted(all_data.items()):
        if dados:
            write_to_sheet(doc, entidade, dados)

    geral_sheet.getCellByPosition(1, 3).setString("Importação concluída com sucesso!")


def parse_dat_recursive(file_path, all_data, entidades_validas, processed_includes):
    """
    Função recursiva para fazer o parsing de arquivos .dat.
    - Lê o arquivo linha a linha.
    - Identifica includes e faz parsing recursivo dos arquivos incluídos.
    - Identifica blocos comentados (iniciados por ';') e blocos ativos (entidades).
    - Coleta atributos de cada bloco e armazena em dicionário.
    - Adiciona os dados no dicionário all_data, agrupando por entidade.
    """
    if file_path in processed_includes or not os.path.exists(file_path):
        return
    processed_includes.add(file_path)

    file_key = os.path.basename(file_path).lower().split('.')[0]
    base_dir = os.path.dirname(file_path)

    with open(file_path, 'r', encoding='latin-1', errors='ignore') as f:
        lines = f.readlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip(' \t\r\n')  # Remove apenas espaços laterais
        i += 1

        if not line:
            continue

        # Trata blocos comentados (iniciados por ';')
        if line.startswith(';'):
            match = re.match(r'^; *([A-Z_]+)', line, re.IGNORECASE)
            if match and match.group(1).upper() in entidades_validas:
                # Cabeçalho de bloco comentado: coleta atributos
                point_dict = {'type': 'c', 'identifier': match.group(1).upper(), 'attributes': {}, 'file': os.path.basename(file_path)}
                while i < len(lines) and lines[i].strip().startswith(';'):
                    attr_line = lines[i].strip()[1:].strip()
                    if '=' in attr_line:
                        key, value = attr_line.split('=', 1)
                        point_dict['attributes'][key.strip()] = value.strip()
                    i += 1
                entidade_ponto = point_dict['identifier'].lower()
                all_data.setdefault(entidade_ponto, []).append(point_dict)
            else:
                # Comentário simples
                all_data.setdefault(file_key, []).append({'type': 'n', 'data': line, 'file': os.path.basename(file_path)})
            continue

        # Trata blocos ativos (entidades)
        if line.upper() in entidades_validas:
            point_dict = {'type': 'x', 'identifier': line.upper(), 'attributes': {}, 'file': os.path.basename(file_path)}
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line or next_line.upper() in entidades_validas or next_line.startswith(';'):
                    break
                if '=' in next_line:
                    key, value = next_line.split('=', 1)
                    point_dict['attributes'][key.strip()] = value.strip()
                i += 1
            if point_dict['attributes']:
                entidade_ponto = point_dict['identifier'].lower()
                all_data.setdefault(entidade_ponto, []).append(point_dict)
            continue


def write_to_sheet(doc, sheet_name, lines):
    """
    Cria uma nova aba na planilha para a entidade e preenche os dados importados.
    - Define os cabeçalhos das colunas (Gera, atributos encontrados).
    - Preenche cada linha com os dados de cada ponto (bloco ativo, comentado ou comentário).
    """
    new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
    doc.getSheets().insertByName(sheet_name, new_sheet)
    sheet = doc.getSheets().getByName(sheet_name)

    headers_base = ["Gera"]
    all_attr_keys = set()
    for line in lines:
        if line['type'] in ['x', 'c'] and 'attributes' in line:
            all_attr_keys.update(line['attributes'].keys())

    sorted_headers = sorted(list(all_attr_keys))
    final_headers = headers_base + sorted_headers
    for i, h in enumerate(final_headers):
        sheet.getCellByPosition(i, 0).setString(h)

    row_idx = 1
    for line_data in lines:
        sheet.getCellByPosition(0, row_idx).setString(line_data['type'])

        if line_data['type'] == 'n':
            if len(final_headers) > 2:
                sheet.getCellByPosition(1, row_idx).setString(line_data['data'])
        elif line_data['type'] in ['x', 'c']:
            if 'attributes' in line_data:
                for attr_key, attr_value in line_data['attributes'].items():
                    try:
                        col_idx = final_headers.index(attr_key)
                        sheet.getCellByPosition(col_idx, row_idx).setString(attr_value)
                    except ValueError:
                        pass
        row_idx += 1

# ===============================================================
# ==================== FUNÇÃO DE EXPORTAÇÃO =====================
# ===============================================================

def exportar_dats():
    """
    Função principal de exportação dos dados das abas do Calc para arquivos .dat.
    - Lê o caminho de exportação informado na aba 'geral'.
    - Para cada aba (exceto 'geral'), gera um arquivo .dat com os dados formatados.
    - Blocos ativos e comentados são exportados conforme o tipo ('x' ou 'c').
    - Comentários são exportados como linhas simples.
    - Atualiza o status na aba 'geral'.
    """
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    try:
        geral_sheet = doc.getSheets().getByName("geral")
        export_path_cell = geral_sheet.getCellByPosition(0, 6)
        export_folder = export_path_cell.getString()
        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(1, 6).setString(f"ERRO: Caminho invalido!")
            return
        geral_sheet.getCellByPosition(1, 6).setString("Processando...")
    except Exception:
        try:
            geral_sheet = doc.getSheets().getByName("geral")
        except Exception:
            return
        geral_sheet.getCellByPosition(1, 6).setString("ERRO: Caminho em Branco!")
        return

    for sheet in doc.getSheets():
        sheet_name = sheet.getName()
        if sheet_name.lower() == 'geral':
            continue

        output_filename = os.path.join(export_folder, f"{sheet_name.lower()}.dat")
        file_content = []

        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        data_range = sheet.getCellRangeByPosition(0, 0, cursor.getRangeAddress().EndColumn, cursor.getRangeAddress().EndRow)
        data = data_range.getDataArray()

        if not data or not data[0]:
            continue

        headers = data[0]
        try:
            gera_col_idx = headers.index("Gera")
        except ValueError:
            continue

        for row_idx, row_data in enumerate(data[1:], 1):
            control_code = row_data[gera_col_idx].lower() if len(row_data) > gera_col_idx else ""

            if not control_code or control_code == 'q':
                continue

            if control_code == 'n':
                # Exporta comentários simples
                comment = row_data[gera_col_idx + 1] if len(row_data) > gera_col_idx + 1 else ""
                if comment:
                    file_content.append(comment)
                continue

            if control_code in ['x', 'c']:
                # Exporta blocos ativos ou comentados
                point_lines = [sheet_name.upper()]
                for col_idx, header in enumerate(headers):
                    if header == "Gera": continue
                    value = row_data[col_idx] if len(row_data) > col_idx else ""
                    if value:
                        point_lines.append(f"\t{header} = {value}")

                if len(point_lines) > 1:
                    final_block = "\n".join(point_lines)
                    if control_code == 'c':
                        final_block = "\n".join([f";{line}" for line in final_block.split('\n')])
                    file_content.append(final_block)

        try:
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(file_content))
                f.write("\n")
        except Exception as e:
            geral_sheet.getCellByPosition(1, 6).setString(f"ERRO ao escrever {output_filename}: {e}")
            return

    geral_sheet.getCellByPosition(1, 6).setString("Exportação concluída com sucesso!")


# Scripts a serem expostos para o LibreOffice
# Permite que as funções sejam chamadas como macros no Calc
g_exportedScripts = importar_dats, exportar_dats,