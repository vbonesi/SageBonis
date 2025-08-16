# -*- coding: utf-8 -*-

import os
import re

# ===============================================================
# ================ MACRO SAGE - VERSÃO 0.6 ======================
# ===============================================================
# Este script é utilizado como macro no LibreOffice Calc para importar e exportar arquivos .dat
# do Sistema Aberto de Gerenciamento de Energia (SAGE).
#
# A partir desta versão, o script lê dinamicamente as configurações de
# ordenação, cores e validação das abas "MaisUsadas" e "EntidadesValoresAtributos".
#
# Desenvolvido para rodar com a planilha SageBonis.ods
# Duvidas/Bugs/Sugestões - (11) 95456-4510 - Victor Bonesi - https://github.com/vbonesi/SageBonis

# ===============================================================
# ==================== CONFIGURAÇÃO GERAL =======================
# ===============================================================

# --- Nomes de Abas ---
NOME_ABA_GERAL = "geral"
NOME_ABA_MAIS_USADAS = "MaisUsadas"
NOME_ABA_VALIDACAO = "EntidadeAtributoValor"

# --- Lista de Abas a Ignorar ---
FOLHAS_IGNORADAS = [NOME_ABA_GERAL, NOME_ABA_MAIS_USADAS, NOME_ABA_VALIDACAO, "opmsk"]

# --- Posições das Células na Aba "geral" ---
CELULA_CAMINHO_IMPORTACAO = (0, 3)  # A4
CELULA_STATUS_IMPORTACAO = (1, 3)   # B4
CELULA_CAMINHO_EXPORTACAO = (0, 6)  # A7
CELULA_STATUS_EXPORTACAO = (1, 6)   # B7
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
CABEÇALHO_COLUNA_DADOS = "Aux"

# --- NOVO: Cores para Linhas Alternadas (Efeito Zebra) ---
# Cores em formato numérico (Decimal de Hex BGR: Blue-Green-Red)
COR_LINHA_PAR = 16777215   # Branco (0xFFFFFF)
COR_LINHA_IMPAR = 15790320  # Cinza muito claro (0xF0F0F0)

# --- Constantes Técnicas ---
FLAGS_LIMPAR_TUDO = 1048575
LIMITE_CARACTERES_VALIDACAO = 250 # Manteremos para a próxima etapa

# --- Expressões Regulares ---
REGEX_INCLUDE = re.compile(r'^\s*#\s*include\s+(.*)', re.IGNORECASE)
REGEX_INCLUDE_COMENTADO = re.compile(r'^\s*;\s*#\s*include\s+(.*)', re.IGNORECASE)

# ===============================================================
# =================== CLASSE DE CONFIGURAÇÃO ====================
# ===============================================================

class SageConfig:
    """Carrega e armazena todas as configurações das abas auxiliares."""
    def __init__(self, doc):
        self.doc = doc
        self.ordem_entidades = []
        self.cores_entidades = {}
        self.ordem_atributos = {}
        self.regras_validacao = {} # Mantido por segurança, mas não será preenchido
        
        self._carregar_configuracoes()

    def _carregar_configuracoes(self):
        """Método principal para chamar os carregadores."""
        self._carregar_mais_usadas()
        # A LINHA ABAIXO FOI REMOVIDA:
        # self._carregar_validacao()

    def _carregar_mais_usadas(self):
        """Lê a aba 'MaisUsadas' para obter ordem, cores e atributos prioritários."""
        try:
            sheet = self.doc.getSheets().getByName(NOME_ABA_MAIS_USADAS)
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            data_range = cursor.getRangeAddress()
            data = sheet.getCellRangeByPosition(0, 0, data_range.EndColumn, data_range.EndRow).getDataArray()

            if not data or len(data) < 2: return

            for row_idx, row_data in enumerate(data[1:], 1):
                if not row_data or not row_data[0]: continue
                entidade_nome = str(row_data[0]).lower().strip()
                if not entidade_nome: continue

                self.ordem_entidades.append(entidade_nome)
                cell = sheet.getCellByPosition(0, row_idx)
                self.cores_entidades[entidade_nome] = cell.CellBackColor
                atributos = [str(attr).upper() for attr in row_data[1:] if attr]
                if atributos:
                    self.ordem_atributos[entidade_nome] = atributos
        except Exception as e:
            print(f"AVISO: Não foi possível carregar as configurações da aba '{NOME_ABA_MAIS_USADAS}'. {e}")

    # A FUNÇÃO _carregar_validacao FOI COMPLETAMENTE REMOVIDA DESTA CLASSE

# ===============================================================
# ================= FUNÇÕES DE IMPORTAÇÃO =======================
# ===============================================================

def importar_dats(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # (O código interno desta função não muda)
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
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # (O código interno desta função não muda)
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
    Função interna que executa a importação, agora usando as configurações carregadas.
    """
    # ALTERAÇÃO: Carrega as configurações da planilha
    config = SageConfig(doc)
    all_data = {}
    
    # (A lógica de varrer os arquivos permanece a mesma)
    for root, _, files in os.walk(base_folder_path):
        for file_name in files:
            if not file_name.lower().endswith('.dat'):
                continue
            entidade_nome = os.path.splitext(file_name)[0].lower()
            if lista_entidades is not None and entidade_nome not in lista_entidades:
                continue
            full_path = os.path.join(root, file_name)
            relative_path = os.path.relpath(full_path, base_folder_path)
            entidades_validas_set = {os.path.splitext(f)[0].upper() for f in os.listdir(root) if f.lower().endswith('.dat')}
            parse_dat_file(full_path, relative_path, all_data, entidades_validas_set)

    # ALTERAÇÃO: Ordena as entidades a serem escritas com base na configuração
    entidades_importadas = all_data.keys()
    abas_ordenadas = sorted(
        entidades_importadas,
        key=lambda e: config.ordem_entidades.index(e) if e in config.ordem_entidades else float('inf')
    )

    # Lógica de escrita na planilha
    abas_a_escrever = lista_entidades if lista_entidades is not None else abas_ordenadas
    for entidade_nome in abas_a_escrever:
        pontos = all_data.get(entidade_nome)
        if pontos:
            # Passa o objeto de configuração para a função de escrita
            write_to_sheet(doc, entidade_nome, pontos, modo_importacao, config)


def write_to_sheet(doc, sheet_name, pontos_importados, modo, config):
    """
    Versão limpa e otimizada. Escreve os dados e aplica formatação visual básica,
    incluindo o efeito zebrado nas linhas importadas + 20 linhas extras.
    """
    # --- Bloco de Limpeza e Criação de Aba (sem alterações) ---
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

    # --- Aplicação de Cores de Aba e Ordenação de Colunas (sem alterações) ---
    cor_aba = config.cores_entidades.get(sheet_name.lower())
    if cor_aba is not None and cor_aba != -1:
        sheet.TabColor = cor_aba
    todos_atributos = {attr for p in pontos_importados if 'attributes' in p for attr in p['attributes']}
    atributos_ordenados = sorted(
        list(todos_atributos),
        key=lambda a: config.ordem_atributos.get(sheet_name.lower(), []).index(a) if a in config.ordem_atributos.get(sheet_name.lower(), []) else float('inf')
    )
    cabecalhos = [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS] + atributos_ordenados
    
    # --- Preenchimento dos Dados (sem alterações) ---
    for i, header_text in enumerate(cabecalhos):
        sheet.getCellByPosition(i, 0).setString(header_text)
    for row_idx, ponto in enumerate(pontos_importados, start=1):
        sheet.getCellByPosition(0, row_idx).setString(ponto.get('origem', ''))
        sheet.getCellByPosition(1, row_idx).setString(ponto['type'])
        if ponto['type'] in [CODIGO_COMENTARIO_SIMPLES, CODIGO_INCLUDE, CODIGO_INCLUDE_COMENTADO]:
            sheet.getCellByPosition(2, row_idx).setString(ponto.get('data', ''))
        elif 'attributes' in ponto:
            for attr_key, attr_value in ponto['attributes'].items():
                try: col_idx = cabecalhos.index(attr_key); sheet.getCellByPosition(col_idx, row_idx).setString(attr_value)
                except ValueError: pass

    # --- PACOTE DE POLIMENTO VISUAL SIMPLIFICADO ---
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    last_col = cursor.getRangeAddress().EndColumn
    last_row = cursor.getRangeAddress().EndRow
    
    # Formatação do Cabeçalho
    if last_row >= 0:
        header_range = sheet.getCellRangeByPosition(0, 0, last_col, 0)
        header_range.HoriJustify = 2 # CENTER
        if cor_aba is not None and cor_aba != -1:
            header_range.CellBackColor = cor_aba

    # Alinhamento da Coluna "Gera"
    try:
        gera_col_idx = cabecalhos.index(CABEÇALHO_COLUNA_CONTROLE)
        if last_row > 0:
            gera_col_range = sheet.getCellRangeByPosition(gera_col_idx, 1, gera_col_idx, last_row)
            gera_col_range.HoriJustify = 2
    except ValueError: pass 

    # Largura Ótima das Colunas
    columns = sheet.getColumns()
    for i in range(last_col + 1):
        columns.getByIndex(i).OptimalWidth = True

    # --- LÓGICA OTIMIZADA: PINTURA DE LINHAS ALTERNADAS ---
    # Aplica a formatação nas linhas de dados + 20 linhas extras.
    num_linhas_formatar = last_row + 21
    if last_row > 0:
        for r in range(1, num_linhas_formatar):
            cor_a_aplicar = COR_LINHA_IMPAR if r % 2 != 0 else COR_LINHA_PAR
            row_range = sheet.getCellRangeByPosition(0, r, last_col, r)
            row_range.CellBackColor = cor_a_aplicar

    # O BLOCO DE CÓDIGO PARA VALIDAÇÃO DE DADOS FOI COMPLETAMENTE REMOVIDO

# ===============================================================
# =================== LÓGICA DE PARSING =========================
# ===============================================================
# (A função parse_dat_file não muda nesta atualização)
def parse_dat_file(file_path, relative_path, all_data, entidades_validas):
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
        if not line: continue
        entidade_chave = os.path.splitext(os.path.basename(file_path))[0].lower()
        include_comentado_match = REGEX_INCLUDE_COMENTADO.match(original_line)
        if include_comentado_match:
            caminho_do_include = include_comentado_match.group(1).strip()
            ponto = {'type': CODIGO_INCLUDE_COMENTADO, 'data': caminho_do_include, 'origem': relative_path}
            all_data.setdefault(entidade_chave, []).append(ponto)
            continue
        include_match = REGEX_INCLUDE.match(original_line)
        if include_match:
            caminho_do_include = include_match.group(1).strip()
            ponto = {'type': CODIGO_INCLUDE, 'data': caminho_do_include, 'origem': relative_path}
            all_data.setdefault(entidade_chave, []).append(ponto)
            continue
        if line.startswith(';'):
            match = re.match(r'^; *([A-Z_]+)', line, re.IGNORECASE)
            if match and match.group(1).upper() in entidades_validas:
                ponto = {'type': CODIGO_BLOCO_COMENTADO, 'identifier': match.group(1).upper(), 'attributes': {}, 'origem': relative_path}
                while i < len(lines) and lines[i].strip().startswith(';'):
                    attr_line = lines[i].strip()[1:].strip()
                    if '=' in attr_line: key, value = attr_line.split('=', 1); ponto['attributes'][key.strip()] = value.strip()
                    i += 1
                entidade_ponto = ponto['identifier'].lower()
                all_data.setdefault(entidade_ponto, []).append(ponto)
            else:
                comentario_limpo = original_line.lstrip(';').lstrip()
                all_data.setdefault(entidade_chave, []).append({'type': CODIGO_COMENTARIO_SIMPLES, 'data': comentario_limpo, 'origem': relative_path})
            continue
        if line.upper() in entidades_validas:
            ponto = {'type': CODIGO_BLOCO_ATIVO, 'identifier': line.upper(), 'attributes': {}, 'origem': relative_path}
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

# ===============================================================
# ================= FUNÇÕES DE EXPORTAÇÃO =======================
# ===============================================================

def exportar_dats(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # ALTERAÇÃO: Usa a lista de folhas ignoradas
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
    abas_a_exportar = [s for s in doc.getSheets() if s.getName().lower() not in [ign.lower() for ign in FOLHAS_IGNORADAS]]
    erros = [_exportar_folha(sheet, export_folder) for sheet in abas_a_exportar]
    erros = [e for e in erros if e]
    
    if erros:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: {'; '.join(erros)}")
    else:
        geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("Exportação total concluída com sucesso!")


def exportar_parcial(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # (O código interno desta função não muda)
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
        # Garante que a aba ativa não seja uma aba ignorada
        if active_sheet_name.lower() not in [ign.lower() for ign in FOLHAS_IGNORADAS]:
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
    Exporta uma única aba, criando um backup (.bak) do arquivo anterior
    antes de salvar a nova versão.
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
        if not origem_path or not control_code or control_code == CODIGO_IGNORAR_LINHA: continue
        dados_agrupados_por_arquivo.setdefault(origem_path, [])
        bloco_final = None
        dado_principal = str(row_data[dados_col_idx])
        if control_code == CODIGO_INCLUDE and dado_principal:
            bloco_final = f'#include {dado_principal}'
        elif control_code == CODIGO_INCLUDE_COMENTADO and dado_principal:
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
            
    # Itera sobre os dados agrupados e escreve cada arquivo.
    for relative_path, file_content_list in dados_agrupados_por_arquivo.items():
        try:
            full_output_path = os.path.join(export_folder, relative_path)
            
            # --- INÍCIO DA LÓGICA DE BACKUP ---
            if os.path.exists(full_output_path):
                backup_path = full_output_path + ".bak"
                # Remove um backup antigo, se existir, para evitar erros no rename
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                os.rename(full_output_path, backup_path)
            # --- FIM DA LÓGICA DE BACKUP ---

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