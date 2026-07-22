# -*- coding: utf-8 -*-

import os
import re
import time

# ===============================================================
# ========== MACRO SAGE - TRILHA COMPLETA - 0.9.3+c1 ============
# ===============================================================
# Esta é a variante COMPLETA (forkada da Simples). Mantém todo o
# import/export da Simples e adiciona recursos avançados. Primeiro
# recurso: VERIFICADOR DE BASE (linter) — ver o bloco no fim do arquivo.
# Estratégia das duas trilhas documentada em PLANEJAMENTO.md.
# ===============================================================
# Este script é utilizado como macro no LibreOffice Calc para importar e exportar arquivos .dat
# do Sistema Aberto de Gerenciamento de Energia (SAGE).
#
# A partir desta versão, o script lê dinamicamente as configurações de
# ordenação, cores e validação das abas "MaisUsadas" e "EntidadesValoresAtributos".
#
# Adicionado a aba cores para facilitar a aplicação de temas de cores.
# Correção de importação de comentarios
#
# Alterações da versão 0.9.2 por Felipe Santos:
# - Compatibilidade com LibreOffice Flatpak usando macro embutida no documento.
# - Correção do nome da aba Geral e busca de abas tolerante a maiúsculas/minúsculas.
# - Tratamento mais seguro para erros de configuração.
# - A planilha SageBonis.ods agora pode carregar este script a partir do próprio documento,
#   usando URLs de macro com location=document.
# - Foram adicionados menu e barra de ferramentas SageBonis dentro do arquivo ODS.
#
# Alterações da versão 0.9.3:
# - Reincorporada a sanitização para latin-1 (_sanitizar_para_latin1) na exportação,
#   evitando UnicodeEncodeError quando há caracteres Unicode (aspas curvas, travessões,
#   etc.) que não existem no ISO-8859-1 esperado pelo SAGE.
# - Resultado da reconciliação entre a macro embutida no .ods (0.9.2) e o histórico do repo.
#
# Desenvolvido para rodar com a planilha SageBonis.ods
# Duvidas/Bugs/Sugestões - (11) 95456-4510 - Victor Bonesi - https://github.com/vbonesi/SageBonis

# ===============================================================
# ==================== CONFIGURAÇÃO GERAL =======================
# ===============================================================

# --- Nomes de Abas ---
# Importante: o LibreOffice diferencia maiúsculas/minúsculas em getByName().
# A planilha possui a aba "Geral" com G maiúsculo; usar "geral" causava falha
# nas rotinas de importação/exportação antes mesmo de ler os caminhos.
NOME_ABA_GERAL = "Geral"
NOME_ABA_MAIS_USADAS = "MaisUsadas"
NOME_ABA_VALIDACAO = "EntidadeAtributoValor"
NOME_ABA_OPMSK = "opmsk"
NOME_ABA_CORES = "Cores"

# --- Lista de Abas a Ignorar ---
FOLHAS_IGNORADAS = [NOME_ABA_GERAL, NOME_ABA_MAIS_USADAS, NOME_ABA_VALIDACAO, NOME_ABA_OPMSK, NOME_ABA_CORES]

# --- Posições das Células na Aba "Geral" ---
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
CABEÇALHO_COLUNA_DADOS = "Comentario/Include"

# --- NOVO: Cores para Linhas Alternadas (Efeito Zebra) ---
# Cores em formato numérico (Decimal de Hex BGR: Blue-Green-Red)
COR_LINHA_PAR = 16777215   # Branco (0xFFFFFF)
COR_LINHA_IMPAR = 15790320  # Cinza muito claro (0xF0F0F0)

# --- Constantes Técnicas ---
FLAGS_LIMPAR_TUDO = 1048575
LIMITE_CARACTERES_VALIDACAO = 250 # Manteremos para a próxima etapa

# --- Codificação dos Arquivos DAT do SAGE ---
ENCODING_EXPORTACAO_SAGE = 'latin-1'  # ISO-8859-1 (padrão esperado pelo SAGE)
ENCODINGS_IMPORTACAO_SAGE = ('latin-1', 'utf-8')  # Aceita os dois formatos na importação

# Mapeamento de caracteres Unicode comuns que não existem no ISO-8859-1
_UNICODE_PARA_LATIN1 = str.maketrans({
    '–': '-',    # en dash
    '—': '-',    # em dash
    '‘': "'",    # aspas simples esquerda
    '’': "'",    # aspas simples direita
    '“': '"',    # aspas duplas esquerda
    '”': '"',    # aspas duplas direita
    '…': '...',  # reticências
    ' ': ' ',    # espaço não-quebrável
    '•': '-',    # bullet
})

def _sanitizar_para_latin1(texto):
    texto = texto.translate(_UNICODE_PARA_LATIN1)
    return texto.encode('latin-1', errors='replace').decode('latin-1')


# --- Expressões Regulares ---
REGEX_INCLUDE = re.compile(r'^\s*#\s*include\s+(.*)', re.IGNORECASE)
REGEX_INCLUDE_COMENTADO = re.compile(r'^\s*;\s*#\s*include\s+(.*)', re.IGNORECASE)
REGEX_INICIO_BLOCO_COMENTADO = re.compile(r'^\s*;\s*([A-Z_]+)\s*$', re.IGNORECASE)

# --- Debug/Diagnóstico de Importação ---
DEBUG_IMPORTACAO = False
LOG_IMPORTACAO_RESUMO = True
LOG_IMPORTACAO_AVISOS = True
WATCHDOG_MAX_ITERACOES_SEM_PROGRESSO = 1000


def _get_sheet(doc, sheet_name):
    """
    Obtém uma aba pelo nome, aceitando diferença de maiúsculas/minúsculas.

    Esta função foi adicionada para deixar a macro mais resistente a variações
    nos nomes das abas. Sem ela, qualquer diferença como "Geral" vs "geral"
    faz o LibreOffice lançar exceção em getByName().
    """
    sheets = doc.getSheets()
    if sheets.hasByName(sheet_name):
        return sheets.getByName(sheet_name)

    wanted = sheet_name.lower()
    for sheet in sheets:
        if sheet.getName().lower() == wanted:
            return sheet

    raise KeyError(f"Aba '{sheet_name}' não encontrada.")


def _log_importacao(level, message, force=False):
    """Logger simples e opcional para diagnóstico da importação."""
    if force:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level == 'DEBUG' and DEBUG_IMPORTACAO:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level == 'INFO' and LOG_IMPORTACAO_RESUMO:
        print(f"[IMPORTACAO:{level}] {message}")
        return
    if level in ['WARN', 'ERROR'] and LOG_IMPORTACAO_AVISOS:
        print(f"[IMPORTACAO:{level}] {message}")


def _classificar_linha_dat(raw_line, entidades_validas):
    """Classifica a linha do arquivo DAT para manter o parser determinístico."""
    original_line = raw_line.strip('\r\n')
    stripped_line = original_line.strip()

    if not stripped_line:
        return {'type': 'blank', 'original': original_line, 'stripped': stripped_line}

    include_comentado_match = REGEX_INCLUDE_COMENTADO.match(original_line)
    if include_comentado_match:
        return {
            'type': 'include_commented',
            'original': original_line,
            'stripped': stripped_line,
            'value': include_comentado_match.group(1).strip()
        }

    include_match = REGEX_INCLUDE.match(original_line)
    if include_match:
        return {
            'type': 'include',
            'original': original_line,
            'stripped': stripped_line,
            'value': include_match.group(1).strip()
        }

    if stripped_line.upper() in entidades_validas:
        return {
            'type': 'entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': stripped_line.upper()
        }

    commented_block_match = REGEX_INICIO_BLOCO_COMENTADO.match(original_line)
    if commented_block_match and commented_block_match.group(1).upper() in entidades_validas:
        return {
            'type': 'commented_entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': commented_block_match.group(1).upper()
        }

    if stripped_line.startswith(';'):
        return {
            'type': 'comment',
            'original': original_line,
            'stripped': stripped_line,
            'value': original_line.lstrip(';').lstrip()
        }

    if '=' in stripped_line:
        key, value = stripped_line.split('=', 1)
        return {
            'type': 'attribute',
            'original': original_line,
            'stripped': stripped_line,
            'key': key.strip(),
            'value': value.strip()
        }

    return {'type': 'invalid', 'original': original_line, 'stripped': stripped_line}


def _classificar_linha_bloco_comentado(raw_line, entidades_validas):
    """
    Classifica linhas quando estamos dentro de um bloco comentado.
    O conteúdo útil está sempre após o primeiro ';'.
    """
    original_line = raw_line.strip('\r\n')
    stripped_line = original_line.strip()

    if not stripped_line:
        return {'type': 'blank', 'original': original_line, 'stripped': stripped_line}

    if not stripped_line.startswith(';'):
        return {'type': 'block_end', 'original': original_line, 'stripped': stripped_line}

    inner_line = stripped_line[1:].strip()
    if not inner_line:
        return {'type': 'comment', 'original': original_line, 'stripped': stripped_line, 'value': ''}

    if inner_line.upper() in entidades_validas:
        return {
            'type': 'commented_entity_start',
            'original': original_line,
            'stripped': stripped_line,
            'entity': inner_line.upper()
        }

    if REGEX_INCLUDE_COMENTADO.match(stripped_line):
        return {'type': 'block_end', 'original': original_line, 'stripped': stripped_line}

    if '=' in inner_line:
        key, value = inner_line.split('=', 1)
        return {
            'type': 'attribute',
            'original': original_line,
            'stripped': stripped_line,
            'key': key.strip(),
            'value': value.strip()
        }

    return {'type': 'comment', 'original': original_line, 'stripped': stripped_line, 'value': inner_line}


def _iniciar_bloco(entidade_nome, tipo_bloco, relative_path, comentarios_iniciais=None):
    bloco = {
        'type': tipo_bloco,
        'identifier': entidade_nome,
        'attributes': {},
        'comments': [],
        'origem': relative_path
    }
    if comentarios_iniciais:
        bloco['comments'].extend(comentarios_iniciais)
    return bloco


def _finalizar_bloco(current_block, all_data, relative_path, stats, line_no):
    if not current_block:
        return

    ponto = {
        'type': current_block['type'],
        'identifier': current_block['identifier'],
        'attributes': current_block['attributes'],
        'origem': relative_path
    }

    if current_block['comments']:
        ponto['comment'] = "\n".join(current_block['comments'])

    if not current_block['attributes']:
        stats['warnings'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} bloco {current_block['identifier']} finalizado sem atributos.",
            force=True
        )

    if current_block['attributes'] and 'ID' not in current_block['attributes']:
        stats['warnings'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} bloco {current_block['identifier']} finalizado sem ID.",
            force=True
        )

    if current_block['attributes'] or current_block['comments']:
        chave = current_block['identifier'].lower()
        all_data.setdefault(chave, []).append(ponto)
        stats['entities_imported'] += 1

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
            sheet = _get_sheet(self.doc, NOME_ABA_MAIS_USADAS)
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
    # Mantém a variável inicializada para que o bloco except não tente escrever
    # status em uma aba que falhou ao ser localizada.
    geral_sheet = None
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
        return

    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Processando importação total...")
    _executar_importacao(doc, folder_path, lista_entidades=None, modo_importacao='REPLACE')
    geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("Importação total concluída com sucesso!")


def importar_parcial(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    # A aba Geral pode não ser encontrada se o arquivo for alterado manualmente;
    # por isso o tratamento de erro precisa funcionar mesmo sem geral_sheet.
    geral_sheet = None
    
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_IMPORTACAO)
        folder_path = path_cell.getString()
        if not os.path.isdir(folder_path):
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString("ERRO: O caminho especificado não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_IMPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
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
    prioridade_entidades = {entidade: idx for idx, entidade in enumerate(config.ordem_entidades)}
    
    # (A lógica de varrer os arquivos permanece a mesma)
    for root, _, files in os.walk(base_folder_path):
        entidades_validas_set = {os.path.splitext(f)[0].upper() for f in files if f.lower().endswith('.dat')}
        for file_name in files:
            if not file_name.lower().endswith('.dat'):
                continue
            entidade_nome = os.path.splitext(file_name)[0].lower()
            if lista_entidades is not None and entidade_nome not in lista_entidades:
                continue
            full_path = os.path.join(root, file_name)
            relative_path = os.path.relpath(full_path, base_folder_path)
            parse_dat_file(full_path, relative_path, all_data, entidades_validas_set)

    # ALTERAÇÃO: Ordena as entidades a serem escritas com base na configuração
    entidades_importadas = all_data.keys()
    abas_ordenadas = sorted(
        entidades_importadas,
        key=lambda e: prioridade_entidades.get(e, float('inf'))
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
        sheet = _get_sheet(doc, sheet_name)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        range_to_clear = sheet.getCellRangeByPosition(0, 0, cursor.getRangeAddress().EndColumn, cursor.getRangeAddress().EndRow)
        range_to_clear.clearContents(FLAGS_LIMPAR_TUDO)
    else:
        if doc.getSheets().hasByName(sheet_name):
            doc.getSheets().removeByName(sheet_name)
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(sheet_name, new_sheet)
        sheet = _get_sheet(doc, sheet_name)

    # --- Aplicação de Cores de Aba e Ordenação de Colunas (sem alterações) ---
    cor_aba = config.cores_entidades.get(sheet_name.lower())
    if cor_aba is not None and cor_aba != -1:
        sheet.TabColor = cor_aba
    todos_atributos = {attr for p in pontos_importados if 'attributes' in p for attr in p['attributes']}
    ordem_atributos_aba = config.ordem_atributos.get(sheet_name.lower(), [])
    prioridade_atributos = {attr: idx for idx, attr in enumerate(ordem_atributos_aba)}
    atributos_ordenados = sorted(
        list(todos_atributos),
        key=lambda a: prioridade_atributos.get(a, float('inf'))
    )
    cabecalhos = [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS] + atributos_ordenados
    header_to_col = {header: idx for idx, header in enumerate(cabecalhos)}
    
    # --- Preenchimento dos Dados (agora em lote para reduzir chamadas UNO) ---
    data_matrix = [cabecalhos]
    for ponto in pontos_importados:
        row_data = [''] * len(cabecalhos)
        row_data[0] = ponto.get('origem', '')
        row_data[1] = ponto['type']
        if ponto['type'] in [CODIGO_COMENTARIO_SIMPLES, CODIGO_INCLUDE, CODIGO_INCLUDE_COMENTADO]:
            row_data[2] = ponto.get('data', '')
        elif ponto['type'] in [CODIGO_BLOCO_ATIVO, CODIGO_BLOCO_COMENTADO]:
            row_data[2] = ponto.get('comment', '')
        if 'attributes' in ponto:
            for attr_key, attr_value in ponto['attributes'].items():
                col_idx = header_to_col.get(attr_key)
                if col_idx is not None:
                    row_data[col_idx] = attr_value
        data_matrix.append(row_data)

    if data_matrix:
        num_rows = len(data_matrix) - 1
        num_cols = len(cabecalhos) - 1
        target_range = sheet.getCellRangeByPosition(0, 0, num_cols, num_rows)
        target_range.setDataArray(tuple(tuple(str(cell) for cell in row) for row in data_matrix))

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
def parse_dat_file(file_path, relative_path, all_data, entidades_validas):
    start_time = time.perf_counter()
    lines = None
    for encoding in ENCODINGS_IMPORTACAO_SAGE:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue
        except IOError as e:
            print(f"Erro ao ler o arquivo {file_path}: {e}")
            return

    if lines is None:
        # Fallback resiliente para evitar falha total em arquivos com bytes inválidos.
        try:
            with open(file_path, 'r', encoding=ENCODING_EXPORTACAO_SAGE, errors='ignore') as f:
                lines = f.readlines()
        except IOError as e:
            print(f"Erro ao ler o arquivo {file_path}: {e}")
            return
        
    i = 0
    current_entidade_chave = os.path.splitext(os.path.basename(file_path))[0].lower()
    pending_comments = []
    current_block = None
    stats = {
        'lines_total': len(lines),
        'entities_imported': 0,
        'comments': 0,
        'ignored_lines': 0,
        'invalid_lines': 0,
        'warnings': 0
    }
    last_i = -1
    iteracoes_sem_progresso = 0

    _log_importacao('DEBUG', f"Iniciando parse de {relative_path} com {len(lines)} linhas.")

    while i < len(lines):
        if i == last_i:
            iteracoes_sem_progresso += 1
            if iteracoes_sem_progresso >= WATCHDOG_MAX_ITERACOES_SEM_PROGRESSO:
                raise RuntimeError(
                    f"Watchdog de importação disparado em {relative_path} na linha {i + 1}: iterações sem avanço."
                )
        else:
            iteracoes_sem_progresso = 0
            last_i = i

        raw_line = lines[i]
        line_no = i + 1

        if current_block and current_block['type'] == CODIGO_BLOCO_COMENTADO:
            line_info = _classificar_linha_bloco_comentado(raw_line, entidades_validas)
        else:
            line_info = _classificar_linha_dat(raw_line, entidades_validas)

        _log_importacao(
            'DEBUG',
            f"{relative_path}:{line_no} bloco={current_block['identifier'] if current_block else '-'} tipo={line_info['type']}"
        )

        if current_block:
            if line_info['type'] in ['entity_start', 'commented_entity_start', 'include', 'include_commented', 'block_end']:
                _finalizar_bloco(current_block, all_data, relative_path, stats, line_no)
                current_block = None
                continue

            if line_info['type'] == 'blank':
                stats['ignored_lines'] += 1
                i += 1
                continue

            if line_info['type'] == 'comment':
                current_block['comments'].append(line_info.get('value', ''))
                stats['comments'] += 1
                i += 1
                continue

            if line_info['type'] == 'attribute':
                current_block['attributes'][line_info['key']] = line_info['value']
                i += 1
                continue

            stats['invalid_lines'] += 1
            stats['warnings'] += 1
            _log_importacao(
                'WARN',
                f"{relative_path}:{line_no} linha inválida dentro do bloco {current_block['identifier']}: {line_info['original']}",
                force=True
            )
            i += 1
            continue

        if line_info['type'] == 'blank':
            stats['ignored_lines'] += 1
            i += 1
            continue

        if line_info['type'] == 'include_commented':
            ponto = {'type': CODIGO_INCLUDE_COMENTADO, 'data': line_info['value'], 'origem': relative_path}
            all_data.setdefault(current_entidade_chave, []).append(ponto)
            if pending_comments:
                stats['warnings'] += 1
                _log_importacao(
                    'WARN',
                    f"{relative_path}:{line_no} comentários pendentes descartados antes de include comentado.",
                    force=True
                )
                pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'include':
            ponto = {'type': CODIGO_INCLUDE, 'data': line_info['value'], 'origem': relative_path}
            all_data.setdefault(current_entidade_chave, []).append(ponto)
            if pending_comments:
                stats['warnings'] += 1
                _log_importacao(
                    'WARN',
                    f"{relative_path}:{line_no} comentários pendentes descartados antes de include.",
                    force=True
                )
                pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'entity_start':
            entidade_nome = line_info['entity']
            current_entidade_chave = entidade_nome.lower()
            current_block = _iniciar_bloco(
                entidade_nome,
                CODIGO_BLOCO_ATIVO,
                relative_path,
                comentarios_iniciais=pending_comments
            )
            pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'commented_entity_start':
            entidade_nome = line_info['entity']
            current_entidade_chave = entidade_nome.lower()
            current_block = _iniciar_bloco(
                entidade_nome,
                CODIGO_BLOCO_COMENTADO,
                relative_path,
                comentarios_iniciais=pending_comments
            )
            pending_comments = []
            i += 1
            continue

        if line_info['type'] == 'comment':
            pending_comments.append(line_info.get('value', ''))
            stats['comments'] += 1
            i += 1
            continue

        if line_info['type'] == 'attribute':
            stats['warnings'] += 1
            stats['invalid_lines'] += 1
            _log_importacao(
                'WARN',
                f"{relative_path}:{line_no} atributo fora de bloco ignorado: {line_info['original']}",
                force=True
            )
            i += 1
            continue

        stats['warnings'] += 1
        stats['invalid_lines'] += 1
        _log_importacao(
            'WARN',
            f"{relative_path}:{line_no} linha não reconhecida ignorada: {line_info['original']}",
            force=True
        )
        i += 1

    if current_block:
        _finalizar_bloco(current_block, all_data, relative_path, stats, len(lines))

    elapsed = time.perf_counter() - start_time
    _log_importacao(
        'INFO',
        (
            f"Arquivo {relative_path} processado em {elapsed:.3f}s. "
            f"linhas={stats['lines_total']} entidades={stats['entities_imported']} "
            f"comentarios={stats['comments']} ignoradas={stats['ignored_lines']} "
            f"invalidas={stats['invalid_lines']} avisos={stats['warnings']}"
        )
    )

# ===============================================================
# ================= FUNÇÕES DE EXPORTAÇÃO =======================
# ===============================================================

def exportar_dats(*args):
    doc = XSCRIPTCONTEXT.getDocument() # type: ignore
    # Mesma proteção das rotinas de importação: evita erro secundário no except
    # quando a aba Geral não existe ou foi renomeada.
    geral_sheet = None
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()
        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
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
    controller = doc.getCurrentController()
    active_sheet = controller.getActiveSheet()
    active_sheet_name = active_sheet.getName()
    # Exportação parcial também depende da aba Geral para ler pasta/lista.
    # Inicializar com None permite emitir diagnóstico seguro em caso de falha.
    geral_sheet = None
    
    try:
        geral_sheet = _get_sheet(doc, NOME_ABA_GERAL)
        export_path_cell = geral_sheet.getCellByPosition(*CELULA_CAMINHO_EXPORTACAO)
        export_folder = export_path_cell.getString()
        if not os.path.isdir(export_folder):
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString("ERRO: O caminho de destino não é uma pasta válida.")
            return
    except Exception as e:
        if geral_sheet:
            geral_sheet.getCellByPosition(*CELULA_STATUS_EXPORTACAO).setString(f"ERRO: Falha ao ler configurações. {e}")
        else:
            print(f"ERRO: Falha ao localizar a aba '{NOME_ABA_GERAL}'. {e}")
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
            try:
                abas_a_exportar.append(_get_sheet(doc, nome))
            except KeyError:
                pass
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
            comment_lines = [line for line in dado_principal.splitlines()]
            attribute_lines = []
            for col_idx, header in enumerate(headers):
                if header in [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS]:
                    continue
                value = str(row_data[col_idx]) if len(row_data) > col_idx else ""
                if value:
                    attribute_lines.append(f"\t{header} = {value}")

            if comment_lines or attribute_lines:
                if control_code == CODIGO_BLOCO_COMENTADO:
                    point_lines = [f";{sheet_name.upper()}"]
                    point_lines.extend([f";{line}" for line in comment_lines])
                    point_lines.extend([f";{line}" for line in attribute_lines])
                else:
                    point_lines = [sheet_name.upper()]
                    point_lines.extend([f";{line}" for line in comment_lines])
                    point_lines.extend(attribute_lines)
                bloco_final = "\n".join(point_lines)
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
            conteudo = _sanitizar_para_latin1("\n\n".join(file_content_list) + "\n")
            with open(full_output_path, 'w', encoding=ENCODING_EXPORTACAO_SAGE) as f:
                f.write(conteudo)
        except IOError as e:
            return f"Falha ao escrever {relative_path}: {e}"
            
    return None

# ===============================================================
# ================= FUNÇÃO DE CORES DO TEMA =====================
# ===============================================================

# --- CONFIGURAÇÃO DA ABA DE CORES ---
# Verifique o nome da aba onde a tabela de cores se encontra.
# Se for "TEMA ESCURO" use-o, caso contrário ajuste.
NOME_ABA_TEMA_CORES = "Cores" 
# Colunas:
COL_R_DEC = 5  # Coluna F (índice 5)
COL_G_DEC = 6  # Coluna G (índice 6)
COL_B_DEC = 7  # Coluna H (índice 7)
COL_COR_AMOSTRA = 11 # Coluna L (índice 11)


def rgb_to_bgr_decimal(r, g, b):
    """
    Converte os valores RGB (Red, Green, Blue) de 0-255 para
    o formato BGR Decimal (Blue-Green-Red), que é o padrão
    de cor numérica para CellBackColor no LibreOffice.
    Fórmula: (B * 256^2) + (G * 256^1) + (R * 256^0)
    """
    try:
        # Garante que os valores são inteiros e no intervalo 0-255.
        r = int(r) if 0 <= r <= 255 else 0
        g = int(g) if 0 <= g <= 255 else 0
        b = int(b) if 0 <= b <= 255 else 0
        
        # Se os três valores forem zero (preto), retorna a cor
        if r == 0 and g == 0 and b == 0:
             return 0 # Preto é o valor 0
        
        # Se houver algum valor diferente de zero, calcula o BGR.
        # B é o mais significativo (<< 16), G é o intermediário (<< 8), R é o menos significativo.
        bgr_decimal = (r * 65536) + (g * 256) + b
        return bgr_decimal
    except:
        # Retorna -1 para sinalizar que a célula deve ficar sem cor
        return -1


def atualizar_amostras_cores(*args):
    """
    Lê os valores RGB Decimais (Colunas H, I, J) da aba do tema
    e aplica a cor de fundo (CellBackColor) na coluna de amostra (L).
    """
    try:
        doc = XSCRIPTCONTEXT.getDocument() # type: ignore
        sheets = doc.getSheets()

        if not sheets.hasByName(NOME_ABA_TEMA_CORES):
            print(f"ERRO: A aba de tema '{NOME_ABA_TEMA_CORES}' não foi encontrada.")
            return

        sheet = _get_sheet(doc, NOME_ABA_TEMA_CORES)
        
        # 1. Determina a última linha preenchida para otimizar a leitura
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_row = cursor.getRangeAddress().EndRow
        
        # 2. Leitura otimizada de um bloco de dados (Colunas R DEC a B DEC)
        # Lemos de R DEC (H) até B DEC (J) da linha 1 até a última.
        # Os índices iniciais são: COL_R_DEC (7) e LINHA 1.
        data_range = sheet.getCellRangeByPosition(COL_R_DEC, 1, COL_B_DEC, last_row)
        data = data_range.getDataArray()
        
        # 3. Itera sobre os dados lidos
        for row_idx, row_data in enumerate(data):
            # Índices relativos ao bloco de dados lido: 0=R, 1=G, 2=B
            r_dec = row_data[0]
            g_dec = row_data[1]
            b_dec = row_data[2]
            
            # Converte para BGR Decimal
            bgr_cor = rgb_to_bgr_decimal(r_dec, g_dec, b_dec)
            
            # A linha de destino é o índice da linha atual (row_idx) + 1 (cabeçalho)
            target_row = row_idx + 1 
            
            amostra_cell = sheet.getCellByPosition(COL_COR_AMOSTRA, target_row)
            
            if bgr_cor == -1:
                # Se a conversão falhar ou os valores não existirem, remove a cor
                amostra_cell.CellBackColor = -1 # O valor -1 no LibreOffice remove o preenchimento
            else:
                # Aplica a cor
                amostra_cell.CellBackColor = bgr_cor
                
        print("Amostras de cores do tema atualizadas com sucesso!")

    except Exception as e:
        print(f"ERRO ao aplicar as cores do tema: {e}")

# ===============================================================
# ============ TRILHA COMPLETA: VERIFICADOR DE BASE =============
# ===============================================================
# Linter de integridade da base SAGE. Lê as abas de entidade e reporta os
# achados numa aba "Análise". É read-only sobre as entidades (nunca altera
# dados de ponto), então não há risco para a exportação.
#
# Princípio de design (ver PLANEJAMENTO.md): as regras de integridade
# referencial NÃO ficam fixas no código — moram na aba de config
# "VerificacaoRefs", que é criada com exemplos (inativos) na primeira execução
# para o usuário ativar/ajustar conforme o padrão da sua base.

NOME_ABA_ANALISE = "Análise"
NOME_ABA_VERIFICACAO_REFS = "VerificacaoRefs"

# Severidades dos achados
SEV_ERRO = "ERRO"
SEV_AVISO = "AVISO"
SEV_OK = "OK"
SEV_INFO = "INFO"

CABECALHOS_ANALISE = ["Severidade", "Entidade", "Linha", "Atributo", "Valor", "Descrição"]
CABECALHOS_REFS = ["EntidadeOrigem", "AtributoOrigem", "EntidadeDestino", "AtributoDestino", "Ativa"]

# Valores aceitos na coluna "Ativa" da aba VerificacaoRefs.
_VALORES_ATIVO = ('s', 'sim', 'x', '1', 'true', 'v')

# Delimitador para múltiplos destinos numa regra de integridade referencial -- modela FK
# "ambígua" do SAGE (ex.: PDF.PNT pode ser um PDS na aquisição OU um PDD na distribuição).
_DELIM_MULTI_DESTINO = "|"


def _parse_entidades_destino(texto):
    """'PDS|PDD' -> ('PDS', 'PDD'); 'PDS' -> ('PDS',). Espaços em volta do delimitador são ok."""
    return tuple(p.strip() for p in str(texto).split(_DELIM_MULTI_DESTINO) if p.strip())


# Regras de integridade referencial gravadas na aba de config na 1ª execução (todas INATIVAS
# por padrão -- o usuário ativa as que fazem sentido para a sua base). Curadas a partir do
# levantamento de entidades/relacionamentos do projeto SkillSAGE (fk_graph.csv, ~85 arestas
# derivadas dos manuais CEPEL + bases reais), excluindo relações para entidades-CATÁLOGO do
# SAGE (TN1/TN2/TCV/TTP): são valores pré-definidos pelo CEPEL sem aba própria na planilha, e
# portanto não há o que checar aqui (ver a checagem de domínio, _check_dominios, para esses
# casos -- ex.: NV2.TN2 validado contra a lista de valores válidos, não contra uma aba "TN2").
# Formato de cada linha: "EntidadeOrigem  AtributoOrigem  EntidadeDestino[|EntidadeDestino2...]"
# -- o atributo de destino é sempre "ID" (é assim que toda FK do SAGE funciona: aponta a chave
# primária da entidade referenciada).
_REGRAS_REFS_PADRAO_TXT = """
INP      NOH     NOH
INP      PRO     PRO
NOCT     NOH     NOH
NOCT     CTX     CTX
PRCT     PRO     PRO
PRCT     CTX     CTX
PRCT     AOR     AOR
SXP      PRO     PRO
SXP      SEV     SEV
OCR      SEVER   SEV
OCR      GRPOCR  GRPOCR
OCR      TELA    TELA
GSD      NO1     NOH
GSD      NO2     NOH
GSD      SITE    SITE
CXU      GSD     GSD
ENU      CXU     CXU
UTR      CNF     CNF
UTR      CXU     CXU
MUL      CNF     CNF
MUL      GSD     GSD
ENM      MUL     MUL
CNM      MUL     MUL
CNF      LSC     LSC
NV1      CNF     CNF
NV2      NV1     NV1
PDF      NV2     NV2
PDF      PNT     PDS|PDD
PAF      NV2     NV2
PAF      PNT     PAS|PAD
CGF      NV2     NV2
CGF      CNF     CNF
CGF      CGS     CGS
PTF      NV2     NV2
PTF      PNT     PTS
PIF      NV2     NV2
PIF      PNT     PIS
LSC      GSD     GSD
LSC      MAP     MAP
INS      AOR     AOR
INS      PTC     PTC
INS      TELA    TELA
TAC      INS     INS
TAC      LSC     LSC
PAS      TAC     TAC
PAS      TCL     TCL
PAS      OCR     OCR
PAS      PTC     PTC
PDS      TAC     TAC
PDS      TCL     TCL
PDS      OCR     OCR
PTS      TAC     TAC
PTS      OCR     OCR
CGS      TAC     TAC
CGS      PINT    PDS|PAS
CGS      PAC     PDS|PAS
CGS      TIPOE   TCTL
RCA      PARC    PDS|PAS|PTS
RCA      PNT     PDS|PAS|PTS
RFC      PARC    PDF|PAF|PTF
RFC      PNT     PDS|PAS|PTS
RFI      PNT     PDF|PAF|PTF
TDD      LSC     LSC
PDD      PDS     PDS
PDD      TDD     TDD
PAD      PAS     PAS
PAD      TDD     TDD
PTD      PTS     PTS
PTD      TDD     TDD
E2M      IDPTO   OCR|PAS|PDS|PTS
E2M      MAP     MAP
GRCMP    GRUPO   GRUPO
GRCMP    PNT     PDS|PAS|CGS|GRUPO
GRCMP    ACAO    ACAO
GRUPO    PNT     PDS|PAS|PTS|CGS|GRUPO
GRUPO    COR     COR
GR2ACT   ACAO    ACAO
GR2ACT   GRACT   GRACT
AUTOZ    AOR     AOR
AUTOZ    GRACT   GRACT
AUTOZ    PAPEL   PAPEL
"""


def _montar_regras_refs_padrao():
    """Parseia _REGRAS_REFS_PADRAO_TXT -> [(EntOrigem, AtrOrigem, EntDestino, 'ID'), ...].
    EntDestino fica como string (possivelmente "A|B") -- só vira tupla ao ser lida de volta da
    aba, em _carregar_regras_refs, o mesmo caminho usado para regras editadas pelo usuário."""
    regras = []
    for linha in _REGRAS_REFS_PADRAO_TXT.strip().splitlines():
        partes = linha.split()
        if len(partes) != 3:
            continue
        ent_o, attr_o, ent_d = partes
        regras.append((ent_o, attr_o, ent_d, "ID"))
    return regras


REGRAS_REFS_PADRAO = _montar_regras_refs_padrao()

# Ordem de severidade para ordenar o relatório (erros primeiro).
_ORDEM_SEV = {SEV_ERRO: 0, SEV_AVISO: 1, SEV_INFO: 2, SEV_OK: 3}

# Peso de fonte "negrito" no LibreOffice (com.sun.star.awt.FontWeight.BOLD).
_FONT_BOLD = 150.0


def _aba_existe_ci(doc, nome):
    """True se existe uma aba com esse nome (ignorando maiúsc/minúsc)."""
    sheets = doc.getSheets()
    if sheets.hasByName(nome):
        return True
    alvo = nome.strip().lower()
    for sheet in sheets:
        if sheet.getName().strip().lower() == alvo:
            return True
    return False


def _ler_entidade(sheet):
    """Lê uma aba -> (headers:list[str], linhas:list[tuple]). (None, None) se vazia."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    addr = cursor.getRangeAddress()
    data = sheet.getCellRangeByPosition(0, 0, addr.EndColumn, addr.EndRow).getDataArray()
    if not data or len(data) < 2:
        return None, None
    headers = [str(h) for h in data[0]]
    return headers, list(data[1:])


def _idx_coluna(headers, nome):
    """Índice da coluna pelo nome (case-insensitive); -1 se não existir."""
    alvo = str(nome).strip().lower()
    for i, h in enumerate(headers):
        if str(h).strip().lower() == alvo:
            return i
    return -1


def _is_ponto_ativo(row, col_gera):
    """True se a linha é um ponto ATIVO (coluna Gera == 'x')."""
    if col_gera < 0 or len(row) <= col_gera:
        return False
    return str(row[col_gera]).strip().lower() == CODIGO_BLOCO_ATIVO


class _Analise:
    """Acumula os achados do verificador."""
    def __init__(self):
        self.achados = []

    def add(self, sev, entidade, linha, atributo, valor, descr):
        self.achados.append({
            "sev": sev, "entidade": entidade, "linha": linha,
            "atributo": atributo, "valor": valor, "descr": descr,
        })

    @property
    def erros(self):
        return sum(1 for a in self.achados if a["sev"] == SEV_ERRO)

    @property
    def avisos(self):
        return sum(1 for a in self.achados if a["sev"] == SEV_AVISO)


# --- Checagens individuais (cada uma só adiciona achados; fácil estender) ---

def _check_ids(sheet_name, headers, linhas, analise):
    """ID vazio em ponto ativo e IDs duplicados dentro da entidade."""
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    col_id = _idx_coluna(headers, "ID")
    if col_id < 0:
        return  # entidade sem coluna ID — nada a checar aqui
    vistos = {}
    for i, row in enumerate(linhas):
        if not _is_ponto_ativo(row, col_gera):
            continue
        linha_planilha = i + 2  # +1 cabeçalho, +1 base-1
        valor = str(row[col_id]).strip() if len(row) > col_id else ""
        if not valor:
            analise.add(SEV_ERRO, sheet_name, linha_planilha, "ID", "", "Ponto ativo sem ID")
            continue
        if valor in vistos:
            analise.add(SEV_ERRO, sheet_name, linha_planilha, "ID", valor,
                        "ID duplicado (1a ocorrencia na linha %d)" % vistos[valor])
        else:
            vistos[valor] = linha_planilha


# Restrição de tamanho de ID por entidade -- fonte: SkillSAGE (_PADRAO/convencao_nomenclatura.md,
# "Restrição SAGE"), ainda não conferida contra o manual CEPEL oficial correspondente. Por isso a
# checagem usa AVISO, não ERRO: um limite errado aqui não deveria travar a confiança na ferramenta.
LIMITES_TAMANHO_ID = {
    "GSD": 8, "LSC": 8,
    "TAC": 12,
    "CNF": 16, "NV1": 16, "UTR": 16, "ENU": 16,
    "NV2": 40,
    "PDS": 32,
}


def _check_tamanho_id(sheet_name, headers, linhas, analise):
    """ID mais longo que o limite conhecido da entidade (ver LIMITES_TAMANHO_ID)."""
    limite = LIMITES_TAMANHO_ID.get(sheet_name.strip().upper())
    if limite is None:
        return
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    col_id = _idx_coluna(headers, "ID")
    if col_id < 0:
        return
    for i, row in enumerate(linhas):
        if not _is_ponto_ativo(row, col_gera):
            continue
        valor = str(row[col_id]).strip() if len(row) > col_id else ""
        if valor and len(valor) > limite:
            analise.add(SEV_AVISO, sheet_name, i + 2, "ID", valor,
                        "ID com %d caracteres, acima do limite conhecido (%d) para %s" %
                        (len(valor), limite, sheet_name))


# Pares entidade.atributo do dominios.csv (SkillSAGE) ainda ausentes na aba EntidadeAtributoValor
# desta planilha (conferido em 2026-07: a aba já cobre ~106 pares, é mais completa que o SkillSAGE
# na maioria dos casos -- só estes 5 faltavam). Só entram se a aba ainda não tiver a própria regra
# para o mesmo par -- nunca sobrescrevem o que o usuário já tem.
_DOMINIOS_SUPLEMENTARES = {
    ("cgf", "KCONV"): {"SBOw TERM", "SBOw", "DIR TERM", "NUL_1_D", "NUL_D", "CO_1"},
    ("grupo", "APLIC"): {"VTelas", "OUTROS"},
    ("tctl", "ID"): {"CTCL", "BLOQ", "HABD", "LIGD", "RSTC", "AUMD"},
    ("tn2", "TIPO"): {"ADAQ", "AAAQ", "ADUP", "CSIM", "CDUP", "AANL", "ALAT", "AMCD", "APFL",
                       "AA32", "ASTP", "ASIM", "CREL", "CSTP"},
    ("utr", "ORDEM"): {"PRI", "REV"},
}


def _mesclar_dominios_suplementares(dominios):
    """Preenche pares entidade.atributo do _DOMINIOS_SUPLEMENTARES que a aba ainda não tem."""
    for chave, valores in _DOMINIOS_SUPLEMENTARES.items():
        dominios.setdefault(chave, valores)


def _carregar_dominios(doc):
    """Lê a aba 'EntidadeAtributoValor' (já existente na planilha; formato largo
    Entidade|Atributo|Valor1..N) -> {(entidade.lower, atributo.upper): set(valores)}.

    Diferente de VerificacaoRefs, esta aba NÃO é criada aqui se faltar -- ela já é uma aba de
    config estabelecida da Trilha Completa (usada no passado para validação de dropdown; ver
    comentário em SageConfig). Se a aba não existir, a checagem de domínio simplesmente não roda."""
    dominios = {}
    if not _aba_existe_ci(doc, NOME_ABA_VALIDACAO):
        return dominios
    sheet = _get_sheet(doc, NOME_ABA_VALIDACAO)
    headers, linhas = _ler_entidade(sheet)
    if not linhas:
        return dominios
    for row in linhas:
        if len(row) < 3:
            continue
        entidade = str(row[0]).strip().lower()
        atributo = str(row[1]).strip().upper()
        if not entidade or not atributo:
            continue
        valores = {str(v).strip() for v in row[2:] if str(v).strip()}
        if valores:
            dominios[(entidade, atributo)] = valores
    _mesclar_dominios_suplementares(dominios)
    return dominios


# Colunas de controle/metadado da planilha (existem em toda aba de entidade) -- nunca são
# atributos SAGE de verdade, então ficam fora da checagem de domínio mesmo que colidam (ex.:
# a coluna de controle "Origem", que guarda o arquivo de include da linha, não tem nada a ver
# com o atributo real "PAS.ORIGEM" do SAGE -- SCADA/MONRES/RCALC/PDO -- apesar do mesmo nome).
_COLUNAS_CONTROLE = frozenset(
    c.strip().lower() for c in
    (CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_DADOS)
)


def _check_dominios(sheet_name, headers, linhas, dominios, analise):
    """Valor de atributo fora do domínio conhecido (aba EntidadeAtributoValor).

    Severidade AVISO: a aba pode não cobrir todos os valores reais de uma base específica,
    então um "fora do domínio" aqui é um sinal para revisar, não necessariamente um erro."""
    if not dominios:
        return
    entidade = sheet_name.strip().lower()
    colunas_com_regra = [
        (i, (entidade, str(h).strip().upper()))
        for i, h in enumerate(headers)
        if str(h).strip().lower() not in _COLUNAS_CONTROLE
        and (entidade, str(h).strip().upper()) in dominios
    ]
    if not colunas_com_regra:
        return
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    for i, row in enumerate(linhas):
        if not _is_ponto_ativo(row, col_gera):
            continue
        linha_planilha = i + 2
        for col, chave in colunas_com_regra:
            valor = str(row[col]).strip() if len(row) > col else ""
            if valor and valor not in dominios[chave]:
                analise.add(SEV_AVISO, sheet_name, linha_planilha, chave[1], valor,
                            "Valor fora do dominio conhecido (%s)" % "/".join(sorted(dominios[chave])))


def _carregar_regras_refs(doc):
    """Lê regras ATIVAS da aba VerificacaoRefs; cria a aba com exemplos se faltar."""
    if not _aba_existe_ci(doc, NOME_ABA_VERIFICACAO_REFS):
        _criar_aba_refs_exemplo(doc)
        return []  # exemplos nascem inativos -> nenhuma regra ativa na 1a execução
    sheet = _get_sheet(doc, NOME_ABA_VERIFICACAO_REFS)
    headers, linhas = _ler_entidade(sheet)
    if not linhas:
        return []
    regras = []
    for row in linhas:
        if len(row) < 5:
            continue
        ent_o, attr_o = str(row[0]).strip(), str(row[1]).strip()
        ent_d = _parse_entidades_destino(row[2])
        attr_d = str(row[3]).strip()
        ativa = str(row[4]).strip().lower()
        if ent_o and attr_o and ent_d and attr_d and ativa in _VALORES_ATIVO:
            regras.append((ent_o, attr_o, ent_d, attr_d))
    return regras


def _criar_aba_refs_exemplo(doc):
    """Cria a aba de config com cabeçalho + regras de exemplo inativas."""
    new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
    doc.getSheets().insertByName(NOME_ABA_VERIFICACAO_REFS, new_sheet)
    sheet = _get_sheet(doc, NOME_ABA_VERIFICACAO_REFS)
    matriz = [CABECALHOS_REFS]
    for ent_o, attr_o, ent_d, attr_d in REGRAS_REFS_PADRAO:
        matriz.append([ent_o, attr_o, ent_d, attr_d, "N"])  # N = inativa
    _escrever_matriz(sheet, matriz, negrito_cabecalho=True)


def _buscar_entidade(entidades, nome):
    """Lookup case-insensitive no mapa {nome: (headers, linhas)} -> valor ou None."""
    if nome in entidades:
        return entidades[nome]
    alvo = str(nome).strip().lower()
    for k, v in entidades.items():
        if k.strip().lower() == alvo:
            return v
    return None


def _check_integridade_referencial(entidades, regras, analise):
    """Para cada regra ativa, verifica se o valor de origem existe em algum dos destinos.

    'entidades' é o mapa {nome: (headers, linhas)} já lido das abas — assim esta
    lógica é pura (não depende do LibreOffice) e pode ser exercitada pelo testador
    standalone (completa/testar_verificacao.py). 'ents_d' é sempre uma tupla de 1+
    entidades (ver _parse_entidades_destino) -- suporta FK ambígua do SAGE (ex.:
    PDF.PNT pode ser um PDS ou um PDD; a validação usa a UNIÃO dos IDs de ambos)."""
    cache_destino = {}  # (ents_d ordenado+lower, attr_d.lower) -> set de valores (ou None)
    for ent_o, attr_o, ents_d, attr_d in regras:
        chave = (tuple(sorted(e.lower() for e in ents_d)), attr_d.lower())
        if chave not in cache_destino:
            cache_destino[chave] = _carregar_ids_destino(entidades, ents_d, attr_d)
        destino_ids = cache_destino[chave]
        destino_desc = "/".join(ents_d)
        if destino_ids is None:
            analise.add(SEV_AVISO, ent_o, "-", attr_o, "",
                        "Regra ignorada: destino %s.%s nao encontrado" % (destino_desc, attr_d))
            continue
        origem = _buscar_entidade(entidades, ent_o)
        if origem is None:
            analise.add(SEV_AVISO, ent_o, "-", attr_o, "",
                        "Regra ignorada: entidade de origem nao existe")
            continue
        headers, linhas = origem
        col_o = _idx_coluna(headers, attr_o)
        col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
        if col_o < 0:
            analise.add(SEV_AVISO, ent_o, "-", attr_o, "",
                        "Regra ignorada: atributo nao existe na origem")
            continue
        for i, row in enumerate(linhas):
            if not _is_ponto_ativo(row, col_gera):
                continue
            valor = str(row[col_o]).strip() if len(row) > col_o else ""
            if valor and valor not in destino_ids:
                analise.add(SEV_ERRO, ent_o, i + 2, attr_o, valor,
                            "Referencia nao encontrada em %s.%s" % (destino_desc, attr_d))


def _carregar_ids_destino(entidades, ents_d, attr_d):
    """União dos valores do atributo destino em qualquer uma das entidades de 'ents_d'.

    Uma entidade individual que não existe (ou não tem o atributo) simplesmente não
    contribui valores -- só retorna None se NENHUMA das entidades listadas resolver
    (mesma semântica do valida_fk.py do SkillSAGE para FKs ambíguas)."""
    valores = set()
    encontrou_alguma = False
    for ent_d in ents_d:
        destino = _buscar_entidade(entidades, ent_d)
        if destino is None:
            continue
        headers, linhas = destino
        col = _idx_coluna(headers, attr_d)
        if col < 0:
            continue
        encontrou_alguma = True
        valores.update(str(r[col]).strip() for r in linhas if len(r) > col and str(r[col]).strip())
    return valores if encontrou_alguma else None


# --- Escrita do relatório ---

def _escrever_matriz(sheet, matriz, negrito_cabecalho=False):
    """Escreve uma matriz (lista de listas) a partir de A1, com largura ótima."""
    nrows = len(matriz)
    ncols = max(len(r) for r in matriz)
    normalizada = tuple(
        tuple(str(c) for c in row) + ("",) * (ncols - len(row)) for row in matriz
    )
    sheet.getCellRangeByPosition(0, 0, ncols - 1, nrows - 1).setDataArray(normalizada)
    if negrito_cabecalho:
        cab = sheet.getCellRangeByPosition(0, 0, ncols - 1, 0)
        cab.CharWeight = _FONT_BOLD
        cab.HoriJustify = 2  # CENTER
    columns = sheet.getColumns()
    for i in range(ncols):
        columns.getByIndex(i).OptimalWidth = True


def _escrever_relatorio_analise(doc, analise):
    """Cria/limpa a aba 'Análise' e escreve os achados (erros primeiro)."""
    if _aba_existe_ci(doc, NOME_ABA_ANALISE):
        sheet = _get_sheet(doc, NOME_ABA_ANALISE)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        sheet.getCellRangeByPosition(0, 0, addr.EndColumn, addr.EndRow).clearContents(FLAGS_LIMPAR_TUDO)
    else:
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(NOME_ABA_ANALISE, new_sheet)
        sheet = _get_sheet(doc, NOME_ABA_ANALISE)

    achados = sorted(analise.achados, key=lambda a: _ORDEM_SEV.get(a["sev"], 9))
    matriz = [CABECALHOS_ANALISE]
    for a in achados:
        matriz.append([a["sev"], a["entidade"], a["linha"], a["atributo"], a["valor"], a["descr"]])
    if len(matriz) == 1:
        matriz.append([SEV_OK, "-", "-", "-", "-", "Nenhum problema encontrado."])
    _escrever_matriz(sheet, matriz, negrito_cabecalho=True)

    # Sinal visual rápido pela cor da aba: vermelho se há erro, verde se limpo.
    try:
        sheet.TabColor = 0xCC0000 if analise.erros else 0x2E7D32
    except Exception:
        pass
    try:
        doc.getCurrentController().setActiveSheet(sheet)
    except Exception:
        pass


def _abas_nao_entidade():
    """Conjunto (lower) das abas que não são entidades: config + relatório."""
    ignoradas = set(n.lower() for n in FOLHAS_IGNORADAS)
    ignoradas.update({NOME_ABA_ANALISE.lower(), NOME_ABA_VERIFICACAO_REFS.lower()})
    return ignoradas


def _coletar_entidades(doc):
    """Lê todas as abas de entidade -> {nome: (headers, linhas)}.

    É entidade a aba que não é de config/relatório e tem a coluna de controle 'Gera'.
    Centralizar aqui evita reler cada aba várias vezes e dá um mapa em memória que as
    checagens (puras) consomem."""
    ignoradas = _abas_nao_entidade()
    entidades = {}
    sheets = doc.getSheets()
    for i in range(sheets.getCount()):
        sheet = sheets.getByIndex(i)
        nome = sheet.getName()
        if nome.lower() in ignoradas:
            continue
        headers, linhas = _ler_entidade(sheet)
        if headers is None:
            continue
        if _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE) < 0:
            continue
        entidades[nome] = (headers, linhas)
    return entidades


def _rodar_checagens(entidades, regras, dominios=None):
    """Executa todas as checagens (lógica PURA) e devolve a _Analise preenchida.

    Compartilhado entre a macro (verificar_base) e o testador standalone."""
    analise = _Analise()
    dominios = dominios or {}
    for nome, (headers, linhas) in entidades.items():
        _check_ids(nome, headers, linhas, analise)
        _check_tamanho_id(nome, headers, linhas, analise)
        _check_dominios(nome, headers, linhas, dominios, analise)
    _check_integridade_referencial(entidades, regras, analise)
    return analise


def verificar_base(*args):
    """Macro: roda o linter de integridade e escreve o relatório na aba 'Análise'."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    entidades = _coletar_entidades(doc)
    # Integridade referencial — dirigida pela aba de config (criada se faltar).
    regras = _carregar_regras_refs(doc)
    # Domínios de valores válidos — lidos da aba já existente 'EntidadeAtributoValor'.
    dominios = _carregar_dominios(doc)
    analise = _rodar_checagens(entidades, regras, dominios)
    _escrever_relatorio_analise(doc, analise)


# ===============================================================
# ========= TRILHA COMPLETA: UNIFICAÇÃO DE PONTOS ================
# ===============================================================
# Gera as entidades .dat relacionadas (PDF/PDS/PDD, PAF/PAS/PAD, CGF/CGS) a partir de
# uma definição única de ponto físico, dirigida por abas de config. Cada ponto físico
# é declarado uma vez (uma linha) em PontoDigital/PontoAnalogico; a mesma linha define
# se ele tem comando associado (Comando=S -> gera CGF/CGS com o MESMO ID do PDS/PAS,
# regra fixa do SAGE). Comandos SEM ponto de status próprio (ex.: um "COM_SAGE"
# genérico ligado a um TAC local, usado por vários comandos ao mesmo tempo) entram por
# ComandoAvulso. Distribuição (PDD/PAD) é opt-in por ponto via DistribuicaoPontos,
# usando o Método (Prefixo/Sufixo/Substituir) do canal em CanaisDistribuicao.
#
# Redundância de origem física: várias linhas de PontoDigital/PontoAnalogico podem
# repetir o mesmo ID_Logico (uma por origem). 1 origem -> PDF/PAF direto (TPFIL=NLFL);
# 2+ origens -> um PDF/PAF por origem + RFC em cadeia (fan-in "ou válido") + PDS/PAS
# com TPFIL=FIL5. Não assume nenhuma convenção fixa de IED físico/virtual (isso é
# escopo do futuro Assistente de Protocolo/IED) -- só o número de origens declaradas.
#
# Escrita idempotente: casa por ID nas abas de entidade já existentes (upsert, não
# reescreve a aba do zero) -- regenerar depois de ajustar a config não duplica linhas
# nem apaga colunas/linhas que não vêm daqui (ex.: pontos importados de .dat reais).

NOME_ABA_PONTO_DIGITAL = "PontoDigital"
NOME_ABA_PONTO_ANALOGICO = "PontoAnalogico"
NOME_ABA_COMANDO_AVULSO = "ComandoAvulso"
NOME_ABA_CANAIS_DISTRIBUICAO = "CanaisDistribuicao"
NOME_ABA_DISTRIBUICAO_PONTOS = "DistribuicaoPontos"

# Marca em "Origem" as linhas que este gerador escreveu/atualizou (diferencia de
# linhas vindas de importação real de .dat, que trazem o caminho do arquivo/include).
ORIGEM_GERADO = "UnificacaoPontos"

CABECALHOS_PONTO_DIGITAL = ["ID_Logico", "ID_Fisico", "NOME", "NV2", "KCONV", "TAC", "OCR",
                            "Comando", "ID_Fisico_Comando", "KCONV_Comando", "Gera"]
CABECALHOS_PONTO_ANALOGICO = ["ID_Logico", "ID_Fisico", "NOME", "NV2", "KCONV1", "KCONV2",
                              "KCONV3", "TAC", "OCR", "Gera"]
CABECALHOS_COMANDO_AVULSO = ["ID", "ID_Fisico", "NOME", "NV2", "KCONV", "TAC", "PAC", "PINT",
                             "TIPOE", "TPCTL", "Gera"]
CABECALHOS_CANAIS_DISTRIBUICAO = ["Nome", "TDD", "Metodo", "Valor1", "Valor2", "Ativo"]
CABECALHOS_DISTRIBUICAO_PONTOS = ["ID_Logico", "Canal", "Ativo"]

# Métodos aceitos na coluna "Metodo" de CanaisDistribuicao (case-insensitive).
_METODO_PREFIXO = ("prefixo", "prefix")
_METODO_SUFIXO = ("sufixo", "suffix")
_METODO_SUBSTITUIR = ("substituir", "replace")


def _valor(row, headers, nome_coluna, default=""):
    """Valor de uma coluna pelo nome (case-insensitive), com default se ausente/vazio."""
    col = _idx_coluna(headers, nome_coluna)
    if col < 0 or col >= len(row):
        return default
    v = str(row[col]).strip()
    return v if v else default


def _garantir_aba_config(doc, nome_aba, cabecalhos):
    """Cria a aba de config só com o cabeçalho se ela ainda não existir. Nunca
    sobrescreve uma aba já existente (mesma cautela de _criar_aba_refs_exemplo)."""
    if _aba_existe_ci(doc, nome_aba):
        return
    new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
    doc.getSheets().insertByName(nome_aba, new_sheet)
    sheet = _get_sheet(doc, nome_aba)
    _escrever_matriz(sheet, [cabecalhos], negrito_cabecalho=True)


def _agrupar_por_id_logico(linhas, headers):
    """{ID_Logico: [row, ...]} só com linhas ativas (Gera=x), na ordem em que aparecem."""
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    grupos = {}
    for row in linhas:
        if not _is_ponto_ativo(row, col_gera):
            continue
        id_logico = _valor(row, headers, "ID_Logico")
        if id_logico:
            grupos.setdefault(id_logico, []).append(row)
    return grupos


def _aplicar_metodo(id_original, metodo, valor1, valor2):
    """Transforma um ID conforme o Método do canal de distribuição. Método
    desconhecido/vazio não transforma (defensivo -- nunca gera ID vazio)."""
    m = str(metodo).strip().lower()
    if m in _METODO_PREFIXO:
        return "%s%s" % (valor1, id_original)
    if m in _METODO_SUFIXO:
        return "%s%s" % (id_original, valor1)
    if m in _METODO_SUBSTITUIR:
        return id_original.replace(valor1, valor2)
    return id_original


def _carregar_canais_distribuicao(doc):
    """{nome_canal.lower(): {"TDD":..., "Metodo":..., "Valor1":..., "Valor2":...}} só
    dos canais Ativos. Não cria a aba (ver unificar_pontos, que garante via
    _garantir_aba_config antes de chamar esta função)."""
    canais = {}
    if not _aba_existe_ci(doc, NOME_ABA_CANAIS_DISTRIBUICAO):
        return canais
    sheet = _get_sheet(doc, NOME_ABA_CANAIS_DISTRIBUICAO)
    headers, linhas = _ler_entidade(sheet)
    if not linhas:
        return canais
    for row in linhas:
        nome = _valor(row, headers, "Nome")
        ativo = _valor(row, headers, "Ativo").lower()
        if not nome or ativo not in _VALORES_ATIVO:
            continue
        canais[nome.lower()] = {
            "TDD": _valor(row, headers, "TDD"),
            "Metodo": _valor(row, headers, "Metodo"),
            "Valor1": _valor(row, headers, "Valor1"),
            "Valor2": _valor(row, headers, "Valor2"),
        }
    return canais


def _carregar_distribuicoes_por_ponto(doc):
    """{ID_Logico: [nome_canal, ...]} só das linhas Ativas."""
    dist = {}
    if not _aba_existe_ci(doc, NOME_ABA_DISTRIBUICAO_PONTOS):
        return dist
    sheet = _get_sheet(doc, NOME_ABA_DISTRIBUICAO_PONTOS)
    headers, linhas = _ler_entidade(sheet)
    if not linhas:
        return dist
    for row in linhas:
        id_logico = _valor(row, headers, "ID_Logico")
        canal = _valor(row, headers, "Canal")
        ativo = _valor(row, headers, "Ativo").lower()
        if id_logico and canal and ativo in _VALORES_ATIVO:
            dist.setdefault(id_logico, []).append(canal)
    return dist


def _gerar_distribuicao(id_logico, entidade_pnt, canais, distribuicoes):
    """Lista de linhas PDD/PAD (uma por canal ativo do ponto). 'entidade_pnt' é "PDS"
    ou "PAS" -- o atributo de FK que aponta de volta ao ponto lógico chama-se igual
    (PDD.PDS / PAD.PAS)."""
    linhas = []
    for canal_nome in distribuicoes.get(id_logico, []):
        canal = canais.get(canal_nome.strip().lower())
        if canal is None:
            continue  # canal inativo ou inexistente -- ignora silenciosamente (opt-in)
        novo_id = _aplicar_metodo(id_logico, canal["Metodo"], canal["Valor1"], canal["Valor2"])
        linhas.append({"ID": novo_id, entidade_pnt: id_logico, "TDD": canal["TDD"], "ORDEM": "1"})
    return linhas


def _gerar_fan_out_digital(linhas, headers, canais, distribuicoes):
    """Lógica PURA: {entidade: [linha_dict, ...]} a upsertar, a partir da aba PontoDigital
    já lida em memória. Testável fora do LibreOffice (mesmo espírito do verificador)."""
    saida = {"pdf": [], "pds": [], "pdd": [], "rfc": [], "cgf": [], "cgs": []}
    for id_logico, origens in _agrupar_por_id_logico(linhas, headers).items():
        # Só linhas com ID_Fisico preenchido viram PDF/RFC -- uma origem sem ID_Fisico
        # (ex.: ponto calculado extraído sem PDF correspondente) não tem componente
        # físico nenhum; o PDS ainda é gerado normalmente logo abaixo.
        origens_com_fisico = [o for o in origens if _valor(o, headers, "ID_Fisico")]
        redundante = len(origens_com_fisico) > 1
        tpfil = "FIL5" if redundante else "NLFL"
        primeira = origens[0]
        for ordem, origem in enumerate(origens_com_fisico, start=1):
            id_fisico = _valor(origem, headers, "ID_Fisico")
            saida["pdf"].append({
                "ID": id_fisico, "NV2": _valor(origem, headers, "NV2"),
                "PNT": id_logico, "TPPNT": "PDS",
                "KCONV": _valor(origem, headers, "KCONV"),
                "DESC1": _valor(origem, headers, "NOME"),
            })
            if redundante:
                saida["rfc"].append({
                    "ORDEM": str(ordem), "PARC": id_fisico, "PNT": id_logico,
                    "TPPARC": "PDF", "TPPNT": "PDS", "TIPOP": "EDC",
                })
        saida["pds"].append({
            "ID": id_logico, "NOME": _valor(primeira, headers, "NOME"),
            "TAC": _valor(primeira, headers, "TAC"), "OCR": _valor(primeira, headers, "OCR"),
            "TPFIL": tpfil,
        })
        comandos = [o for o in origens if _valor(o, headers, "Comando").lower() in _VALORES_ATIVO]
        if comandos:
            saida["cgs"].append({
                "ID": id_logico, "NOME": _valor(primeira, headers, "NOME"),
                "TAC": _valor(primeira, headers, "TAC"), "PAC": id_logico,
            })
            for origem_cmd in comandos:
                saida["cgf"].append({
                    "ID": _valor(origem_cmd, headers, "ID_Fisico_Comando"),
                    "NV2": _valor(origem_cmd, headers, "NV2"), "CGS": id_logico,
                    "KCONV": _valor(origem_cmd, headers, "KCONV_Comando"),
                })
        saida["pdd"].extend(_gerar_distribuicao(id_logico, "PDS", canais, distribuicoes))
    return saida


def _gerar_fan_out_analogico(linhas, headers, canais, distribuicoes):
    """Mesma lógica de _gerar_fan_out_digital, para PAF/PAS/PAD (sem comando)."""
    saida = {"paf": [], "pas": [], "pad": [], "rfc": []}
    for id_logico, origens in _agrupar_por_id_logico(linhas, headers).items():
        # Ver nota equivalente em _gerar_fan_out_digital sobre origens sem ID_Fisico.
        origens_com_fisico = [o for o in origens if _valor(o, headers, "ID_Fisico")]
        redundante = len(origens_com_fisico) > 1
        tpfil = "FIL5" if redundante else "NLFL"
        primeira = origens[0]
        for ordem, origem in enumerate(origens_com_fisico, start=1):
            id_fisico = _valor(origem, headers, "ID_Fisico")
            saida["paf"].append({
                "ID": id_fisico, "NV2": _valor(origem, headers, "NV2"),
                "PNT": id_logico, "TPPNT": "PAS",
                "KCONV1": _valor(origem, headers, "KCONV1"),
                "KCONV2": _valor(origem, headers, "KCONV2"),
                "KCONV3": _valor(origem, headers, "KCONV3"),
                "DESC1": _valor(origem, headers, "NOME"),
            })
            if redundante:
                saida["rfc"].append({
                    "ORDEM": str(ordem), "PARC": id_fisico, "PNT": id_logico,
                    "TPPARC": "PAF", "TPPNT": "PAS", "TIPOP": "VAC",
                })
        saida["pas"].append({
            "ID": id_logico, "NOME": _valor(primeira, headers, "NOME"),
            "TAC": _valor(primeira, headers, "TAC"), "OCR": _valor(primeira, headers, "OCR"),
            "TPFIL": tpfil,
        })
        saida["pad"].extend(_gerar_distribuicao(id_logico, "PAS", canais, distribuicoes))
    return saida


def _gerar_comandos_avulsos(linhas, headers):
    """Comandos sem ponto de status próprio (ex.: genérico ligado a um TAC local).
    Cada linha tem seu próprio ID de CGS/CGF -- várias linhas podem repetir o mesmo
    TAC/PAC (o ponto genérico), é exatamente o caso que este mecanismo cobre."""
    saida = {"cgf": [], "cgs": []}
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    for row in linhas:
        if not _is_ponto_ativo(row, col_gera):
            continue
        id_cgs = _valor(row, headers, "ID")
        if not id_cgs:
            continue
        saida["cgs"].append({
            "ID": id_cgs, "NOME": _valor(row, headers, "NOME"),
            "TAC": _valor(row, headers, "TAC"), "PAC": _valor(row, headers, "PAC"),
            "PINT": _valor(row, headers, "PINT"), "TIPOE": _valor(row, headers, "TIPOE"),
            "TPCTL": _valor(row, headers, "TPCTL"),
        })
        saida["cgf"].append({
            "ID": _valor(row, headers, "ID_Fisico"), "NV2": _valor(row, headers, "NV2"),
            "CGS": id_cgs, "KCONV": _valor(row, headers, "KCONV"),
        })
    return saida


def _mesclar_saidas(*saidas):
    """Combina vários dicts {entidade: [linha, ...]} num só (concatena as listas)."""
    combinado = {}
    for saida in saidas:
        for entidade, linhas in saida.items():
            combinado.setdefault(entidade, []).extend(linhas)
    return combinado


def _mesclar_linhas_upsert(headers, linhas_atuais, linhas_novas, colunas_chave=("ID",)):
    """Lógica PURA do upsert: recebe headers/linhas já lidos de uma aba + as linhas
    novas a aplicar, devolve (headers_finais, linhas_finais) com o upsert feito casando
    por 'colunas_chave' (atualiza a linha existente com aquela chave; cria uma linha
    nova, marcada Gera=x e Origem=UnificacaoPontos quando essas colunas existirem, se
    a chave ainda não existir). 'colunas_chave' aceita mais de uma coluna (chave
    composta -- ex.: ("ID_Logico","Canal") para tabelas de ligação N:N como
    DistribuicaoPontos, que não têm uma coluna "ID" única). Não toca UNO -- é a parte
    testável fora do LibreOffice (mesmo espírito de _rodar_checagens no verificador)."""
    headers = list(headers)
    linhas_atuais = [list(r) for r in linhas_atuais]
    for col in colunas_chave:
        if _idx_coluna(headers, col) < 0:
            headers.append(col)
    for linha in linhas_novas:
        for atributo in linha:
            if _idx_coluna(headers, atributo) < 0:
                headers.append(atributo)
    for r in linhas_atuais:
        if len(r) < len(headers):
            r.extend([""] * (len(headers) - len(r)))

    idx_chave = [_idx_coluna(headers, c) for c in colunas_chave]
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    col_origem = _idx_coluna(headers, CABEÇALHO_COLUNA_ORIGEM)

    def _chave_da_linha(r):
        return tuple(str(r[i]).strip() if i < len(r) else "" for i in idx_chave)

    index_por_chave = {}
    for i, r in enumerate(linhas_atuais):
        chave = _chave_da_linha(r)
        if all(chave):
            index_por_chave[chave] = i

    for linha in linhas_novas:
        chave = tuple(str(linha.get(c, "")).strip() for c in colunas_chave)
        if not all(chave):
            continue
        if chave in index_por_chave:
            r = linhas_atuais[index_por_chave[chave]]
        else:
            r = [""] * len(headers)
            if col_gera >= 0:
                r[col_gera] = CODIGO_BLOCO_ATIVO
            if col_origem >= 0:
                r[col_origem] = ORIGEM_GERADO
            linhas_atuais.append(r)
            index_por_chave[chave] = len(linhas_atuais) - 1
        for atributo, valor in linha.items():
            idx = _idx_coluna(headers, atributo)
            r[idx] = "" if valor is None else str(valor)
    return headers, linhas_atuais


def _upsert_linhas_entidade(doc, sheet_name, linhas_novas, colunas_chave=("ID",)):
    """Insere ou atualiza linhas na aba 'sheet_name', casando por 'colunas_chave' (ver
    _mesclar_linhas_upsert). Cria a aba (com o cabeçalho padrão Origem/Gera/Comentario-
    Include) se faltar; se já existir, preserva todas as colunas e linhas não tocadas
    (nunca reescreve a aba a partir do zero como o importador faz em write_to_sheet) --
    convive com pontos importados de .dat reais na mesma aba."""
    if not linhas_novas:
        return
    if _aba_existe_ci(doc, sheet_name):
        sheet = _get_sheet(doc, sheet_name)
        headers, linhas_lidas = _ler_entidade(sheet)
        headers = headers or [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS]
        linhas_atuais = linhas_lidas or []
    else:
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(sheet_name, new_sheet)
        sheet = _get_sheet(doc, sheet_name)
        headers = [CABEÇALHO_COLUNA_ORIGEM, CABEÇALHO_COLUNA_CONTROLE, CABEÇALHO_COLUNA_DADOS]
        linhas_atuais = []

    headers_final, linhas_final = _mesclar_linhas_upsert(headers, linhas_atuais, linhas_novas, colunas_chave)
    matriz = [headers_final] + [[str(c) for c in r] for r in linhas_final]
    _escrever_matriz(sheet, matriz, negrito_cabecalho=True)


def unificar_pontos(*args):
    """Macro: lê PontoDigital/PontoAnalogico/ComandoAvulso/CanaisDistribuicao/
    DistribuicaoPontos e escreve (upsert) as entidades relacionadas."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    for nome, cabecalhos in (
        (NOME_ABA_PONTO_DIGITAL, CABECALHOS_PONTO_DIGITAL),
        (NOME_ABA_PONTO_ANALOGICO, CABECALHOS_PONTO_ANALOGICO),
        (NOME_ABA_COMANDO_AVULSO, CABECALHOS_COMANDO_AVULSO),
        (NOME_ABA_CANAIS_DISTRIBUICAO, CABECALHOS_CANAIS_DISTRIBUICAO),
        (NOME_ABA_DISTRIBUICAO_PONTOS, CABECALHOS_DISTRIBUICAO_PONTOS),
    ):
        _garantir_aba_config(doc, nome, cabecalhos)

    canais = _carregar_canais_distribuicao(doc)
    distribuicoes = _carregar_distribuicoes_por_ponto(doc)

    headers_dig, linhas_dig = _ler_entidade(_get_sheet(doc, NOME_ABA_PONTO_DIGITAL))
    headers_ana, linhas_ana = _ler_entidade(_get_sheet(doc, NOME_ABA_PONTO_ANALOGICO))
    headers_cmd, linhas_cmd = _ler_entidade(_get_sheet(doc, NOME_ABA_COMANDO_AVULSO))

    saida = _mesclar_saidas(
        _gerar_fan_out_digital(linhas_dig or [], headers_dig or CABECALHOS_PONTO_DIGITAL,
                                canais, distribuicoes),
        _gerar_fan_out_analogico(linhas_ana or [], headers_ana or CABECALHOS_PONTO_ANALOGICO,
                                  canais, distribuicoes),
        _gerar_comandos_avulsos(linhas_cmd or [], headers_cmd or CABECALHOS_COMANDO_AVULSO),
    )
    for entidade, linhas in saida.items():
        # RFC não tem coluna "ID" própria (ver dicionário de atributos do SAGE) -- sem
        # isso, o upsert (chave "ID" por padrão) descartaria toda linha de RFC em
        # silêncio, porque nenhum dict de RFC tem essa chave.
        chave = ("PARC", "PNT") if entidade == "rfc" else ("ID",)
        _upsert_linhas_entidade(doc, entidade.upper(), linhas, colunas_chave=chave)


# ---------------------------------------------------------------
# ---- Extração reversa: entidades existentes -> abas de config ----
# ---------------------------------------------------------------
# Espelho de unificar_pontos: reconstrói PontoDigital/PontoAnalogico/ComandoAvulso/
# CanaisDistribuicao/DistribuicaoPontos a partir de pontos já importados de .dat reais
# (PDS/PDF/PAS/PAF/CGS/CGF/PDD/PAD). Sem isso, o modelo unificado só serviria pra
# pontos novos -- uma base real já existente nunca apareceria em PontoDigital.
#
# É uma reconstrução de melhor esforço, não um inverso perfeito: quando uma
# transformação de distribuição não é reconhecidamente Prefixo/Sufixo (ver
# _inferir_metodo), ou quando há mais canais de comando (CGF) do que origens físicas
# declaradas, alguns campos ficam para o usuário completar/conferir manualmente.

def _linhas_ativas_como_dicts(headers, linhas):
    """[{atributo: valor}, ...] só das linhas ativas (Gera=x). headers/linhas podem
    ser None (entidade ausente) -- devolve [] nesse caso."""
    if not headers:
        return []
    col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
    saida = []
    for row in linhas or []:
        if not _is_ponto_ativo(row, col_gera):
            continue
        saida.append({h: (str(row[i]).strip() if i < len(row) else "") for i, h in enumerate(headers)})
    return saida


def _extrair_ponto_digital(entidades):
    """Lógica PURA: linhas de PontoDigital a partir de PDS/PDF/CGS/CGF já lidos
    (entidades = {"pds": (headers,linhas), ...}, entidades ausentes viram (None,None)).
    Cada PDS ativo vira 1+ linhas (uma por PDF cujo PNT aponta pra ele); PDS sem PDF
    correspondente ainda gera 1 linha (só com os campos do PDS, usando o próprio
    ID_Logico como ID_Fisico provisório) -- preserva o ponto pro usuário completar."""
    headers_pds, linhas_pds = entidades.get("pds", (None, None))
    if not headers_pds:
        return []
    pds_dicts = _linhas_ativas_como_dicts(headers_pds, linhas_pds)
    pdf_dicts = _linhas_ativas_como_dicts(*entidades.get("pdf", (None, None)))
    cgs_dicts = _linhas_ativas_como_dicts(*entidades.get("cgs", (None, None)))
    cgf_dicts = _linhas_ativas_como_dicts(*entidades.get("cgf", (None, None)))

    pdf_por_pnt = {}
    for pdf in pdf_dicts:
        pdf_por_pnt.setdefault(pdf.get("PNT", ""), []).append(pdf)
    cgs_por_id = {cgs["ID"]: cgs for cgs in cgs_dicts if cgs.get("ID")}
    cgf_por_cgs = {}
    for cgf in cgf_dicts:
        cgf_por_cgs.setdefault(cgf.get("CGS", ""), []).append(cgf)

    saida = []
    for pds in pds_dicts:
        id_logico = pds.get("ID", "")
        if not id_logico:
            continue
        origens = pdf_por_pnt.get(id_logico)
        if not origens:
            # Sem PDF correspondente -- tipicamente um ponto calculado (RCA/TCL), que
            # não tem origem física nenhuma. Fora do escopo de PontoDigital (que
            # descreve pontos COM origem física); o PDS em si continua intocado na
            # aba "pds". Fabricar aqui um ID_Fisico só geraria um PDF fantasma na
            # próxima vez que unificar_pontos() rodasse.
            continue
        cgfs = cgf_por_cgs.get(id_logico, []) if id_logico in cgs_por_id else []
        for i, pdf in enumerate(origens):
            linha = {
                "ID_Logico": id_logico,
                "ID_Fisico": pdf.get("ID", ""),
                "NOME": pdf.get("DESC1") or pds.get("NOME", ""),
                "NV2": pdf.get("NV2", ""),
                "KCONV": pdf.get("KCONV", ""),
                "TAC": pds.get("TAC", ""),
                "OCR": pds.get("OCR", ""),
                "Comando": "S" if id_logico in cgs_por_id else "N",
            }
            if id_logico in cgs_por_id:
                # Casa a i-ésima origem com o i-ésimo CGF; sobrando menos CGF que
                # origens, reaproveita o primeiro (caso comum: 1 canal de comando
                # para N origens redundantes -- ver nota em completa/README.md).
                cgf = cgfs[i] if i < len(cgfs) else (cgfs[0] if cgfs else None)
                if cgf:
                    linha["ID_Fisico_Comando"] = cgf.get("ID", "")
                    linha["KCONV_Comando"] = cgf.get("KCONV", "")
            saida.append(linha)
    return saida


def _extrair_ponto_analogico(entidades):
    """Mesma lógica de _extrair_ponto_digital, para PAS/PAF (sem comando)."""
    headers_pas, linhas_pas = entidades.get("pas", (None, None))
    if not headers_pas:
        return []
    pas_dicts = _linhas_ativas_como_dicts(headers_pas, linhas_pas)
    paf_dicts = _linhas_ativas_como_dicts(*entidades.get("paf", (None, None)))
    paf_por_pnt = {}
    for paf in paf_dicts:
        paf_por_pnt.setdefault(paf.get("PNT", ""), []).append(paf)

    saida = []
    for pas in pas_dicts:
        id_logico = pas.get("ID", "")
        if not id_logico:
            continue
        origens = paf_por_pnt.get(id_logico)
        if not origens:
            continue  # sem PAF correspondente -- ver nota equivalente em _extrair_ponto_digital
        for paf in origens:
            saida.append({
                "ID_Logico": id_logico,
                "ID_Fisico": paf.get("ID", ""),
                "NOME": paf.get("DESC1") or pas.get("NOME", ""),
                "NV2": paf.get("NV2", ""),
                "KCONV1": paf.get("KCONV1", ""), "KCONV2": paf.get("KCONV2", ""),
                "KCONV3": paf.get("KCONV3", ""),
                "TAC": pas.get("TAC", ""), "OCR": pas.get("OCR", ""),
            })
    return saida


def _extrair_comandos_avulsos(entidades, ids_logicos_com_ponto):
    """CGS cujo ID não corresponde a nenhum PDS/PAS (comando "solto", ex.: COM_SAGE)."""
    cgs_dicts = _linhas_ativas_como_dicts(*entidades.get("cgs", (None, None)))
    if not cgs_dicts:
        return []
    cgf_dicts = _linhas_ativas_como_dicts(*entidades.get("cgf", (None, None)))
    cgf_por_cgs = {}
    for cgf in cgf_dicts:
        cgf_por_cgs.setdefault(cgf.get("CGS", ""), []).append(cgf)

    saida = []
    for cgs in cgs_dicts:
        id_cgs = cgs.get("ID", "")
        if not id_cgs or id_cgs in ids_logicos_com_ponto:
            continue
        cgf = next(iter(cgf_por_cgs.get(id_cgs, [])), {})
        saida.append({
            "ID": id_cgs, "ID_Fisico": cgf.get("ID", ""),
            "NOME": cgs.get("NOME", ""), "NV2": cgf.get("NV2", ""),
            "KCONV": cgf.get("KCONV", ""), "TAC": cgs.get("TAC", ""),
            "PAC": cgs.get("PAC", ""), "PINT": cgs.get("PINT", ""),
            "TIPOE": cgs.get("TIPOE", ""), "TPCTL": cgs.get("TPCTL", ""),
        })
    return saida


def _inferir_metodo(id_logico, id_transformado):
    """Tenta inferir (Metodo, Valor1) de um par (ID lógico, ID já transformado numa
    distribuição existente). Reconhece Prefixo/Sufixo com confiança (basta conferir se
    um é sufixo/prefixo literal do outro); quando não dá pra inferir, devolve
    ("Substituir", None) -- o usuário confere/completa Valor1/Valor2 manualmente."""
    if not id_transformado or id_transformado == id_logico:
        return None, None
    if id_transformado.endswith(id_logico):
        return "Prefixo", id_transformado[:-len(id_logico)]
    if id_transformado.startswith(id_logico):
        return "Sufixo", id_transformado[len(id_logico):]
    return "Substituir", None


def _extrair_canais_e_distribuicao(entidades):
    """(linhas_canais, linhas_distribuicao) a partir de PDD/PAD já existentes. Um canal
    novo por TDD distinto encontrado (Metodo/Valor1 inferidos do 1º par visto para
    aquele TDD); linhas de DistribuicaoPontos ligam cada ID lógico ao TDD/canal."""
    canais_vistos = {}
    distribuicao = []
    for ent, attr_pnt in (("pdd", "PDS"), ("pad", "PAS")):
        for row in _linhas_ativas_como_dicts(*entidades.get(ent, (None, None))):
            id_logico, tdd = row.get(attr_pnt, ""), row.get("TDD", "")
            if not id_logico or not tdd:
                continue
            if tdd not in canais_vistos:
                metodo, valor1 = _inferir_metodo(id_logico, row.get("ID", ""))
                canais_vistos[tdd] = {"Nome": tdd, "TDD": tdd, "Metodo": metodo or "",
                                      "Valor1": valor1 or "", "Valor2": "", "Ativo": "S"}
            distribuicao.append({"ID_Logico": id_logico, "Canal": tdd, "Ativo": "S"})
    return list(canais_vistos.values()), distribuicao


def extrair_pontos(*args):
    """Macro: espelho de unificar_pontos -- reconstrói PontoDigital/PontoAnalogico/
    ComandoAvulso/CanaisDistribuicao/DistribuicaoPontos a partir das entidades já
    importadas (PDS/PDF/PAS/PAF/CGS/CGF/PDD/PAD). Fecha o ciclo pra bases já
    existentes; rodar antes de editar/estender uma base real pelo modelo unificado."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    entidades = {}
    for nome in ("pds", "pdf", "pas", "paf", "cgs", "cgf", "pdd", "pad"):
        if _aba_existe_ci(doc, nome.upper()):
            entidades[nome] = _ler_entidade(_get_sheet(doc, nome.upper()))

    for nome_aba, cabecalhos in (
        (NOME_ABA_PONTO_DIGITAL, CABECALHOS_PONTO_DIGITAL),
        (NOME_ABA_PONTO_ANALOGICO, CABECALHOS_PONTO_ANALOGICO),
        (NOME_ABA_COMANDO_AVULSO, CABECALHOS_COMANDO_AVULSO),
        (NOME_ABA_CANAIS_DISTRIBUICAO, CABECALHOS_CANAIS_DISTRIBUICAO),
        (NOME_ABA_DISTRIBUICAO_PONTOS, CABECALHOS_DISTRIBUICAO_PONTOS),
    ):
        _garantir_aba_config(doc, nome_aba, cabecalhos)

    linhas_pd = _extrair_ponto_digital(entidades)
    linhas_pa = _extrair_ponto_analogico(entidades)
    ids_com_ponto = {l["ID_Logico"] for l in linhas_pd} | {l["ID_Logico"] for l in linhas_pa}
    linhas_ca = _extrair_comandos_avulsos(entidades, ids_com_ponto)
    linhas_canais, linhas_dist = _extrair_canais_e_distribuicao(entidades)

    _upsert_linhas_entidade(doc, NOME_ABA_PONTO_DIGITAL, linhas_pd, colunas_chave=("ID_Fisico",))
    _upsert_linhas_entidade(doc, NOME_ABA_PONTO_ANALOGICO, linhas_pa, colunas_chave=("ID_Fisico",))
    _upsert_linhas_entidade(doc, NOME_ABA_COMANDO_AVULSO, linhas_ca, colunas_chave=("ID",))
    _upsert_linhas_entidade(doc, NOME_ABA_CANAIS_DISTRIBUICAO, linhas_canais, colunas_chave=("Nome",))
    _upsert_linhas_entidade(doc, NOME_ABA_DISTRIBUICAO_PONTOS, linhas_dist,
                            colunas_chave=("ID_Logico", "Canal"))


# ===============================================================
# ============= TRILHA COMPLETA: GANHOS RÁPIDOS ==================
# ===============================================================
# Três recursos de baixo esforço/alto retorno do PLANEJAMENTO.md: troca de ID global
# (com propagação de referências), estatística de pontos por entidade, e gestão de
# includes (listar + corrigir em lote). Cada um é read/write direto nas abas de
# entidade já existentes -- nenhum toca em arquivo .dat.

def _valor_bruto(row, headers, nome_coluna):
    """Como _valor, mas sem default -- devolve '' se ausente (uso interno aqui onde
    'row' pode ser lista OU tupla, ambos suportados por _idx_coluna/indexação)."""
    col = _idx_coluna(headers, nome_coluna)
    if col < 0 or col >= len(row):
        return ""
    return str(row[col]).strip()


# --- Troca de ID global ---------------------------------------------------

NOME_ABA_TROCA_ID = "TrocaId"
NOME_ABA_RELATORIO_TROCA_ID = "RelatorioTrocaId"
CABECALHOS_TROCA_ID = ["IDAntigo", "IDNovo", "Ativa"]


def _construir_mapa_referencias_por_destino():
    """{ENTIDADE_DESTINO: [(entidade_origem, atributo_origem), ...]} a partir de
    REGRAS_REFS_PADRAO -- quem referencia essa entidade e por qual atributo. É o
    inverso do grafo de FK usado pelo verificador; aqui serve pra saber onde propagar
    uma troca de ID (ex.: renomear um PDS precisa achar todo PDD.PDS/RCA.PARC/... que
    aponta pra ele)."""
    mapa = {}
    for ent_o, attr_o, ent_d_txt, _attr_d in REGRAS_REFS_PADRAO:
        for ent_d in _parse_entidades_destino(ent_d_txt):
            mapa.setdefault(ent_d.upper(), []).append((ent_o, attr_o))
    return mapa


def _preparar_entidades_mutaveis(entidades_brutas):
    """{nome: (headers, [linha_lista, ...])} -- converte as linhas (tuplas imutáveis
    vindas de _ler_entidade/getDataArray) em listas, pra poder alterar célula a célula."""
    return {nome: (headers, [list(r) for r in linhas]) for nome, (headers, linhas) in entidades_brutas.items()}


def _resolver_nome_entidade(entidades, nome):
    """Nome EXATO da chave em 'entidades' (preserva o case real do nome da aba) que
    corresponde a 'nome' (case-insensitive); None se não existir. Necessário porque
    mapa_referencias/REGRAS_REFS_PADRAO guardam nomes em MAIÚSCULAS, mas as abas reais
    costumam estar em minúsculas (ex.: "pds") -- sem isso, 'tocadas' ficaria com o case
    do mapa em vez do case real da aba, e o write-back (entidades[nome]) quebraria."""
    if nome in entidades:
        return nome
    alvo = str(nome).strip().lower()
    for k in entidades:
        if k.strip().lower() == alvo:
            return k
    return None


def _entidade_do_id(entidades, id_valor):
    """Nome da entidade onde 'id_valor' aparece na coluna ID (None se não achar, ou se
    achar em mais de uma -- ambíguo demais pra trocar sem confirmação do usuário)."""
    achadas = []
    for nome, (headers, linhas) in entidades.items():
        col_id = _idx_coluna(headers, "ID")
        if col_id < 0:
            continue
        if any(_valor_bruto(row, headers, "ID") == id_valor for row in linhas):
            achadas.append(nome)
    return achadas[0] if len(achadas) == 1 else None


def _trocar_id_em_entidades(entidades, id_antigo, id_novo, mapa_referencias):
    """Lógica PURA: troca id_antigo->id_novo na entidade onde ele é a própria chave
    (coluna ID) e em toda coluna que a referencia (mapa_referencias). Muta as linhas de
    'entidades' (listas, ver _preparar_entidades_mutaveis) IN-PLACE. Devolve
    (entidades_tocadas: set[str], relatorio: list[str])."""
    tocadas = set()
    relatorio = []
    entidade_origem = _entidade_do_id(entidades, id_antigo)
    if entidade_origem is None:
        relatorio.append("ID '%s' não encontrado (ou ambíguo em mais de uma entidade) -- nada alterado" % id_antigo)
        return tocadas, relatorio

    headers, linhas = entidades[entidade_origem]
    col_id = _idx_coluna(headers, "ID")
    for row in linhas:
        if len(row) > col_id and str(row[col_id]).strip() == id_antigo:
            row[col_id] = id_novo
            tocadas.add(entidade_origem)
    relatorio.append("%s.ID: '%s' -> '%s'" % (entidade_origem, id_antigo, id_novo))

    for ent_ref, attr_ref in mapa_referencias.get(entidade_origem.upper(), []):
        nome_real = _resolver_nome_entidade(entidades, ent_ref)
        if nome_real is None:
            continue
        headers_ref, linhas_ref = entidades[nome_real]
        col_ref = _idx_coluna(headers_ref, attr_ref)
        if col_ref < 0:
            continue
        n = 0
        for row in linhas_ref:
            if len(row) > col_ref and str(row[col_ref]).strip() == id_antigo:
                row[col_ref] = id_novo
                n += 1
        if n:
            tocadas.add(nome_real)
            relatorio.append("%s.%s: %d referência(s) atualizada(s)" % (nome_real, attr_ref, n))
    return tocadas, relatorio


def _escrever_relatorio_simples(doc, nome_aba, cabecalho, linhas_texto):
    """Cria/limpa 'nome_aba' e escreve um relatório de 1 coluna (cabecalho + linhas)."""
    if _aba_existe_ci(doc, nome_aba):
        sheet = _get_sheet(doc, nome_aba)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        sheet.getCellRangeByPosition(0, 0, addr.EndColumn, addr.EndRow).clearContents(FLAGS_LIMPAR_TUDO)
    else:
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(nome_aba, new_sheet)
        sheet = _get_sheet(doc, nome_aba)
    matriz = [cabecalho] + [[linha] for linha in (linhas_texto or ["Nenhuma alteração ativa a processar."])]
    _escrever_matriz(sheet, matriz, negrito_cabecalho=True)


def trocar_id_global(*args):
    """Macro: lê a aba 'TrocaId' (criada vazia na 1ª execução) e, para cada linha
    ativa (IDAntigo, IDNovo), renomeia o ID naquela entidade e propaga a troca por
    toda coluna que referencia essa entidade (grafo de REGRAS_REFS_PADRAO). Suporta
    várias trocas em lote (linhas processadas em ordem -- uma troca pode encadear na
    seguinte). Escreve o relatório em 'RelatorioTrocaId'."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    _garantir_aba_config(doc, NOME_ABA_TROCA_ID, CABECALHOS_TROCA_ID)

    entidades = _preparar_entidades_mutaveis(_coletar_entidades(doc))
    mapa_referencias = _construir_mapa_referencias_por_destino()

    headers_troca, linhas_troca = _ler_entidade(_get_sheet(doc, NOME_ABA_TROCA_ID))
    relatorio_geral = []
    entidades_tocadas = set()
    if headers_troca:
        col_ativa = _idx_coluna(headers_troca, "Ativa")
        for row in (linhas_troca or []):
            ativa = str(row[col_ativa]).strip().lower() if col_ativa >= 0 and len(row) > col_ativa else ""
            if ativa not in _VALORES_ATIVO:
                continue
            id_antigo = _valor_bruto(row, headers_troca, "IDAntigo")
            id_novo = _valor_bruto(row, headers_troca, "IDNovo")
            if not id_antigo or not id_novo:
                continue
            tocadas, relatorio = _trocar_id_em_entidades(entidades, id_antigo, id_novo, mapa_referencias)
            entidades_tocadas.update(tocadas)
            relatorio_geral.extend(relatorio)

    for nome in entidades_tocadas:
        headers, linhas = entidades[nome]
        matriz = [headers] + [[str(c) for c in r] for r in linhas]
        _escrever_matriz(_get_sheet(doc, nome.upper()), matriz, negrito_cabecalho=True)
    _escrever_relatorio_simples(doc, NOME_ABA_RELATORIO_TROCA_ID, ["Alteração"], relatorio_geral)


# --- Estatística -----------------------------------------------------------

NOME_ABA_ESTATISTICA = "Estatistica"


def _calcular_estatisticas(entidades):
    """[(nome_entidade, total_linhas, total_ativas), ...] ordenado por nome. Lógica
    PURA -- 'entidades' é o mesmo mapa {nome:(headers,linhas)} do verificador."""
    stats = []
    for nome, (headers, linhas) in entidades.items():
        col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
        ativas = sum(1 for row in linhas if _is_ponto_ativo(row, col_gera))
        stats.append((nome, len(linhas), ativas))
    return sorted(stats, key=lambda t: t[0].lower())


def estatistica_base(*args):
    """Macro: conta linhas (total e ativas, Gera=x) por entidade e escreve na aba
    'Estatistica', com uma linha TOTAL no fim."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    stats = _calcular_estatisticas(_coletar_entidades(doc))
    if _aba_existe_ci(doc, NOME_ABA_ESTATISTICA):
        sheet = _get_sheet(doc, NOME_ABA_ESTATISTICA)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        addr = cursor.getRangeAddress()
        sheet.getCellRangeByPosition(0, 0, addr.EndColumn, addr.EndRow).clearContents(FLAGS_LIMPAR_TUDO)
    else:
        new_sheet = doc.createInstance("com.sun.star.sheet.Spreadsheet")
        doc.getSheets().insertByName(NOME_ABA_ESTATISTICA, new_sheet)
        sheet = _get_sheet(doc, NOME_ABA_ESTATISTICA)
    matriz = [["Entidade", "Total de linhas", "Linhas ativas (Gera=x)"]]
    for nome, total, ativas in stats:
        matriz.append([nome.upper(), str(total), str(ativas)])
    matriz.append(["TOTAL", str(sum(t[1] for t in stats)), str(sum(t[2] for t in stats))])
    _escrever_matriz(sheet, matriz, negrito_cabecalho=True)


# --- Gestão de includes -----------------------------------------------------

NOME_ABA_SUBSTITUIR_INCLUDES = "SubstituirIncludes"
NOME_ABA_RELATORIO_INCLUDES = "RelatorioIncludes"
CABECALHOS_SUBSTITUIR_INCLUDES = ["Buscar", "Substituir", "Ativa"]

_TIPOS_INCLUDE = (CODIGO_INCLUDE, CODIGO_INCLUDE_COMENTADO)


def _substituir_em_includes(entidades, buscar, substituir):
    """Substitui a substring 'buscar' por 'substituir' no path (coluna
    Comentario/Include) de toda linha de include (Gera=i ou u), em qualquer entidade.
    Muta 'entidades' IN-PLACE. Devolve (entidades_tocadas, quantidade_substituida)."""
    tocadas = set()
    n = 0
    for nome, (headers, linhas) in entidades.items():
        col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
        col_dados = _idx_coluna(headers, CABEÇALHO_COLUNA_DADOS)
        if col_gera < 0 or col_dados < 0:
            continue
        for row in linhas:
            codigo = str(row[col_gera]).strip().lower() if len(row) > col_gera else ""
            if codigo not in _TIPOS_INCLUDE:
                continue
            atual = str(row[col_dados]) if len(row) > col_dados else ""
            if buscar and buscar in atual:
                row[col_dados] = atual.replace(buscar, substituir)
                tocadas.add(nome)
                n += 1
    return tocadas, n


def _listar_includes(entidades):
    """[(entidade, linha_planilha, path)] de toda linha de include (Gera=i ou u), em
    qualquer entidade. Lógica PURA."""
    saida = []
    for nome, (headers, linhas) in entidades.items():
        col_gera = _idx_coluna(headers, CABEÇALHO_COLUNA_CONTROLE)
        col_dados = _idx_coluna(headers, CABEÇALHO_COLUNA_DADOS)
        if col_gera < 0 or col_dados < 0:
            continue
        for i, row in enumerate(linhas):
            codigo = str(row[col_gera]).strip().lower() if len(row) > col_gera else ""
            if codigo in _TIPOS_INCLUDE:
                saida.append((nome, i + 2, str(row[col_dados]).strip() if len(row) > col_dados else ""))
    return saida


def gerir_includes(*args):
    """Macro: aplica as substituições ativas da aba 'SubstituirIncludes' (criada vazia
    na 1ª execução) em toda linha de include de qualquer entidade, e sempre escreve o
    estado atual (path de cada include, já com as trocas aplicadas) em
    'RelatorioIncludes' -- serve tanto pra listar quanto pra corrigir em lote."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    _garantir_aba_config(doc, NOME_ABA_SUBSTITUIR_INCLUDES, CABECALHOS_SUBSTITUIR_INCLUDES)

    entidades = _preparar_entidades_mutaveis(_coletar_entidades(doc))
    headers_sub, linhas_sub = _ler_entidade(_get_sheet(doc, NOME_ABA_SUBSTITUIR_INCLUDES))
    entidades_tocadas = set()
    if headers_sub:
        col_ativa = _idx_coluna(headers_sub, "Ativa")
        for row in (linhas_sub or []):
            ativa = str(row[col_ativa]).strip().lower() if col_ativa >= 0 and len(row) > col_ativa else ""
            if ativa not in _VALORES_ATIVO:
                continue
            buscar = _valor_bruto(row, headers_sub, "Buscar")
            substituir = _valor_bruto(row, headers_sub, "Substituir")
            if not buscar:
                continue
            tocadas, _n = _substituir_em_includes(entidades, buscar, substituir)
            entidades_tocadas.update(tocadas)

    for nome in entidades_tocadas:
        headers, linhas = entidades[nome]
        matriz = [headers] + [[str(c) for c in r] for r in linhas]
        _escrever_matriz(_get_sheet(doc, nome.upper()), matriz, negrito_cabecalho=True)

    linhas_relatorio = ["%s (linha %d): %s" % (ent, linha, path)
                        for ent, linha, path in _listar_includes(entidades)]
    _escrever_relatorio_simples(doc, NOME_ABA_RELATORIO_INCLUDES, ["Include"], linhas_relatorio)


# ===============================================================
# ================= EXPOSIÇÃO PARA LIBREOFFICE ==================
# ===============================================================
g_exportedScripts = (importar_dats, exportar_dats, importar_parcial, exportar_parcial,
                     atualizar_amostras_cores, verificar_base, unificar_pontos, extrair_pontos,
                     trocar_id_global, estatistica_base, gerir_includes)
