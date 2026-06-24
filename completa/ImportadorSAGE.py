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

# Regras de exemplo gravadas na aba de config na 1ª execução (inativas por padrão,
# para não gerar falso-positivo antes de o usuário confirmar que se aplicam à base).
# Formato: (EntidadeOrigem, AtributoOrigem, EntidadeDestino, AtributoDestino)
REGRAS_REFS_EXEMPLO = [
    ("PDS", "TAC", "TAC", "ID"),
    ("PDD", "PDS", "PDS", "ID"),
    ("PAS", "TAC", "TAC", "ID"),
]

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
        ent_o, attr_o, ent_d, attr_d = (str(row[k]).strip() for k in range(4))
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
    for ent_o, attr_o, ent_d, attr_d in REGRAS_REFS_EXEMPLO:
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
    """Para cada regra ativa, verifica se o valor de origem existe no destino.

    'entidades' é o mapa {nome: (headers, linhas)} já lido das abas — assim esta
    lógica é pura (não depende do LibreOffice) e pode ser exercitada pelo testador
    standalone (completa/testar_verificacao.py)."""
    cache_destino = {}  # (ent_d.lower, attr_d.lower) -> set de valores (ou None)
    for ent_o, attr_o, ent_d, attr_d in regras:
        chave = (ent_d.lower(), attr_d.lower())
        if chave not in cache_destino:
            cache_destino[chave] = _carregar_ids_destino(entidades, ent_d, attr_d)
        destino_ids = cache_destino[chave]
        if destino_ids is None:
            analise.add(SEV_AVISO, ent_o, "-", attr_o, "",
                        "Regra ignorada: destino %s.%s nao encontrado" % (ent_d, attr_d))
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
                            "Referencia nao encontrada em %s.%s" % (ent_d, attr_d))


def _carregar_ids_destino(entidades, ent_d, attr_d):
    """Conjunto de valores do atributo destino (qualquer linha); None se inválido."""
    destino = _buscar_entidade(entidades, ent_d)
    if destino is None:
        return None
    headers, linhas = destino
    col = _idx_coluna(headers, attr_d)
    if col < 0:
        return None
    return {str(r[col]).strip() for r in linhas if len(r) > col and str(r[col]).strip()}


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


def _rodar_checagens(entidades, regras):
    """Executa todas as checagens (lógica PURA) e devolve a _Analise preenchida.

    Compartilhado entre a macro (verificar_base) e o testador standalone."""
    analise = _Analise()
    for nome, (headers, linhas) in entidades.items():
        _check_ids(nome, headers, linhas, analise)
    _check_integridade_referencial(entidades, regras, analise)
    return analise


def verificar_base(*args):
    """Macro: roda o linter de integridade e escreve o relatório na aba 'Análise'."""
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore
    entidades = _coletar_entidades(doc)
    # Integridade referencial — dirigida pela aba de config (criada se faltar).
    regras = _carregar_regras_refs(doc)
    analise = _rodar_checagens(entidades, regras)
    _escrever_relatorio_analise(doc, analise)


# ===============================================================
# ================= EXPOSIÇÃO PARA LIBREOFFICE ==================
# ===============================================================
g_exportedScripts = importar_dats, exportar_dats, importar_parcial, exportar_parcial, atualizar_amostras_cores, verificar_base
