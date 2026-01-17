"""
Módulo para leitura de arquivos de planilha (Excel/CSV)
Substitui a integração com Google Sheets quando arquivos são enviados
"""
import os
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
import pandas as pd

logger = logging.getLogger(__name__)


def read_spreadsheet_file(
    file_path: str,
    sheet_name: Optional[str] = None,
    header: Optional[int] = None
) -> Dict[str, Any]:
    """
    Lê dados de um arquivo de planilha (Excel ou CSV)
    
    Args:
        file_path: Caminho para o arquivo
        sheet_name: Nome da aba (para Excel). Se None, usa a primeira aba
        header: Linha a usar como cabeçalho (0-indexed). Se None, usa primeira linha (0).
                O pandas automaticamente pula as linhas antes do header.
        
    Returns:
        Dicionário com 'values' (lista de linhas) e 'headers' (primeira linha)
        
    Raises:
        FileNotFoundError: Se o arquivo não existir
        ValueError: Se o arquivo não puder ser lido
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    
    try:
        # Determinar tipo de arquivo pela extensão
        file_ext = Path(file_path).suffix.lower()
        
        if file_ext == '.csv':
            # Ler CSV
            read_params = {'encoding': 'utf-8'}
            if header is not None:
                read_params['header'] = header
            df = pd.read_csv(file_path, **read_params)
            logger.info(f"Arquivo CSV lido: {len(df)} linhas, {len(df.columns)} colunas (header={header})")
        elif file_ext in ['.xlsx', '.xls']:
            # Ler Excel
            read_params = {'engine': 'openpyxl'}
            if header is not None:
                read_params['header'] = header
            
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, **read_params)
            else:
                # Usar primeira aba
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                sheet_name = excel_file.sheet_names[0]
                df = pd.read_excel(file_path, sheet_name=sheet_name, **read_params)
                logger.info(f"Usando primeira aba: {sheet_name}")
            logger.info(f"Arquivo Excel lido: {len(df)} linhas, {len(df.columns)} colunas (header={header})")
        else:
            raise ValueError(f"Formato de arquivo não suportado: {file_ext}")
        
        # Converter DataFrame para formato compatível com google_sheets
        # Substituir NaN por strings vazias
        df = df.fillna('')
        
        # Converter para lista de listas
        values = df.values.tolist()
        headers = df.columns.tolist()
        
        # Converter headers para string se necessário
        headers = [str(h) for h in headers]
        
        # Converter valores para string se necessário (exceto números)
        processed_values = []
        for row in values:
            processed_row = []
            for val in row:
                if pd.isna(val) or val == '':
                    processed_row.append('')
                else:
                    processed_row.append(str(val))
            processed_values.append(processed_row)
        
        logger.info(f"Dados carregados: {len(processed_values)} linhas, {len(headers)} colunas")
        
        return {
            'values': processed_values,
            'headers': headers,
            'sheet_name': sheet_name if sheet_name else 'Sheet1'
        }
        
    except Exception as e:
        logger.error(f"Erro ao ler arquivo {file_path}: {str(e)}", exc_info=True)
        raise ValueError(f"Erro ao ler arquivo: {str(e)}")


def parse_status_data(
    data: Dict[str, Any],
    status_column: str,
    years: List[int],
    status_config: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Processa dados da planilha e agrupa por status e ano
    (Função compatível com google_sheets.py)
    
    Args:
        data: Dados retornados por read_spreadsheet_file
        status_column: Nome da coluna que contém os status
        years: Lista de anos para processar
        status_config: Configuração de status (main_statuses, other_statuses, etc.)
        
    Returns:
        Dicionário com 'main_statuses' e 'other_statuses' processados
    """
    headers = data.get('headers', [])
    rows = data.get('values', [])
    
    if not headers or not rows:
        logger.warning("Dados vazios para processar")
        return {
            'main_statuses': [],
            'other_statuses': []
        }
    
    # Encontrar índices das colunas
    try:
        status_col_idx = headers.index(status_column)
    except ValueError:
        logger.error(f"Coluna '{status_column}' não encontrada. Colunas disponíveis: {headers}")
        raise ValueError(f"Coluna '{status_column}' não encontrada")
    
    # Encontrar índices das colunas de anos
    year_prefix = status_config.get('columns', {}).get('year_prefix', 'Acionados em')
    year_col_indices = {}
    for year in years:
        year_col_name = f"{year_prefix} {year}"
        try:
            year_col_indices[year] = headers.index(year_col_name)
        except ValueError:
            logger.warning(f"Coluna '{year_col_name}' não encontrada para ano {year}")
            year_col_indices[year] = None
    
    # Encontrar índices de TOTAL e Percentual
    total_col_name = status_config.get('columns', {}).get('total_column', 'TOTAL')
    percentage_col_name = status_config.get('columns', {}).get('percentage_column', 'Percentual')
    
    try:
        total_col_idx = headers.index(total_col_name)
    except ValueError:
        total_col_idx = None
        logger.warning(f"Coluna '{total_col_name}' não encontrada")
    
    try:
        percentage_col_idx = headers.index(percentage_col_name)
    except ValueError:
        percentage_col_idx = None
        logger.warning(f"Coluna '{percentage_col_name}' não encontrada")
    
    # Processar linhas
    status_counts = {}
    include_blank = status_config.get('include_blank', False)
    
    for row in rows:
        if len(row) <= status_col_idx:
            continue
        
        status_value = row[status_col_idx].strip() if status_col_idx < len(row) else ""
        
        # Pular valores em branco se não incluídos
        if not status_value and not include_blank:
            continue
        
        if status_value not in status_counts:
            status_counts[status_value] = {
                'years': {year: 0 for year in years},
                'total': 0,
                'percentage': 0.0
            }
        
        # Processar valores por ano
        for year in years:
            if year_col_indices.get(year) is not None:
                col_idx = year_col_indices[year]
                if col_idx < len(row) and row[col_idx]:
                    try:
                        # Remover espaços e converter para int
                        value_str = str(row[col_idx]).strip().replace(',', '').replace('.', '')
                        value = int(value_str) if value_str else 0
                        status_counts[status_value]['years'][year] += value
                    except (ValueError, TypeError):
                        pass
        
        # Processar TOTAL
        if total_col_idx is not None and total_col_idx < len(row) and row[total_col_idx]:
            try:
                total_str = str(row[total_col_idx]).strip().replace(',', '').replace('.', '')
                status_counts[status_value]['total'] = int(total_str) if total_str else 0
            except (ValueError, TypeError):
                pass
        
        # Processar Percentual
        if percentage_col_idx is not None and percentage_col_idx < len(row) and row[percentage_col_idx]:
            try:
                # Remover % se presente
                pct_str = str(row[percentage_col_idx]).replace('%', '').replace(',', '.').strip()
                status_counts[status_value]['percentage'] = float(pct_str) if pct_str else 0.0
            except (ValueError, TypeError):
                pass
    
    # Separar status principais e outros
    main_statuses = []
    other_statuses = []
    
    main_status_values = {s['sheet_value'] for s in status_config.get('main_statuses', [])}
    other_status_values = {s['sheet_value'] for s in status_config.get('other_statuses', [])}
    
    for status_value, counts in status_counts.items():
        if status_value in main_status_values:
            # Encontrar display_name correspondente
            main_status_config = next(
                (s for s in status_config.get('main_statuses', []) if s['sheet_value'] == status_value),
                None
            )
            display_name = main_status_config['display_name'] if main_status_config else status_value
            
            main_statuses.append({
                'name': display_name,
                'sheet_value': status_value,
                'years': counts['years'],
                'total': counts['total'],
                'percentage': counts['percentage']
            })
        elif status_value in other_status_values:
            # Encontrar display_name correspondente
            other_status_config = next(
                (s for s in status_config.get('other_statuses', []) if s['sheet_value'] == status_value),
                None
            )
            display_name = other_status_config['display_name'] if other_status_config else status_value
            
            other_statuses.append({
                'name': display_name,
                'sheet_value': status_value,
                'years': counts['years'],
                'total': counts['total'],
                'percentage': counts['percentage']
            })
    
    # Ordenar conforme ordem na configuração
    main_statuses_ordered = []
    for config_status in status_config.get('main_statuses', []):
        found = next((s for s in main_statuses if s['sheet_value'] == config_status['sheet_value']), None)
        if found:
            main_statuses_ordered.append(found)
    
    other_statuses_ordered = []
    for config_status in status_config.get('other_statuses', []):
        found = next((s for s in other_statuses if s['sheet_value'] == config_status['sheet_value']), None)
        if found:
            other_statuses_ordered.append(found)
    
    return {
        'main_statuses': main_statuses_ordered,
        'other_statuses': other_statuses_ordered
    }
