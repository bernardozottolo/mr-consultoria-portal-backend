"""
Módulo de integração com Google Sheets API
"""
import os
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

# Cache do cliente para evitar múltiplas inicializações
_sheets_client = None

def get_sheets_client(credentials_path: str):
    """
    Inicializa e retorna o cliente Google Sheets usando service account
    
    Args:
        credentials_path: Caminho para o arquivo JSON de credenciais
        
    Returns:
        Cliente Google Sheets API
        
    Raises:
        FileNotFoundError: Se o arquivo de credenciais não existir
        ValueError: Se as credenciais forem inválidas
    """
    global _sheets_client
    
    if _sheets_client is not None:
        return _sheets_client
    
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(f"Arquivo de credenciais não encontrado: {credentials_path}")
    
    try:
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
        )
        _sheets_client = build('sheets', 'v4', credentials=credentials)
        logger.info(f"Cliente Google Sheets inicializado com sucesso")
        return _sheets_client
    except Exception as e:
        logger.error(f"Erro ao inicializar cliente Google Sheets: {str(e)}")
        raise ValueError(f"Erro ao carregar credenciais: {str(e)}")


def get_spreadsheet_data(
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
    range_name: Optional[str] = None,
    credentials_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Busca dados de uma planilha Google Sheets
    
    Args:
        spreadsheet_id: ID da planilha (da URL: .../d/{SPREADSHEET_ID}/...)
        sheet_name: Nome da aba (se None, usa a primeira aba)
        range_name: Range específico (ex: 'A1:Z1000'). Se None, busca toda a aba
        credentials_path: Caminho para credenciais (usa config se None)
        
    Returns:
        Dicionário com 'values' (lista de linhas) e 'headers' (primeira linha)
        
    Raises:
        HttpError: Se houver erro na API (planilha não compartilhada, etc.)
    """
    if credentials_path is None:
        from . import config
        credentials_path = config.GOOGLE_SERVICE_ACCOUNT_FILE
    
    try:
        service = get_sheets_client(credentials_path)
        
        # Se sheet_name não foi fornecido, buscar primeira aba
        if sheet_name is None:
            spreadsheet_metadata = service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            sheets = spreadsheet_metadata.get('sheets', [])
            if not sheets:
                raise ValueError("Planilha não possui abas")
            sheet_name = sheets[0]['properties']['title']
            logger.info(f"Usando primeira aba: {sheet_name}")
        
        # Construir range
        if range_name:
            range_full = f"{sheet_name}!{range_name}"
        else:
            range_full = sheet_name
        
        # Buscar dados
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_full
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            logger.warning(f"Nenhum dado encontrado na planilha {spreadsheet_id}, aba {sheet_name}")
            return {'values': [], 'headers': []}
        
        # Primeira linha são os headers
        headers = values[0] if values else []
        data_rows = values[1:] if len(values) > 1 else []
        
        logger.info(f"Dados carregados: {len(data_rows)} linhas, {len(headers)} colunas")
        
        return {
            'values': data_rows,
            'headers': headers,
            'sheet_name': sheet_name
        }
        
    except HttpError as e:
        error_details = json.loads(e.content.decode('utf-8'))
        error_message = error_details.get('error', {}).get('message', str(e))
        
        if e.resp.status == 403:
            logger.error(f"Acesso negado à planilha {spreadsheet_id}. Verifique se a planilha está compartilhada com o service account.")
            raise HttpError(
                resp=e.resp,
                content=f"Acesso negado. Compartilhe a planilha com: google-sheets-service@mr-consultoria-reports-app.iam.gserviceaccount.com".encode('utf-8')
            )
        elif e.resp.status == 404:
            logger.error(f"Planilha {spreadsheet_id} não encontrada")
            raise HttpError(
                resp=e.resp,
                content=f"Planilha não encontrada. Verifique o ID da planilha.".encode('utf-8')
            )
        else:
            logger.error(f"Erro ao buscar dados da planilha: {error_message}")
            raise
    
    except Exception as e:
        logger.error(f"Erro inesperado ao buscar dados: {str(e)}")
        raise


def parse_status_data(
    data: Dict[str, Any],
    status_column: str,
    years: List[int],
    status_config: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Processa dados da planilha e agrupa por status e ano
    
    Args:
        data: Dados retornados por get_spreadsheet_data
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
                        value = int(row[col_idx])
                        status_counts[status_value]['years'][year] += value
                    except (ValueError, TypeError):
                        pass
        
        # Processar TOTAL
        if total_col_idx is not None and total_col_idx < len(row) and row[total_col_idx]:
            try:
                status_counts[status_value]['total'] = int(row[total_col_idx])
            except (ValueError, TypeError):
                pass
        
        # Processar Percentual
        if percentage_col_idx is not None and percentage_col_idx < len(row) and row[percentage_col_idx]:
            try:
                # Remover % se presente
                pct_str = row[percentage_col_idx].replace('%', '').replace(',', '.').strip()
                status_counts[status_value]['percentage'] = float(pct_str)
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

