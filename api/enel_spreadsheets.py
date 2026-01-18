"""
Blueprint para gerenciar upload e configuração de planilhas específicas do Enel
"""
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename
from .auth import login_required
from . import config
from data.database import get_db_connection
import os
import re
import logging
from pathlib import Path
import json
from datetime import datetime

enel_spreadsheets_bp = Blueprint('enel_spreadsheets', __name__, url_prefix='/api/enel-spreadsheets')
logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# Nomes das planilhas necessárias para o relatório Enel
ENEL_REQUIRED_SPREADSHEETS = [
    'Base Ceara Alvarás de funcionamento',
    'CTEEP ATUALIZADA - BASE MR 2025',
    'ENEL - Legalização CE',
    'LEGALIZAÇÃO RJ_28-04',
    'Legalização SP',
    'Regularizações SP'
]


def allowed_file(filename):
    """Verifica se o arquivo tem extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@enel_spreadsheets_bp.route('/upload', methods=['POST'])
@login_required
def upload_enel_spreadsheet():
    """
    Endpoint para upload de planilha específica do Enel
    Espera um arquivo e parâmetro: spreadsheet_name (nome da planilha)
    """
    try:
        # Verificar se arquivo foi enviado
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        # Verificar extensão
        if not allowed_file(file.filename):
            return jsonify({
                'error': f'Formato não suportado. Use: {", ".join(ALLOWED_EXTENSIONS)}'
            }), 400
        
        # Obter parâmetros
        spreadsheet_name = request.form.get('spreadsheet_name', '').strip()
        if not spreadsheet_name:
            return jsonify({'error': 'Parâmetro "spreadsheet_name" é obrigatório'}), 400
        
        # Validar se o nome da planilha é um dos permitidos
        if spreadsheet_name not in ENEL_REQUIRED_SPREADSHEETS:
            return jsonify({
                'error': f'Nome de planilha inválido. Deve ser um dos: {", ".join(ENEL_REQUIRED_SPREADSHEETS)}'
            }), 400
        
        sheet_name = request.form.get('sheet_name', None)
        status_column = request.form.get('status_column', 'Relatório Status detalhado')
        
        # Criar nome seguro para o arquivo baseado apenas no nome da planilha
        # Usar nome da planilha para criar identificador único (evitar duplicação)
        safe_spreadsheet_id = spreadsheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_').replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A')
        
        # Obter extensão do arquivo original
        original_filename = secure_filename(file.filename)
        file_ext = Path(original_filename).suffix if '.' in original_filename else '.xlsx'
        
        # Criar nome único usando apenas o nome da planilha (evita duplicação)
        safe_filename = f"ENEL_{safe_spreadsheet_id}{file_ext}"
        
        # Salvar arquivo
        file_path = config.SPREADSHEETS_DIR / safe_filename
        
        try:
            file.save(str(file_path))
        except Exception as save_error:
            logger.error(f"ERRO ao salvar arquivo: {save_error}", exc_info=True)
            return jsonify({'error': f'Erro ao salvar arquivo: {str(save_error)}'}), 500
        
        # Validar que o arquivo foi salvo corretamente
        if not file_path.exists():
            logger.error(f"ERRO: Arquivo não foi salvo corretamente em {file_path}")
            return jsonify({'error': 'Erro ao salvar arquivo'}), 500
        
        logger.info(f"Arquivo salvo e validado: {file_path} para planilha Enel: {spreadsheet_name}")
        
        # Testar acesso ao arquivo imediatamente após salvar
        try:
            from .spreadsheet_files import read_spreadsheet_file
            # Tentar ler a primeira aba do arquivo para validar acesso
            test_data = read_spreadsheet_file(
                file_path=str(file_path),
                sheet_name=None  # Primeira aba
            )
            if test_data is None or len(test_data) == 0:
                logger.warning(f"Arquivo salvo mas parece estar vazio ou inacessível: {file_path}")
            else:
                logger.info(f"Arquivo acessível e legível: {file_path} ({len(test_data)} linhas lidas)")
        except Exception as e:
            logger.error(f"ERRO ao testar acesso ao arquivo salvo: {e}")
            # Não falhar o upload, mas registrar o erro
            logger.warning(f"Arquivo salvo em {file_path} mas houve erro ao testar acesso: {str(e)}")
        
        # Salvar informações no banco de dados
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Verificar se já existe registro para esta planilha
        cursor.execute(
            'SELECT id, file_path FROM enel_spreadsheets WHERE spreadsheet_name = ?',
            (spreadsheet_name,)
        )
        existing = cursor.fetchone()
        
        if existing:
            # Atualizar registro existente
            old_file_path = existing['file_path']
            
            # Remover arquivo antigo se existir E for diferente do novo
            if os.path.exists(old_file_path) and old_file_path != str(file_path):
                try:
                    os.remove(old_file_path)
                    logger.info(f"Arquivo antigo removido: {old_file_path}")
                except Exception as e:
                    logger.warning(f"Erro ao remover arquivo antigo: {e}")
            
            cursor.execute('''
                UPDATE enel_spreadsheets 
                SET file_path = ?, file_name = ?, sheet_name = ?, status_column = ?, uploaded_at = CURRENT_TIMESTAMP
                WHERE spreadsheet_name = ?
            ''', (str(file_path), safe_filename, sheet_name, status_column, spreadsheet_name))
        else:
            # Inserir novo registro
            cursor.execute('''
                INSERT INTO enel_spreadsheets (spreadsheet_name, file_path, file_name, sheet_name, status_column)
                VALUES (?, ?, ?, ?, ?)
            ''', (spreadsheet_name, str(file_path), safe_filename, sheet_name, status_column))
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': 'Planilha enviada com sucesso',
            'spreadsheet_name': spreadsheet_name,
            'file_name': safe_filename,
            'file_path': str(file_path)
        }), 200
        
    except Exception as e:
        logger.error(f"Erro ao fazer upload: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao fazer upload: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/list', methods=['GET'])
@login_required
def list_enel_spreadsheets():
    """Lista todas as planilhas do Enel (incluindo as que ainda não foram enviadas)"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT spreadsheet_name, file_name, sheet_name, status_column, uploaded_at
            FROM enel_spreadsheets
            ORDER BY spreadsheet_name
        ''')
        uploaded_spreadsheets = {row['spreadsheet_name']: dict(row) for row in cursor.fetchall()}
        conn.close()
        
        # Criar lista com todas as planilhas necessárias, indicando quais foram enviadas
        result = []
        for spreadsheet_name in ENEL_REQUIRED_SPREADSHEETS:
            if spreadsheet_name in uploaded_spreadsheets:
                result.append({
                    'spreadsheet_name': spreadsheet_name,
                    'file_name': uploaded_spreadsheets[spreadsheet_name]['file_name'],
                    'sheet_name': uploaded_spreadsheets[spreadsheet_name]['sheet_name'],
                    'status_column': uploaded_spreadsheets[spreadsheet_name]['status_column'],
                    'uploaded_at': uploaded_spreadsheets[spreadsheet_name]['uploaded_at'],
                    'is_uploaded': True
                })
            else:
                result.append({
                    'spreadsheet_name': spreadsheet_name,
                    'file_name': None,
                    'sheet_name': None,
                    'status_column': None,
                    'uploaded_at': None,
                    'is_uploaded': False
                })
        
        return jsonify({'spreadsheets': result}), 200
    except Exception as e:
        logger.error(f"Erro ao listar planilhas: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao listar planilhas: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/<spreadsheet_name>', methods=['GET'])
@login_required
def get_enel_spreadsheet_info(spreadsheet_name):
    """Obtém informações de uma planilha específica"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT spreadsheet_name, file_name, file_path, sheet_name, status_column, uploaded_at
            FROM enel_spreadsheets
            WHERE spreadsheet_name = ?
        ''', (spreadsheet_name,))
        
        result = cursor.fetchone()
        conn.close()
        
        if not result:
            return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404
        
        return jsonify(dict(result)), 200
    except Exception as e:
        logger.error(f"Erro ao buscar planilha: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao buscar planilha: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/required', methods=['GET'])
@login_required
def get_required_spreadsheets():
    """Retorna lista de planilhas necessárias para o relatório Enel"""
    return jsonify({'required_spreadsheets': ENEL_REQUIRED_SPREADSHEETS}), 200


@enel_spreadsheets_bp.route('/debug/files', methods=['GET'])
@login_required
def debug_list_files():
    """Endpoint de debug para listar arquivos no diretório de planilhas"""
    try:
        files = []
        if config.SPREADSHEETS_DIR.exists():
            for file_path in config.SPREADSHEETS_DIR.iterdir():
                if file_path.is_file():
                    files.append({
                        'name': file_path.name,
                        'path': str(file_path),
                        'exists': file_path.exists(),
                        'size': file_path.stat().st_size if file_path.exists() else 0
                    })
        
        # Também listar registros do banco
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT spreadsheet_name, file_path, file_name
            FROM enel_spreadsheets
        ''')
        db_records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return jsonify({
            'spreadsheets_dir': str(config.SPREADSHEETS_DIR),
            'dir_exists': config.SPREADSHEETS_DIR.exists(),
            'files_in_dir': files,
            'db_records': db_records
        }), 200
    except Exception as e:
        logger.error(f"Erro ao listar arquivos: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao listar arquivos: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/debug/logs/debug', methods=['GET'])
@login_required
def debug_get_debug_logs():
    """Endpoint para acessar logs de debug (NDJSON)"""
    try:
        debug_log_path = Path('.cursor') / 'debug.log'
        
        if not debug_log_path.exists():
            return jsonify({
                'error': 'Arquivo de log não encontrado',
                'path': str(debug_log_path),
                'exists': False
            }), 404
        
        # Ler últimas N linhas (padrão: 100)
        lines_limit = request.args.get('lines', type=int, default=100)
        
        with open(debug_log_path, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # Pegar últimas N linhas
        lines_to_return = all_lines[-lines_limit:] if len(all_lines) > lines_limit else all_lines
        
        # Tentar parsear cada linha como JSON
        parsed_logs = []
        for line_num, line in enumerate(lines_to_return, start=len(all_lines) - len(lines_to_return) + 1):
            line = line.strip()
            if not line:
                continue
            try:
                log_entry = json.loads(line)
                parsed_logs.append(log_entry)
            except json.JSONDecodeError:
                # Se não for JSON válido, adicionar como texto
                parsed_logs.append({
                    'line': line_num,
                    'raw': line,
                    'parse_error': True
                })
        
        return jsonify({
            'path': str(debug_log_path),
            'total_lines': len(all_lines),
            'returned_lines': len(parsed_logs),
            'logs': parsed_logs
        }), 200
    except Exception as e:
        logger.error(f"Erro ao ler logs de debug: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao ler logs de debug: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/debug/logs/app', methods=['GET'])
@login_required
def debug_get_app_logs():
    """Endpoint para acessar logs do Flask"""
    try:
        from datetime import datetime
        from .config import ROOT_DIR
        
        # Tentar arquivo de hoje primeiro
        today_log = ROOT_DIR / 'logs' / f'app_{datetime.now().strftime("%Y%m%d")}.log'
        
        # Se não existir, listar todos os arquivos de log disponíveis
        log_dir = ROOT_DIR / 'logs'
        if not log_dir.exists():
            return jsonify({
                'error': 'Diretório de logs não encontrado',
                'path': str(log_dir),
                'exists': False
            }), 404
        
        # Listar todos os arquivos de log
        log_files = sorted([f for f in log_dir.glob('app_*.log')], reverse=True)
        
        # Se especificou um arquivo específico via query param
        log_file_name = request.args.get('file', None)
        if log_file_name:
            log_path = log_dir / log_file_name
            if not log_path.exists():
                return jsonify({
                    'error': f'Arquivo de log não encontrado: {log_file_name}',
                    'available_files': [f.name for f in log_files]
                }), 404
        else:
            # Usar o mais recente ou o de hoje
            log_path = today_log if today_log.exists() else (log_files[0] if log_files else None)
        
        if not log_path or not log_path.exists():
            return jsonify({
                'error': 'Nenhum arquivo de log encontrado',
                'available_files': [f.name for f in log_files] if log_files else []
            }), 404
        
        # Ler últimas N linhas (padrão: 100)
        lines_limit = request.args.get('lines', type=int, default=100)
        
        with open(log_path, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # Pegar últimas N linhas
        lines_to_return = all_lines[-lines_limit:] if len(all_lines) > lines_limit else all_lines
        
        return jsonify({
            'path': str(log_path),
            'file_name': log_path.name,
            'total_lines': len(all_lines),
            'returned_lines': len(lines_to_return),
            'available_files': [f.name for f in log_files],
            'logs': [line.strip() for line in lines_to_return if line.strip()]
        }), 200
    except Exception as e:
        logger.error(f"Erro ao ler logs do Flask: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao ler logs do Flask: {str(e)}'}), 500


@enel_spreadsheets_bp.route('/<spreadsheet_name>/data', methods=['GET'])
@login_required
def get_enel_spreadsheet_data(spreadsheet_name):
    """
    Obtém dados processados de uma planilha específica do Enel
    Processa dados para criar estrutura hierárquica de estatísticas
    """
    try:
        from .spreadsheet_files import read_spreadsheet_file
        
        # Buscar informações da planilha
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT file_path, file_name, sheet_name, status_column
            FROM enel_spreadsheets
            WHERE spreadsheet_name = ?
        ''', (spreadsheet_name,))
        
        result = cursor.fetchone()
        conn.close()
        
        if not result:
            return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404
        
        # Converter Row para dict para facilitar acesso
        result_dict = dict(result)
        
        file_path = result_dict['file_path']
        # Sempre usar a primeira aba (ignorar o nome salvo no banco)
        # Para 'ENEL - Legalização CE', usar coluna 'Relatório Status detalhado acionamento'
        # Para outras planilhas, usar coluna padrão 'Relatório Status detalhado'
        status_column_override = request.args.get('status_column', None)
        if status_column_override:
            status_column = status_column_override
        elif spreadsheet_name == 'ENEL - Legalização CE':
            status_column = 'Relatório Status detalhado acionamento'
        else:
            status_column = result_dict['status_column'] if result_dict['status_column'] else 'Relatório Status detalhado'
        
        logger.info(f"Usando planilha: {spreadsheet_name}, primeira aba (automática), coluna: {status_column}")
        
        # Verificar se o arquivo existe
        # Converter para Path se necessário
        if isinstance(file_path, str):
            file_path_obj = Path(file_path)
        else:
            file_path_obj = file_path
        
        # Se o caminho não for absoluto, tentar construir caminho relativo ao SPREADSHEETS_DIR
        if not file_path_obj.is_absolute():
            # Tentar usar o caminho completo primeiro
            if config.SPREADSHEETS_DIR.exists():
                file_path_obj = config.SPREADSHEETS_DIR / file_path_obj.name
            else:
                # Se o diretório não existe, tentar usar o caminho como está
                pass
        
        # Verificar se arquivo existe
        if not file_path_obj.exists():
            logger.warning(f"Arquivo não encontrado no caminho esperado: {file_path_obj}")
            logger.info(f"Caminho original do banco: {file_path}")
            logger.info(f"SPREADSHEETS_DIR: {config.SPREADSHEETS_DIR}")
            
            # Tentar encontrar o arquivo pelo nome no diretório de planilhas
            file_name = result_dict.get('file_name', '')
            found_file = None
            
            # PRIMEIRO: Tentar construir o nome esperado usando a mesma lógica do upload
            if config.SPREADSHEETS_DIR.exists():
                # Usar a mesma lógica do upload para construir o nome esperado
                safe_spreadsheet_id = spreadsheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_').replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A')
                # Tentar extensões comuns
                for ext in ['.xlsx', '.xls']:
                    expected_filename = f"ENEL_{safe_spreadsheet_id}{ext}"
                    expected_path = config.SPREADSHEETS_DIR / expected_filename
                    if expected_path.exists():
                        logger.info(f"Arquivo encontrado pelo nome esperado (lógica upload): {expected_path}")
                        found_file = expected_path
                        break
            
            # SEGUNDO: Tentar pelo file_name do banco de dados
            if not found_file and file_name and config.SPREADSHEETS_DIR.exists():
                # Tentar encontrar por nome exato
                alternative_path = config.SPREADSHEETS_DIR / file_name
                
                if alternative_path.exists():
                    logger.info(f"Arquivo encontrado por nome: {alternative_path}")
                    found_file = alternative_path
                else:
                    # Procurar arquivos que contenham parte do nome da planilha
                    # Normalizar nome da planilha para busca
                    spreadsheet_name_clean = spreadsheet_name.replace(' ', '_').replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A').lower()
                    all_files = list(config.SPREADSHEETS_DIR.glob('*'))
                    
                    # Buscar arquivos ENEL relacionados a esta planilha
                    # Primeiro, tentar encontrar por padrão ENEL_ + nome da planilha
                    expected_pattern = f"ENEL_{spreadsheet_name_clean}"
                    for possible_file in all_files:
                        if possible_file.is_file():
                            file_name_lower = possible_file.name.lower()
                            # Verificar se começa com ENEL_ e contém partes do nome da planilha
                            if file_name_lower.startswith('enel_'):
                                # Verificar palavras-chave específicas da planilha
                                keywords = []
                                if 'ceara' in spreadsheet_name_clean or 'ceara' in file_name_lower:
                                    keywords.append('ceara')
                                if 'alvaras' in spreadsheet_name_clean or 'alvaras' in file_name_lower or 'alvarás' in file_name_lower:
                                    keywords.append('alvaras')
                                
                                # Se encontrou palavras-chave relevantes, usar este arquivo
                                if keywords:
                                    logger.info(f"Arquivo possível encontrado por palavras-chave: {possible_file}")
                                    found_file = possible_file
                                    
                                    break
            
            if not found_file:
                # ÚLTIMA TENTATIVA: Procurar qualquer arquivo ENEL_ que contenha palavras-chave relacionadas
                if config.SPREADSHEETS_DIR.exists():
                    try:
                        all_enel_files = [f for f in config.SPREADSHEETS_DIR.glob('ENEL_*') if f.is_file()]
                        # Normalizar nome da planilha para busca
                        spreadsheet_name_lower = spreadsheet_name.lower()
                        # Palavras-chave específicas para cada tipo de planilha
                        keywords_map = {
                            'ceara': ['ceara', 'ceará', 'cear'],
                            'alvaras': ['alvaras', 'alvarás', 'alvara'],
                            'legalizacao': ['legalizacao', 'legalização', 'legaliza'],
                            'regularizacao': ['regularizacao', 'regularização', 'regulariza']
                        }
                        
                        # Identificar palavras-chave relevantes para esta planilha
                        relevant_keywords = []
                        for key, variants in keywords_map.items():
                            if any(variant in spreadsheet_name_lower for variant in variants):
                                relevant_keywords.extend(variants)
                        
                        # Procurar arquivo que contenha essas palavras-chave
                        for enel_file in all_enel_files:
                            file_name_lower = enel_file.name.lower()
                            if any(keyword in file_name_lower for keyword in relevant_keywords):
                                logger.info(f"Arquivo encontrado por palavras-chave finais: {enel_file}")
                                found_file = enel_file
                                break
                    except Exception as e:
                        logger.error(f"Erro na busca final por palavras-chave: {e}")
                
                # Se ainda não encontrou, listar arquivos no diretório para debug
                if not found_file:
                    files_in_dir = []
                    if config.SPREADSHEETS_DIR.exists():
                        try:
                            files_in_dir = [str(f.name) for f in config.SPREADSHEETS_DIR.glob('*') if f.is_file()]
                        except Exception as e:
                            logger.error(f"Erro ao listar arquivos: {e}")
                    
                    return jsonify({
                    'error': f'Arquivo não encontrado: {file_path_obj}',
                    'original_path': str(file_path),
                    'searched_path': str(file_path_obj),
                    'file_name': file_name,
                    'spreadsheets_dir': str(config.SPREADSHEETS_DIR),
                    'files_in_dir': files_in_dir,
                    'hint': 'Verifique se o arquivo foi enviado corretamente. Use /api/enel-spreadsheets/debug/files para ver arquivos disponíveis.'
                }), 404
            
            file_path_obj = found_file
            
        # Obter anos da query string
        years_param = request.args.get('years', '')
        if years_param:
            try:
                years = [int(y.strip()) for y in years_param.split(',') if y.strip()]
            except ValueError:
                years = []
        else:
            # Usar anos padrão baseado em report_year_start e report_year_end
            report_year_start = request.args.get('report_year_start', type=int) or 2024
            report_year_end = request.args.get('report_year_end', type=int) or datetime.now().year
            years = list(range(report_year_start, report_year_end + 1))
        
        # Obter filtro de natureza da operação (para Licença Sanitária)
        filter_natureza = request.args.get('filter_natureza', None)
        
        # Decodificar URL se necessário
        if filter_natureza:
            from urllib.parse import unquote
            filter_natureza = unquote(filter_natureza)

        # Coluna de ano customizada (ex: Legalização SP)
        year_column_name = request.args.get('year_column', None)
        year_parse_mode = request.args.get('year_parse', None)
        if spreadsheet_name == 'Legalização SP' and not year_column_name:
            year_column_name = 'Data de acionamento MR'
            year_parse_mode = 'last4'
        
        # Parâmetros opcionais para filtros específicos
        sheet_name = request.args.get('sheet_name', None)
        header_row_param = request.args.get('header_row', None)
        header_row = int(header_row_param) if header_row_param is not None else None
        item_column = request.args.get('item_column', None)
        item_not_equals = request.args.get('item_not_equals', None)
        concluido_statuses_param = request.args.get('concluido_statuses', None)
        concluido_statuses = None
        if concluido_statuses_param:
            concluido_statuses = [v.strip() for v in concluido_statuses_param.split(',') if v.strip()]
        cancelado_statuses_param = request.args.get('cancelado_statuses', None)
        cancelado_statuses = None
        if cancelado_statuses_param:
            cancelado_statuses = [v.strip() for v in cancelado_statuses_param.split(',') if v.strip()]
        status_exclude_param = request.args.get('status_exclude', None)
        status_exclude = None
        if status_exclude_param:
            status_exclude = [v.strip() for v in status_exclude_param.split(',') if v.strip()]

        
        if not years:
            years = [2024, 2025]  # Fallback
        
        # Ler arquivo
        logger.info(f"Lendo arquivo: {file_path_obj}")
        try:
            # Somente 'ENEL - Legalização CE' tem tabela começando na 5ª linha (índice 4)
            # 'Base Ceara Alvarás de funcionamento' começa na primeira linha (normal)
            if header_row is None:
                if spreadsheet_name == 'ENEL - Legalização CE':
                    header_row = 4  # Linha 4 (0-indexed) = 5ª linha (pandas pula linhas 0-3 automaticamente)
                    logger.info(f"Planilha 'ENEL - Legalização CE' detectada: usando linha {header_row} (5ª linha) como cabeçalho")
                else:
                    # Outras planilhas (incluindo 'Base Ceara Alvarás de funcionamento') começam na primeira linha
                    header_row = None  # None = primeira linha (0) como cabeçalho
                    logger.info(f"Planilha '{spreadsheet_name}': usando primeira linha como cabeçalho")
            
            sheet_data = read_spreadsheet_file(
                file_path=str(file_path_obj),
                sheet_name=sheet_name,  # None = primeira aba automaticamente
                header=header_row
            )
        except FileNotFoundError as e:
            logger.error(f"Arquivo não encontrado: {e}")
            # Tentar buscar por nome similar no diretório
            file_name = result_dict.get('file_name', '')
            if file_name:
                # Procurar arquivos que contenham parte do nome
                possible_files = list(config.SPREADSHEETS_DIR.glob(f'*{file_name}*'))
                if possible_files:
                    logger.info(f"Arquivos similares encontrados: {[str(f) for f in possible_files]}")
                    # Tentar usar o primeiro arquivo encontrado
                    file_path_obj = possible_files[0]
                    logger.info(f"Tentando usar arquivo: {file_path_obj} (primeira aba)")
                    # Somente 'ENEL - Legalização CE' tem tabela começando na 5ª linha
                    if header_row is None:
                        header_row = None
                        if spreadsheet_name == 'ENEL - Legalização CE':
                            header_row = 4  # Linha 4 (0-indexed) = 5ª linha
                    sheet_data = read_spreadsheet_file(
                        file_path=str(file_path_obj),
                        sheet_name=sheet_name,  # None = primeira aba automaticamente
                        header=header_row
                    )
                else:
                    return jsonify({
                        'error': f'Arquivo não encontrado: {file_path_obj}',
                        'original_path': str(file_path),
                        'file_name': file_name,
                        'spreadsheets_dir': str(config.SPREADSHEETS_DIR),
                        'hint': 'Verifique se o arquivo foi enviado corretamente'
                    }), 404
            else:
                raise
        
        # Processar dados para estrutura hierárquica
        try:
            # Processar dados usando a coluna de status determinada acima
            processed_data = process_enel_legalizacao_data(
                data=sheet_data,
                status_column=status_column,
                years=years,
                filter_natureza=filter_natureza,
                year_column_name=year_column_name,
                year_parse_mode=year_parse_mode,
                item_column=item_column,
                item_not_equals=item_not_equals,
                concluido_statuses=concluido_statuses,
                cancelado_statuses=cancelado_statuses,
                status_exclude=status_exclude
            )
        except ValueError as ve:
            # Se a coluna não foi encontrada, retornar dados vazios com informações sobre colunas disponíveis
            if "não encontrada" in str(ve):
                headers = sheet_data.get('headers', [])
                
                processed_data = {
                    'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
                    'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                    'em_andamento': {
                        'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                        'subcategorias': []
                    },
                    'warning': f"Coluna '{status_column}' não encontrada",
                    'available_columns': headers,
                    'requested_column': status_column
                }
            else:
                raise
        
        return jsonify(processed_data), 200
        
    except FileNotFoundError as e:
        logger.error(f"Arquivo não encontrado: {str(e)}", exc_info=True)
        return jsonify({
            'error': f'Arquivo não encontrado: {str(e)}',
            'hint': 'Verifique se a planilha foi enviada corretamente através do upload'
        }), 404
    except Exception as e:
        logger.error(f"Erro ao buscar dados da planilha: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao buscar dados: {str(e)}'}), 500


def _get_enel_spreadsheet_data_internal(
    spreadsheet_name: str,
    years: list = None,
    filter_natureza: str = None,
    year_column_name: str = None,
    year_parse_mode: str = None,
    sheet_name: str = None,
    header_row: int = None,
    item_column: str = None,
    item_not_equals: str = None,
    concluido_statuses: list = None,
    cancelado_statuses: list = None,
    status_exclude: list = None,
    status_column_override: str = None
):
    """
    Função interna para obter dados de planilha sem depender do contexto Flask.
    Pode ser chamada diretamente com parâmetros.
    """
    from .spreadsheet_files import read_spreadsheet_file
    
    # Buscar informações da planilha
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT file_path, file_name, sheet_name, status_column
        FROM enel_spreadsheets
        WHERE spreadsheet_name = ?
    ''', (spreadsheet_name,))
    
    result = cursor.fetchone()
    conn.close()
    
    if not result:
        return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404
    
    result_dict = dict(result)
    file_path = result_dict['file_path']
    
    # Determinar coluna de status
    if status_column_override:
        status_column = status_column_override
    elif spreadsheet_name == 'ENEL - Legalização CE':
        status_column = 'Relatório Status detalhado acionamento'
    else:
        status_column = result_dict['status_column'] if result_dict['status_column'] else 'Relatório Status detalhado'
    
    # Converter para Path
    if isinstance(file_path, str):
        file_path_obj = Path(file_path)
    else:
        file_path_obj = file_path
    
    # Verificar se arquivo existe
    if not file_path_obj.exists():
        # Buscar arquivo alternativo (mesma lógica da função principal)
        file_name = result_dict.get('file_name', '')
        found_file = None
        
        if config.SPREADSHEETS_DIR.exists():
            safe_spreadsheet_id = spreadsheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_').replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A')
            for ext in ['.xlsx', '.xls']:
                expected_filename = f"ENEL_{safe_spreadsheet_id}{ext}"
                expected_path = config.SPREADSHEETS_DIR / expected_filename
                if expected_path.exists():
                    found_file = expected_path
                    break
        
        if not found_file and file_name and config.SPREADSHEETS_DIR.exists():
            alternative_path = config.SPREADSHEETS_DIR / file_name
            if alternative_path.exists():
                found_file = alternative_path
        
        if not found_file:
            return jsonify({'error': f'Arquivo não encontrado: {file_path_obj}'}), 404
        
        file_path_obj = found_file
    
    # Ler arquivo
    if header_row is None:
        if spreadsheet_name == 'ENEL - Legalização CE':
            header_row = 4
        else:
            header_row = None
    
    sheet_data = read_spreadsheet_file(
        file_path=str(file_path_obj),
        sheet_name=sheet_name,
        header=header_row
    )
    
    # Processar dados
    if years is None:
        years = [2024, 2025, 2026]  # Default

    if spreadsheet_name == 'Legalização SP' and not year_column_name:
        year_column_name = 'Data de acionamento MR'
        year_parse_mode = 'last4'
    
    processed_data = process_enel_legalizacao_data(
        data=sheet_data,
        status_column=status_column,
        years=years,
        filter_natureza=filter_natureza,
        year_column_name=year_column_name,
        year_parse_mode=year_parse_mode,
        item_column=item_column,
        item_not_equals=item_not_equals,
        concluido_statuses=concluido_statuses,
        cancelado_statuses=cancelado_statuses,
        status_exclude=status_exclude
    )
    
    return jsonify(processed_data), 200


def process_enel_legalizacao_data(
    data: dict,
    status_column: str,
    years: list,
    filter_natureza: str = None,
    year_column_name: str = None,
    year_parse_mode: str = None,
    item_column: str = None,
    item_not_equals: str = None,
    concluido_statuses: list = None,
    cancelado_statuses: list = None,
    status_exclude: list = None
) -> dict:
    """
    Processa dados da planilha para criar estrutura hierárquica:
    - Total demandado (total de registros)
    - Concluídos (status = "Concluído")
    - Alvarás em andamento (outros status como subcategorias)
    
    Args:
        data: Dados da planilha
        status_column: Nome da coluna de status
        years: Lista de anos para processar
        filter_natureza: Valor para filtrar na coluna 'Relatório Natureza da Operação' (opcional)
    """
    headers = data.get('headers', [])
    rows = data.get('values', [])
    
    if not headers or not rows:
        return {
            'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
            'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
            'em_andamento': {
                'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                'subcategorias': []
            }
        }
    
    # Encontrar índice da coluna de status (case-insensitive, com trim)
    status_col_idx = None
    status_column_normalized = status_column.strip().lower()
    for idx, header in enumerate(headers):
        if header.strip().lower() == status_column_normalized:
            status_col_idx = idx
            break
    
    if status_col_idx is None:
        logger.error(f"Coluna '{status_column}' não encontrada. Colunas disponíveis: {headers}")
        # Retornar dados vazios com informações sobre colunas disponíveis
        return {
            'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
            'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
            'em_andamento': {
                'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                'subcategorias': []
            },
            'warning': f"Coluna '{status_column}' não encontrada",
            'available_columns': headers,
            'requested_column': status_column
        }
    
    # Encontrar índice da coluna de ano (default: 'ano Acionamento')
    year_column_name = year_column_name or 'ano Acionamento'
    year_col_idx = None
    for idx, header in enumerate(headers):
        if header.strip().lower() == year_column_name.lower():
            year_col_idx = idx
            break
    
    if year_col_idx is None:
        logger.warning(f"Coluna '{year_column_name}' não encontrada. Colunas disponíveis: {headers}")
        # Se não encontrar a coluna de ano, retornar dados vazios com informações
        return {
            'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
            'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
            'em_andamento': {
                'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                'subcategorias': []
            },
            'warning': f"Coluna '{year_column_name}' não encontrada",
            'available_columns': headers,
            'requested_year_column': year_column_name
        }
    
    # Encontrar índice da coluna 'Relatório Natureza da Operação' se filtro for necessário
    natureza_col_idx = None
    if filter_natureza:
        natureza_column_name = 'Relatório Natureza da Operação'
        
        for idx, header in enumerate(headers):
            if header.strip().lower() == natureza_column_name.lower():
                natureza_col_idx = idx
                break
        
        if natureza_col_idx is None:
            logger.warning(f"Coluna '{natureza_column_name}' não encontrada para filtro. Colunas disponíveis: {headers}")
            
            # Se não encontrar a coluna de natureza, retornar dados vazios com informações
            return {
                'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
                'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                'em_andamento': {
                    'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                    'subcategorias': []
                },
                'warning': f"Coluna '{natureza_column_name}' não encontrada para filtro",
                'available_columns': headers,
                'requested_natureza_column': natureza_column_name
            }
    
    # Encontrar índice da coluna 'Item' se filtro for necessário
    item_col_idx = None
    if item_column:
        for idx, header in enumerate(headers):
            if header.strip().lower() == item_column.strip().lower():
                item_col_idx = idx
                break
        if item_col_idx is None:
            return {
                'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
                'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                'em_andamento': {
                    'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                    'subcategorias': []
                },
                'warning': f"Coluna '{item_column}' não encontrada",
                'available_columns': headers,
                'requested_item_column': item_column
            }

    
    # Normalizações de status para filtros/agrupamentos
    concluido_values_normalized = None
    cancelado_values_normalized = None
    status_exclude_normalized = None
    if concluido_statuses:
        concluido_values_normalized = {' '.join(v.split()).lower() for v in concluido_statuses if isinstance(v, str)}
    if cancelado_statuses:
        cancelado_values_normalized = {' '.join(v.split()).lower() for v in cancelado_statuses if isinstance(v, str)}
    if status_exclude:
        status_exclude_normalized = {' '.join(v.split()).lower() for v in status_exclude if isinstance(v, str)}

    # Processar linhas
    status_counts = {}
    total_by_year = {year: 0 for year in years}
    total_all = 0
    
    # Coletar valores únicos de natureza para debug
    natureza_values_set = set()
    rows_before_filter = 0
    rows_after_filter = 0
    rows_filtered_out = 0
    rows_skipped_empty_status = 0
    rows_skipped_empty_year = 0
    rows_skipped_year_parse = 0
    rows_skipped_year_not_in_range = 0
    status_value_samples = []
    year_value_samples = []
    for row in rows:
        rows_before_filter += 1
        
        # Verificar se a linha tem colunas suficientes
        required_indices = [status_col_idx, year_col_idx]
        if natureza_col_idx is not None:
            required_indices.append(natureza_col_idx)
        if item_col_idx is not None:
            required_indices.append(item_col_idx)
        if len(row) <= max(required_indices):
            continue
        
        # Aplicar filtro de item (ex: Item != 53)
        if item_col_idx is not None and item_not_equals is not None:
            item_value = row[item_col_idx] if item_col_idx < len(row) else ""
            item_value_str = str(item_value).strip()
            compare_value_str = str(item_not_equals).strip()
            
            def _values_equal(left, right):
                try:
                    return float(left) == float(right)
                except (ValueError, TypeError):
                    return left.strip().lower() == right.strip().lower()
            
            if _values_equal(item_value_str, compare_value_str):
                continue
        
        # Aplicar filtro de natureza da operação se necessário
        if filter_natureza and natureza_col_idx is not None:
            natureza_value = str(row[natureza_col_idx]).strip() if natureza_col_idx < len(row) else ""
            
            # Coletar valores únicos para debug
            if natureza_value:
                natureza_values_set.add(natureza_value)
            
            # Comparar valores (case-insensitive, com normalização de espaços)
            natureza_value_normalized = ' '.join(natureza_value.split()).lower()
            filter_natureza_normalized = ' '.join(filter_natureza.split()).lower()
            match = natureza_value_normalized == filter_natureza_normalized
            
            if not match:
                rows_filtered_out += 1
                continue  # Pular linhas que não correspondem ao filtro
        
        rows_after_filter += 1
        
        # Obter status
        status_value = row[status_col_idx].strip() if status_col_idx < len(row) else ""
        if not status_value:
            rows_skipped_empty_status += 1
            continue
        if len(status_value_samples) < 5:
            status_value_samples.append(status_value)

        # Normalizar status e aplicar exclusões, se houver
        status_normalized = ' '.join(status_value.split()).lower()
        if status_exclude_normalized and status_normalized in status_exclude_normalized:
            continue
        
        # Obter ano da coluna configurada
        year_value_str = str(row[year_col_idx]).strip() if year_col_idx < len(row) else ""
        if not year_value_str:
            rows_skipped_empty_year += 1
            continue  # Pular linhas sem ano
        if len(year_value_samples) < 5:
            year_value_samples.append({
                'raw': row[year_col_idx],
                'type': type(row[year_col_idx]).__name__,
                'str': year_value_str
            })
        
        # Interpretar ano conforme modo
        if year_parse_mode == 'last4':
            if year_value_str.lower() == 'não acionado':
                continue
            if len(year_value_str) < 4:
                continue
            year_suffix = year_value_str[-4:]
            if year_suffix.isdigit():
                row_year = int(year_suffix)
            else:
                # Tentar extrair ano dentro de strings com data/hora (ex: "2024-07-22 00:00:00")
                year_match = re.search(r'(19|20)\d{2}', year_value_str)
                if not year_match:
                    continue
                row_year = int(year_match.group(0))
        elif year_parse_mode == 'extract_year':
            year_match = re.search(r'(19|20)\d{2}', year_value_str)
            if not year_match:
                rows_skipped_year_parse += 1
                continue
            row_year = int(year_match.group(0))
        else:
            # Tentar converter o ano para inteiro (pode vir como float string "2024.0")
            try:
                # Primeiro tentar converter para float e depois para int (para tratar "2024.0")
                row_year = int(float(year_value_str))
            except (ValueError, TypeError):
                # Se não conseguir converter, pular a linha
                rows_skipped_year_parse += 1
                continue
        
        # Verificar se o ano está na lista de anos solicitados
        if row_year not in years:
            rows_skipped_year_not_in_range += 1
            continue  # Pular anos fora do range solicitado
        
        if status_normalized not in status_counts:
            status_counts[status_normalized] = {
                'original': status_value,  # Manter original para exibição
                'years': {year: 0 for year in years},
                'total': 0
            }
        
        # Contar esta linha para o ano correspondente
        status_counts[status_normalized]['years'][row_year] += 1
        total_by_year[row_year] += 1
        status_counts[status_normalized]['total'] += 1
        total_all += 1
    
    
    # Separar Concluídos e outros status
    concluidos_normalized = 'concluído'
    concluidos_data = {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0}
    cancelados_data = {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0}
    em_andamento_subcategorias = []
    
    for status_norm, status_info in status_counts.items():
        if concluido_values_normalized is not None:
            is_concluido = status_norm in concluido_values_normalized
        else:
            is_concluido = concluidos_normalized in status_norm
        if cancelado_values_normalized is not None:
            is_cancelado = status_norm in cancelado_values_normalized
        else:
            is_cancelado = False
        if is_concluido:
            # É concluído
            for year in years:
                concluidos_data['years'][year] += status_info['years'][year]
            concluidos_data['total'] += status_info['total']
        elif is_cancelado:
            for year in years:
                cancelados_data['years'][year] += status_info['years'][year]
            cancelados_data['total'] += status_info['total']
        else:
            # É subcategoria de "Alvarás em andamento"
            subcat = {
                'name': status_info['original'],
                'years': status_info['years'].copy(),
                'total': status_info['total'],
                'percentage': 0.0
            }
            em_andamento_subcategorias.append(subcat)
    
    # Calcular total de "Alvarás em andamento"
    em_andamento_total = {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0}
    for subcat in em_andamento_subcategorias:
        for year in years:
            em_andamento_total['years'][year] += subcat['years'][year]
        em_andamento_total['total'] += subcat['total']
    
    # Calcular percentuais
    if total_all > 0:
        concluidos_data['percentage'] = (concluidos_data['total'] / total_all) * 100
        cancelados_data['percentage'] = (cancelados_data['total'] / total_all) * 100
        em_andamento_total['percentage'] = (em_andamento_total['total'] / total_all) * 100
        for subcat in em_andamento_subcategorias:
            subcat['percentage'] = (subcat['total'] / total_all) * 100
    
    return {
        'total_demandado': {
            'years': total_by_year.copy(),
            'total': total_all,
            'percentage': 100.0
        },
        'concluidos': concluidos_data,
        'cancelados': cancelados_data,
        'em_andamento': {
            'total': em_andamento_total,
            'subcategorias': em_andamento_subcategorias
        },
        'years': years
    }
