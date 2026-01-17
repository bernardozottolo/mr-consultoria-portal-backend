"""
Blueprint para gerenciar upload e configuração de planilhas específicas do Enel
"""
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename
from .auth import login_required
from . import config
from data.database import get_db_connection
import os
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
        
        # #region agent log
        log_dir = Path('.cursor')
        log_dir.mkdir(exist_ok=True)
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,B,C',
                'location': 'enel_spreadsheets.py:84',
                'message': 'ANTES de salvar - estado do diretório',
                'data': {
                    'file_path': str(file_path),
                    'spreadsheets_dir': str(config.SPREADSHEETS_DIR),
                    'dir_exists': config.SPREADSHEETS_DIR.exists(),
                    'dir_writable': os.access(config.SPREADSHEETS_DIR, os.W_OK) if config.SPREADSHEETS_DIR.exists() else False,
                    'file_exists_before': file_path.exists(),
                    'files_in_dir_before': [f.name for f in config.SPREADSHEETS_DIR.iterdir()] if config.SPREADSHEETS_DIR.exists() else []
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
        try:
            file.save(str(file_path))
        except Exception as save_error:
            # #region agent log
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'B',
                    'location': 'enel_spreadsheets.py:85',
                    'message': 'ERRO durante file.save()',
                    'data': {
                        'error': str(save_error),
                        'error_type': type(save_error).__name__,
                        'file_path': str(file_path)
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            logger.error(f"ERRO ao salvar arquivo: {save_error}", exc_info=True)
            return jsonify({'error': f'Erro ao salvar arquivo: {str(save_error)}'}), 500
        
        # #region agent log
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,B,C',
                'location': 'enel_spreadsheets.py:87',
                'message': 'APÓS file.save() - verificação imediata',
                'data': {
                    'file_path': str(file_path),
                    'exists_immediately': file_path.exists(),
                    'file_size': file_path.stat().st_size if file_path.exists() else 0,
                    'files_in_dir_after_save': [f.name for f in config.SPREADSHEETS_DIR.iterdir()] if config.SPREADSHEETS_DIR.exists() else []
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
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
            
            # #region agent log
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'D',
                    'location': 'enel_spreadsheets.py:122',
                    'message': 'Registro existente encontrado - ANTES de remover arquivo antigo',
                    'data': {
                        'old_file_path': old_file_path,
                        'new_file_path': str(file_path),
                        'same_path': old_file_path == str(file_path),
                        'old_file_exists': os.path.exists(old_file_path),
                        'new_file_exists': file_path.exists(),
                        'files_in_dir_before_remove': [f.name for f in config.SPREADSHEETS_DIR.iterdir()] if config.SPREADSHEETS_DIR.exists() else []
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            
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
        
        # #region agent log
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,B,C,D',
                'location': 'enel_spreadsheets.py:146',
                'message': 'APÓS salvar no banco - verificação final',
                'data': {
                    'file_path': str(file_path),
                    'exists_after_db': file_path.exists(),
                    'file_size_after_db': file_path.stat().st_size if file_path.exists() else 0,
                    'files_in_dir_final': [f.name for f in config.SPREADSHEETS_DIR.iterdir()] if config.SPREADSHEETS_DIR.exists() else []
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
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
        
        # #region agent log
        log_dir = Path('.cursor')
        log_dir.mkdir(exist_ok=True)
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,B,C,D,E',
                'location': 'enel_spreadsheets.py:272',
                'message': 'Dados do banco de dados',
                'data': {
                    'spreadsheet_name': spreadsheet_name,
                    'file_path_from_db': result_dict.get('file_path'),
                    'file_name_from_db': result_dict.get('file_name'),
                    'sheet_name_from_db': result_dict.get('sheet_name'),
                    'status_column_from_db': result_dict.get('status_column')
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
        file_path = result_dict['file_path']
        # Sempre usar a primeira aba (ignorar o nome salvo no banco)
        status_column = result_dict['status_column'] if result_dict['status_column'] else 'Relatório Status detalhado'
        
        logger.info(f"Usando planilha: {spreadsheet_name}, primeira aba (automática), coluna: {status_column}")
        
        # Verificar se o arquivo existe
        # Converter para Path se necessário
        if isinstance(file_path, str):
            file_path_obj = Path(file_path)
        else:
            file_path_obj = file_path
        
        # #region agent log
        log_dir = Path('.cursor')
        log_dir.mkdir(exist_ok=True)
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,B,C',
                'location': 'enel_spreadsheets.py:283',
                'message': 'Caminho inicial construído',
                'data': {
                    'file_path_str': str(file_path),
                    'file_path_obj': str(file_path_obj),
                    'is_absolute': file_path_obj.is_absolute(),
                    'spreadsheets_dir': str(config.SPREADSHEETS_DIR),
                    'spreadsheets_dir_exists': config.SPREADSHEETS_DIR.exists() if hasattr(config, 'SPREADSHEETS_DIR') else False
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
        # Se o caminho não for absoluto, tentar construir caminho relativo ao SPREADSHEETS_DIR
        if not file_path_obj.is_absolute():
            # Tentar usar o caminho completo primeiro
            if config.SPREADSHEETS_DIR.exists():
                file_path_obj = config.SPREADSHEETS_DIR / file_path_obj.name
            else:
                # Se o diretório não existe, tentar usar o caminho como está
                pass
        
        # #region agent log
        log_dir = Path('.cursor')
        log_dir.mkdir(exist_ok=True)
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A,C',
                'location': 'enel_spreadsheets.py:298',
                'message': 'Caminho após processamento',
                'data': {
                    'file_path_obj': str(file_path_obj),
                    'exists': file_path_obj.exists() if file_path_obj else False
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
        # #endregion
        
        # Verificar se arquivo existe
        if not file_path_obj.exists():
            logger.warning(f"Arquivo não encontrado no caminho esperado: {file_path_obj}")
            logger.info(f"Caminho original do banco: {file_path}")
            logger.info(f"SPREADSHEETS_DIR: {config.SPREADSHEETS_DIR}")
            
            # Tentar encontrar o arquivo pelo nome no diretório de planilhas
            file_name = result_dict.get('file_name', '')
            found_file = None
            
            # #region agent log
            log_dir = Path('.cursor')
            log_dir.mkdir(exist_ok=True)
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'A,D,E',
                    'location': 'enel_spreadsheets.py:304',
                    'message': 'Iniciando busca alternativa',
                    'data': {
                        'file_name_from_db': file_name,
                        'spreadsheets_dir_exists': config.SPREADSHEETS_DIR.exists() if hasattr(config, 'SPREADSHEETS_DIR') else False
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            
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
                
                # #region agent log
                log_dir = Path('.cursor')
                log_dir.mkdir(exist_ok=True)
                with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                    f.write(json.dumps({
                        'sessionId': 'debug-session',
                        'runId': 'run1',
                        'hypothesisId': 'A',
                        'location': 'enel_spreadsheets.py:310',
                        'message': 'Tentando caminho alternativo por nome exato',
                        'data': {
                            'alternative_path': str(alternative_path),
                            'exists': alternative_path.exists()
                        },
                        'timestamp': int(datetime.now().timestamp() * 1000)
                    }) + '\n')
                # #endregion
                
                if alternative_path.exists():
                    logger.info(f"Arquivo encontrado por nome: {alternative_path}")
                    found_file = alternative_path
                else:
                    # Procurar arquivos que contenham parte do nome da planilha
                    # Normalizar nome da planilha para busca
                    spreadsheet_name_clean = spreadsheet_name.replace(' ', '_').replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A').lower()
                    all_files = list(config.SPREADSHEETS_DIR.glob('*'))
                    
                    # #region agent log
                    log_dir = Path('.cursor')
                    log_dir.mkdir(exist_ok=True)
                    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                        f.write(json.dumps({
                            'sessionId': 'debug-session',
                            'runId': 'run1',
                            'hypothesisId': 'E',
                            'location': 'enel_spreadsheets.py:316',
                            'message': 'Listando todos os arquivos no diretório',
                            'data': {
                                'all_files': [str(f.name) for f in all_files if f.is_file()],
                                'spreadsheet_name_clean': spreadsheet_name_clean
                            },
                            'timestamp': int(datetime.now().timestamp() * 1000)
                        }) + '\n')
                    # #endregion
                    
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
                                    
                                    # #region agent log
                                    log_dir = Path('.cursor')
                                    log_dir.mkdir(exist_ok=True)
                                    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                                        f.write(json.dumps({
                                            'sessionId': 'debug-session',
                                            'runId': 'run1',
                                            'hypothesisId': 'E',
                                            'location': 'enel_spreadsheets.py:323',
                                            'message': 'Arquivo encontrado por busca parcial',
                                            'data': {
                                                'found_file': str(found_file),
                                                'keywords_matched': keywords
                                            },
                                            'timestamp': int(datetime.now().timestamp() * 1000)
                                        }) + '\n')
                                    # #endregion
                                    
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
                    
                    # #region agent log
                    log_dir = Path('.cursor')
                    log_dir.mkdir(exist_ok=True)
                    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                        f.write(json.dumps({
                            'sessionId': 'debug-session',
                            'runId': 'run1',
                            'hypothesisId': 'A,C,D,E',
                            'location': 'enel_spreadsheets.py:326',
                            'message': 'Arquivo não encontrado - listando diretório',
                            'data': {
                                'searched_path': str(file_path_obj),
                                'file_name_from_db': file_name,
                                'files_in_dir': files_in_dir,
                                'spreadsheets_dir': str(config.SPREADSHEETS_DIR)
                            },
                            'timestamp': int(datetime.now().timestamp() * 1000)
                        }) + '\n')
                    # #endregion
                    
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
            
            # #region agent log
            log_dir = Path('.cursor')
            log_dir.mkdir(exist_ok=True)
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'A,C',
                    'location': 'enel_spreadsheets.py:345',
                    'message': 'Arquivo encontrado via busca alternativa',
                    'data': {
                        'final_file_path': str(file_path_obj),
                        'exists': file_path_obj.exists() if file_path_obj else False
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
        
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
        
        if not years:
            years = [2024, 2025]  # Fallback
        
        # Ler arquivo
        logger.info(f"Lendo arquivo: {file_path_obj}")
        try:
            sheet_data = read_spreadsheet_file(
                file_path=str(file_path_obj),
                sheet_name=None  # None = primeira aba automaticamente
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
                    sheet_data = read_spreadsheet_file(
                        file_path=str(file_path_obj),
                        sheet_name=None  # None = primeira aba automaticamente
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
            # #region agent log
            log_dir = Path('.cursor')
            log_dir.mkdir(exist_ok=True)
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'A',
                    'location': 'enel_spreadsheets.py:869',
                    'message': 'ANTES de processar dados - verificando colunas',
                    'data': {
                        'status_column': status_column,
                        'headers_available': sheet_data.get('headers', []),
                        'rows_count': len(sheet_data.get('values', []))
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            
            # Processar dados de Alvarás de Funcionamento (coluna padrão)
            processed_data = process_enel_legalizacao_data(
                data=sheet_data,
                status_column=status_column,
                years=years
            )
            
            # Processar dados de Licença Sanitária - Renovação (coluna 'Relatório Status detalhado acionamento')
            licenca_sanitaria_data = None
            try:
                licenca_sanitaria_status_column = 'Relatório Status detalhado acionamento'
                licenca_sanitaria_data = process_enel_legalizacao_data(
                    data=sheet_data,
                    status_column=licenca_sanitaria_status_column,
                    years=years
                )
                # Adicionar aos dados processados
                processed_data['licenca_sanitaria'] = licenca_sanitaria_data
            except Exception as e:
                logger.warning(f"Erro ao processar dados de Licença Sanitária: {e}")
                # Adicionar dados vazios se não conseguir processar
                processed_data['licenca_sanitaria'] = {
                    'total_demandado': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 100.0},
                    'concluidos': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                    'em_andamento': {
                        'total': {'years': {y: 0 for y in years}, 'total': 0, 'percentage': 0.0},
                        'subcategorias': []
                    }
                }
        except ValueError as ve:
            # Se a coluna não foi encontrada, retornar dados vazios com informações sobre colunas disponíveis
            if "não encontrada" in str(ve):
                headers = sheet_data.get('headers', [])
                
                # #region agent log
                with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                    f.write(json.dumps({
                        'sessionId': 'debug-session',
                        'runId': 'run1',
                        'hypothesisId': 'A',
                        'location': 'enel_spreadsheets.py:883',
                        'message': 'Coluna não encontrada - retornando dados vazios com info',
                        'data': {
                            'requested_column': status_column,
                            'available_columns': headers,
                            'error': str(ve)
                        },
                        'timestamp': int(datetime.now().timestamp() * 1000)
                    }) + '\n')
                # #endregion
                
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


def process_enel_legalizacao_data(data: dict, status_column: str, years: list) -> dict:
    """
    Processa dados da planilha para criar estrutura hierárquica:
    - Total demandado (total de registros)
    - Concluídos (status = "Concluído")
    - Alvarás em andamento (outros status como subcategorias)
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
    
    # Encontrar índice da coluna 'ano Acionamento' (case-insensitive, com trim)
    year_column_name = 'ano Acionamento'
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
    
    # Processar linhas
    status_counts = {}
    total_by_year = {year: 0 for year in years}
    total_all = 0
    
    for row in rows:
        # Verificar se a linha tem colunas suficientes
        if len(row) <= max(status_col_idx, year_col_idx):
            continue
        
        # Obter status
        status_value = row[status_col_idx].strip() if status_col_idx < len(row) else ""
        if not status_value:
            continue
        
        # Obter ano da coluna 'ano Acionamento'
        year_value_str = str(row[year_col_idx]).strip() if year_col_idx < len(row) else ""
        if not year_value_str:
            continue  # Pular linhas sem ano
        
        # Tentar converter o ano para inteiro
        try:
            row_year = int(year_value_str)
        except (ValueError, TypeError):
            # Se não conseguir converter, pular a linha
            continue
        
        # Verificar se o ano está na lista de anos solicitados
        if row_year not in years:
            continue  # Pular anos fora do range solicitado
        
        # Normalizar status (case-insensitive, remover espaços extras)
        status_normalized = ' '.join(status_value.split()).lower()
        
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
    em_andamento_subcategorias = []
    
    for status_norm, status_info in status_counts.items():
        if concluidos_normalized in status_norm:
            # É concluído
            for year in years:
                concluidos_data['years'][year] += status_info['years'][year]
            concluidos_data['total'] += status_info['total']
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
        'em_andamento': {
            'total': em_andamento_total,
            'subcategorias': em_andamento_subcategorias
        },
        'years': years
    }
