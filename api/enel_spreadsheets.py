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
        
        # Criar nome seguro para o arquivo
        filename = secure_filename(file.filename)
        # Criar nome único baseado no nome da planilha
        file_ext = Path(filename).suffix
        file_base = Path(filename).stem
        # Usar nome da planilha para criar identificador único
        safe_spreadsheet_id = spreadsheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        safe_filename = f"ENEL_{safe_spreadsheet_id}_{file_base}{file_ext}"
        
        # Salvar arquivo
        file_path = config.SPREADSHEETS_DIR / safe_filename
        file.save(str(file_path))
        
        logger.info(f"Arquivo salvo: {file_path} para planilha Enel: {spreadsheet_name}")
        
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
            # Remover arquivo antigo se existir
            if os.path.exists(old_file_path):
                try:
                    os.remove(old_file_path)
                    logger.info(f"Arquivo antigo removido: {old_file_path}")
                except Exception as e:
                    logger.warning(f"Erro ao remover arquivo antigo: {e}")
            
            cursor.execute('''
                UPDATE enel_spreadsheets 
                SET file_path = ?, file_name = ?, sheet_name = ?, status_column = ?, uploaded_at = CURRENT_TIMESTAMP
                WHERE spreadsheet_name = ?
            ''', (str(file_path), filename, sheet_name, status_column, spreadsheet_name))
        else:
            # Inserir novo registro
            cursor.execute('''
                INSERT INTO enel_spreadsheets (spreadsheet_name, file_path, file_name, sheet_name, status_column)
                VALUES (?, ?, ?, ?, ?)
            ''', (spreadsheet_name, str(file_path), filename, sheet_name, status_column))
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': 'Planilha enviada com sucesso',
            'spreadsheet_name': spreadsheet_name,
            'file_name': filename,
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
