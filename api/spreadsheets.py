"""
Blueprint para gerenciar upload e configuração de planilhas
"""
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename
from .auth import login_required
from . import config
from data.database import get_db_connection
import os
import logging
from pathlib import Path

spreadsheets_bp = Blueprint('spreadsheets', __name__, url_prefix='/api/spreadsheets')
logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}


def allowed_file(filename):
    """Verifica se o arquivo tem extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@spreadsheets_bp.route('/upload', methods=['POST'])
@login_required
def upload_spreadsheet():
    """
    Endpoint para upload de planilha
    Espera um arquivo e parâmetros: regional, sheet_name (opcional), status_column (opcional)
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
        regional = request.form.get('regional', '').upper()
        if not regional:
            return jsonify({'error': 'Parâmetro "regional" é obrigatório'}), 400
        
        sheet_name = request.form.get('sheet_name', None)
        status_column = request.form.get('status_column', 'Relatório Status detalhado')
        
        # Criar nome seguro para o arquivo
        filename = secure_filename(file.filename)
        # Adicionar regional ao nome do arquivo para evitar conflitos
        file_ext = Path(filename).suffix
        file_base = Path(filename).stem
        safe_filename = f"{regional}_{file_base}{file_ext}"
        
        # Salvar arquivo
        file_path = config.SPREADSHEETS_DIR / safe_filename
        file.save(str(file_path))
        
        logger.info(f"Arquivo salvo: {file_path} para regional {regional}")
        
        # Salvar informações no banco de dados
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Verificar se já existe registro para esta regional
        cursor.execute(
            'SELECT id, file_path FROM spreadsheets WHERE regional = ?',
            (regional,)
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
                UPDATE spreadsheets 
                SET file_path = ?, file_name = ?, sheet_name = ?, status_column = ?
                WHERE regional = ?
            ''', (str(file_path), filename, sheet_name, status_column, regional))
        else:
            # Inserir novo registro
            cursor.execute('''
                INSERT INTO spreadsheets (regional, file_path, file_name, sheet_name, status_column)
                VALUES (?, ?, ?, ?, ?)
            ''', (regional, str(file_path), filename, sheet_name, status_column))
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'message': 'Planilha enviada com sucesso',
            'regional': regional,
            'file_name': filename,
            'file_path': str(file_path)
        }), 200
        
    except Exception as e:
        logger.error(f"Erro ao fazer upload: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao fazer upload: {str(e)}'}), 500


@spreadsheets_bp.route('/list', methods=['GET'])
@login_required
def list_spreadsheets():
    """Lista todas as planilhas configuradas"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT regional, file_name, sheet_name, status_column, uploaded_at
            FROM spreadsheets
            ORDER BY regional
        ''')
        spreadsheets = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return jsonify({'spreadsheets': spreadsheets}), 200
    except Exception as e:
        logger.error(f"Erro ao listar planilhas: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao listar planilhas: {str(e)}'}), 500


@spreadsheets_bp.route('/<regional>', methods=['GET'])
@login_required
def get_spreadsheet_info(regional):
    """Obtém informações de uma planilha específica"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT regional, file_name, file_path, sheet_name, status_column, uploaded_at
            FROM spreadsheets
            WHERE regional = ?
        ''', (regional.upper(),))
        
        result = cursor.fetchone()
        conn.close()
        
        if not result:
            return jsonify({'error': f'Planilha não encontrada para regional {regional}'}), 404
        
        return jsonify(dict(result)), 200
    except Exception as e:
        logger.error(f"Erro ao buscar planilha: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao buscar planilha: {str(e)}'}), 500


@spreadsheets_bp.route('/<regional>', methods=['DELETE'])
@login_required
def delete_spreadsheet(regional):
    """Remove uma planilha"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT file_path FROM spreadsheets WHERE regional = ?', (regional.upper(),))
        result = cursor.fetchone()
        
        if not result:
            return jsonify({'error': f'Planilha não encontrada para regional {regional}'}), 404
        
        file_path = result['file_path']
        
        # Remover arquivo
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                logger.info(f"Arquivo removido: {file_path}")
            except Exception as e:
                logger.warning(f"Erro ao remover arquivo: {e}")
        
        # Remover do banco
        cursor.execute('DELETE FROM spreadsheets WHERE regional = ?', (regional.upper(),))
        conn.commit()
        conn.close()
        
        return jsonify({'message': f'Planilha removida para regional {regional}'}), 200
    except Exception as e:
        logger.error(f"Erro ao remover planilha: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao remover planilha: {str(e)}'}), 500
