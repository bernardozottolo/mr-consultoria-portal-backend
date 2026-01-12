from flask import Blueprint, jsonify, render_template, Response, request, current_app
from .auth import login_required
from . import config
import os
import json
import logging
from pathlib import Path
try:
    from weasyprint import HTML
    from weasyprint.text.fonts import FontConfiguration
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

# Remover sys.path.insert - usar imports normais
from data import reports_db
from data.database import get_db_connection
import plotly.graph_objs as go
from .config import ROOT_DIR, REPORT_CONFIG_PATH, IMAGES_DIR
from .google_sheets import get_spreadsheet_data, parse_status_data
from .spreadsheet_files import read_spreadsheet_file, parse_status_data as parse_status_data_file

reports_bp = Blueprint('reports', __name__, url_prefix='/api', template_folder='templates')
logger = logging.getLogger(__name__)

def load_report_config():
    """Carrega configuração de relatórios do JSON"""
    try:
        if not os.path.exists(REPORT_CONFIG_PATH):
            logger.error(f"Arquivo de configuração não encontrado: {REPORT_CONFIG_PATH}")
            return None
        with open(REPORT_CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Erro ao carregar configuração: {str(e)}")
        return None

def get_spreadsheet_config_for_regional(config_data, regional):
    """Busca configuração de planilha para uma regional específica"""
    if not config_data:
        return None
    spreadsheets = config_data.get('spreadsheets', [])
    for sheet_config in spreadsheets:
        if sheet_config.get('regional') == regional.upper():
            return sheet_config
    return None

@reports_bp.route('/clients', methods=['GET'])
@login_required
def get_clients():
    """Lista todos os clientes disponíveis"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id, nome, logo_path FROM clients')
    clients = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'clients': clients})

@reports_bp.route('/reports/<client_id>/data/<regional>', methods=['GET'])
@login_required
def get_regional_data(client_id, regional):
    """Retorna dados de uma regional específica (arquivo ou Google Sheets)"""
    logger.info(f"Buscando dados para client_id={client_id}, regional={regional}")
    
    # Carregar configuração
    config_data = load_report_config()
    if not config_data:
        return jsonify({'error': 'Configuração de relatórios não encontrada'}), 500
    
    # Verificar se existe arquivo de planilha no banco de dados
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT file_path, sheet_name, status_column
        FROM spreadsheets
        WHERE regional = ?
    ''', (regional.upper(),))
    file_config = cursor.fetchone()
    conn.close()
    
    # Obter anos da query string ou usar padrão
    years_param = request.args.get('years', '')
    if years_param:
        try:
            years = [int(y.strip()) for y in years_param.split(',') if y.strip()]
        except ValueError:
            years = config_data.get('default_years', [])
    else:
        years = config_data.get('default_years', [])
    
    if not years:
        years = config_data.get('years', [])
    
    status_config = config_data.get('status_config', {})
    
    try:
        # Se existe arquivo, usar arquivo
        if file_config:
            logger.info(f"Usando arquivo de planilha para regional {regional}")
            file_path = file_config['file_path']
            sheet_name = file_config['sheet_name']
            status_column = file_config['status_column'] or 'Relatório Status detalhado'
            
            # Ler arquivo
            sheet_data = read_spreadsheet_file(
                file_path=file_path,
                sheet_name=sheet_name
            )
            
            # Processar dados
            processed_data = parse_status_data_file(
                data=sheet_data,
                status_column=status_column,
                years=years,
                status_config=status_config
            )
        else:
            # Usar Google Sheets (comportamento antigo)
            logger.info(f"Usando Google Sheets para regional {regional}")
            sheet_config = get_spreadsheet_config_for_regional(config_data, regional)
            if not sheet_config:
                return jsonify({'error': f'Configuração não encontrada para regional {regional}'}), 404
            
            spreadsheet_id = sheet_config.get('spreadsheet_id')
            if not spreadsheet_id:
                return jsonify({'error': f'Spreadsheet ID não configurado para regional {regional}'}), 400
            
            # Buscar dados do Google Sheets
            sheet_data = get_spreadsheet_data(
                spreadsheet_id=spreadsheet_id,
                sheet_name=sheet_config.get('sheet_name'),
                credentials_path=config.GOOGLE_SERVICE_ACCOUNT_FILE
            )
            
            # Processar dados
            status_column = sheet_config.get('status_column', 'Relatório Status detalhado')
            
            processed_data = parse_status_data(
                data=sheet_data,
                status_column=status_column,
                years=years,
                status_config=status_config
            )
        
        # Adicionar informações de configuração
        processed_data['config'] = {
            'years': years,
            'columns': config_data.get('columns', {}),
            'other_statuses_group': status_config.get('other_statuses_group', 'Outros')
        }
        
        return jsonify(processed_data)
        
    except Exception as e:
        logger.error(f"Erro ao buscar dados: {str(e)}", exc_info=True)
        error_message = str(e)
        if 'Acesso negado' in error_message or '403' in error_message:
            error_message = 'Acesso negado. Compartilhe a planilha com: google-sheets-service@mr-consultoria-reports-app.iam.gserviceaccount.com'
        return jsonify({'error': error_message}), 500

@reports_bp.route('/reports/<client_id>', methods=['GET'])
@login_required
def get_report(client_id):
    """Retorna os dados do relatório para um cliente (legado - mantido para compatibilidade)"""
    data = reports_db.get_report_data(client_id)
    if not data:
        return jsonify({'error': 'Cliente não encontrado'}), 404
    
    # Criar gráfico de barras verticais
    categorias = data['chart_data']['categories']
    valores = data['chart_data']['values']
    
    fig = go.Figure(data=[
        go.Bar(
            x=categorias,
            y=valores,
            marker_color='#2f84a8',
            text=valores,
            textposition='auto',
        )
    ])
    
    fig.update_layout(
        title='Gráfico de Barras Verticais',
        xaxis_title='Mês',
        yaxis_title='Quantidade',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(size=12, color='#2f84a8'),
        title_font=dict(size=16, color='#2f84a8'),
        height=400
    )
    
    # Converter gráfico para dict
    import plotly.utils
    graphJSON = plotly.utils.PlotlyJSONEncoder().encode(fig)
    graph_dict = json.loads(graphJSON)
    
    return jsonify({
        'table_data': data['table_data'],
        'graph_data': graph_dict
    })

@reports_bp.route('/reports/<client_id>/pdf', methods=['GET'])
@login_required
def generate_pdf(client_id):
    """Gera PDF do relatório para um cliente"""
    logger.info(f"=== PDF ENDPOINT CHAMADO ===")
    logger.info(f"PDF request recebido: client_id={client_id}, mes={request.args.get('mes')}, ano={request.args.get('ano')}")
    logger.info(f"Request path: {request.path}, Request url: {request.url}")
    
    if not WEASYPRINT_AVAILABLE:
        logger.error("WeasyPrint não está disponível")
        return jsonify({'error': 'WeasyPrint não está instalado. Execute: pip install WeasyPrint>=66.0'}), 500
    
    # Obter parâmetros de mês/ano
    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)
    
    # Obter estados selecionados (padrão: CE&SP&RJ)
    estados_param = request.args.get('estados', 'CE&SP&RJ')
    # Validar e limpar estados
    estados_validos = ['CE', 'SP', 'RJ']
    estados_lista = [e.strip().upper() for e in estados_param.split('&') if e.strip().upper() in estados_validos]
    if not estados_lista:
        estados_lista = ['CE', 'SP', 'RJ']  # Fallback para padrão
    estados_str = '|'.join(estados_lista)  # Usar | para exibição no PDF
    
    # Obter anos selecionados
    years_param = request.args.get('years', '')
    if years_param:
        try:
            years = [int(y.strip()) for y in years_param.split(',') if y.strip()]
        except ValueError:
            years = []
    else:
        years = []
    
    # Obter nomes de status customizados (JSON)
    status_names_param = request.args.get('status_names', '{}')
    try:
        status_names = json.loads(status_names_param) if status_names_param else {}
    except json.JSONDecodeError:
        status_names = {}
    
    # Obter comentários (JSON array)
    comments_param = request.args.get('comments', '[]')
    try:
        comments = json.loads(comments_param) if comments_param else []
    except json.JSONDecodeError:
        comments = []
    
    # Valores padrão se não fornecidos
    from datetime import datetime
    if mes is None:
        mes = datetime.now().month - 1  # 0-11
    if ano is None:
        ano = datetime.now().year
    
    # Validar valores
    if mes < 0 or mes > 11:
        mes = datetime.now().month - 1
    if ano < 2020 or ano > 2100:
        ano = datetime.now().year
    
    # Meses em português
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
             'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    month_name = meses[mes]
    
    # Buscar dados do cliente
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id, nome, logo_path FROM clients WHERE id = ?', (client_id,))
    client = cursor.fetchone()
    conn.close()
    
    if not client:
        return jsonify({'error': 'Cliente não encontrado'}), 404
    
    client_dict = dict(client)
    
    # Preparar paths das imagens e converter para base64 (WeasyPrint funciona melhor com base64)
    images_dir = config.IMAGES_DIR
    
    def get_image_base64(image_path):
        """Converte imagem para base64"""
        import base64
        try:
            if os.path.exists(image_path):
                with open(image_path, 'rb') as img_file:
                    img_data = img_file.read()
                    img_ext = os.path.splitext(image_path)[1].lower()
                    mime_type = 'image/png' if img_ext == '.png' else 'image/jpeg'
                    base64_str = f"data:{mime_type};base64,{base64.b64encode(img_data).decode('utf-8')}"
                    logger.info(f"Imagem convertida: {image_path} -> {len(base64_str)} chars")
                    return base64_str
            else:
                logger.warning(f"Arquivo de imagem não encontrado: {image_path}")
                return ""
        except Exception as e:
            logger.error(f"Erro ao converter imagem {image_path}: {e}")
            return ""
    
    # Construir caminhos corretos das imagens
    mr_logo_path = str(images_dir / 'mr-consultoria-logo.png')
    # Remover 'images/' do logo_path se existir e construir caminho completo
    client_logo_filename = client_dict['logo_path'].replace('images/', '').replace('static/images/', '')
    client_logo_path = str(images_dir / client_logo_filename)
    # Log de caminhos para debug
    logger.info(f"Procurando imagens em: {images_dir}")
    
    # Converter para base64 (apenas logos, sem background)
    mr_logo_base64 = get_image_base64(mr_logo_path)
    client_logo_base64 = get_image_base64(client_logo_path)
    
    # Log para debug
    logger.info(f"Imagens carregadas - MR Logo: {len(mr_logo_base64) > 0}, Client Logo: {len(client_logo_base64) > 0}")
    logger.info(f"Caminhos - MR: {mr_logo_path}, Client: {client_logo_path}")
    logger.info(f"Arquivos existem - MR: {os.path.exists(mr_logo_path)}, Client: {os.path.exists(client_logo_path)}")
    
    # Não buscar dados das planilhas - apenas títulos
    # Renderizar template HTML (versão simplificada - apenas títulos)
    html_content = render_template(
        'report_pdf.html',
        client_name=client_dict['nome'],
        month_name=month_name,
        year=ano,
        estados=estados_str,
        estados_lista=estados_lista,
        comments=comments,
        mr_logo_path=mr_logo_base64,
        client_logo_path=client_logo_base64
    )
    
    # Log do HTML gerado (primeiros 500 caracteres para debug)
    logger.info(f"HTML gerado (primeiros 500 chars): {html_content[:500]}")
    
    # Gerar PDF com WeasyPrint
    try:
        font_config = FontConfiguration()
        # base_url não é mais necessário pois imagens são base64
        pdf_bytes = HTML(string=html_content, base_url=str(images_dir) if images_dir.exists() else None).write_pdf(
            font_config=font_config
        )
        
        logger.info(f"PDF gerado com sucesso. Tamanho: {len(pdf_bytes)} bytes")
        
        # Verificar se é preview ou download (via query param)
        is_preview = request.args.get('preview', 'false').lower() == 'true'
        
        # Criar resposta com PDF
        filename = f"relatorio-{client_id}-{ano}-{mes+1:02d}.pdf"
        
        # Se for preview, usar 'inline' para abrir no navegador
        # Se for download, usar 'attachment' para forçar download
        disposition = 'inline' if is_preview else 'attachment'
        
        response = Response(
            pdf_bytes,
            mimetype='application/pdf',
            headers={
                'Content-Disposition': f'{disposition}; filename="{filename}"',
                'Content-Type': 'application/pdf'
            }
        )
        return response
    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)}", exc_info=True)
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500

