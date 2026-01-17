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
    
    # Obter parâmetros de configuração do relatório
    # report_month e report_year são os valores de referência do relatório
    report_month = request.args.get('report_month', type=int)
    report_year = request.args.get('report_year', type=int)
    
    # Se não fornecidos, usar mes/ano como fallback (compatibilidade)
    mes = report_month if report_month is not None else request.args.get('mes', type=int)
    ano = report_year if report_year is not None else request.args.get('ano', type=int)
    
    # Obter anos para estatísticas
    report_year_start = request.args.get('report_year_start', type=int)
    report_year_end = request.args.get('report_year_end', type=int)
    
    # Calcular anos baseado em report_year_start e report_year_end
    if report_year_start is None:
        report_year_start = 2024  # Padrão
    if report_year_end is None:
        from datetime import datetime
        report_year_end = datetime.now().year  # Padrão: ano atual
    
    # Gerar lista de anos
    years = list(range(report_year_start, report_year_end + 1))
    
    # Obter categorias de legalização e regularização
    legalizacao_param = request.args.get('legalizacao', 'CE,SP,RJ')
    regularizacao_param = request.args.get('regularizacao', 'RJ,SP,CTEEP')
    
    # Processar legalização (separado por vírgula)
    legalizacao_lista = [e.strip().upper() for e in legalizacao_param.split(',') if e.strip()]
    if not legalizacao_lista:
        legalizacao_lista = ['CE', 'SP', 'RJ']  # Padrão
    
    # Processar regularização (separado por vírgula)
    regularizacao_lista = [e.strip().upper() for e in regularizacao_param.split(',') if e.strip()]
    if not regularizacao_lista:
        regularizacao_lista = ['RJ', 'SP', 'CTEEP']  # Padrão
    
    # Obter estados selecionados (padrão: CE&SP&RJ) - mantido para compatibilidade
    estados_param = request.args.get('estados', 'CE&SP&RJ')
    # Validar e limpar estados
    estados_validos = ['CE', 'SP', 'RJ']
    estados_lista = [e.strip().upper() for e in estados_param.split('&') if e.strip().upper() in estados_validos]
    if not estados_lista:
        estados_lista = ['CE', 'SP', 'RJ']  # Fallback para padrão
    estados_str = '|'.join(estados_lista)  # Usar | para exibição no PDF
    
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
    
    # Separar comentários por página
    alvaras_comments = []
    licenca_comments = []
    for comment in comments:
        if isinstance(comment, dict):
            page = comment.get('page', '')
            if page == 'Licença Sanitária - Renovação':
                licenca_comments.append(comment)
            elif not page or page == 'Visão Geral - Alvarás de Funcionamento':
                alvaras_comments.append(comment)
        else:
            # Comentários antigos sem página definida vão para Alvarás
            alvaras_comments.append(comment)
    
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
    
    # Import datetime para logs (antes de usar)
    from datetime import datetime as dt
    
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
    
    # #region agent log
    log_dir = Path('.cursor')
    log_dir.mkdir(exist_ok=True)
    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
        f.write(json.dumps({
            'sessionId': 'debug-session',
            'runId': 'run1',
            'hypothesisId': 'A,B,C',
            'location': 'reports.py:351',
            'message': 'Dados do cliente obtidos do banco',
            'data': {
                'client_id': client_id,
                'client_id_lower': client_id.lower(),
                'client_dict': client_dict,
                'logo_path_from_db': client_dict.get('logo_path'),
                'client_nome': client_dict.get('nome')
            },
            'timestamp': int(dt.now().timestamp() * 1000)
        }) + '\n')
    # #endregion
    
    # Para ENEL, sempre buscar no frontend primeiro (caminho exato informado pelo usuário)
    # Tentar múltiplos caminhos possíveis para o frontend
    possible_frontend_dirs = [
        ROOT_DIR / 'portal-frontend' / 'images',  # Primeira opção: mesmo nível que portal-backend
        ROOT_DIR.parent / 'portal-frontend' / 'images',  # Segunda opção: nível acima
        Path('/app') / 'portal-frontend' / 'images',  # Caminho absoluto no container
        Path('/app/portal-frontend/images'),  # Caminho absoluto direto
    ]
    
    frontend_enel_path = None
    frontend_images_dir = None
    
    # #region agent log
    # Listar conteúdo de /app para entender estrutura
    app_dir_contents = []
    if os.path.exists('/app'):
        try:
            app_dir_contents = [item.name for item in Path('/app').iterdir() if item.is_dir()]
        except Exception as e:
            app_dir_contents = [f"Erro ao listar: {str(e)}"]
    
    # Tentar encontrar o arquivo usando busca recursiva limitada
    found_enel_logo_paths = []
    search_paths = ['/app', '/app/portal-frontend', '/app/portal-frontend/images']
    for search_path in search_paths:
        if os.path.exists(search_path):
            try:
                for root, dirs, files in os.walk(search_path):
                    if 'enel-logo.png' in files:
                        found_enel_logo_paths.append(os.path.join(root, 'enel-logo.png'))
                    # Limitar profundidade para não demorar muito
                    if root.count(os.sep) - search_path.count(os.sep) > 2:
                        dirs[:] = []
            except Exception as e:
                pass
    
    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
        f.write(json.dumps({
            'sessionId': 'debug-session',
            'runId': 'run1',
            'hypothesisId': 'A',
            'location': 'reports.py:368',
            'message': 'Tentando encontrar logo ENEL - listando caminhos',
            'data': {
                'possible_paths': [str(p / 'enel-logo.png') for p in possible_frontend_dirs],
                'paths_exist': [os.path.exists(str(p / 'enel-logo.png')) for p in possible_frontend_dirs],
                'app_dir_contents': app_dir_contents,
                'found_enel_logo_paths': found_enel_logo_paths,
                'root_dir': str(ROOT_DIR),
                'root_dir_parent': str(ROOT_DIR.parent)
            },
            'timestamp': int(dt.now().timestamp() * 1000)
        }) + '\n')
    # #endregion
    
    # Se encontrou o arquivo na busca, usar o primeiro resultado
    if found_enel_logo_paths:
        frontend_enel_path = found_enel_logo_paths[0]
        frontend_images_dir = Path(frontend_enel_path).parent
        logger.info(f"Logo ENEL encontrado via busca recursiva: {frontend_enel_path}")
    
    for frontend_dir in possible_frontend_dirs:
        enel_path = frontend_dir / 'enel-logo.png'
        enel_path_str = str(enel_path)
        if os.path.exists(enel_path_str):
            frontend_enel_path = enel_path_str
            frontend_images_dir = frontend_dir
            logger.info(f"Diretório frontend encontrado: {frontend_dir}, Logo ENEL: {frontend_enel_path}")
            # #region agent log
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'A',
                    'location': 'reports.py:385',
                    'message': 'Logo ENEL encontrado!',
                    'data': {
                        'found_path': frontend_enel_path,
                        'found_dir': str(frontend_images_dir)
                    },
                    'timestamp': int(dt.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            break
        else:
            logger.debug(f"Logo ENEL não encontrado em: {enel_path_str}")
    
    # #region agent log
    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
        f.write(json.dumps({
            'sessionId': 'debug-session',
            'runId': 'run1',
            'hypothesisId': 'A,B,C',
            'location': 'reports.py:380',
            'message': 'Verificando caminhos de logo',
            'data': {
                'client_id': client_id,
                'client_id_lower': client_id.lower(),
                'is_enel': client_id.lower() == 'enel' or client_dict.get('nome', '').upper() == 'ENEL',
                'frontend_images_dir': str(frontend_images_dir) if frontend_images_dir else None,
                'frontend_enel_path': frontend_enel_path,
                'frontend_enel_exists': os.path.exists(frontend_enel_path) if frontend_enel_path else False,
                'images_dir': str(images_dir),
                'client_logo_filename': client_logo_filename,
                'backend_logo_path': str(images_dir / client_logo_filename),
                'backend_logo_exists': os.path.exists(str(images_dir / client_logo_filename)),
                'root_dir': str(ROOT_DIR),
                'root_dir_parent': str(ROOT_DIR.parent)
            },
            'timestamp': int(dt.now().timestamp() * 1000)
        }) + '\n')
    # #endregion
    
    # Determinar caminho do logo do cliente
    client_logo_path = None  # Inicializar variável
    
    if client_id.lower() == 'enel' or client_dict.get('nome', '').upper() == 'ENEL':
        # Sempre usar logo do frontend para ENEL se encontrado
        if frontend_enel_path and os.path.exists(frontend_enel_path):
            client_logo_path = frontend_enel_path
            logger.info(f"Logo ENEL encontrado no frontend: {frontend_enel_path}")
        else:
            # Tentar backend como fallback
            backend_path = str(images_dir / client_logo_filename)
            if os.path.exists(backend_path):
                client_logo_path = backend_path
                logger.warning(f"Logo ENEL não encontrado no frontend, usando backend: {backend_path}")
            else:
                # Último recurso: tentar variações no backend
                alt_paths = [
                    str(images_dir / 'enel-logo.png'),
                    str(images_dir / 'ENEL-logo.png'),
                    str(images_dir / 'enel_logo.png'),
                    str(images_dir / 'ENEL_logo.png'),
                ]
                for alt_path in alt_paths:
                    if os.path.exists(alt_path):
                        client_logo_path = alt_path
                        logger.info(f"Logo ENEL encontrado em caminho alternativo: {client_logo_path}")
                        break
                else:
                    # Se não encontrou nada, usar o caminho do frontend mesmo que não exista (para debug)
                    client_logo_path = frontend_enel_path if frontend_enel_path else str(images_dir / client_logo_filename)
                    logger.error(f"Logo ENEL não encontrado em nenhum lugar! Tentando: {client_logo_path}")
    else:
        # Para outros clientes, tentar backend primeiro
        client_logo_path = str(images_dir / client_logo_filename)
        if not os.path.exists(client_logo_path) and frontend_images_dir:
            # Se não encontrou no backend, tentar no frontend
            frontend_client_logo_path = str(frontend_images_dir / client_logo_filename)
            if os.path.exists(frontend_client_logo_path):
                client_logo_path = frontend_client_logo_path
                logger.info(f"Logo encontrado no frontend: {frontend_client_logo_path}")
    
    # Garantir que client_logo_path está definido
    if not client_logo_path:
        client_logo_path = str(images_dir / client_logo_filename)
        logger.warning(f"client_logo_path não definido, usando padrão: {client_logo_path}")
    
    # Log de caminhos para debug
    logger.info(f"Procurando imagens em: {images_dir}")
    logger.info(f"Logo do cliente - Caminho: {client_logo_path}, Existe: {os.path.exists(client_logo_path)}")
    
    # Converter para base64 (apenas logos, sem background)
    mr_logo_base64 = get_image_base64(mr_logo_path)
    client_logo_base64 = get_image_base64(client_logo_path)
    
    # #region agent log
    with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
        f.write(json.dumps({
            'sessionId': 'debug-session',
            'runId': 'run1',
            'hypothesisId': 'C',
            'location': 'reports.py:375',
            'message': 'APÓS primeira tentativa de conversão',
            'data': {
                'client_logo_path': client_logo_path,
                'client_logo_path_exists': os.path.exists(client_logo_path),
                'client_logo_base64_length': len(client_logo_base64) if client_logo_base64 else 0,
                'mr_logo_base64_length': len(mr_logo_base64) if mr_logo_base64 else 0
            },
            'timestamp': int(dt.now().timestamp() * 1000)
        }) + '\n')
    # #endregion
    
    # Se ainda não encontrou o logo do cliente após converter, tentar buscar via HTTP do frontend
    if not client_logo_base64:
        # Para ENEL, tentar baixar via HTTP do container frontend (nginx)
        if client_id.lower() == 'enel' or client_dict.get('nome', '').upper() == 'ENEL':
            try:
                import urllib.request
                import socket
                
                # Tentar diferentes URLs possíveis para o logo
                # No Docker, o frontend está acessível pelo nome do serviço 'frontend' na mesma rede
                possible_urls = []
                
                # 1. Via nome do serviço Docker (mais provável em produção)
                possible_urls.append('http://frontend/images/enel-logo.png')
                
                # 2. Via variável de ambiente se configurada
                frontend_url = os.environ.get('FRONTEND_URL', '')
                if frontend_url:
                    possible_urls.insert(0, f'{frontend_url.rstrip("/")}/images/enel-logo.png')
                
                # 3. Via host da requisição atual (se frontend e backend estão no mesmo domínio)
                if request.host:
                    # Remover porta se houver
                    host_without_port = request.host.split(':')[0]
                    possible_urls.append(f'http://{host_without_port}/images/enel-logo.png')
                
                # 4. Tentativas locais para desenvolvimento
                possible_urls.extend([
                    'http://localhost/images/enel-logo.png',
                    'http://127.0.0.1/images/enel-logo.png',
                ])
                
                # #region agent log
                with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                    f.write(json.dumps({
                        'sessionId': 'debug-session',
                        'runId': 'run1',
                        'hypothesisId': 'C',
                        'location': 'reports.py:520',
                        'message': 'Tentando baixar logo ENEL via HTTP',
                        'data': {
                            'possible_urls': possible_urls,
                            'request_host': request.host if hasattr(request, 'host') else None
                        },
                        'timestamp': int(dt.now().timestamp() * 1000)
                    }) + '\n')
                # #endregion
                
                for url in possible_urls:
                    try:
                        req = urllib.request.Request(url)
                        req.add_header('User-Agent', 'Mozilla/5.0')
                        with urllib.request.urlopen(req, timeout=3) as response:
                            if response.status == 200:
                                img_data = response.read()
                                import base64
                                base64_str = f"data:image/png;base64,{base64.b64encode(img_data).decode('utf-8')}"
                                client_logo_base64 = base64_str
                                logger.info(f"Logo ENEL baixado via HTTP: {url}")
                                # #region agent log
                                with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                                    f.write(json.dumps({
                                        'sessionId': 'debug-session',
                                        'runId': 'run1',
                                        'hypothesisId': 'C',
                                        'location': 'reports.py:550',
                                        'message': 'Logo ENEL baixado via HTTP com sucesso!',
                                        'data': {
                                            'url': url,
                                            'base64_length': len(client_logo_base64) if client_logo_base64 else 0,
                                            'image_size_bytes': len(img_data)
                                        },
                                        'timestamp': int(dt.now().timestamp() * 1000)
                                    }) + '\n')
                                # #endregion
                                break
                    except urllib.error.HTTPError as e:
                        logger.debug(f"HTTP {e.code} ao baixar logo de {url}")
                        continue
                    except (urllib.error.URLError, socket.timeout, socket.gaierror) as e:
                        logger.debug(f"Erro de conexão ao baixar logo de {url}: {e}")
                        continue
                    except Exception as e:
                        logger.debug(f"Erro ao baixar logo de {url}: {e}")
                        continue
            except Exception as e:
                logger.warning(f"Erro ao tentar baixar logo via HTTP: {e}")
        
        # Se ainda não encontrou, tentar caminhos locais alternativos
        if not client_logo_base64 and frontend_images_dir:
            exact_enel_path = str(frontend_images_dir / 'enel-logo.png')
            if os.path.exists(exact_enel_path):
                client_logo_base64 = get_image_base64(exact_enel_path)
                logger.info(f"Logo encontrado no frontend (caminho exato): {exact_enel_path}")
            else:
                # Tentar também variações do nome no frontend
                for alt_name in ['ENEL-logo.png', 'enel_logo.png', 'ENEL_logo.png', client_logo_filename]:
                    alt_frontend_path = str(frontend_images_dir / alt_name)
                    if os.path.exists(alt_frontend_path):
                        client_logo_base64 = get_image_base64(alt_frontend_path)
                        logger.info(f"Logo encontrado no frontend (alternativo): {alt_frontend_path}")
                        break
    
    # Log para debug
    logger.info(f"Imagens carregadas - MR Logo: {len(mr_logo_base64) > 0}, Client Logo: {len(client_logo_base64) > 0}")
    logger.info(f"Caminhos - MR: {mr_logo_path}, Client: {client_logo_path}")
    logger.info(f"Arquivos existem - MR: {os.path.exists(mr_logo_path)}, Client: {os.path.exists(client_logo_path)}")
    
    # Buscar dados de Legalização CE se CE estiver na lista
    legalizacao_ce_data = None
    licenca_sanitaria_data = None
    
    if 'CE' in legalizacao_lista:
        try:
            from .enel_spreadsheets import get_enel_spreadsheet_data
            
            # Função auxiliar para converter chaves de anos
            def convert_years_keys(years_dict):
                """Converte chaves de string para int se necessário"""
                if not years_dict:
                    return years_dict
                keys_list = list(years_dict.keys())
                if keys_list and isinstance(keys_list[0], str):
                    return {int(k): v for k, v in years_dict.items()}
                return years_dict
            
            # 1. Buscar dados de Alvarás de Funcionamento da planilha 'Base Ceara Alvarás de funcionamento'
            spreadsheet_name_alvaras = 'Base Ceara Alvarás de funcionamento'
            years_str = ','.join(map(str, years))
            with current_app.test_request_context(
                path=f'/api/enel-spreadsheets/{spreadsheet_name_alvaras}/data',
                query_string=f'years={years_str}',
                headers={'Authorization': request.headers.get('Authorization', '')}
            ):
                result = get_enel_spreadsheet_data(spreadsheet_name_alvaras)
                if hasattr(result, 'get_json'):
                    legalizacao_ce_data = result.get_json()
                elif isinstance(result, tuple) and len(result) > 0:
                    if result[1] == 200:
                        legalizacao_ce_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                    else:
                        logger.warning(f"Erro ao buscar dados de Alvarás: status {result[1]}")
            
            # Converter anos nos dados de Alvarás
            if legalizacao_ce_data:
                if legalizacao_ce_data.get('total_demandado', {}).get('years'):
                    legalizacao_ce_data['total_demandado']['years'] = convert_years_keys(legalizacao_ce_data['total_demandado']['years'])
                if legalizacao_ce_data.get('concluidos', {}).get('years'):
                    legalizacao_ce_data['concluidos']['years'] = convert_years_keys(legalizacao_ce_data['concluidos']['years'])
                if legalizacao_ce_data.get('em_andamento', {}).get('total', {}).get('years'):
                    legalizacao_ce_data['em_andamento']['total']['years'] = convert_years_keys(legalizacao_ce_data['em_andamento']['total']['years'])
                for subcat in legalizacao_ce_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])
            
            # 2. Buscar dados de Licença Sanitária da planilha 'ENEL - Legalização CE'
            # Filtrar apenas registros onde 'Relatório Natureza da Operação' = 'Renovação Licença Sanitária'
            from urllib.parse import quote_plus
            spreadsheet_name_licenca = 'ENEL - Legalização CE'
            filter_natureza_value = 'Renovação Licença Sanitária'
            query_string_licenca = f'years={years_str}&filter_natureza={quote_plus(filter_natureza_value)}'
            # #region agent log
            import json
            from pathlib import Path
            from datetime import datetime as dt
            log_dir = Path('.cursor')
            log_dir.mkdir(exist_ok=True)
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'E',
                    'location': 'reports.py:715',
                    'message': 'Chamando API para Licença Sanitária com filtro',
                    'data': {
                        'spreadsheet_name': spreadsheet_name_licenca,
                        'query_string': query_string_licenca,
                        'filter_natureza_value': filter_natureza_value,
                        'years_str': years_str
                    },
                    'timestamp': int(dt.now().timestamp() * 1000)
                }) + '\n')
            # #endregion
            with current_app.test_request_context(
                path=f'/api/enel-spreadsheets/{spreadsheet_name_licenca}/data',
                query_string=query_string_licenca,
                headers={'Authorization': request.headers.get('Authorization', '')}
            ):
                result = get_enel_spreadsheet_data(spreadsheet_name_licenca)
                if hasattr(result, 'get_json'):
                    licenca_sanitaria_data = result.get_json()
                elif isinstance(result, tuple) and len(result) > 0:
                    if result[1] == 200:
                        licenca_sanitaria_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                    else:
                        logger.warning(f"Erro ao buscar dados de Licença Sanitária: status {result[1]}")
            
            # Converter anos nos dados de Licença Sanitária
            if licenca_sanitaria_data:
                if licenca_sanitaria_data.get('total_demandado', {}).get('years'):
                    licenca_sanitaria_data['total_demandado']['years'] = convert_years_keys(licenca_sanitaria_data['total_demandado']['years'])
                if licenca_sanitaria_data.get('concluidos', {}).get('years'):
                    licenca_sanitaria_data['concluidos']['years'] = convert_years_keys(licenca_sanitaria_data['concluidos']['years'])
                if licenca_sanitaria_data.get('em_andamento', {}).get('total', {}).get('years'):
                    licenca_sanitaria_data['em_andamento']['total']['years'] = convert_years_keys(licenca_sanitaria_data['em_andamento']['total']['years'])
                for subcat in licenca_sanitaria_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Legalização CE: {e}", exc_info=True)
    
    # Renderizar template HTML
    html_content = render_template(
        'report_pdf.html',
        client_name=client_dict['nome'],
        month_name=month_name,
        year=ano,
        estados=estados_str,
        estados_lista=estados_lista,
        legalizacao_lista=legalizacao_lista,
        legalizacao_ce_data=legalizacao_ce_data,
        licenca_sanitaria_data=licenca_sanitaria_data,
        years=years,
        comments=comments,
        alvaras_comments=alvaras_comments,
        licenca_comments=licenca_comments,
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
        import traceback
        logger.error(f"Traceback completo: {traceback.format_exc()}")
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500

