from flask import Blueprint, jsonify, render_template, Response, request, current_app
from .auth import login_required
from . import config
import os
import json
import re
import logging
from pathlib import Path
from datetime import datetime
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
from .config import ROOT_DIR, IMAGES_DIR
from .spreadsheet_files import read_spreadsheet_file

reports_bp = Blueprint('reports', __name__, url_prefix='/api', template_folder='templates')
logger = logging.getLogger(__name__)

def _find_enel_spreadsheet_file(spreadsheet_name: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT file_path, file_name
        FROM enel_spreadsheets
        WHERE spreadsheet_name = ?
    ''', (spreadsheet_name,))
    result = cursor.fetchone()
    conn.close()
    if not result:
        return None

    file_path = result['file_path']
    file_name = result['file_name']
    file_path_obj = Path(file_path) if isinstance(file_path, str) else file_path

    if file_path_obj.exists():
        return file_path_obj

    if config.SPREADSHEETS_DIR.exists():
        safe_spreadsheet_id = spreadsheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_') \
            .replace('á', 'a').replace('Á', 'A').replace('ã', 'a').replace('Ã', 'A')
        for ext in ['.xlsx', '.xls']:
            expected_path = config.SPREADSHEETS_DIR / f"ENEL_{safe_spreadsheet_id}{ext}"
            if expected_path.exists():
                return expected_path

    if file_name and config.SPREADSHEETS_DIR.exists():
        alternative_path = config.SPREADSHEETS_DIR / file_name
        if alternative_path.exists():
            return alternative_path

    return None

def _build_regularizacao_sp_macroprocess(sheet_data: dict):
    headers = sheet_data.get('headers', [])
    rows = sheet_data.get('values', [])
    macro_idx = None
    for idx, header in enumerate(headers):
        if header.strip().lower() == 'macroprocesso':
            macro_idx = idx
            break
    if macro_idx is None:
        return {
            'items': [],
            'total_all': 0,
            'macro_idx': None,
            'warning': "Coluna 'Macroprocesso' não encontrada",
            'available_columns': headers
        }

    counts = {}
    for row in rows:
        if len(row) <= macro_idx:
            continue
        raw_value = str(row[macro_idx]).strip()
        if not raw_value or raw_value.lower() in ('nan', 'none'):
            continue
        counts[raw_value] = counts.get(raw_value, 0) + 1

    def sort_key(name: str):
        match = re.match(r'\s*(\d+)', name)
        return (int(match.group(1)) if match else 9999, name)

    total_all = sum(counts.values())
    items = []
    for name in sorted(counts.keys(), key=sort_key):
        total = counts[name]
        percentage = (total / total_all * 100) if total_all else 0.0
        items.append({
            'name': name,
            'total': total,
            'percentage': round(percentage, 2)
        })

    return {
        'items': items,
        'total_all': total_all,
        'macro_idx': macro_idx
    }

def _build_regularizacao_rj_macro_microprocess(sheet_data: dict):
    """Processa dados de Regularização RJ com macroprocessos e microprocessos"""
    headers = sheet_data.get('headers', [])
    rows = sheet_data.get('values', [])
    
    macro_idx = None
    micro_idx = None
    
    for idx, header in enumerate(headers):
        header_lower = header.strip().upper()
        if header_lower == 'MACROPROCESSO':
            macro_idx = idx
        elif header_lower == 'MICROPROCESSO':
            micro_idx = idx
    
    if macro_idx is None:
        return {
            'items': [],
            'total_all': 0,
            'warning': "Coluna 'MACROPROCESSO' não encontrada",
            'available_columns': headers
        }
    
    # Agrupar por macroprocesso e microprocesso
    macro_data = {}  # {macro_name: {'micros': {micro_name: count}, 'total': count}}
    
    for row in rows:
        if len(row) <= max(macro_idx, micro_idx if micro_idx is not None else -1):
            continue
        
        macro_value = str(row[macro_idx]).strip() if macro_idx < len(row) else ''
        micro_value = str(row[micro_idx]).strip() if micro_idx is not None and micro_idx < len(row) else ''
        
        if not macro_value or macro_value.lower() in ('nan', 'none', ''):
            continue
        
        # Inicializar macro se não existir
        if macro_value not in macro_data:
            macro_data[macro_value] = {'micros': {}, 'total': 0}
        
        # Se tem microprocesso, contar separadamente
        if micro_value and micro_value.lower() not in ('nan', 'none', ''):
            if micro_value not in macro_data[macro_value]['micros']:
                macro_data[macro_value]['micros'][micro_value] = 0
            macro_data[macro_value]['micros'][micro_value] += 1
        
        # Contar no total do macroprocesso
        macro_data[macro_value]['total'] += 1
    
    # Converter para estrutura hierárquica
    def sort_key(name: str):
        # Extrair números do início para ordenação
        match = re.match(r'\s*(\d+(?:\.\d+)*)', name)
        if match:
            parts = match.group(1).split('.')
            return tuple(int(p) for p in parts) + (name,)
        return (9999, name)
    
    items = []
    total_all = sum(macro['total'] for macro in macro_data.values())
    
    for macro_name in sorted(macro_data.keys(), key=sort_key):
        macro_info = macro_data[macro_name]
        macro_total = macro_info['total']
        macro_percentage = (macro_total / total_all * 100) if total_all else 0.0
        
        # Adicionar linha do macroprocesso
        items.append({
            'type': 'macro',
            'name': macro_name,
            'micro_name': '',
            'total': macro_total,
            'percentage': round(macro_percentage, 2)
        })
            
        # Adicionar microprocessos ordenados
        for micro_name in sorted(macro_info['micros'].keys(), key=sort_key):
            micro_count = macro_info['micros'][micro_name]
            micro_percentage = (micro_count / total_all * 100) if total_all else 0.0
            
            items.append({
                'type': 'micro',
                'name': macro_name,
                'micro_name': micro_name,
                'total': micro_count,
                'percentage': round(micro_percentage, 2)
            })
    
    return {
        'items': items,
        'total_all': total_all
    }

def _build_regularizacao_cteep_etapa_macro_microprocess(sheet_data: dict):
    """Processa dados de Regularização CTEEP com etapas, macroprocessos e microprocessos"""
    headers = sheet_data.get('headers', [])
    rows = sheet_data.get('values', [])

    etapa_idx = None
    macro_idx = None
    micro_idx = None

    for idx, header in enumerate(headers):
        header_upper = header.strip().upper()
        if header_upper == 'ETAPAS':
            etapa_idx = idx
        elif header_upper == 'MACROPROCESSO':
            macro_idx = idx
        elif header_upper == 'MICROPROCESSO':
            micro_idx = idx

    if etapa_idx is None:
        return {
            'items': [],
            'total_all': 0,
            'warning': "Coluna 'Etapas' não encontrada",
            'available_columns': headers
        }

    if macro_idx is None:
        return {
            'items': [],
            'total_all': 0,
            'warning': "Coluna 'Macroprocesso' não encontrada",
            'available_columns': headers
        }

    etapa_data = {}  # {etapa: {macro: {'micros': {micro: count}, 'total': count}}}

    for row in rows:
        max_idx = max(etapa_idx, macro_idx, micro_idx if micro_idx is not None else -1)
        if len(row) <= max_idx:
            continue

        etapa_value = str(row[etapa_idx]).strip() if etapa_idx < len(row) else ''
        macro_value = str(row[macro_idx]).strip() if macro_idx < len(row) else ''
        micro_value = str(row[micro_idx]).strip() if micro_idx is not None and micro_idx < len(row) else ''

        if not etapa_value or etapa_value.lower() in ('nan', 'none', ''):
            continue
        if not macro_value or macro_value.lower() in ('nan', 'none', ''):
            continue

        if etapa_value not in etapa_data:
            etapa_data[etapa_value] = {}
        if macro_value not in etapa_data[etapa_value]:
            etapa_data[etapa_value][macro_value] = {'micros': {}, 'total': 0}

        if micro_value and micro_value.lower() not in ('nan', 'none', ''):
            micros = etapa_data[etapa_value][macro_value]['micros']
            micros[micro_value] = micros.get(micro_value, 0) + 1

        etapa_data[etapa_value][macro_value]['total'] += 1

    def sort_key(name: str):
        match = re.match(r'\s*(\d+(?:\.\d+)*)', name)
        if match:
            parts = match.group(1).split('.')
            return tuple(int(p) for p in parts) + (name,)
        return (9999, name)

    items = []
    total_all = 0
    for etapa_name in etapa_data:
        for macro_name in etapa_data[etapa_name]:
            total_all += etapa_data[etapa_name][macro_name]['total']

    for etapa_name in sorted(etapa_data.keys(), key=sort_key):
        etapa_total = sum(macro_info['total'] for macro_info in etapa_data[etapa_name].values())
        etapa_percentage = (etapa_total / total_all * 100) if total_all else 0.0
        items.append({
            'type': 'etapa',
            'etapa_name': etapa_name,
            'macro_name': '',
            'micro_name': '',
            'total': etapa_total,
            'percentage': round(etapa_percentage, 2)
        })

        for macro_name in sorted(etapa_data[etapa_name].keys(), key=sort_key):
            macro_info = etapa_data[etapa_name][macro_name]
            macro_total = macro_info['total']
            macro_percentage = (macro_total / total_all * 100) if total_all else 0.0
            items.append({
                'type': 'macro',
                'etapa_name': etapa_name,
                'macro_name': macro_name,
                'micro_name': '',
                'total': macro_total,
                'percentage': round(macro_percentage, 2)
            })

            for micro_name in sorted(macro_info['micros'].keys(), key=sort_key):
                micro_count = macro_info['micros'][micro_name]
                micro_percentage = (micro_count / total_all * 100) if total_all else 0.0
                items.append({
                    'type': 'micro',
                    'etapa_name': etapa_name,
                    'macro_name': macro_name,
                    'micro_name': micro_name,
                    'total': micro_count,
                    'percentage': round(micro_percentage, 2)
                })

    return {
        'items': items,
        'total_all': total_all
    }

@reports_bp.route('/regularizacao/sp', methods=['GET'])
@login_required
def get_regularizacao_sp():
    """Retorna dados de Regularização SP (Macroprocesso)"""
    spreadsheet_name = 'Regularizações SP'
    sheet_name = request.args.get('sheet_name', None)
    file_path_obj = _find_enel_spreadsheet_file(spreadsheet_name)
    if not file_path_obj:
        return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404

    sheet_data = read_spreadsheet_file(
        file_path=str(file_path_obj),
        sheet_name=sheet_name,
        header=None
    )
    processed = _build_regularizacao_sp_macroprocess(sheet_data)
    return jsonify(processed), 200

@reports_bp.route('/regularizacao/rj', methods=['GET'])
@login_required
def get_regularizacao_rj():
    """Retorna dados de Regularização RJ (Macroprocesso e Microprocesso)"""
    spreadsheet_name = 'Registral e Notarial - Regularização RJ'
    sheet_name = request.args.get('sheet_name', None)
    file_path_obj = _find_enel_spreadsheet_file(spreadsheet_name)
    if not file_path_obj:
        return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404

    # Ler planilha começando na linha 3 (header_row=2 pois é 0-indexed)
    sheet_data = read_spreadsheet_file(
        file_path=str(file_path_obj),
        sheet_name=sheet_name,
        header=2  # Linha 3 (0-indexed = 2)
    )
    processed = _build_regularizacao_rj_macro_microprocess(sheet_data)
    return jsonify(processed), 200

@reports_bp.route('/regularizacao/cteep', methods=['GET'])
@login_required
def get_regularizacao_cteep():
    """Retorna dados de Regularização CTEEP (Etapas, Macroprocesso e Microprocesso)"""
    spreadsheet_name = 'CTEEP ATUALIZADA - BASE MR 2025'
    sheet_name = request.args.get('sheet_name', None)
    file_path_obj = _find_enel_spreadsheet_file(spreadsheet_name)
    if not file_path_obj:
        return jsonify({'error': f'Planilha não encontrada: {spreadsheet_name}'}), 404

    sheet_data = read_spreadsheet_file(
        file_path=str(file_path_obj),
        sheet_name=sheet_name,
        header=0
    )
    processed = _build_regularizacao_cteep_etapa_macro_microprocess(sheet_data)
    return jsonify(processed), 200

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
    anuencia_comments = []
    certificado_bombeiro_comments = []
    legalizacao_sp_comments = []
    servicos_diversos_sp_comments = []
    legalizacao_rj_comments = []
    legalizacao_rj_bombeiro_comments = []
    regularizacao_sp_comments = []
    regularizacao_rj_comments = []
    regularizacao_cteep_comments = []
    for comment in comments:
        if isinstance(comment, dict):
            page = comment.get('page', '')
            if page == 'Licença Sanitária - Renovação':
                licenca_comments.append(comment)
            elif page == 'Anuência Ambiental':
                anuencia_comments.append(comment)
            elif page == 'Certificado de aprovação Bombeiro':
                certificado_bombeiro_comments.append(comment)
            elif page == 'Alvarás de Funcionamento - Renovação (SP)':
                legalizacao_sp_comments.append(comment)
            elif page == 'Serviços Diversos (SP)':
                servicos_diversos_sp_comments.append(comment)
            elif page == 'Visão Geral - Alvarás de Funcionamento (RJ)':
                legalizacao_rj_comments.append(comment)
            elif page == 'Certificado de Aprovação dos Bombeiros (RJ)':
                legalizacao_rj_bombeiro_comments.append(comment)
            elif page == 'Regularização - SP':
                regularizacao_sp_comments.append(comment)
            elif page == 'Regularização - RJ':
                regularizacao_rj_comments.append(comment)
            elif page == 'Regularização - CTEEP':
                regularizacao_cteep_comments.append(comment)
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
            break
        else:
            logger.debug(f"Logo ENEL não encontrado em: {enel_path_str}")
    
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

    # Fluxograma CTEEP (para última página)
    fluxograma_cteep_path = ''
    fluxograma_cteep_base64 = ''
    try:
        possible_fluxograma_paths = [
            str(ROOT_DIR / 'portal-frontend' / 'images' / 'fluxograma_cteep.png'),
            str(ROOT_DIR.parent / 'portal-frontend' / 'images' / 'fluxograma_cteep.png'),
            str(IMAGES_DIR / 'fluxograma_cteep.png'),
        ]
        for flux_path in possible_fluxograma_paths:
            if os.path.exists(flux_path):
                fluxograma_cteep_path = flux_path
                break
        if fluxograma_cteep_path:
            fluxograma_cteep_base64 = get_image_base64(fluxograma_cteep_path)
    except Exception as e:
        logger.warning(f"Erro ao carregar fluxograma CTEEP: {e}")
    
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
    
    # Função auxiliar para converter chaves de anos
    def convert_years_keys(years_dict):
        """Converte chaves de string para int se necessário"""
        if not years_dict:
            return years_dict
        keys_list = list(years_dict.keys())
        if keys_list and isinstance(keys_list[0], str):
            return {int(k): v for k, v in years_dict.items()}
        return years_dict

    # Buscar dados de Legalização CE se CE estiver na lista
    legalizacao_ce_data = None
    licenca_sanitaria_data = None
    anuencia_ambiental_data = None
    certificado_bombeiro_data = None
    legalizacao_sp_data = None
    legalizacao_sp_servicos_data = None
    legalizacao_rj_data = None
    legalizacao_rj_bombeiro_data = None
    regularizacao_sp_data = None
    
    if 'CE' in legalizacao_lista:
        try:
            from .enel_spreadsheets import get_enel_spreadsheet_data
            
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
            from .enel_spreadsheets import _get_enel_spreadsheet_data_internal
            
            spreadsheet_name_licenca = 'ENEL - Legalização CE'
            filter_natureza_value = 'Renovação Licença Sanitária'
            
            result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name=spreadsheet_name_licenca,
                years=years,
                filter_natureza=filter_natureza_value
            )
            if isinstance(result, tuple) and len(result) > 0:
                if result[1] == 200:
                    licenca_sanitaria_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Licença Sanitária: status {result[1]}")
            elif hasattr(result, 'get_json'):
                licenca_sanitaria_data = result.get_json()
            
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
            
            # 3. Buscar dados de Anuência Ambiental da planilha 'ENEL - Legalização CE'
            # Filtrar apenas registros onde 'Relatório Natureza da Operação' = 'Anuência Ambiental'
            spreadsheet_name_anuencia = 'ENEL - Legalização CE'
            filter_natureza_anuencia = 'Anuência Ambiental'
            result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name=spreadsheet_name_anuencia,
                years=years,
                filter_natureza=filter_natureza_anuencia
            )
            
            if isinstance(result, tuple) and len(result) > 0:
                if result[1] == 200:
                    anuencia_ambiental_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Anuência Ambiental: status {result[1]}")
            elif hasattr(result, 'get_json'):
                anuencia_ambiental_data = result.get_json()
            
            # Converter anos nos dados de Anuência Ambiental
            if anuencia_ambiental_data:
                if anuencia_ambiental_data.get('total_demandado', {}).get('years'):
                    anuencia_ambiental_data['total_demandado']['years'] = convert_years_keys(anuencia_ambiental_data['total_demandado']['years'])
                if anuencia_ambiental_data.get('concluidos', {}).get('years'):
                    anuencia_ambiental_data['concluidos']['years'] = convert_years_keys(anuencia_ambiental_data['concluidos']['years'])
                if anuencia_ambiental_data.get('em_andamento', {}).get('total', {}).get('years'):
                    anuencia_ambiental_data['em_andamento']['total']['years'] = convert_years_keys(anuencia_ambiental_data['em_andamento']['total']['years'])
                for subcat in anuencia_ambiental_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])

            # 4. Buscar dados de Certificado de aprovação Bombeiro da planilha 'ENEL - Legalização CE'
            # Filtrar apenas registros onde 'Relatório Natureza da Operação' = 'Certificado de aprovação Bombeiro'
            spreadsheet_name_bombeiro = 'ENEL - Legalização CE'
            filter_natureza_bombeiro = 'Certificado de aprovação Bombeiro'
            result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name=spreadsheet_name_bombeiro,
                years=years,
                filter_natureza=filter_natureza_bombeiro
            )

            if isinstance(result, tuple) and len(result) > 0:
                if result[1] == 200:
                    certificado_bombeiro_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Certificado de aprovação Bombeiro: status {result[1]}")
            elif hasattr(result, 'get_json'):
                certificado_bombeiro_data = result.get_json()

            # Converter anos nos dados de Certificado de aprovação Bombeiro
            if certificado_bombeiro_data:
                if certificado_bombeiro_data.get('total_demandado', {}).get('years'):
                    certificado_bombeiro_data['total_demandado']['years'] = convert_years_keys(certificado_bombeiro_data['total_demandado']['years'])
                if certificado_bombeiro_data.get('concluidos', {}).get('years'):
                    certificado_bombeiro_data['concluidos']['years'] = convert_years_keys(certificado_bombeiro_data['concluidos']['years'])
                if certificado_bombeiro_data.get('em_andamento', {}).get('total', {}).get('years'):
                    certificado_bombeiro_data['em_andamento']['total']['years'] = convert_years_keys(certificado_bombeiro_data['em_andamento']['total']['years'])
                for subcat in certificado_bombeiro_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Legalização CE: {e}", exc_info=True)

    if 'SP' in legalizacao_lista:
        try:
            from .enel_spreadsheets import _get_enel_spreadsheet_data_internal
            result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name='Legalização SP',
                years=years
            )
            if isinstance(result, tuple) and len(result) > 0:
                if result[1] == 200:
                    legalizacao_sp_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Legalização SP: status {result[1]}")
            elif hasattr(result, 'get_json'):
                legalizacao_sp_data = result.get_json()

            if legalizacao_sp_data:
                if legalizacao_sp_data.get('total_demandado', {}).get('years'):
                    legalizacao_sp_data['total_demandado']['years'] = convert_years_keys(legalizacao_sp_data['total_demandado']['years'])
                if legalizacao_sp_data.get('concluidos', {}).get('years'):
                    legalizacao_sp_data['concluidos']['years'] = convert_years_keys(legalizacao_sp_data['concluidos']['years'])
                if legalizacao_sp_data.get('em_andamento', {}).get('total', {}).get('years'):
                    legalizacao_sp_data['em_andamento']['total']['years'] = convert_years_keys(legalizacao_sp_data['em_andamento']['total']['years'])
                for subcat in legalizacao_sp_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])

            # Serviços Diversos (SP) - aba "MR - Outros Serviços"
            servicos_result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name='Legalização SP',
                years=years,
                sheet_name='MR - Outros Serviços',
                header_row=1,
                item_column='Item',
                item_not_equals='53',
                year_column_name='ano Acionamento',
                concluido_statuses=['Serviços diversos concluídos'],
                status_column_override='Relatório Status detalhado'
            )
            if isinstance(servicos_result, tuple) and len(servicos_result) > 0:
                if servicos_result[1] == 200:
                    legalizacao_sp_servicos_data = servicos_result[0].get_json() if hasattr(servicos_result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Serviços Diversos (SP): status {servicos_result[1]}")
            elif hasattr(servicos_result, 'get_json'):
                legalizacao_sp_servicos_data = servicos_result.get_json()

            if legalizacao_sp_servicos_data:
                if legalizacao_sp_servicos_data.get('total_demandado', {}).get('years'):
                    legalizacao_sp_servicos_data['total_demandado']['years'] = convert_years_keys(legalizacao_sp_servicos_data['total_demandado']['years'])
                if legalizacao_sp_servicos_data.get('concluidos', {}).get('years'):
                    legalizacao_sp_servicos_data['concluidos']['years'] = convert_years_keys(legalizacao_sp_servicos_data['concluidos']['years'])
                if legalizacao_sp_servicos_data.get('em_andamento', {}).get('total', {}).get('years'):
                    legalizacao_sp_servicos_data['em_andamento']['total']['years'] = convert_years_keys(legalizacao_sp_servicos_data['em_andamento']['total']['years'])
                for subcat in legalizacao_sp_servicos_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Legalização SP: {e}", exc_info=True)

    if 'RJ' in legalizacao_lista:
        try:
            from .enel_spreadsheets import _get_enel_spreadsheet_data_internal
            result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name='LEGALIZAÇÃO RJ_28-04',
                years=years,
                sheet_name='Base Alvarás',
                status_column_override='Status detalhado Relatório',
                year_column_name='ano Acionamento',
                year_parse_mode='extract_year',
                concluido_statuses=['Concluído'],
                cancelado_statuses=['Cancelado']
            )
            if isinstance(result, tuple) and len(result) > 0:
                if result[1] == 200:
                    legalizacao_rj_data = result[0].get_json() if hasattr(result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Legalização RJ: status {result[1]}")
            elif hasattr(result, 'get_json'):
                legalizacao_rj_data = result.get_json()

            if legalizacao_rj_data:
                if legalizacao_rj_data.get('total_demandado', {}).get('years'):
                    legalizacao_rj_data['total_demandado']['years'] = convert_years_keys(legalizacao_rj_data['total_demandado']['years'])
                if legalizacao_rj_data.get('concluidos', {}).get('years'):
                    legalizacao_rj_data['concluidos']['years'] = convert_years_keys(legalizacao_rj_data['concluidos']['years'])
                if legalizacao_rj_data.get('cancelados', {}).get('years'):
                    legalizacao_rj_data['cancelados']['years'] = convert_years_keys(legalizacao_rj_data['cancelados']['years'])
                if legalizacao_rj_data.get('em_andamento', {}).get('total', {}).get('years'):
                    legalizacao_rj_data['em_andamento']['total']['years'] = convert_years_keys(legalizacao_rj_data['em_andamento']['total']['years'])
                for subcat in legalizacao_rj_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])

            # Certificado de Aprovação dos Bombeiros (RJ) - aba "Base Bombeiro"
            bombeiro_result = _get_enel_spreadsheet_data_internal(
                spreadsheet_name='LEGALIZAÇÃO RJ_28-04',
                years=years,
                sheet_name='Base Bombeiro',
                status_column_override='Status Geral do imóvel',
                year_column_name='Ano Acionamento',
                year_parse_mode='extract_year',
                concluido_statuses=['CA emitido'],
                status_exclude=['*']
            )
            if isinstance(bombeiro_result, tuple) and len(bombeiro_result) > 0:
                if bombeiro_result[1] == 200:
                    legalizacao_rj_bombeiro_data = bombeiro_result[0].get_json() if hasattr(bombeiro_result[0], 'get_json') else None
                else:
                    logger.warning(f"Erro ao buscar dados de Bombeiros RJ: status {bombeiro_result[1]}")
            elif hasattr(bombeiro_result, 'get_json'):
                legalizacao_rj_bombeiro_data = bombeiro_result.get_json()

            if legalizacao_rj_bombeiro_data:
                if legalizacao_rj_bombeiro_data.get('total_demandado', {}).get('years'):
                    legalizacao_rj_bombeiro_data['total_demandado']['years'] = convert_years_keys(legalizacao_rj_bombeiro_data['total_demandado']['years'])
                if legalizacao_rj_bombeiro_data.get('concluidos', {}).get('years'):
                    legalizacao_rj_bombeiro_data['concluidos']['years'] = convert_years_keys(legalizacao_rj_bombeiro_data['concluidos']['years'])
                if legalizacao_rj_bombeiro_data.get('em_andamento', {}).get('total', {}).get('years'):
                    legalizacao_rj_bombeiro_data['em_andamento']['total']['years'] = convert_years_keys(legalizacao_rj_bombeiro_data['em_andamento']['total']['years'])
                for subcat in legalizacao_rj_bombeiro_data.get('em_andamento', {}).get('subcategorias', []):
                    if subcat.get('years'):
                        subcat['years'] = convert_years_keys(subcat['years'])

                # Separar "Não Iniciado" a partir de status específicos
                nao_iniciado_statuses = ['Aguardando obra Sist. Incêndio - Enel']
                nao_iniciado_norm = {s.strip().lower() for s in nao_iniciado_statuses}
                total_demandado_total = legalizacao_rj_bombeiro_data.get('total_demandado', {}).get('total', 0)

                em_andamento_total = legalizacao_rj_bombeiro_data.get('em_andamento', {}).get('total', {})
                em_andamento_subcats = legalizacao_rj_bombeiro_data.get('em_andamento', {}).get('subcategorias', [])

                nao_iniciado_subcats = []
                remaining_subcats = []
                nao_iniciado_years = {y: 0 for y in years}
                nao_iniciado_total = 0

                for subcat in em_andamento_subcats:
                    subcat_name = str(subcat.get('name', '')).strip()
                    if subcat_name.lower() in nao_iniciado_norm:
                        nao_iniciado_subcats.append(subcat)
                        nao_iniciado_total += subcat.get('total', 0)
                        for year in years:
                            nao_iniciado_years[year] += subcat.get('years', {}).get(year, 0)
                    else:
                        remaining_subcats.append(subcat)

                if nao_iniciado_total > 0:
                    # Atualizar em_andamento total removendo nao_iniciado
                    em_andamento_years = em_andamento_total.get('years', {})
                    for year in years:
                        em_andamento_years[year] = em_andamento_years.get(year, 0) - nao_iniciado_years.get(year, 0)
                    em_andamento_total['years'] = em_andamento_years
                    em_andamento_total['total'] = max(em_andamento_total.get('total', 0) - nao_iniciado_total, 0)

                    legalizacao_rj_bombeiro_data['em_andamento']['total'] = em_andamento_total
                    legalizacao_rj_bombeiro_data['em_andamento']['subcategorias'] = remaining_subcats

                    def calc_pct(value, total):
                        return round((value / total) * 100, 1) if total else 0.0

                    # Recalcular percentuais
                    legalizacao_rj_bombeiro_data['concluidos']['percentage'] = calc_pct(
                        legalizacao_rj_bombeiro_data.get('concluidos', {}).get('total', 0),
                        total_demandado_total
                    )
                    legalizacao_rj_bombeiro_data['em_andamento']['total']['percentage'] = calc_pct(
                        legalizacao_rj_bombeiro_data.get('em_andamento', {}).get('total', {}).get('total', 0),
                        total_demandado_total
                    )
                    for subcat in legalizacao_rj_bombeiro_data['em_andamento'].get('subcategorias', []):
                        subcat['percentage'] = calc_pct(subcat.get('total', 0), total_demandado_total)

                    # Criar estrutura de "Não Iniciado"
                    legalizacao_rj_bombeiro_data['nao_iniciado'] = {
                        'total': nao_iniciado_total,
                        'years': nao_iniciado_years,
                        'percentage': calc_pct(nao_iniciado_total, total_demandado_total),
                        'subcategorias': []
                    }
                    for subcat in nao_iniciado_subcats:
                        legalizacao_rj_bombeiro_data['nao_iniciado']['subcategorias'].append({
                            'name': subcat.get('name', ''),
                            'years': subcat.get('years', {}),
                            'total': subcat.get('total', 0),
                            'percentage': calc_pct(subcat.get('total', 0), total_demandado_total)
                        })
                else:
                    # Garantir estrutura vazia para renderização
                    legalizacao_rj_bombeiro_data['nao_iniciado'] = {
                        'total': 0,
                        'years': {y: 0 for y in years},
                        'percentage': 0.0,
                        'subcategorias': []
                    }
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Legalização RJ: {e}", exc_info=True)
    
    # Buscar dados de Regularização SP se SP estiver na lista
    if 'SP' in regularizacao_lista:
        try:
            file_path_obj = _find_enel_spreadsheet_file('Regularizações SP')
            if file_path_obj:
                sheet_data = read_spreadsheet_file(
                    file_path=str(file_path_obj),
                    sheet_name=None,
                    header=None
                )
                regularizacao_sp_data_dict = _build_regularizacao_sp_macroprocess(sheet_data)
                # Converter dicionário para objeto com atributos para evitar conflito com .items() do dict
                from types import SimpleNamespace
                # Converter cada item da lista também em objeto com atributos
                items_list = []
                max_total = 0
                for item_dict in regularizacao_sp_data_dict.get('items', []):
                    item_total = item_dict.get('total', 0)
                    if item_total > max_total:
                        max_total = item_total
                    items_list.append(SimpleNamespace(
                        name=item_dict.get('name', ''),
                        total=item_total,
                        percentage=item_dict.get('percentage', 0.0)
                    ))
                regularizacao_sp_data = SimpleNamespace(
                    items=items_list,
                    total_all=regularizacao_sp_data_dict.get('total_all', 0),
                    macro_idx=regularizacao_sp_data_dict.get('macro_idx'),
                    max_total=max_total
                )
            else:
                logger.warning("Planilha Regularizações SP não encontrada")
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Regularização SP: {e}", exc_info=True)
    
    # Buscar dados de Regularização RJ se RJ estiver na lista
    regularizacao_rj_data = None
    if 'RJ' in regularizacao_lista:
        try:
            file_path_obj = _find_enel_spreadsheet_file('Registral e Notarial - Regularização RJ')
            if file_path_obj:
                # Ler planilha começando na linha 3 (header_row=2 pois é 0-indexed)
                sheet_data = read_spreadsheet_file(
                    file_path=str(file_path_obj),
                    sheet_name=None,
                    header=2  # Linha 3 (0-indexed = 2)
                )
                regularizacao_rj_data_dict = _build_regularizacao_rj_macro_microprocess(sheet_data)
                # Converter dicionário para objeto com atributos
                from types import SimpleNamespace
                items_list = []
                max_total = 0
                for item_dict in regularizacao_rj_data_dict.get('items', []):
                    # Calcular max_total apenas dos macroprocessos
                    if item_dict.get('type') == 'macro':
                        item_total = item_dict.get('total', 0)
                        if item_total > max_total:
                            max_total = item_total
                    items_list.append(SimpleNamespace(
                        type=item_dict.get('type', ''),
                        name=item_dict.get('name', ''),
                        micro_name=item_dict.get('micro_name', ''),
                        total=item_dict.get('total', 0),
                        percentage=item_dict.get('percentage', 0.0)
                    ))
                regularizacao_rj_data = SimpleNamespace(
                    items=items_list,
                    total_all=regularizacao_rj_data_dict.get('total_all', 0),
                    max_total=max_total
                )
            else:
                logger.warning("Planilha Registral e Notarial - Regularização RJ não encontrada")
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Regularização RJ: {e}", exc_info=True)

    # Buscar dados de Regularização CTEEP se CTEEP estiver na lista
    regularizacao_cteep_data = None
    if 'CTEEP' in regularizacao_lista:
        try:
            file_path_obj = _find_enel_spreadsheet_file('CTEEP ATUALIZADA - BASE MR 2025')
            if file_path_obj:
                sheet_data = read_spreadsheet_file(
                    file_path=str(file_path_obj),
                    sheet_name=None,
                    header=0
                )
                regularizacao_cteep_data_dict = _build_regularizacao_cteep_etapa_macro_microprocess(sheet_data)
                from types import SimpleNamespace
                items_list = []
                max_total = 0
                for item_dict in regularizacao_cteep_data_dict.get('items', []):
                    if item_dict.get('type') == 'macro':
                        item_total = item_dict.get('total', 0)
                        if item_total > max_total:
                            max_total = item_total
                    items_list.append(SimpleNamespace(
                        type=item_dict.get('type', ''),
                        etapa_name=item_dict.get('etapa_name', ''),
                        macro_name=item_dict.get('macro_name', ''),
                        micro_name=item_dict.get('micro_name', ''),
                        total=item_dict.get('total', 0),
                        percentage=item_dict.get('percentage', 0.0)
                    ))
                regularizacao_cteep_data = SimpleNamespace(
                    items=items_list,
                    total_all=regularizacao_cteep_data_dict.get('total_all', 0),
                    max_total=max_total
                )
            else:
                logger.warning("Planilha CTEEP ATUALIZADA - BASE MR 2025 não encontrada")
        except Exception as e:
            logger.error(f"Erro ao buscar dados de Regularização CTEEP: {e}", exc_info=True)
    
    # Renderizar template HTML
    # #region agent log
    log_dir = Path('.cursor')
    log_dir.mkdir(exist_ok=True)
    try:
        with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
            f.write(json.dumps({
                'sessionId': 'debug-session',
                'runId': 'run1',
                'hypothesisId': 'A',
                'location': 'reports.py:819',
                'message': 'Before render_template',
                'data': {
                    'regularizacao_lista': regularizacao_lista,
                    'regularizacao_sp_data_present': regularizacao_sp_data is not None,
                    'regularizacao_sp_data_items_count': len(regularizacao_sp_data.get('items', [])) if regularizacao_sp_data else 0
                },
                'timestamp': int(datetime.now().timestamp() * 1000)
            }) + '\n')
    except Exception:
        pass
    # #endregion
    
    try:
        html_content = render_template(
            'report_pdf.html',
            client_name=client_dict['nome'],
            month_name=month_name,
            year=ano,
            estados=estados_str,
            estados_lista=estados_lista,
            legalizacao_lista=legalizacao_lista,
            legalizacao_ce_data=legalizacao_ce_data,
            legalizacao_sp_data=legalizacao_sp_data,
            legalizacao_sp_servicos_data=legalizacao_sp_servicos_data,
            legalizacao_rj_data=legalizacao_rj_data,
            legalizacao_rj_bombeiro_data=legalizacao_rj_bombeiro_data,
            regularizacao_lista=regularizacao_lista,
            regularizacao_sp_data=regularizacao_sp_data,
            regularizacao_rj_data=regularizacao_rj_data,
            regularizacao_cteep_data=regularizacao_cteep_data,
            licenca_sanitaria_data=licenca_sanitaria_data,
            anuencia_ambiental_data=anuencia_ambiental_data,
            certificado_bombeiro_data=certificado_bombeiro_data,
            years=years,
            comments=comments,
            alvaras_comments=alvaras_comments,
            licenca_comments=licenca_comments,
            anuencia_comments=anuencia_comments,
            certificado_bombeiro_comments=certificado_bombeiro_comments,
            legalizacao_sp_comments=legalizacao_sp_comments,
            servicos_diversos_sp_comments=servicos_diversos_sp_comments,
            legalizacao_rj_comments=legalizacao_rj_comments,
            legalizacao_rj_bombeiro_comments=legalizacao_rj_bombeiro_comments,
            regularizacao_sp_comments=regularizacao_sp_comments,
            regularizacao_rj_comments=regularizacao_rj_comments,
            regularizacao_cteep_comments=regularizacao_cteep_comments,
            mr_logo_path=mr_logo_base64,
            client_logo_path=client_logo_base64,
            fluxograma_cteep_path=fluxograma_cteep_base64
    )
        # #region agent log
        try:
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'B',
                    'location': 'reports.py:870',
                    'message': 'Template rendered successfully',
                    'data': {
                        'html_length': len(html_content),
                        'html_preview': html_content[:200]
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
        except Exception:
            pass
        # #endregion
    except Exception as e:
        logger.error(f"Erro ao renderizar template: {e}", exc_info=True)
        # #region agent log
        try:
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'C',
                    'location': 'reports.py:render_template_error',
                    'message': 'Template render error',
                    'data': {
                        'error': str(e),
                        'error_type': type(e).__name__
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
        except Exception:
            pass
        # #endregion
        return jsonify({'error': f'Erro ao renderizar template: {str(e)}'}), 500
    
    # Log do HTML gerado (primeiros 500 caracteres para debug)
    logger.info(f"HTML gerado (primeiros 500 chars): {html_content[:500]}")
    
    # Gerar PDF com WeasyPrint
    try:
        # #region agent log
        try:
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'D',
                    'location': 'reports.py:before_weasyprint',
                    'message': 'Before WeasyPrint PDF generation',
                    'data': {
                        'html_length': len(html_content),
                        'images_dir_exists': images_dir.exists() if images_dir else False
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
        except Exception:
            pass
        # #endregion
        
        font_config = FontConfiguration()

        # base_url não é mais necessário pois imagens são base64
        pdf_bytes = HTML(string=html_content, base_url=str(images_dir) if images_dir.exists() else None).write_pdf(
            font_config=font_config
        )

        # #region agent log
        try:
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'E',
                    'location': 'reports.py:after_weasyprint',
                    'message': 'WeasyPrint PDF generated successfully',
                    'data': {
                        'pdf_bytes_length': len(pdf_bytes)
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
        except Exception:
            pass
        # #endregion
        
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
        # #region agent log
        try:
            with open('.cursor/debug.log', 'a', encoding='utf-8') as f:
                f.write(json.dumps({
                    'sessionId': 'debug-session',
                    'runId': 'run1',
                    'hypothesisId': 'F',
                    'location': 'reports.py:weasyprint_exception',
                    'message': 'WeasyPrint PDF generation exception',
                    'data': {
                        'error': str(e),
                        'error_type': type(e).__name__,
                        'traceback': traceback.format_exc()
                    },
                    'timestamp': int(datetime.now().timestamp() * 1000)
                }) + '\n')
        except Exception:
            pass
        # #endregion
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500

