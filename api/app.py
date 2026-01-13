from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from .auth import generate_token, verify_token, get_token_from_request
from .users import users_bp
from .reports import reports_bp
from .spreadsheets import spreadsheets_bp
from .enel_spreadsheets import enel_spreadsheets_bp
from .config import (
    DEBUG, IS_PRODUCTION, ENABLE_CORS, CORS_ORIGINS,
    FLASK_HOST, FLASK_PORT, JWT_SECRET, ROOT_DIR
)
import os
import pyotp
import logging
from pathlib import Path
from datetime import datetime

# Remover sys.path.insert - usar imports normais
from data import users_db, database

# Remover static_folder (backend não serve frontend)
app = Flask(__name__, 
            template_folder=os.path.join(os.path.dirname(__file__), 'templates'))

# Configurar logging
log_dir = ROOT_DIR / 'logs'
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f'app_{datetime.now().strftime("%Y%m%d")}.log'

logging.basicConfig(
    level=logging.DEBUG if DEBUG else logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
logger.info(f"Logging configurado. Arquivo: {log_file}")

# Configurar CORS baseado em variáveis de ambiente
if ENABLE_CORS:
    if CORS_ORIGINS == ['*']:
        CORS(app)  # Permitir todas as origens (desenvolvimento)
    else:
        CORS(app, origins=CORS_ORIGINS)  # Origens específicas (produção)
else:
    # CORS desabilitado (produção padrão quando frontend está na mesma origem)
    pass

# Registrar blueprints
app.register_blueprint(users_bp)
app.register_blueprint(reports_bp)
app.register_blueprint(spreadsheets_bp)
app.register_blueprint(enel_spreadsheets_bp)
logger.info("Blueprints registrados: users_bp, reports_bp, spreadsheets_bp, enel_spreadsheets_bp")
with app.app_context():
    routes = [str(rule) for rule in app.url_map.iter_rules()]
    logger.info(f"Rotas disponíveis ({len(routes)}): {routes}")
    # Verificar especificamente a rota de PDF
    pdf_routes = [r for r in routes if 'pdf' in r.lower()]
    logger.info(f"Rotas de PDF: {pdf_routes}")

# Inicializar banco de dados
database.init_database()

@app.route('/api/auth/login', methods=['POST'])
def login():
    """Endpoint de login"""
    data = request.json
    email = data.get('email')
    senha = data.get('senha')
    
    if not email or not senha:
        return jsonify({'error': 'Email e senha são obrigatórios'}), 400
    
    # Verificar credenciais
    if users_db.verify_password(email, senha):
        user = users_db.get_user_by_email(email)
        token = generate_token(email, user['role'])
        return jsonify({
            'token': token,
            'user': {
                'email': user['email'],
                'nome': user['nome'],
                'role': user['role']
            }
        })
    else:
        return jsonify({'error': 'Credenciais inválidas'}), 401

@app.route('/api/auth/forgot-password', methods=['POST'])
def forgot_password():
    """Endpoint para iniciar recuperação de senha (retorna TOTP necessário)"""
    data = request.json
    email = data.get('email')
    
    if not email:
        return jsonify({'error': 'Email é obrigatório'}), 400
    
    user = users_db.get_user_by_email(email)
    if not user:
        # Por segurança, não revelar se o email existe ou não
        return jsonify({'message': 'Se o email existir, você receberá instruções'})
    
    # Retornar que precisa do código TOTP
    return jsonify({
        'message': 'Forneça o código TOTP de 6 dígitos',
        'requires_totp': True
    })

@app.route('/api/auth/reset-password', methods=['POST'])
def reset_password():
    """Endpoint para resetar senha com código TOTP"""
    data = request.json
    email = data.get('email')
    totp_code = data.get('totp_code')
    new_password = data.get('new_password')
    
    if not all([email, totp_code, new_password]):
        return jsonify({'error': 'Dados incompletos'}), 400
    
    user = users_db.get_user_by_email(email)
    if not user:
        return jsonify({'error': 'Usuário não encontrado'}), 404
    
    # Verificar TOTP
    totp = pyotp.TOTP(user['totp_secret'])
    if not totp.verify(totp_code, valid_window=1):
        return jsonify({'error': 'Código TOTP inválido'}), 400
    
    # Atualizar senha
    success = users_db.update_user(email, senha=new_password)
    if success:
        return jsonify({'message': 'Senha alterada com sucesso'})
    else:
        return jsonify({'error': 'Erro ao alterar senha'}), 500

@app.route('/api/health', methods=['GET'])
def health():
    """Endpoint de health check"""
    return jsonify({'status': 'ok'})

# Rotas estáticas removidas - frontend é servido pelo nginx no portal-frontend
# Em produção, backend serve apenas API

if __name__ == '__main__':
    app.run(debug=DEBUG, host=FLASK_HOST, port=FLASK_PORT)
