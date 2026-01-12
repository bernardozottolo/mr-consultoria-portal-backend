import jwt
import datetime
from functools import wraps
from flask import request, jsonify
from .config import JWT_SECRET, JWT_ALGORITHM, JWT_EXPIRATION_HOURS

def generate_token(email, role):
    """Gera um token JWT para o usuário"""
    payload = {
        'email': email,
        'role': role,
        'exp': datetime.datetime.utcnow() + datetime.timedelta(hours=JWT_EXPIRATION_HOURS),
        'iat': datetime.datetime.utcnow()
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)

def verify_token(token):
    """Verifica e decodifica um token JWT"""
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        return payload
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None

def get_token_from_request():
    """Extrai o token JWT do header Authorization"""
    auth_header = request.headers.get('Authorization')
    if not auth_header:
        return None
    try:
        token = auth_header.split(' ')[1]  # Formato: "Bearer <token>"
        return token
    except IndexError:
        return None

def login_required(f):
    """Decorator para rotas que requerem autenticação"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        token = get_token_from_request()
        if not token:
            return jsonify({'error': 'Token não fornecido'}), 401
        
        payload = verify_token(token)
        if not payload:
            return jsonify({'error': 'Token inválido ou expirado'}), 401
        
        # Adicionar informações do usuário ao request
        request.current_user = payload
        return f(*args, **kwargs)
    return decorated_function

def role_required(required_role):
    """Decorator para rotas que requerem um role específico"""
    def decorator(f):
        @wraps(f)
        @login_required
        def decorated_function(*args, **kwargs):
            if request.current_user.get('role') != required_role:
                return jsonify({'error': 'Acesso negado. Role insuficiente.'}), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator

