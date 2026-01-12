from flask import Blueprint, request, jsonify
from .auth import login_required, role_required
# Remover sys.path.insert - usar imports normais
from data import users_db
import pyotp

users_bp = Blueprint('users', __name__)

@users_bp.route('/api/users', methods=['GET'])
@login_required
@role_required('dev-master')
def get_users():
    """Lista todos os usuários (apenas dev-master)"""
    users = users_db.get_all_users()
    return jsonify({'users': users})

@users_bp.route('/api/users', methods=['POST'])
@login_required
@role_required('dev-master')
def create_user():
    """Cria um novo usuário (apenas dev-master)"""
    data = request.json
    email = data.get('email')
    nome = data.get('nome')
    senha = data.get('senha')
    role = data.get('role')
    
    if not all([email, nome, senha, role]):
        return jsonify({'error': 'Dados incompletos'}), 400
    
    # Gerar TOTP secret
    totp_secret = pyotp.random_base32()
    
    success = users_db.create_user(email, nome, senha, role, totp_secret)
    if success:
        return jsonify({'message': 'Usuário criado com sucesso', 'totp_secret': totp_secret}), 201
    else:
        return jsonify({'error': 'Email já existe'}), 400

@users_bp.route('/api/users/<email>', methods=['GET'])
@login_required
@role_required('dev-master')
def get_user(email):
    """Obtém um usuário específico (apenas dev-master)"""
    user = users_db.get_user_by_email(email)
    if not user:
        return jsonify({'error': 'Usuário não encontrado'}), 404
    
    # Não retornar senha_hash
    user_data = {
        'email': user['email'],
        'nome': user['nome'],
        'role': user['role']
    }
    return jsonify(user_data)

@users_bp.route('/api/users/<email>', methods=['PUT'])
@login_required
@role_required('dev-master')
def update_user(email):
    """Atualiza um usuário (apenas dev-master)"""
    data = request.json
    nome = data.get('nome')
    senha = data.get('senha')
    role = data.get('role')
    
    success = users_db.update_user(email, nome=nome, senha=senha, role=role)
    if success:
        return jsonify({'message': 'Usuário atualizado com sucesso'})
    else:
        return jsonify({'error': 'Usuário não encontrado'}), 404

@users_bp.route('/api/users/<email>', methods=['DELETE'])
@login_required
@role_required('dev-master')
def delete_user(email):
    """Remove um usuário (apenas dev-master)"""
    success = users_db.delete_user(email)
    if success:
        return jsonify({'message': 'Usuário removido com sucesso'})
    else:
        return jsonify({'error': 'Usuário não encontrado'}), 404

