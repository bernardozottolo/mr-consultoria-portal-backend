from .database import get_db_connection
import bcrypt
import sqlite3
import re

def normalize_email(email):
    """
    Normaliza email: trim + lowercase
    Valida formato básico
    """
    if not email:
        return None
    email = email.strip().lower()
    # Validação básica de formato
    if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
        raise ValueError('Email com formato inválido')
    return email

def create_user(email, nome, senha, role, totp_secret):
    """Cria um novo usuário no banco de dados"""
    email = normalize_email(email)
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Hash da senha
    senha_hash = bcrypt.hashpw(senha.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    try:
        cursor.execute('''
            INSERT INTO users (email, nome, senha_hash, role, totp_secret)
            VALUES (?, ?, ?, ?, ?)
        ''', (email, nome, senha_hash, role, totp_secret))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def get_user_by_email(email):
    """Retorna um usuário pelo email (email normalizado)"""
    email = normalize_email(email)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM users WHERE email = ?', (email,))
    user = cursor.fetchone()
    conn.close()
    return dict(user) if user else None

def get_all_users():
    """Retorna todos os usuários"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT email, nome, role FROM users')
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users

def update_user(email, nome=None, senha=None, role=None, totp_secret=None):
    """Atualiza um usuário (email normalizado)"""
    email = normalize_email(email)
    conn = get_db_connection()
    cursor = conn.cursor()
    
    updates = []
    params = []
    
    if nome is not None:
        updates.append('nome = ?')
        params.append(nome)
    if senha is not None:
        senha_hash = bcrypt.hashpw(senha.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        updates.append('senha_hash = ?')
        params.append(senha_hash)
    if role is not None:
        updates.append('role = ?')
        params.append(role)
    if totp_secret is not None:
        updates.append('totp_secret = ?')
        params.append(totp_secret)
    
    if not updates:
        conn.close()
        return False
    
    params.append(email)
    query = f'UPDATE users SET {", ".join(updates)} WHERE email = ?'
    cursor.execute(query, params)
    conn.commit()
    success = cursor.rowcount > 0
    conn.close()
    return success

def delete_user(email):
    """Remove um usuário (email normalizado)"""
    email = normalize_email(email)
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM users WHERE email = ?', (email,))
    conn.commit()
    success = cursor.rowcount > 0
    conn.close()
    return success

def verify_password(email, senha):
    """Verifica se a senha está correta (email normalizado)"""
    email = normalize_email(email)
    user = get_user_by_email(email)
    if not user:
        return False
    return bcrypt.checkpw(senha.encode('utf-8'), user['senha_hash'].encode('utf-8'))
