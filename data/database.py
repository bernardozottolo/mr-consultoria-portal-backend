import sqlite3
import os
from pathlib import Path

# Remover sys.path.insert e imports de repo.api.config
DB_PATH = os.environ.get('DB_PATH')
if not DB_PATH:
    # Default relativo à raiz do portal-backend
    root_dir = Path(__file__).parent.parent
    DB_PATH = str(root_dir / 'data' / 'database.db')
# Garantir path absoluto
DB_PATH = os.path.abspath(DB_PATH)

def get_db_connection():
    """
    Retorna uma conexão com o banco de dados SQLite
    Configurado com WAL mode e timeouts para suportar multi-processo (Gunicorn)
    """
    conn = sqlite3.connect(DB_PATH, timeout=5.0)
    conn.row_factory = sqlite3.Row
    
    # Configurar PRAGMAs para produção (WAL mode, timeouts, foreign keys)
    conn.execute('PRAGMA journal_mode=WAL;')  # Write-Ahead Logging para multi-processo
    conn.execute('PRAGMA busy_timeout=5000;')  # Timeout de 5 segundos para locks
    conn.execute('PRAGMA foreign_keys=ON;')    # Habilitar foreign keys
    
    return conn

def init_database():
    """Inicializa o banco de dados com as tabelas necessárias"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Tabela de usuários
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            senha_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            totp_secret TEXT NOT NULL
        )
    ''')
    
    # Tabela de clientes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS clients (
            id TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            logo_path TEXT NOT NULL
        )
    ''')
    
    # Tabela de planilhas enviadas (por regional)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS spreadsheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            regional TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            sheet_name TEXT,
            status_column TEXT,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(regional)
        )
    ''')
    
    # Tabela de planilhas específicas do Enel
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS enel_spreadsheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spreadsheet_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            sheet_name TEXT,
            status_column TEXT,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(spreadsheet_name)
        )
    ''')
    
    # Inserir cliente ENEL se não existir
    cursor.execute('''
        INSERT OR IGNORE INTO clients (id, nome, logo_path)
        VALUES ('enel', 'ENEL', 'images/enel-logo.png')
    ''')
    
    conn.commit()
    conn.close()

def reset_database():
    """Remove todas as tabelas (usado para testes)"""
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    # Também remover arquivos WAL e SHM
    wal_path = DB_PATH + '-wal'
    shm_path = DB_PATH + '-shm'
    if os.path.exists(wal_path):
        os.remove(wal_path)
    if os.path.exists(shm_path):
        os.remove(shm_path)
    init_database()

if __name__ == '__main__':
    init_database()
    print("Database initialized successfully!")
    print(f"DB_PATH: {DB_PATH}")
