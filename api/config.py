"""
Configurações centralizadas do projeto
Lê variáveis de ambiente com fallbacks seguros para desenvolvimento
"""
import os
from pathlib import Path

# Diretório raiz do projeto (assumindo que config.py está em api/)
ROOT_DIR = Path(__file__).parent.parent  # api/config.py → portal-backend/

# Ambiente
FLASK_ENV = os.environ.get('FLASK_ENV', 'development')
DEBUG = os.environ.get('FLASK_DEBUG', 'True').lower() in ('true', '1', 'yes')
IS_PRODUCTION = FLASK_ENV == 'production'

# JWT Secret - CRÍTICO: Deve estar definido em produção
JWT_SECRET = os.environ.get('JWT_SECRET')
if not JWT_SECRET:
    if IS_PRODUCTION:
        raise ValueError(
            "JWT_SECRET deve ser definido em produção! "
            "Defina a variável de ambiente JWT_SECRET."
        )
    else:
        JWT_SECRET = 'dev-secret-key-change-in-production'
        print("⚠️  AVISO: JWT_SECRET usando valor padrão inseguro. Defina JWT_SECRET em produção!")

JWT_ALGORITHM = 'HS256'
JWT_EXPIRATION_HOURS = int(os.environ.get('JWT_EXPIRATION_HOURS', '24'))

# Database
DB_PATH = os.environ.get('DB_PATH')
if not DB_PATH:
    # Default: data/database.db (relativo ao root do projeto)
    DB_PATH = str(ROOT_DIR / 'data' / 'database.db')
else:
    # Garantir path absoluto
    DB_PATH = os.path.abspath(DB_PATH)

# CORS - Controlado por variáveis de ambiente
ENABLE_CORS = os.environ.get('ENABLE_CORS', 'true' if not IS_PRODUCTION else 'false').lower() in ('true', '1', 'yes')
CORS_ORIGINS = os.environ.get('CORS_ORIGINS', '*')
if CORS_ORIGINS != '*' and CORS_ORIGINS:
    # Converter string separada por vírgulas em lista
    CORS_ORIGINS = [origin.strip() for origin in CORS_ORIGINS.split(',')]
else:
    CORS_ORIGINS = ['*'] if ENABLE_CORS else []

# Flask host/port (para desenvolvimento)
FLASK_HOST = os.environ.get('FLASK_HOST', '127.0.0.1' if IS_PRODUCTION else '0.0.0.0')
FLASK_PORT = int(os.environ.get('FLASK_PORT', '5000'))

# Google Sheets - Padronizado para GOOGLE_SERVICE_ACCOUNT_FILE
GOOGLE_SERVICE_ACCOUNT_FILE = os.environ.get(
    'GOOGLE_SERVICE_ACCOUNT_FILE',
    '/run/secrets/google.json'
)

# Mantido para compatibilidade (deprecated)
GOOGLE_SHEETS_SPREADSHEET_ID = os.environ.get('GOOGLE_SHEETS_SPREADSHEET_ID', '')
GOOGLE_SHEETS_CREDENTIALS_PATH = GOOGLE_SERVICE_ACCOUNT_FILE  # Alias para compatibilidade

# Report Configuration
REPORT_CONFIG_PATH = os.environ.get(
    'REPORT_CONFIG_PATH',
    str(ROOT_DIR / 'config' / 'report_config.json')
)

# Logging
LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO' if IS_PRODUCTION else 'DEBUG')

# Paths para assets (para geração de PDF)
IMAGES_DIR = ROOT_DIR / 'assets' / 'images'
TEMPLATES_DIR = ROOT_DIR / 'api' / 'templates'

# Diretório para armazenar planilhas enviadas
SPREADSHEETS_DIR = ROOT_DIR / 'data' / 'spreadsheets'
SPREADSHEETS_DIR.mkdir(parents=True, exist_ok=True)
