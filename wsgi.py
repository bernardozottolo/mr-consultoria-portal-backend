"""
WSGI entry point para Gunicorn
"""
from api.app import app

if __name__ == "__main__":
    app.run()
