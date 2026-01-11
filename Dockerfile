FROM python:3.12-slim

WORKDIR /app

# Instalar dependências do sistema necessárias para WeasyPrint
RUN apt-get update && apt-get install -y \
    libpango-1.0-0 \
    libharfbuzz0b \
    libpangoft2-1.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements e instalar dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código da aplicação
COPY . .

# Garantir que PYTHONPATH inclui /app
ENV PYTHONPATH=/app

# Criar diretórios necessários
RUN mkdir -p /app/data /app/logs

# Expor porta 5000
EXPOSE 5000

# Comando para iniciar Gunicorn
# IMPORTANTE: Bind em 0.0.0.0:5000 para aceitar conexões de outros containers
CMD ["gunicorn", "-b", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "--access-logfile", "-", "--error-logfile", "-", "wsgi:app"]
