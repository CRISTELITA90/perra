#!/usr/bin/env bash
set -euo pipefail

cd /home/site/wwwroot

# Limpiar caché Python para forzar recompilación
find . -name "*.pyc" -delete 2>/dev/null || true
find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true

# Instala dependencias en una ruta escribible y estándar para App Service
python -m pip install --upgrade pip
python -m pip install --no-cache-dir --target=/home/site/.python_packages/lib/site-packages -r requirements.txt

export PYTHONPATH=/home/site/.python_packages/lib/site-packages

# Arranca con logs a stdout/stderr (para que queden en logs)
exec python -m gunicorn --chdir /home/site/wwwroot \
  -k uvicorn.workers.UvicornWorker main:app \
  --bind 0.0.0.0:${PORT:-8000} \
  --workers 1 \
  --timeout 180 \
  --log-level debug \
  --access-logfile - \
  --error-logfile -
