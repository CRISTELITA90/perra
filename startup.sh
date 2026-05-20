#!/usr/bin/env bash
set -euxo pipefail

cd /home/site/wwwroot

# Instala siempre en una ruta escribible fuera de wwwroot (evita problemas de run-from-package)
python -m pip install --upgrade pip
python -m pip install --no-cache-dir --target=/home/site/.python_packages/lib/site-packages -r requirements.txt

export PYTHONPATH=/home/site/.python_packages/lib/site-packages

# Arranque con logs a consola (para ver el error real en logs)
exec python -m gunicorn --chdir /home/site/wwwroot \
  -k uvicorn.workers.UvicornWorker main:app \
  --bind 0.0.0.0:8000 \
  --workers 1 \
  --timeout 180 \
  --log-level debug \
  --access-logfile - \
  --error-logfile -
