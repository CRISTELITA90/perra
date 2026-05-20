#!/usr/bin/env bash
set -euxo pipefail

cd /home/site/wwwroot

python -m pip install --upgrade pip
python -m pip install --no-cache-dir --target=/home/site/.python_packages/lib/site-packages -r requirements.txt
export PYTHONPATH=/home/site/.python_packages/lib/site-packages

# Arranque con logs a consola (se verán en logs de App Service)
exec python -m gunicorn -k uvicorn.workers.UvicornWorker main:app \
  --bind 0.0.0.0:8000 \
  --workers 1 \
  --timeout 180 \
  --log-level debug \
  --access-logfile - \
  --error-logfile -
