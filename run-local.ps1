$ErrorActionPreference = "Stop"

Write-Host "== Building frontend to backend/web =="
Push-Location "frontend"
npm install
npm run build
Pop-Location

Write-Host "== Starting FastAPI over HTTPS on https://localhost:8443 =="
# If you use a venv, activate it before running this script or adjust accordingly.
python -m uvicorn backend.main:app `
  --host 127.0.0.1 `
  --port 8443 `
  --ssl-keyfile "backend/certs/localhost+2-key.pem" `
  --ssl-certfile "backend/certs/localhost+2.pem"
