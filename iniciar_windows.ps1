Set-Location $PSScriptRoot
if (!(Test-Path .venv)) {
  Write-Host "Criando ambiente virtual..."
  py -m venv .venv
}
. .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if (!(Test-Path .env)) { Copy-Item .env.example .env }
if (!(Test-Path data)) { New-Item -ItemType Directory -Path data | Out-Null }
Start-Process "http://127.0.0.1:8000"
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
