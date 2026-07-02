Set-Location $PSScriptRoot

$pagesUrl = "https://victorgadiola-dev.github.io/ora-nfse-automatico/"
if (Test-Path .env) {
  Get-Content .env | ForEach-Object {
    if ($_ -match "^GITHUB_PAGES_URL=(.+)$") {
      $pagesUrl = $Matches[1].Trim()
    }
  }
}

if (!(Test-Path .venv)) {
  Write-Host "Criando ambiente virtual..."
  py -m venv .venv
}

. .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if (!(Test-Path .env)) { Copy-Item .env.example .env }
if (!(Test-Path data)) { New-Item -ItemType Directory -Path data | Out-Null }

Write-Host ""
Write-Host "Abrindo interface publica ORA em: $pagesUrl"
Write-Host "Agente local: http://127.0.0.1:8000"
Write-Host "Nao feche esta janela enquanto estiver usando o sistema."
Write-Host ""

Start-Process $pagesUrl
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
