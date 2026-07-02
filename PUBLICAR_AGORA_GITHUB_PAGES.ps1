param(
    [string]$Repositorio
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "ORA NFS-e Automatico - Publicar interface no GitHub Pages" -ForegroundColor Cyan
Write-Host ""

try {
    git --version | Out-Null
} catch {
    Write-Host "ERRO: Git nao foi encontrado no Windows." -ForegroundColor Red
    exit 1
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptDir = (Resolve-Path $scriptDir).Path

Write-Host "Pasta onde este assistente esta:"
Write-Host $scriptDir
Write-Host ""

if (Test-Path (Join-Path $scriptDir ".git")) {
    $repoDir = $scriptDir
} else {
    Write-Host "Esta pasta tem os arquivos do sistema, mas NAO tem .git." -ForegroundColor Yellow
    Write-Host "Por isso o erro anterior apareceu: fatal: not a git repository."
    Write-Host ""
    if (-not $Repositorio) {
        $Repositorio = Read-Host "Cole aqui a pasta REAL do repositorio local"
    }
    $Repositorio = $Repositorio.Trim('"')
    $repoDir = (Resolve-Path $Repositorio).Path

    if (-not (Test-Path (Join-Path $repoDir ".git"))) {
        Write-Host "ERRO: esta pasta tambem nao tem .git: $repoDir" -ForegroundColor Red
        Write-Host "No GitHub Desktop, use Repository > Show in Explorer para achar a pasta correta."
        exit 1
    }

    Write-Host "Copiando arquivos para a raiz do repositorio..."
    robocopy $scriptDir $repoDir /E /XD ".git" "data" "logs" "cache" ".venv" "venv" "__pycache__" /XF ".env" "*.pfx" "*.p12" "*.pem" "*.key" "*.pyc" "*.sqlite" "*.db" | Out-Null
    if ($LASTEXITCODE -ge 8) {
        throw "Falha ao copiar arquivos para o repositorio."
    }
}

Set-Location $repoDir

if (-not (Test-Path "index.html")) {
    throw "index.html nao esta na raiz do repositorio."
}
if (-not (Test-Path ".nojekyll")) {
    New-Item -ItemType File -Path ".nojekyll" | Out-Null
}

Write-Host "OK: index.html encontrado."
Write-Host "OK: .nojekyll encontrado."
Write-Host ""

git status --short

git add index.html .nojekyll README.md PUBLICAR_GITHUB_PAGES.md GITHUB_SEM_BANCO.md requirements.txt .env.example .gitignore iniciar_windows.bat iniciar_windows_rede.bat iniciar_windows.ps1 iniciar_github_pages_windows.bat iniciar_github_pages_windows.ps1 corrigir_github_pages_windows.bat corrigir_github_pages_windows.ps1 PUBLICAR_AGORA_GITHUB_PAGES.bat app docs samples tests

git diff --cached --quiet
if ($LASTEXITCODE -eq 1) {
    git commit -m "Publica interface do sistema no GitHub Pages"
} else {
    Write-Host "Nenhuma alteracao nova para commit."
}

git push

Write-Host ""
Write-Host "Publicacao enviada com sucesso." -ForegroundColor Green
Write-Host "Aguarde 1 a 5 minutos e acesse:"
Write-Host "https://victorgadiola-dev.github.io/ora-nfse-automatico/?v=13"
Write-Host "Use Ctrl + F5 para forcar atualizacao."
