@echo off
chcp 65001 >nul
title ORA - Publicar no Render

echo ==========================================================
echo ORA NFS-e Automatico - Publicar operacao online no Render
echo ==========================================================
echo.
echo Este assistente copia esta versao para a pasta REAL do
echo repositorio GitHub e remove arquivos antigos de GitHub Pages
echo e inicializadores locais que nao sao usados em producao.
echo.

set "SRC=%~dp0"
set /p "REPO=Arraste ou cole aqui a pasta do repositorio ora-nfse-automatico e pressione ENTER: "
set "REPO=%REPO:"=%"

if not exist "%REPO%\.git" (
  echo.
  echo ERRO: a pasta informada nao tem .git.
  echo Abra o GitHub Desktop, clique em "Show in Explorer" e use aquela pasta.
  pause
  exit /b 1
)

echo.
echo Copiando arquivos para o repositorio...
robocopy "%SRC%" "%REPO%" /E /XD ".git" ".venv" "venv" "data" "__pycache__" ".pytest_cache" /XF ".env" "*.pfx" "*.p12" "*.pem" "*.key" "*.cer" "*.crt" >nul

echo.
echo Removendo arquivos antigos ou locais, se existirem...
del "%REPO%\index.html" 2>nul
del "%REPO%\.nojekyll" 2>nul
del "%REPO%\PUBLICAR_GITHUB_PAGES.md" 2>nul
del "%REPO%\PUBLICAR_AGORA_GITHUB_PAGES.bat" 2>nul
del "%REPO%\PUBLICAR_AGORA_GITHUB_PAGES.ps1" 2>nul
del "%REPO%\corrigir_github_pages_windows.bat" 2>nul
del "%REPO%\corrigir_github_pages_windows.ps1" 2>nul
del "%REPO%\iniciar_github_pages_windows.bat" 2>nul
del "%REPO%\iniciar_github_pages_windows.ps1" 2>nul
del "%REPO%\iniciar_windows.bat" 2>nul
del "%REPO%\iniciar_windows.ps1" 2>nul
del "%REPO%\iniciar_windows_rede.bat" 2>nul

echo.
echo Publicando no Git...
cd /d "%REPO%"
git add -A
git commit -m "Prepara sistema para operacao online no Render"
git push

echo.
echo Finalizado.
echo Agora abra seu servico no Render e clique em Manual Deploy / Deploy latest commit.
echo Depois acesse /ambiente para validar o armazenamento e a senha.
echo.
pause
