@echo off
setlocal EnableExtensions
chcp 65001 >nul
title ORA - Publicar GitHub Pages

echo.
echo =====================================================
echo ORA NFS-e Automatico - Publicar interface no GitHub Pages
echo =====================================================
echo.

git --version >nul 2>nul
if errorlevel 1 (
  echo ERRO: Git nao foi encontrado no Windows.
  echo Instale o Git ou use o GitHub Desktop.
  echo.
  pause
  exit /b 1
)

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%.") do set "SCRIPT_DIR=%%~fI"

echo Pasta onde este assistente esta:
echo %SCRIPT_DIR%
echo.

if exist "%SCRIPT_DIR%\.git\" (
  set "REPO_DIR=%SCRIPT_DIR%"
  goto :REPO_OK
)

echo Esta pasta tem os arquivos do sistema, mas NAO tem a pasta .git.
echo Por isso o erro anterior apareceu: fatal: not a git repository.
echo.
echo Agora informe a pasta REAL do repositorio clonado do GitHub.
echo E a pasta onde voce ja fez o git init / git clone / git push.
echo.
echo Exemplos comuns:
echo C:\Users\%USERNAME%\Documents\GitHub\ora-nfse-automatico
echo C:\Users\%USERNAME%\OneDrive\Documentos\GitHub\ora-nfse-automatico
echo C:\Projetos\ora-nfse-automatico
echo.
set /p REPO_DIR=Arraste ou cole aqui a pasta do repositorio e pressione ENTER: 
set "REPO_DIR=%REPO_DIR:"=%"
for %%I in ("%REPO_DIR%") do set "REPO_DIR=%%~fI"

if not exist "%REPO_DIR%\.git\" (
  echo.
  echo ERRO: esta pasta tambem nao tem .git:
  echo %REPO_DIR%
  echo.
  echo Abra a pasta correta do repositorio local.
  echo Dica: no GitHub Desktop, clique em Repository ^> Show in Explorer.
  echo Depois rode este assistente novamente e informe essa pasta.
  echo.
  pause
  exit /b 1
)

echo.
echo Repositorio encontrado:
echo %REPO_DIR%
echo.

echo Copiando arquivos da versao atual para a raiz do repositorio...
robocopy "%SCRIPT_DIR%" "%REPO_DIR%" /E /XD ".git" "data" "logs" "cache" ".venv" "venv" "__pycache__" /XF ".env" "*.pfx" "*.p12" "*.pem" "*.key" "*.pyc" "*.sqlite" "*.db" >nul
if errorlevel 8 (
  echo ERRO: falha ao copiar os arquivos para o repositorio.
  echo Verifique permissoes e tente novamente.
  echo.
  pause
  exit /b 1
)

:REPO_OK
cd /d "%REPO_DIR%"

echo.
echo Conferindo arquivos obrigatorios na raiz...
if not exist "index.html" (
  echo ERRO: index.html ainda nao esta na raiz do repositorio.
  echo Pasta atual:
  cd
  echo.
  pause
  exit /b 1
)

if not exist ".nojekyll" (
  echo Criando .nojekyll...
  type nul > ".nojekyll"
)

echo OK: index.html encontrado.
echo OK: .nojekyll encontrado.
echo.

echo Arquivos principais:
dir /b index.html .nojekyll README.md 2>nul
echo.

echo Status do Git antes do envio:
git status --short
echo.

echo Adicionando arquivos ao Git...
git add index.html .nojekyll README.md PUBLICAR_GITHUB_PAGES.md GITHUB_SEM_BANCO.md requirements.txt .env.example .gitignore iniciar_windows.bat iniciar_windows_rede.bat iniciar_windows.ps1 iniciar_github_pages_windows.bat iniciar_github_pages_windows.ps1 corrigir_github_pages_windows.bat corrigir_github_pages_windows.ps1 PUBLICAR_AGORA_GITHUB_PAGES.bat app docs samples tests

echo.
git diff --cached --quiet
if errorlevel 1 (
  echo Criando commit...
  git commit -m "Publica interface do sistema no GitHub Pages"
  if errorlevel 1 (
    echo.
    echo ERRO: o commit falhou.
    echo Se aparecer erro de usuario/email, rode:
    echo git config --global user.name "Seu Nome"
    echo git config --global user.email "seuemail@email.com"
    echo.
    pause
    exit /b 1
  )
) else (
  echo Nenhuma alteracao nova para commit.
)

echo.
echo Enviando para o GitHub...
git push
if errorlevel 1 (
  echo.
  echo ERRO: o envio para o GitHub falhou.
  echo Verifique se o repositorio tem remote origin e se voce esta autenticado.
  echo.
  echo Rode estes comandos na pasta do repositorio para diagnosticar:
  echo git remote -v
  echo git branch
  echo.
  pause
  exit /b 1
)

echo.
echo =====================================================
echo Publicacao enviada com sucesso.
echo Aguarde 1 a 5 minutos e acesse:
echo https://victorgadiola-dev.github.io/ora-nfse-automatico/?v=13
echo.
echo Para forcar atualizacao no navegador: Ctrl + F5
echo =====================================================
echo.
pause
exit /b 0
