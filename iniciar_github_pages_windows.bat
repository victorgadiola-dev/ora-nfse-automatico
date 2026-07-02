@echo off
setlocal
cd /d %~dp0

set GITHUB_PAGES_URL=https://victorgadiola-dev.github.io/ora-nfse-automatico/

if exist .env (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    if /I "%%A"=="GITHUB_PAGES_URL" set GITHUB_PAGES_URL=%%B
  )
)

if not exist .venv (
  echo Criando ambiente virtual...
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if not exist .env (
  copy .env.example .env >nul
)

if not exist data mkdir data

echo.
echo Abrindo interface publica ORA em:
echo %GITHUB_PAGES_URL%
echo.
echo O agente local sera iniciado em http://127.0.0.1:8000
echo Nao feche esta janela enquanto estiver usando o sistema.
echo.

start "" "%GITHUB_PAGES_URL%"
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
