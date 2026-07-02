@echo off
setlocal
cd /d %~dp0

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
echo Abrindo ORA NFS-e em http://127.0.0.1:8000
echo Nao feche esta janela enquanto estiver usando o sistema.
echo.
start "" http://127.0.0.1:8000
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
