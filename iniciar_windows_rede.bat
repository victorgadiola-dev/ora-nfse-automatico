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
echo ORA NFS-e sera iniciado para acesso na rede interna.
echo Descubra o IP deste computador com: ipconfig
echo Acesse em outro computador usando: http://IP_DO_COMPUTADOR:8000
echo Nao exponha este sistema em rede publica.
echo.
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
pause
