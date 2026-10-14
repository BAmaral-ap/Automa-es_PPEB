@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv
    if errorlevel 1 goto erro
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto erro
.venv\Scripts\python.exe run.py
if errorlevel 1 goto erro
echo Cartazes gerados. Confira a pasta cartazes_gerados.
pause
exit /b 0
:erro
echo Falha. Confira a mensagem acima e as instrucoes no README.md.
pause
exit /b 1
