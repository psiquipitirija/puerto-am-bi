@echo off
cd /d "%~dp0"
echo ============================================
echo  Puerto A&M - Business Intelligence
echo  Datos ficticios para uso academico
echo ============================================
echo.
python -m uvicorn app.main:app --host 127.0.0.1 --port 8124
pause
