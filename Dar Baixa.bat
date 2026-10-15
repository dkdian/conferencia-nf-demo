@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ==========================================================
echo   Dar baixa no ERP - demonstracao
echo ==========================================================
echo.

python dar_baixa.py %*
echo.
echo Concluido. Pode fechar esta janela.
pause
