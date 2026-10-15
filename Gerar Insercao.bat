@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ==========================================================
echo   Gerar arquivo de insercao - demonstracao
echo ==========================================================
echo.

python gerar_insercao.py %*
if errorlevel 1 goto :erro

echo.
echo Concluido. Pode fechar esta janela.
pause
exit /b 0

:erro
echo.
echo *** Nada foi gerado. Leia a mensagem acima. ***
pause
exit /b 1
