@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ==========================================================
echo   Conferencia de Notas Fiscais - demonstracao
echo ==========================================================
echo.

python conferir_nf.py %*
if errorlevel 1 goto :erro

echo.
echo Concluido. Pode fechar esta janela.
pause
exit /b 0

:erro
echo.
echo *** A execucao terminou com erro. Leia a mensagem acima. ***
pause
exit /b 1
