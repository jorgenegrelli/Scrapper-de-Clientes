@echo off
chcp 65001 >nul
title Buscar clientes sem site
cd /d "%~dp0"

if not exist "sistema\python.txt" (
    echo Primeiro rode o INSTALAR.bat.
    pause
    exit /b 1
)
set /p PY=<"sistema\python.txt"

powershell -NoProfile -ExecutionPolicy Bypass -File "sistema\iniciar-motor.ps1"
if errorlevel 1 (
    pause
    exit /b 1
)
echo.
echo === Buscar clientes sem site proprio no Google Maps ===
echo.
set /p NICHO=Tipo de negocio (ex: dentista):
set /p CIDADE=Cidade (ex: Curitiba, PR):
set /p MINIMO=Quantos clientes no minimo (ex: 80):
if "%MINIMO%"=="" set MINIMO=80
echo.
"%PY%" "sistema\buscar_clientes.py" --nicho "%NICHO%" --cidade "%CIDADE%" --minimo %MINIMO%
echo.
echo Pronto. A planilha e o PAINEL (para enviar as mensagens) estao na pasta "Planilhas".
start "" "%~dp0Planilhas"
pause
