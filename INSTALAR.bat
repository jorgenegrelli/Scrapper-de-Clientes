@echo off
chcp 65001 >nul
title Instalar - Clientes para Sites
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "sistema\instalar.ps1"
echo.
pause
