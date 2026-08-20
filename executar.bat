@echo off
title Calculadora de Laudas e Orcamentos
cd /d "%~dp0"
py app.py
if errorlevel 1 (
    echo.
    echo Nao foi possivel iniciar o programa.
    echo Instale as dependencias com: py -m pip install -r requirements.txt
    echo.
    pause
)

