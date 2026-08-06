@echo off
title Sistema RH

echo ============================================
echo   SISTEMA RH
echo ============================================
echo.

REM === Verifica Python ===
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado!
    echo Instale em https://python.org e marque "Add to PATH"
    pause
    exit /b 1
)

echo [OK] Python encontrado.
echo.

REM === Instala dependencias (se necessario) ===
echo [INFO] Verificando dependencias...
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Instalando Streamlit, aguarde...
    python -m pip install streamlit pandas openpyxl Pillow
)
python -c "import pandas" >nul 2>&1
if errorlevel 1 python -m pip install pandas -q
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 python -m pip install openpyxl -q
python -c "from PIL import Image" >nul 2>&1
if errorlevel 1 python -m pip install Pillow -q
python -c "import matplotlib" >nul 2>&1
if errorlevel 1 python -m pip install matplotlib -q

echo [OK] Dependencias OK.
echo.

REM === Roda o launcher Python (mais robusto) ===
python launcher.py
