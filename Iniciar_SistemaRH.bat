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
if errorlevel 1 (
    echo [INFO] Instalando pandas...
    python -m pip install pandas
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Instalando openpyxl...
    python -m pip install openpyxl
)

echo [OK] Dependencias OK.
echo.

REM === Inicia o servidor Streamlit ===
echo [INFO] Iniciando servidor Streamlit na porta 8501...
echo [INFO] Aguarde alguns segundos...
echo [INFO] O navegador abrira automaticamente.
echo.
echo ----------------------------------------
echo Para encerrar, feche esta janela.
echo ----------------------------------------
echo.

REM Abre o navegador apos 5 segundos (em background)
start /min cmd /c "timeout /t 5 /nobreak >nul && start http://localhost:8501"

REM Inicia o Streamlit (logs aparecem aqui na tela)
python -m streamlit run app_rh.py --server.port 8501 --server.headless true --browser.gatherUsageStats false --server.enableCORS false

echo.
echo Servidor encerrado.
pause
