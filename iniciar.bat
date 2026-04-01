@echo off
setlocal
cd /d "%~dp0"

echo =============================================
echo Renomeador de Imagens - Windows Bootstrapper
echo =============================================

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado ou nao esta no PATH. Instale o Python da Microsoft Store.
    pause
    exit /b
)

if not exist venv\ (
    echo Criando ambiente virtual...
    python -m venv venv
)

call venv\Scripts\activate.bat

python -c "import openpyxl, PIL, fitz" >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando dependencias (openpyxl, pillow, pymupdf)...
    python -m pip install --upgrade pip
    python -m pip install openpyxl pillow pillow-heif pymupdf
)

echo Iniciando aplicacao...
python renomeador.py

call deactivate
