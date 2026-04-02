@echo off
setlocal
cd /d %~dp0

echo ==========================================================
echo       🏗️ COMPILADOR ECOWAVE PRO v1.4.5 (WIN - CLEAN)
echo ==========================================================

echo 1. Limpando pastas antigas...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist venv rmdir /s /q venv
if exist *.spec del /f /q *.spec
if exist *.zip del /f /q *.zip

echo 2. Criando ambiente virtual...
python -m venv venv
call venv\Scripts\activate.bat

echo 3. Instalando dependências estáveis...
python -m pip install --upgrade pip
pip install openpyxl pillow pillow-heif pymupdf pyinstaller customtkinter darkdetect pdf2docx pdfplumber reportlab pytesseract python-docx

echo 4. Verificando Tesseract OCR...
where tesseract >nul 2>&1
if errorlevel 1 (
    winget install -e --id UB-Mannheim.TesseractOCR --silent 2>nul
)

echo 5. Compilando o aplicativo (PyInstaller)...
pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --icon "icon.ico" --add-data "logo_ecowave.png;." \
--collect-all customtkinter \
--hidden-import pytesseract \
renomeador.py

echo 6. Finalizando e criando Zip...
powershell -Command "Compress-Archive -Path 'dist\RenomeadorApp\*' -DestinationPath 'EcoRenamer_Win_v1.4.5.zip' -Force"

echo --------------------------------------------------------
echo ✅ SUCESSO! Versão 1.4.5 (Limpa) gerada.
echo --------------------------------------------------------
pause
