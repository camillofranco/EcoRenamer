#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "=========================================================="
echo "      🏗️ COMPILADOR ECOWAVE PRO v1.4.5 (MACOS - CLEAN)"
echo "=========================================================="

echo "1. Limpando pastas antigas..."
rm -rf build dist venv __pycache__ *.spec *.zip

echo "2. Criando ambiente virtual..."
python3 -m venv venv
source venv/bin/activate

echo "3. Instalando dependências estáveis..."
pip install --upgrade pip
pip install openpyxl pillow pillow-heif pymupdf pyinstaller customtkinter darkdetect pdf2docx pdfplumber reportlab pytesseract python-docx

echo "4. Verificando Tesseract OCR..."
if ! command -v tesseract &>/dev/null; then
    if command -v brew &>/dev/null; then
        brew install tesseract tesseract-lang
    fi
fi

echo "5. Preparando ícone..."
if [ -f "icon.png" ]; then
    sips -z 512 512 icon.png --out icon_mac.png > /dev/null 2>&1
    ICON_FILE="icon_mac.png"
else
    ICON_FILE=""
fi

echo "6. Compilando o aplicativo (PyInstaller)..."
if [ -n "$ICON_FILE" ]; then
    pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --icon "$ICON_FILE" \
    --collect-all customtkinter \
    --hidden-import pytesseract \
    --add-data "logo_ecowave.png:." renomeador.py
else
    pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" \
    --collect-all customtkinter \
    --hidden-import pytesseract \
    --add-data "logo_ecowave.png:." renomeador.py
fi

echo "7. Finalizando e criando pacote Zip..."
cd dist
zip -ry ../EcoRenamer_Mac_v1.4.5.zip RenomeadorApp.app
cd ..

echo "--------------------------------------------------------"
echo "✅ SUCESSO! Versão 1.4.5 (Limpa) gerada."
echo "--------------------------------------------------------"
deactivate
