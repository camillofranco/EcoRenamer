#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "=========================================================="
echo "      🏗️ COMPILADOR ECOWAVE PRO v1.4.4 (MACOS)"
echo "=========================================================="

echo "1. Limpando pastas antigas para evitar erros de cache..."
rm -rf build dist venv __pycache__ *.spec *.zip

echo "2. Criando ambiente virtual (Clean Slate)..."
python3 -m venv venv
source venv/bin/activate

echo "3. Instalando dependências oficiais..."
pip install --upgrade pip
pip install openpyxl pillow pillow-heif pymupdf pyinstaller customtkinter darkdetect pdf2docx pdfplumber reportlab

# Preparar ícone em alta resolução (Garante que fique nítido no Mac)
echo "4. Preparando ícone de alta resolução..."
# Redimensiona para quadrado 512x512 para o Mac
if [ -f "icon.png" ]; then
    sips -z 512 512 icon.png --out icon_mac.png > /dev/null 2>&1
    ICON_FILE="icon_mac.png"
else
    ICON_FILE=""
fi

echo "5. Compilando o aplicativo (PyInstaller)..."
if [ -n "$ICON_FILE" ]; then
    pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --icon "$ICON_FILE" --collect-all customtkinter --add-data "logo_ecowave.png:." renomeador.py
else
    pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --collect-all customtkinter --add-data "logo_ecowave.png:." renomeador.py
fi

echo "6. Finalizando e criando pacote Zip Seguro (-ry)..."
cd dist
zip -ry ../EcoRenamer_Mac_v1.4.4.zip RenomeadorApp.app
cd ..

echo "--------------------------------------------------------"
echo "✅ SUCESSO ABSOLUTO!"
echo "O arquivo pronto para o GitHub é: EcoRenamer_Mac_v1.4.4.zip"
echo "O aplicativo para uso direto está em: dist/RenomeadorApp.app"
echo "--------------------------------------------------------"
deactivate
