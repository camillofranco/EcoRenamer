#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "Compilando Aplicativo macOS Oficial..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi
source venv/bin/activate
pip install openpyxl pillow pillow-heif pymupdf pyinstaller

pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --icon "icon.ico" --add-data "logo_ecowave.png:." renomeador.py
echo "--------------------------------------------------------"
echo "COMPILACAO CONCLUIDA!"
echo "O aplicativo (RenomeadorApp.app) esta dentro da pasta 'dist/RenomeadorApp.app'."
echo "--------------------------------------------------------"
