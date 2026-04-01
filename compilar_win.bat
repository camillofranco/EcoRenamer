@echo off
setlocal
cd /d "%~dp0"

echo Compilando Executavel Windows Oficial...
if not exist venv\ (
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install openpyxl pillow pillow-heif pymupdf pyinstaller

pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" renomeador.py
echo --------------------------------------------------------
echo COMPILACAO CONCLUIDA!
echo O executavel (.exe) esta dentro da pasta 'dist'.
echo --------------------------------------------------------
pause
