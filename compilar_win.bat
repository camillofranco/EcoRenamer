@echo off
setlocal
cd /d %~dp0

echo ==========================================================
echo       🏗️ COMPILADOR ECOWAVE PRO v1.4.1 (WINDOWS)
echo ==========================================================

echo 1. Limpando lixos antigos...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist venv rmdir /s /q venv
if exist *.spec del /f /q *.spec
if exist *.zip del /f /q *.zip

echo 2. Criando Ambiente Virtual (Clean Slate)...
python -m venv venv
call venv\Scripts\activate.bat

echo 3. Instalando Dependencias Oficiais...
python -m pip install --upgrade pip
pip install openpyxl pillow pillow-heif pymupdf pyinstaller customtkinter darkdetect

echo 4. Compilando o aplicativo (PyInstaller)...
:: O Windows usa o .ico nativo do PyInstaller se ele existir
pyinstaller --noconfirm --windowed --noconsole --name "RenomeadorApp" --icon "icon.ico" --collect-all customtkinter --add-data "logo_ecowave.png;." renomeador.py

echo 5. Finalizando e Criando Zip Profissional (via PowerShell)...
powershell -Command "Compress-Archive -Path 'dist\RenomeadorApp\*' -DestinationPath 'EcoRenamer_Win_v1.4.1.zip' -Force"

echo --------------------------------------------------------
echo ✅ SUCESSO ABSOLUTO!
echo O arquivo pronto para o GitHub: EcoRenamer_Win_v1.4.1.zip
echo O aplicativo .exe esta em: dist\RenomeadorApp\RenomeadorApp.exe
echo --------------------------------------------------------
pause
