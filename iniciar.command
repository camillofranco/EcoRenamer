#!/bin/bash

# Este script descobre em qual diretório se encontra para sempre rodar no contexto correto
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "============================================="
echo "Renomeador de Imagens por Excel - Inicializador"
echo "============================================="

# Verifica se o Python 3 está instalado no sistema
if ! command -v python3 &> /dev/null
then
    echo "[ERRO] Python 3 não foi encontrado."
    echo "Por favor, instale o Python 3 ou verifique suas variáveis de ambiente."
    exit 1
fi

# Usa um ambiente virtual (venv) para evitar o erro "externally-managed-environment" no Mac
if [ ! -d "venv" ]; then
    echo "Configurando ambiente isolado para o programa pela primeira vez..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "[ERRO] Falha ao criar ambiente virtual. O Python pode não ter o módulo venv configurado."
        exit 1
    fi
fi

# Ativa o ambiente virtual local
source venv/bin/activate

# Verifica a existência do módulo openpyxl e tenta instalar caso não exista no venv
python3 -c "import openpyxl; import PIL; import fitz" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Bibliotecas requeridas não detectadas. Instalando..."
    python3 -m pip install --upgrade pip
    python3 -m pip install openpyxl pillow pillow-heif pymupdf
    
    if [ $? -ne 0 ]; then
        echo "[ERRO] Falha ao instalar openpyxl. Verifique sua conexão."
        exit 1
    fi
fi

echo "Iniciando aplicação..."
# Executa o aplicativo
python3 renomeador.py

# Desativa o ambiente virtual ao terminar
deactivate
