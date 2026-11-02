#!/bin/bash
cd "$(dirname "$0")"

echo "=========================================="
echo "Gestione Sostituzioni 3.3 - Avvio in corso..."
echo "Processore: $(uname -m)"
echo "OS: $(sw_vers -productVersion)"
echo "=========================================="

# Trova una versione di Python compatibile se serve creare il venv
PYTHON_CMD="python3"
if [ -f "/usr/local/bin/python3.13" ]; then
    PYTHON_CMD="/usr/local/bin/python3.13"
elif [ -f "/usr/local/bin/python3.12" ]; then
    PYTHON_CMD="/usr/local/bin/python3.12"
elif [ -f "/usr/local/bin/python3.11" ]; then
    PYTHON_CMD="/usr/local/bin/python3.11"
elif [ -f "/opt/homebrew/bin/python3" ]; then
    PYTHON_CMD="/opt/homebrew/bin/python3"
fi

# Crea venv se non esiste
if [ ! -d "venv" ]; then
    echo "Creazione ambiente virtuale con $PYTHON_CMD..."
    $PYTHON_CMD -m venv venv
    echo "Installazione dipendenze..."
    ./venv/bin/pip install -r requirements.txt
fi

echo "Lancio di Gestione Sostituzioni 3.2 (architettura modulare)..."
if [ -f "main.py" ]; then
    ./venv/bin/python3 main.py
else
    ./venv/bin/python3 Gestione_docenti_refactored.py
fi
