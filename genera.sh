#!/bin/bash
# Script di generazione rapida senza browser

# Naviga nella cartella del progetto
CDPATH= cd "$(dirname "$0")"

# Attiva l'ambiente virtuale
if [ -d ".venv" ]; then
    source .venv/bin/activate
else
    echo "Ambiente virtuale .venv non trovato. Avvio prima run.sh per crearlo..."
    bash run.sh
    exit 0
fi

# Avvia lo script automatico
python3 genera.py
