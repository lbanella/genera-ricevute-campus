#!/bin/bash
# Script di avvio per l'applicazione di generazione ricevute

# Naviga nella cartella del progetto
CDPATH= cd "$(dirname "$0")"

# Crea l'ambiente virtuale se non esiste
if [ ! -d ".venv" ]; then
    echo "Creazione dell'ambiente virtuale Python (.venv)..."
    python3 -m venv .venv
fi

# Attiva l'ambiente virtuale
source .venv/bin/activate

# Aggiorna pip e installa le dipendenze
echo "Verifica e installazione delle dipendenze..."
pip install --upgrade pip
pip install -r requirements.txt

# Avvia l'applicazione
echo "Avvio dell'applicazione..."
python3 src/app.py
