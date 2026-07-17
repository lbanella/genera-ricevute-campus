@echo off
:: Script di avvio rapido per Windows (non-web)
cd /d "%~dp0"

echo ==================================================
echo       AVVIO GENERATORE RICEVUTE (WINDOWS)
echo ==================================================

:: Controlla se Python e installato
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERRORE: Python non e installato su questo computer.
    echo Per favore, installa Python da https://www.python.org/downloads/
    echo Assicurati di spuntare la casella "Add Python to PATH" durante l'installazione.
    echo.
    pause
    exit /b 1
)

:: Crea l'ambiente virtuale se non esiste
if not exist .venv (
    echo Creazione dell'ambiente virtuale Python (.venv)...
    python -m venv .venv
)

:: Attiva l'ambiente virtuale
call .venv\Scripts\activate

:: Aggiorna pip e installa le dipendenze
echo Verifica e installazione delle dipendenze...
python -m pip install --upgrade pip
pip install -r requirements.txt

:: Avvia il compilatore automatico
python genera.py

echo.
pause
