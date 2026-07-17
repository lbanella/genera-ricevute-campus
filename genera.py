import os
import re
import shutil
import subprocess
import sys
import pandas as pd
import openpyxl

# ==========================================
# FUNZIONI UTILI (SPLITTING E SOSTITUZIONE)
# ==========================================

def split_italian_name(full_name):
    """
    Divide un nominativo completo (Cognome e Nome) in Cognome e Nome separati.
    Gestisce i prefissi comuni italiani (Di, De, Da, Lo, La, Del, Della, Delle, ecc.).
    """
    if not full_name or not isinstance(full_name, str):
        return "", ""
    
    parts = full_name.strip().split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    
    prefixes = {"di", "de", "da", "lo", "la", "li", "del", "della", "delle", "delli", "dell", "van", "von", "d'"}
    
    first_word_lower = parts[0].lower().replace("'", "")
    if (first_word_lower in prefixes or parts[0].endswith("'")) and len(parts) > 2:
        cognome = parts[0] + " " + parts[1]
        nome = " ".join(parts[2:])
    else:
        cognome = parts[0]
        nome = " ".join(parts[1:])
        
    return cognome, nome

def make_safe_filename(cognome, nome):
    """
    Genera un nome file sicuro per il file system.
    """
    clean_cognome = "".join(c for c in cognome if c.isalnum())
    clean_nome = "".join(c for c in nome if c.isalnum())
    if not clean_cognome and not clean_nome:
        return "ricevuta"
    return f"{clean_cognome}{clean_nome}-ricevuta"

def replace_placeholders(text, mapping):
    """
    Sostituisce i segnaposto tipo {{campo}} in modo case-insensitive.
    """
    if not isinstance(text, str):
        return text
    
    def repl(match):
        key = match.group(1).strip().lower()
        if key == 'città':
            key = 'citta'
            
        if key in mapping:
            return str(mapping[key])
            
        col_mappings = {
            'cognome e nome del bambino': mapping.get('nome_bambino_completo', ''),
            'codice fiscale del bambino': mapping.get('cf_bambino', ''),
            'nominativo genitore a cui intestare la ricevuta': mapping.get('nome_genitore_completo', ''),
            'codice fiscale del genitore a cui intestare la ricevuta': mapping.get('cf_genitore', ''),
            'via di residenza': mapping.get('via', ''),
            'citta': mapping.get('citta', ''),
            'città': mapping.get('citta', ''),
            'cap': mapping.get('cap', '')
        }
        if key in col_mappings:
            return str(col_mappings[key])
            
        return match.group(0)

    return re.sub(r'\{\{?([^{}]+)\}\}?', repl, text)

# ==========================================
# LOGICA PRINCIPALE DI COMPILAZIONE
# ==========================================

def main():
    print("==================================================")
    print("      COMPILATORE AUTOMATICO DI RICEVUTE PDF       ")
    print("==================================================")
    
    # 1. Ricerca dei file nella cartella corrente
    dati_path = None
    template_path = None
    
    current_files = os.listdir('.')
    
    # Cerca file dei paganti
    for f in current_files:
        if 'paganti' in f.lower() and (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlxsx')):
            dati_path = f
            break
            
    # Cerca il modello di ricevuta
    for f in current_files:
        if ('modello' in f.lower() or 'ricevuta' in f.lower()) and f.endswith('.xlsx'):
            template_path = f
            break
            
    # Se non trovati, prova in templates/
    if not dati_path and os.path.exists('templates'):
        for f in os.listdir('templates'):
            if 'paganti' in f.lower() and (f.endswith('.xlsx') or f.endswith('.xls')):
                dati_path = os.path.join('templates', f)
                break
                
    if not template_path and os.path.exists('templates'):
        for f in os.listdir('templates'):
            if ('modello' in f.lower() or 'ricevuta' in f.lower()) and f.endswith('.xlsx'):
                template_path = os.path.join('templates', f)
                break

    # Se mancano i file
    if not dati_path:
        print("ERRORE: Impossibile trovare il file Excel dei paganti (es: paganti.xlsx)")
        print("Assicurati che sia presente in questa cartella.")
        sys.exit(1)
        
    if not template_path:
        print("ERRORE: Impossibile trovare il file Excel modello della ricevuta (es: modello_ricevuta.xlsx)")
        print("Assicurati che sia presente in questa cartella.")
        sys.exit(1)

    print(f"-> File Dati rilevato: {dati_path}")
    print(f"-> File Modello rilevato: {template_path}")
    
    output_dir = "output_ricevute"
    os.makedirs(output_dir, exist_ok=True)
    print(f"-> I PDF verranno salvati in: {os.path.abspath(output_dir)}")
    print("--------------------------------------------------")
    
    try:
        # Carica foglio dati (tutto come testo)
        df = pd.read_excel(dati_path, dtype=str)
        df.columns = [c.strip() for c in df.columns]
        
        required_cols = [
            "Cognome e Nome del Bambino",
            "Codice Fiscale del Bambino",
            "Nominativo Genitore a cui Intestare la Ricevuta",
            "Codice Fiscale del Genitore a cui Intestare la Ricevuta",
            "Via di Residenza",
            "CITTA",
            "CAP"
        ]
        
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            print(f"ERRORE: Mancano colonne nell'Excel dei dati:\n{', '.join(missing_cols)}")
            sys.exit(1)
            
        total = len(df)
        print(f"Trovate {total} righe da elaborare. Avvio conversione...")
        
        def safe_str(val):
            if pd.isna(val):
                return ""
            s = str(val).strip()
            if s.endswith(".0"):
                s = s[:-2]
            if s.lower() == 'nan':
                return ""
            return s
            
        for index, row in df.iterrows():
            nominativo_bambino = safe_str(row["Cognome e Nome del Bambino"])
            cf_bambino = safe_str(row["Codice Fiscale del Bambino"])
            nominativo_genitore = safe_str(row["Nominativo Genitore a cui Intestare la Ricevuta"])
            cf_genitore = safe_str(row["Codice Fiscale del Genitore a cui Intestare la Ricevuta"])
            via = safe_str(row["Via di Residenza"])
            citta = safe_str(row["CITTA"])
            cap = safe_str(row["CAP"])
            
            cognome_bambino, nome_bambino = split_italian_name(nominativo_bambino)
            cognome_genitore, nome_genitore = split_italian_name(nominativo_genitore)
            
            row_data = {
                'nome_bambino_completo': nominativo_bambino,
                'cognome_bambino': cognome_bambino,
                'nome_bambino': nome_bambino,
                'cf_bambino': cf_bambino,
                'nome_genitore_completo': nominativo_genitore,
                'cognome_genitore': cognome_genitore,
                'nome_genitore': nome_genitore,
                'cf_genitore': cf_genitore,
                'via': via,
                'citta': citta,
                'cap': cap
            }
            
            # Compila Excel
            wb = openpyxl.load_workbook(template_path)
            for ws in wb.worksheets:
                for r in ws.iter_rows():
                    for cell in r:
                        if cell.value is not None and isinstance(cell.value, str):
                            cell.value = replace_placeholders(cell.value, row_data)
            
            import tempfile
            filename_base = make_safe_filename(cognome_bambino, nome_bambino)
            temp_xlsx = os.path.join(tempfile.gettempdir(), f"{filename_base}.xlsx")
            wb.save(temp_xlsx)
            wb.close()

            
            # Converte in PDF usando LibreOffice
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                output_dir,
                temp_xlsx
            ]
            
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if result.returncode != 0:
                # Prova con soffice
                cmd[0] = "soffice"
                result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                if result.returncode != 0:
                    print(f"[{index+1}/{total}] ERRORE nella conversione PDF per: {nominativo_bambino}")
                    if os.path.exists(temp_xlsx):
                        os.remove(temp_xlsx)
                    continue
            
            # Pulisce temporaneo
            if os.path.exists(temp_xlsx):
                os.remove(temp_xlsx)
                
            print(f"[{index+1}/{total}] PDF Generato con successo: {filename_base}.pdf")
            
        print("--------------------------------------------------")
        print("PROCESSO COMPLETATO! Trovi i PDF nella cartella 'output_ricevute'.")
        
    except Exception as e:
        print(f"ERRORE GENERALE DI ELABORAZIONE: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
