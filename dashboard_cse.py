import base64
from datetime import datetime, timedelta
import io
import json
import os
import re
import sqlite3
import time
import urllib.parse
import zipfile
import zoneinfo

import docx2txt
import dropbox
from groq import Groq
import pandas as pd
import pypdfium2 as pdfium
import requests
import streamlit as st

st.set_page_config(layout="wide", page_title="Dashboard CSE — Controllo Completo Attestati")

# --- CONFIGURAZIONE DROPBOX E DATABASE ---
DB_FILE_NAME = "database_sicurezza.db"

DURATA_CORSI_ANNI = {
    "preposto": 2,
    "primo soccorso": 3,
    "antincendio": 3,
    "visita medica": 1,
    "idoneita": 1,
    "idoneità": 1,
    "sorveglianza sanitaria": 1,
    "ple": 5,
    "gru": 5,
    "carrello": 5,
    "muletto": 5,
    "patentino": 5,
    "escavatore": 5,
    "spazi confinati": 5,
    "lavori in quota": 5,
    "formazione lavoratori": 5,
    "generale": 5,
    "specifica": 5,
    "dirigente": 5,
    "rspp": 5,
    "aspp": 5,
    "rls": 1
}

PAROLE_DA_SCARTARE_ASSOLUTE = [
    "pos", "p.o.s", "piano operativo", "fattura", "fatture", "preventivo", 
    "ordine", "ddt", "bolla", "contabilita", "estratto", "pagamento", 
    "acconto", "saldo", "banca", "bonifico", "contratto", "computo", 
    "durc", "dvr", "valutazione rischi", "visura", "targa", "libretto", 
    "omologazione", "assicurazione", "verbale", "verbale cse", "nomina"
]

def pulisci_percorso_dropbox(percorso_input):
    if not percorso_input:
        return ""
    p = urllib.parse.unquote(str(percorso_input)).strip()
    if "?" in p:
        p = p.split("?")[0]
    if "dropbox.com" in p:
        for pattern in ["/home", "/fo/", "/sh/", "/scl/fo/", "/scl/fi/"]:
            if pattern in p:
                p = p.split(pattern)[-1]
                break
        if "dropbox.com" in p:
            parts = p.split("dropbox.com")
            p = parts[-1]
    p = p.strip()
    p = re.sub(r'/+', '/', p)
    if p and not p.startswith("/"):
        p = "/" + p
    if p == "/":
        return ""
    return p

def estrai_nome_lavoratore_da_percorso(path_file):
    try:
        parti = [p.strip() for p in path_file.split('/') if p.strip()]
        if len(parti) >= 2:
            cartella_padre = parti[-2]
            cartelle_escluse = [
                "1. ditta", "2. dipendenti", "3. mezzi", "ditte", "attestati", 
                "documenti", "pdf", "sub", "pos", "durc", "generale", "sicurezza"
            ]
            if not any(c in cartella_padre.lower() for c in cartelle_escluse) and len(cartella_padre) > 3:
                return cartella_padre
    except Exception:
        pass
    return None

def get_dropbox_client():
    refresh_token = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")
    app_key = st.secrets.get("DROPBOX_APP_KEY", "")
    app_secret = st.secrets.get("DROPBOX_APP_SECRET", "")
    token = st.secrets.get("DROPBOX_TOKEN", "")

    if refresh_token and app_key and app_secret:
        try:
            return dropbox.Dropbox(
                app_key=app_key,
                app_secret=app_secret,
                oauth2_refresh_token=refresh_token
            )
        except Exception as e:
            st.error(f"⚠️ Errore connessione Refresh Token Dropbox: {e}")

    if token:
        try:
            return dropbox.Dropbox(token)
        except Exception as e:
            st.error(f"⚠️ Errore connessione Token Dropbox: {e}")

    return None

def get_db_connection():
    conn = sqlite3.connect(DB_FILE_NAME, timeout=30)
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn

def download_db_from_dropbox():
    dbx = get_dropbox_client()
    if dbx:
        try:
            metadata, res = dbx.files_download(f"/{DB_FILE_NAME}")
            with open(DB_FILE_NAME, "wb") as f:
                f.write(res.content)
            
            try:
                temp_conn = sqlite3.connect(DB_FILE_NAME)
                temp_conn.execute("PRAGMA journal_mode=DELETE;")
                temp_conn.close()
            except Exception:
                pass
        except Exception:
            pass

def upload_db_to_dropbox():
    dbx = get_dropbox_client()
    if dbx:
        try:
            with open(DB_FILE_NAME, "rb") as f:
                dbx.files_upload(f.read(), f"/{DB_FILE_NAME}", mode=dropbox.files.WriteMode.overwrite)
        except Exception as e:
            st.error(f"⚠️ Errore nel salvataggio su Dropbox: {e}")

download_db_from_dropbox()

def inizializza_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS aziende (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            nome TEXT UNIQUE,
            percorso_dropbox TEXT DEFAULT ''
        )
        """)

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS lavoratori (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            azienda_id INTEGER,
            nominativo TEXT, 
            mansione TEXT, 
            stato_scadenza_totale TEXT,
            prescrizioni_mediche TEXT DEFAULT 'Nessuna prescrizione rilevata',
            FOREIGN KEY(azienda_id) REFERENCES aziende(id),
            UNIQUE(azienda_id, nominativo)
        )
        """)
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS documenti_lavoratori (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            lavoratore_id INTEGER,
            tipo_documento TEXT, 
            stato_scadenza TEXT, 
            data_scadenza TEXT,
            nome_file_origine TEXT DEFAULT '',
            FOREIGN KEY(lavoratore_id) REFERENCES lavoratori(id)
        )
        """)
        
        cursor.execute("PRAGMA table_info(documenti_lavoratori)")
        colonne = [col[1] for col in cursor.fetchall()]
        if "nome_file_origine" not in colonne:
            cursor.execute("ALTER TABLE documenti_lavoratori ADD COLUMN nome_file_origine TEXT DEFAULT ''")
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS file_processati (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            path_file TEXT UNIQUE,
            data_elaborazione TEXT
        )
        """)
        conn.commit()

inizializza_db()

def e_file_gia_processato(path_file):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM file_processati WHERE path_file = ?", (path_file,))
        return cursor.fetchone() is not None

def registra_file_processato(path_file):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT OR REPLACE INTO file_processati (path_file, data_elaborazione) VALUES (?, ?)", (path_file, now_str))
        conn.commit()

def cancella_memoria_file_cartella(percorso_cartella):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        pattern = f"%{percorso_cartella.lower()}%"
        cursor.execute("DELETE FROM file_processati WHERE LOWER(path_file) LIKE ?", (pattern,))
        conn.commit()
    upload_db_to_dropbox()

def pulisci_nome_rigido(nome_grezzo):
    if not nome_grezzo:
        return "SCONOSCIUTO"
    nome = re.sub(r'\(.*?\)', '', str(nome_grezzo))
    nome = re.sub(r'[-–—]', ' ', nome)
    nome_pulito = re.sub(r'\s+', ' ', nome).strip().upper()
    
    if len(nome_pulito) <= 3 or any(char.isdigit() for char in nome_pulito):
        return "SCONOSCIUTO"
        
    parole_aziendali = ["SRL", "SPA", "SAS", "SNC", "EDIL", "COSTRUZIONI", "IMPRESA", "DITTA", "S.R.L."]
    if any(p in nome_pulito for p in parole_aziendali):
        return "SCONOSCIUTO"
        
    return nome_pulito

def trova_o_crea_lavoratore(cursor, az_id, nom_lav, mans_lav, prescr):
    tokens_nuovi = set(nom_lav.split())
    
    cursor.execute("SELECT id, nominativo FROM lavoratori WHERE azienda_id = ?", (az_id,))
    lavoratori_esistenti = cursor.fetchall()
    
    for op_id, nom_db in lavoratori_esistenti:
        tokens_db = set(nom_db.upper().split())
        if tokens_nuovi == tokens_db:
            if prescr != 'Nessuna prescrizione rilevata':
                cursor.execute("UPDATE lavoratori SET prescrizioni_mediche = ? WHERE id = ?", (prescr, op_id))
            return op_id

    cursor.execute(
        "INSERT INTO lavoratori (azienda_id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche) VALUES (?, ?, ?, '🔴 Da Verificare', ?)", 
        (az_id, nom_lav, mans_lav, prescr)
    )
    return cursor.lastrowid

def aggiorna_stato_generale_lavoratore(lavoratore_id, conn_esistente=None):
    if conn_esistente:
        cursor = conn_esistente.cursor()
        cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (lavoratore_id,))
        tutti_stati = [r[0] for r in cursor.fetchall()]
        stringa_totale = "".join(tutti_stati)
        
        if "🔴" in stringa_totale or not tutti_stati:
            nuovo_accesso = "🔴 INTERDETTO"
        elif "🟡" in stringa_totale:
            nuovo_accesso = "🟡 MONITORARE"
        else:
            nuovo_accesso = "🟢 ABILITATO"
        
        cursor.execute("UPDATE lavoratori SET stato_scadenza_totale = ? WHERE id = ?", (nuovo_accesso, lavoratore_id))
    else:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (lavoratore_id,))
            tutti_stati = [r[0] for r in cursor.fetchall()]
            stringa_totale = "".join(tutti_stati)
            
            if "🔴" in stringa_totale or not tutti_stati:
                nuovo_accesso = "🔴 INTERDETTO"
            elif "🟡" in stringa_totale:
                nuovo_accesso = "🟡 MONITORARE"
            else:
                nuovo_accesso = "🟢 ABILITATO"
            
            cursor.execute("UPDATE lavoratori SET stato_scadenza_totale = ? WHERE id = ?", (nuovo_accesso, lavoratore_id))
            conn.commit()

def unifica_tutti_i_duplicati_azienda(azienda_nome):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_nome,))
        az_row = cursor.fetchone()
        if not az_row:
            return 0
        az_id = az_row[0]
        
        cursor.execute("SELECT id, nominativo FROM lavoratori WHERE azienda_id = ? ORDER BY id ASC", (az_id,))
        lavoratori = cursor.fetchall()
        
        gruppi = {}
        unificati_conteggio = 0
        lavoratori_da_aggiornare = set()
        
        for op_id, nom in lavoratori:
            key = frozenset(nom.upper().split())
            if key not in gruppi:
                gruppi[key] = op_id
            else:
                target_id = gruppi[key]
                
                cursor.execute("""
                    DELETE FROM documenti_lavoratori
                    WHERE lavoratore_id = ?
                      AND UPPER(tipo_documento) IN (
                          SELECT UPPER(tipo_documento) 
                          FROM documenti_lavoratori 
                          WHERE lavoratore_id = ?
                      )
                """, (op_id, target_id))

                cursor.execute("UPDATE documenti_lavoratori SET lavoratore_id = ? WHERE lavoratore_id = ?", (target_id, op_id))
                cursor.execute("DELETE FROM lavoratori WHERE id = ?", (op_id,))
                
                unificati_conteggio += 1
                lavoratori_da_aggiornare.add(target_id)
        
        for target_id in lavoratori_da_aggiornare:
            aggiorna_stato_generale_lavoratore(target_id, conn_esistente=conn)
                
        conn.commit()
        
    upload_db_to_dropbox()
    return unificati_conteggio

def normalizza_nome_documento(tipo_doc, testo_estratto=""):
    t = (tipo_doc or "").upper().strip()
    testo_l = (testo_estratto or "").lower()
    
    if "UNILAV" in t or "ASSUNZIONE" in t or ("COMUNICAZIONE" in t and "LAVORATORE" in testo_l):
        if "INDETERMINATO" in t or "T.I." in t or "tempo indeterminato" in testo_l or "a tempo indeterminato" in testo_l:
            return "UNILAV (Tempo Indeterminato)"
        elif "DETERMINATO" in t or "T.D." in t or "tempo determinato" in testo_l or "a tempo determinato" in testo_l:
            return "UNILAV (Tempo Determinato)"
        else:
            return "UNILAV (Tempo Indeterminato)"
            
    return tipo_doc.strip().title()

def estrai_pagine_da_pdf(file_bytes):
    pagine_estratte = []
    try:
        pdf = pdfium.PdfDocument(file_bytes)
        for i, page in enumerate(pdf):
            textpage = page.get_textpage()
            testo_pagina = textpage.get_text_range().strip()
            
            pagine_estratte.append({
                "numero_pagina": i + 1,
                "testo": testo_pagina
            })
    except Exception:
        pass
    return pagine_estratte

def stima_anni_validita_da_tipo(tipo_documento):
    doc_lower = (tipo_documento or "").lower()
    for chiave, anni in DURATA_CORSI_ANNI.items():
        if chiave in doc_lower:
            return anni
    return 5

def calcola_stato_da_stringa_data(data_scad_str, tipo_documento=""):
    doc_lower = (tipo_documento or "").lower()
    
    parole_senza_scadenza = ["dpi", "consegna dpi", "dispositivi", "tesserino", "badge", "riconoscimento", "unilav (tempo indeterminato)", "indeterminato"]
    if any(k in doc_lower for k in parole_senza_scadenza):
        data_mostrata = f"Consegnati ({data_scad_str})" if (data_scad_str and data_scad_str not in ["Da Verificare", "NON_PRESENTI", "Tempo Indeterminato"]) else "Tempo Indeterminato"
        return data_mostrata, "🟢 In Regola"

    if not data_scad_str or data_scad_str in ["Da Verificare", "NON_PRESENTI"]:
        return "Da Verificare", "🟡 In Scadenza"

    data_pulita = str(data_scad_str).strip()
    data_lower = data_pulita.lower()

    if any(k in data_lower for k in ["indeterminato", "t.i.", "tempo indeterminato", "illimitata", "presente", "consegnato", "consegnati"]):
        return "Tempo Indeterminato", "🟢 In Regola"

    fuso_orario = zoneinfo.ZoneInfo("Europe/Rome")
    oggi = datetime.now(fuso_orario).date()
    
    dt_scad = None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt_scad = datetime.strptime(data_pulita, fmt).date()
            break
        except ValueError:
            pass

    if not dt_scad:
        return data_pulita, "🟢 In Regola"

    data_formattata = dt_scad.strftime("%d/%m/%Y")
    giorni_rimanenti = (dt_scad - oggi).days

    if giorni_rimanenti < 0:
        return data_formattata, "🔴 Scaduto"
    elif giorni_rimanenti <= 60:
        return data_formattata, "🟡 In Scadenza"
    else:
        return data_formattata, "🟢 In Regola"

def calcola_stato_e_data_python(data_scad_str, data_emissione_str, anni_validita, tipo_documento="", testo_estratto=""):
    doc_lower = (tipo_documento or "").lower()
    testo_lower = (testo_estratto or "").lower()
    
    is_senza_scadenza = any(termine in doc_lower for termine in ["dpi", "consegna dpi", "dispositivi", "tesserino", "badge", "riconoscimento", "unilav (tempo indeterminato)", "indeterminato"])
    is_unilav_indeterminato = "unilav" in doc_lower and ("indeterminato" in doc_lower or "t.i." in doc_lower or "tempo indeterminato" in testo_lower or "a tempo indeterminato" in testo_lower)

    if is_senza_scadenza or is_unilav_indeterminato:
        data_notazione = f"Consegnati ({data_emissione_str})" if (data_emissione_str and data_emissione_str != "NON_PRESENTI") else "Tempo Indeterminato"
        return data_notazione, "🟢 In Regola"

    data_finale_str = data_scad_str

    if (not data_scad_str or data_scad_str == "NON_PRESENTI") and data_emissione_str and data_emissione_str != "NON_PRESENTI":
        if not anni_validita or int(anni_validita) <= 0:
            anni_validita = stima_anni_validita_da_tipo(tipo_documento)
            
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt_em = datetime.strptime(data_emissione_str, fmt).date()
                dt_scad = dt_em.replace(year=dt_em.year + int(anni_validita))
                data_finale_str = dt_scad.strftime("%d/%m/%Y")
                break
            except Exception:
                pass

    return calcola_stato_da_stringa_data(data_finale_str, tipo_documento=tipo_documento)

def e_documento_sicurezza_pertinente(nome_file, testo_estratto):
    nome_lower = nome_file.lower()
    if any(p in nome_lower for p in PAROLE_DA_SCARTARE_ASSOLUTE):
        return False

    parole_chiave_personale = [
        "attestato", "formazione", "corso", "idoneità", "idoneita", 
        "visita", "medica", "medico", "lavoratore", "preposto", "dirigente", "antincendio", 
        "primo soccorso", "spazi confinati", "ple", "gru", "muletto", "carrello", "imbracatore", 
        "dpi", "tesserino", "badge", "riconoscimento", "sorveglianza sanitaria", "unilav", "assunzione"
    ]
    
    if any(p in nome_lower for p in parole_chiave_personale):
        return True
        
    testo_lower = (testo_estratto or "").lower()
    if any(p in testo_lower for p in PAROLE_DA_SCARTARE_ASSOLUTE):
        return False
        
    match_testo = sum(1 for p in parole_chiave_personale if p in testo_lower)
    return match_testo >= 1

def _esegui_chiamata_ai_e_salvataggio(testo_estratto, nome_file, path_univoco, client, azienda_selezionata):
    lavoratore_fallback = estrai_nome_lavoratore_da_percorso(path_univoco)
    
    system_prompt = f"""Sei un esperto verificatore di documenti di cantiere (CSE).
Analizza il documento e determina SE E SOLO SE si tratta di un documento INDIVIDUALE di un LAVORATORE (Es. Attestato corso, Idoneità Medica, UNILAV, Tesserino, Consegna DPI).

REGOLE RIGIDE:
1. Se il documento è aziendale (POS, DURC, Fattura, Contratto, Scheda Mezzo/Macchinario, Verbale), imposta "documento_pertinente": false.
2. Se NON è presente un NOME E COGNOME di una persona fisica (lavoratore), imposta "documento_pertinente": false.
3. Se "documento_pertinente" è false, lascia gli altri campi vuoti o null.

Restituisci ESCLUSIVAMENTE un JSON valido con questo schema:
{{
    "documento_pertinente": true,
    "lavoratore": "NOME COGNOME",
    "mansione": "Operaio/Preposto/ecc",
    "documento_nome": "Nome Specifico del Documento",
    "data_emissione": "GG/MM/AAAA oppure NON_PRESENTI",
    "data_scadenza": "GG/MM/AAAA oppure Tempo Indeterminato oppure NON_PRESENTI",
    "anni_validita": null,
    "prescrizione_medica": "Nessuna prescrizione rilevata"
}}"""

    prompt_user = f"Nome File: {nome_file}\nPercorso: {path_univoco}\nTesto Estratto:\n{testo_estratto if testo_estratto else 'Documento Senza Testo Estratto'}"

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": str(system_prompt)},
                {"role": "user", "content": str(prompt_user)}
            ],
            response_format={"type": "json_object"}
        )

        contenuto_risposta = response.choices[0].message.content
        dati_ai = json.loads(contenuto_risposta)

        if not dati_ai.get("documento_pertinente", False):
            registra_file_processato(path_univoco)
            return None, False

        nom_lav = pulisci_nome_rigido(dati_ai.get("lavoratore"))
        
        if nom_lav == "SCONOSCIUTO" and lavoratore_fallback:
            nom_lav = pulisci_nome_rigido(lavoratore_fallback)

        if nom_lav == "SCONOSCIUTO":
            registra_file_processato(path_univoco)
            return None, False

        mans_lav = (dati_ai.get("mansione") or "Operaio").strip().title()
        
        doc_grezzo = dati_ai.get("documento_nome") or "Attestato Generico"
        doc_nome = normalizza_nome_documento(doc_grezzo, testo_estratto)
        
        data_em = dati_ai.get("data_emissione", "NON_PRESENTI")

        data_scad, stato_calc = calcola_stato_e_data_python(
            dati_ai.get("data_scadenza"),
            data_em,
            dati_ai.get("anni_validita"),
            tipo_documento=doc_nome,
            testo_estratto=testo_estratto
        )

        prescr_raw = dati_ai.get("prescrizione_medica")
        prescr = prescr_raw.strip() if (prescr_raw and str(prescr_raw).lower() != "null") else 'Nessuna prescrizione rilevata'

        op_id = None
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
            az_row = cursor.fetchone()
            if az_row:
                az_id = az_row[0]
                
                op_id = trova_o_crea_lavoratore(cursor, az_id, nom_lav, mans_lav, prescr)
                
                cursor.execute("SELECT id FROM documenti_lavoratori WHERE lavoratore_id = ? AND UPPER(tipo_documento) = ?", (op_id, doc_nome.upper()))
                doc_esistente = cursor.fetchone()
                
                if doc_esistente:
                    cursor.execute("""
                        UPDATE documenti_lavoratori 
                        SET stato_scadenza = ?, data_scadenza = ?, nome_file_origine = ?
                        WHERE id = ?
                    """, (stato_calc, data_scad, nome_file, doc_esistente[0]))
                    msg_esito = f"🔄 Aggiornato: {doc_nome} ({nom_lav}) -> Stato: {data_scad}"
                else:
                    cursor.execute("""
                        INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza, nome_file_origine)
                        VALUES (?, ?, ?, ?, ?)
                    """, (op_id, doc_nome, stato_calc, data_scad, nome_file))
                    msg_esito = f"✨ Registrato: {doc_nome} ({nom_lav}) -> Stato: {data_scad}"
                    
                conn.commit()

        if op_id:
            aggiorna_stato_generale_lavoratore(op_id)
            
        registra_file_processato(path_univoco)
        upload_db_to_dropbox()
        return msg_esito, False

    except Exception as e:
        err_msg = str(e)
        if "429" in err_msg or "rate_limit" in err_msg.lower():
            return "🛑 LIMIT RATE GROQ RAGGIUNTO: Token temporaneamente esauriti.", True
        return f"Errore AI: {err_msg}", False

def elabora_singolo_documento_con_ai(file_bytes, nome_file, path_univoco, client, azienda_selezionata):
    nome_lower = nome_file.lower()
    
    if any(p in nome_lower for p in PAROLE_DA_SCARTARE_ASSOLUTE):
        registra_file_processato(path_univoco)
        return None, False

    if nome_lower.endswith(".zip"):
        risultati = []
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                for filename in z.namelist():
                    if filename.lower().endswith(('.pdf', '.docx')) and not filename.startswith('__MACOSX'):
                        if any(p in filename.lower() for p in PAROLE_DA_SCARTARE_ASSOLUTE):
                            continue
                        unzipped_bytes = z.read(filename)
                        sub_path = f"{path_univoco}/{filename}"
                        if e_file_gia_processato(sub_path):
                            continue
                        res, stopped = elabora_singolo_documento_con_ai(
                            file_bytes=unzipped_bytes, 
                            nome_file=os.path.basename(filename), 
                            path_univoco=sub_path,
                            client=client, 
                            azienda_selezionata=azienda_selezionata
                        )
                        if stopped:
                            return res, True
                        if res:
                            risultati.append(res)
                        time.sleep(0.05)
            registra_file_processato(path_univoco)
            return f"ZIP elaborato ({len(risultati)} nuovi attestati trovati)", False
        except Exception as e:
            return f"Errore estrazione ZIP: {str(e)}", False

    if not (nome_lower.endswith(".pdf") or nome_lower.endswith(".docx")):
        return None, False

    if nome_lower.endswith(".pdf"):
        pagine = estrai_pagine_da_pdf(file_bytes)
        messaggi_esito = []
        testo_totale_pdf = " ".join([p["testo"] for p in pagine])
        
        if not e_documento_sicurezza_pertinente(nome_file, testo_totale_pdf):
            registra_file_processato(path_univoco)
            return None, False

        for pag in pagine:
            testo_estratto = pag["testo"]
            num_pag = pag["numero_pagina"]
            sub_path_pag = f"{path_univoco}_pag_{num_pag}"
            if e_file_gia_processato(sub_path_pag):
                continue
            nome_file_pagina = f"{nome_file} (Pag. {num_pag})"
            esito_pag, stopped = _esegui_chiamata_ai_e_salvataggio(
                testo_estratto=testo_estratto,
                nome_file=nome_file_pagina,
                path_univoco=sub_path_pag,
                client=client,
                azienda_selezionata=azienda_selezionata
            )
            if stopped:
                return esito_pag, True
            if esito_pag:
                messaggi_esito.append(esito_pag)
            time.sleep(0.05)
        registra_file_processato(path_univoco)
        if messaggi_esito:
            return " | ".join(messaggi_esito), False
        return None, False
    else:
        testo_estratto = ""
        try:
            testo_estratto = docx2txt.process(io.BytesIO(file_bytes))
        except Exception:
            testo_estratto = ""
        if not e_documento_sicurezza_pertinente(nome_file, testo_estratto):
            registra_file_processato(path_univoco)
            return None, False
        esito, stopped = _esegui_chiamata_ai_e_salvataggio(
            testo_estratto=testo_estratto,
            nome_file=nome_file,
            path_univoco=path_univoco,
            client=client,
            azienda_selezionata=azienda_selezionata
        )
        return esito, stopped

# --- INTERFACCIA STREAMLIT ---
st.markdown("""
    <style>
    html, body, [data-testid="stAppViewContainer"] { background-color: #f4f6f8; font-family: 'IBM Plex Sans', sans-serif; color: #0f1923; }
    .metric-card { background: white; padding: 20px; border-radius: 6px; border: 1px solid #dde3e9; box-shadow: 0 1px 3px rgba(0,0,0,0.05); text-align: center; }
    .metric-card h3 { margin: 0; font-size: 16px; color: #555; }
    .metric-card h2 { margin: 10px 0 0 0; font-size: 28px; color: #111; }
    .prescrizione-box { background-color: #fff3e0; border-left: 5px solid #ff9800; padding: 10px; margin-bottom: 15px; border-radius: 4px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

PASSWORD_CORRETTA = st.secrets.get("ADMIN_PASSWORD", "Criansa2026")

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "scansione_in_corso" not in st.session_state:
    st.session_state["scansione_in_corso"] = False

with st.sidebar:
    st.markdown("### 🧠 CONFIGURAZIONE GROQ AI")
    api_key_manuale = st.text_input("Inserisci l'API Key di Groq (gsk_...):", type="password")
    api_key_segreta = st.secrets.get("GROQ_API_KEY", "")
    
    if api_key_manuale.strip() != "":
        api_key_inserita = api_key_manuale.strip()
    elif api_key_segreta:
        api_key_inserita = api_key_segreta
    else:
        api_key_inserita = ""

    if api_key_inserita:
        try:
            url_test = "https://api.groq.com/openai/v1/chat/completions"
            headers_test = {
                "Authorization": f"Bearer {api_key_inserita}",
                "Content-Type": "application/json"
            }
            body_test = {
                "model": "llama-3.3-70b-versatile",
                "messages": [{"role": "user", "content": "ping"}],
                "max_tokens": 1
            }
            resp_test = requests.post(url_test, headers=headers_test, json=body_test, timeout=10)
            if resp_test.status_code == 200:
                st.success("🟢 API Key Valida e Funzionante!")
            elif resp_test.status_code == 401:
                st.error("🔴 API Key NON valida o disattivata.")
            elif resp_test.status_code == 429:
                st.warning("🟡 Limite di token/minuto raggiunto. Attendi circa 60 secondi.")
        except Exception as e_test:
            st.error(f"⚠️ Impossibile verificare la chiave: {e_test}")

    st.write("---")
    st.markdown("### 🔐 ACCESSO UTENTE")
    ruolo = st.selectbox("Seleziona il tuo ruolo:", ["👀 Solo Visualizzazione", "🛠️ Coordinatore (Modifica)"])
    
    ha_permesso_modifica = False
    if ruolo == "🛠️ Coordinatore (Modifica)":
        password_inserita = st.text_input("Inserisci la Password Amministratore:", type="password")
        if password_inserita == PASSWORD_CORRETTA:
            st.success("🔓 Accesso Abilitato!")
            ha_permesso_modifica = True
        elif password_inserita != "":
            st.error("🔒 Password Errata!")
            
    st.write("---")
    st.markdown("### 🏢 AZIENDE IN CANTIERE")
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT nome FROM aziende ORDER BY nome ASC")
        lista_aziende = [riga[0] for riga in cursor.fetchall()]

    azienda_selezionata = st.selectbox("Seleziona l'azienda:", lista_aziende) if lista_aziende else None
        
    if ha_permesso_modifica:
        st.write("---")
        st.markdown("#### 👤 Inserisci Dipendente Manualmente")
        if azienda_selezionata:
            with st.expander("➕ Aggiungi Nuovo Dipendente", expanded=False):
                with st.form("form_aggiungi_dipendente_manuale", clear_on_submit=True):
                    nom_manuale = st.text_input("Nome e Cognome Lavoratore:*", placeholder="Es. ANGELO PROETTO").strip().upper()
                    mans_manuale = st.text_input("Mansione:", value="Operaio").strip().title()
                    prescr_manuale = st.text_input("Prescrizioni Mediche:", value="Nessuna prescrizione rilevata").strip()
                    
                    st.markdown("**Opzionale: Registra Primo Documento**")
                    doc_manuale = st.text_input("Tipo Documento:", placeholder="Es. Idoneità Medica / Corso Sicurezza").strip()
                    data_scad_manuale = st.text_input("Data Scadenza o Stato:", placeholder="GG/MM/AAAA oppure Tempo Indeterminato").strip()
                    
                    btn_salva_manuale = st.form_submit_button("💾 Salva Dipendente")
                    
                    if btn_salva_manuale:
                        if not nom_manuale:
                            st.error("⚠️ Il campo Nome e Cognome è obbligatorio!")
                        else:
                            with get_db_connection() as conn:
                                cursor = conn.cursor()
                                cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
                                row_az = cursor.fetchone()
                                if row_az:
                                    az_id = row_az[0]
                                    op_id = trova_o_crea_lavoratore(cursor, az_id, nom_manuale, mans_manuale, prescr_manuale)
                                    
                                    if doc_manuale:
                                        doc_norm = normalizza_nome_documento(doc_manuale)
                                        data_scad_calc, stato_calc = calcola_stato_da_stringa_data(data_scad_manuale, tipo_documento=doc_norm)
                                        
                                        cursor.execute("""
                                            INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza, nome_file_origine)
                                            VALUES (?, ?, ?, ?, 'Inserito Manualmente')
                                        """, (op_id, doc_norm, stato_calc, data_scad_calc))
                                    
                                    conn.commit()
                                    aggiorna_stato_generale_lavoratore(op_id)
                                    upload_db_to_dropbox()
                                    st.success(f"✅ Dipendente **{nom_manuale}** registrato con successo!")
                                    st.rerun()
        else:
            st.info("💡 Seleziona un'azienda per abilitare l'inserimento manuale.")

        st.write("---")
        st.markdown("#### 🏢 Configurazione Cantiere")
        nuova_azienda = st.text_input("Aggiungi Nuova Ditta:").strip()
        percorso_dbx_nuova_raw = st.text_input("Percorso o Link Cartella Dropbox:", placeholder="Incolla link o percorso...")
        
        if st.button("Salva Azienda") and nuova_azienda:
            percorso_pulito = pulisci_percorso_dropbox(percorso_dbx_nuova_raw)
            try:
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO aziende (nome, percorso_dropbox) VALUES (?, ?)", (nuova_azienda, percorso_pulito))
                    conn.commit()
                upload_db_to_dropbox()
                st.session_state["azienda_selezionata"] = nuova_azienda
                st.success(f"Azienda registrata! Percorso Dropbox: `{percorso_pulito}`")
                st.rerun()
            except sqlite3.IntegrityError: 
                st.error("Azienda già esistente.")

        if azienda_selezionata:
            if st.button(f"🗑️ Elimina Azienda ({azienda_selezionata})"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM lavoratori WHERE azienda_id = (SELECT id FROM aziende WHERE nome = ?)", (azienda_selezionata,))
                    cursor.execute("DELETE FROM aziende WHERE nome = ?", (azienda_selezionata,))
                    conn.commit()
                upload_db_to_dropbox()
                st.success("Azienda eliminata!")
                st.rerun()

# --- DASHBOARD PRINCIPALE ---
if azienda_selezionata:
    st.markdown("# 🛡️ Dashboard CSE — Sistema di Controllo Integrato")
    st.markdown(f"### 🏢 Impresa in analisi: **{azienda_selezionata}**")
    st.write("---")
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT percorso_dropbox FROM aziende WHERE nome = ?", (azienda_selezionata,))
        row_p = cursor.fetchone()
        percorso_dropbox_ditta = row_p[0] if (row_p and row_p[0]) else ""
