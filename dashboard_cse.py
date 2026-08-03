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

# Durate legali predefinite dei corsi di formazione sulla sicurezza (in Anni) - Normativa Italiana D.Lgs 81/08
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

def pulisci_percorso_dropbox(percorso_input):
    if not percorso_input:
        return ""
    p = str(percorso_input).strip()
    p = urllib.parse.unquote(p)
    
    if "dropbox.com" in p:
        if "/home" in p:
            p = p.split("/home")[-1]
        elif "/fo/" in p:
            p = p.split("/fo/")[-1]
        elif "/sh/" in p:
            p = p.split("/sh/")[-1]
        elif "/scl/fo/" in p:
            p = p.split("/scl/fo/")[-1]
            
    if "?" in p:
        p = p.split("?")[0]
        
    p = p.strip()
    if p and not p.startswith("/"):
        p = "/" + p
    return p

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

def download_db_from_dropbox():
    dbx = get_dropbox_client()
    if dbx:
        try:
            metadata, res = dbx.files_download(f"/{DB_FILE_NAME}")
            with open(DB_FILE_NAME, "wb") as f:
                f.write(res.content)
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

def get_db_connection():
    conn = sqlite3.connect(DB_FILE_NAME, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn

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

# --- GESTIONE TRACCIAMENTO FILE ---
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

# --- UTILITIES ---
def pulisci_nome_rigido(nome_grezzo):
    if not nome_grezzo:
        return "SCONOSCIUTO"
    nome = re.sub(r'\(.*?\)', '', str(nome_grezzo))
    nome = re.sub(r'[-–—]', ' ', nome)
    nome_pulito = re.sub(r'\s+', ' ', nome).strip().upper()
    return nome_pulito if len(nome_pulito) > 3 else "SCONOSCIUTO"

def aggiorna_stato_generale_lavoratore(lavoratore_id):
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

def estrai_pagine_da_pdf(file_bytes):
    pagine_estratte = []
    try:
        pdf = pdfium.PdfDocument(file_bytes)
        for i, page in enumerate(pdf):
            textpage = page.get_textpage()
            testo_pagina = textpage.get_text_range().strip()
            
            img_base64 = None
            if len(testo_pagina) < 30:
                try:
                    image = page.render(scale=2).to_pil()
                    buffered = io.BytesIO()
                    image.save(buffered, format="JPEG")
                    img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
                except Exception:
                    pass
            
            pagine_estratte.append({
                "numero_pagina": i + 1,
                "testo": testo_pagina,
                "immagine": img_base64
            })
    except Exception:
        pass
    return pagine_estratte

def stima_anni_validita_da_tipo(tipo_documento):
    doc_lower = (tipo_documento or "").lower()
    for chiave, anni in DURATA_CORSI_ANNI.items():
        if chiave in doc_lower:
            return anni
    return 5  # Valore predefinito D.Lgs 81/08

def calcola_stato_e_data_python(data_scad_str, data_emissione_str, anni_validita, tipo_documento=""):
    fuso_orario = zoneinfo.ZoneInfo("Europe/Rome")
    oggi = datetime.now(fuso_orario).date()
    
    doc_lower = (tipo_documento or "").lower()
    is_senza_scadenza = any(termine in doc_lower for termine in ["tesserino", "badge", "riconoscimento", "dpi", "consegna", "dispositivi"])
    
    if is_senza_scadenza:
        data_rif = data_emissione_str if (data_emissione_str and data_emissione_str != "NON_PRESENTI") else "Presente"
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt_rif = datetime.strptime(data_rif, fmt).date()
                data_rif = dt_rif.strftime("%d/%m/%Y")
                break
            except Exception:
                pass
        return data_rif, "🟢 In Regola"

    data_finale = None

    # 1. Se c'è una data di scadenza esplicita
    if data_scad_str and data_scad_str != "NON_PRESENTI":
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                data_finale = datetime.strptime(data_scad_str, fmt).date()
                break
            except ValueError:
                pass

    # 2. Se manca la data di scadenza esplicita, calcolala dalla data di fine corso/emissione
    if not data_finale and data_emissione_str and data_emissione_str != "NON_PRESENTI":
        if not anni_validita or int(anni_validita) <= 0:
            anni_validita = stima_anni_validita_da_tipo(tipo_documento)
            
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt_em = datetime.strptime(data_emissione_str, fmt).date()
                data_finale = dt_em.replace(year=dt_em.year + int(anni_validita))
                break
            except Exception:
                pass

    if not data_finale:
        return "Da Verificare", "🟡 In Scadenza"

    data_formattata = data_finale.strftime("%d/%m/%Y")
    giorni_rimanenti = (data_finale - oggi).days

    if giorni_rimanenti < 0:
        return data_formattata, "🔴 Scaduto"
    elif giorni_rimanenti <= 60:
        return data_formattata, "🟡 In Scadenza"
    else:
        return data_formattata, "🟢 In Regola"

def e_documento_sicurezza_pertinente(nome_file, testo_estratto):
    nome_lower = nome_file.lower()
    
    parole_da_scartare = ["pos", "p.o.s", "piano operativo", "fattura", "fatture", "preventivo", "ordine", "ddt", "bolla", "contabilita", "estratto", "pagamento", "acconto", "saldo", "banca", "bonifico", "contratto", "computo"]
    if any(p in nome_lower for p in parole_da_scartare):
        return False

    parole_chiave_sicurezza = [
        "attestato", "formazione", "corso", "sicurezza", "patentino", "idoneità", "idoneita", 
        "visita", "medica", "medico", "lavoratore", "preposto", "dirigente", "antincendio", 
        "primo soccorso", "spazi confinati", "ple", "gru", "muletto", "carrello", "imbracatore", 
        "dpi", "consegna", "tesserino", "badge", "riconoscimento", "asl", "ispettorato", 
        "dlgs 81", "formazione lavoratori", "formazione accordo stato regioni", "sorveglianza sanitaria"
    ]
    
    if any(p in nome_lower for p in parole_chiave_sicurezza):
        return True
        
    testo_lower = (testo_estratto or "").lower()
    match_testo = sum(1 for p in parole_chiave_sicurezza if p in testo_lower)
    
    return match_testo >= 2

def _esegui_chiamata_ai_e_salvataggio(testo_estratto, img_base64, nome_file, path_univoco, client, azienda_selezionata):
    system_prompt = """
Sei un esperto verificatore di documenti di cantiere e sicurezza sul lavoro (CSE).
Analizza il documento ed estrai con la massima precisione i dati del lavoratore e del documento.

REGOLE ESSENZIALI PER DATE E CALCOLO SCADENZE:
1. **Data Rilascio/Emissione/Fine Corso**: Cerca la data in cui il corso è stato completato, concluso o la data di rilascio del documento (es. "Svoltosi il...", "Concluso il...", "Data rilascio"). Formato GG/MM/AAAA.
2. **Data Scadenza**: Se è presente nel testo estraila. Se NON è indicata nel testo, metti "NON_PRESENTI" (verrà calcolata automaticamente dal sistema in base alla tipologia di corso).
3. **Anni Validità**: Se specificati esplicitamente estraili (es. 2, 3, 5), altrimenti imposta null.
4. **Tesserini e Consegna DPI**: Imposta sempre "data_scadenza": "NON_PRESENTI".

Dati richiesti:
1. Lavoratore (Nome e Cognome).
2. Mansione.
3. Nome ESATTO e SPECIFICO del documento (es. "Attestato Preposto", "Idoneità Sanitaria", "Patentino PLE", "Verbale Consegna DPI").
4. Data emissione / fine corso (GG/MM/AAAA) oppure NON_PRESENTI.
5. Data di scadenza esplicita (GG/MM/AAAA) oppure NON_PRESENTI.
6. Anni di validità se presenti nel testo, altrimenti null.
7. Prescrizioni sanitarie (se visita medica).

Restituisci ESCLUSIVAMENTE un JSON con questo schema:
{
    "documento_pertinente": true,
    "lavoratore": "NOME COGNOME",
    "mansione": "Operaio/Preposto/ecc",
    "documento_nome": "Nome Specifico del Documento",
    "data_emissione": "GG/MM/AAAA oppure NON_PRESENTI",
    "data_scadenza": "GG/MM/AAAA oppure NON_PRESENTI",
    "anni_validita": null,
    "prescrizione_medica": "Nessuna prescrizione rilevata"
}
"""

    try:
        if img_base64:
            # Utilizza il modello Vision attivo su Groq
            response = client.chat.completions.create(
                model="llama-3.2-90b-vision-preview",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": system_prompt},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}}
                        ]
                    }
                ],
                response_format={"type": "json_object"}
            )
        else:
            prompt_user = f"Nome del File: {nome_file}\nTesto Estratto:\n{testo_estratto}"
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt_user}
                ],
                response_format={"type": "json_object"}
            )

        contenuto_risposta = response.choices[0].message.content
        dati_ai = json.loads(contenuto_risposta)

        if not dati_ai.get("documento_pertinente", True):
            registra_file_processato(path_univoco)
            return None, False

        nom_lav = pulisci_nome_rigido(dati_ai.get("lavoratore"))
        if nom_lav == "SCONOSCIUTO":
            registra_file_processato(path_univoco)
            return None, False

        mans_lav = (dati_ai.get("mansione") or "Operaio").strip().title()
        doc_nome = (dati_ai.get("documento_nome") or "Attestato Generico").strip()
        data_em = dati_ai.get("data_emissione", "NON_PRESENTI")

        data_scad, stato_calc = calcola_stato_e_data_python(
            dati_ai.get("data_scadenza"),
            data_em,
            dati_ai.get("anni_validita"),
            tipo_documento=doc_nome
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
                cursor.execute("SELECT id FROM lavoratori WHERE azienda_id = ? AND UPPER(nominativo) = ?", (az_id, nom_lav))
                operaio_db = cursor.fetchone()
                
                if operaio_db:
                    op_id = operaio_db[0]
                    if prescr != 'Nessuna prescrizione rilevata':
                        cursor.execute("UPDATE lavoratori SET prescrizioni_mediche = ? WHERE id = ?", (prescr, op_id))
                else:
                    cursor.execute("INSERT INTO lavoratori (azienda_id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche) VALUES (?, ?, ?, '🔴 Da Verificare', ?)", (az_id, nom_lav, mans_lav, prescr))
                    op_id = cursor.lastrowid
                
                cursor.execute("SELECT id FROM documenti_lavoratori WHERE lavoratore_id = ? AND UPPER(tipo_documento) = ?", (op_id, doc_nome.upper()))
                doc_esistente = cursor.fetchone()
                
                if doc_esistente:
                    cursor.execute("""
                        UPDATE documenti_lavoratori 
                        SET stato_scadenza = ?, data_scadenza = ?, nome_file_origine = ?
                        WHERE id = ?
                    """, (stato_calc, data_scad, nome_file, doc_esistente[0]))
                    msg_esito = f"🔄 Aggiornato: {doc_nome} ({nom_lav}) -> Scadenza: {data_scad}"
                else:
                    cursor.execute("""
                        INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza, nome_file_origine)
                        VALUES (?, ?, ?, ?, ?)
                    """, (op_id, doc_nome, stato_calc, data_scad, nome_file))
                    msg_esito = f"✨ Registrato: {doc_nome} ({nom_lav}) -> Scadenza: {data_scad}"
                    
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
    
    parole_da_scartare = ["pos", "p.o.s", "piano operativo", "fattura", "fatture", "preventivo", "ordine", "ddt", "bolla", "contabilita", "estratto", "pagamento", "acconto", "saldo", "banca", "bonifico", "contratto", "computo"]
    if any(p in nome_lower for p in parole_da_scartare):
        registra_file_processato(path_univoco)
        return None, False

    if nome_lower.endswith(".zip"):
        risultati = []
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                for filename in z.namelist():
                    if filename.lower().endswith(('.pdf', '.docx')) and not filename.startswith('__MACOSX'):
                        sub_name_lower = filename.lower()
                        if any(p in sub_name_lower for p in parole_da_scartare):
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
            img_base64 = pag["immagine"]
            num_pag = pag["numero_pagina"]
            
            if not testo_estratto and not img_base64:
                continue
                
            sub_path_pag = f"{path_univoco}_pag_{num_pag}"
            if e_file_gia_processato(sub_path_pag):
                continue
                
            nome_file_pagina = f"{nome_file} (Pag. {num_pag})"
            esito_pag, stopped = _esegui_chiamata_ai_e_salvataggio(
                testo_estratto=testo_estratto,
                img_base64=img_base64,
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
            img_base64=None,
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
    .status-meter { background-color: #ffffff; padding: 10px 14px; border-radius: 6px; border: 1px solid #d1d5db; margin-top: 10px; font-size: 13px; }
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
                
                req_remaining_day = resp_test.headers.get("x-ratelimit-remaining-requests-day", "N/D")
                req_limit_day = resp_test.headers.get("x-ratelimit-limit-requests-day", "14400")
                req_remaining_min = resp_test.headers.get("x-ratelimit-remaining-requests-minute", "N/D")
                tok_remaining_min = resp_test.headers.get("x-ratelimit-remaining-tokens-minute", "N/D")
                
                st.markdown(f"""
                <div class="status-meter">
                    📊 <b>Stato Richieste Groq (Rimanenti):</b><br/>
                    • 📆 <b>Giornaliere (RPD):</b> {req_remaining_day} / {req_limit_day}<br/>
                    • ⏱️ <b>Al Minuto (RPM):</b> {req_remaining_min}<br/>
                    • 🔤 <b>Token Minuto (TPM):</b> {tok_remaining_min}
                </div>
                """, unsafe_allow_html=True)
            elif resp_test.status_code == 401:
                st.error("🔴 API Key NON valida o disattivata.")
            elif resp_test.status_code == 429:
                st.warning("🟡 Limite di token/minuto raggiunto. Attendi circa 60 secondi.")
            else:
                st.warning(f"⚠️ Errore test API: {resp_test.status_code}")
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
                st.success(f"Azienda registrata! Percorso Dropbox: {percorso_pulito}")
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
                    
        st.write("---")
        st.markdown("### 📤 UPLOAD MANUALE (PDF, DOCX, ZIP)")
        file_caricato = st.file_uploader("Carica File o Archivio ZIP", type=["pdf", "docx", "zip"], key=f"uploader_{st.session_state['uploader_key']}")
    else:
        file_caricato = None

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

    if ha_permesso_modifica:
        c_left, c_right = st.columns(2)
        
        with c_left:
            st.markdown("#### 📦 Analisi Automatica Cartella Dropbox")
            if percorso_dropbox_ditta:
                st.caption(f"Cartella base ditta: `{percorso_dropbox_ditta}`")
            
            sottocartella_specifica = st.text_input(
                "Sottocartella specifica (opzionale):", 
                value="", 
                placeholder="Es. /MioCantiere/Aggiornamenti"
            ).strip()
            
            percorso_da_usare = pulisci_percorso_dropbox(sottocartella_specifica) if sottocartella_specifica else percorso_dropbox_ditta

            if percorso_da_usare:
                st.caption(f"Percorso effettivo di scansione: `{percorso_da_usare}`")
                
                col_Avvia, col_Stop = st.columns(2)
                with col_Avvia:
                    avvia_scansione = st.button("🚀 SCANSIONA / RIPRENDI")
                with col_Stop:
                    ferma_scansione = st.button("🛑 INTERROMPI SCANSIONE")

                if ferma_scansione:
                    st.session_state["scansione_in_corso"] = False
                    st.warning("⚠️ Richiesta di interruzione registrata.")

                if avvia_scansione:
                    st.session_state["scansione_in_corso"] = True

                if st.session_state["scansione_in_corso"]:
                    dbx = get_dropbox_client()
                    if not dbx:
                        st.error("⚠️ Nessun Token Dropbox attivo.")
                        st.session_state["scansione_in_corso"] = False
                    elif not api_key_inserita:
                        st.error("🚨 Inserisci la chiave Groq API a sinistra!")
                        st.session_state["scansione_in_corso"] = False
                    else:
                        with st.spinner(f"📦 Scansione della cartella `{percorso_da_usare}` in corso..."):
                            client = Groq(api_key=api_key_inserita)
                            try:
                                res = dbx.files_list_folder(percorso_da_usare, recursive=True)
                                file_list = res.entries
                                
                                while res.has_more:
                                    res = dbx.files_list_folder_continue(res.cursor)
                                    file_list.extend(res.entries)

                                processed = 0
                                scartati = 0
                                bloccato_per_limit = False
                                
                                for entry in file_list:
                                    if not st.session_state["scansione_in_corso"]:
                                        st.info("🛑 Scansione interrotta manualmente dall'utente.")
                                        break

                                    if isinstance(entry, dropbox.files.FileMetadata):
                                        if entry.name.lower().endswith(('.pdf', '.docx', '.zip')):
                                            path_univoco = entry.path_lower
                                            
                                            if e_file_gia_processato(path_univoco):
                                                continue
                                                
                                            parole_da_scartare = ["pos", "p.o.s", "piano operativo", "fattura", "preventivo"]
                                            if any(p in entry.name.lower() for p in parole_da_scartare):
                                                registra_file_processato(path_univoco)
                                                scartati += 1
                                                continue

                                            _, res_file = dbx.files_download(entry.path_lower)
                                            esito_file, stopped = elabora_singolo_documento_con_ai(
                                                file_bytes=res_file.content,
                                                nome_file=entry.name,
                                                path_univoco=path_univoco,
                                                client=client,
                                                azienda_selezionata=azienda_selezionata
                                            )
                                            
                                            if stopped:
                                                st.error(esito_file)
                                                bloccato_per_limit = True
                                                st.session_state["scansione_in_corso"] = False
                                                break
                                                
                                            if esito_file:
                                                st.write(esito_file)
                                                processed += 1

                                if not bloccato_per_limit:
                                    st.success(f"✅ Scansione completata! Processati {processed} nuovi file (scartati {scartati}).")
                                    st.session_state["scansione_in_corso"] = False
                                    st.rerun()

                            except Exception as e_dbx:
                                st.error(f"⚠️ Errore durante l'accesso a Dropbox: {e_dbx}")
                                st.session_state["scansione_in_corso"] = False

        with c_right:
            st.markdown("#### 📄 Upload Diretto Locale")
            if file_caricato and api_key_inserita:
                client = Groq(api_key=api_key_inserita)
                with st.spinner("⏳ Analisi del file caricato in corso..."):
                    bytes_data = file_caricato.read()
                    esito_up, stopped = elabora_singolo_documento_con_ai(
                        file_bytes=bytes_data,
                        nome_file=file_caricato.name,
                        path_univoco=f"upload_locale_{file_caricato.name}",
                        client=client,
                        azienda_selezionata=azienda_selezionata
                    )
                    if stopped:
                        st.error(esito_up)
                    elif esito_up:
                        st.success(esito_up)
                        st.session_state["uploader_key"] += 1
                        st.rerun()
                    else:
                        st.info("ℹ️ Nessun documento rilevante estratto o file non di sicurezza.")

    # --- VISUALIZZAZIONE TABELLARE E METRICHE ---
    st.write("---")
    
    with get_db_connection() as conn:
        df_lavoratori = pd.read_sql_query("""
            SELECT l.id, l.nominativo AS [Lavoratore], l.mansione AS [Mansione], 
                   l.stato_scadenza_totale AS [Stato Idoneità Cantiere], 
                   l.prescrizioni_mediche AS [Prescrizioni Sanitarie]
            FROM lavoratori l
            JOIN aziende a ON l.azienda_id = a.id
            WHERE a.nome = ?
            ORDER BY l.nominativo ASC
        """, conn, params=(azienda_selezionata,))

    if not df_lavoratori.empty:
        totale_lav = len(df_lavoratori)
        abilitati = len(df_lavoratori[df_lavoratori["Stato Idoneità Cantiere"].str.contains("ABILITATO", na=False)])
        monitorare = len(df_lavoratori[df_lavoratori["Stato Idoneità Cantiere"].str.contains("MONITORARE", na=False)])
        interdetti = len(df_lavoratori[df_lavoratori["Stato Idoneità Cantiere"].str.contains("INTERDETTO", na=False)])

        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(f'<div class="metric-card"><h3>👥 Lavoratori</h3><h2>{totale_lav}</h2></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="metric-card"><h3>🟢 Abilitati</h3><h2>{abilitati}</h2></div>', unsafe_allow_html=True)
        m3.markdown(f'<div class="metric-card"><h3>🟡 Monitorare</h3><h2>{monitorare}</h2></div>', unsafe_allow_html=True)
        m4.markdown(f'<div class="metric-card"><h3>🔴 Interdetti</h3><h2>{interdetti}</h2></div>', unsafe_allow_html=True)

        st.write("---")
        st.markdown("### 📋 Elenco Personale e Stato Documentale")

        for idx, row in df_lavoratori.iterrows():
            lav_id = row["id"]
            nom = row["Lavoratore"]
            mans = row["Mansione"]
            stato = row["Stato Idoneità Cantiere"]
            prescr = row["Prescrizioni Sanitarie"]

            # Recupera documenti e calcola il numero di attestati
            with get_db_connection() as conn:
                df_docs = pd.read_sql_query("""
                    SELECT id, tipo_documento AS [Documento], stato_scadenza AS [Stato], 
                           data_scadenza AS [Data Scadenza / Riferimento], nome_file_origine AS [File di Origine]
                    FROM documenti_lavoratori
                    WHERE lavoratore_id = ?
                    ORDER BY data_scadenza ASC
                """, conn, params=(lav_id,))

            num_attestati = len(df_docs)

            with st.expander(f"{stato} | **{nom}** ({mans}) — 📜 Attestati registrati: {num_attestati}"):
                if prescr and prescr != 'Nessuna prescrizione rilevata':
                    st.markdown(f'<div class="prescrizione-box">⚠️ <b>Prescrizioni Sanitarie:</b> {prescr}</div>', unsafe_allow_html=True)

                if not df_docs.empty:
                    st.dataframe(df_docs.drop(columns=["id"]), use_container_width=True)
                else:
                    st.write("❌ Nessun documento attualmente registrato per questo lavoratore.")

                if ha_permesso_modifica:
                    if st.button(f"🗑️ Rimuovi Lavoratore ({nom})", key=f"del_lav_{lav_id}"):
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM documenti_lavoratori WHERE lavoratore_id = ?", (lav_id,))
                            cursor.execute("DELETE FROM lavoratori WHERE id = ?", (lav_id,))
                            conn.commit()
                        upload_db_to_dropbox()
                        st.success(f"Lavoratore {nom} eliminato!")
                        st.rerun()
    else:
        st.info("ℹ️ Nessun lavoratore o attestato registrato per questa ditta. Esegui la scansione della cartella Dropbox o carica un file manualmente.")
else:
    st.info("👈 Seleziona o crea un'azienda dalla barra laterale per accedere al Pannello di Controllo CSE.")
