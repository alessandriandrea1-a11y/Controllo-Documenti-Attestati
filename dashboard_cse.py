import streamlit as st
import pandas as pd
import sqlite3
import json
import io
import os
import zipfile
import dropbox
import docx2txt
import re
import requests
from datetime import datetime
import zoneinfo
from groq import Groq
import pypdfium2 as pdfium

st.set_page_config(layout="wide", page_title="Dashboard CSE — Controllo Totale Diretto")

# --- CONFIGURAZIONE DROPBOX E DATABASE ---
DB_FILE_NAME = "database_sicurezza.db"

DROPBOX_APP_KEY = st.secrets.get("DROPBOX_APP_KEY", "lz3k1850lvdbpe2")
DROPBOX_APP_SECRET = st.secrets.get("DROPBOX_APP_SECRET", "jcqww6ots1z1r9t")
DROPBOX_REFRESH_TOKEN = st.secrets.get("DROPBOX_REFRESH_TOKEN", "")

def get_dropbox_client():
    refresh_token = st.secrets.get("DROPBOX_REFRESH_TOKEN", DROPBOX_REFRESH_TOKEN)
    app_key = st.secrets.get("DROPBOX_APP_KEY", DROPBOX_APP_KEY)
    app_secret = st.secrets.get("DROPBOX_APP_SECRET", DROPBOX_APP_SECRET)
    
    if refresh_token:
        try:
            return dropbox.Dropbox(
                app_key=app_key,
                app_secret=app_secret,
                oauth2_refresh_token=refresh_token
            )
        except Exception as e:
            st.error(f"⚠️ Errore connessione Refresh Token Dropbox: {e}")
            
    token = st.secrets.get("DROPBOX_TOKEN", "")
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

# --- DATABASE LOCAL SQLITE E CREAZIONE TABELLE ---
conn = sqlite3.connect(DB_FILE_NAME, check_same_thread=False, timeout=20)
cursor = conn.cursor()

def inizializza_db():
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS aziende (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        nome TEXT UNIQUE,
        percorso_dropbox TEXT DEFAULT ''
    )
    """)
    
    cursor.execute("PRAGMA table_info(aziende)")
    colonne = [riga[1] for riga in cursor.fetchall()]
    if "percorso_dropbox" not in colonne:
        try:
            cursor.execute("ALTER TABLE aziende ADD COLUMN percorso_dropbox TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            pass

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
        FOREIGN KEY(lavoratore_id) REFERENCES lavoratori(id),
        UNIQUE(lavoratore_id, tipo_documento)
    )
    """)
    conn.commit()

inizializza_db()

# --- UTILITIES PER PULIZIA E NORMALIZZAZIONE ---
def pulisci_nome_rigido(nome_grezzo):
    if not nome_grezzo:
        return "SCONOSCIUTO"
    nome = re.sub(r'\(.*?\)', '', str(nome_grezzo))
    nome = re.sub(r'[-–—]', ' ', nome)
    nome_pulito = re.sub(r'\s+', ' ', nome).strip().upper()
    return nome_pulito if len(nome_pulito) > 3 else "SCONOSCIUTO"

def normalizza_nome_documento(testo_doc):
    if not testo_doc:
        return "Attestato Formazione Generico"
    
    t = str(testo_doc).lower()
    
    if "confinat" in t or "sospett" in t or "inquinament" in t:
        return "Formazione Ambienti Confinati"
    elif "antincendio" in t:
        return "Formazione Antincendio"
    elif "primo soccorso" in t:
        return "Formazione Primo Soccorso"
    elif "quota" in t or "cadute" in t:
        return "Formazione Lavori in Quota"
    elif "rls" in t:
        return "Formazione RLS"
    elif "rspp" in t:
        return "Formazione RSPP"
    elif "preposto" in t:
        return "Formazione Preposto"
    elif "medica" in t or "idoneit" in t or "sanitar" in t or "doc" in t:
        return "Idoneità Sanitaria"
    elif "accordo" in t or "generale" in t or "specifica" in t:
        return "Formazione Generale / Specifica"
    else:
        return testo_doc.strip().title()

def estrai_testo_da_bytes(file_bytes, nome_file):
    testo = ""
    nome_lower = nome_file.lower()
    try:
        if nome_lower.endswith(".pdf"):
            pdf = pdfium.PdfDocument(file_bytes)
            for page in pdf:
                textpage = page.get_textpage()
                t_pag = textpage.get_text_range()
                testo += t_pag + "\n"
        elif nome_lower.endswith(".docx"):
            testo = docx2txt.process(io.BytesIO(file_bytes))
    except Exception:
        pass
    
    testo_pulito = testo.strip()
    if len(testo_pulito) < 15:
        testo_pulito = f"NOME DEL FILE: {nome_file}"
        
    return testo_pulito

def elabora_singolo_documento_con_ai(file_bytes, nome_file, client, azienda_selezionata):
    nome_lower = nome_file.lower()
    
    if nome_lower.endswith(".zip"):
        risultati = []
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                for filename in z.namelist():
                    if filename.lower().endswith(('.pdf', '.docx')) and not filename.startswith('__MACOSX'):
                        unzipped_bytes = z.read(filename)
                        res = elabora_singolo_documento_con_ai(unzipped_bytes, os.path.basename(filename), client, azienda_selezionata)
                        if res:
                            risultati.append(res)
            return f"ZIP elaborato ({len(risultati)} file utili trovati)"
        except Exception as e:
            return f"Errore ZIP: {str(e)}"

    if not (nome_lower.endswith(".pdf") or nome_lower.endswith(".docx")):
        return None

    testo_estratto = estrai_testo_da_bytes(file_bytes, nome_file)

    fuso_orario = zoneinfo.ZoneInfo("Europe/Rome")
    data_oggi = datetime.now(fuso_orario).strftime("%d/%m/%Y")

    prompt = f"""
    Sei un addetto alla verifica documenti CSE in un cantiere.
    DATA ODIERNA DI CONFRONTO: {data_oggi}.
    NOME DEL FILE ANALIZZATO: "{nome_file}"
    TESTO DEL FILE:
    ---
    {testo_estratto}
    ---

    ISTRUZIONI TASSATIVE PER L'ESTRAZIONE:
    1. PERTINENZA: Tratta il file come "documento_pertinente": true solo se riguarda un LAVORATORE INDIVIDUALE (Attestato corso, Idoneità medica, Nomina, etc.). Se si tratta di documenti generali aziendali (POS, DURC, DVR, Verbali) imposta "documento_pertinente": false.
    2. NOME LAVORATORE: Estrai SOLO Nome e Cognome reali dell'operaio/lavoratore (dal testo o dal nome file). NON inventare o unire frasi casuali. Se non c'è un nome chiaro, imposta "documento_pertinente": false.
    3. DATA SCADENZA:
       - Cerca la data esplicita di SCADENZA o di ESECUZIONE del corso/visita.
       - Formato richiesto: GG/MM/AAAA.
       - Se nel testo NON è presente alcuna data certa o leggibile, scrivi ESATTAMENTE "Da Verificare" nella data.
    4. STATO:
       - Se la data di scadenza è passata rispetto a {data_oggi}, imposta "Scaduto".
       - Se scade nei prossimi 60 giorni, imposta "In Scadenza".
       - Se la data è futura (> 60 giorni), imposta "In Regola".
       - Se la data è "Da Verificare", imposta "In Scadenza".

    Rispondi SOLO in formato JSON:
    {{
        "documento_pertinente": true,
        "lavoratore": "NOME COGNOME",
        "mansione": "Operaio",
        "documento_nome": "Nome Corso / Visita Medica",
        "data_scadenza": "GG/MM/AAAA oppure Da Verificare",
        "stato_calcolato": "In Regola / In Scadenza / Scaduto",
        "prescrizione_medica": "Nessuna prescrizione rilevata"
    }}
    """

    try:
        try:
            chat_completion = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model="llama-3.3-70b-versatile",
                response_format={"type": "json_object"}
            )
        except Exception:
            chat_completion = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model="llama-3.1-8b-instant",
                response_format={"type": "json_object"}
            )

        dati_ai = json.loads(chat_completion.choices[0].message.content)

        if not dati_ai.get("documento_pertinente", True):
            return "Documento non individuale o non pertinente ignorato"

        nom_lav = pulisci_nome_rigido(dati_ai.get("lavoratore"))
        if nom_lav == "SCONOSCIUTO":
            return "Impossibile rilevare il nome del lavoratore"

        mans_lav = (dati_ai.get("mansione") or "Operaio").strip().title()
        doc_grezzo = (dati_ai.get("documento_nome") or "Attestato Formazione").strip()
        doc_nome = normalizza_nome_documento(doc_grezzo)
        data_scad = (dati_ai.get("data_scadenza") or "Da Verificare").strip()

        stato_raw = str(dati_ai.get("stato_calcolato", "")).lower()
        if "scaduto" in stato_raw:
            stato_calc = "🔴 Scaduto"
        elif "scadenza" in stato_raw:
            stato_calc = "🟡 In Scadenza"
        else:
            stato_calc = "🟢 In Regola"

        prescr_raw = dati_ai.get("prescrizione_medica")
        prescr = prescr_raw.strip() if (prescr_raw and str(prescr_raw).lower() != "null") else 'Nessuna prescrizione rilevata'

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
                conn.commit()
                op_id = cursor.lastrowid
            
            cursor.execute("DELETE FROM documenti_lavoratori WHERE lavoratore_id = ? AND tipo_documento = ?", (op_id, doc_nome))
            cursor.execute("""
                INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza)
                VALUES (?, ?, ?, ?)
            """, (op_id, doc_nome, stato_calc, data_scad))
            
            cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (op_id,))
            tutti_stati = [r[0] for r in cursor.fetchall()]
            stringa_totale = "".join(tutti_stati)
            
            if "🔴" in stringa_totale:
                nuovo_accesso = "🔴 INTERDETTO"
            elif "🟡" in stringa_totale:
                nuovo_accesso = "🟡 MONITORARE"
            else:
                nuovo_accesso = "🟢 ABILITATO"
            
            cursor.execute("UPDATE lavoratori SET stato_scadenza_totale = ? WHERE id = ?", (nuovo_accesso, op_id))
            conn.commit()
            upload_db_to_dropbox()

            return f"Registrato: {doc_nome} - {nom_lav} ({data_scad})"

    except Exception as e:
        return f"Errore AI: {str(e)}"

# Grafica CSS
st.markdown("""
    <style>
    html, body, [data-testid="stAppViewContainer"] { background-color: #f4f6f8; font-family: 'IBM Plex Sans', sans-serif; color: #0f1923; }
    .metric-card { background: white; padding: 20px; border-radius: 6px; border: 1px solid #dde3e9; box-shadow: 0 1px 3px rgba(0,0,0,0.05); text-align: center; }
    .metric-card h3 { margin: 0; font-size: 16px; color: #555; }
    .metric-card h2 { margin: 10px 0 0 0; font-size: 28px; color: #111; }
    .prescrizione-box { background-color: #fff3e0; border-left: 5px solid #ff9800; padding: 10px; margin-bottom: 15px; border-radius: 4px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

PASSWORD_CORRETTA = "Criansa2026"

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# --- BARRA LATERALE ---
with st.sidebar:
    st.markdown("### 🧠 CONFIGURAZIONE GROQ AI")
    api_key_manuale = st.text_input("Inserisci l'API Key di Groq (gsk_...):", type="password")
    api_key_segreta = st.secrets.get("GROQ_API_KEY", "")
    
    if api_key_manuale.strip() != "":
        api_key_inserita = api_key_manuale.strip()
        st.success("🔑 API Key Groq Manuale attiva!")
    elif api_key_segreta:
        api_key_inserita = api_key_segreta
        st.info("🤖 API Key da Secrets attiva.")
    else:
        api_key_inserita = ""

    st.write("---")
    st.markdown("### 🔑 CONNETTI DROPBOX PERMANENTEMENTE")
    with st.expander("🛠️ Genera Refresh Token"):
        st.markdown(f"""
        1. [👉 **Clicca qui per autorizzare Dropbox**](https://www.dropbox.com/oauth2/authorize?client_id={DROPBOX_APP_KEY}&response_type=code&token_access_type=offline)
        2. Autorizza l'app e copia il codice fornito.
        3. Incollalo subito sotto e premi il pulsante.
        """)
        code_input = st.text_input("Incolla il Codice Autorizzazione:")
        if st.button("🚀 Ottieni Refresh Token"):
            if code_input:
                try:
                    res = requests.post(
                        "https://api.dropbox.com/oauth2/token",
                        data={
                            "code": code_input.strip(),
                            "grant_type": "authorization_code",
                        },
                        auth=(DROPBOX_APP_KEY, DROPBOX_APP_SECRET)
                    )
                    data = res.json()
                    if "refresh_token" in data:
                        r_token = data["refresh_token"]
                        st.success("✅ Token Generato!")
                        st.code(f'DROPBOX_REFRESH_TOKEN = "{r_token}"', language="toml")
                        st.info("💡 Copia la riga di codice sopra e incollala nei tuoi Secrets su Streamlit Cloud!")
                    else:
                        st.error(f"Errore: {data.get('error_description', data)}")
                except Exception as ex:
                    st.error(f"Errore: {ex}")

    st.write("---")
    st.markdown("### 🔐 ACCESSO UTENTE")
    ruolo = st.selectbox("Seleziona il tuo ruolo:", ["👀 Solo Visualizzazione", "🛠️ Coordinatore (Modifica)"])
    
    ha_permesso_modifica = False
    if ruolo == "🛠️ Coordinatore (Modifica)":
        password_inserita = st.text_input("Inserisci la Password Amministratore:", type="password")
        if password_inserita == PASSWORD_CORRETTA:
            st.success("🔓 Accesso Modifica Abilitato!")
            ha_permesso_modifica = True
        elif password_inserita != "":
            st.error("🔒 Password Errata!")
            
    st.write("---")
    st.markdown("### 🏢 AZIENDE IN CANTIERE")
    cursor.execute("SELECT nome FROM aziende ORDER BY nome ASC")
    lista_aziende = [riga[0] for riga in cursor.fetchall()]
    azienda_selezionata = st.selectbox("Seleziona l'azienda:", lista_aziende) if lista_aziende else None
        
    if ha_permesso_modifica:
        st.write("---")
        st.markdown("#### 🏢 Configurazione Cantiere")
        nuova_azienda = st.text_input("Aggiungi Nuova Ditta:").strip()
        percorso_dbx_nuova = st.text_input("Percorso Cartella Dropbox Ditta (es. /Cantiere/Ditta_A):").strip()
        
        if st.button("Salva Azienda") and nuova_azienda:
            try:
                cursor.execute("INSERT INTO aziende (nome, percorso_dropbox) VALUES (?, ?)", (nuova_azienda, percorso_dbx_nuova))
                conn.commit()
                upload_db_to_dropbox()
                st.success("Azienda registrata!")
                st.rerun()
            except sqlite3.IntegrityError: 
                st.error("Esiste già.")

        if azienda_selezionata:
            if st.button(f"🗑️ Elimina Azienda ({azienda_selezionata})"):
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

# --- INTERFACCIA PRINCIPALE ---
if azienda_selezionata:
    st.markdown(f"# 🛡️ Dashboard CSE — Sistema di Controllo Integrato")
    st.markdown(f"### 🏢 Impresa in analisi: **{azienda_selezionata}**")
    st.write("---")
    
    cursor.execute("SELECT percorso_dropbox FROM aziende WHERE nome = ?", (azienda_selezionata,))
    row_p = cursor.fetchone()
    percorso_dropbox_ditta = row_p[0] if (row_p and row_p[0]) else ""

    if ha_permesso_modifica:
        c_left, c_right = st.columns(2)
        
        with c_left:
            st.markdown("#### 📦 Analisi Automatica Cartella Dropbox")
            if percorso_dropbox_ditta:
                st.caption(f"Cartella collegata: `{percorso_dropbox_ditta}`")
                if st.button("🚀 SCANSIONA ED ELABORA TUTTI I FILE DELLA CARTELLA DROPBOX"):
                    dbx = get_dropbox_client()
                    if not dbx:
                        st.error("⚠️ Nessun Token Dropbox attivo.")
                    elif not api_key_inserita:
                        st.error("🚨 Manca la Groq API Key! Inseriscila nella barra a sinistra.")
                    else:
                        with st.spinner("📦 Scansione cartella Dropbox ed elaborazione AI in corso..."):
                            client = Groq(api_key=api_key_inserita)
                            try:
                                res = dbx.files_list_folder(percorso_dropbox_ditta, recursive=True)
                                file_list = res.entries
                                processed = 0
                                
                                for entry in file_list:
                                    if isinstance(entry, dropbox.files.FileMetadata):
                                        if entry.name.lower().endswith(('.pdf', '.docx', '.zip')):
                                            st.write(f"🔄 Lettura file: `{entry.name}`...")
                                            _, file_res = dbx.files_download(entry.path_lower)
                                            f_bytes = file_res.content
                                            msg = elabora_singolo_documento_con_ai(f_bytes, entry.name, client, azienda_selezionata)
                                            if msg:
                                                st.info(f"➡️ {entry.name}: {msg}")
                                                processed += 1
                                st.success(f"✅ Scansione completata su {processed} file trovati.")
                                st.rerun()
                            except Exception as ex_dbx:
                                st.error(f"⚠️ Errore lettura cartella Dropbox: {ex_dbx}")
            else:
                st.warning("Nessun percorso Dropbox inserito per questa ditta.")

    if file_caricato is not None and ha_permesso_modifica:
        if not api_key_inserita:
            st.error("🚨 Inserisci la tua chiave API Groq a sinistra!")
        else:
            with st.spinner("🧠 Elaborazione file / archivio ZIP in corso..."):
                client = Groq(api_key=api_key_inserita)
                f_bytes = file_caricato.read()
                esito = elabora_singolo_documento_con_ai(f_bytes, file_caricato.name, client, azienda_selezionata)
                st.session_state["uploader_key"] += 1
                st.success(f"🎉 Risultato: {esito}")
                st.rerun()

    # --- TABELLA E STATISTICHE ---
    cursor.execute("""
        SELECT id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche FROM lavoratori 
        WHERE azienda_id = (SELECT id FROM aziende WHERE nome = ?)
        ORDER BY nominativo ASC
    """, (azienda_selezionata,))
    lavoratori = cursor.fetchall()
    
    if lavoratori:
        tot_lav = len(lavoratori)
        interdetti = sum(1 for l in lavoratori if "🔴" in l[3] or "INTERDETTO" in l[3])
        monitorare = sum(1 for l in lavoratori if "🟡" in l[3] or "MONITORARE" in l[3])
        abilitati = tot_lav - interdetti - monitorare
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><h3>👥 Forza Lavoro Totale</h3><h2>{tot_lav}</h2></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #4caf50;"><h3>🟢 Abilitati</h3><h2>{abilitati}</h2></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #ff9800;"><h3>🟡 Da Monitorare</h3><h2>{monitorare}</h2></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #f44336;"><h3>🔴 Interdetti</h3><h2>{interdetti}</h2></div>', unsafe_allow_html=True)
        
        st.write("---")
        st.markdown("### 📋 Fascicolo Elettronico dei Dipendenti")
        for lav in lavoratori:
            lav_id, nome, mansione, accesso, prescrizioni = lav
            
            cursor.execute("SELECT tipo_documento, stato_scadenza, data_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (lav_id,))
            docs = cursor.fetchall()
            
            tabella_pulita = []
            for doc_nome, validita, data_scad in docs:
                if not data_scad or data_scad == "None":
                    data_scad = "Da Verificare"
                tabella_pulita.append([doc_nome, validita, data_scad])
            
            with st.expander(f"{accesso} — 👤 {nome} ({mansione})"):
                if prescrizioni and prescrizioni != 'Nessuna prescrizione rilevata':
                    st.markdown(f'<div class="prescrizione-box">⚠️ **Prescrizioni Sanitarie:** {prescrizioni}</div>', unsafe_allow_html=True)
                
                if tabella_pulita:
                    df = pd.DataFrame(tabella_pulita, columns=["Documento Caricato", "Validità AI", "Scadenza Calcolata"])
                    st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.info("Nessun documento associato a questo lavoratore.")
                
                if ha_permesso_modifica:
                    if st.button(f"❌ Rimuovi Questo Lavoratore", key=f"del_{lav_id}"):
                        cursor.execute("DELETE FROM documenti_lavoratori WHERE lavoratore_id = ?", (lav_id,))
                        cursor.execute("DELETE FROM lavoratori WHERE id = ?", (lav_id,))
                        conn.commit()
                        upload_db_to_dropbox()
                        st.rerun()

        if ha_permesso_modifica:
            st.write("---")
            if st.button("🧹 PULISCILA ORA (ELIMINA RIGHE DUPLICATE VECCHIE)"):
                cursor.execute("SELECT id, lavoratore_id, tipo_documento, stato_scadenza, data_scadenza FROM documenti_lavoratori")
                tutti_i_docs = cursor.fetchall()
                cursor.execute("DELETE FROM documenti_lavoratori")
                
                mantenuti = {}
                for d_id, lav_id, t_doc, st_scad, d_scad in tutti_i_docs:
                    doc_norm = normalizza_nome_documento(t_doc)
                    chiave = (lav_id, doc_norm)
                    mantenuti[chiave] = (st_scad, d_scad)
                
                for (lav_id, doc_norm), (st_scad, d_scad) in mantenuti.items():
                    cursor.execute("""
                        INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza)
                        VALUES (?, ?, ?, ?)
                    """, (lav_id, doc_norm, st_scad, d_scad))
                
                conn.commit()
                upload_db_to_dropbox()
                st.success("✨ Tabella pulita con successo!")
                st.rerun()
    else:
        st.info("Nessun lavoratore registrato per questa azienda.")
else:
    st.info("👋 Seleziona o aggiungi un'azienda dalla barra laterale.")
