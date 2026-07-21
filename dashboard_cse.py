import streamlit as st
import pandas as pd
import sqlite3
import json
from google import genai
from google.genai import types
import io
import dropbox
import docx2txt
import re

st.set_page_config(layout="wide", page_title="Dashboard CSE — Controllo Totale Diretto")

# --- CONFIGURAZIONE DROPBOX ---
DROPBOX_TOKEN = st.secrets.get("DROPBOX_TOKEN", "IL_TUO_TOKEN_TEMPORANEO_QUI")
DB_FILE_NAME = "database_sicurezza.db"

def download_db_from_dropbox():
    if DROPBOX_TOKEN and DROPBOX_TOKEN != "IL_TUO_TOKEN_TEMPORANEO_QUI":
        try:
            dbx = dropbox.Dropbox(DROPBOX_TOKEN)
            metadata, res = dbx.files_download(f"/{DB_FILE_NAME}")
            with open(DB_FILE_NAME, "wb") as f:
                f.write(res.content)
        except Exception as e:
            pass

def upload_db_to_dropbox():
    if DROPBOX_TOKEN and DROPBOX_TOKEN != "IL_TUO_TOKEN_TEMPORANEO_QUI":
        try:
            dbx = dropbox.Dropbox(DROPBOX_TOKEN)
            with open(DB_FILE_NAME, "rb") as f:
                dbx.files_upload(f.read(), f"/{DB_FILE_NAME}", mode=dropbox.files.WriteMode.overwrite)
        except Exception as e:
            st.error(f"Errore nel salvataggio su Dropbox: {e}")

download_db_from_dropbox()

# --- DATABASE LOCAL SQLITE ---
conn = sqlite3.connect(DB_FILE_NAME, check_same_thread=False)
cursor = conn.cursor()

cursor.execute("CREATE TABLE IF NOT EXISTS aziende (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT UNIQUE)")
cursor.execute("""
CREATE TABLE IF NOT EXISTS lavoratori (
    id INTEGER PRIMARY KEY AUTOINCREMENT, azienda_id INTEGER,
    nominativo TEXT, mansione TEXT, stato_scadenza_totale TEXT,
    prescrizioni_mediche TEXT DEFAULT 'Nessuna prescrizione rilevata',
    FOREIGN KEY(azienda_id) REFERENCES aziende(id),
    UNIQUE(azienda_id, nominativo)
)
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS documenti_lavoratori (
    id INTEGER PRIMARY KEY AUTOINCREMENT, lavoratore_id INTEGER,
    tipo_documento TEXT, stato_scadenza TEXT, data_scadenza TEXT,
    FOREIGN KEY(lavoratore_id) REFERENCES lavoratori(id),
    UNIQUE(lavoratore_id, tipo_documento)
)
""")
conn.commit()

# Aggiornamento colonne silente
try:
    cursor.execute("PRAGMA table_info(documenti_lavoratori)")
    colonne_doc = [c[1] for c in cursor.fetchall()]
    if 'data_scadenza' not in colonne_doc:
        cursor.execute("ALTER TABLE documenti_lavoratori ADD COLUMN data_scadenza TEXT")
        conn.commit()
except:
    pass

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

# --- BARRA LATERALE ---
with st.sidebar:
    st.markdown("### 🧠 CONFIGURAZIONE AI")
    api_key_segreta = st.secrets.get("GEMINI_API_KEY", "")
    if api_key_segreta and api_key_segreta != "da_inserire_su_chrome":
        api_key_inserita = api_key_segreta
        st.info("🤖 Chiave Gemini caricata dai Secrets.")
    else:
        api_key_inserita = st.text_input("Inserisci l'API Key di Google Gemini:", type="password")

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
        if st.button("Salva Azienda") and nuova_azienda:
            try:
                cursor.execute("INSERT INTO aziende (nome) VALUES (?)", (nuova_azienda,))
                conn.commit()
                upload_db_to_dropbox()
                st.success("Azienda registrata!")
                st.rerun()
            except sqlite3.IntegrityError: 
                st.error("Esiste già.")
                    
        st.write("---")
        st.markdown("### 📤 LETTORE AUTOMATICO MULTIMODALE")
        file_caricato = st.file_uploader("Carica Documento (PDF, PNG, JPG, DOCX)", type=["pdf", "png", "jpg", "jpeg", "docx"])
    else:
        file_caricato = None

# --- INTERFACCIA PRINCIPALE ---
if azienda_selezionata:
    st.markdown(f"# 🛡️ Dashboard CSE — Sistema di Controllo Integrato")
    st.markdown(f"### 🏢 Impresa in analisi: **{azienda_selezionata}**")
    st.write("---")
    
    if file_caricato is not None and ha_permesso_modifica:
        if not api_key_inserita or api_key_inserita == "da_inserire_su_chrome":
            st.error("🚨 Inserisci la tua chiave API Gemini qui a sinistra per elaborare il documento!")
        else:
            with st.spinner("🧠 L'AI sta analizzando la grafica e i testi (anche a penna) del documento..."):
                try:
                    file_bytes = file_caricato.read()
                    nome_file = file_caricato.name.lower()
                    contenuto_gemini = []
                    
                    if nome_file.endswith(".pdf"):
                        contenuto_gemini.append(types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"))
                    elif nome_file.endswith((".png", ".jpg", ".jpeg")):
                        mime = "image/png" if nome_file.endswith(".png") else "image/jpeg"
                        contenuto_gemini.append(types.Part.from_bytes(data=file_bytes, mime_type=mime))
                    elif nome_file.endswith(".docx"):
                        testo_word = docx2txt.process(io.BytesIO(file_bytes))
                        contenuto_gemini.append(testo_word)
                    
                    client = genai.Client(api_key=api_key_inserita)
                    
                    prompt = """
                    Analizza questo documento di sicurezza sul lavoro o idoneità sanitaria. La data odierna di riferimento è il 13/07/2026.
                    
                    ISTRUZIONI SPECIALI DI LETTURA CRITICA:
                    - Fai massima attenzione alle scritte a mano o a penna, specialmente nelle sezioni dedicate a date, scadenze, limitazioni o prescrizioni.
                    - Se trovi solo MESE ed ANNO di scadenza (es: "Maggio 2028", "06/2027"), interpretalo come l'ultimo giorno di quel mese (es: "31/05/2028", "30/06/2027").
                    - Cerca attentamente nel verbale o certificato medico qualsiasi "limitazione", "prescrizione", nota o dicitura come "con prescrizione", "adibito a con la condizione di...". Estrai l'intera frase in modo chiaro.

                    Compiti di analisi:
                    1. Trova Nome e Cognome del lavoratore e la mansione.
                    2. Identifica il tipo preciso di documento.
                    3. CALCOLA/ESTRAI LA DATA DI SCADENZA (anche se scritta a penna o desunta da mese/anno).
                    4. Calcola lo stato rispetto al 13/07/2026: "🟢 In Regola", "🟡 In Scadenza", "🔴 Scaduto".
                    5. PRESCRIZIONI MEDICHE: Se presenti estraile dettagliatamente. Se non ce ne sono, restituisci null.

                    Rispondi RIGOROSAMENTE in formato JSON con questa struttura:
                    {
                        "lavoratore": "NOME COGNOME",
                        "mansione": "MANSIONE",
                        "documento_nome": "Nome Identificato del Documento",
                        "data_scadenza": "DD/MM/AAAA oppure 'Illimitato'",
                        "stato_calcolato": "🟢 In Regola / 🟡 In Scadenza / 🔴 Scaduto",
                        "prescrizione_medica": "Testo dettagliato delle prescrizioni/limitazioni estratte o null"
                    }
                    """
                    contenuto_gemini.append(prompt)
                    
                    response = client.models.generate_content(
                        model='gemini-2.0-flash-lite',
                        contents=contenuto_gemini,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    
                    if response.text:
                        dati_ai = json.loads(response.text.strip())
                        
                        cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
                        az_id = cursor.fetchone()[0]
                        
                        cursor.execute("SELECT id FROM lavoratori WHERE azienda_id = ? AND LOWER(nominativo) = LOWER(?)", (az_id, dati_ai["lavoratore"].strip()))
                        operaio_db = cursor.fetchone()
                        
                        prescr = dati_ai["prescrizione_medica"] if dati_ai["prescrizione_medica"] else 'Nessuna prescrizione rilevata'
                        
                        if operaio_db:
                            op_id = operaio_db[0]
                            if dati_ai["prescrizione_medica"]:
                                cursor.execute("UPDATE lavoratori SET prescrizioni_mediche = ? WHERE id = ?", (prescr, op_id))
                        else:
                            cursor.execute("INSERT INTO lavoratori (azienda_id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche) VALUES (?, ?, ?, '🔴 Da Verificare', ?)", (az_id, dati_ai["lavoratore"].strip(), dati_ai["mansione"], prescr))
                            conn.commit()
                            op_id = cursor.lastrowid
                        
                        stato_pulito = dati_ai["stato_calcolato"].split("(")[0].strip()
                        
                        cursor.execute("""
                            INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza)
                            VALUES (?, ?, ?, ?)
                            ON CONFLICT(lavoratore_id, tipo_documento) 
                            DO UPDATE SET stato_scadenza=excluded.stato_scadenza, data_scadenza=excluded.data_scadenza
                        """, (op_id, dati_ai["documento_nome"], stato_pulito, dati_ai["data_scadenza"]))
                        
                        cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (op_id,))
                        tutti_stati = [r[0] for r in cursor.fetchall()]
                        stringa_totale = "".join(tutti_stati)
                        nuovo_accesso = "🔴 INTERDETTO" if "🔴" in stringa_totale else ("🟡 MONITORARE" if "🟡" in stringa_totale else "🟢 ABILITATO")
                        
                        cursor.execute("UPDATE lavoratori SET stato_scadenza_totale = ? WHERE id = ?", (nuovo_accesso, op_id))
                        conn.commit()
                        
                        upload_db_to_dropbox()
                        st.success(f"🎉 Registrato/Aggiornato: {dati_ai['documento_nome']} per {dati_ai['lavoratore']}")
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"Errore durante l'elaborazione AI: {str(e)}")

    # --- REPERIMENTO E VISUALIZZAZIONE DATI ---
    cursor.execute("""
        SELECT id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche FROM lavoratori 
        WHERE azienda_id = (SELECT id FROM aziende WHERE nome = ?)
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
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #4caf50;"><h3>🟢 Abilitati all\'Ingresso</h3><h2>{abilitati}</h2></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #ff9800;"><h3>🟡 Da Monitorare</h3><h2>{monitorare}</h2></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card" style="border-top: 4px solid #f44336;"><h3>🔴 Interdetti (Bloccati)</h3><h2>{interdetti}</h2></div>', unsafe_allow_html=True)
        
        st.write("---")
        
        st.markdown("### 📋 Fascicolo Elettronico dei Dipendenti")
        for lav in lavoratori:
            lav_id, nome, mansione, accesso, prescrizioni = lav
            
            cursor.execute("SELECT tipo_documento, stato_scadenza, data_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (lav_id,))
            docs = cursor.fetchall()
            
            tabella_pulita = []
            for doc_nome, validita, data_scad in docs:
                if not data_scad or data_scad == "None":
                    match = re.search(r'Scad\.\s*([\w/]+)', validita)
                    if match:
                        data_scad = match.group(1)
                    else:
                        data_scad = "Non richiesta / Illimitata"
                
                validita_pulita = validita.split("(")[0].strip()
                tabella_pulita.append([doc_nome, validita_pulita, data_scad])
            
            with st.expander(f"{accesso} — 👤 {nome} ({mansione})"):
                if prescrizioni and prescrizioni != 'Nessuna prescrizione rilevata':
                    st.markdown(f'<div class="prescrizione-box">⚠️ **Prescrizioni Sanitarie / Limitazioni:** {prescrizioni}</div>', unsafe_allow_html=True)
                
                if tabella_pulita:
                    df = pd.DataFrame(tabella_pulita, columns=["Documento Caricato", "Validità AI", "Scadenza Calcolata"])
                    st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.info("Nessun documento associato a questo lavoratore.")
                
                if ha_permesso_modifica:
                    if st.button(f"❌ Rimuovi {nome}", key=f"del_{lav_id}"):
                        cursor.execute("DELETE FROM documenti_lavoratori WHERE lavoratore_id = ?", (lav_id,))
                        cursor.execute("DELETE FROM lavoratori WHERE id = ?", (lav_id,))
                        conn.commit()
                        upload_db_to_dropbox()
                        st.rerun()
    else:
        st.info("Nessun lavoratore registrato per questa azienda. Passa al ruolo 'Coordinatore' per aggiungere file o aziende.")
