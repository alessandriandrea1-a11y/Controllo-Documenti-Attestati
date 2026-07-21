import streamlit as st
import pandas as pd
import sqlite3
import json
import io
import dropbox
import docx2txt
import re
from datetime import datetime
import base64
from groq import Groq
import pypdfium2 as pdfium

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
    st.markdown("### 🧠 CONFIGURAZIONE GROQ AI")
    api_key_manuale = st.text_input("Inserisci l'API Key di Groq (gsk_...):", type="password")
    api_key_segreta = st.secrets.get("GROQ_API_KEY", "")
    
    if api_key_manuale.strip() != "":
        api_key_inserita = api_key_manuale.strip()
        st.success("🔑 Usando l'API Key Groq inserita a mano!")
    elif api_key_segreta:
        api_key_inserita = api_key_segreta
        st.info("🤖 Usando l'API Key dai Secrets.")
    else:
        api_key_inserita = ""

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
        if not api_key_inserita:
            st.error("🚨 Inserisci la tua chiave API Groq (gsk_...) a sinistra per elaborare il documento!")
        else:
            with st.spinner("🧠 Groq AI sta analizzando il documento (PDF / Immagine)..."):
                try:
                    file_bytes = file_caricato.read()
                    nome_file = file_caricato.name.lower()
                    client = Groq(api_key=api_key_inserita)
                    data_oggi = datetime.now().strftime("%d/%m/%Y")
                    
                    prompt = f"""
                    Analizza questo documento di sicurezza sul lavoro o idoneità sanitaria. La data odierna di riferimento è il {data_oggi}.
                    
                    ISTRUZIONI SPECIALI DI LETTURA CRITICA:
                    - Fai massima attenzione alle scritte a mano o a penna, specialmente nelle sezioni dedicate a date, scadenze, limitazioni o prescrizioni.
                    - Se trovi solo MESE ed ANNO di scadenza (es: "Maggio 2028", "06/2027"), interpretalo come l'ultimo giorno di quel mese (es: "31/05/2028", "30/06/2027").
                    - Cerca attentamente nel verbale o certificato medico qualsiasi "limitazione", "prescrizione", nota o dicitura come "con prescrizione", "adibito a con la condizione di...". Estrai l'intera frase in modo chiaro.

                    Compiti di analisi:
                    1. Trova Nome e Cognome del lavoratore e la mansione.
                    2. Identifica il tipo preciso di documento.
                    3. CALCOLA/ESTRAI LA DATA DI SCADENZA (anche se scritta a penna o desunta da mese/anno).
                    4. Calcola lo stato rispetto al {data_oggi}: "🟢 In Regola", "🟡 In Scadenza", "🔴 Scaduto".
                    5. PRESCRIZIONI MEDICHE: Se presenti estraile dettagliatamente. Se non ce ne sono, restituisci null.

                    Rispondi ESCLUSIVAMENTE con un oggetto JSON valido:
                    {{
                        "lavoratore": "NOME COGNOME",
                        "mansione": "MANSIONE",
                        "documento_nome": "Nome Identificato del Documento",
                        "data_scadenza": "DD/MM/AAAA oppure 'Illimitato'",
                        "stato_calcolato": "🟢 In Regola / 🟡 In Scadenza / 🔴 Scaduto",
                        "prescrizione_medica": "Testo dettagliato delle prescrizioni/limitazioni estratte o null"
                    }}
                    """

                    if nome_file.endswith(".pdf"):
                        # Converte la prima pagina del PDF in un'immagine per Groq Vision
                        pdf = pdfium.PdfDocument(file_bytes)
                        page = pdf[0] # Prima pagina
                        image = page.render(scale=2).to_pil()
                        img_byte_arr = io.BytesIO()
                        image.save(img_byte_arr, format='PNG')
                        base64_image = base64.b64encode(img_byte_arr.getvalue()).decode('utf-8')
                        
                        messages = [
                            {
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": prompt},
                                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                                ]
                            }
                        ]
                        model_to_use = "llama-3.2-90b-vision-preview"

                    elif nome_file.endswith((".png", ".jpg", ".jpeg")):
                        base64_image = base64.b64encode(file_bytes).decode('utf-8')
                        mime = "image/png" if nome_file.endswith(".png") else "image/jpeg"
                        messages = [
                            {
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": prompt},
                                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{base64_image}"}}
                                ]
                            }
                        ]
                        model_to_use = "llama-3.2-90b-vision-preview"

                    elif nome_file.endswith(".docx"):
                        testo_word = docx2txt.process(io.BytesIO(file_bytes))
                        messages = [{"role": "user", "content": f"{prompt}\n\nContenuto del documento:\n{testo_word}"}]
                        model_to_use = "llama-3.3-70b-versatile"

                    chat_completion = client.chat.completions.create(
                        messages=messages,
                        model=model_to_use,
                        response_format={"type": "json_object"}
                    )
                    
                    risposta_testo = chat_completion.choices[0].message.content
                    dati_ai = json.loads(risposta_testo)
                    
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
                    st.success(f"🎉 Registrato/Aggiornato con Groq: {dati_ai['documento_nome']} per {dati_ai['lavoratore']}")
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
