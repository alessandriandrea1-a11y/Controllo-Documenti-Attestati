import streamlit as st
import sqlite3
import os
import io
import base64
import json
from datetime import datetime
import pypdfium2 as pdfium
from groq import Groq

# Configurazione della pagina Streamlit
st.set_page_config(
    page_title="Dashboard Sicurezza CSE",
    page_icon="🛡️",
    layout="wide"
)

# Inizializzazione client Groq (assicurati di avere la chiave nei Secrets di Streamlit o variabile d'ambiente)
# Se non è impostata, legge da st.secrets
groq_api_key = st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", ""))
client = Groq(api_key=groq_api_key)

# ---------------------------------------------------------
# DATABASE SETUP
# ---------------------------------------------------------
DB_NAME = "database_sicurezza.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS lavoratori (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cognome TEXT NOT NULL,
            codice_fiscale TEXT UNIQUE,
            mansione TEXT,
            stato TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS certificati (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lavoratore_id INTEGER,
            nome_corso TEXT,
            data_emissione TEXT,
            data_scadenza TEXT,
            file_name TEXT,
            FOREIGN KEY (lavoratore_id) REFERENCES lavoratori(id)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# FUNZIONI DI ESTRAZIONE PDF (MULTIPAGINA)
# ---------------------------------------------------------
def estrai_pagine_da_pdf(file_bytes):
    """Estrae il testo e l'immagine pagina per pagina da un PDF multipagina."""
    pagine_estratte = []
    try:
        pdf = pdfium.PdfDocument(file_bytes)
        for i, page in enumerate(pdf):
            textpage = page.get_textpage()
            testo_pagina = textpage.get_text_range().strip()
            
            img_base64 = None
            # Se il testo è troppo corto (es. PDF scannerizzato), estraiamo l'immagine di QUELLA pagina
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
    except Exception as e:
        st.error(f"Errore nella lettura del PDF: {e}")
        
    return pagine_estratte

# ---------------------------------------------------------
# INTERAZIONE CON GROQ AI
# ---------------------------------------------------------
def analizza_contenuto_con_ai(testo, immagine_base64=None):
    """Invia il testo o l'immagine della pagina a Groq per estrarre i dati del certificato."""
    prompt = """
    Analizza il documento di sicurezza allegato ed estrai le seguenti informazioni in formato JSON puro:
    - nome (Nome del lavoratore)
    - cognome (Cognome del lavoratore)
    - codice_fiscale (Codice fiscale del lavoratore, se presente)
    - mansione (Mansione o ruolo, se presente)
    - nome_corso (Titolo specifico del corso di formazione, es. 'RSPP', 'Muletto', 'Spazi Confinati', ecc.)
    - data_emissione (Data di emissione nel formato YYYY-MM-DD, se presente)
    - data_scadenza (Data di scadenza nel formato YYYY-MM-DD, se presente)

    Rispondi SOLO con un oggetto JSON valido, senza markdown o commenti aggiuntivi. Se un campo non è presente, metti null.
    """
    
    messages = []
    if immagine_base64:
        messages.append({
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{immagine_base64}"
                    }
                }
            ]
        })
        model_to_use = "llama-3.2-11b-vision-preview"
    else:
        messages.append({
            "role": "user",
            "content": f"{prompt}\n\nTesto del documento:\n{testo}"
        })
        model_to_use = "llama-3.1-8b-instant"

    try:
        completion = client.chat.completions.create(
            model=model_to_use,
            messages=messages,
            temperature=0.1,
            response_format={"type": "json_object"}
        )
        return json.loads(completion.choices[0].message.content)
    except Exception as e:
        st.error(f"Errore durante la chiamata a Groq AI: {e}")
        return None

# ---------------------------------------------------------
# LOGICA DI SALVATAGGIO NEL DATABASE
# ---------------------------------------------------------
def salva_dati_estratti(dati, nome_file):
    if not dati or not dati.get("nome") or not dati.get("cognome"):
        return False

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cf = dati.get("codice_fiscale")
    # Cerca lavoratore esistente per CF o per Nome e Cognome
    if cf:
        cursor.execute("SELECT id FROM lavoratori WHERE codice_fiscale = ?", (cf,))
    else:
        cursor.execute("SELECT id FROM lavoratori WHERE nome = ? AND cognome = ?", (dati.get("nome"), dati.get("cognome")))
    
    row = cursor.fetchone()
    
    if row:
        lavoratore_id = row[0]
    else:
        cursor.execute('''
            INSERT INTO lavoratori (nome, cognome, codice_fiscale, mansione, stato)
            VALUES (?, ?, ?, ?, ?)
        ''', (dati.get("nome"), dati.get("cognome"), cf, dati.get("mansione"), "ABILITATO"))
        lavoratore_id = cursor.lastrowid

    # Inserisci il certificato associato
    cursor.execute('''
        INSERT INTO certificati (lavoratore_id, nome_corso, data_emissione, data_scadenza, file_name)
        VALUES (?, ?, ?, ?, ?)
    ''', (
        lavoratore_id, 
        dati.get("nome_corso", "Corso Generico"), 
        dati.get("data_emissione"), 
        dati.get("data_scadenza"), 
        nome_file
    ))
    
    conn.commit()
    conn.close()
    return True

# ---------------------------------------------------------
# INTERFACCIA UTENTE STREAMLIT
# ---------------------------------------------------------
st.title("🛡️ Dashboard Sicurezza CSE - Gestione Attestati")
st.write("Carica i PDF riepiloghi o singoli attestati dei lavoratori. Il sistema leggerà **ogni singola pagina** autonomamente.")

menu = st.sidebar.selectbox("Navigazione", ["Carica Documenti", "Elenco Lavoratori & Stato"])

if menu == "Carica Documenti":
    st.subheader("📤 Caricamento e Analisi IA Multipagina")
    
    uploaded_files = st.file_uploader("Seleziona file PDF o immagini", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if uploaded_files and st.button("Avvia Analisi e Importazione"):
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.read()
            nome_basso = uploaded_file.name.lower()
            
            with st.spinner(f"Elaborazione in corso per: {uploaded_file.name}..."):
                if nome_basso.endswith(".pdf"):
                    pagine = estrai_pagine_da_pdf(file_bytes)
                    
                    for pag in pagine:
                        testo = pag["testo"]
                        img = pag["immagine"]
                        num_pag = pag["numero_pagina"]
                        
                        if not testo and not img:
                            continue
                            
                        dati_ai = analizza_contenuto_con_ai(testo, img)
                        if dati_ai:
                            nome_file_pagina = f"{uploaded_file.name} (Pag. {num_pag})"
                            salva_dati_estratti(dati_ai, nome_file_pagina)
                            
                    st.success(f"Completato file multipagina: {uploaded_file.name}")
                    
                else:
                    # Gestione immagini singole (png, jpg)
                    img_base64 = base64.b64encode(file_bytes).decode('utf-8')
                    dati_ai = analizza_contenuto_con_ai("", img_base64)
                    if dati_ai:
                        salva_dati_estratti(dati_ai, uploaded_file.name)
                    st.success(f"Completato file: {uploaded_file.name}")
                    
        st.balloons()
        st.info("Tutti i documenti sono stati elaborati e salvati nel database!")

elif menu == "Elenco Lavoratori & Stato":
    st.subheader("👥 Gestione Lavoratori e Certificati")
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome, cognome, codice_fiscale, mansione, stato FROM lavoratori")
    lavoratori = cursor.fetchall()
    
    if not lavoratori:
        st.warning("Nessun lavoratore presente nel database. Inizia caricando dei documenti.")
    else:
        for lav in lavoratori:
            lav_id, nome, cognome, cf, mansione, stato = lav
            with st.expander(f"{cognome} {nome} - Mansione: {mansione or 'ND'} (CF: {cf or 'ND'})"):
                cursor.execute("SELECT id, nome_corso, data_emissione, data_scadenza, file_name FROM certificati WHERE lavoratore_id = ?", (lav_id,))
                certificati = cursor.fetchall()
                
                if certificati:
                    st.markdown("**Certificati Associati:**")
                    for cert in certificati:
                        c_id, corso, emissione, scadenza, f_name = cert
                        st.write(f"- **{corso}** | Emesso: {emissione or 'ND'} | Scadenza: {scadenza or 'ND'} | *File: {f_name}*")
                else:
                    st.info("Nessun certificato registrato per questo lavoratore.")
                    
    conn.close()
