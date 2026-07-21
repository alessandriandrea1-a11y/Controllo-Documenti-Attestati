import streamlit as st
import pandas as pd
import sqlite3
import json
import io
import dropbox
import docx2txt
import re
from datetime import datetime
import zoneinfo
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
            st.error(f"⚠️ Errore nel salvataggio su Dropbox: {e}")

download_db_from_dropbox()

# --- DATABASE LOCAL SQLITE E CREAZIONE TABELLE GARANTITA ---
conn = sqlite3.connect(DB_FILE_NAME, check_same_thread=False)
cursor = conn.cursor()

def inizializza_db():
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS aziende (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        nome TEXT UNIQUE
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
    return nome_pulito if nome_pulito else "SCONOSCIUTO"

def normalizza_nome_documento(testo_doc):
    if not testo_doc:
        return "Attestato Formazione Generico"
    
    t = str(testo_doc).lower()
    
    if "confinat" in t or "sospett" in t:
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
    elif "medica" in t or "idoneit" in t or "sanitar" in t:
        return "Idoneità Sanitaria"
    elif "accordo stato regioni" in t or "generale" in t or "specifica" in t:
        return "Formazione Generale / Specifica"
    else:
        return testo_doc.strip().title()

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
        file_caricato = st.file_uploader("Carica Documento (PDF, DOCX)", type=["pdf", "docx"], key=f"uploader_{st.session_state['uploader_key']}")
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
            with st.spinner("🧠 Groq AI sta analizzando il documento..."):
                try:
                    fuso_orario = zoneinfo.ZoneInfo("Europe/Rome")
                    data_oggi = datetime.now(fuso_orario).strftime("%d/%m/%Y")
                    
                    file_bytes = file_caricato.read()
                    nome_file = file_caricato.name.lower()
                    client = Groq(api_key=api_key_inserita)
                    
                    testo_estratto = ""
                    
                    if nome_file.endswith(".pdf"):
                        pdf = pdfium.PdfDocument(file_bytes)
                        for page in pdf:
                            textpage = page.get_textpage()
                            testo_estratto += textpage.get_text_range() + "\n"
                    elif nome_file.endswith(".docx"):
                        testo_estratto = docx2txt.process(io.BytesIO(file_bytes))

                    testo_estratto = testo_estratto.strip()

                    if not testo_estratto:
                        st.error("📄 Il PDF caricato non contiene testo selezionabile. Salvalo come PDF stampabile/con testo.")
                    else:
                        testo_ottimizzato = testo_estratto[:3500]

                        prompt = f"""
                        Sei un esperto CSE di sicurezza sul lavoro. Analizza questo documento.
                        LA DATA ODIERNA DI RIFERIMENTO È TASSATIVAMENTE: {data_oggi}.

                        ISTRUZIONI RIGIDE PER L'ANALISI:
                        1. NOME LAVORATORE:
                           - Estrai SOLO ED ESCLUSIVAMENTE il Nome e Cognome della persona (es. "MAHMOUD ALI EZZAT ALI").
                           - NON INCLUDERE MAI il nome del corso o parentesi nel campo del lavoratore.
                        
                        2. DATA DI SCADENZA:
                           - Distingui la data di SVOLGIMENTO del corso dalla DATA DI SCADENZA.
                           - Se è un ATTESTATO DI FORMAZIONE (es. Ambienti Confinati, Antincendio, Primo Soccorso, Quota, ecc.) e non c'è scritta la scadenza, CALCOLA la scadenza aggiungendo 5 ANNI alla data del corso (es. corso del 11/05/2026 -> scadenza 11/05/2031).
                           - Se è una VISITA MEDICA e c'è scritto "scadenza Maggio 2028", usa "31/05/2028".

                        3. STATO RISPETTO A OGGI ({data_oggi}):
                           - Se la data di scadenza è FUTURA (oltre 60 giorni da {data_oggi}): rispondi "In Regola"
                           - Se scade nei prossimi 60 giorni rispetto a {data_oggi}: rispondi "In Scadenza"
                           - Se la data di scadenza è PASSATA rispetto a {data_oggi}: rispondi "Scaduto"

                        4. PRESCRIZIONI MEDICHE:
                           - Rileva prescrizioni SOLO per Idoneità Sanitaria/Visita Medica.
                           - Per gli attestati di formazione rispondi sempre "Nessuna prescrizione rilevata".

                        Rispondi ESCLUSIVAMENTE in JSON (USA SOLO TESTO SEMPLICE, SENZA EMOJI):
                        {{
                            "lavoratore": "NOME COGNOME",
                            "mansione": "MANSIONE (se assente usa 'Operaio')",
                            "documento_nome": "Nome esatto del corso/documento (es. Formazione Ambienti Confinati)",
                            "data_scadenza": "DD/MM/AAAA",
                            "stato_calcolato": "In Regola / In Scadenza / Scaduto",
                            "prescrizione_medica": "Nessuna prescrizione rilevata"
                        }}

                        TESTO DOCUMENTO:
                        {testo_ottimizzato}
                        """

                        try:
                            chat_completion = client.chat.completions.create(
                                messages=[{"role": "user", "content": prompt}],
                                model="llama-3.3-70b-versatile",
                                response_format={"type": "json_object"}
                            )
                        except Exception as req_err:
                            if "429" in str(req_err) or "rate_limit" in str(req_err):
                                chat_completion = client.chat.completions.create(
                                    messages=[{"role": "user", "content": prompt}],
                                    model="llama-3.1-8b-instant",
                                    response_format={"type": "json_object"}
                                )
                            else:
                                raise req_err
                        
                        risposta_testo = chat_completion.choices[0].message.content
                        dati_ai = json.loads(risposta_testo)
                        
                        nom_lav = pulisci_nome_rigido(dati_ai.get("lavoratore"))
                        mans_lav = (dati_ai.get("mansione") or "Operaio").strip().title()
                        
                        # Normalizzazione rigida del nome documento per evitare duplicati
                        doc_grezzo = (dati_ai.get("documento_nome") or "Attestato Formazione").strip()
                        doc_nome = normalizza_nome_documento(doc_grezzo)
                        
                        data_scad = (dati_ai.get("data_scadenza") or "Illimitato").strip()
                        
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
                            
                            cursor.execute("""
                                INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza)
                                VALUES (?, ?, ?, ?)
                                ON CONFLICT(lavoratore_id, tipo_documento) 
                                DO UPDATE SET stato_scadenza=excluded.stato_scadenza, data_scadenza=excluded.data_scadenza
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
                            
                            st.session_state["uploader_key"] += 1
                            st.success(f"🎉 Registrato con successo: **{doc_nome}** per **{nom_lav}**")
                            st.rerun()
                    
                except Exception as e:
                    st.error(f"⚠️ Errore durante l'elaborazione AI: {str(e)}")

    # --- REPERIMENTO E VISUALIZZAZIONE DATI ---
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
                
                tabella_pulita.append([doc_nome, validita, data_scad])
            
            with st.expander(f"{accesso} — 👤 {nome} ({mansione})"):
                if prescrizioni and prescrizioni != 'Nessuna prescrizione rilevata':
                    st.markdown(f'<div class="prescrizione-box">⚠️ **Prescrizioni Sanitarie / Limitazioni:** {prescrizioni}</div>', unsafe_allow_html=True)
                
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

        # PULIZIA DUPLICATI PERSONE E DOCUMENTI
        if ha_permesso_modifica:
            st.write("---")
            if st.button("🧹 PULISCI E UNIFICA TABELLE ED OPERAI DUPLICATI"):
                # 1. Normalizza i nomi delle persone
                cursor.execute("SELECT id, nominativo FROM lavoratori")
                tutti = cursor.fetchall()
                for l_id, nom in tutti:
                    nom_p = pulisci_nome_rigido(nom)
                    cursor.execute("UPDATE lavoratori SET nominativo = ? WHERE id = ?", (nom_p, l_id))
                conn.commit()
                
                # 2. Unifica persone duplicate
                cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
                az_row = cursor.fetchone()
                if az_row:
                    az_id = az_row[0]
                    cursor.execute("SELECT nominativo, COUNT(*) FROM lavoratori WHERE azienda_id = ? GROUP BY nominativo HAVING COUNT(*) > 1", (az_id,))
                    duplicati = cursor.fetchall()
                    for nom_dup, _ in duplicati:
                        cursor.execute("SELECT id FROM lavoratori WHERE azienda_id = ? AND nominativo = ? ORDER BY id ASC", (az_id, nom_dup))
                        ids = [r[0] for r in cursor.fetchall()]
                        id_principale = ids[0]
                        for id_vecchio in ids[1:]:
                            cursor.execute("UPDATE OR IGNORE documenti_lavoratori SET lavoratore_id = ? WHERE lavoratore_id = ?", (id_principale, id_vecchio))
                            cursor.execute("DELETE FROM lavoratori WHERE id = ?", (id_vecchio))
                    conn.commit()

                # 3. Normalizza e unifica i titoli dei documenti duplicati
                cursor.execute("SELECT id, tipo_documento FROM documenti_lavoratori")
                docs_tutti = cursor.fetchall()
                for d_id, d_nome in docs_tutti:
                    d_norm = normalizza_nome_documento(d_nome)
                    cursor.execute("UPDATE OR IGNORE documenti_lavoratori SET tipo_documento = ? WHERE id = ?", (d_norm, d_id))
                
                # Rimuove righe duplicate createsi dopo la normalizzazione tenendo l'ultima inserita
                cursor.execute("""
                    DELETE FROM documenti_lavoratori 
                    WHERE id NOT IN (
                        SELECT MAX(id) 
                        FROM documenti_lavoratori 
                        GROUP BY lavoratore_id, tipo_documento
                    )
                """)
                conn.commit()
                upload_db_to_dropbox()
                st.success("✨ Tutte le tabelle e le persone duplicate sono state unificate con successo!")
                st.rerun()

    else:
        st.info("Nessun lavoratore registrato per questa azienda. Passa al ruolo 'Coordinatore' per aggiungere file o aziende.")
else:
    st.info("👋 Benvenuto! Aggiungi o seleziona un'azienda dalla barra laterale a sinistra per iniziare.")
