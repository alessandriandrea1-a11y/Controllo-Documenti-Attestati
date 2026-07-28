import os
import re
import json
import sqlite3
import datetime
import io
import streamlit as st
from groq import Groq
import dropbox
import pypdf

# ==========================================
# 1. CONFIGURAZIONE E INIZIALIZZAZIONE DB
# ==========================================
DB_FILE = "gestionale_cantieri.db"

GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", ""))
DROPBOX_TOKEN = st.secrets.get("DROPBOX_TOKEN", os.environ.get("DROPBOX_TOKEN", ""))

client_groq = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

def get_db_connection():
    return sqlite3.connect(DB_FILE)

def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS aziende (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT UNIQUE NOT NULL,
                percorso_dropbox TEXT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS lavoratori (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                azienda_id INTEGER,
                nominativo TEXT NOT NULL,
                mansione TEXT,
                stato_scadenza_totale TEXT DEFAULT '🔴 Da Verificare',
                prescrizioni_mediche TEXT DEFAULT 'Nessuna prescrizione rilevata',
                FOREIGN KEY (azienda_id) REFERENCES aziende (id) ON DELETE CASCADE
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS documenti_lavoratori (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lavoratore_id INTEGER,
                tipo_documento TEXT NOT NULL,
                stato_scadenza TEXT DEFAULT '🔴 Scaduto / Assente',
                data_scadenza TEXT,
                nome_file_origine TEXT,
                FOREIGN KEY (lavoratore_id) REFERENCES lavoratori (id) ON DELETE CASCADE
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS file_processati (
                path_univoco TEXT PRIMARY KEY
            )
        """)
        conn.commit()

init_db()

# ==========================================
# 2. FUNZIONI UTILI & DROPBOX
# ==========================================
def pulisci_percorso_dropbox(link_raw):
    if not link_raw:
        return ""
    path = link_raw.strip()
    path = re.sub(r"^https?://(www\.)?dropbox\.com/home", "", path)
    path = re.sub(r"\?quickview=.*$", "", path)
    path = re.sub(r"\?dl=.*$", "", path)
    path = path.replace("%20", " ")
    if path and not path.startswith("/"):
        path = "/" + path
    return path.rstrip("/")

def get_dropbox_client():
    if DROPBOX_TOKEN:
        return dropbox.Dropbox(DROPBOX_TOKEN)
    return None

def upload_db_to_dropbox():
    dbx = get_dropbox_client()
    if dbx and os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                dbx.files_upload(f.read(), f"/{DB_FILE}", mode=dropbox.files.WriteMode.overwrite)
        except Exception:
            pass

def file_gia_processato(path_univoco):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM file_processati WHERE path_univoco = ?", (path_univoco,))
        return cursor.fetchone() is not None

def registra_file_processato(path_univoco):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT OR IGNORE INTO file_processati (path_univoco) VALUES (?)", (path_univoco,))
        conn.commit()

def pulisci_nome_rigido(nome_raw):
    if not nome_raw or str(nome_raw).lower() in ["null", "none", "sconosciuto", ""]:
        return "SCONOSCIUTO"
    nome_pulisce = re.sub(r'[^a-zA-ZàèéìòùÁÉÍÓÚàèéìòù\s\'-]', '', str(nome_raw))
    parti = [p.capitalize() for p in nome_pulisce.split() if len(p) > 1]
    return " ".join(parti) if parti else "SCONOSCIUTO"

def estrai_testo_completo_pdf(file_bytes):
    """Estrae il testo da tutte le pagine del PDF usando pypdf."""
    testo_totale = ""
    try:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for idx, page in enumerate(reader.pages):
            t_pag = page.extract_text()
            if t_pag:
                testo_totale += f"\n--- PAGINA {idx+1} ---\n" + t_pag
    except Exception as e:
        st.error(f"Errore lettura testo PDF: {e}")
    return testo_totale

# ==========================================
# 3. CALCOLO STATI E DATE
# ==========================================
def calcola_stato_e_data_python(data_scad_str, data_emiss_str, anni_val):
    oggi = datetime.date.today()
    data_finale = None

    if data_scad_str and data_scad_str != "NON_PRESENTI":
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                data_finale = datetime.datetime.strptime(data_scad_str.strip(), fmt).date()
                break
            except ValueError:
                pass

    if not data_finale and data_emiss_str and data_emiss_str != "NON_PRESENTI" and anni_val:
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt_em = datetime.datetime.strptime(data_emiss_str.strip(), fmt).date()
                try:
                    data_finale = dt_em.replace(year=dt_em.year + int(anni_val))
                except ValueError:
                    data_finale = dt_em + datetime.timedelta(days=365 * int(anni_val))
                break
            except ValueError:
                pass

    if not data_finale:
        return "🔴 Scaduto / Assente", "Non indicata"

    diff_giorni = (data_finale - oggi).days
    str_data = data_finale.strftime("%d/%m/%Y")

    if diff_giorni < 0:
        return "🔴 Scaduto", str_data
    elif diff_giorni <= 60:
        return "🟡 In Scadenza", str_data
    else:
        return "🟢 Valido", str_data

def aggiorna_stato_generale_lavoratore(lavoratore_id):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id = ?", (lavoratore_id,))
        stati = [r[0] for r in cursor.fetchall()]
        
        if not stati:
            stato_tot = "🔴 Nessun Documento"
        elif any("🔴" in s for s in stati):
            stato_tot = "🔴 Documenti Scaduti/Mancanti"
        elif any("🟡" in s for s in stati):
            stato_tot = "🟡 Documenti in Scadenza"
        else:
            stato_tot = "🟢 In Regola"
            
        cursor.execute("UPDATE lavoratori SET stato_scadenza_totale = ? WHERE id = ?", (stato_tot, lavoratore_id))
        conn.commit()

# ==========================================
# 4. ESTRAZIONE AI MULTI-DOCUMENTO
# ==========================================
def elabora_singolo_documento_con_ai(file_bytes, nome_file, azienda_selezionata, path_univoco):
    if file_gia_processato(path_univoco):
        return f"File già elaborato: {nome_file}", True

    testo_pdf = estrai_testo_completo_pdf(file_bytes)
    if not testo_pdf.strip():
        registra_file_processato(path_univoco)
        return f"Impossibile leggere il testo da {nome_file} (potrebbe essere una scansione senza testo).", False

    if not client_groq:
        return "Chiave API Groq non trovata.", False

    system_prompt = """
Sei un esperto verificatore di documenti di cantiere e sicurezza sul lavoro (CSE).
Analizza il testo estratto da un file PDF multipagina. 

ATTENZIONE: Un singolo file PDF POTREBBE CONTENERE PIÙ DOCUMENTI O ATTESTATI DIVERSI per lo stesso lavoratore (es. Visita Medica, Formazione Generale/Specifica, Antincendio, Primo Soccorso, DPI, Unilav, Tesserino, ecc.).

Estrai TUTTI i documenti/attestati rilevanti e le prescrizioni mediche trovate nel file.

Restituisci ESCLUSIVAMENTE un JSON valido (senza testo di contorno o blocchi markdown) con questa struttura:
{
    "lavoratore": "NOME COGNOME",
    "mansione": "Manovale/Carpentiere/ecc",
    "prescrizione_medica": "Eventuali prescrizioni sanitarie oppure 'Nessuna prescrizione rilevata'",
    "documenti": [
        {
            "documento_nome": "Nome Specifico del Corso o Documento (es. Certificato Medico Idoneità, Corso Formazione Generale e Specifica, Verbale Consegna DPI)",
            "data_emissione": "GG/MM/AAAA oppure NON_PRESENTI",
            "data_scadenza": "GG/MM/AAAA oppure NON_PRESENTI",
            "anni_validita": 5
        }
    ]
}
"""

    try:
        response = client_groq.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Nome File: {nome_file}\n\nTesto Documento Multipagina:\n{testo_pdf[:16000]}"}
            ],
            response_format={"type": "json_object"},
            temperature=0.1
        )

        dati_ai = json.loads(response.choices[0].message.content)

        nom_lav = pulisci_nome_rigido(dati_ai.get("lavoratore"))
        if nom_lav == "SCONOSCIUTO":
            registra_file_processato(path_univoco)
            return f"Lavoratore non identificato in {nome_file}", False

        mans_lav = (dati_ai.get("mansione") or "Operaio").strip().title()
        prescr_raw = dati_ai.get("prescrizione_medica")
        prescr = prescr_raw.strip() if (prescr_raw and str(prescr_raw).lower() != "null") else 'Nessuna prescrizione rilevata'

        lista_doc = dati_ai.get("documenti", [])
        if not lista_doc:
            registra_file_processato(path_univoco)
            return f"Nessun documento rilevante trovato in {nome_file}", False

        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
            az_row = cursor.fetchone()
            if not az_row:
                return f"Azienda {azienda_selezionata} non trovata.", False
            az_id = az_row[0]

            cursor.execute("SELECT id FROM lavoratori WHERE azienda_id = ? AND UPPER(nominativo) = ?", (az_id, nom_lav.upper()))
            operaio_db = cursor.fetchone()

            if operaio_db:
                op_id = operaio_db[0]
                if prescr != 'Nessuna prescrizione rilevata':
                    cursor.execute("UPDATE lavoratori SET prescrizioni_mediche = ? WHERE id = ?", (prescr, op_id))
            else:
                cursor.execute("""
                    INSERT INTO lavoratori (azienda_id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche)
                    VALUES (?, ?, ?, '🔴 Da Verificare', ?)
                """, (az_id, nom_lav, mans_lav, prescr))
                op_id = cursor.lastrowid

            # Salva CIASCUN documento trovato nel PDF multipagina
            count_inseriti = 0
            for doc_item in lista_doc:
                doc_nome = (doc_item.get("documento_nome") or "Attestato Generico").strip()
                data_em = doc_item.get("data_emissione", "NON_PRESENTI")

                data_scad, stato_calc = calcola_stato_e_data_python(
                    doc_item.get("data_scadenza"),
                    data_em,
                    doc_item.get("anni_validita")
                )

                cursor.execute("""
                    SELECT id FROM documenti_lavoratori 
                    WHERE lavoratore_id = ? AND nome_file_origine = ? AND tipo_documento = ?
                """, (op_id, nome_file, doc_nome))

                doc_esistente = cursor.fetchone()

                if doc_esistente:
                    cursor.execute("""
                        UPDATE documenti_lavoratori 
                        SET stato_scadenza = ?, data_scadenza = ?
                        WHERE id = ?
                    """, (stato_calc, data_scad, doc_esistente[0]))
                else:
                    cursor.execute("""
                        INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza, nome_file_origine)
                        VALUES (?, ?, ?, ?, ?)
                    """, (op_id, doc_nome, stato_calc, data_scad, nome_file))
                count_inseriti += 1

            conn.commit()

        aggiorna_stato_generale_lavoratore(op_id)
        registra_file_processato(path_univoco)
        upload_db_to_dropbox()
        return f"✨ Estratti {count_inseriti} attestati per {nom_lav} da {nome_file}", True

    except Exception as e:
        return f"Errore nell'analisi del file {nome_file}: {e}", False

# ==========================================
# 5. INTERFACCIA UTENTE STREAMLIT
# ==========================================
st.set_page_config(page_title="Gestione Cantiere CSE", layout="wide", page_icon="🏗️")
st.title("🏗️ Monitoraggio IDONEITÀ E ATTESTATI CANTIERE")

with st.sidebar:
    st.header("⚙️ Configurazione Ditte e Dropbox")
    
    with st.expander("🏢 Aggiungi / Aggiorna Ditta", expanded=False):
        nuova_azienda = st.text_input("Nome Impresa / Ditta")
        link_dbx = st.text_input("Link o Percorso Cartella Dropbox")
        if st.button("💾 Salva Azienda"):
            if nuova_azienda:
                percorso_clean = pulisci_percorso_dropbox(link_dbx)
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO aziende (nome, percorso_dropbox) VALUES (?, ?)
                        ON CONFLICT(nome) DO UPDATE SET percorso_dropbox=excluded.percorso_dropbox
                    """, (nuova_azienda.strip(), percorso_clean))
                    conn.commit()
                st.success(f"Ditta '{nuova_azienda}' salvata!")
                st.rerun()

    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, percorso_dropbox FROM aziende")
        aziende_list = cursor.fetchall()

    nomi_aziende = [a[1] for a in aziende_list]
    if nomi_aziende:
        azienda_selezionata = st.selectbox("Seleziona Impresa", nomi_aziende)
        az_info = next((a for a in aziende_list if a[1] == azienda_selezionata), None)
        percorso_az = az_info[2] if az_info else ""
        if percorso_az:
            st.caption(f"📁 Percorso Dropbox: `{percorso_az}`")
    else:
        azienda_selezionata = None
        st.warning("Inserisci prima un'azienda.")

    st.markdown("---")
    
    if azienda_selezionata:
        st.subheader("📤 Upload Manuale")
        uploaded_files = st.file_uploader("Carica file PDF", type=["pdf"], accept_multiple_files=True)
        if uploaded_files and st.button("🚀 Elabora File Caricati"):
            for f in uploaded_files:
                bytes_data = f.read()
                msg, ok = elabora_singolo_documento_con_ai(bytes_data, f.name, azienda_selezionata, f"manual_{f.name}")
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
            st.rerun()

    st.markdown("---")
    if azienda_selezionata and percorso_az:
        if st.button("🚀 SCANSIONA / RIPRENDI SCANSIONE"):
            dbx = get_dropbox_client()
            if not dbx:
                st.error("Token Dropbox non configurato!")
            else:
                with st.spinner("Scansione in corso..."):
                    try:
                        res = dbx.files_list_folder(percorso_az, recursive=True)
                        files_to_process = [entry for entry in res.entries if isinstance(entry, dropbox.files.FileMetadata) and entry.name.lower().endswith(".pdf")]
                        
                        num_f = len(files_to_process)
                        st.info(f"Trovati {num_f} PDF.")
                        
                        prog_bar = st.progress(0)
                        for idx, f_meta in enumerate(files_to_process):
                            _, response = dbx.files_download(f_meta.path_lower)
                            msg, ok = elabora_singolo_documento_con_ai(response.content, f_meta.name, azienda_selezionata, f_meta.path_lower)
                            prog_bar.progress((idx + 1) / num_f)
                        st.success("Scansione completata!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Errore scansione: {e}")

# DASHBOARD
if azienda_selezionata:
    st.subheader(f"👥 Elenco Lavoratori - Impresa: {azienda_selezionata}")

    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM aziende WHERE nome = ?", (azienda_selezionata,))
        az_id = cursor.fetchone()[0]

        cursor.execute("""
            SELECT id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche 
            FROM lavoratori WHERE azienda_id = ? ORDER BY nominativo ASC
        """, (az_id,))
        lavoratori = cursor.fetchall()

    if not lavoratori:
        st.info("Nessun lavoratore presente.")
    else:
        for lav in lavoratori:
            lav_id, nom, mans, stato_tot, prescr = lav
            
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, tipo_documento, stato_scadenza, data_scadenza, nome_file_origine 
                    FROM documenti_lavoratori WHERE lavoratore_id = ?
                """, (lav_id,))
                docs = cursor.fetchall()

            expander_title = f"{stato_tot} — 👤 {nom} ({mans}) — [{len(docs)} attestati registrati]"
            
            with st.expander(expander_title, expanded=False):
                if prescr and prescr != 'Nessuna prescrizione rilevata':
                    st.warning(f"⚠️ **Prescrizioni Sanitarie:** {prescr}")

                if docs:
                    table_data = []
                    for d in docs:
                        doc_id, t_doc, st_scad, d_scad, f_orig = d
                        table_data.append({
                            "Attestato / Visita Rilevata": t_doc,
                            "Validità AI": st_scad,
                            "Scadenza Calcolata": d_scad,
                            "File Origine": f_orig
                        })
                    st.dataframe(table_data, use_container_width=True)
                else:
                    st.write("Nessun documento registrato.")

                if st.button(f"🗑️ Rimuovi Questo Lavoratore", key=f"del_{lav_id}"):
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM lavoratori WHERE id = ?", (lav_id,))
                        conn.commit()
                    st.rerun()
