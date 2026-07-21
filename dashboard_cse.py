import streamlit as st
import pandas as pd
import sqlite3
import json
from google import genai
from google.genai import types
import os
import dropbox
import tempfile
import zipfile

# ---------------------------------------------------------
# 1. CONFIGURAZIONE PAGINA & CSS CUSTOM
# ---------------------------------------------------------
st.set_page_config(
    page_title="CSE Master Control - Sicurezza Cantiere",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1E3A8A;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #4B5563;
        margin-bottom: 2rem;
    }
    .badge-ok { background-color: #DEF7EC; color: #03543F; padding: 4px 12px; border-radius: 12px; font-weight: bold; }
    .badge-warn { background-color: #FEF08A; color: #713F12; padding: 4px 12px; border-radius: 12px; font-weight: bold; }
    .badge-danger { background-color: #FDE8E8; color: #9B1C1C; padding: 4px 12px; border-radius: 12px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. INIZIALIZZAZIONE CLIENT API & DATABASE (SICURA)
# ---------------------------------------------------------
try:
    gemini_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=gemini_key)
except Exception as e:
    st.error("⚠️ Chiave GEMINI_API_KEY non trovata nei Secrets di Streamlit.")
    st.stop()

# Connessione al DB SQLite con timeout esteso
conn = sqlite3.connect("database_sicurezza.db", check_same_thread=False, timeout=20)
cursor = conn.cursor()

def inizializza_db():
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS aziende (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        nome TEXT UNIQUE,
        percorso_dropbox TEXT DEFAULT ''
    )
    """)
    
    # Controllo sicuro per la colonna percorso_dropbox
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

# ---------------------------------------------------------
# 3. HELPER DROPBOX
# ---------------------------------------------------------
def get_dropbox_client():
    token = st.secrets.get("DROPBOX_TOKEN", None)
    if token:
        return dropbox.Dropbox(token)
    return None

def scarica_file_da_dropbox(dbx, path_dropbox):
    """Scarica un file temporaneo da Dropbox e ne restituisce il percorso locale"""
    try:
        _, res = dbx.files_download(path_dropbox)
        ext = os.path.splitext(path_dropbox)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(res.content)
            return tmp.name
    except Exception as e:
        return None

def esplora_e_elabora_dropbox(dbx, path_cartella, id_azienda, status_placeholder):
    """Esplora la cartella Dropbox ed elabora PDF/ZIP con Gemini"""
    file_processati = 0
    
    try:
        res = dbx.files_list_folder(path_cartella, recursive=True)
        entries = res.entries
        
        while res.has_more:
            res = dbx.files_list_folder_continue(res.cursor)
            entries.extend(res.entries)
            
        file_validi = [e for e in entries if isinstance(e, dropbox.files.FileMetadata) 
                       and e.name.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.zip'))]
        
        total = len(file_validi)
        
        for idx, entry in enumerate(file_validi):
            status_placeholder.info(f"⏳ Elaborazione file ({idx+1}/{total}): **{entry.name}**")
            
            temp_path = scarica_file_da_dropbox(dbx, entry.path_lower)
            if not temp_path:
                continue
                
            if entry.name.lower().endswith('.zip'):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    try:
                        with zipfile.ZipFile(temp_path, 'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)
                        for root, _, files in os.walk(tmp_dir):
                            for f in files:
                                if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                                    fp = os.path.join(root, f)
                                    processa_file_locale(fp, f, id_azienda)
                                    file_processati += 1
                    except Exception:
                        pass
            else:
                processa_file_locale(temp_path, entry.name, id_azienda)
                file_processati += 1
                
            os.remove(temp_path)
            
    except Exception as e:
        st.error(f"Errore durante l'accesso alla cartella Dropbox: {e}")
        
    return file_processati

def processa_file_locale(file_path, file_name, id_azienda):
    """Analizza il singolo file con l'API Gemini e salva su SQLite"""
    try:
        with open(file_path, "rb") as f:
            content = f.read()
            
        ext = os.path.splitext(file_name)[1].lower()
        mime_type = "application/pdf" if ext == ".pdf" else "image/jpeg"
        
        uploaded_file = client.files.upload(
            file=content,
            config=types.UploadFileConfig(mime_type=mime_type)
        )
        
        prompt = """
        Sei un esperto CSE (Coordinatore Sicurezza Esecuzione).
        Analizza questo documento relativo alla sicurezza sul lavoro ed estrai in formato JSON:
        - "nominativo": Nome e Cognome del lavoratore
        - "mansione": Mansione del lavoratore (se presente, altrimenti null)
        - "tipo_documento": Es. "Idoneità Medica", "Formazione Generale", "Accordo Stato Regioni", "Antincendio", "Primo Soccorso", "PLE", etc.
        - "stato_scadenza": Solo uno tra "REGOLARE", "IN SCADENZA", "SCADUTO"
        - "data_scadenza": Data formato AAAA-MM-GG (o null)
        - "prescrizioni_mediche": Eventuali prescrizioni mediche della visita se è una visita medica, altrimenti "Nessuna"
        
        Rispondi ESCLUSIVAMENTE con un oggetto JSON valido.
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[uploaded_file, prompt]
        )
        
        raw_text = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(raw_text)
        
        if data.get("nominativo"):
            salva_dati_nel_db(
                id_azienda,
                data["nominativo"],
                data.get("mansione", "Non specificata"),
                data.get("tipo_documento", "Documento"),
                data.get("stato_scadenza", "REGOLARE"),
                data.get("data_scadenza", ""),
                data.get("prescrizioni_mediche", "Nessuna")
            )
    except Exception as e:
        pass

def salva_dati_nel_db(azienda_id, nominativo, mansione, tipo_doc, stato_scadenza, data_scadenza, prescrizioni):
    """Inserisce o aggiorna i record nel DB SQLite"""
    try:
        cursor.execute("""
            INSERT INTO lavoratori (azienda_id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(azienda_id, nominativo) DO UPDATE SET
            mansione=excluded.mansione,
            prescrizioni_mediche=CASE WHEN excluded.prescrizioni_mediche != 'Nessuna' THEN excluded.prescrizioni_mediche ELSE prescrizioni_mediche END
        """, (azienda_id, nominativo, mansione, stato_scadenza, prescrizioni))
        
        cursor.execute("SELECT id FROM lavoratori WHERE azienda_id=? AND nominativo=?", (azienda_id, nominativo))
        lavoratore_id = cursor.fetchone()[0]
        
        cursor.execute("""
            INSERT INTO documenti_lavoratori (lavoratore_id, tipo_documento, stato_scadenza, data_scadenza)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(lavoratore_id, tipo_documento) DO UPDATE SET
            stato_scadenza=excluded.stato_scadenza,
            data_scadenza=excluded.data_scadenza
        """, (lavoratore_id, tipo_doc, stato_scadenza, data_scadenza))
        
        # Aggiorna lo stato globale del lavoratore
        cursor.execute("SELECT stato_scadenza FROM documenti_lavoratori WHERE lavoratore_id=?", (lavoratore_id,))
        stati = [r[0] for r in cursor.fetchall()]
        
        if "SCADUTO" in stati:
            stato_globale = "SCADUTO"
        elif "IN SCADENZA" in stati:
            stato_globale = "IN SCADENZA"
        else:
            stato_globale = "REGOLARE"
            
        cursor.execute("UPDATE lavoratori SET stato_scadenza_totale=? WHERE id=?", (stato_globale, lavoratore_id))
        conn.commit()
    except Exception as e:
        pass

# ---------------------------------------------------------
# 4. BARRA LATERALE (SIDEBAR)
# ---------------------------------------------------------
st.sidebar.image("https://img.icons8.com/color/96/000000/worker-with-roadblock.png", width=70)
st.sidebar.title("CSE Control Center")

# --- SELEZIONE/CREAZIONE AZIENDA ---
st.sidebar.markdown("### 🏢 Configurazione Cantiere")

cursor.execute("SELECT id, nome, percorso_dropbox FROM aziende")
aziende = cursor.fetchall()

mappa_aziende = {a[1]: {"id": a[0], "path": a[2]} for a in aziende}
nomi_aziende = list(mappa_aziende.keys())

azienda_selezionata = st.sidebar.selectbox("Seleziona l'azienda", ["--- Scegli Azienda ---"] + nomi_aziende)

with st.sidebar.expander("➕ Aggiungi / Modifica Azienda"):
    nuova_azienda = st.text_input("Aggiungi Nuova Ditta")
    percorso_dropbox = st.text_input("Percorso Cartella Dropbox Ditta", placeholder="/CANTIERE/DITTE/SUBAPPALTI/REEVIVE")
    
    if st.button("Salva Azienda"):
        if nuova_azienda:
            try:
                cursor.execute(
                    "INSERT INTO aziende (nome, percorso_dropbox) VALUES (?, ?) ON CONFLICT(nome) DO UPDATE SET percorso_dropbox=excluded.percorso_dropbox", 
                    (nuova_azienda, percorso_dropbox)
                )
                conn.commit()
                st.success(f"Azienda {nuova_azienda} salvata!")
                st.rerun()
            except Exception as e:
                st.error(f"Errore durante il salvataggio: {e}")

# --- UPLOAD MANUALE ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 📤 Upload Manuale (PDF/ZIP)")
uploaded_files = st.sidebar.file_uploader(
    "Carica File o Archivio ZIP", 
    type=["pdf", "png", "jpg", "jpeg", "zip"], 
    accept_multiple_files=True
)

if uploaded_files and azienda_selezionata != "--- Scegli Azienda ---":
    if st.sidebar.button("⚡ Analizza File Caricati"):
        id_az = mappa_aziende[azienda_selezionata]["id"]
        with st.spinner("Analisi manuale in corso..."):
            for uf in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uf.name)[1]) as tmp:
                    tmp.write(uf.getbuffer())
                    tmp_path = tmp.name
                
                if uf.name.lower().endswith(".zip"):
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        with zipfile.ZipFile(tmp_path, 'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)
                        for root, _, files in os.walk(tmp_dir):
                            for f in files:
                                if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                                    processa_file_locale(os.path.join(root, f), f, id_az)
                else:
                    processa_file_locale(tmp_path, uf.name, id_az)
                os.remove(tmp_path)
            st.sidebar.success("Upload ed elaborazione completati!")
            st.rerun()

# ---------------------------------------------------------
# 5. DASHBOARD PRINCIPALE
# ---------------------------------------------------------
st.markdown('<div class="main-header">🛡️ Dashboard Coordinatore Sicurezza (CSE)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Monitoraggio Conformità Documentale Cantiere & Attestati</div>', unsafe_allow_html=True)

if azienda_selezionata == "--- Scegli Azienda ---":
    st.info("👈 Seleziona o crea un'azienda dal menu a sinistra per iniziare.")
else:
    info_az = mappa_aziende[azienda_selezionata]
    id_azienda = info_az["id"]
    path_dropbox_az = info_az["path"]
    
    st.subheader(f"Ditta: {azienda_selezionata}")
    
    # --- BOX SCANSIONE DROPBOX ---
    with st.container():
        st.markdown("##### 📦 Analisi Automatica Cartella Dropbox")
        if path_dropbox_az:
            st.text(f"Percorso collegato: {path_dropbox_az}")
            if st.button("🚀 SCANSIONA ED ELABORA TUTTI I FILE DELLA CARTELLA DROPBOX", type="primary"):
                dbx = get_dropbox_client()
                if not dbx:
                    st.error("⚠️ Token Dropbox non configurato nei Secrets (`DROPBOX_TOKEN`).")
                else:
                    status_placeholder = st.empty()
                    count = esplora_e_elabora_dropbox(dbx, path_dropbox_az, id_azienda, status_placeholder)
                    status_placeholder.success(f"🎉 Scansione completata! Elaborati {count} file.")
                    st.rerun()
        else:
            st.warning("Nessuna cartella Dropbox associata a questa ditta. Inserisci il percorso nella barra a sinistra se vuoi la scansione automatica.")

    st.markdown("---")

    # --- DATI E METRICHE ---
    cursor.execute("SELECT id, nominativo, mansione, stato_scadenza_totale, prescrizioni_mediche FROM lavoratori WHERE azienda_id=?", (id_azienda,))
    lavoratori = cursor.fetchall()
    
    total_lav = len(lavoratori)
    tot_ok = sum(1 for l in lavoratori if l[3] == "REGOLARE")
    tot_warn = sum(1 for l in lavoratori if l[3] == "IN SCADENZA")
    tot_danger = sum(1 for l in lavoratori if l[3] == "SCADUTO")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Totale Lavoratori", total_lav)
    m2.metric("🟢 Regolari", tot_ok)
    m3.metric("🟡 In Scadenza", tot_warn)
    m4.metric("🔴 Scaduti / Non Idonei", tot_danger)
    
    st.markdown("### 📋 Registro Idoneità e Attestati Lavoratori")
    
    if total_lav == 0:
        st.info("Nessun lavoratore registrato per questa impresa. Lancia una scansione da Dropbox o carica un file PDF/ZIP.")
    else:
        for lav in lavoratori:
            lav_id, nom, mans, stato_tot, prescrizioni = lav
            
            # Colore stato
            if stato_tot == "REGOLARE":
                badge = '<span class="badge-ok">REGOLARE</span>'
            elif stato_tot == "IN SCADENZA":
                badge = '<span class="badge-warn">IN SCADENZA</span>'
            else:
                badge = '<span class="badge-danger">DOCUMENTI SCADUTI</span>'
                
            with st.expander(f"👤 **{nom}** - Mansione: *{mans}* | Stato: {stato_tot}", expanded=(stato_tot != "REGOLARE")):
                st.markdown(f"**Prescrizioni / Note Mediche:** `{prescrizioni}`")
                
                cursor.execute("SELECT tipo_documento, stato_scadenza, data_scadenza FROM documenti_lavoratori WHERE lavoratore_id=?", (lav_id,))
                docs = cursor.fetchall()
                
                if docs:
                    df_docs = pd.DataFrame(docs, columns=["Documento / Attestato", "Stato Scadenza", "Data Scadenza"])
                    st.dataframe(df_docs, use_container_width=True)
                else:
                    st.write("Nessun dettaglio documento trovato.")
