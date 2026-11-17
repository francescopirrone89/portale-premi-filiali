import streamlit as st
import pandas as pd
import openpyxl
from supabase import create_client
import io

# Configurazione pagina
st.set_page_config(page_title="Portale Premi Filiali", page_icon="📊", layout="wide")

# Ordine fisso delle colonne interne
COLONNE_FILIALE = [
    "Operatore",
    "T+KM WI (RA)", "T+KM WI (€)", "T+KM WI (50%)",
    "GOLD (Q)", "GOLD (€)",
    "CPP (Q)", "CPP (€)",
    "RAP (Q)", "RAP (€)",
    "PAI (Q)", "PAI (€)",
    "UPSELL (Q)", "UPSELL (€)",
    "SEC. GUIDA (Q)", "SEC. GUIDA (€)",
    "Totale (€)", "% Operatore", "Premio Lordo (€)"
]

# Connessione a Supabase tramite Secrets
@st.cache_resource
def init_supabase():
    url = st.secrets["SUPABASE_URL"].strip()
    key = st.secrets["SUPABASE_KEY"].strip()
    return create_client(url, key)

def init_supabase_admin():
    url = st.secrets["SUPABASE_URL"].strip()
    service_key = st.secrets["SUPABASE_SERVICE_KEY"].strip()
    return create_client(url, service_key)

supabase = init_supabase()

# Gestione stato sessione
if "user" not in st.session_state:
    st.session_state.user = None
if "profilo" not in st.session_state:
    st.session_state.profilo = None

# --- FUNZIONI DI FORMATTAZIONE E PARSING EXCEL ---
def to_float(val):
    if val is None or str(val).strip() == "":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("€", "").replace("%", "").strip()
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def fmt_euro(val):
    num = to_float(val)
    if num is None:
        return "-"
    return f"{num:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_num(val, dec=2):
    num = to_float(val)
    if num is None:
        return "-"
    if dec == 0:
        return f"{int(round(num)):,}".replace(",", ".")
    return f"{num:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_perc(val):
    num = to_float(val)
    if num is None:
        return "-"
    if abs(num) <= 2.0 and num != 0:
        num = num * 100
    return f"{num:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def parse_foglio_filiale(ws):
    righe_tabella = []
    kpi = {"perc_fatturato": "-", "risultato_nps": "-", "premio_lordo": "-", "vs_obj": ""}

    for r in range(4, ws.max_row + 1):
        nome = ws.cell(row=r, column=2).value
        if nome is not None and str(nome).strip() != "":
            nome_str = str(nome).strip()
            if nome_str.upper() == "TOTALI":
                break
            
            is_totale = "totale" in nome_str.lower()
            riga_dict = {
                "Operatore": nome_str,
                "T+KM WI (RA)": fmt_num(ws.cell(row=r, column=3).value, 0),
                "T+KM WI (€)": fmt_euro(ws.cell(row=r, column=4).value),
                "T+KM WI (50%)": fmt_euro(ws.cell(row=r, column=5).value),
                "GOLD (Q)": fmt_num(ws.cell(row=r, column=6).value, 0 if is_totale else 2),
                "GOLD (€)": fmt_euro(ws.cell(row=r, column=7).value),
                "CPP (Q)": fmt_num(ws.cell(row=r, column=8).value, 0 if is_totale else 2),
                "CPP (€)": fmt_euro(ws.cell(row=r, column=9).value),
                "RAP (Q)": fmt_num(ws.cell(row=r, column=10).value, 0 if is_totale else 2),
                "RAP (€)": fmt_euro(ws.cell(row=r, column=11).value),
                "PAI (Q)": fmt_num(ws.cell(row=r, column=12).value, 0 if is_totale else 2),
                "PAI (€)": fmt_euro(ws.cell(row=r, column=13).value),
                "UPSELL (Q)": fmt_num(ws.cell(row=r, column=14).value, 0 if is_totale else 2),
                "UPSELL (€)": fmt_euro(ws.cell(row=r, column=15).value),
                "SEC. GUIDA (Q)": fmt_num(ws.cell(row=r, column=16).value, 0 if is_totale else 2),
                "SEC. GUIDA (€)": fmt_euro(ws.cell(row=r, column=17).value),
                "Totale (€)": fmt_euro(ws.cell(row=r, column=18).value),
                "% Operatore": fmt_num(ws.cell(row=r, column=19).value, 0 if is_totale else 2) + "%",
                "Premio Lordo (€)": fmt_euro(ws.cell(row=r, column=20).value),
            }
            righe_tabella.append(riga_dict)
            if is_totale:
                break

    for r in range(25, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            if "Perc. Fatturato" in val:
                kpi["perc_fatturato"] = fmt_perc(ws.cell(row=r + 1, column=c).value)
                kpi["risultato_nps"] = fmt_perc(ws.cell(row=r + 1, column=c + 2).value)
                kpi["premio_lordo"] = fmt_euro(ws.cell(row=r + 1, column=c + 4).value)
                obj_val = ws.cell(row=r + 2, column=c + 2).value
                if obj_val is not None and str(obj_val).strip() != "":
                    kpi["vs_obj"] = str(obj_val).strip()
                break

    return {"righe": righe_tabella, "kpi": kpi}

def genera_tabella_html_unita(df):
    """Genera una tabella HTML con intestazioni unite su 2 livelli identiche al foglio Excel."""
    html = """
    <style>
    .table-wrapper {
        overflow-x: auto;
        border-radius: 8px;
        border: 1px solid #3a3f4b;
        margin-top: 8px;
        margin-bottom: 16px;
    }
    .premi-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        font-size: 13px;
        white-space: nowrap;
    }
    .premi-table th {
        background-color: #c6e0b4;
        color: #111111;
        font-weight: 700;
        text-align: center;
        padding: 7px 10px;
        border: 1px solid #7f9c6b;
    }
    .premi-table td {
        padding: 7px 10px;
        border: 1px solid #3a3f4b;
        text-align: right;
        color: #f0f2f6;
        background-color: #161920;
    }
    .premi-table td.op-cell {
        text-align: left;
        font-weight: 600;
        position: sticky;
        left: 0;
        background-color: #1e222b;
        z-index: 1;
    }
    .premi-table td.col-tot {
        background-color: #1d2633;
        font-weight: 600;
    }
    .premi-table tr.totale-row td {
        background-color: #d9ead3 !important;
        color: #111111 !important;
        font-weight: 700;
        border-top: 2px solid #5b7a47;
    }
    .premi-table tr.totale-row td.premio-cell {
        background-color: #f4cccc !important;
    }
    </style>
    <div class="table-wrapper">
    <table class="premi-table">
        <thead>
            <tr>
                <th rowspan="2" style="position: sticky; left: 0; z-index: 2;">Operatore</th>
                <th colspan="3">T+KM WI</th>
                <th colspan="2">GOLD</th>
                <th colspan="2">CPP</th>
                <th colspan="2">RAP</th>
                <th colspan="2">PAI</th>
                <th colspan="2">UPSELL</th>
                <th colspan="2">SEC. GUIDA</th>
                <th rowspan="2">Totale</th>
                <th rowspan="2">%<br>Operatore</th>
                <th rowspan="2">Premio di<br>risultato Lordo</th>
            </tr>
            <tr>
                <th>RA TOT</th><th>€</th><th>50%</th>
                <th>QTÀ</th><th>€</th>
                <th>QTÀ</th><th>€</th>
                <th>QTÀ</th><th>€</th>
                <th>QTÀ</th><th>€</th>
                <th>QTÀ</th><th>€</th>
                <th>QTÀ</th><th>€</th>
            </tr>
        </thead>
        <tbody>
    """
    for _, row in df.iterrows():
        is_tot = "totale" in str(row["Operatore"]).lower()
        tr_class = ' class="totale-row"' if is_tot else ''
        html += f"<tr{tr_class}>"
        for col in COLONNE_FILIALE:
            val = row.get(col, "-")
            if col == "Operatore":
                html += f'<td class="op-cell">{val}</td>'
            elif col == "Totale (€)":
                html += f'<td class="col-tot">{val}</td>'
            elif col == "Premio Lordo (€)":
                html += f'<td class="premio-cell">{val}</td>'
            else:
                html += f"<td>{val}</td>"
        html += "</tr>"

    html += """
        </tbody>
    </table>
    </div>
    """
    return html

def render_scheda_pulita(dati_json, nome_scheda, mese_rif):
    """Visualizza la scheda formattata con intestazioni unite e KPI."""
    if isinstance(dati_json, dict) and "righe" in dati_json:
        righe = dati_json["righe"]
        kpi = dati_json.get("kpi", {})
        
        if kpi and any(v != "-" for v in kpi.values()):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("📈 Perc. Fatturato", kpi.get("perc_fatturato", "-"))
            with c2:
                delta_txt = kpi.get("vs_obj", "")
                st.metric("⭐ Risultato NPS", kpi.get("risultato_nps", "-"), delta=delta_txt if delta_txt else None, delta_color="off")
            with c3:
                st.metric("🏆 Premio Lordo Filiale", kpi.get("premio_lordo", "-"))
            st.divider()

        df = pd.DataFrame(righe)
        if not df.empty:
            colonne_presenti = [c for c in COLONNE_FILIALE if c in df.columns]
            df = df[colonne_presenti]
            st.markdown(genera_tabella_html_unita(df), unsafe_allow_html=True)
            return df
    else:
        df = pd.DataFrame(dati_json)
        st.dataframe(df, use_container_width=True, hide_index=True)
        return df

# --- FUNZIONI AUTH ---
def login(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        prof = supabase.table("profili_filiali").select("*").eq("id", res.user.id).single().execute()
        st.session_state.profilo = prof.data
        st.rerun()
    except Exception:
        st.error("Credenziali non valide o profilo non configurato.")

def logout():
    supabase.auth.sign_out()
    st.session_state.user = None
    st.session_state.profilo = None
    st.rerun()

# --- SCHERMATA DI LOGIN ---
if not st.session_state.user:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.title("📊 Portale Riepilogo Premi")
        st.markdown("Accedi con le tue credenziali aziendali per visualizzare i dati della tua filiale.")
        with st.form("login_form"):
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Accedi", use_container_width=True)
            if submitted:
                login(email, password)

# --- AREA RISERVATA (DOPO IL LOGIN) ---
else:
    profilo = st.session_state.profilo
    ruolo = profilo.get("ruolo", "filiale")
    
    with st.sidebar:
        st.write(f"👤 **{profilo['email']}**")
        st.write(f"Ruolo: `{ruolo.upper()}`")
        if ruolo == "filiale":
            st.write(f"Filiale: **{profilo['nome_scheda_excel']}**")
        st.divider()
        if st.button("Esci (Logout)", use_container_width=True):
            logout()

    # ==========================================
    # SEZIONE ADMIN
    # ==========================================
    if ruolo == "admin":
        st.title("🛠️ Pannello Amministrazione")
        tab1, tab2, tab3 = st.tabs(["📤 Carica Excel Mensile", "👥 Gestione Utenti e Filiali", "👀 Gestione e Anteprima Dati"])

        # --- TAB 1: CARICAMENTO EXCEL ---
        with tab1:
            st.subheader("Carica il riepilogo mensile")
            st.info("Il sistema pulirà automaticamente le righe vuote, arrotonderà gli importi ed estrarrà i KPI per ogni filiale.")
            
            mese_rif = st.text_input("Mese / Periodo di riferimento (es. Ottobre 2026)", placeholder="Es. Ottobre 2026")
            uploaded_file = st.file_uploader("Seleziona il file Excel (.xlsx)", type=["xlsx"], key="upload_premi")

            if uploaded_file and mese_rif:
                if st.button("Elabora e Pubblica Schede", type="primary"):
                    try:
                        file_bytes = uploaded_file.getvalue()
                        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                        schede_elaborate = []
                        admin_client = init_supabase_admin()
                        
                        admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_rif.strip()).execute()

                        for sheet_name in wb.sheetnames:
                            ws = wb[sheet_name]
                            test_header = str(ws.cell(row=2, column=3).value or "") + str(ws.cell(row=3, column=3).value or "")
                            if "RA TOT" in test_header or "T+KM" in test_header:
                                dati_salvare = parse_foglio_filiale(ws)
                            else:
                                df_gen = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
                                df_gen = df_gen.dropna(how="all").dropna(axis=1, how="all").fillna("")
                                df_gen.columns = [str(c) if not str(c).startswith("Unnamed") else "" for c in df_gen.columns]
                                dati_salvare = df_gen.astype(str).to_dict(orient="records")

                            admin_client.table("dati_premi").insert({
                                "mese_riferimento": mese_rif.strip(),
                                "nome_scheda": sheet_name.strip(),
                                "dati_json": dati_salvare
                            }).execute()
                            schede_elaborate.append(sheet_name.strip())

                        st.success(f"✅ Pubblicazione completata per **{mese_rif}**!")
                        st.write("Schede divise e salvate:", ", ".join(schede_elaborate))
                    except Exception as e:
                        st.error(f"Errore durante l'elaborazione del file: {e}")

        # --- TAB 2: GESTIONE ACCOUNT ---
        with tab2:
            st.subheader("📥 Importazione massiva utenti da Excel")
            st.markdown("Carica un file `.xlsx` con esattamente queste 3 colonne: **`email`**, **`password`** (min. 6 caratteri) e **`nome_scheda_excel`**.")
            
            df_modello = pd.DataFrame([
                {"email": "filiale.milazzo@azienda.it", "password": "Password123!", "nome_scheda_excel": "Milazzo"},
                {"email": "filiale.roma@azienda.it", "password": "Password456!", "nome_scheda_excel": "Roma"}
            ])
            buf_modello = io.BytesIO()
            with pd.ExcelWriter(buf_modello, engine='openpyxl') as writer:
                df_modello.to_excel(writer, index=False, sheet_name="Utenti")
            
            st.download_button(
                label="📄 Scarica file Excel di esempio",
                data=buf_modello.getvalue(),
                file_name="modello_importazione_utenti.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            file_utenti = st.file_uploader("Carica l'elenco utenti (.xlsx)", type=["xlsx"], key="upload_utenti")
            if file_utenti:
                if st.button("Importa e Crea Tutti gli Utenti", type="primary"):
                    try:
                        df_imp = pd.read_excel(file_utenti)
                        df_imp.columns = [str(c).strip().lower() for c in df_imp.columns]
                        
                        colonne_richieste = {"email", "password", "nome_scheda_excel"}
                        if not colonne_richieste.issubset(set(df_imp.columns)):
                            st.error("Il file Excel deve contenere le colonne: email, password, nome_scheda_excel")
                        else:
                            df_imp = df_imp.dropna(subset=["email", "password", "nome_scheda_excel"])
                            admin_client = init_supabase_admin()
                            creati = 0
                            errori = []
                            
                            progress_bar = st.progress(0)
                            totale = len(df_imp)
                            
                            for idx, row in df_imp.iterrows():
                                em = str(row["email"]).strip()
                                pw = str(row["password"]).strip()
                                sc = str(row["nome_scheda_excel"]).strip()
                                
                                if len(pw) < 6:
                                    errori.append(f"{em}: password inferiore a 6 caratteri")
                                    continue
                                
                                try:
                                    new_user = admin_client.auth.admin.create_user({
                                        "email": em,
                                        "password": pw,
                                        "email_confirm": True
                                    })
                                    admin_client.table("profili_filiali").insert({
                                        "id": new_user.user.id,
                                        "email": em,
                                        "ruolo": "filiale",
                                        "nome_scheda_excel": sc
                                    }).execute()
                                    creati += 1
                                except Exception as ex_u:
                                    errori.append(f"{em}: {ex_u}")
                                
                                progress_bar.progress((idx + 1) / totale)
                            
                            if creati > 0:
                                st.success(f"✅ Importazione completata: **{creati}** nuovi utenti creati!")
                            if errori:
                                st.warning("Alcune righe non sono state importate:")
                                for err in errori:
                                    st.write(f"- {err}")
                    except Exception as e:
                        st.error(f"Errore nella lettura del file Excel: {e}")

            st.divider()

            with st.expander("➕ Crea singolo account manualmente", expanded=False):
                with st.form("nuovo_utente_form"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        nuova_email = st.text_input("Email utente")
                        nuova_pass = st.text_input("Password temporanea (min. 6 caratteri)", type="password")
                    with col_b:
                        nome_scheda = st.text_input("Nome esatto della scheda Excel (es. Milazzo)", help="Deve coincidere con il nome del foglio nel file Excel.")
                        nuovo_ruolo = st.selectbox("Ruolo", ["filiale", "admin"])
                    
                    crea_btn = st.form_submit_button("Crea Utente")
                    if crea_btn:
                        if nuova_email and len(nuova_pass) >= 6 and nome_scheda:
                            try:
                                admin_client = init_supabase_admin()
                                new_user = admin_client.auth.admin.create_user({
                                    "email": nuova_email.strip(),
                                    "password":
