import streamlit as st
import pandas as pd
from supabase import create_client, Client
import io

# Configurazione pagina
st.set_page_config(page_title="Portale Premi Filiali", page_icon="📊", layout="wide")

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

# Funzione Login
def login(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        prof = supabase.table("profili_filiali").select("*").eq("id", res.user.id).single().execute()
        st.session_state.profilo = prof.data
        st.rerun()
    except Exception:
        st.error("Credenziali non valide o profilo non configurato.")

# Funzione Logout
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
    
    # Barra laterale
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
            st.info("Il sistema dividerà automaticamente tutte le schede presenti nel file Excel in base al loro nome.")
            
            mese_rif = st.text_input("Mese / Periodo di riferimento (es. Ottobre 2026)", placeholder="Es. Ottobre 2026")
            uploaded_file = st.file_uploader("Seleziona il file Excel (.xlsx)", type=["xlsx"], key="upload_premi")

            if uploaded_file and mese_rif:
                if st.button("Elabora e Pubblica Schede", type="primary"):
                    try:
                        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
                        schede_elaborate = []
                        admin_client = init_supabase_admin()
                        
                        admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_rif.strip()).execute()

                        for nome_scheda, df in excel_data.items():
                            df = df.dropna(how="all")
                            df = df.fillna("")
                            records = df.astype(str).to_dict(orient="records")
                            
                            admin_client.table("dati_premi").insert({
                                "mese_riferimento": mese_rif.strip(),
                                "nome_scheda": nome_scheda.strip(),
                                "dati_json": records
                            }).execute()
                            schede_elaborate.append(nome_scheda.strip())

                        st.success(f"✅ Pubblicazione completata per **{mese_rif}**!")
                        st.write("Schede divise e salvate:", ", ".join(schede_elaborate))
                    except Exception as e:
                        st.error(f"Errore durante l'elaborazione del file: {e}")

        # --- TAB 2: GESTIONE ACCOUNT ---
        with tab2:
            # --- SEZIONE A: CARICAMENTO MASSIVO DA EXCEL ---
            st.subheader("📥 Importazione massiva utenti da Excel")
            st.markdown("Carica un file `.xlsx` con esattamente queste 3 colonne: **`email`**, **`password`** (min. 6 caratteri) e **`nome_scheda_excel`**.")
            
            # Generatore file template scaricabile
            df_modello = pd.DataFrame([
                {"email": "filiale.milano@azienda.it", "password": "Password123!", "nome_scheda_excel": "Milano"},
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
                        # Normalizza nomi colonne in minuscolo senza spazi
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

            # --- SEZIONE B: CREAZIONE SINGOLO UTENTE ---
            with st.expander("➕ Crea singolo account manualmente", expanded=False):
                with st.form("nuovo_utente_form"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        nuova_email = st.text_input("Email utente")
                        nuova_pass = st.text_input("Password temporanea (min. 6 caratteri)", type="password")
                    with col_b:
                        nome_scheda = st.text_input("Nome esatto della scheda Excel (es. Milano)", help="Deve coincidere con il nome del foglio nel file Excel.")
                        nuovo_ruolo = st.selectbox("Ruolo", ["filiale", "admin"])
                    
                    crea_btn = st.form_submit_button("Crea Utente")
                    if crea_btn:
                        if nuova_email and len(nuova_pass) >= 6 and nome_scheda:
                            try:
                                admin_client = init_supabase_admin()
                                new_user = admin_client.auth.admin.create_user({
                                    "email": nuova_email.strip(),
                                    "password": nuova_pass,
                                    "email_confirm": True
                                })
                                admin_client.table("profili_filiali").insert({
                                    "id": new_user.user.id,
                                    "email": nuova_email.strip(),
                                    "ruolo": nuovo_ruolo,
                                    "nome_scheda_excel": nome_scheda.strip()
                                }).execute()
                                st.success(f"Utente {nuova_email} creato e associato alla scheda '{nome_scheda}'!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Errore nella creazione utente: {e}")
                        else:
                            st.warning("Compila tutti i campi (la password deve avere almeno 6 caratteri).")

            st.divider()

            # --- SEZIONE C: ELENCO E ELIMINAZIONE UTENTI ---
            st.subheader("Utenti registrati")
            admin_client = init_supabase_admin()
            utenti_res = admin_client.table("profili_filiali").select("id, email, ruolo, nome_scheda_excel").order("email").execute()
            
            if utenti_res.data:
                df_utenti = pd.DataFrame(utenti_res.data)[["email", "ruolo", "nome_scheda_excel"]]
                df_utenti.columns = ["Email", "Ruolo", "Scheda Excel Associata"]
                st.dataframe(df_utenti, use_container_width=True, hide_index=True)

                st.markdown("#### 🗑️ Elimina un utente")
                altri_utenti = [u for u in utenti_res.data if u["id"] != st.session_state.user.id]
                
                if altri_utenti:
                    opzioni_utenti = {f"{u['email']} ({u['nome_scheda_excel']})": u["id"] for u in altri_utenti}
                    col_del1, col_del2 = st.columns([3, 1])
                    with col_del1:
                        utente_da_eliminare = st.selectbox("Seleziona l'account da rimuovere", list(opzioni_utenti.keys()), label_visibility="collapsed")
                    with col_del2:
                        if st.button("Elimina Utente", type="secondary", use_container_width=True):
                            try:
                                uid_target = opzioni_utenti[utente_da_eliminare]
                                admin_client.auth.admin.delete_user(uid_target)
                                st.success(f"Account {utente_da_eliminare} eliminato definitivamente.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Errore durante l'eliminazione: {e}")
                else:
                    st.caption("Non ci sono altri utenti oltre al tuo account Admin.")

        # --- TAB 3: ANTEPRIMA E GESTIONE DATI ---
        with tab3:
            st.subheader("Verifica ed elimina dati pubblicati")
            admin_client = init_supabase_admin()
            tutti_dati = admin_client.table("dati_premi").select("mese_riferimento, nome_scheda, dati_json").order("id", desc=True).execute()
            
            if tutti_dati.data:
                mesi = list(dict.fromkeys([r["mese_riferimento"] for r in tutti_dati.data]))
                
                col_m1, col_m2 = st.columns([3, 1])
                with col_m1:
                    mese_sel = st.selectbox("Seleziona periodo di riferimento", mesi)
                with col_m2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Elimina intero mese", type="secondary", use_container_width=True):
                        try:
                            admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_sel).execute()
                            st.success(f"Tutti i dati di '{mese_sel}' sono stati eliminati.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Errore nell'eliminazione: {e}")

                schede_mese = [r for r in tutti_dati.data if r["mese_riferimento"] == mese_sel]
                scheda_sel = st.selectbox("Visualizza anteprima scheda", [r["nome_scheda"] for r in schede_mese])
                
                dati_scelti = next(r["dati_json"] for r in schede_mese if r["nome_scheda"] == scheda_sel)
                st.dataframe(pd.DataFrame(dati_scelti), use_container_width=True, hide_index=True)
            else:
                st.info("Nessun dato ancora caricato.")

    # ==========================================
    # SEZIONE FILIALE (SOLA LETTURA PROPRIA SCHEDA)
    # ==========================================
    else:
        mia_scheda = profilo["nome_scheda_excel"]
        st.title(f"📊 Riepilogo Premi - {mia_scheda}")
        
        dati_res = supabase.table("dati_premi").select("mese_riferimento, dati_json, data_caricamento").eq("nome_scheda", mia_scheda).order("id", desc=True).execute()
        
        if dati_res.data:
            mesi_disponibili = [r["mese_riferimento"] for r in dati_res.data]
            mese_scelto = st.selectbox("Seleziona il periodo di riferimento:", mesi_disponibili)
            
            record_selezionato = next(r for r in dati_res.data if r["mese_riferimento"] == mese_scelto)
            df_filiale = pd.DataFrame(record_selezionato["dati_json"])
            
            st.dataframe(df_filiale, use_container_width=True, hide_index=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_filiale.to_excel(writer, index=False, sheet_name=mia_scheda[:31])
            
            st.download_button(
                label="📥 Scarica questa scheda in Excel",
                data=buffer.getvalue(),
                file_name=f"Premi_{mia_scheda}_{mese_scelto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Non sono ancora presenti dati pubblicati per la tua filiale.")
