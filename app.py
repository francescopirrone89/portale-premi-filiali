import streamlit as st
import pandas as pd
import openpyxl
from supabase import create_client
import io

st.set_page_config(page_title="Portale Premi Filiali", page_icon="📊", layout="wide")

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

@st.cache_resource
def init_supabase():
    return create_client(st.secrets["SUPABASE_URL"].strip(), st.secrets["SUPABASE_KEY"].strip())

def init_supabase_admin():
    return create_client(st.secrets["SUPABASE_URL"].strip(), st.secrets["SUPABASE_SERVICE_KEY"].strip())

supabase = init_supabase()

if "user" not in st.session_state:
    st.session_state.user = None
if "profilo" not in st.session_state:
    st.session_state.profilo = None

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
    return "-" if num is None else f"{num:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

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
            is_tot = "totale" in nome_str.lower()
            d = 0 if is_tot else 2
            righe_tabella.append({
                "Operatore": nome_str,
                "T+KM WI (RA)": fmt_num(ws.cell(row=r, column=3).value, 0),
                "T+KM WI (€)": fmt_euro(ws.cell(row=r, column=4).value),
                "T+KM WI (50%)": fmt_euro(ws.cell(row=r, column=5).value),
                "GOLD (Q)": fmt_num(ws.cell(row=r, column=6).value, d),
                "GOLD (€)": fmt_euro(ws.cell(row=r, column=7).value),
                "CPP (Q)": fmt_num(ws.cell(row=r, column=8).value, d),
                "CPP (€)": fmt_euro(ws.cell(row=r, column=9).value),
                "RAP (Q)": fmt_num(ws.cell(row=r, column=10).value, d),
                "RAP (€)": fmt_euro(ws.cell(row=r, column=11).value),
                "PAI (Q)": fmt_num(ws.cell(row=r, column=12).value, d),
                "PAI (€)": fmt_euro(ws.cell(row=r, column=13).value),
                "UPSELL (Q)": fmt_num(ws.cell(row=r, column=14).value, d),
                "UPSELL (€)": fmt_euro(ws.cell(row=r, column=15).value),
                "SEC. GUIDA (Q)": fmt_num(ws.cell(row=r, column=16).value, d),
                "SEC. GUIDA (€)": fmt_euro(ws.cell(row=r, column=17).value),
                "Totale (€)": fmt_euro(ws.cell(row=r, column=18).value),
                "% Operatore": fmt_num(ws.cell(row=r, column=19).value, d) + "%",
                "Premio Lordo (€)": fmt_euro(ws.cell(row=r, column=20).value),
            })
            if is_tot:
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
    html = (
        '<style>'
        '.table-wrapper { overflow-x: auto; border-radius: 8px; border: 1px solid #3a3f4b; margin: 8px 0 16px 0; }'
        '.premi-table { width: 100%; border-collapse: collapse; font-family: "Segoe UI", sans-serif; font-size: 13px; white-space: nowrap; }'
        '.premi-table th { background-color: #c6e0b4; color: #111; font-weight: 700; text-align: center; padding: 7px 10px; border: 1px solid #7f9c6b; }'
        '.premi-table td { padding: 7px 10px; border: 1px solid #3a3f4b; text-align: right; color: #f0f2f6; background-color: #161920; }'
        '.premi-table td.op-cell { text-align: left; font-weight: 600; position: sticky; left: 0; background-color: #1e222b; z-index: 1; }'
        '.premi-table td.col-tot { background-color: #1d2633; font-weight: 600; }'
        '.premi-table tr.totale-row td { background-color: #d9ead3 !important; color: #111 !important; font-weight: 700; border-top: 2px solid #5b7a47; }'
        '.premi-table tr.totale-row td.premio-cell { background-color: #f4cccc !important; }'
        '</style>'
        '<div class="table-wrapper"><table class="premi-table"><thead>'
        '<tr>'
        '<th rowspan="2" style="position:sticky;left:0;z-index:2;">Operatore</th>'
        '<th colspan="3">T+KM WI</th><th colspan="2">GOLD</th><th colspan="2">CPP</th>'
        '<th colspan="2">RAP</th><th colspan="2">PAI</th><th colspan="2">UPSELL</th><th colspan="2">SEC. GUIDA</th>'
        '<th rowspan="2">Totale</th><th rowspan="2">%<br>Operatore</th><th rowspan="2">Premio di<br>risultato Lordo</th>'
        '</tr>'
        '<tr>'
        '<th>RA TOT</th><th>€</th><th>50%</th><th>QTÀ</th><th>€</th><th>QTÀ</th><th>€</th>'
        '<th>QTÀ</th><th>€</th><th>QTÀ</th><th>€</th><th>QTÀ</th><th>€</th><th>QTÀ</th><th>€</th>'
        '</tr></thead><tbody>'
    )
    for _, row in df.iterrows():
        is_tot = "totale" in str(row["Operatore"]).lower()
        html += f'<tr{" class=\"totale-row\"" if is_tot else ""}>'
        for col in COLONNE_FILIALE:
            val = row.get(col, "-")
            cls = ' class="op-cell"' if col == "Operatore" else (' class="col-tot"' if col == "Totale (€)" else (' class="premio-cell"' if col == "Premio Lordo (€)" else ""))
            html += f"<td{cls}>{val}</td>"
        html += "</tr>"
    return html + "</tbody></table></div>"

def render_scheda_pulita(dati_json, nome_scheda, mese_rif):
    if isinstance(dati_json, dict) and "righe" in dati_json:
        kpi = dati_json.get("kpi", {})
        if kpi and any(v != "-" for v in kpi.values()):
            c1, c2, c3 = st.columns(3)
            c1.metric("📈 Perc. Fatturato", kpi.get("perc_fatturato", "-"))
            delta_txt = kpi.get("vs_obj", "")
            c2.metric("⭐ Risultato NPS", kpi.get("risultato_nps", "-"), delta=delta_txt if delta_txt else None, delta_color="off")
            c3.metric("🏆 Premio Lordo Filiale", kpi.get("premio_lordo", "-"))
            st.divider()
        df = pd.DataFrame(dati_json["righe"])
        if not df.empty:
            df = df[[c for c in COLONNE_FILIALE if c in df.columns]]
            st.markdown(genera_tabella_html_unita(df), unsafe_allow_html=True)
            return df
    else:
        df = pd.DataFrame(dati_json)
        st.dataframe(df, use_container_width=True, hide_index=True)
        return df

def login(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        st.session_state.profilo = supabase.table("profili_filiali").select("*").eq("id", res.user.id).single().execute().data
        st.rerun()
    except Exception:
        st.error("Credenziali non valide o profilo non configurato.")

def logout():
    supabase.auth.sign_out()
    st.session_state.user = None
    st.session_state.profilo = None
    st.rerun()

if not st.session_state.user:
    _, col2, _ = st.columns([1, 1.5, 1])
    with col2:
        st.title("📊 Portale Riepilogo Premi")
        st.markdown("Accedi con le tue credenziali aziendali per visualizzare i dati della tua filiale.")
        with st.form("login_form"):
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Accedi", use_container_width=True):
                login(email, password)
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

    if ruolo == "admin":
        st.title("🛠️ Pannello Amministrazione")
        tab1, tab2, tab3 = st.tabs(["📤 Carica Excel Mensile", "👥 Gestione Utenti e Filiali", "👀 Gestione e Anteprima Dati"])

        with tab1:
            st.subheader("Carica il riepilogo mensile")
            mese_rif = st.text_input("Mese / Periodo di riferimento (es. Ottobre 2026)", placeholder="Es. Ottobre 2026")
            uploaded_file = st.file_uploader("Seleziona il file Excel (.xlsx)", type=["xlsx"], key="upload_premi")
            if uploaded_file and mese_rif and st.button("Elabora e Pubblica Schede", type="primary"):
                try:
                    file_bytes = uploaded_file.getvalue()
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                    schede_elaborate = []
                    admin_client = init_supabase_admin()
                    admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_rif.strip()).execute()
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        test_h = str(ws.cell(row=2, column=3).value or "") + str(ws.cell(row=3, column=3).value or "")
                        if "RA TOT" in test_h or "T+KM" in test_h:
                            dati_salvare = parse_foglio_filiale(ws)
                        else:
                            df_gen = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name).dropna(how="all").dropna(axis=1, how="all").fillna("")
                            df_gen.columns = [str(c) if not str(c).startswith("Unnamed") else "" for c in df_gen.columns]
                            dati_salvare = df_gen.astype(str).to_dict(orient="records")
                        admin_client.table("dati_premi").insert({"mese_riferimento": mese_rif.strip(), "nome_scheda": sheet_name.strip(), "dati_json": dati_salvare}).execute()
                        schede_elaborate.append(sheet_name.strip())
                    st.success(f"✅ Pubblicazione completata per **{mese_rif}**!")
                    st.write("Schede divise e salvate:", ", ".join(schede_elaborate))
                except Exception as e:
                    st.error(f"Errore durante l'elaborazione: {e}")

        with tab2:
            st.subheader("📥 Importazione massiva utenti da Excel")
            df_modello = pd.DataFrame([
                {"email": "filiale.milazzo@azienda.it", "password": "Password123!", "nome_scheda_excel": "Milazzo"},
                {"email": "filiale.roma@azienda.it", "password": "Password456!", "nome_scheda_excel": "Roma"}
            ])
            buf_modello = io.BytesIO()
            with pd.ExcelWriter(buf_modello, engine='openpyxl') as writer:
                df_modello.to_excel(writer, index=False, sheet_name="Utenti")
            st.download_button("📄 Scarica file Excel di esempio", data=buf_modello.getvalue(), file_name="modello_importazione_utenti.xlsx")

            file_utenti = st.file_uploader("Carica l'elenco utenti (.xlsx)", type=["xlsx"], key="upload_utenti")
            if file_utenti and st.button("Importa e Crea Tutti gli Utenti", type="primary"):
                try:
                    df_imp = pd.read_excel(file_utenti)
                    df_imp.columns = [str(c).strip().lower() for c in df_imp.columns]
                    if not {"email", "password", "nome_scheda_excel"}.issubset(set(df_imp.columns)):
                        st.error("Colonne richieste: email, password, nome_scheda_excel")
                    else:
                        df_imp = df_imp.dropna(subset=["email", "password", "nome_scheda_excel"])
                        admin_client = init_supabase_admin()
                        creati, errori = 0, []
                        for _, row in df_imp.iterrows():
                            em, pw, sc = str(row["email"]).strip(), str(row["password"]).strip(), str(row["nome_scheda_excel"]).strip()
                            if len(pw) < 6:
                                errori.append(f"{em}: password < 6 caratteri")
                                continue
                            try:
                                nu = admin_client.auth.admin.create_user({"email": em, "password": pw, "email_confirm": True})
                                admin_client.table("profili_filiali").insert({"id": nu.user.id, "email": em, "ruolo": "filiale", "nome_scheda_excel": sc}).execute()
                                creati += 1
                            except Exception as ex_u:
                                errori.append(f"{em}: {ex_u}")
                        if creati > 0:
                            st.success(f"✅ {creati} utenti creati!")
                        for err in errori:
                            st.warning(err)
                except Exception as e:
                    st.error(f"Errore lettura Excel: {e}")

            st.divider()
            with st.expander("➕ Crea singolo account manualmente", expanded=False):
                with st.form("nuovo_utente_form"):
                    c1, c2 = st.columns(2)
                    nuova_email = c1.text_input("Email utente")
                    nuova_pass = c1.text_input("Password (min. 6 caratteri)", type="password")
                    nome_scheda = c2.text_input("Nome esatto scheda Excel (es. Milazzo)")
                    nuovo_ruolo = c2.selectbox("Ruolo", ["filiale", "admin"])
                    if st.form_submit_button("Crea Utente"):
                        if nuova_email and len(nuova_pass) >= 6 and nome_scheda:
                            try:
                                admin_client = init_supabase_admin()
                                nu = admin_client.auth.admin.create_user({"email": nuova_email.strip(), "password": nuova_pass, "email_confirm": True})
                                admin_client.table("profili_filiali").insert({"id": nu.user.id, "email": nuova_email.strip(), "ruolo": nuovo_ruolo, "nome_scheda_excel": nome_scheda.strip()}).execute()
                                st.success(f"Utente {nuova_email} creato!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Errore: {e}")
                        else:
                            st.warning("Compila tutti i campi.")

            st.divider()
            st.subheader("Utenti registrati")
            admin_client = init_supabase_admin()
            utenti_res = admin_client.table("profili_filiali").select("id, email, ruolo, nome_scheda_excel").order("email").execute()
            if utenti_res.data:
                df_utenti = pd.DataFrame(utenti_res.data)[["email", "ruolo", "nome_scheda_excel"]]
                df_utenti.columns = ["Email", "Ruolo", "Scheda Excel Associata"]
                st.dataframe(df_utenti, use_container_width=True, hide_index=True)
                altri = [u for u in utenti_res.data if u["id"] != st.session_state.user.id]
                if altri:
                    opzioni = {f"{u['email']} ({u['nome_scheda_excel']})": u["id"] for u in altri}
                    cd1, cd2 = st.columns([3, 1])
                    u_del = cd1.selectbox("Rimuovi account", list(opzioni.keys()), label_visibility="collapsed")
                    if cd2.button("Elimina Utente", use_container_width=True):
                        admin_client.auth.admin.delete_user(opzioni[u_del])
                        st.rerun()

        with tab3:
            st.subheader("Verifica ed elimina dati pubblicati")
            admin_client = init_supabase_admin()
            tutti = admin_client.table("dati_premi").select("mese_riferimento, nome_scheda, dati_json").order("id", desc=True).execute()
            if tutti.data:
                mesi = list(dict.fromkeys([r["mese_riferimento"] for r in tutti.data]))
                cm1, cm2 = st.columns([3, 1])
                mese_sel = cm1.selectbox("Seleziona periodo", mesi)
                if cm2.button("🗑️ Elimina intero mese", use_container_width=True):
                    admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_sel).execute()
                    st.rerun()
                schede_m = [r for r in tutti.data if r["mese_riferimento"] == mese_sel]
                scheda_sel = st.selectbox("Visualizza scheda", [r["nome_scheda"] for r in schede_m])
                dati_scelti = next(r["dati_json"] for r in schede_m if r["nome_scheda"] == scheda_sel)
                render_scheda_pulita(dati_scelti, scheda_sel, mese_sel)
            else:
                st.info("Nessun dato ancora caricato.")
    else:
        mia_scheda = profilo["nome_scheda_excel"]
        st.title(f"📊 Riepilogo Premi - {mia_scheda}")
        dati_res = supabase.table("dati_premi").select("mese_riferimento, dati_json").eq("nome_scheda", mia_scheda).order("id", desc=True).execute()
        if dati_res.data:
            mesi_disp = [r["mese_riferimento"] for r in dati_res.data]
            mese_scelto = st.selectbox("Seleziona il periodo di riferimento:", mesi_disp)
            rec = next(r for r in dati_res.data if r["mese_riferimento"] == mese_scelto)
            df_esp = render_scheda_pulita(rec["dati_json"], mia_scheda, mese_scelto)
            if df_esp is not None and not df_esp.empty:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    df_esp.to_excel(writer, index=False, sheet_name=mia_scheda[:31])
                st.download_button("📥 Scarica questa scheda in Excel", data=buf.getvalue(), file_name=f"Premi_{mia_scheda}_{mese_scelto}.xlsx")
        else:
            st.info("Non sono ancora presenti dati pubblicati per la tua filiale.")
