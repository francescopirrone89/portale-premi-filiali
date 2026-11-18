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

# --- CARICAMENTO COMPLETO DA SUPABASE (SUPERA IL LIMITE DI 1000 RIGHE) ---
def carica_tutti_dati_admin(admin_client, order_desc=False):
    tutti_record = []
    batch_size = 1000
    offset = 0
    while True:
        query = admin_client.table("dati_premi").select("id, mese_riferimento, nome_scheda, dati_json")
        query = query.order("id", desc=order_desc)
        res = query.range(offset, offset + batch_size - 1).execute()
        dati = res.data or []
        tutti_record.extend(dati)
        if len(dati) < batch_size:
            break
        offset += batch_size
    return tutti_record

# --- FUNZIONI DI FORMATTAZIONE E CONVERSIONE ---
def to_float(val):
    if val is None or str(val).strip() in ("", "-"):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("€", "").replace("%", "").replace("\xa0", "").strip()
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

def fmt_perc(val, show_sign=False):
    num = to_float(val)
    if num is None:
        return "-"
    if abs(num) <= 2.0 and num != 0:
        num = num * 100
    sign = "+" if (show_sign and num > 0) else ""
    return f"{sign}{num:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

# --- PARSING FOGLIO EXCEL ---
def parse_foglio_filiale(ws):
    righe_tabella = []
    kpi = {"perc_fatturato": "-", "risultato_nps": "-", "premio_lordo": "-", "vs_obj": ""}

    col_perc, col_premio = 19, 20
    for c in range(15, ws.max_column + 1):
        h_txt = (str(ws.cell(row=2, column=c).value or "") + " " + str(ws.cell(row=3, column=c).value or "")).strip().lower()
        if "operatore" in h_txt or "%" in h_txt:
            col_perc = c
        elif "premio" in h_txt or "lordo" in h_txt:
            col_premio = c

    for r in range(4, ws.max_row + 1):
        nome = ws.cell(row=r, column=2).value
        if nome is not None and str(nome).strip() != "":
            nome_str = str(nome).strip()
            if nome_str.upper() == "TOTALI":
                break
            is_tot = "totale" in nome_str.lower()
            d = 0 if is_tot else 2

            v_tkm_50 = to_float(ws.cell(row=r, column=5).value) or 0.0
            v_gold   = to_float(ws.cell(row=r, column=7).value) or 0.0
            v_cpp    = to_float(ws.cell(row=r, column=9).value) or 0.0
            v_rap    = to_float(ws.cell(row=r, column=11).value) or 0.0
            v_pai    = to_float(ws.cell(row=r, column=13).value) or 0.0
            v_upsell = to_float(ws.cell(row=r, column=15).value) or 0.0
            v_sec    = to_float(ws.cell(row=r, column=17).value) or 0.0

            totale_calcolato = v_tkm_50 + v_gold + v_cpp + v_rap + v_pai + v_upsell + v_sec

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
                "Totale (€)": fmt_euro(totale_calcolato),
                "% Operatore": fmt_num(ws.cell(row=r, column=col_perc).value, d) + "%",
                "Premio Lordo (€)": fmt_euro(ws.cell(row=r, column=col_premio).value),
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
                
                parti_obj = []
                for col_scan in range(max(1, c), min(ws.max_column + 1, c + 6)):
                    cell_v = ws.cell(row=r + 2, column=col_scan).value
                    if cell_v is not None and str(cell_v).strip() != "":
                        if isinstance(cell_v, (int, float)):
                            parti_obj.append(fmt_perc(cell_v, show_sign=True))
                        else:
                            s_v = str(cell_v).strip()
                            num_test = to_float(s_v)
                            if num_test is not None and "obj" not in s_v.lower():
                                parti_obj.append(fmt_perc(num_test, show_sign=True))
                            else:
                                parti_obj.append(s_v)
                if parti_obj:
                    kpi["vs_obj"] = " ".join(parti_obj)
                break

    return {"righe": righe_tabella, "kpi": kpi}

# --- AGGREGAZIONE GRAN TOTALE ANNUALE ---
def calcola_gran_totale_filiale(lista_dati_json):
    somme_op = {}
    ordine_op = []
    col_numeriche = [c for c in COLONNE_FILIALE if c not in ("Operatore", "% Operatore")]

    for d_json in lista_dati_json:
        if not isinstance(d_json, dict) or "righe" not in d_json:
            continue
        for r in d_json["righe"]:
            op = str(r.get("Operatore", "")).strip()
            if not op or "totale" in op.lower():
                continue
            if op not in somme_op:
                somme_op[op] = {c: 0.0 for c in col_numeriche}
                ordine_op.append(op)
            for c in col_numeriche:
                v = to_float(r.get(c))
                if v is not None:
                    somme_op[op][c] += v

    if not ordine_op:
        return None

    for op in ordine_op:
        s = somme_op[op]
        s["Totale (€)"] = (
            s["T+KM WI (50%)"] + s["GOLD (€)"] + s["CPP (€)"] +
            s["RAP (€)"] + s["PAI (€)"] + s["UPSELL (€)"] + s["SEC. GUIDA (€)"]
        )

    grand_tot = {c: sum(somme_op[op][c] for op in ordine_op) for c in col_numeriche}
    tot_generale_euro = grand_tot["Totale (€)"]

    righe_finali = []
    for op in ordine_op:
        s = somme_op[op]
        perc_op = (s["Totale (€)"] / tot_generale_euro * 100) if tot_generale_euro > 0 else 0.0
        righe_finali.append({
            "Operatore": op,
            "T+KM WI (RA)": fmt_num(s["T+KM WI (RA)"], 0),
            "T+KM WI (€)": fmt_euro(s["T+KM WI (€)"]),
            "T+KM WI (50%)": fmt_euro(s["T+KM WI (50%)"]),
            "GOLD (Q)": fmt_num(s["GOLD (Q)"], 2),
            "GOLD (€)": fmt_euro(s["GOLD (€)"]),
            "CPP (Q)": fmt_num(s["CPP (Q)"], 2),
            "CPP (€)": fmt_euro(s["CPP (€)"]),
            "RAP (Q)": fmt_num(s["RAP (Q)"], 2),
            "RAP (€)": fmt_euro(s["RAP (€)"]),
            "PAI (Q)": fmt_num(s["PAI (Q)"], 2),
            "PAI (€)": fmt_euro(s["PAI (€)"]),
            "UPSELL (Q)": fmt_num(s["UPSELL (Q)"], 2),
            "UPSELL (€)": fmt_euro(s["UPSELL (€)"]),
            "SEC. GUIDA (Q)": fmt_num(s["SEC. GUIDA (Q)"], 2),
            "SEC. GUIDA (€)": fmt_euro(s["SEC. GUIDA (€)"]),
            "Totale (€)": fmt_euro(s["Totale (€)"]),
            "% Operatore": fmt_num(perc_op, 2) + "%",
            "Premio Lordo (€)": fmt_euro(s["Premio Lordo (€)"]),
        })

    righe_finali.append({
        "Operatore": "Totale complessivo",
        "T+KM WI (RA)": fmt_num(grand_tot["T+KM WI (RA)"], 0),
        "T+KM WI (€)": fmt_euro(grand_tot["T+KM WI (€)"]),
        "T+KM WI (50%)": fmt_euro(grand_tot["T+KM WI (50%)"]),
        "GOLD (Q)": fmt_num(grand_tot["GOLD (Q)"], 0),
        "GOLD (€)": fmt_euro(grand_tot["GOLD (€)"]),
        "CPP (Q)": fmt_num(grand_tot["CPP (Q)"], 0),
        "CPP (€)": fmt_euro(grand_tot["CPP (€)"]),
        "RAP (Q)": fmt_num(grand_tot["RAP (Q)"], 0),
        "RAP (€)": fmt_euro(grand_tot["RAP (€)"]),
        "PAI (Q)": fmt_num(grand_tot["PAI (Q)"], 0),
        "PAI (€)": fmt_euro(grand_tot["PAI (€)"]),
        "UPSELL (Q)": fmt_num(grand_tot["UPSELL (Q)"], 0),
        "UPSELL (€)": fmt_euro(grand_tot["UPSELL (€)"]),
        "SEC. GUIDA (Q)": fmt_num(grand_tot["SEC. GUIDA (Q)"], 0),
        "SEC. GUIDA (€)": fmt_euro(grand_tot["SEC. GUIDA (€)"]),
        "Totale (€)": fmt_euro(grand_tot["Totale (€)"]),
        "% Operatore": "100%",
        "Premio Lordo (€)": fmt_euro(grand_tot["Premio Lordo (€)"]),
    })

    return {"righe": righe_finali}

# --- CLASSIFICA OPERATORI (USA T+KM WI € PIENO + ALTRE VOCI €) ---
def calcola_classifica_operatori(records_filtrati):
    """Usa calcola_gran_totale_filiale per ogni filiale e somma T+KM WI (€) pieno + le altre voci €."""
    per_filiale = {}
    for rec in records_filtrati:
        fil = str(rec.get("nome_scheda", "")).strip()
        d_json = rec.get("dati_json", {})
        if isinstance(d_json, dict) and "righe" in d_json:
            per_filiale.setdefault(fil, []).append(d_json)

    agg_op = {}
    for fil, lista_json in per_filiale.items():
        gt = calcola_gran_totale_filiale(lista_json)
        if not gt or "righe" not in gt:
            continue
        for riga in gt["righe"]:
            op = str(riga.get("Operatore", "")).strip()
            if not op or "totale" in op.lower():
                continue
            val_tot = (
                (to_float(riga.get("T+KM WI (€)")) or 0.0) +
                (to_float(riga.get("GOLD (€)")) or 0.0) +
                (to_float(riga.get("CPP (€)")) or 0.0) +
                (to_float(riga.get("RAP (€)")) or 0.0) +
                (to_float(riga.get("PAI (€)")) or 0.0) +
                (to_float(riga.get("UPSELL (€)")) or 0.0) +
                (to_float(riga.get("SEC. GUIDA (€)")) or 0.0)
            )
            chiave = " ".join(op.upper().split())
            if chiave not in agg_op:
                agg_op[chiave] = {
                    "Operatore": op,
                    "Filiali": set(),
                    "Totale (€)": 0.0
                }
            if fil:
                agg_op[chiave]["Filiali"].add(fil)
            agg_op[chiave]["Totale (€)"] += val_tot

    risultati = [
        {
            "Operatore": v["Operatore"],
            "Filiale": ", ".join(sorted(v["Filiali"])),
            "Totale (€)": v["Totale (€)"]
        }
        for v in agg_op.values()
    ]
    return sorted(risultati, key=lambda x: x["Totale (€)"], reverse=True)

# --- RENDERING GRAFICO ---
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

def render_scheda_pulita(dati_json, mostra_kpi=True):
    if isinstance(dati_json, dict) and "righe" in dati_json:
        kpi = dati_json.get("kpi", {})
        if mostra_kpi and kpi and any(v != "-" for v in kpi.values()):
            c1, c2, c3 = st.columns(3)
            c1.metric("📈 Perc. Fatturato", kpi.get("perc_fatturato", "-"))
            
            delta_txt = str(kpi.get("vs_obj", "")).strip()
            nps_val = kpi.get("risultato_nps", "-")
            
            if "-" in delta_txt:
                badge_color = "#ff4b4b"
                bg_badge = "rgba(255, 75, 75, 0.15)"
                arrow = "▼ "
            elif "+" in delta_txt:
                badge_color = "#21c354"
                bg_badge = "rgba(33, 195, 84, 0.15)"
                arrow = "▲ "
            else:
                badge_color = "#a3a8b8"
                bg_badge = "rgba(163, 168, 184, 0.15)"
                arrow = ""
            
            badge_html = (
                f'<span style="display:inline-block; margin-top:4px; padding:2px 8px; border-radius:12px; '
                f'font-size:13px; font-weight:600; color:{badge_color}; background-color:{bg_badge};">'
                f'{arrow}{delta_txt}</span>'
            ) if delta_txt else ""

            c2.markdown(
                f'<div style="line-height:1.2;">'
                f'<div style="font-size:14px; color:#fafafa; margin-bottom:4px;">⭐ Risultato NPS</div>'
                f'<div style="font-size:32px; font-weight:600; color:#ffffff;">{nps_val}</div>'
                f'{badge_html}'
                f'</div>',
                unsafe_allow_html=True
            )
            
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

# --- AUTH ---
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

    # ==========================================
    # PANNELLO ADMIN
    # ==========================================
    if ruolo == "admin":
        st.title("🛠️ Pannello Amministrazione")
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📤 Carica Excel Mensile",
            "👥 Gestione Utenti e Filiali",
            "👀 Gestione e Anteprima Dati",
            "📊 Gran Totale Annuale",
            "🏆 Classifica Operatori"
        ])

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
            tutti = carica_tutti_dati_admin(admin_client, order_desc=True)
            if tutti:
                mesi = list(dict.fromkeys([r["mese_riferimento"] for r in tutti]))
                cm1, cm2 = st.columns([3, 1])
                mese_sel = cm1.selectbox("Seleziona periodo", mesi)
                if cm2.button("🗑️ Elimina intero mese", use_container_width=True):
                    admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_sel).execute()
                    st.rerun()
                schede_m = [r for r in tutti if r["mese_riferimento"] == mese_sel]
                scheda_sel = st.selectbox("Visualizza scheda", [r["nome_scheda"] for r in schede_m])
                dati_scelti = next(r["dati_json"] for r in schede_m if r["nome_scheda"] == scheda_sel)
                render_scheda_pulita(dati_scelti, mostra_kpi=True)
            else:
                st.info("Nessun dato ancora caricato.")

        with tab4:
            st.subheader("📊 Gran Totale Annuale per Ufficio")
            admin_client = init_supabase_admin()
            tutti_ann = carica_tutti_dati_admin(admin_client)
            
            if tutti_ann:
                mesi_disponibili = list(dict.fromkeys([r["mese_riferimento"] for r in tutti_ann]))
                filiali_uniche = sorted(list(set(
                    r["nome_scheda"] for r in tutti_ann
                    if isinstance(r["dati_json"], dict) and "righe" in r["dati_json"]
                )))
                
                col_f1, col_f2 = st.columns([1, 2])
                with col_f1:
                    fil_sel = st.selectbox("Seleziona l'Ufficio / Filiale:", filiali_uniche, key="ann_fil_sel")
                with col_f2:
                    mesi_scelti_ann = st.multiselect("Mesi da sommare:", mesi_disponibili, default=mesi_disponibili, key="ann_mesi_sel")
                
                if fil_sel and mesi_scelti_ann:
                    lista_json_fil = [
                        r["dati_json"] for r in tutti_ann
                        if r["nome_scheda"] == fil_sel and r["mese_riferimento"] in mesi_scelti_ann
                    ]
                    gran_tot_res = calcola_gran_totale_filiale(lista_json_fil)
                    if gran_tot_res:
                        df_ann = render_scheda_pulita(gran_tot_res, mostra_kpi=False)
                        if df_ann is not None and not df_ann.empty:
                            buf_a = io.BytesIO()
                            with pd.ExcelWriter(buf_a, engine='openpyxl') as writer:
                                df_ann.to_excel(writer, index=False, sheet_name=f"Totale_{fil_sel}"[:31])
                            st.download_button(
                                f"📥 Scarica Gran Totale {fil_sel} (Excel)",
                                data=buf_a.getvalue(),
                                file_name=f"GranTotale_{fil_sel}.xlsx"
                            )
                    else:
                        st.warning("Nessun dato valido trovato per i mesi selezionati.")
            else:
                st.info("Carica almeno un file mensile per visualizzare il Gran Totale.")

        with tab5:
            st.subheader("🏆 Classifica Operatori (Somma T+KM WI € + Voci €)")
            admin_client = init_supabase_admin()
            tutti_op = carica_tutti_dati_admin(admin_client)

            if tutti_op:
                mesi_op = list(dict.fromkeys([r["mese_riferimento"] for r in tutti_op]))
                filiali_op = sorted(list(set(
                    r["nome_scheda"] for r in tutti_op
                    if isinstance(r["dati_json"], dict) and "righe" in r["dati_json"]
                )))

                c_op1, c_op2 = st.columns([2, 1])
                with c_op1:
                    mesi_sel_op = st.multiselect("Mesi da sommare:", mesi_op, default=mesi_op, key="op_mesi_sel")
                with c_op2:
                    filtro_fil_op = st.selectbox("Filtra per Ufficio:", ["Tutte le Filiali"] + filiali_op, key="op_fil_sel")

                if mesi_sel_op:
                    rec_filtrati = [
                        r for r in tutti_op
                        if r["mese_riferimento"] in mesi_sel_op
                        and (filtro_fil_op == "Tutte le Filiali" or r["nome_scheda"] == filtro_fil_op)
                    ]
                    classifica = calcola_classifica_operatori(rec_filtrati)

                    if classifica:
                        righe_classifica = [
                            {
                                "Pos.": f"{pos}°",
                                "Operatore": item["Operatore"],
                                "Filiale": item["Filiale"],
                                "Totale (€)": fmt_euro(item["Totale (€)"])
                            }
                            for pos, item in enumerate(classifica, start=1)
                        ]

                        df_classifica = pd.DataFrame(righe_classifica)
                        st.dataframe(df_classifica, use_container_width=True, hide_index=True)

                        buf_op = io.BytesIO()
                        with pd.ExcelWriter(buf_op, engine='openpyxl') as writer:
                            df_classifica.to_excel(writer, index=False, sheet_name="Classifica_Operatori")
                        st.download_button(
                            "📥 Scarica Classifica Operatori (Excel)",
                            data=buf_op.getvalue(),
                            file_name="Classifica_Operatori_Totale.xlsx"
                        )
                    else:
                        st.warning("Nessun operatore trovato per i filtri selezionati.")
            else:
                st.info("Carica almeno un file mensile per visualizzare la classifica.")

    # ==========================================
    # PANNELLO FILIALE
    # ==========================================
    else:
        mia_scheda = profilo["nome_scheda_excel"]
        st.title(f"📊 Riepilogo Premi - {mia_scheda}")
        dati_res = supabase.table("dati_premi").select("mese_riferimento, dati_json").eq("nome_scheda", mia_scheda).order("id", desc=True).execute()
        
        if dati_res.data:
            vista_filiale = st.radio("Modalità di consultazione:", ["📅 Dettaglio Mensile", "📊 Gran Totale Annuale"], horizontal=True)
            
            if vista_filiale == "📅 Dettaglio Mensile":
                mesi_disp = [r["mese_riferimento"] for r in dati_res.data]
                mese_scelto = st.selectbox("Seleziona il periodo di riferimento:", mesi_disp)
                rec = next(r for r in dati_res.data if r["mese_riferimento"] == mese_scelto)
                df_esp = render_scheda_pulita(rec["dati_json"], mostra_kpi=True)
                if df_esp is not None and not df_esp.empty:
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                        df_esp.to_excel(writer, index=False, sheet_name=mia_scheda[:31])
                    st.download_button("📥 Scarica questa scheda in Excel", data=buf.getvalue(), file_name=f"Premi_{mia_scheda}_{mese_scelto}.xlsx")
            else:
                mesi_disp = [r["mese_riferimento"] for r in dati_res.data]
                mesi_sel_f = st.multiselect("Seleziona i mesi da sommare:", mesi_disp, default=mesi_disp)
                if mesi_sel_f:
                    lista_j_f = [r["dati_json"] for r in dati_res.data if r["mese_riferimento"] in mesi_sel_f]
                    gt_f = calcola_gran_totale_filiale(lista_j_f)
                    if gt_f:
                        df_esp_ann = render_scheda_pulita(gt_f, mostra_kpi=False)
                        if df_esp_ann is not None and not df_esp_ann.empty:
                            buf_fa = io.BytesIO()
                            with pd.ExcelWriter(buf_fa, engine='openpyxl') as writer:
                                df_esp_ann.to_excel(writer, index=False, sheet_name=f"Totale_{mia_scheda}"[:31])
                            st.download_button("📥 Scarica Gran Totale in Excel", data=buf_fa.getvalue(), file_name=f"GranTotale_{mia_scheda}.xlsx")
        else:
            st.info("Non sono ancora presenti dati pubblicati per la tua filiale.")
