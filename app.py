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
                "GOLD (€)": fmt_euro(
