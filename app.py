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
    return "" if num is None else f"{num:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_num(val, dec=2):
    num = to_float(val)
    if num is None:
        return ""
    if dec == 0:
        return f"{int(round(num)):,}".replace(",", ".")
    return f"{num:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_perc(val, show_sign=False):
    num = to_float(val)
    if num is None:
        return ""
    if abs(num) <= 2.0 and num != 0:
        num = num * 100
    sign = "+" if (show_sign and num > 0) else ""
    return f"{sign}{num:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def classifica_tipo_colonna_totali(intestazione_completa, sample_fmt=""):
    h = intestazione_completa.strip().lower()
    if (
        "€" in sample_fmt
        or "€" in h
        or "50%" in h
        or any(k in h for k in ("premio", "boost", "qualit", "fatturato"))
    ):
        return "euro"
    if (
        "%" in sample_fmt
        or "%" in h
        or any(k in h for k in ("nps", "aliquota", "incidenza", "abbattim", "obj"))
    ):
        return "perc"
    if any(k in h for k in ("cod. staz", "ra tot", "t+km", "gold", "cpp", "rap", "pai", "upsell", "sec. guida", "(q)", "qtà")):
        return "int"
    return "auto"

# --- PARSER FEDELE AL FOGLIO EXCEL GENERICO / TOTALI (MANTIENE CELLE UNITE E ORDINE COLONNE) ---
def parse_foglio_generico_excel(ws, max_col_limit=None):
    max_r = ws.max_row
    max_c = min(ws.max_column, max_col_limit) if max_col_limit else ws.max_column

    col_valide = []
    for c in range(1, max_c + 1):
        ha_valori = any(ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != "" for r in range(1, max_r + 1))
        if ha_valori:
            col_valide.append(c)
    if not col_valide:
        return []
    c_start, c_end = min(col_valide), max(col_valide)

    r_head_1 = 1
    for r in range(1, max_r + 1):
        if any(ws.cell(row=r, column=c).value is not None and str(ws.cell(row=r, column=c).value).strip() != "" for c in range(c_start, c_end + 1)):
            r_head_1 = r
            break

    has_sub_header = False
    if r_head_1 + 1 <= max_r:
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r_head_1 <= rng.max_row and rng.max_row > r_head_1:
                has_sub_header = True
                break
        if not has_sub_header:
            sub_vals = [str(ws.cell(row=r_head_1 + 1, column=c).value or "").strip() for c in range(c_start, c_end + 1)]
            if any(x in ("Q", "QTÀ", "RA TOT", "€", "50%", "%") for x in sub_vals):
                has_sub_header = True

    r_head_end = r_head_1 + 1 if has_sub_header else r_head_1

    merged_map = {}
    skip_cells = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row >= r_head_1 and rng.min_row <= r_head_end and rng.min_col <= c_end:
            rs = min(rng.max_row, r_head_end) - rng.min_row + 1
            cs = min(rng.max_col, c_end) - max(rng.min_col, c_start) + 1
            c0 = max(rng.min_col, c_start)
            if cs >= 1 and rs >= 1:
                merged_map[(rng.min_row, c0)] = (rs, cs)
                for rr in range(rng.min_row, min(rng.max_row, r_head_end) + 1):
                    for cc in range(c0, min(rng.max_col, c_end) + 1):
                        if (rr, cc) != (rng.min_row, c0):
                            skip_cells.add((rr, cc))

    if has_sub_header:
        for c in range(c_start, c_end + 1):
            if (r_head_1, c) not in merged_map and (r_head_1, c) not in skip_cells:
                v2 = ws.cell(row=r_head_end, column=c).value
                if (v2 is None or str(v2).strip() == "") and (r_head_end, c) not in skip_cells:
                    merged_map[(r_head_1, c)] = (2, 1)
                    skip_cells.add((r_head_end, c))

    header_rows = []
    for r in range(r_head_1, r_head_end + 1):
        row_cells = []
        for c in range(c_start, c_end + 1):
            if (r, c) in skip_cells:
                continue
            val = str(ws.cell(row=r, column=c).value or "").strip()
            rs, cs = merged_map.get((r, c), (1, 1))
            row_cells.append({"text": val, "rowspan": rs, "colspan": cs, "col": c})
        header_rows.append(row_cells)

    col_types = {}
    for c in range(c_start, c_end + 1):
        h1 = ""
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r_head_1 <= rng.max_row and rng.min_col <= c <= rng.max_col:
                h1 = str(ws.cell(row=rng.min_row, column=rng.min_col).value or "").strip()
                break
        if not h1:
            h1 = str(ws.cell(row=r_head_1, column=c).value or "").strip()
        h2 = str(ws.cell(row=r_head_end, column=c).value or "").strip() if has_sub_header else ""
        full_h = f"{h1} {h2}".strip()

        sample_fmt = ""
        for rr in range(r_head_end + 1, min(max_r + 1, r_head_end + 10)):
            cell_s = ws.cell(row=rr, column=c)
            if isinstance(cell_s.value, (int, float)) and cell_s.number_format:
                sample_fmt = str(cell_s.number_format)
                break

        col_types[c] = classifica_tipo_colonna_totali(full_h, sample_fmt)

    body_rows = []
    for r in range(r_head_end + 1, max_r + 1):
        raw_vals = [ws.cell(row=r, column=c).value for c in range(c_start, c_end + 1)]
        if all(v is None or str(v).strip() in ("", "-") for v in raw_vals):
            continue

        row_formatted = []
        is_tot_row = False
        for c in range(c_start, c_end + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None or str(val).strip() in ("", "-"):
                row_formatted.append("")
                continue

            if isinstance(val, str):
                s_val = val.strip()
                if "totale" in s_val.lower() or "complessiv" in s_val.lower():
                    is_tot_row = True
                row_formatted.append(s_val)
            elif isinstance(val, (int, float)):
                ctype = col_types.get(c, "auto")
                if ctype == "perc":
                    row_formatted.append(fmt_perc(val))
                elif ctype == "euro":
                    row_formatted.append(fmt_euro(val))
                elif ctype == "int":
                    row_formatted.append(fmt_num(val, 0))
                else:
                    row_formatted.append(fmt_num(val, 0) if abs(val - round(val)) < 1e-6 else fmt_num(val, 2))
            else:
                row_formatted.append(str(val).strip())

        body_rows.append({"cells": row_formatted, "is_tot": is_tot_row})
        if is_tot_row:
            break

    return {"tipo": "tabella_excel_unita", "header_rows": header_rows, "body_rows": body_rows}

def render_tabella_excel_unita_html(dati_tab):
    header_rows = dati_tab.get("header_rows", [])
    body_rows = dati_tab.get("body_rows", [])

    col_headers = {}
    for hrow in header_rows:
        col_ptr = 0
        for cell in hrow:
            c_idx = cell.get("col", col_ptr)
            cs = cell.get("colspan", 1)
            txt = str(cell.get("text", "")).strip()
            for offset in range(cs):
                col_headers[c_idx + offset] = (col_headers.get(c_idx + offset, "") + " " + txt).strip()
            col_ptr = c_idx + cs
    ordered_col_keys = sorted(col_headers.keys())

    html = (
        '<style>'
        '.ex-wrapper { overflow-x: auto; -webkit-overflow-scrolling: touch; border-radius: 8px; border: 1px solid light-dark(#c5cad3, #3a3f4b); margin: 8px 0 16px 0; }'
        '.ex-table { width: 100%; border-collapse: collapse; font-family: "Segoe UI", sans-serif; font-size: 13px; white-space: nowrap; }'
        '.ex-table th { background-color: #c6e0b4; color: #111111; font-weight: 700; text-align: center; padding: 7px 10px; border: 1px solid #7f9c6b; }'
        '.ex-table td { padding: 7px 10px; border: 1px solid light-dark(#d0d7de, #3a3f4b); text-align: right; color: light-dark(#111111, #f0f2f6); background-color: light-dark(#ffffff, #0e1117); }'
        '.ex-table td.first-col { text-align: left; font-weight: 600; position: sticky; left: 0; background-color: light-dark(#f6f8fa, #1e222b); color: light-dark(#111111, #f0f2f6); z-index: 1; }'
        '.ex-table td.text-col { text-align: center; }'
        '.ex-table tr.totale-row td { background-color: #d9ead3 !important; color: #111111 !important; font-weight: 700; border-top: 2px solid #5b7a47; }'
        '@media (max-width: 768px) {'
        '  .ex-table { font-size: 11px; }'
        '  .ex-table th, .ex-table td { padding: 5px 6px; }'
        '}'
        '</style>'
        '<div class="ex-wrapper"><table class="ex-table"><thead>'
    )

    for r_idx, hrow in enumerate(header_rows):
        html += "<tr>"
        for c_idx, cell in enumerate(hrow):
            rs = f' rowspan="{cell["rowspan"]}"' if cell["rowspan"] > 1 else ""
            cs = f' colspan="{cell["colspan"]}"' if cell["colspan"] > 1 else ""
            sticky = ' style="position:sticky;left:0;z-index:2;"' if (r_idx == 0 and c_idx == 0) else ""
            html += f"<th{rs}{cs}{sticky}>{cell['text']}</th>"
        html += "</tr>"

    html += "</thead><tbody>"
    for brow in body_rows:
        celle_pulite = ["" if str(v).strip() == "-" else str(v).strip() for v in brow.get("cells", [])]
        if all(c == "" for c in celle_pulite):
            continue

        prima_cella = celle_pulite[0].lower() if celle_pulite else ""
        is_tot = brow.get("is_tot") or ("totale" in prima_cella or "complessiv" in prima_cella)
        tr_cls = ' class="totale-row"' if is_tot else ""
        html += f"<tr{tr_cls}>"

        for idx, s_val in enumerate(celle_pulite):
            if idx < len(ordered_col_keys) and s_val != "":
                h_txt = col_headers.get(ordered_col_keys[idx], "")
                if classifica_tipo_colonna_totali(h_txt) == "euro" and "€" not in s_val:
                    num_test = to_float(s_val)
                    if num_test is not None:
                        s_val = fmt_euro(num_test)

            if idx == 0:
                td_cls = ' class="first-col"'
            elif s_val != "" and to_float(s_val) is None:
                td_cls = ' class="text-col"'
            else:
                td_cls = ""
            html += f"<td{td_cls}>{s_val}</td>"
        html += "</tr>"

        if is_tot:
            break

    html += "</tbody></table></div>"
    return html

# --- PARSING FOGLIO EXCEL SINGOLA FILIALE ---
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
            perc_op_fmt = fmt_num(ws.cell(row=r, column=col_perc).value, d)

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
                "% Operatore": f"{perc_op_fmt}%" if perc_op_fmt != "" else "",
                "Premio Lordo (€)": fmt_euro(ws.cell(row=r, column=col_premio).value),
            })
            if is_tot:
                break

    for r in range(25, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            if "Perc. Fatturato" in val:
                kpi["perc_fatturato"] = fmt_perc(ws.cell(row=r + 1, column=c).value) or "-"
                kpi["risultato_nps"] = fmt_perc(ws.cell(row=r + 1, column=c + 2).value) or "-"
                kpi["premio_lordo"] = fmt_euro(ws.cell(row=r + 1, column=c + 4).value) or "-"
                
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

# --- RENDERING GRAFICO (ADATTIVO NATIVO CON light-dark() + RESPONSIVE MOBILE) ---
def genera_tabella_html_unita(df):
    html = (
        '<style>'
        '.table-wrapper { overflow-x: auto; -webkit-overflow-scrolling: touch; border-radius: 8px; border: 1px solid light-dark(#c5cad3, #3a3f4b); margin: 8px 0 16px 0; }'
        '.premi-table { width: 100%; border-collapse: collapse; font-family: "Segoe UI", sans-serif; font-size: 13px; white-space: nowrap; }'
        '.premi-table th { background-color: #c6e0b4; color: #111111; font-weight: 700; text-align: center; padding: 7px 10px; border: 1px solid #7f9c6b; }'
        '.premi-table td { padding: 7px 10px; border: 1px solid light-dark(#d0d7de, #3a3f4b); text-align: right; color: light-dark(#111111, #f0f2f6); background-color: light-dark(#ffffff, #0e1117); }'
        '.premi-table td.op-cell { text-align: left; font-weight: 600; position: sticky; left: 0; background-color: light-dark(#f6f8fa, #1e222b); color: light-dark(#111111, #f0f2f6); z-index: 1; }'
        '.premi-table td.col-tot { background-color: light-dark(#eef5e9, #1d2633); font-weight: 600; }'
        '.premi-table tr.totale-row td { background-color: #d9ead3 !important; color: #111111 !important; font-weight: 700; border-top: 2px solid #5b7a47; }'
        '.premi-table tr.totale-row td.premio-cell { background-color: #f4cccc !important; color: #111111 !important; }'
        '@media (max-width: 768px) {'
        '  .premi-table { font-size: 11px; }'
        '  .premi-table th, .premi-table td { padding: 5px 6px; }'
        '  .premi-table td.op-cell { max-width: 115px; overflow: hidden; text-overflow: ellipsis; }'
        '}'
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
            raw_val = row.get(col, "")
            val = "" if str(raw_val).strip() in ("-", "nan", "None") else str(raw_val).strip()
            cls = ' class="op-cell"' if col == "Operatore" else (' class="col-tot"' if col == "Totale (€)" else (' class="premio-cell"' if col == "Premio Lordo (€)" else ""))
            html += f"<td{cls}>{val}</td>"
        html += "</tr>"
    return html + "</tbody></table></div>"

def render_scheda_pulita(dati_json, mostra_kpi=True):
    if isinstance(dati_json, dict) and "righe" in dati_json:
        kpi = dati_json.get("kpi", {})
        if mostra_kpi and kpi and any(v not in ("-", "") for v in kpi.values()):
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
                badge_color = "#808495"
                bg_badge = "rgba(128, 132, 149, 0.15)"
                arrow = ""
            
            badge_html = (
                f'<span style="display:inline-block; margin-top:4px; padding:2px 8px; border-radius:12px; '
                f'font-size:13px; font-weight:600; color:{badge_color}; background-color:{bg_badge};">'
                f'{arrow}{delta_txt}</span>'
            ) if delta_txt else ""

            with c2:
                st.metric("⭐ Risultato NPS", nps_val)
                if badge_html:
                    st.markdown(f'<div style="margin-top:-12px;">{badge_html}</div>', unsafe_allow_html=True)
            
            c3.metric("🏆 Premio Lordo Filiale", kpi.get("premio_lordo", "-"))
            st.divider()

        df = pd.DataFrame(dati_json["righe"])
        if not df.empty:
            df = df[[c for c in COLONNE_FILIALE if c in df.columns]]
            st.markdown(genera_tabella_html_unita(df), unsafe_allow_html=True)
            return df
    elif isinstance(dati_json, dict) and dati_json.get("tipo") == "tabella_excel_unita":
        st.markdown(render_tabella_excel_unita_html(dati_json), unsafe_allow_html=True)
        return None
    else:
        df = pd.DataFrame(dati_json)
        st.dataframe(df, use_container_width=True, hide_index=True)
        return df

# --- TABELLA RANGE ALIQUOTE (ADATTIVA AL TEMA LIGHT / DARK DI STREAMLIT) ---
def render_tabella_range_aliquote():
    html = (
        '<style>'
        '.range-box { display: flex; justify-content: center; margin-bottom: 6px; }'
        '.range-table { border-collapse: collapse; font-family: "Segoe UI", Calibri, sans-serif; font-size: 13px; border: 1.5px solid light-dark(#000000, #8c959f); background-color: light-dark(#ffffff, #0e1117); }'
        '.range-table th { background-color: #e2efda; color: #002060; font-weight: 700; text-align: center; padding: 6px 14px; border-bottom: 1.5px solid light-dark(#000000, #8c959f); letter-spacing: 0.3px; }'
        '.range-table td { color: light-dark(#002060, #f0f2f6); padding: 2px 10px; text-align: center; background-color: light-dark(#ffffff, #0e1117); line-height: 1.35; }'
        '.range-table td.col-da { text-align: right; padding-left: 14px; padding-right: 6px; }'
        '.range-table td.col-a { text-align: right; padding-left: 6px; padding-right: 10px; border-right: 1.5px dotted light-dark(#000000, #8c959f); }'
        '.range-table td.col-aliq { text-align: center; padding-left: 12px; padding-right: 14px; }'
        '.range-table tr:first-child td { padding-top: 5px; }'
        '.range-table tr:last-child td { padding-bottom: 5px; }'
        '</style>'
        '<div class="range-box"><table class="range-table">'
        '<thead><tr><th colspan="3">RANGE ALIQUOTE</th></tr></thead>'
        '<tbody>'
        '<tr><td class="col-da">0,00%</td><td class="col-a">7,99%</td><td class="col-aliq">NO PREMIO</td></tr>'
        '<tr><td class="col-da">8,00%</td><td class="col-a">14,99%</td><td class="col-aliq">6%</td></tr>'
        '<tr><td class="col-da">15,00%</td><td class="col-a">19,99%</td><td class="col-aliq">8%</td></tr>'
        '<tr><td class="col-da">20,00%</td><td class="col-a">24,99%</td><td class="col-aliq">10%</td></tr>'
        '<tr><td class="col-da">25,00%</td><td class="col-a">29,99%</td><td class="col-aliq">12%</td></tr>'
        '<tr><td class="col-da">30,00%</td><td class="col-a">100,00%</td><td class="col-aliq">14%</td></tr>'
        '</tbody></table></div>'
    )
    st.markdown(html, unsafe_allow_html=True)

# --- HELPER SEZIONI CONDIVISE TRA SUPER ADMIN E ADMIN ---
def render_sezione_anteprima(is_super_admin):
    col_tit_ante, col_range_ante, col_spazio_ante = st.columns([1.4, 1.3, 1.3])
    with col_tit_ante:
        st.subheader("Verifica ed elimina dati pubblicati" if is_super_admin else "Verifica dati pubblicati")
    with col_range_ante:
        render_tabella_range_aliquote()

    admin_client = init_supabase_admin()
    tutti = carica_tutti_dati_admin(admin_client, order_desc=True)
    if tutti:
        mesi = list(dict.fromkeys([r["mese_riferimento"] for r in tutti]))
        if is_super_admin:
            cm1, cm2 = st.columns([3, 1])
            mese_sel = cm1.selectbox("Seleziona periodo", mesi)
            if cm2.button("🗑️ Elimina intero mese", use_container_width=True):
                admin_client.table("dati_premi").delete().eq("mese_riferimento", mese_sel).execute()
                st.rerun()
        else:
            mese_sel = st.selectbox("Seleziona periodo", mesi)

        schede_m = [r for r in tutti if r["mese_riferimento"] == mese_sel]
        scheda_sel = st.selectbox("Visualizza scheda", sorted([r["nome_scheda"] for r in schede_m], key=lambda x: x.lower()))
        dati_scelti = next(r["dati_json"] for r in schede_m if r["nome_scheda"] == scheda_sel)
        render_scheda_pulita(dati_scelti, mostra_kpi=True)
    else:
        st.info("Nessun dato ancora caricato.")

def render_sezione_gran_totale():
    st.subheader("📊 Gran Totale Annuale per Ufficio")
    admin_client = init_supabase_admin()
    tutti_ann = carica_tutti_dati_admin(admin_client)
    
    if tutti_ann:
        mesi_disponibili = list(dict.fromkeys([r["mese_riferimento"] for r in tutti_ann]))
        filiali_uniche = sorted(list(set(
            r["nome_scheda"] for r in tutti_ann
            if isinstance(r["dati_json"], dict) and "righe" in r["dati_json"]
        )), key=lambda x: x.lower())
        
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

def render_sezione_classifica():
    st.subheader("🏆 Classifica Operatori")
    admin_client = init_supabase_admin()
    tutti_op = carica_tutti_dati_admin(admin_client)

    if tutti_op:
        mesi_op = list(dict.fromkeys([r["mese_riferimento"] for r in tutti_op]))
        filiali_op = sorted(list(set(
            r["nome_scheda"] for r in tutti_op
            if isinstance(r["dati_json"], dict) and "righe" in r["dati_json"]
        )), key=lambda x: x.lower())

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
    ruolo = str(profilo.get("ruolo", "filiale")).strip().lower()
    with st.sidebar:
        st.write(f"👤 **{profilo['email']}**")
        st.write(f"Ruolo: `{ruolo.replace('_', ' ').upper()}`")
        if ruolo == "filiale":
            st.write(f"Filiale: **{profilo['nome_scheda_excel']}**")
        st.divider()
        if st.button("Esci (Logout)", use_container_width=True):
            logout()

    # ==========================================
    # PANNELLO SUPER ADMIN E ADMIN
    # ==========================================
    if ruolo in ("super_admin", "admin"):
        is_super = (ruolo == "super_admin")

        st.markdown(
            """
            <style>
            @media (min-width: 769px) {
                .admin-mobile-bar { display: none !important; }
            }
            @media (max-width: 768px) {
                .block-container {
                    padding-top: 2.6rem !important;
                    padding-left: 0.7rem !important;
                    padding-right: 0.7rem !important;
                    padding-bottom: 1.5rem !important;
                }
                h1 { font-size: 1.45rem !important; margin-bottom: 0.4rem !important; }
                h2, h3 { font-size: 1.15rem !important; }
                .stTabs [data-baseweb="tab-list"] {
                    gap: 4px !important;
                    overflow-x: auto !important;
                    flex-wrap: nowrap !important;
                    -webkit-overflow-scrolling: touch;
                    padding-bottom: 4px;
                }
                .stTabs [data-baseweb="tab"] {
                    font-size: 12px !important;
                    padding: 6px 10px !important;
                    white-space: nowrap !important;
                }
                [data-testid="stMetricValue"] { font-size: 1.35rem !important; }
                [data-testid="stMetricLabel"] { font-size: 0.8rem !important; }
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        col_head, col_mob_out = st.columns([4, 1])
        with col_head:
            st.title("🛠️ Pannello Super Amministrazione" if is_super else "🛠️ Pannello Amministrazione")
        with col_mob_out:
            st.markdown('<div class="admin-mobile-bar">', unsafe_allow_html=True)
            if st.button("🚪 Esci", key="mobile_logout_admin", use_container_width=True):
                logout()
            st.markdown('</div>', unsafe_allow_html=True)

        if is_super:
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
                            nome_pulito = sheet_name.strip()
                            test_h = str(ws.cell(row=2, column=3).value or "") + str(ws.cell(row=3, column=3).value or "")
                            if nome_pulito.upper() == "TOTALI":
                                dati_salvare = parse_foglio_generico_excel(ws, max_col_limit=34)
                            elif "RA TOT" in test_h or "T+KM" in test_h:
                                dati_salvare = parse_foglio_filiale(ws)
                            else:
                                dati_salvare = parse_foglio_generico_excel(ws)
                            admin_client.table("dati_premi").insert({"mese_riferimento": mese_rif.strip(), "nome_scheda": nome_pulito, "dati_json": dati_salvare}).execute()
                            schede_elaborate.append(nome_pulito)
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
                if file_utenti and st.button("Importa e Aggiorna Tutti gli Utenti", type="primary"):
                    try:
                        df_imp = pd.read_excel(file_utenti)
                        df_imp.columns = [str(c).strip().lower() for c in df_imp.columns]
                        if not {"email", "password", "nome_scheda_excel"}.issubset(set(df_imp.columns)):
                            st.error("Colonne richieste: email, password, nome_scheda_excel")
                        else:
                            df_imp = df_imp.dropna(subset=["email", "password", "nome_scheda_excel"])
                            admin_client = init_supabase_admin()
                            esistenti_res = admin_client.table("profili_filiali").select("id, email").execute()
                            mappa_esistenti = {u["email"].strip().lower(): u["id"] for u in (esistenti_res.data or [])}

                            creati, aggiornati, errori = 0, 0, []
                            for _, row in df_imp.iterrows():
                                em, pw, sc = str(row["email"]).strip(), str(row["password"]).strip(), str(row["nome_scheda_excel"]).strip()
                                if len(pw) < 6:
                                    errori.append(f"{em}: password < 6 caratteri")
                                    continue
                                try:
                                    if em.lower() in mappa_esistenti:
                                        uid = mappa_esistenti[em.lower()]
                                        admin_client.auth.admin.update_user_by_id(uid, {"password": pw})
                                        admin_client.table("profili_filiali").update({
                                            "nome_scheda_excel": sc,
                                            "password_visibile": pw,
                                            "must_change_password": True
                                        }).eq("id", uid).execute()
                                        aggiornati += 1
                                    else:
                                        nu = admin_client.auth.admin.create_user({"email": em, "password": pw, "email_confirm": True})
                                        admin_client.table("profili_filiali").insert({
                                            "id": nu.user.id,
                                            "email": em,
                                            "ruolo": "filiale",
                                            "nome_scheda_excel": sc,
                                            "password_visibile": pw,
                                            "must_change_password": True
                                        }).execute()
                                        creati += 1
                                except Exception as ex_u:
                                    errori.append(f"{em}: {ex_u}")
                            if creati > 0 or aggiornati > 0:
                                st.success(f"✅ {creati} utenti creati, {aggiornati} utenti aggiornati!")
                                st.rerun()
                            for err in errori:
                                st.warning(err)
                    except Exception as e:
                        st.error(f"Errore lettura Excel: {e}")

                st.divider()
                with st.expander("➕ Crea singolo account manualmente", expanded=False):
                    with st.form("nuovo_utente_form"):
                        c1, c2 = st.columns(2)
                        nuova_email = c1.text_input("Email utente")
                        nuova_pass = c1.text_input("Password provvisoria (min. 6 caratteri)", type="password")
                        nome_scheda = c2.text_input("Nome esatto scheda Excel (es. Milazzo)")
                        nuovo_ruolo = c2.selectbox("Ruolo", ["filiale", "admin", "super_admin"])
                        if st.form_submit_button("Crea Utente"):
                            if nuova_email and len(nuova_pass) >= 6 and nome_scheda:
                                try:
                                    admin_client = init_supabase_admin()
                                    nu = admin_client.auth.admin.create_user({"email": nuova_email.strip(), "password": nuova_pass, "email_confirm": True})
                                    admin_client.table("profili_filiali").insert({
                                        "id": nu.user.id,
                                        "email": nuova_email.strip(),
                                        "ruolo": nuovo_ruolo,
                                        "nome_scheda_excel": nome_scheda.strip(),
                                        "password_visibile": nuova_pass,
                                        "must_change_password": (nuovo_ruolo == "filiale")
                                    }).execute()
                                    st.success(f"Utente {nuova_email} creato!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Errore: {e}")
                            else:
                                st.warning("Compila tutti i campi.")

                st.divider()
                st.subheader("🔑 Visualizza e Cambia Password Utente")
                admin_client = init_supabase_admin()
                utenti_res = admin_client.table("profili_filiali").select("id, email, ruolo, nome_scheda_excel, password_visibile, must_change_password").order("email").execute()

                if utenti_res.data:
                    opzioni_pw = {f"{u['email']} ({u['nome_scheda_excel']})": u for u in utenti_res.data}
                    col_p1, col_p2, col_p3 = st.columns([2, 1.5, 1])
                    with col_p1:
                        utente_scelto_label = st.selectbox("Seleziona utente:", list(opzioni_pw.keys()), key="sel_utente_pw")
                        utente_scelto = opzioni_pw[utente_scelto_label]
                        pw_attuale = utente_scelto.get("password_visibile") or "Non salvata"
                        st.caption(f"Password attuale registrata: **`{pw_attuale}`**")
                    with col_p2:
                        nuova_pw_reset = st.text_input("Nuova password provvisoria (min. 6 caratteri)", key="input_nuova_pw")
                    with col_p3:
                        st.write("")
                        st.write("")
                        if st.button("Aggiorna Password", type="primary", use_container_width=True):
                            if len(nuova_pw_reset.strip()) >= 6:
                                try:
                                    admin_client.auth.admin.update_user_by_id(utente_scelto["id"], {"password": nuova_pw_reset.strip()})
                                    admin_client.table("profili_filiali").update({
                                        "password_visibile": nuova_pw_reset.strip(),
                                        "must_change_password": (utente_scelto["ruolo"] == "filiale")
                                    }).eq("id", utente_scelto["id"]).execute()
                                    st.success(f"✅ Password aggiornata per {utente_scelto['email']}!")
                                    st.rerun()
                                except Exception as e_pw:
                                    st.error(f"Errore aggiornamento password: {e_pw}")
                            else:
                                st.warning("La password deve avere almeno 6 caratteri.")

                    st.divider()
                    st.subheader("Utenti registrati")
                    df_utenti = pd.DataFrame(utenti_res.data)[["email", "ruolo", "nome_scheda_excel", "password_visibile", "must_change_password"]]
                    df_utenti.columns = ["Email", "Ruolo", "Scheda Excel Associata", "Password Attuale", "Stato Password"]
                    df_utenti["Password Attuale"] = df_utenti["Password Attuale"].fillna("-")
                    df_utenti["Stato Password"] = df_utenti["Stato Password"].apply(lambda x: "⏳ Da cambiare" if x else "✅ Personale")
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
                render_sezione_anteprima(is_super_admin=True)

            with tab4:
                render_sezione_gran_totale()

            with tab5:
                render_sezione_classifica()

        else:
            # Ruolo: ADMIN (solo consultazione, Gran Totale Annuale e Classifica Operatori)
            tab_a1, tab_a2, tab_a3 = st.tabs([
                "👀 Anteprima Dati",
                "📊 Gran Totale Annuale",
                "🏆 Classifica Operatori"
            ])

            with tab_a1:
                render_sezione_anteprima(is_super_admin=False)

            with tab_a2:
                render_sezione_gran_totale()

            with tab_a3:
                render_sezione_classifica()

    # ==========================================
    # PANNELLO FILIALE (CON CAMBIO PASSWORD OBBLIGATORIO AL PRIMO ACCESSO)
    # ==========================================
    else:
        st.markdown(
            """
            <style>
            [data-testid="stToolbar"],
            [data-testid="stDecoration"],
            [data-testid="stStatusWidget"],
            #MainMenu {
                display: none !important;
                visibility: hidden !important;
            }
            header[data-testid="stHeader"] {
                background: transparent !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )

        mia_scheda = profilo["nome_scheda_excel"]
        must_change = profilo.get("must_change_password", True)

        if must_change:
            _, col_pw, _ = st.columns([1, 1.6, 1])
            with col_pw:
                st.title("🔐 Cambio Password Obbligatorio")
                st.info(
                    f"Benvenuto **{mia_scheda}**! Per motivi di sicurezza, al primo accesso è necessario "
                    "sostituire la password provvisoria con una password personale."
                )
                with st.form("cambio_password_primo_accesso"):
                    pw_nuova_1 = st.text_input("Nuova password (min. 6 caratteri)", type="password")
                    pw_nuova_2 = st.text_input("Conferma nuova password", type="password")
                    if st.form_submit_button("Salva nuova password e accedi", type="primary", use_container_width=True):
                        pw1 = pw_nuova_1.strip()
                        pw2 = pw_nuova_2.strip()
                        pw_vecchia = str(profilo.get("password_visibile") or "").strip()

                        if len(pw1) < 6:
                            st.error("La password deve contenere almeno 6 caratteri.")
                        elif pw1 != pw2:
                            st.error("Le due password non coincidono.")
                        elif pw_vecchia and pw1 == pw_vecchia:
                            st.error("La nuova password deve essere diversa da quella provvisoria.")
                        else:
                            try:
                                admin_client = init_supabase_admin()
                                admin_client.auth.admin.update_user_by_id(st.session_state.user.id, {"password": pw1})
                                admin_client.table("profili_filiali").update({
                                    "password_visibile": pw1,
                                    "must_change_password": False
                                }).eq("id", st.session_state.user.id).execute()

                                st.session_state.profilo["password_visibile"] = pw1
                                st.session_state.profilo["must_change_password"] = False
                                st.success("✅ Password aggiornata con successo! Accesso in corso...")
                                st.rerun()
                            except Exception as e_chg:
                                st.error(f"Errore durante il salvataggio: {e_chg}")
        else:
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
