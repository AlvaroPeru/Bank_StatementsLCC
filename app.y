import streamlit as st
import pdfplumber
import pandas as pd
import re, io, tempfile
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="RBC → Excel", page_icon="🏦", layout="centered")
st.title("🏦 RBC Bank Statements → Excel")
st.caption("Sube uno o varios PDFs de RBC — descarga todos los movimientos en Excel")

MONTH_MAP = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
             "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
DATE = re.compile(r"^(\d{1,2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", re.I)

def get_years(text):
    pat = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{1,2}[,\s]+(\d{4})", re.I)
    return [(MONTH_MAP[m.group(1).capitalize()[:3]], int(m.group(2))) for m in pat.finditer(text)]

def best_year(month, anchors):
    if not anchors: return datetime.now().year
    for pm, py in anchors:
        if pm == month: return py
    return min(anchors, key=lambda x: abs(x[0]-month))[1]

def to_num(s):
    s = s.strip()
    return float(s.replace(",","")) if re.match(r"^[\d,]+\.\d{2}$", s) else None

def parse_pdf(path):
    raw = []  # list of {date, desc, debit, credit, balance} — one per PDF row
    with pdfplumber.open(path) as pdf:
        anchors = get_years("\n".join(p.extract_text() or "" for p in pdf.pages))
        cur_date = None

        for page in pdf.pages:
            words = page.extract_words()
            if not words: continue

            # Detect column x-positions from header row
            x_desc = x_debit = x_credit = x_bal = None
            for w in words:
                t = w["text"]
                if t == "Description":  x_desc   = w["x0"]
                if "Cheques"   in t:    x_debit  = w["x0"]
                if "Deposits"  in t:    x_credit = w["x0"]
                if t == "Balance($)":   x_bal    = w["x0"]
            if not x_debit: continue

            # Find activity section start Y
            act_y = next((w["top"] for w in words if "AccountActivity" in w["text"].replace(" ","")), 0)

            # Group words by row Y
            by_y = defaultdict(list)
            for w in words:
                if w["top"] > act_y:
                    by_y[round(w["top"])].append(w)

            for y in sorted(by_y):
                rw = sorted(by_y[y], key=lambda w: w["x0"])
                line = "".join(w["text"] for w in rw)

                if re.search(r"Description|Cheques|Deposits|Balance\(\$\)|Closingbalance|AccountFees|\dof\d|Openingbalance", line, re.I):
                    continue

                # Assign words to columns by x position
                date_t = desc_t = debit_t = credit_t = bal_t = ""
                for w in rw:
                    x, t = w["x0"], w["text"]
                    if x < x_desc:                  date_t   += t + " "
                    elif x < x_debit:               desc_t   += t + " "
                    elif x < x_credit:              debit_t  += t + " "
                    elif x < x_bal:                 credit_t += t + " "
                    else:                           bal_t    += t + " "

                # Parse date if present
                dm = DATE.match(date_t.strip().replace(" ",""))
                if dm:
                    day = int(dm.group(1)); mon = MONTH_MAP[dm.group(2).capitalize()[:3]]
                    cur_date = datetime(best_year(mon, anchors), mon, day).date()

                if not cur_date: continue

                # Clean description: remove Airbnb ref numbers and hex hashes
                desc = re.sub(r"\b\d{9,15}\b", "", desc_t).strip()
                desc = re.sub(r"\b[0-9a-f]{30,}\b", "", desc).strip()

                # Skip rows that are purely reference numbers with no amounts
                if not desc and not debit_t.strip() and not credit_t.strip() and not bal_t.strip():
                    continue

                raw.append({
                    "date":    cur_date,
                    "desc":    desc,
                    "debit":   to_num(debit_t.strip()),
                    "credit":  to_num(credit_t.strip()),
                    "balance": to_num(bal_t.strip()),
                })

    # Merge: if a row has desc but no amounts, merge with the next row that has amounts
    merged = []
    i = 0
    while i < len(raw):
        r = raw[i]
        has_amt = r["debit"] or r["credit"] or r["balance"]
        if r["desc"] and not has_amt:
            # Look ahead for the amounts row
            if i+1 < len(raw) and raw[i+1]["date"] == r["date"] and not raw[i+1]["desc"]:
                nxt = raw[i+1]
                merged.append({**r, "debit": nxt["debit"], "credit": nxt["credit"], "balance": nxt["balance"]})
                i += 2
                continue
        merged.append(r)
        i += 1

    # Remove leftover orphan amount-only rows (already merged above)
    result = [r for r in merged if r["desc"]]

    return [{
        "Fecha":       r["date"],
        "Descripcion": r["desc"],
        "Credito":     r["credit"],
        "Debito":      r["debit"],
        "Balance":     r["balance"],
    } for r in result]

# ── UI ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Sube los PDFs de RBC", type="pdf", accept_multiple_files=True,
    help="Puedes subir varios meses a la vez"
)

if uploaded:
    all_rows = []
    for f in uploaded:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(f.read())
        try:
            rows = parse_pdf(Path(tmp.name))
            for r in rows: r["Archivo"] = f.name
            all_rows.extend(rows)
            st.success(f"✅ {f.name} → {len(rows)} transacciones")
        except Exception as e:
            st.error(f"❌ {f.name}: {e}")

    if all_rows:
        df = pd.DataFrame(all_rows).reset_index(drop=True)
        df["Fecha"] = pd.to_datetime(df["Fecha"])
        df["Año"]  = df["Fecha"].dt.year
        df["Mes"]  = df["Fecha"].dt.month
        df["Dia"]  = df["Fecha"].dt.day
        df = df[["Año","Mes","Dia","Fecha","Descripcion","Credito","Debito","Balance","Archivo"]]

        st.divider()
        st.subheader(f"📋 {len(df)} transacciones")
        st.dataframe(df, use_container_width=True, hide_index=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Movimientos")
            ws = writer.sheets["Movimientos"]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF", name="Arial")
                cell.fill = PatternFill("solid", fgColor="1F3864")
                cell.alignment = Alignment(horizontal="center")
            for col, w in zip("ABCDEFGHI", [6,5,5,14,52,12,12,12,30]):
                ws.column_dimensions[col].width = w
            for row in ws.iter_rows(min_row=2):
                row[3].number_format = "YYYY-MM-DD"
                for c in [5,6,7]: row[c].number_format = '#,##0.00'
            ws.freeze_panes = "A2"
        buf.seek(0)

        st.download_button(
            "⬇️ Descargar Excel",
            data=buf, file_name="RBC_movimientos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
