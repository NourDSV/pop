# app.py
import streamlit as st
import openpyxl
import datetime as dt
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.workbook.properties import CalcProperties
import re
from pathlib import Path

st.set_page_config(page_title="POP files", layout="centered")
st.title("Création des fichiers POP")
st.subheader("📂 Déposez tous les fichiers POP ici")

st.markdown("""
L’application :
- 🔄 met à jour les **jours équivalents**
- 🧮 applique les **formules Excel**
- 📦 génère une **archive ZIP** contenant tous les fichiers créés
""")

# -------------------------
# Uploads
# -------------------------

MAPPING_PATH = Path("Calendrier_comparatif_2026_vs_2025.xlsx")
mapping_file = BytesIO(MAPPING_PATH.read_bytes())
mapping_file.name = MAPPING_PATH.name

source_PATH = Path("Book2.xlsx")
source_file = BytesIO(source_PATH.read_bytes())
source_file.name = source_PATH.name

target_files = st.file_uploader("Deposer ici les fichers POP",type=["xlsx"],accept_multiple_files=True)

# ✅ Show how many target files were uploaded
if target_files:
    st.info(f"📦 {len(target_files)} file(s) uploaded.")

# -------------------------
# Config (your rules)
# -------------------------
RANGES_TO_COPY = ["E32:AE48", "E58:AE72"]
HEADER_ROW_2026 = 31
HEADER_ROW_2025 = 3
START_COL = 5     # E
END_COL   = 31    # AE

# -------------------------
# Helpers (range copy)
# -------------------------
def col_to_num(col: str) -> int:
    col = col.upper()
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n

def parse_range(a1_range: str):
    m = re.fullmatch(r"\s*([A-Za-z]+)(\d+)\s*:\s*([A-Za-z]+)(\d+)\s*", a1_range)
    if not m:
        raise ValueError(f"Invalid range '{a1_range}'. Use format like E32:AE48")
    c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
    row1, row2 = int(r1), int(r2)
    col1, col2 = col_to_num(c1), col_to_num(c2)
    if row2 < row1 or col2 < col1:
        raise ValueError(f"Range '{a1_range}' must end bottom-right of the start")
    return row1, col1, row2, col2

def copy_range(ws_src, ws_dst, a1_range: str):
    row1, col1, row2, col2 = parse_range(a1_range)
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            src_cell = ws_src.cell(row=row, column=col)
            dst_cell = ws_dst.cell(row=row, column=col)

            
            dst_cell.value = src_cell.value

            
            dst_cell.number_format = "General"


# -------------------------
# Helpers (date parsing + mapping)
# -------------------------
def cell_to_date(cell, default_year=None):
    v = cell.value
    if v is None:
        return None

    if cell.is_date:
        if isinstance(v, dt.datetime):
            return v.date()
        if isinstance(v, dt.date):
            return v

    if isinstance(v, (int, float)):
        try:
            d = from_excel(v)
            return d.date() if isinstance(d, dt.datetime) else d
        except Exception:
            return None

    if isinstance(v, str):
        s = v.strip().replace("'", "")
        try:
            return dt.datetime.strptime(s, "%d/%m/%Y").date()
        except ValueError:
            pass
        try:
            dm = dt.datetime.strptime(s, "%d/%m")
            if default_year is None:
                return None
            return dt.date(default_year, dm.month, dm.day)
        except ValueError:
            return None

    return None

def norm_full_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        try:
            d = from_excel(v)
            return d.date() if isinstance(d, dt.datetime) else d
        except Exception:
            return None
    if isinstance(v, str):
        s = v.strip().replace("'", "")
        try:
            return dt.datetime.strptime(s, "%d/%m/%Y").date()
        except ValueError:
            return None
    return None

def read_mapping(file_bytes: bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    mapping = {}
    for r in range(1, ws.max_row + 1):
        d26 = norm_full_date(ws.cell(row=r, column=1).value)  # A
        d25 = norm_full_date(ws.cell(row=r, column=2).value)  # B
        if d26 and d25:
            mapping[d26] = d25
    return mapping

def build_date_to_col(ws, header_row, default_year=None):
    result = {}
    for c in range(START_COL, END_COL + 1):
        d = cell_to_date(ws.cell(row=header_row, column=c), default_year=default_year)
        if d:
            result[d] = c
    return result

def apply_mapping_formulas(ws, mapping):
    col_2026 = build_date_to_col(ws, HEADER_ROW_2026, default_year=2026)
    col_2025 = build_date_to_col(ws, HEADER_ROW_2025, default_year=2025)

    for d26, d25 in mapping.items():
        if d26 not in col_2026 or d25 not in col_2025:
            continue

        tgt_col = col_2026[d26]
        src_col = col_2025[d25]

        T = get_column_letter(tgt_col)
        S = get_column_letter(src_col)

        # IMPORTANT: use commas in file formulas
        ws[f"{T}35"].value = (
            f'=IF({S}6*(1+$E$52)+{S}7*(1+$H$52)=0,0,'
            f'{S}6*(1+$E$52)+{S}7*(1+$H$52))'
        )
        ws[f"{T}41"].value = f'=IF({S}6*(1+$E$52)="","",{S}6*(1+$E$52))'
        ws[f"{T}47"].value = (
            f'=IF({S}23*(1+$N$52)+{S}24*(1+$K$52)=0,0,'
            f'{S}23*(1+$N$52)+{S}24*(1+$K$52))'
        )

def force_recalc_on_open(wb: openpyxl.Workbook):
    if wb.calculation is None:
        wb.calculation = CalcProperties(fullCalcOnLoad=True)
    else:
        wb.calculation.fullCalcOnLoad = True

# -------------------------
# Main action
# -------------------------
ready = mapping_file and source_file and target_files and len(target_files) > 0
if not ready:
    st.caption("Upload  at least 1 pop file.")
    st.stop()

if st.button("✅ Apply and generate ZIP"):
    n_files = len(target_files)

    # ✅ Waiting widget
    with st.spinner(f"Processing {n_files} file(s)..."):
        try:
            mapping = read_mapping(mapping_file.getvalue())
            if not mapping:
                st.error("Mapping file is empty or dates are not readable. Column A/B must contain valid dates.")
                st.stop()

            wb_src = openpyxl.load_workbook(BytesIO(source_file.getvalue()), data_only=False)
            ws_src = wb_src[wb_src.sheetnames[0]]

            zip_buffer = BytesIO()
            updated_count = 0

            with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as zf:
                for up in target_files:
                    wb_dst = openpyxl.load_workbook(BytesIO(up.getvalue()), data_only=False)
                    ws_dst = wb_dst[wb_dst.sheetnames[0]]

                    # 1) Copy formula ranges
                    for r in RANGES_TO_COPY:
                        copy_range(ws_src, ws_dst, r)

                    # 2) Apply mapping formulas
                    apply_mapping_formulas(ws_dst, mapping)

                    # 3) Force recalc
                    force_recalc_on_open(wb_dst)

                    out = BytesIO()
                    wb_dst.save(out)
                    file_bytes = out.getvalue()

                    original_name = up.name.rsplit(".", 1)[0]
                    zip_name = f"{original_name}_updated.xlsx"
                    zf.writestr(zip_name, file_bytes)

                    updated_count += 1

            zip_buffer.seek(0)

        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    # ✅ Finished message
    st.success(f"✅ All done — {updated_count}/{n_files} file(s) updated.")

    st.download_button(
        "⬇️ Download ZIP (all updated files)",
        data=zip_buffer.getvalue(),
        file_name="updated_excels.zip",
        mime="application/zip",
    )
