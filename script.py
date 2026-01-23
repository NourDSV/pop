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
# Uploads (local bundled files)
# -------------------------
MAPPING_PATH = Path("Calendrier_comparatif_2026_vs_2025.xlsx")
mapping_file = BytesIO(MAPPING_PATH.read_bytes())
mapping_file.name = MAPPING_PATH.name

SOURCE_PATH = Path("Book2.xlsx")  # <-- your NEW "max" source template (E..AE days, AF Total)
source_file = BytesIO(SOURCE_PATH.read_bytes())
source_file.name = SOURCE_PATH.name

target_files = st.file_uploader("Deposer ici les fichers POP", type=["xlsx"], accept_multiple_files=True)

if target_files:
    st.info(f"📦 {len(target_files)} file(s) uploaded.")

# -------------------------
# Config (your rules)
# -------------------------
HEADER_ROW_2026 = 31
HEADER_ROW_2025 = 3

# Where "Total" label is in the target (header row)
TOTAL_HEADER_ROW = 30

# Day columns in the MAX source template
START_COL = 5      # E
MAX_DAY_COL = 31   # AE  (max day columns)
SOURCE_TOTAL_COL = 32  # AF (Total column in the source template)

# Blocks to copy (rows) - dynamic columns will be used
COPY_BLOCKS = [
    (32, 48),
    (58, 72),
]

# Rows where we rewrite formulas with mapping (only on day columns)
MAPPING_ROWS = (35, 41, 47)

# -------------------------
# Helpers (range + copy)
# -------------------------
def copy_block(ws_src, ws_dst, r1, r2, c1, c2):
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            dst_cell = ws_dst.cell(row=row, column=col)
            dst_cell.value = ws_src.cell(row=row, column=col).value
            dst_cell.number_format = "General"  # avoid formulas-as-text

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

def find_last_day_col(ws):
    """
    Find the last column (E..AE) that contains a readable date in row 31.
    """
    last = None
    for c in range(START_COL, MAX_DAY_COL + 1):
        d = cell_to_date(ws.cell(row=HEADER_ROW_2026, column=c), default_year=2026)
        if d:
            last = c
    if last is None:
        raise ValueError(f"No 2026 date found in row {HEADER_ROW_2026} between E and AE.")
    return last

def build_date_to_col(ws, header_row, end_col, default_year=None):
    result = {}
    for c in range(START_COL, end_col + 1):
        d = cell_to_date(ws.cell(row=header_row, column=c), default_year=default_year)
        if d:
            result[d] = c
    return result

def apply_mapping_formulas(ws, mapping, end_day_col):
    """
    Apply mapping formulas only on day columns E..end_day_col (NOT Total column).
    """
    col_2026 = build_date_to_col(ws, HEADER_ROW_2026, end_day_col, default_year=2026)
    col_2025 = build_date_to_col(ws, HEADER_ROW_2025, end_day_col, default_year=2025)

    for d26, d25 in mapping.items():
        if d26 not in col_2026 or d25 not in col_2025:
            continue

        tgt_col = col_2026[d26]
        src_col = col_2025[d25]

        T = get_column_letter(tgt_col)
        S = get_column_letter(src_col)

        # IMPORTANT: use commas in file formulas (not semicolons)
        ws[f"{T}35"].value = (
            f'=IF({S}6*(1+$E$52)+{S}7*(1+$H$52)=0,0,'
            f'{S}6*(1+$E$52)+{S}7*(1+$H$52))'
        )
        ws[f"{T}41"].value = f'=IF({S}6*(1+$E$52)="","",{S}6*(1+$E$52))'
        ws[f"{T}47"].value = (
            f'=IF({S}23*(1+$N$52)+{S}24*(1+$K$52)=0,0,'
            f'{S}23*(1+$N$52)+{S}24*(1+$K$52))'
        )

def rewrite_total_formula(formula: str, source_total_letter: str, target_total_letter: str) -> str:
    """
    Replace all references to the source Total column (e.g. AF) with the target Total column (e.g. Y/AA/...).
    Handles AF32, AF$32, $AF32, $AF$32
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    pat = re.compile(rf'(\$?){source_total_letter}(\$?\d+)')
    return pat.sub(rf'\1{target_total_letter}\2', formula)

def copy_total_column(ws_src, ws_dst, target_total_col, row_from, row_to):
    """
    Copy the Total column from source (AF) into the target total column (last_day+1),
    and rewrite formulas that reference AF to reference the actual total column.
    """
    src_total_letter = get_column_letter(SOURCE_TOTAL_COL)   # "AF"
    tgt_total_letter = get_column_letter(target_total_col)   # e.g. "Y" / "AA" / ...

    # Write header label
    ws_dst.cell(row=TOTAL_HEADER_ROW, column=target_total_col).value = "Total"

    for r in range(row_from, row_to + 1):
        src_cell = ws_src.cell(row=r, column=SOURCE_TOTAL_COL)
        dst_cell = ws_dst.cell(row=r, column=target_total_col)

        v = src_cell.value
        if isinstance(v, str) and v.startswith("="):
            v = rewrite_total_formula(v, src_total_letter, tgt_total_letter)

        dst_cell.value = v
        dst_cell.number_format = "General"

def force_recalc_on_open(wb: openpyxl.Workbook):
    if wb.calculation is None:
        wb.calculation = CalcProperties(fullCalcOnLoad=True)
    else:
        wb.calculation.fullCalcOnLoad = True

# -------------------------
# Main action
# -------------------------
ready = target_files and len(target_files) > 0
if not ready:
    st.caption("Upload at least 1 pop file.")
    st.stop()

if st.button("✅ Apply and generate ZIP"):
    n_files = len(target_files)

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
            total = len(target_files)

            progress = st.progress(0)

            with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as zf:
                for up in target_files:
                    wb_dst = openpyxl.load_workbook(BytesIO(up.getvalue()), data_only=False)
                    ws_dst = wb_dst[wb_dst.sheetnames[0]]

                    # 1) Detect last day column in target (E..AE)
                    last_day_col = find_last_day_col(ws_dst)

                    # 2) Copy day blocks from source template up to last_day_col
                    for (r1, r2) in COPY_BLOCKS:
                        copy_block(ws_src, ws_dst, r1, r2, START_COL, last_day_col)

                    # 3) Copy Total column from source (AF) to target at (last_day_col + 1)
                    target_total_col = last_day_col + 1
                    copy_total_column(ws_src, ws_dst, target_total_col, row_from=32, row_to=72)

                    # 4) Apply mapping only on day columns E..last_day_col
                    apply_mapping_formulas(ws_dst, mapping, end_day_col=last_day_col)

                    # 5) Force recalc
                    force_recalc_on_open(wb_dst)

                    # Save
                    out = BytesIO()
                    wb_dst.save(out)
                    file_bytes = out.getvalue()

                    original_name = up.name.rsplit(".", 1)[0]
                    zip_name = f"{original_name}_updated.xlsx"
                    zf.writestr(zip_name, file_bytes)

                    updated_count += 1
                    progress.progress(updated_count / total)

            zip_buffer.seek(0)

        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    st.success(f"✅ All done — {updated_count}/{n_files} file(s) updated.")

    st.download_button(
        "⬇️ Download ZIP (all updated files)",
        data=zip_buffer.getvalue(),
        file_name="updated_excels.zip",
        mime="application/zip",
    )
