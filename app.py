import io
import re
import pandas as pd
import streamlit as st

# ===== CONFIG =====
DEFAULT_MASTER_PATH = "hsn_master.xlsx"   # your bundled master workbook
RESULT_FILENAME_PREFIX = "HSN_Result_"
MAX_HEADER_SCAN_ROWS = 20
LOGO_PATH = "logo.png"                     # place logo.png beside app.py

DIGIT_RUN = re.compile(r"\d+")

# ===== Helpers =====
def keep_digits(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())

def extract_digit_tokens(text: str):
    tokens = []
    seen = set()
    for m in DIGIT_RUN.finditer(str(text)):
        tok = m.group(0)
        if 2 <= len(tok) <= 12 and tok not in seen:
            seen.add(tok)
            tokens.append(tok)
    return tokens

def cell_matches_hsn(cell_text: str, search_hsn: str, level: int) -> bool:
    tokens = extract_digit_tokens(cell_text)
    if level == 1:  # full exact
        return any(tok == search_hsn for tok in tokens)
    if level == 2:  # 4-digit prefix
        return any(tok.startswith(search_hsn) or tok == search_hsn for tok in tokens)
    if level == 3:  # 2-digit prefix
        return any(tok.startswith(search_hsn) or tok == search_hsn for tok in tokens)
    return False

def find_header_row(df: pd.DataFrame, max_scan_rows: int = MAX_HEADER_SCAN_ROWS) -> int:
    """Heuristic similar to your VBA."""
    max_rows = min(max_scan_rows, len(df))
    df_str = df.astype(str).applymap(lambda x: x.strip().lower())
    for r in range(max_rows):
        row_vals = list(df_str.iloc[r, :])
        from_found = any("from" in v for v in row_vals)
        to_found = any(v == "to" or " to " in v or "to" in v for v in row_vals)
        if from_found and to_found:
            return r
        a_val = df_str.iloc[r, 0] if df_str.shape[1] > 0 else ""
        if (a_val.startswith("s") and "no" in a_val) or a_val in {"s. no.", "s.no"}:
            return r
    return 0

def get_ordered_sheet_names(xls: pd.ExcelFile):
    names = xls.sheet_names
    rate_idx = next((i for i, n in enumerate(names) if n.strip().lower() == "rate change"), None)
    if rate_idx is not None:
        return [names[rate_idx]] + [n for i, n in enumerate(names) if i != rate_idx]
    return names

def search_sheet(df: pd.DataFrame, hsn: str):
    """Return dict: {header_row, header_df, matches}."""
    if df.empty:
        return {"header_row": 0, "header_df": None, "matches": []}

    header_row = find_header_row(df)
    df2 = df.copy()
    df2.columns = df2.iloc[header_row]
    df2 = df2.iloc[header_row + 1:].reset_index(drop=True)

    col_to_scan = None
    if df2.shape[1] >= 2:
        col_to_scan = df2.columns[1]

    matches = []
    for level in (1, 2, 3):
        tmp = []
        if level == 1:
            look = hsn
        elif level == 2 and len(hsn) >= 4:
            look = hsn[:4]
        else:
            look = hsn[:2]

        if col_to_scan is not None and col_to_scan in df2.columns:
            for idx, val in df2[col_to_scan].items():
                if pd.isna(val):
                    continue
                if cell_matches_hsn(str(val), look, level):
                    tmp.append(df2.iloc[idx])
        else:
            for idx, row in df2.iterrows():
                hit = False
                for val in row.values:
                    if pd.isna(val):
                        continue
                    if cell_matches_hsn(str(val), look, level):
                        hit = True
                        break
                if hit:
                    tmp.append(df2.iloc[idx])

        if tmp:
            matches = tmp
            break

    header_df = pd.DataFrame([df.iloc[header_row].values])
    header_df.columns = [f"col_{i+1}" for i in range(header_df.shape[1])]
    return {"header_row": header_row, "header_df": header_df, "matches": matches}

# -------- Excel formatting helpers --------
def fmt(wb, **kwargs):
    return wb.add_format(kwargs)

def write_pretty_block(ws, wb, row_ptr: int, sheet_name: str, header_df: pd.DataFrame, matches: list) -> int:
    title_fmt = fmt(wb, bold=True, font_size=12)
    hdr_fmt   = fmt(wb, bold=True, bg_color="#F2F2F2", border=1, bottom=1)
    cell_fmt  = fmt(wb, text_wrap=True, border=1)
    cell_alt  = fmt(wb, text_wrap=True, border=1, bg_color="#FBFBFB")

    row_ptr += 1
    ws.write(row_ptr, 0, sheet_name, title_fmt)
    row_ptr += 1

    max_cols = 0
    if header_df is not None and not header_df.empty:
        max_cols = header_df.shape[1]
        for c in range(max_cols):
            ws.write(row_ptr, c, "" if pd.isna(header_df.iloc[0, c]) else header_df.iloc[0, c], hdr_fmt)
    row_ptr += 1

    if matches:
        mdf = pd.DataFrame(matches)
        max_cols = max(max_cols, mdf.shape[1])
        for r in range(mdf.shape[0]):
            rf = cell_alt if (r % 2 == 0) else cell_fmt
            for c in range(mdf.shape[1]):
                val = mdf.iat[r, c]
                ws.write(row_ptr + r, c, "" if pd.isna(val) else val, rf)
        row_ptr += mdf.shape[0]

    row_ptr += 2
    return row_ptr

def build_details_workbook(xls_bytes_or_path, hsn: str, place_logo_in_excel: bool = True):
    """Build the Excel 'Details' sheet; return (bytes, found_any)."""
    if isinstance(xls_bytes_or_path, (bytes, bytearray, io.BytesIO)):
        xls = pd.ExcelFile(io.BytesIO(xls_bytes_or_path), engine="openpyxl")
    else:
        xls = pd.ExcelFile(xls_bytes_or_path, engine="openpyxl")

    sheet_order = get_ordered_sheet_names(xls)

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    wb = writer.book
    ws = wb.add_worksheet("Details")
    writer.sheets["Details"] = ws

    ws.hide_gridlines(2)

    center_bold   = fmt(wb, bold=True, align="center", valign="vcenter", text_wrap=True)
    top_disc_fmt  = fmt(wb, bold=True, font_size=12, align="left",  valign="top",   text_wrap=True)
    bottom_disc_fmt = fmt(wb, bold=True, align="center", valign="vcenter", text_wrap=True)

    # ======= HEADER AREA (more space so nothing overlaps) =======
    # Reserve rows 0..15; then place disclaimer at rows 14..15; start content after row 17.
    HEADER_END_ROW = 15

    for r in range(0, HEADER_END_ROW + 1):
        ws.set_row(r, 22)  # tweak if you want taller rows

    if place_logo_in_excel:
        try:
            ws.insert_image(
                0, 0, LOGO_PATH,
                {
                    "x_scale": 0.48, "y_scale": 0.48,  # shrink a bit
                    "x_offset": 6, "y_offset": 6,
                    "object_position": 2               # don't move/size with cells
                }
            )
        except Exception:
            pass

    # Top disclaimer FAR below the logo: rows 14..15, merged across A:E
    ws.merge_range(
        14, 0, 15, 4,
        "This document is for internal office use only. We shall not be responsible for any inaccuracies or misinterpretations.",
        top_disc_fmt
    )

    row_ptr = HEADER_END_ROW + 2   # start at row 17+2 = 17? (actually 15+2 = 17)
    found_any = False
    wrote_other_heading = False

    for sname in sheet_order:
        df = pd.read_excel(xls, sheet_name=sname, header=None, engine="openpyxl")
        result = search_sheet(df, hsn)
        matches = result["matches"]
        if not matches:
            continue

        found_any = True
        row_ptr = write_pretty_block(ws, wb, row_ptr, sname, result["header_df"], matches)

        if sname.strip().lower() == "rate change" and not wrote_other_heading:
            row_ptr += 1
            ws.merge_range(row_ptr, 0, row_ptr, 4, "Other important information related to HSN", center_bold)
            row_ptr += 1
            wrote_other_heading = True

    # Two-row gap BEFORE bottom disclaimer
    row_ptr += 2

    # ======= UPDATED bottom disclaimer text =======
    bottom_disclaimer = (
        "If your desired HSN code is not listed above, it implies that there has been no change in the applicable "
        "GST rate for that particular code. The said HSN code continues to fall under the existing 0%, 5%, or 18% tax brackets."
    )
    ws.merge_range(row_ptr, 0, row_ptr + 4, 4, bottom_disclaimer, bottom_disc_fmt)
    row_ptr += 5

    # Column widths (A:E)
    ws.set_column(0, 0, 8.82)   # A
    ws.set_column(1, 1, 23.82)  # B
    ws.set_column(2, 2, 84)     # C
    ws.set_column(3, 3, 13.18)  # D
    ws.set_column(4, 4, 13.18)  # E

    writer.close()
    return output.getvalue(), found_any

# ===== UI =====
st.set_page_config(page_title="HSN Search Tool", page_icon="🔎", layout="centered")

# Topbar/logo in the web app
cols = st.columns([1, 6, 1])
with cols[0]:
    try:
        st.image(LOGO_PATH, use_container_width=True)  # modern Streamlit param
    except Exception:
        pass
with cols[1]:
    st.title("HSN Search Tool")
    st.caption("Type an HSN (2–8 digits) and click **Search**. Get a polished Excel with your results.")
with cols[2]:
    pass

# Sidebar options
st.sidebar.header("Options")
use_uploaded = st.sidebar.checkbox("Upload an Excel instead of using bundled master", value=False)
place_logo_in_excel = st.sidebar.checkbox("Embed logo in Excel header", value=True)
uploaded_file = None
if use_uploaded:
    uploaded_file = st.sidebar.file_uploader("Upload .xlsx file", type=["xlsx"])

# Main controls
hsn_input = st.text_input("HSN Code", value="", max_chars=12, help="Only digits are considered")
do_search = st.button("Search")

if do_search:
    user_hsn = keep_digits(hsn_input)
    if len(user_hsn) < 2:
        st.error("Please enter at least 2 digits.")
    else:
        try:
            if use_uploaded:
                if uploaded_file is None:
                    st.error("Please upload an Excel file or uncheck the upload option.")
                    st.stop()
                data_bytes = uploaded_file.read()
                data, found = build_details_workbook(data_bytes, user_hsn, place_logo_in_excel)
            else:
                data, found = build_details_workbook(DEFAULT_MASTER_PATH, user_hsn, place_logo_in_excel)

            if not found:
                st.info(f"No matching rows found for HSN: {user_hsn} (tried full, then 4-digit, then 2-digit prefix).")

            st.success("Search complete. Download your Excel below.")
            st.download_button(
                label="Download Details.xlsx",
                data=data,
                file_name=f"{RESULT_FILENAME_PREFIX}{user_hsn}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

st.caption("Tip: Keep 'hsn_master.xlsx' next to this app for bundled mode. Or upload a workbook from the sidebar.")
