import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.title("Spreadsheet Processor for Custom Columns")

uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls"])


# ---------- helpers ----------

def get_header_map(ws, header_row=1):
    """Map header text -> column index (1-based)."""
    header_map = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            header_map[str(cell.value).strip()] = cell.col_idx
    return header_map


def is_nonempty(x):
    """True if cell value is not NaN and not empty after strip."""
    if x is None:
        return False
    if isinstance(x, float) and pd.isna(x):
        return False
    return str(x).strip() != ""


def split_color_size(spec):
    """
    Supports:
      Red/XL
      Red-XL
      Navy-Blue-XL (splits on LAST dash)
    """
    if spec is None or (isinstance(spec, float) and pd.isna(spec)):
        return "", ""

    s = str(spec).strip()
    if "/" in s:
        c, s2 = s.split("/", 1)
        return c.strip(), s2.strip()
    if "-" in s:
        c, s2 = s.rsplit("-", 1)
        return c.strip(), s2.strip()

    return s, ""


def extract_style_code(source_id):
    """
    款号编码: ONLY the first single digit after 'A'
      A2-... -> A2
      A8250... -> A8
    """
    m = re.search(r"A(\d)", source_id)
    return f"A{m.group(1)}" if m else "A2"


def extract_image_code(source_id):
    """
    图片编码 extraction:
      - A2-20250703381-Navy-XL -> 20250703381
      - Otherwise: first long digit sequence (>=6)
      - Otherwise: before first '-' if present
      - Otherwise: whole string
    """
    source_id = str(source_id).strip()

    m = re.match(r"^A\d+-(\d+)", source_id)
    if m:
        return m.group(1)

    m2 = re.search(r"\d{6,}", source_id)
    if m2:
        return m2.group(0)

    if "-" in source_id:
        return source_id.split("-", 1)[0]

    return source_id


# ---------- main ----------

if uploaded_file is not None:
    try:
        # Load workbook for writing (keeps formatting better)
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet = workbook.active

        # Reset pointer for pandas
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.write("### Original Spreadsheet")
        st.dataframe(df)

        # Required source columns
        if "规格属性" not in df.columns or "SKCID" not in df.columns:
            st.error("Uploaded file must contain '规格属性' and 'SKCID' columns.")
            st.stop()

        # Map headers in the real Excel file (prevents column shifting)
        header_map = get_header_map(sheet, header_row=1)

        target_headers = ["*款号编码", "*颜色编码", "*尺寸编码", "*图片编码", "*工艺类型"]
        missing = [h for h in target_headers if h not in header_map]
        if missing:
            st.error(f"Missing required target columns: {missing}")
            st.stop()

        col_style = header_map["*款号编码"]
        col_color = header_map["*颜色编码"]
        col_size  = header_map["*尺寸编码"]
        col_img   = header_map["*图片编码"]
        col_proc  = header_map["*工艺类型"]

        has_merchant_col = "商家编码" in df.columns

        for index, row in df.iterrows():
            skcid = str(row.get("SKCID", "")).strip()
            spec = row.get("规格属性")

            # ---------- 款号编码 source ----------
            # If 商家编码 exists and non-empty, use it; else use SKCID
            if has_merchant_col and is_nonempty(row.get("商家编码")):
                style_source = str(row.get("商家编码")).strip()
            else:
                style_source = skcid

            款号编码 = extract_style_code(style_source)

            # ---------- 颜色/尺寸 from 规格属性 ----------
            颜色编码, 尺寸编码 = split_color_size(spec)

            # ---------- 图片编码 source (YOUR REQUEST) ----------
            # Try 商家编码 first; if empty -> straight use SKCID
            if has_merchant_col and is_nonempty(row.get("商家编码")):
                img_source = str(row.get("商家编码")).strip()
            else:
                img_source = skcid

            图片编码 = extract_image_code(img_source)

            # Write into Excel (row 2 corresponds to df index 0)
            excel_row = index + 2
            sheet.cell(excel_row, col_style, 款号编码)
            sheet.cell(excel_row, col_color, 颜色编码)
            sheet.cell(excel_row, col_size,  尺寸编码)
            sheet.cell(excel_row, col_img,   图片编码)
            sheet.cell(excel_row, col_proc,  "白墨烫画")

        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)

        st.success("Done — 图片编码 now prefers 商家编码, otherwise falls back to SKCID.")

        st.download_button(
            "Download Processed Spreadsheet",
            buffer,
            "processed_spreadsheet.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
