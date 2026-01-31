import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.title("Spreadsheet Processor for Custom Columns")

uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls"])

def get_header_map(ws, header_row=1):
    """
    Map header text -> column index (1-based)
    """
    header_map = {}
    for cell in ws[header_row]:
        val = cell.value
        if val is not None:
            header_map[str(val).strip()] = cell.col_idx
    return header_map

def split_color_size(spec):
    """
    Supports:
      Red/XS
      Red-XS
      Navy-XL
    Strategy:
      - if '/' present: split on first '/'
      - elif '-' present: split on last '-' (safer if color contains dashes)
      - else: color=spec, size=""
    """
    if spec is None or (isinstance(spec, float) and pd.isna(spec)):
        return "", ""
    s = str(spec).strip()
    if "/" in s:
        parts = s.split("/", 1)
        return parts[0].strip(), parts[1].strip()
    if "-" in s:
        left, right = s.rsplit("-", 1)
        return left.strip(), right.strip()
    return s, ""

def extract_style_code(source_id):
    """
    款号编码: detect ANY number of digits after 'A'
    Examples:
      A2-20250703381-Navy-XL -> A2
      A8250523149R          -> A8250523149
    """
    m = re.search(r"A\d+", source_id)
    return m.group(0) if m else "A2"

def extract_image_code(source_id):
    """
    图片编码:
      - If starts with A + digits + '-' + digits -> take those digits after dash
        A2-20250703381-Navy-XL -> 20250703381
      - Else: try first long digit sequence (>=6)
      - Else: if '-' exists -> take first token
      - Else: whole string
    """
    m = re.match(r"^A\d+-(\d+)", source_id)
    if m:
        return m.group(1)

    m2 = re.search(r"\d{6,}", source_id)
    if m2:
        return m2.group(0)

    if "-" in source_id:
        return source_id.split("-", 1)[0]

    return source_id

if uploaded_file is not None:
    try:
        # Load workbook for writing (preserves formatting more than pandas)
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet = workbook.active

        # Reset pointer for pandas read
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.write("### Original Spreadsheet:")
        st.dataframe(df)

        # Basic required columns in dataframe
        if "规格属性" not in df.columns or "SKCID" not in df.columns:
            st.error("Uploaded file must contain '规格属性' and 'SKCID' columns.")
            st.stop()

        # Build header->col mapping from the actual Excel sheet
        header_map = get_header_map(sheet, header_row=1)

        required_targets = ["*款号编码", "*颜色编码", "*尺寸编码", "*图片编码", "*工艺类型"]
        missing_targets = [h for h in required_targets if h not in header_map]
        if missing_targets:
            st.error(f"Your sheet is missing these target columns: {missing_targets}")
            st.stop()

        # Column indices to write into
        col_style = header_map["*款号编码"]
        col_color = header_map["*颜色编码"]
        col_size  = header_map["*尺寸编码"]
        col_img   = header_map["*图片编码"]
        col_proc  = header_map["*工艺类型"]

        for index, row in df.iterrows():
            skcid = str(row.get("SKCID", "")).strip()
            spec = row.get("规格属性")

            # ✅ If 商家编码 exists and non-empty, use it instead of SKCID
            source_id = skcid
            if "商家编码" in df.columns:
                merchant_code = row.get("商家编码")
                if pd.notna(merchant_code) and str(merchant_code).strip() != "":
                    source_id = str(merchant_code).strip()

            款号编码 = extract_style_code(source_id)
            颜色编码, 尺寸编码 = split_color_size(spec)
            图片编码 = extract_image_code(source_id)

            excel_row = index + 2  # header is row 1

            sheet.cell(row=excel_row, column=col_style, value=款号编码)
            sheet.cell(row=excel_row, column=col_color, value=颜色编码)
            sheet.cell(row=excel_row, column=col_size,  value=尺寸编码)
            sheet.cell(row=excel_row, column=col_img,   value=图片编码)
            sheet.cell(row=excel_row, column=col_proc,  value="白墨烫画")

        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)

        st.success("Done — columns filled by header names (no shifting).")

        st.download_button(
            "Download Processed Spreadsheet",
            buffer,
            "processed_spreadsheet.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
