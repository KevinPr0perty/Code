import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.title("Spreadsheet Processor for Custom Columns")

uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read workbook for writing cells (keeps original formatting as much as possible)
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet = workbook.active

        # IMPORTANT: reset pointer before pandas reads, because openpyxl already consumed it
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.write("### Original Spreadsheet:")
        st.dataframe(df)

        if "规格属性" in df.columns and "SKCID" in df.columns:
            for index, row in df.iterrows():
                skcid = str(row.get("SKCID", "")).strip()
                spec = row.get("规格属性")

                # ✅ If 商家编码 exists and is not empty, use it instead of SKCID for extraction
                source_id = skcid
                if "商家编码" in df.columns:
                    merchant_code = row.get("商家编码")
                    if pd.notna(merchant_code) and str(merchant_code).strip() != "":
                        source_id = str(merchant_code).strip()

                # 款号编码 logic (from source_id)
                match = re.search(r"A\d", source_id)
                款号编码 = match.group() if match else "A2"

                # 颜色编码 and 尺寸编码 (from 规格属性)
                颜色编码 = spec.split("/")[0] if pd.notna(spec) else ""
                尺寸编码 = spec.split("/")[1] if pd.notna(spec) and "/" in str(spec) else ""

                # ✅ 图片编码 logic:
                # Fix: A2-20250703381-Navy-XL  -> 20250703381 (NOT XL)
                m = re.match(r"^A\d-(\d+)", source_id)
                if m:
                    图片编码 = m.group(1)
                elif "-" in source_id:
                    图片编码 = source_id.split("-")[0]
                else:
                    图片编码 = source_id

                # Write into Excel (row 2 corresponds to index 0)
                sheet[f"K{index+2}"] = 款号编码
                sheet[f"L{index+2}"] = 颜色编码
                sheet[f"M{index+2}"] = 尺寸编码
                sheet[f"N{index+2}"] = 图片编码
                sheet[f"O{index+2}"] = "白墨烫画"

            st.write("### Processed Spreadsheet (Preview):")
            st.dataframe(df)

            buffer = io.BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Processed Spreadsheet",
                buffer,
                "processed_spreadsheet.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("Uploaded file must contain '规格属性' and 'SKCID' columns.")

    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
