import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.title("Spreadsheet Processor for Custom Columns")

uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet = workbook.active

        df = pd.read_excel(uploaded_file, engine='openpyxl')

        st.write("### Original Spreadsheet:")
        st.dataframe(df)

        if "规格属性" in df.columns and "SKCID" in df.columns:
            for index, row in df.iterrows():
                skcid = str(row["SKCID"])
                spec = row["规格属性"]

                # 款号编码 logic
                match = re.search(r"A\d", skcid)
                款号编码 = match.group() if match else "A2"

                # 颜色编码 and 尺寸编码
                颜色编码 = spec.split("/")[0] if pd.notna(spec) else ""
                尺寸编码 = spec.split("/")[1] if pd.notna(spec) and "/" in spec else ""

                # 图片编码 logic
                if re.match(r"A\d-\d+", skcid):
                    图片编码 = skcid.split("-")[-1]
                elif "-" in skcid:
                    图片编码 = skcid.split("-")[0]
                else:
                    图片编码 = skcid

                sheet[f"K{index+2}"] = 款号编码
                sheet[f"L{index+2}"] = 颜色编码
                sheet[f"M{index+2}"] = 尺寸编码
                sheet[f"N{index+2}"] = 图片编码
                sheet[f"O{index+2}"] = "白墨烫画"

            st.write("### Processed Spreadsheet (Same Format):")
            st.dataframe(df)

            buffer = io.BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            st.download_button(
                "Download Processed Spreadsheet",
                buffer,
                "processed_spreadsheet.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.error("Uploaded file must contain '规格属性' and 'SKCID' columns.")

    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
