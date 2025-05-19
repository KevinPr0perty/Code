import streamlit as st
import pandas as pd
import openpyxl
import io

st.title("Spreadsheet Processor for Custom Columns")

# File upload
uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Load the file with original formatting including styles
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet = workbook.active

        # Process the data using pandas
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        st.write("### Original Spreadsheet:")
        st.dataframe(df)

        # Preserve original structure
        if "规格属性" in df.columns and "SKCID" in df.columns:
            # Correctly update the specified columns
            for index, row in df.iterrows():
                skcid = row["SKCID"]
                spec = row["规格属性"]

                # Directly update specified cells without adding new rows or columns
                sheet[f"K{index+2}"] = skcid[:2]  # 款号编码
                sheet[f"L{index+2}"] = spec.split("/")[0] if pd.notna(spec) else ""  # 颜色编码
                sheet[f"M{index+2}"] = spec.split("/")[1] if pd.notna(spec) else ""  # 尺寸编码
                sheet[f"N{index+2}"] = skcid.rsplit("-", 2)[0]  # 图片编码
                sheet[f"O{index+2}"] = "白墨烫画"  # 工艺类型

            st.write("### Processed Spreadsheet (Same Format):")
            st.dataframe(df)

            # Save processed spreadsheet with styles preserved
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
