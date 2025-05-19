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
            # Fill only the specified columns without changing layout
            df["款号编码"] = df["SKCID"].apply(lambda x: x[:2] if pd.notna(x) else "")
            df["颜色编码"] = df["规格属性"].apply(lambda x: x.split("/")[0] if pd.notna(x) else "")
            df["尺寸编码"] = df["规格属性"].apply(lambda x: x.split("/")[1] if pd.notna(x) else "")
            df["图片编码"] = df["SKCID"].apply(lambda x: x.rsplit("-", 2)[0] if pd.notna(x) else "")
            df["工艺类型"] = "白墨烫画"

            st.write("### Processed Spreadsheet (Same Format):")
            st.dataframe(df)

            # Write back to the original Excel with styles preserved
            for row_idx, row in enumerate(df.itertuples(index=False), start=2):
                sheet.cell(row=row_idx, column=sheet.max_column + 1, value=row.款号编码)
                sheet.cell(row=row_idx, column=sheet.max_column + 1, value=row.颜色编码)
                sheet.cell(row=row_idx, column=sheet.max_column + 1, value=row.尺寸编码)
                sheet.cell(row=row_idx, column=sheet.max_column + 1, value=row.图片编码)
                sheet.cell(row=row_idx, column=sheet.max_column + 1, value=row.工艺类型)

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
