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
            for index, row in df.iterrows():
                skcid = row["SKCID"]
                spec = row["规格属性"]

                sheet.cell(row=index+2, column=sheet.max_column + 1, value=skcid[:2])
                sheet.cell(row=index+2, column=sheet.max_column + 2, value=spec.split("/")[0] if pd.notna(spec) else "")
                sheet.cell(row=index+2, column=sheet.max_column + 3, value=spec.split("/")[1] if pd.notna(spec) else "")
                sheet.cell(row=index+2, column=sheet.max_column + 4, value=skcid.rsplit("-", 2)[0])
                sheet.cell(row=index+2, column=sheet.max_column + 5, value="白墨烫画")

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
