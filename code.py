import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Filler Tool", layout="wide")
st.title("Excel Filler Tool")

uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Creating a new DataFrame with the filled data
    filled_df = pd.DataFrame()

    # Extracting A# for Column 1
    def extract_a_number(value):
        try:
            if "A#" in str(value):
                return str(value).split()[0].split("A#")[0] + "A#"
            else:
                return "ERROR"
        except:
            return "ERROR"

    filled_df['Column 1'] = df.iloc[:, 0].apply(extract_a_number)
    filled_df['Column 2'] = ['Black/XL'] * len(df)
    filled_df['Column 3'] = ['Black/XL'] * len(df)

    # Extracting SKU ID without the extra parts with error handling
    def extract_sku(value):
        try:
            if "A#" in str(value):
                return str(value).split()[0]
            else:
                return "ERROR"
        except:
            return "ERROR"

    filled_df['Column 4'] = df.iloc[:, 0].apply(extract_sku)

    # Filling other columns with "白墨烫画"
    for i in range(5, len(df.columns) + 1):
        filled_df[f'Column {i}'] = ['白墨烫画'] * len(df)

    st.dataframe(filled_df)

    # Allowing download of the processed file
    output = filled_df.to_excel(index=False, engine='openpyxl')
    st.download_button("Download Filled Excel File", data=output, file_name="filled_excel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
