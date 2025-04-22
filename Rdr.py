import streamlit as st
import pandas as pd
import os

def process_rdr1_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)
    docnum_mapping = {}
    current_docnum = 1
    prev_ref = None
    
    def get_docnum(ref):
        nonlocal current_docnum, prev_ref
        if ref != prev_ref:
            if ref not in docnum_mapping:
                docnum_mapping[ref] = current_docnum
                current_docnum += 1
        prev_ref = ref
        return docnum_mapping[ref]
    
    df['DocNum'] = df['Customer Reference No'].apply(get_docnum)
    df['LineNum'] = df.groupby('DocNum').cumcount()
    
    rdr1_data = {
        "ParentKey": df['DocNum'].tolist(),
        "LineNum": df['LineNum'].tolist(),
        "ItemCode": df['Item Code'].tolist(),
        "SupplierCatNum": df['Part No'].tolist(),
        "Quantity": df['Quantity'].tolist(),
        "UnitPrice": df['Price'].tolist(),
        "AccountCode": [410000] * len(df),
        "TaxCode": df['Tax code'].tolist(),
        "WarehouseCode": df['Warehouse'].tolist()
    }
    rdr1_df = pd.DataFrame(rdr1_data)
    
    old_columns = ["DocNum", "LineNum", "ItemCode", "SubCatNum", "Quantity", "PriceBefDi", "AcctCode", "TaxCode", "WhsCode"]
    old_row = pd.DataFrame([old_columns], columns=rdr1_df.columns)
    rdr1_df = pd.concat([old_row, rdr1_df], ignore_index=True)
    
    output_excel_file = "RDR1.xlsx"
    rdr1_df.to_excel(output_excel_file, index=False, sheet_name='RDR1')
    
    output_txt_file = "RDR1.txt"
    with open(output_txt_file, "w") as txt_file:
        txt_file.write(rdr1_df.to_csv(sep='\t', index=False))
    
    return output_excel_file, output_txt_file

def process_ordr_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)
    docnum_mapping = {}
    current_docnum = 1
    prev_ref = None
    
    def get_docnum(ref):
        nonlocal current_docnum, prev_ref
        if ref != prev_ref:
            if ref not in docnum_mapping:
                docnum_mapping[ref] = current_docnum
                current_docnum += 1
        prev_ref = ref
        return docnum_mapping[ref]
    
    df['DocNum'] = df['Customer Reference No'].apply(get_docnum)
    df = df.drop_duplicates(subset=['DocNum', 'Customer Reference No'])
    
    ordr_data = {
        "DocNum": df['DocNum'].tolist(),
        "DocType": ["dDocument_Items"] * len(df),
        "Series": [882] * len(df),
        "NumAtCard": df['Customer Reference No'].tolist(),
        "DocDate": df['Document Date'].tolist(),
        "TaxDate": df['Tax Date'].tolist(),
        "DocDueDate": df.get('Document Due Date', pd.Series(['N/A'] * len(df))).tolist(),
        "DocCurrency": df['DocCur'].tolist(),
        "DocRate": df['Docrate'].tolist(),
        "CardCode": df['Customer CODE'].tolist(),
        "U_FRIEGHT": ["NA"] * len(df),
        "DutyStatus": ["Without Payment of Duty"] * len(df),
        "U_SALESCAT": ["Export"] * len(df),
        "U_DEPT": ["NA"] * len(df),
        "U_XmlFileStatus": [1] * len(df),
        "BPL_IDAssignedToInvoice": [3] * len(df),
        "U_MHXML": [1] * len(df)
    }
    ordr_df = pd.DataFrame(ordr_data)
    
    old_columns = list(ordr_df.columns)
    old_columns[old_columns.index("DocCurrency")] = "DocCurrency"
    old_columns[old_columns.index("BPL_IDAssignedToInvoice")] = "BPL_IDAssignedToInvoice"
    old_row = pd.DataFrame([old_columns], columns=ordr_df.columns)
    ordr_df = pd.concat([old_row, ordr_df], ignore_index=True)
    
    output_ordr_file = "Ordr.xlsx"
    ordr_df.to_excel(output_ordr_file, index=False, sheet_name='Ordr')
    
    output_txt_file = "Ordr.txt"
    with open(output_txt_file, "w") as txt_file:
        txt_file.write(ordr_df.to_csv(sep='\t', index=False))
    
    return output_ordr_file, output_txt_file

st.title("SO Automation")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    output_excel_file, output_txt_file = process_rdr1_excel(uploaded_file)
    output_ordr_file, output_ordr_txt_file = process_ordr_excel(uploaded_file)
    
    st.success("Files generated successfully!")
    
    with open(output_excel_file, "rb") as f:
        st.download_button(
            label="Download RDR1 Excel",
            data=f,
            file_name=output_excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with open(output_txt_file, "rb") as f:
        st.download_button(
            label="Download RDR1 TXT",
            data=f,
            file_name=output_txt_file,
            mime="text/plain"
        )
    
    with open(output_ordr_file, "rb") as f:
        st.download_button(
            label="Download ORDR Excel",
            data=f,
            file_name=output_ordr_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with open(output_ordr_txt_file, "rb") as f:
        st.download_button(
            label="Download ORDR TXT",
            data=f,
            file_name=output_ordr_txt_file,
            mime="text/plain"
        )
