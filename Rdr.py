import streamlit as st
import pandas as pd
import os
from datetime import datetime
import calendar

# Ensure the upload directory exists
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ===========================
# ?? SO DOMESTICS Functions
# ===========================
def process_ordr_domestic(df):
    ordr_column_mapping = {
        'DocNum': 'CustomerRefNo',
        'DocType': 'ItemDescription',
        'Series': 'Series',
        'NumAtCard': 'CustomerRefNo',
        'DocDate': 'Document Date',
        'TaxDate': 'Tax Date',
        'DocDueDate': 'Document Due Date',
        'CardCode': 'Customer Code',
        'U_FRIEGHT': 'Frieght',
        'U_SALESCAT': 'Sales Category',
        'U_DEPT': 'Department',
        'U_MHXML': 'Part No',
    }

    df_ordr = df[list(ordr_column_mapping.values())].rename(columns=ordr_column_mapping)
    df_ordr.columns = ['DocNum', 'DocType', 'Series', 'NumAtCard', 'DocDate', 'TaxDate', 'DocDueDate', 'CardCode',
                       'U_FRIEGHT', 'U_SALESCAT', 'U_DEPT', 'U_MHXML']

    df_ordr['U_XmlFileStatus'] = 1
    df_ordr['BPL_IDAssignedToInvoice'] = 3

    new_columns_order = ['Series', 'NumAtCard', 'DocDate', 'TaxDate', 'DocDueDate', 'CardCode',
                         'U_FRIEGHT', 'U_SALESCAT', 'U_DEPT', 'U_XmlFileStatus', 'BPL_IDAssignedToInvoice', 'U_MHXML']
    df_final = df_ordr[new_columns_order]

    df_final.insert(0, 'DocNum', range(1, len(df_final) + 1))
    df_final.insert(1, 'DocType', 'dDocument_Items')
    df_final['Series'] = 882

    current_date = datetime.now().strftime('%Y%m%d')
    df_final['DocDate'] = current_date
    df_final['TaxDate'] = current_date
    last_day = calendar.monthrange(datetime.now().year, datetime.now().month)[1]
    df_final['DocDueDate'] = datetime.now().replace(day=last_day).strftime('%Y%m%d')

    df_final['U_MHXML'] = df_final['CardCode'].apply(lambda x: '0' if '16M' in x or '16G' in x else '1')

    headers_as_row = pd.DataFrame([df_final.columns], columns=df_final.columns)
    df_final = pd.concat([headers_as_row, df_final], ignore_index=True)
    df_final.at[0, 'BPL_IDAssignedToInvoice'] = 'BPLId'

    return df_final


def process_rdr1_domestic(df):
    rdr1_column_mapping = {
        'ItemCode': 'Item Code',
        'SubCatNum': 'Part No',
        'Quantity': 'Quantity',
        'PriceBefDi': 'Price',
        'TaxCode': 'TaxCode',
        'WhsCode': 'Warehouse',
    }

    df_rdr1 = df[list(rdr1_column_mapping.values())].rename(columns=rdr1_column_mapping)
    df_rdr1.insert(0, 'DocNum', range(1, len(df_rdr1) + 1))
    df_rdr1.insert(1, 'LineNum', 0)

    df_rdr1.columns = ['DocNum', 'LineNum', 'ItemCode', 'SubCatNum', 'Quantity', 'PriceBefDi', 'TaxCode', 'WhsCode']
    header1 = ['ParentKey', 'LineNum', 'ItemCode', 'SupplierCatNum', 'Quantity', 'UnitPrice', 'TaxCode', 'WarehouseCode']
    header2 = ['DocNum', 'LineNum', 'ItemCode', 'SubCatNum', 'Quantity', 'PriceBefDi', 'TaxCode', 'WhsCode']

    df_final = pd.concat([
        pd.DataFrame([header1], columns=df_rdr1.columns),
        pd.DataFrame([header2], columns=df_rdr1.columns),
        df_rdr1
    ], ignore_index=True)

    return df_final

# ===========================
# ?? SO EXPORT Functions
# ===========================
def process_rdr1_export(uploaded_file):
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

    return rdr1_df


def process_ordr_export(uploaded_file):
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

    old_row = pd.DataFrame([ordr_df.columns.tolist()], columns=ordr_df.columns)
    ordr_df = pd.concat([old_row, ordr_df], ignore_index=True)

    return ordr_df

# ===========================
# ?? MAIN UI
# ===========================
st.title("?? SO Automation")

col1, col2 = st.columns(2)
with col1:
    dom_btn = st.button("?? SO Domestics")
with col2:
    exp_btn = st.button("?? SO Export")

uploaded_file = st.file_uploader("?? Upload Excel File", type=["xlsx"])

if uploaded_file:
    if dom_btn:
        df = pd.read_excel(uploaded_file)
        df_ordr = process_ordr_domestic(df)
        df_rdr1 = process_rdr1_domestic(df)

        ordr_xlsx = os.path.join(UPLOAD_DIR, "ORDR_Domestic.xlsx")
        rdr1_xlsx = os.path.join(UPLOAD_DIR, "RDR1_Domestic.xlsx")

        df_ordr.to_excel(ordr_xlsx, index=False)
        df_rdr1.to_excel(rdr1_xlsx, index=False, header=False)

        st.download_button("?? Download ORDR (Domestic)", open(ordr_xlsx, "rb"), file_name="ORDR_Domestic.xlsx")
        st.download_button("?? Download RDR1 (Domestic)", open(rdr1_xlsx, "rb"), file_name="RDR1_Domestic.xlsx")

    elif exp_btn:
        rdr1_df = process_rdr1_export(uploaded_file)
        ordr_df = process_ordr_export(uploaded_file)

        rdr1_file = os.path.join(UPLOAD_DIR, "RDR1_Export.xlsx")
        ordr_file = os.path.join(UPLOAD_DIR, "ORDR_Export.xlsx")

        rdr1_df.to_excel(rdr1_file, index=False)
        ordr_df.to_excel(ordr_file, index=False)

        st.download_button("?? Download RDR1 (Export)", open(rdr1_file, "rb"), file_name="RDR1_Export.xlsx")
        st.download_button("?? Download ORDR (Export)", open(ordr_file, "rb"), file_name="ORDR_Export.xlsx")
