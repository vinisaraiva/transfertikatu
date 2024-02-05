import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

st.set_page_config(
    page_title="App TIKATU",
    page_icon="📑",
)

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def insert_data_to_sheet(client, df, sheet_url):
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    data_list = df.applymap(lambda x: str(x) if pd.notnull(x) else '').values.tolist()
    next_row = len(worksheet.get_all_values()) + 1
    worksheet.insert_rows(data_list, next_row)

def main():
    st.title("Upload e Inserção de Arquivo Excel no Google Sheets")
    
    # CSS para personalizar os botões
    st.markdown("""
        <style>
        .stButton>button {
            color: white;
            background-color: #084d6e;
            font-size: 16px;
            font-weight: bold;
        }
        </style>
        """, unsafe_allow_html=True)

    client = None
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        # Lê o arquivo Excel
        data = pd.read_excel(uploaded_file, header=0)
        # Filtra as linhas onde a data na coluna 'C' é igual à data atual
        data['C'] = pd.to_datetime(data['C'])
        today = pd.to_datetime('today').normalize()
        filtered_data = data[data['C'] == today]
        st.write("Dados filtrados do arquivo Excel para a data atual:")
        st.dataframe(filtered_data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar ao Google Sheets"):
        client = authenticate_google_sheets()
        if client:
            st.success("Conectado com sucesso ao Google Sheets.")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if st.button("Enviar para Google Sheets") and uploaded_file is not None and client:
        insert_data_to_sheet(client, filtered_data, sheet_url)
        st.success("Dados inseridos com sucesso no Google Sheets.")

if __name__ == '__main__':
    main()

