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

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        # Converte a coluna 'Date' para datetime
        data['Date'] = pd.to_datetime(data['Date'])
        # Ordena o DataFrame pela coluna 'Date' em ordem decrescente
        data.sort_values(by='Date', ascending=False, inplace=True)
        # Identifica a data mais recente
        most_recent_date = data['Date'].max()
        # Filtra o DataFrame para incluir apenas linhas com a data mais recente
        filtered_data = data[data['Date'] == most_recent_date]
        st.write("Dados filtrados do arquivo Excel para a data mais recente:")
        st.dataframe(filtered_data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    client = authenticate_google_sheets()

    if st.button("Conectar ao Google Sheets", key='connect_google_sheets'):
        if client:
            st.success("Conectado com sucesso ao Google Sheets.")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if st.button("Enviar para Google Sheets", key='send_to_google_sheets') and uploaded_file is not None and client:
        if not filtered_data.empty:
            insert_data_to_sheet(client, filtered_data, sheet_url)
            st.success("Dados da data mais recente inseridos com sucesso no Google Sheets.")
        else:
            st.error("Não há dados da data mais recente para enviar.")

if __name__ == '__main__':
    main()





