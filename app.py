import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def insert_data_to_sheet(df, sheet_url):
    client = authenticate_google_sheets()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    for index, row in df.iterrows():
        worksheet.append_row(row.values.tolist())

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")
    client = None

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])
    data = None
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar ao Google Sheets"):
        client = authenticate_google_sheets()
        if client:
            st.success("Conectado com sucesso ao Google Sheets!")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if st.button("Enviar para Google Sheets"):
        insert_data_to_sheet()
        st.success("Dados enviados com sucesso para o Google Sheets.")


if __name__ == '__main__':
    main()
