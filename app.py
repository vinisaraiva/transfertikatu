import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

def main():
    st.title("Upload e Inserção de Arquivo Excel no Google Sheets")
    client = authenticate_google_sheets()  # Mova a autenticação para fora do botão
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Conectar ao Google Sheets"):
        if client:
            st.success("Conectado com sucesso ao Google Sheets.")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if st.button("Enviar para Google Sheets") and uploaded_file is not None:
        insert_data_to_sheet(client, data, sheet_url)

