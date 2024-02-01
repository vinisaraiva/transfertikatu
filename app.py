import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

import streamlit as st
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Criar DataFrame de dados fictícios
def create_dummy_data():
    data = np.random.randint(0, 100, size=(10, 24))  # 10 linhas e 24 colunas (A a X)
    df = pd.DataFrame(data, columns=list('ABCDEFGHIJKLMNOPQRSTUVWX'))
    return df

# Função para autenticar e conectar ao Google Sheets
def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    client = gspread.authorize(creds)
    return client

# Função para inserir dados do DataFrame na planilha do Google Sheets
def insert_data_to_sheet(df, sheet_url):
    client = authenticate_google_sheets()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)  # assumindo que queremos inserir na primeira aba
    # Converte o DataFrame para uma lista de listas e insere na planilha
    worksheet.update('A1', df.values.tolist())

def main():
    st.title("Inserção de Dados Fictícios no Google Sheets")

    df_dummy = create_dummy_data()
    st.write("Dados fictícios gerados:")
    st.dataframe(df_dummy)

    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Inserir Dados no Google Sheets"):
        insert_data_to_sheet(df_dummy, sheet_url)
        st.success("Dados inseridos com sucesso no Google Sheets.")

if __name__ == '__main__':
    main()
