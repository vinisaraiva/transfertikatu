import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

def create_dummy_data():
    data = np.random.randint(0, 100, size=(10, 24))
    df = pd.DataFrame(data, columns=list('ABCDEFGHIJKLMNOPQRSTUVWX'))
    return df

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('creden.json', SCOPES)
    client = gspread.authorize(creds)
    return client

def insert_data_to_sheet(df, sheet_url):
    client = authenticate_google_sheets()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    for index, row in df.iterrows():
        worksheet.append_row(row.values.tolist())

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

