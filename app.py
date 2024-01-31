import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creden.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    client = gspread.authorize(creds)
    return client

def upload_data_to_sheet(uploaded_file, sheet_url):
    client = authenticate_google_sheets()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)

    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=1)
        data_list = data.values.tolist()
        worksheet.append_rows(data_list)

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])
    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"

    if st.button("Enviar para Google Sheets"):
        upload_data_to_sheet(uploaded_file, sheet_url)
        st.success("Dados enviados com sucesso para o Google Sheets.")

if __name__ == '__main__':
    main()
