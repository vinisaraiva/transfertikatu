import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import datetime
import os

st.set_page_config(
    page_title="App TIKATU",
    page_icon="📑",
)

def authenticate_google_services():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    client = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return client, drive_service

def insert_data_to_sheet(client, df, sheet_url):
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    data_list = df.applymap(lambda x: str(x) if pd.notnull(x) else '').values.tolist()
    next_row = len(worksheet.get_all_values()) + 1
    worksheet.insert_rows(data_list, next_row)

def upload_file_to_drive(drive_service, filename, filepath, folder_id):
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(filepath, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file.get('id')

def main():
    st.title("Upload e Inserção de Arquivo Excel no Google Sheets e Backup no Google Drive")
    
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
    client, drive_service = None, None

    if uploaded_file:
        data = pd.read_excel(uploaded_file, header=0)
        data['Date'] = pd.to_datetime(data['Date'])
        data.sort_values(by='Date', ascending=False, inplace=True)
        most_recent_date = data['Date'].max()
        filtered_data = data[data['Date'] == most_recent_date]
        st.write("Dados filtrados do arquivo Excel para a data mais recente:")
        st.dataframe(filtered_data)

    cols = st.columns(2)
    with cols[0]:
        if st.button("Conectar ao Google Sheets", key='connect_google_sheets'):
            client, drive_service = authenticate_google_services()
            if client:
                st.success("Conectado com sucesso ao Google Sheets.")
            else:
                st.error("Falha ao conectar ao Google Sheets.")
    
    with cols[1]:
        if st.button("Enviar para Google Sheets", key='send_to_google_sheets') and uploaded_file and client:
            insert_data_to_sheet(client, filtered_data, sheet_url="https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0")
            st.success("Dados da data mais recente inseridos com sucesso no Google Sheets.")
    
    if uploaded_file and client and drive_service:
        folder_id = '14-m18deS8QWYsSGSMFHq3Iyy_vPzQmUa'
        if st.button("Fazer backup no Drive", key='backup_to_drive'):
            file_id = upload_file_to_drive(drive_service, uploaded_file.name, "temp_file.xlsx", folder_id)
            st.success(f"Backup realizado no Google Drive. ID do arquivo: {file_id}")

if __name__ == '__main__':
    main()




