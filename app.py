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
    imagem = Image.open('bannerapp1.png')
    st.image (imagem, caption='')
     # CSS para personalizar os botões
    button_style = """
    <style>
    .stButton>button {
        color: white;
        background-color: #084d6e;
        font-size: 18px;
        font-weight: bold;
    }
    </style>
    """
    st.markdown(button_style, unsafe_allow_html=True)

    
    st.markdown("""
    <style>
        .reportview-container {
            margin-top: -2em;
        }
        #MainMenu {visibility: hidden;}
        .stDeployButton {display:none;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        #stDecoration {display:none;}
    </style>
    """, unsafe_allow_html=True)

    st.set_option('deprecation.showPyplotGlobalUse', False)
    st.markdown("""
        <style>
               .block-container {
                    padding-top: 0.2rem;
                    padding-bottom: 0rem;
                    padding-left: 0.3rem;
                    padding-right: 0.3rem;
    
                }
                
                #root > div:nth-child(1) > div > div > div > div > section > div {padding-bottom: 0rem;}
        </style>
        """, unsafe_allow_html=True)
    st.subheader(":blue[Envio dos dados coletados]")

    uploaded_file = st.file_uploader("SELECIONE ABAIXO O ARQUIVO", type=['xlsx', 'xls'])
    client, drive_service = authenticate_google_services()

    if uploaded_file:
        with open("temp_file.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        data = pd.read_excel("temp_file.xlsx", header=0)
        data['Date'] = pd.to_datetime(data['Date'])
        data.sort_values(by='Date', ascending=False, inplace=True)
        most_recent_date = data['Date'].max()
        filtered_data = data[data['Date'] == most_recent_date]
        st.write("Dados filtrados do arquivo Excel para a data mais recente:")
        st.dataframe(filtered_data)
    st.subheader('', divider='rainbow')
    col1, col2 = st.columns(2)
    with col1:
        connect_button = st.button("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CONECTAR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", key='connect_google_sheets')
    with col2:
        send_button = st.button("TRANSFERIR DADOS", key='send_to_google_sheets_drive')

    if connect_button:
        if client:
            st.success("Conectado com sucesso ao Google Sheets.")
        else:
            st.error("Falha ao conectar ao Google Sheets.")

    if send_button and uploaded_file and client:
        sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"
        folder_id = '14-m18deS8QWYsSGSMFHq3Iyy_vPzQmUa'
        if not filtered_data.empty:
            insert_data_to_sheet(client, filtered_data, sheet_url)
            file_id = upload_file_to_drive(drive_service, uploaded_file.name, "temp_file.xlsx", folder_id)
            st.success(f"Dados da data mais recente inseridos com sucesso no Google Sheets e backup realizado no Google Drive. ID do arquivo: {file_id}")
        else:
            st.error("Não há dados da data mais recente para enviar.")
    elif send_button and not uploaded_file:
        st.error("Por favor, faça o upload de um arquivo.")
    elif send_button and not client:
        st.error("Conexão com o Google Sheets não estabelecida. Tente conectar novamente.")

if __name__ == '__main__':
    main()


