import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import os

st.set_page_config(
    page_title="App TIKATU",
    page_icon="📑",
)

def authenticate_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = ServiceAccountCredentials.from_json_keyfile_name('cacesso.json', SCOPES)
    return gspread.authorize(creds)

def insert_data_to_sheet(client, df, sheet_url):
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    data_list = df.applymap(lambda x: str(x) if pd.notnull(x) else '').values.tolist()
    next_row = len(worksheet.get_all_values()) + 1
    worksheet.insert_rows(data_list, next_row)

def main():
    imagem = Image.open('bannerapp1.png')
    st.image (imagem, caption='')

    st.title(":blue[App para envio de dados do monitoramento da água]")

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
                    padding-top: 0.3rem;
                    padding-bottom: 0rem;
                    padding-left: 0.3rem;
                    padding-right: 0.3rem;
    
                }
                
                
        </style>
        """, unsafe_allow_html=True)
    
    client = authenticate_google_sheets()
    st.subheader('Selecione abaixo o arquivo excel')
    uploaded_file = st.file_uploader("", type=['xlsx', 'xls'], key='file_uploader')
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)
    st.subheader('', divider='rainbow')
    sheet_url = "https://docs.google.com/spreadsheets/d/1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE/edit#gid=0"
    
    col1, col2 = st.columns(2)

    with col1:
        if st.button("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CONECTAR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", key='connect'):
            client = authenticate_google_sheets()
            if client:
                st.session_state['client'] = client
                st.session_state['connection_status'] = 'success'
            else:
                st.session_state['connection_status'] = 'error'
    
    with col2:
        if st.button("TRANSFERIR DADOS", key='send'):
            if 'connection_status' not in st.session_state or st.session_state['connection_status'] != 'success':
                st.error("Por favor, conecte-se ao Google Sheets primeiro.")
            elif uploaded_file is None:
                st.error("Por favor, selecione um arquivo para poder efetuar o envio.")
            elif 'client' in st.session_state:
                data = pd.read_excel(uploaded_file, header=0)
                insert_data_to_sheet(st.session_state['client'], data, sheet_url)
                st.session_state['insert_status'] = 'success'
    
    # Mensagens de status
    if 'connection_status' in st.session_state:
        if st.session_state['connection_status'] == 'success':
            st.success("Conectado com sucesso ao Google Sheets.")
        elif st.session_state['connection_status'] == 'error':
            st.error("Falha ao conectar ao Google Sheets.")
    
    if 'insert_status' in st.session_state and st.session_state['insert_status'] == 'success':
        st.success("Dados inseridos com sucesso no Google Sheets.")
        
if __name__ == '__main__':
    main()
