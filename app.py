import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials
from streamlit_gsheets import GSheetsConnection
import os

def main():
    st.title("Upload de Arquivo Excel para Google Sheets")

    # Crie uma conexão com o Google Sheets
    gsheets_connector = GSheetsConnection(secrets=st.secrets)

    # Widget para fazer upload de arquivo Excel
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        # Lê e exibe os dados da planilha Excel
        data = pd.read_excel(uploaded_file, header=0)
        st.write("Dados lidos do arquivo Excel:")
        st.dataframe(data)

        # Botão para enviar dados para o Google Sheets
        if st.button("Enviar para Google Sheets"):
            try:
                # Abre a planilha e a aba especificada
                sheet = gsheets_connector.open_spreadsheet_by_key('1FPBeAXQBKy8noJ3bTF52p8JL_Eg-ptuSP6djDTsRfKE')
                worksheet = sheet.worksheet_by_title('Página1')
                
                # Converte os dados para lista de listas
                data_list = data.values.tolist()
                
                # Insere os dados na planilha (adiciona novas linhas)
                worksheet.append_rows(data_list, value_input_option='USER_ENTERED')
                st.success("Dados enviados com sucesso para o Google Sheets.")
            except Exception as e:
                st.error(f"Falha ao enviar dados: {e}")

if __name__ == '__main__':
    main()
