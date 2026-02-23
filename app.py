import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Configuração de Acesso
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Função para limpar a chave de caracteres invisíveis
def get_creds():
    info = dict(st.secrets["gcp_service_account"])
    # Remove barras invertidas duplicadas e garante que as quebras de linha sejam reais
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(info, scopes=scope)

try:
    creds = get_creds()
    client = gspread.authorize(creds)
    # Tenta abrir a planilha pelo ID que você passou
    spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
except Exception as e:
    st.error(f"Erro na conexão com Google: {e}")
    st.stop()

st.title("📊 Gestor Financeiro - Status Marcenaria")
# ... resto do código (seleção de mês, ano e upload)
