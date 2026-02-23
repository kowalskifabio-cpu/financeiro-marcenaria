import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuração e Conexão
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_creds():
    info = dict(st.secrets["gcp_service_account"])
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(info, scopes=scope)

try:
    creds = get_creds()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
except Exception as e:
    st.error(f"Erro de conexão: {e}")
    st.stop()

st.set_page_config(page_title="Status Marcenaria", layout="wide")
st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Carga de Dados", "📈 Relatório de Níveis"])

with aba1:
    st.subheader("Upload de Movimentação Mensal")
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col2:
        ano = st.selectbox("Ano", [2025, 2026, 2027])

    arquivo = st.file_uploader("Arraste o Excel aqui", type=["xlsx"])
    
    if arquivo and st.button("🚀 Processar e Salvar"):
        with st.spinner("Lendo arquivo..."):
            df = pd.read_excel(arquivo)
            
            # Limpeza: Pega apenas o código numérico da conta (ex: 01.01.001)
            df['C. Resultado Limpo'] = df['C. Resultado'].astype(str).str.split(' ').str[0]
            
            # Converte valores para número
            df['Valor Baixado'] = pd.to_numeric(df['Valor Baixado'], errors='coerce').fillna(0)
            
            # Regra de sinal: P = Negativo
            df['Valor Ajustado'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).upper() == 'P' else x['Valor Baixado'], axis=1)
            
            nome_aba = f"{mes}_{ano}"
            try:
                worksheet = spreadsheet.worksheet(nome_aba)
                worksheet.clear()
            except:
                worksheet = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="30")
