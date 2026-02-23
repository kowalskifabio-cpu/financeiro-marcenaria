import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 1. Configuração de Acesso ao Google Sheets
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
client = gspread.authorize(creds)

# Abre a planilha pelo ID que você enviou
spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")

st.title("📊 Gestor Financeiro - Status Marcenaria")

# Seleção de Período
col1, col2 = st.columns(2)
with col1:
    mes = st.selectbox("Selecione o Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
with col2:
    ano = st.selectbox("Selecione o Ano", [2025, 2026, 2027])

nome_aba = f"{mes}_{ano}"

# Upload do Arquivo
arquivo = st.file_uploader("Arraste o relatório de fechamento aqui (Excel)", type=["xlsx"])

if arquivo:
    if st.button(f"🚀 Carregar Dados para {nome_aba}"):
        with st.spinner("Processando e enviando para o Google Sheets..."):
            # Lendo o Excel
            df = pd.read_excel(arquivo)
            
            # Aplicando sua regra: P = Negativo, R = Positivo
            # Coluna B (Pag/Rec) e Coluna J (Valor Baixado)
            df['Valor Ajustado'] = df.apply(lambda x: x['Valor Baixado'] * -1 if x['Pag/Rec'] == 'P' else x['Valor Baixado'], axis=1)
            
            # Organizando para o Google Sheets (convertendo datas para texto)
            df = df.astype(str)
            
            # Verificando se a aba já existe para limpar ou criar
            try:
                worksheet = spreadsheet.worksheet(nome_aba)
                worksheet.clear() # Limpa se já existir (evita duplicidade)
            except:
                worksheet = spreadsheet.add_worksheet(title=nome_aba, rows="1000", cols="20")
            
            # Envia os novos dados
            worksheet.update([df.columns.values.tolist()] + df.values.tolist())
            
            st.success(f"✅ Dados de {nome_aba} carregados com sucesso!")
            st.info("O próximo passo será criar a visualização dos indicadores baseada na aba 'Base'.")
