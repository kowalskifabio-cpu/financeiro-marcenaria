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

    arquivo = st.file_uploader("Arraste o Excel de Janeiro aqui", type=["xlsx"])
    
    if arquivo and st.button("🚀 Processar e Salvar"):
        df = pd.read_excel(arquivo)
        # Regra de sinal: P vira negativo
        df['Valor Ajustado'] = df.apply(lambda x: x['Valor Baixado'] * -1 if x['Pag/Rec'] == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{mes}_{ano}"
        try:
            worksheet = spreadsheet.worksheet(nome_aba)
            worksheet.clear()
        except:
            worksheet = spreadsheet.add_worksheet(title=nome_aba, rows="1000", cols="25")
        
        # Salva os dados processados
        df_save = df.astype(str)
        worksheet.update([df_save.columns.values.tolist()] + df_save.values.tolist())
        st.success(f"Dados de {nome_aba} salvos no Google Sheets!")

with aba2:
    st.subheader("Demonstrativo de Resultado (DRE)")
    if st.button("🔄 Gerar Relatório do Mês"):
        # Lendo a Base e os Dados do Mês
        base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
        try:
            mensal = pd.DataFrame(spreadsheet.worksheet(f"{mes}_{ano}").get_all_records())
            mensal['Valor Ajustado'] = pd.to_numeric(mensal['Valor Ajustado'])
            
            # Soma valores por conta (C. Resultado)
            resumo_mes = mensal.groupby('C. Resultado')['Valor Ajustado'].sum().to_dict()
            
            # Atribui valores ao Nível 4
            base['Valor'] = base['Conta'].map(resumo_mes).fillna(0)
            
            # Lógica de Soma dos Níveis (Cascata)
            # Organiza do nível mais profundo para o mais alto
            for n in [3, 2, 1]:
                indices_pai = base[base['Nivel'] == n].index
                for idx in indices_pai:
                    conta_pai = base.at[idx, 'Conta']
                    # Soma todos que começam com o código do pai
                    filhos = base[base['Conta'].str.startswith(str(conta_pai) + ".")]
                    base.at[idx, 'Valor'] = filhos['Valor'].sum()
            
            # Formatação para exibição
            base['Valor'] = base['Valor'].apply(lambda x: f"R$ {x:,.2f}")
            st.table(base[['Nivel', 'Conta', 'Descrição ', 'Valor']])
            
        except:
            st.warning(f"Aba {mes}_{ano} não encontrada. Faça o upload na Aba 1 primeiro.")
