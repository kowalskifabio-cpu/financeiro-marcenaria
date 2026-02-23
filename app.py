import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Status Marcenaria", layout="wide")

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_creds():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(info, scopes=scope)
    except Exception as e:
        st.error(f"Erro nos Segredos (Secrets): {e}")
        return None

creds = get_creds()
if creds:
    client = gspread.authorize(creds)
    # ID da sua planilha fornecido
    spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
else:
    st.stop()

st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Carga de Dados", "📈 Relatório Modelo"])

# --- ABA 1: CARGA DE DADOS ---
with aba1:
    st.subheader("Upload do arquivo Janeiro 2026")
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox("Mês do Relatório", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col2:
        ano = st.selectbox("Ano do Relatório", [2026, 2027, 2025])

    arquivo = st.file_uploader("Selecione o arquivo Excel extraído do sistema", type=["xlsx"])
    
    if arquivo and st.button("🚀 Executar Carga"):
        with st.spinner("Processando..."):
            df = pd.read_excel(arquivo)
            
            # Padronização de Colunas (Evita erro de maiúsculas/minúsculas)
            df.columns = [c.strip() for c in df.columns]
            
            # Extrair apenas o número da conta (Ex: 01.01.001) da coluna C. Resultado
            df['Conta_Limpa'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
            
            # Tratamento de Valores
            df['Valor Baixado'] = pd.to_numeric(df['Valor Baixado'], errors='coerce').fillna(0)
            
            # Regra de Sinal: Receita (R) positivo, Pagamento (P) negativo
            df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
            
            # Gravação no Google Sheets
            nome_aba = f"{mes}_{ano}"
            try:
                try:
                    ws = spreadsheet.worksheet(nome_aba)
                    ws.clear()
                except:
                    ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="30")
                
                ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
                st.success(f"✅ Dados gravados com sucesso na aba: {nome_aba}")
            except Exception as e:
                st.error(f"Erro ao gravar no Google: {e}")

# --- ABA 2: RELATÓRIO DE INDICADORES ---
with aba2:
    st.subheader("Visualização por Níveis (Conforme Modelo)")
    
    if st.button("📊 Gerar Relatório Consolidado"):
        try:
            # 1. Carrega a Base (Hierarquia)
            base_ws = spreadsheet.worksheet("Base")
            df_base = pd.DataFrame(base_ws.get_all_records())
            # Limpa nomes de colunas da base
            df_base.columns = [c.strip() for c in df_base.columns]
            
            # 2. Carrega os dados do Mês
            nome_aba = f"{mes}_{ano}"
            mov_ws = spreadsheet.worksheet(nome_aba)
            df_mov = pd.DataFrame(mov_ws.get_all_records())
            
            # Converter valor para número
            df_mov['Valor_Final'] = pd.to_numeric(df_mov['Valor_Final'], errors='coerce').fillna(0)
            
            # 3. Cruzamento
            resumo_mes = df_mov.groupby('Conta_Limpa')['Valor_Final'].sum().to_dict()
            
            # 4. Cálculo dos Níveis (Lógica do Excel Modelo)
            df_base['Valor'] = df_base['Conta'].astype(str).str.strip().map(resumo_mes).fillna(0)
            
            # Somar de baixo para cima (Nível 4 até Nível 1)
            for nivel in [3, 2, 1]:
                contas_nivel = df_base[df_base['Nivel'] == nivel]
                for idx, row in contas_nivel.iterrows():
                    prefixo = str(row['Conta']).strip()
                    # Soma todos que começam com esse código (Ex: 01.01 soma tudo que começa com 01.01.)
                    filhos = df_base[df_base['Conta'].astype(str).str.startswith(prefixo + ".")]
                    if not filhos.empty:
                        df_base.at[idx, 'Valor'] = filhos['Valor'].sum()
            
            # 5. Exibição Estilizada
            def style_negative(v):
                color = 'red' if v < 0 else 'green' if v > 0 else 'black'
                return f'color: {color}; font-weight: bold' if v != 0 else f'color: {color}'

            st.write(f"### Resultado: {mes} / {ano}")
            
            # Formatação final para visualização
            df_final = df_base[['Nivel', 'Conta', 'Descrição', 'Valor']].copy()
            
            st.dataframe(
                df_final.style.applymap(style_negative, subset=['Valor'])
                .format({'Valor': 'R$ {:,.2f}'}),
                use_container_width=True,
                height=600
            )
            
        except Exception as e:
            st.error(f"Ocorreu um erro ao gerar o relatório: {e}")
            st.info("Dica: Verifique se você já fez o upload dos dados para este mês na Aba 1.")
