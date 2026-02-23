import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Status Marcenaria - Gestão Financeira", layout="wide")

# Estilos Visuais para o Dashboard
st.markdown("""
    <style>
    .stDataFrame { border-radius: 10px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    h1 { color: #1e40af; font-family: 'Helvetica'; }
    </style>
    """, unsafe_allow_html=True)

# --- CONEXÃO COM GOOGLE SHEETS ---
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_creds():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(info, scopes=scope)
    except Exception as e:
        st.error(f"Erro na chave de acesso: {e}")
        return None

creds = get_creds()
if not creds: st.stop()
client = gspread.authorize(creds)
spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")

st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Carga de Dados", "📈 Relatório Consolidado"])

# --- FUNÇÃO DE LIMPEZA DE CONTA (Resolve o erro do 2001 do Google) ---
def limpar_conta(valor):
    v = str(valor).strip()
    if '/' in v or '-' in v: # Se o Google converteu para data
        v = v.replace('/', '.').replace('-', '.')
        partes = v.split('.')
        if len(partes) >= 3:
            # Reconstrói 01.01.001 (ajusta se o final for 2001)
            ano_final = "001" if "2001" in partes[2] else partes[2][-3:]
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}.{ano_final}"
    return v

# --- ABA 1: CARGA ---
with aba1:
    col_m, col_a = st.columns(2)
    with col_m: m_ref = st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col_a: a_ref = st.selectbox("Ano", [2026, 2025, 2027])
    
    arq = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])
    
    if arq and st.button("🚀 Salvar Período"):
        df = pd.read_excel(arq)
        df.columns = [str(c).strip() for c in df.columns]
        df['Conta_ID'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{m_ref}_{a_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        
        ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
        st.success(f"✅ Dados de {nome_aba} salvos no Google Sheets!")

# --- ABA 2: RELATÓRIO ---
with aba2:
    ano_sel = st.sidebar.selectbox("Ano de Análise", [2026, 2025, 2027])
    
    if st.button("📊 Gerar Relatório de Níveis"):
        with st.spinner("Processando cálculos..."):
            # 1. Carrega a Base
            df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
            df_base.columns = [str(c).strip() for c in df_base.columns]
            df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
            df_base['Conta'] = df_base['Conta'].apply(limpar_conta)

            # 2. Identifica meses carregados
            abas = [w.title for w in spreadsheet.worksheets() if f"_{ano_sel}" in w.title]
            lista_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_exibir = [m for m in lista_meses if f"{m}_{ano_sel}" in abas]

            if not meses_exibir:
                st.warning(f"Sem dados carregados para o ano {ano_sel}.")
                st.stop()

            # 3. Processa cada mês
            for m in meses_exibir:
                df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano_sel}").get_all_records())
                df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                mapeamento = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                
                # Inicia valores no Nível 4
                df_base[m] = df_base['Conta'].map(mapeamento).fillna(0)

                # --- LÓGICA DE SOMATÓRIO HIERÁRQUICO ---
                # Nível 3: Soma seus Níveis 4
                for idx, row in df_base[df_base['Nivel'] == 3].iterrows():
                    prefixo = str(row['Conta']) + "."
                    df_base.at[idx, m] = df_base[(df_base['Nivel'] == 4) & (df_base['Conta'].str.startswith(prefixo))][m].sum()
                
                # Nível 2: Soma seus Níveis 3
                for idx, row in df_base[df_base['Nivel'] == 2].iterrows():
                    prefixo = str(row['Conta']) + "."
                    df_base.at[idx, m] = df_base[(df_base['Nivel'] == 3) & (df_base['Conta'].str.startswith(prefixo))][m].sum()
                
                # Nível 1 (Resultado): Soma todos os Níveis 2 (Receitas + Despesas negativas)
                for idx, row in df_base[df_base['Nivel'] == 1].iterrows():
                    df_base.at[idx, m] = df_base[df_base['Nivel'] == 2][m].sum()

            # 4. Cálculo de Totais e Média
            df_base['ACUMULADO'] = df_base[meses_exibir].sum(axis=1)
            df_base['MÉDIA'] = df_base[meses_exibir].mean(axis=1)

            # --- FORMATAÇÃO BRASILEIRA (Verde/Vermelho com Parênteses) ---
            def format_br_currency(val):
                if not isinstance(val, (int, float)): return val
                # Formato: 1.234,56
                f = f"{abs(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"({f})" if val < 0 else f

            def color_negative(val):
                if not isinstance(val, (int, float)): return ''
                color = '#e11d48' if val < 0 else '#16a34a' if val > 0 else '#6b7280'
                return f'color: {color}; font-weight: bold'

            def highlight_rows(row):
                if row['Nivel'] == 1: return ['background-color: #1e40af; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold'] * len(row)
                return [''] * len(row)

            # Exibição Final
            cols_fin = ['Nivel', 'Conta', 'Descrição', 'MÉDIA', 'ACUMULADO'] + meses_exibir
            st.dataframe(
                df_base[cols_fin].style.apply(highlight_rows, axis=1)
                .applymap(color_negative, subset=['MÉDIA', 'ACUMULADO'] + meses_exibir)
                .format({c: format_br_currency for c in cols_fin if c not in ['Nivel', 'Conta', 'Descrição']}),
                use_container_width=True, height=800
            )
