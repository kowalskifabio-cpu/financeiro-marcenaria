import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="BI Status Marcenaria", layout="wide")

# Estilos Visuais
st.markdown("""
    <style>
    .stDataFrame { border-radius: 10px; }
    h1 { color: #1e40af; }
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
# Seu ID da Planilha
spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")

st.title("📊 BI Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Carga de Dados", "📈 Relatório Consolidado"])

# --- FUNÇÃO DE LIMPEZA DE CONTA (O segredo para o erro do 2001) ---
def limpar_conta(valor):
    v = str(valor).strip()
    # Se o Google converteu em data (ex: 01/01/2001)
    if '/' in v or '-' in v:
        v = v.replace('/', '.').replace('-', '.')
        partes = v.split('.')
        if len(partes) >= 3:
            # Reconstrói no formato 01.01.001
            dia, mes, ano = partes[0], partes[1], partes[2]
            final = "001" if "2001" in ano else ano[-3:]
            return f"{dia.zfill(2)}.{mes.zfill(2)}.{final}"
    return v

# --- ABA 1: CARGA ---
with aba1:
    col_m, col_a = st.columns(2)
    with col_m: m_ref = st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col_a: a_ref = st.selectbox("Ano", [2026, 2025, 2027])
    
    arq = st.file_uploader("Subir Excel", type=["xlsx"])
    
    if arq and st.button("🚀 Salvar Dados"):
        df = pd.read_excel(arq)
        df.columns = [str(c).strip() for c in df.columns]
        # Limpa a conta vindo do sistema
        df['Conta_ID'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{m_ref}_{a_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        
        ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
        st.success(f"✅ Aba {nome_aba} atualizada!")

# --- ABA 2: RELATÓRIO ---
with aba2:
    ano_sel = st.sidebar.selectbox("Ano do BI", [2026, 2025, 2027])
    
    if st.button("📊 Gerar Demonstrativo"):
        with st.spinner("Processando níveis financeiros..."):
            # 1. Base e Limpeza de Nomes
            df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
            df_base.columns = [str(c).strip() for c in df_base.columns]
            df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
            
            # Aplica a limpeza para corrigir o erro do 2001
            df_base['Conta'] = df_base['Conta'].apply(limpar_conta)

            # 2. Identifica Meses
            abas_disponiveis = [w.title for w in spreadsheet.worksheets() if f"_{ano_sel}" in w.title]
            lista_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_exibir = [m for m in lista_meses if f"{m}_{ano_sel}" in abas_disponiveis]

            if not meses_exibir:
                st.warning("Sem dados para este ano.")
                st.stop()

            # 3. Consolidação
            for m in meses_exibir:
                df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano_sel}").get_all_records())
                df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                soma_mes = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                df_base[m] = df_base['Conta'].map(soma_mes).fillna(0)

                # SOMA DOS NÍVEIS
                for n in [3, 2, 1]:
                    for idx in df_base[df_base['Nivel'] == n].index:
                        p = df_base.at[idx, 'Conta']
                        filhos = df_base[df_base['Conta'].str.startswith(p + ".")]
                        if not filhos.empty:
                            df_base.at[idx, m] = filhos[m].sum()

            # 4. Indicadores e Visual
            df_base['ACUMULADO'] = df_base[meses_exibir].sum(axis=1)
            df_base['MÉDIA'] = df_base[meses_exibir].mean(axis=1)

            def style_row(row):
                if row['Nivel'] == 1: return ['background-color: #1e40af; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #d1d5db; font-weight: bold'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #f3f4f6; font-weight: bold'] * len(row)
                return [''] * len(row)

            cols = ['Nivel', 'Conta', 'Descrição', 'MÉDIA', 'ACUMULADO'] + meses_exibir
            st.dataframe(
                df_base[cols].style.apply(style_row, axis=1)
                .format({c: "R$ {:,.2f}" for c in cols if c not in ['Nivel', 'Conta', 'Descrição']})
                .applymap(lambda x: 'color: red' if isinstance(x, (float, int)) and x < 0 else 'color: green' if isinstance(x, (float, int)) and x > 0 else '', 
                          subset=['MÉDIA', 'ACUMULADO'] + meses_exibir),
                use_container_width=True, height=800
            )
