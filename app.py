import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Status Marcenaria - Gestão Financeira", layout="wide")

# Estilos Visuais
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

# --- FUNÇÃO DE LIMPEZA DE CONTA (Xerife dos dados) ---
def limpar_conta_blindado(valor, nivel):
    v = str(valor).strip()
    # Corrige se o Google transformou em data (ex: 01/01/2001)
    if '/' in v or '-' in v:
        v = v.replace('/', '.').replace('-', '.')
        partes = v.split('.')
        if len(partes) >= 3:
            # Se for Nivel 4, garante o final 001 em vez de 2001
            p3 = "001" if partes[2] == "2001" else partes[2].zfill(3)
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}.{p3}"
        elif nivel == 3:
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}"
    # Padronização de zeros à esquerda para contas puramente numéricas
    if nivel == 2 and len(v) == 1: return v.zfill(2)
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
        st.success(f"✅ Dados de {nome_aba} salvos!")

# --- ABA 2: RELATÓRIO ---
with aba2:
    ano_sel = st.sidebar.selectbox("Ano de Análise", [2026, 2025, 2027])
    
    if st.button("📊 Gerar Relatório de Níveis"):
        with st.spinner("Calculando somatórios (Nível 3 -> Nível 2)..."):
            # 1. Carrega a Base
            df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
            df_base.columns = [str(c).strip() for c in df_base.columns]
            df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
            # Limpeza crucial para evitar erros de data e falta de zeros
            df_base['Conta'] = df_base.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1)

            # 2. Identifica meses
            abas = [w.title for w in spreadsheet.worksheets() if f"_{ano_sel}" in w.title]
            meses_exibir = [m for m in ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"] if f"{m}_{ano_sel}" in abas]

            if not meses_exibir:
                st.warning("Sem dados para este ano.")
                st.stop()

            # 3. Processa cada mês
            for m in meses_exibir:
                df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano_sel}").get_all_records())
                df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                mapeamento = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                
                df_base[m] = 0.0
                # PASSO 1: Coloca valores reais nas contas de Nível 4
                mask_n4 = df_base['Nivel'] == 4
                df_base.loc[mask_n4, m] = df_base.loc[mask_n4, 'Conta'].map(mapeamento).fillna(0)

                # PASSO 2: SOMATÓRIO DO NÍVEL 3 (Soma os Níveis 4)
                for idx, row in df_base[df_base['Nivel'] == 3].iterrows():
                    prefixo = str(row['Conta']) + "."
                    soma_n4 = df_base[(df_base['Nivel'] == 4) & (df_base['Conta'].str.startswith(prefixo))][m].sum()
                    df_base.at[idx, m] = soma_n4
                
                # PASSO 3: SOMATÓRIO DO NÍVEL 2 (Soma os Níveis 3)
                for idx, row in df_base[df_base['Nivel'] == 2].iterrows():
                    prefixo = str(row['Conta']) + "."
                    soma_n3 = df_base[(df_base['Nivel'] == 3) & (df_base['Conta'].str.startswith(prefixo))][m].sum()
                    df_base.at[idx, m] = soma_n3
                
                # PASSO 4: Nível 1 (Resultado Final) - Soma das Receitas e Despesas (que já são negativas)
                for idx, row in df_base[df_base['Nivel'] == 1].iterrows():
                    df_base.at[idx, m] = df_base[df_base['Nivel'] == 2][m].sum()

            # 4. Cálculo de Totais e Média
            df_base['ACUMULADO'] = df_base[meses_exibir].sum(axis=1)
            df_base['MÉDIA'] = df_base[meses_exibir].mean(axis=1)

            # --- FORMATAÇÃO BRASILEIRA ---
            def format_br(val):
                if not isinstance(val, (int, float)): return val
                v_abs = abs(val)
                f = f"{v_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"({f})" if val < 0 else f

            def color_val(val):
                if not isinstance(val, (int, float)): return ''
                return 'color: #D10000' if val < 0 else 'color: #008000' if val > 0 else 'color: #6b7280'

            def style_rows(row):
                if row['Nivel'] == 1: return ['background-color: #1e40af; color: white; font-weight: bold'] * len(row) # Azul Escuro
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row) # Cinza
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row) # Azul Claro
                return [''] * len(row)

            cols = ['Nivel', 'Conta', 'Descrição', 'MÉDIA', 'ACUMULADO'] + meses_exibir
            st.dataframe(
                df_base[cols].style.apply(style_rows, axis=1)
                .applymap(color_val, subset=['MÉDIA', 'ACUMULADO'] + meses_exibir)
                .format({c: format_br for c in cols if c not in ['Nivel', 'Conta', 'Descrição']}),
                use_container_width=True, height=800
            )
