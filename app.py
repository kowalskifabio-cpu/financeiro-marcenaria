import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Status Marcenaria - BI Financeiro", layout="wide")

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_creds():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(info, scopes=scope)
    except Exception as e:
        st.error(f"Erro na chave: {e}")
        return None

creds = get_creds()
if not creds: st.stop()
client = gspread.authorize(creds)
spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")

# --- FUNÇÃO DE LIMPEZA DE CONTA ---
def limpar_conta_blindado(valor, nivel):
    v = str(valor).strip()
    if '/' in v or '-' in v: 
        v = v.replace('/', '.').replace('-', '.')
        partes = v.split('.')
        if len(partes) >= 3:
            ano_corrigido = "001" if "2001" in partes[2] else partes[2][-3:]
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}.{ano_corrigido}"
    
    if nivel in [2, 3]:
        if not v.startswith('0') and (len(v) == 1 or ('.' in v and len(v.split('.')[0]) == 1)):
            v = '0' + v
    return v

# --- FORMATAÇÃO BRASILEIRA ---
def formatar_moeda_br(val):
    if not isinstance(val, (int, float)): return val
    valor_abs = abs(val)
    f = f"{valor_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({f})" if val < 0 else f

st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2, aba3 = st.tabs(["📥 Carga de Dados", "📈 Relatório Consolidado", "🎯 Indicadores"])

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
        st.success(f"✅ Dados salvos!")

ano_sel = st.sidebar.selectbox("Ano de Análise", [2026, 2025, 2027])

def processar_bi(ano):
    df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
    df_base.columns = [str(c).strip() for c in df_base.columns]
    df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
    df_base['Conta'] = df_base.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)

    ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    abas_existentes = [w.title for w in spreadsheet.worksheets()]
    meses_exibir = [m for m in ordem_meses if f"{m}_{ano}" in abas_existentes]

    if not meses_exibir: return None, []

    for m in meses_exibir:
        df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano}").get_all_records())
        df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
        mapeamento = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
        df_base[m] = 0.0
        df_base.loc[df_base['Nivel'] == 4, m] = df_base['Conta'].map(mapeamento).fillna(0)

        for n in [3, 2]:
            for idx, row in df_base[df_base['Nivel'] == n].iterrows():
                pref = str(row['Conta']).strip()
                total = df_base[(df_base['Nivel'] == 4) & (df_base['Conta'].str.startswith(pref))][m].sum()
                df_base.at[idx, m] = total
        
        for idx, row in df_base[df_base['Nivel'] == 1].iterrows():
            df_base.at[idx, m] = df_base[df_base['Nivel'] == 2][m].sum()

    df_base['ACUMULADO'] = df_base[meses_exibir].sum(axis=1)
    df_base['MÉDIA'] = df_base[meses_exibir].mean(axis=1)
    return df_base, meses_exibir

with aba2:
    # --- AJUSTE CSS: BARRA DE ROLAGEM NO TOPO ---
    st.markdown("""
        <style>
            .stDataFrame div[data-testid="stHorizontalScrollContainer"] {
                transform: rotateX(180deg);
            }
            .stDataFrame div[data-testid="stHorizontalScrollContainer"] > div {
                transform: rotateX(180deg);
            }
        </style>
    """, unsafe_allow_html=True)

    if st.button("📊 Gerar Relatório de Níveis"):
        with st.spinner("Somando tudo..."):
            df_res, meses = processar_bi(ano_sel)
            if df_res is not None:
                def style_rows(row):
                    if row['Nivel'] == 1: return ['background-color: #334155; color: white; font-weight: bold'] * len(row)
                    if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                    if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                    return [''] * len(row)

                cols = ['Nivel', 'Conta', 'Descrição'] + meses + ['MÉDIA', 'ACUMULADO']
                st.dataframe(df_res[cols].style.apply(style_rows, axis=1).format({c: formatar_moeda_br for c in cols if c not in ['Nivel', 'Conta', 'Descrição']}), use_container_width=True, height=800)
            else:
                st.warning("Sem dados para este ano.")

with aba3:
    st.subheader(f"Principais Números de {ano_sel}")
    if st.button("📈 Ver Indicadores"):
        df_ind, meses = processar_bi(ano_sel)
        if df_ind is not None:
            receita_total = df_ind[df_ind['Conta'].str.startswith('01') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            despesa_total = df_ind[df_ind['Conta'].str.startswith('02') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            lucro_ano = receita_total + despesa_total

            c1, c2, c3 = st.columns(3)
            c1.metric("Faturamento Total", formatar_moeda_br(receita_total))
            c2.metric("Despesa Total", formatar_moeda_br(despesa_total))
            c3.metric("Lucro Líquido", formatar_moeda_br(lucro_ano), delta=f"{(lucro_ano/receita_total*100):.1f}% de Margem" if receita_total > 0 else "0%")

            st.divider()
            st.write("### Evolução Mensal")
            df_chart = df_ind[(df_ind['Nivel'] == 2) & (df_ind['Conta'].isin(['01', '02']))].copy()
            df_melted = df_chart.melt(id_vars=['Descrição'], value_vars=meses, var_name='Mês', value_name='Valor')
            df_melted['Valor_Abs'] = df_melted['Valor'].abs()
            
            fig = px.bar(df_melted, 
                         x='Mês', 
                         y='Valor_Abs', 
                         color='Descrição', 
                         barmode='group',
                         labels={'Valor_Abs': 'Valor (R$)', 'Descrição': 'Tipo'},
                         color_discrete_map={'RECEITAS': '#22c55e', 'DESPESAS': '#ef4444'},
                         text_auto='.2s')

            fig.update_layout(xaxis_title="Meses", yaxis_title="Valor em Reais", legend_title="Legenda")
            st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.write("### Maiores Despesas (Top 10)")
            maiores_despesas = df_ind[(df_ind['Nivel'] == 4) & (df_ind['ACUMULADO'] < 0)].sort_values(by='ACUMULADO')
            st.table(maiores_despesas[['Conta', 'Descrição', 'ACUMULADO']].head(10).style.format({'ACUMULADO': formatar_moeda_br}))
