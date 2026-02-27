import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
import plotly.graph_objects as go
import io 
import time
from datetime import datetime
import calendar

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Status Marcenaria - BI Financeiro", layout="wide")

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def get_gspread_client():
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Chave 'gcp_service_account' não encontrada nos Secrets.")
            return None
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Erro ao autorizar Google: {e}")
        return None

client = get_gspread_client()

@st.cache_resource
def abrir_planilha(key):
    try:
        return client.open_by_key(key)
    except Exception as e:
        st.error(f"Erro ao abrir a planilha (Limite de cota do Google): {e}")
        return None

spreadsheet = abrir_planilha("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
if not spreadsheet: st.stop()

# --- FUNÇÃO DE LIMPEZA DE CONTA (PRESERVAÇÃO DO .10) ---
def limpar_conta_blindado(valor, nivel):
    v = str(valor).strip()
    if '/' in v or '-' in v: 
        v = v.replace('/', '.').replace('-', '.')
        partes = v.split('.')
        if len(partes) >= 3:
            ano_corrigido = "001" if "2001" in partes[2] else partes[2][-3:]
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}.{ano_corrigido}"
    
    if nivel == 3 and '.' in v:
        p = v.split('.')
        p0 = p[0].zfill(2)
        p1 = p[1]
        if len(p1) == 1:
            v = f"{p0}.{p1}0"
        else:
            v = f"{p0}.{p1}"
            
    if nivel in [2, 3] and not v.startswith('0') and (len(v) == 1 or ('.' in v and len(v.split('.')[0]) == 1)):
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
    meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    with col_m: m_ref = st.selectbox("Mês", meses_lista)
    with col_a: a_ref = st.selectbox("Ano", [2026, 2025, 2027])
    arq = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])
    
    if arq and st.button("🚀 Salvar Período"):
        df = pd.read_excel(arq)
        df.columns = [str(c).strip() for c in df.columns]
        
        # --- NOVA TRAVA DE SEGURANÇA: VALIDAÇÃO DE DATAS ---
        if 'Data Baixa' in df.columns:
            df['Data Baixa'] = pd.to_datetime(df['Data Baixa'], errors='coerce')
            mes_num = meses_lista.index(m_ref) + 1
            ultimo_dia = calendar.monthrange(a_ref, mes_num)[1]
            data_inicio = datetime(a_ref, mes_num, 1)
            data_fim = datetime(a_ref, mes_num, ultimo_dia)
            
            fora_do_periodo = df[(df['Data Baixa'] < data_inicio) | (df['Data Baixa'] > data_fim)]
            
            if not fora_do_periodo.empty:
                st.error(f"❌ CARGA ABORTADA: O arquivo contém {len(fora_do_periodo)} linhas fora do período de {m_ref}/{a_ref}.")
                st.write("Datas inconsistentes encontradas:", fora_do_periodo['Data Baixa'].dt.strftime('%d/%m/%Y').unique())
                st.stop()
        else:
            st.warning("⚠️ Coluna 'Data Baixa' não encontrada para validação automática de período.")

        if 'Histórico' in df.columns:
            df = df[~df['Histórico'].astype(str).str.contains('baixa vinculo', case=False, na=False)]

        df['Conta_ID'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        
        df_base_check = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
        contas_base = set(df_base_check.iloc[:, 0].astype(str).str.strip().unique())
        contas_carga = set(df['Conta_ID'].unique())
        contas_faltantes = contas_carga - contas_base
        
        if contas_faltantes:
            st.error("⚠️ ERRO: Contas não encontradas na aba 'Base':")
            st.write(list(contas_faltantes))
            st.stop()
        
        df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{m_ref}_{a_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
        st.success(f"✅ Dados de {m_ref}/{a_ref} salvos com sucesso!")

# --- FILTROS SIDEBAR ---
st.sidebar.header("Filtros de Análise")
ano_sel = st.sidebar.selectbox("Ano", [2026, 2025, 2027])
ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

@st.cache_data(ttl=300)
def listar_abas_existentes():
    return [w.title for w in spreadsheet.worksheets()]

abas_existentes = listar_abas_existentes()
meses_disponiveis = [m for m in ordem_meses if f"{m}_{ano_sel}" in abas_existentes]
meses_sel = st.sidebar.multiselect("Meses", meses_disponiveis, default=meses_disponiveis)

@st.cache_data(ttl=600)
def obter_centros_custo(ano, meses):
    centros = set()
    for m in meses:
        try:
            df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano}").get_all_records())
            if 'Centro de Custo' in df_m.columns:
                centros.update(df_m['Centro de Custo'].astype(str).unique())
        except: pass
    return sorted(list(centros))

lista_cc = obter_centros_custo(ano_sel, meses_disponiveis)
cc_sel = st.sidebar.multiselect("Centros de Custo", ["Todos"] + lista_cc, default="Todos")
niveis_sel = st.sidebar.multiselect("Níveis", [1, 2, 3, 4], default=[1, 2, 3, 4])

# --- PROCESSAMENTO (GARANTE AS 237 LINHAS DA ABA BASE) ---
def processar_bi(ano, meses, filtros_cc):
    if not meses: return None, []
    df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
    df_base.columns = [str(c).strip() for c in df_base.columns]
    df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
    df_base['Conta'] = df_base.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)

    for m in meses:
        df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano}").get_all_records())
        df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
        if "Todos" not in filtros_cc and filtros_cc:
            if 'Centro de Custo' in df_m.columns:
                df_m = df_m[df_m['Centro de Custo'].isin(filtros_cc)]
        
        mapeamento = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
        df_base[m] = 0.0
        df_base.loc[df_base['Nivel'] == 4, m] = df_base['Conta'].map(mapeamento).fillna(0)

        for n in [3, 2]:
            for idx, row in df_base[df_base['Nivel'] == n].iterrows():
                pref = str(row['Conta']).strip() + "."
                total = df_base[(df_base['Nivel'] == 4) & (df_base['Conta'].str.startswith(pref))][m].sum()
                df_base.at[idx, m] = total
        
        for idx, row in df_base[df_base['Nivel'] == 1].iterrows():
            df_base.at[idx, m] = df_base[df_base['Nivel'] == 2][m].sum()

    df_base['ACUMULADO'] = df_base[meses].sum(axis=1)
    df_base['MÉDIA'] = df_base[meses].mean(axis=1)
    return df_base, meses

# --- FUNÇÃO AUXILIAR PARA PIZZA ---
def gerar_dados_pizza(df, nivel, limite=10):
    dados = df[(df['Nivel'] == nivel) & (df['ACUMULADO'] < 0)].copy()
    dados['Abs_Acumulado'] = dados['ACUMULADO'].abs()
    dados = dados.sort_values(by='Abs_Acumulado', ascending=False)
    if len(dados) > limite:
        principais = dados.head(limite).copy()
        outros_val = dados.iloc[limite:]['Abs_Acumulado'].sum()
        outros_df = pd.DataFrame({'Descrição': ['OUTRAS DESPESAS'], 'Abs_Acumulado': [outros_val]})
        return pd.concat([principais, outros_df], ignore_index=True)
    return dados

with aba2:
    st.markdown("""<style>.stDataFrame div[data-testid="stHorizontalScrollContainer"] { transform: rotateX(180deg); } .stDataFrame div[data-testid="stHorizontalScrollContainer"] > div { transform: rotateX(180deg); }</style>""", unsafe_allow_html=True)
    if st.button("📊 Gerar Relatório Filtrado"):
        df_res, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_res is not None:
            df_visual = df_res[df_res['Nivel'].isin(niveis_sel)].copy()
            cols_export = ['Nivel', 'Conta', 'Descrição'] + meses_exibir + ['MÉDIA', 'ACUMULADO']
            def style_rows(row):
                if row['Nivel'] == 1: return ['background-color: #334155; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)
            st.dataframe(df_visual[cols_export].style.apply(style_rows, axis=1).format({c: formatar_moeda_br for c in cols_export if c not in ['Nivel', 'Conta', 'Descrição']}), use_container_width=True, height=800)

with aba3:
    st.subheader("Indicadores")
    if st.button("📈 Ver Dashboard"):
        df_ind, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_ind is not None:
            rec = df_ind[df_ind['Conta'].str.startswith('01') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            desp = df_ind[df_ind['Conta'].str.startswith('02') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            lucro = rec + desp
            c1, c2, c3 = st.columns(3)
            c1.metric("Faturamento", formatar_moeda_br(rec))
            c2.metric("Despesa", formatar_moeda_br(desp))
            c3.metric("Lucro Líquido", formatar_moeda_br(lucro), delta=f"{(lucro/rec*100):.1f}%" if rec > 0 else "0%")
            
            df_chart = df_ind[(df_ind['Nivel'] == 2) & (df_ind['Conta'].isin(['01', '02']))].copy()
            df_melted = df_chart.melt(id_vars=['Descrição'], value_vars=meses_exibir, var_name='Mês', value_name='Valor')
            fig = px.bar(df_melted, x='Mês', y=df_melted['Valor'].abs(), color='Descrição', barmode='group',
                         color_discrete_map={'RECEITAS': '#22c55e', 'DESPESAS': '#ef4444'}, text_auto='.2s')
            df_lucro_line = df_ind[df_ind['Nivel'] == 1].melt(value_vars=meses_exibir, var_name='Mês', value_name='Lucro')
            fig.add_trace(go.Scatter(x=df_lucro_line['Mês'], y=df_lucro_line['Lucro'], name='LUCRO LÍQUIDO', line=dict(color='#1e40af', width=3)))
            st.plotly_chart(fig, use_container_width=True)

            st.divider()
            col_top3, col_top4 = st.columns(2)
            with col_top3:
                st.write("### 📉 Maiores Grupos (Nível 3)")
                df_pizza3 = gerar_dados_pizza(df_ind, 3)
                fig_p3 = px.pie(df_pizza3, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
                fig_p3.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_p3, use_container_width=True)
                st.table(df_ind[(df_ind['Nivel'] == 3) & (df_ind['ACUMULADO'] < 0)].sort_values(by='ACUMULADO').head(10)[['Conta', 'Descrição', 'ACUMULADO']].style.format({'ACUMULADO': formatar_moeda_br}))

            with col_top4:
                st.write("### 🔍 Maiores Detalhes (Nível 4)")
                df_pizza4 = gerar_dados_pizza(df_ind, 4)
                fig_p4 = px.pie(df_pizza4, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrRd)
                fig_p4.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_p4, use_container_width=True)
                st.table(df_ind[(df_ind['Nivel'] == 4) & (df_ind['ACUMULADO'] < 0)].sort_values(by='ACUMULADO').head(10)[['Conta', 'Descrição', 'ACUMULADO']].style.format({'ACUMULADO': formatar_moeda_br}))
