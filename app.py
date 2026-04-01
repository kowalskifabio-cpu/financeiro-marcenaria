# STATUS DO SCRIPT: v14.0 (RESTAURAÇÃO TOTAL + CARGA BLINDADA) | DATA: 01/04/2026
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
import google.generativeai as genai

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Status Marcenaria - BI Financeiro", layout="wide")

# Configuração da IA - Kowalski
if "gemini_api_key" in st.secrets:
    genai.configure(api_key=st.secrets["gemini_api_key"])

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

@st.cache_resource(ttl=3600)
def abrir_planilha(key):
    for tentativa in range(4):
        try:
            if client:
                return client.open_by_key(key)
        except Exception as e:
            if "quota" in str(e).lower() or "429" in str(e) or "APIError" in str(e):
                time.sleep(10) # Pausa de segurança para o novo projeto
                continue
            st.error(f"Erro ao abrir a planilha (Cota do Google): {e}")
        time.sleep(2)
    return None

spreadsheet = abrir_planilha("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
if not spreadsheet: st.stop()

# --- FUNÇÕES DE LIMPEZA E FORMATAÇÃO (PRESERVAÇÃO TOTAL) ---
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
        p0, p1 = p[0].zfill(2), p[1]
        v = f"{p0}.{p1}0" if len(p1) == 1 else f"{p0}.{p1}"
            
    if nivel in [2, 3] and not v.startswith('0') and (len(v) == 1 or ('.' in v and len(v.split('.')[0]) == 1)):
        v = '0' + v
    return v

def formatar_moeda_br(val):
    if not isinstance(val, (int, float)): return val
    valor_abs = abs(val)
    f = f"{valor_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({f})" if val < 0 else f

def formatar_pct(val):
    if not isinstance(val, (int, float)): return val
    return f"{val:.1f}%"

def filtrar_linhas_zeradas(df, colunas_valores):
    df = df.copy()
    df['zerado'] = df[colunas_valores].abs().sum(axis=1) == 0
    remover_indices = set(df[(df['Nivel'] == 4) & (df['zerado'])].index)
    for idx, row in df[df['Nivel'] == 3].iterrows():
        pref = str(row['Conta']).strip() + "."
        filhos = df[(df['Nivel'] == 4) & (df['Conta'].str.startswith(pref))]
        if not filhos.empty and filhos['zerado'].all(): remover_indices.add(idx)
    for idx, row in df[df['Nivel'] == 2].iterrows():
        pref = str(row['Conta']).strip() + "."
        filhos_n3 = df[(df['Nivel'] == 3) & (df['Conta'].str.startswith(pref))]
        if not filhos_n3.empty and all(i in remover_indices for i in filhos_n3.index): remover_indices.add(idx)
    return df.drop(index=list(remover_indices)).drop(columns=['zerado'])

@st.cache_data(ttl=600) 
def listar_abas_existentes():
    try:
        return [w.title for w in spreadsheet.worksheets()]
    except:
        time.sleep(5)
        try: return [w.title for w in spreadsheet.worksheets()]
        except: return []
st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs(["📥 Carga", "📈 Relatório", "🎯 Indicadores", "🏢 Obras", "⚖️ Comparativo", "⚠️ Alertas", "📉 Curva ABC"])

with aba1:
    col_m, col_a = st.columns(2)
    meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    with col_m: m_ref = st.selectbox("Mês", meses_lista)
    with col_a: a_ref = st.selectbox("Ano", [2026, 2025, 2027, 2024])
    arq = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])
    
    if arq and st.button("🚀 Salvar Período"):
        df_upload = pd.read_excel(arq)
        df_upload.columns = [str(c).strip() for c in df_upload.columns]
        
        if 'Data Baixa' in df_upload.columns:
            df_upload['Data Baixa'] = pd.to_datetime(df_upload['Data Baixa'], errors='coerce')
            mes_num = meses_lista.index(m_ref) + 1
            ultimo_dia = calendar.monthrange(a_ref, mes_num)[1]
            data_inicio = datetime(a_ref, mes_num, 1)
            data_fim = datetime(a_ref, mes_num, ultimo_dia)
            fora_do_periodo = df_upload[(df_upload['Data Baixa'] < data_inicio) | (df_upload['Data Baixa'] > data_fim)]
            if not fora_do_periodo.empty:
                st.error(f"❌ CARGA ABORTADA: Datas fora de {m_ref}/{a_ref} detectadas.")
                st.stop()

        if 'Histórico' in df_upload.columns:
            total_antes = len(df_upload)
            df_upload = df_upload[~df_upload['Histórico'].astype(str).str.contains('baixa vinculo', case=False, na=False)]
            removidos = total_antes - len(df_upload)
            if removidos > 0:
                st.warning(f"ℹ️ {removidos} lançamentos de 'baixa vinculo' foram ignorados nesta carga.")

        # Validação da Base (Com cache para não estourar o novo projeto)
        @st.cache_data(ttl=3600)
        def carregar_base_val():
            try:
                ws_b = spreadsheet.worksheet("Base")
                return pd.DataFrame(ws_b.get_all_records())
            except:
                time.sleep(5)
                return pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())

        df_base_check = carregar_base_val()
        contas_base = set(df_base_check.iloc[:, 0].astype(str).str.strip().unique())
        df_upload['Conta_ID'] = df_upload['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        contas_carga = set(df_upload['Conta_ID'].unique())
        contas_faltantes = contas_carga - contas_base
        
        if contas_faltantes:
            st.error("⚠️ ERRO: Contas de Resultado novas detectadas. Cadastre na aba 'Base'.")
            st.write(list(contas_faltantes))
            st.stop()

        df_upload['Valor_Final'] = df_upload.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{m_ref}_{a_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            time.sleep(5) 
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        
        # Correção JSON Blindada: Conversão total para String nativa para o gspread
        dados_final = [df_upload.columns.tolist()] + df_upload.fillna('').astype(str).values.tolist()
        ws.update(dados_final)
        
        st.cache_data.clear()
        st.success(f"✅ Dados de {m_ref}/{a_ref} salvos! APP atualizado.")

# --- FILTROS SIDEBAR ---
st.sidebar.header("Filtros de Análise")
abas_existentes = listar_abas_existentes()
ano_sel = st.sidebar.selectbox("Ano de Referência", [2026, 2025, 2027, 2024], index=0)
ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

meses_disponiveis = [m for m in ordem_meses if f"{m}_{ano_sel}" in abas_existentes]
meses_sel = st.sidebar.multiselect("Meses (Filtro Geral)", meses_disponiveis, default=meses_disponiveis)

@st.cache_data(ttl=600)
def obter_centros_custo(abas_tuple): 
    centros = set()
    for aba_nome in abas_tuple:
        try:
            df_m = pd.DataFrame(spreadsheet.worksheet(aba_nome).get_all_records())
            if 'Centro de Custo' in df_m.columns:
                centros.update(df_m['Centro de Custo'].astype(str).unique())
        except: pass
    return sorted(list(centros))

lista_cc = obter_centros_custo(tuple(abas_existentes))
cc_sel = st.sidebar.multiselect("Centros de Custo", ["Todos"] + lista_cc, default="Todos")
niveis_sel = st.sidebar.multiselect("Níveis", [1, 2, 3, 4], default=[1, 2, 3, 4])
@st.cache_data(ttl=600)
def carregar_aba_base():
    try:
        return pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
    except:
        time.sleep(3)
        try: return pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
        except: return pd.DataFrame()

def processar_bi(ano, meses, filtros_cc):
    if not meses: return None, []
    df_base = carregar_aba_base().copy()
    if df_base.empty: return None, []
    df_base.columns = [str(c).strip() for c in df_base.columns]
    df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
    df_base['Conta'] = df_base.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)

    for m in meses:
        try:
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
        except: df_base[m] = 0.0

    df_base['ACUMULADO'] = df_base[meses].sum(axis=1)
    df_base['MÉDIA'] = df_base[meses].mean(axis=1)
    return df_base, meses

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
    ocultar_vazios_aba2 = st.checkbox("🚫 Ocultar Contas sem Movimento", value=False, key="ocultar_aba2")
    if st.button("📊 Gerar Relatório Filtrado", key="btn_aba2"):
        df_res, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_res is not None:
            if ocultar_vazios_aba2:
                df_res = filtrar_linhas_zeradas(df_res, meses_exibir + ['ACUMULADO'])
            df_visual = df_res[df_res['Nivel'].isin(niveis_sel)].copy()
            cols_export = ['Nivel', 'Conta', 'Descrição'] + meses_exibir + ['MÉDIA', 'ACUMULADO']
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_visual[cols_export].to_excel(writer, index=False, sheet_name='Consolidado')
            st.download_button(label="📥 Exportar Relatório (Excel)", data=buffer.getvalue(), file_name=f"Relatorio_{ano_sel}.xlsx")
            def style_rows(row):
                if row['Nivel'] == 1: return ['background-color: #334155; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)
            st.dataframe(df_visual[cols_export].style.apply(style_rows, axis=1).format({c: formatar_moeda_br for c in cols_export if c not in ['Nivel', 'Conta', 'Descrição']}), use_container_width=True, height=800)

with aba3:
    st.subheader("Indicadores de Gestão")
    if st.button("📈 Ver Dashboard", key="btn_aba3"):
        df_ind, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_ind is not None:
            rec = df_ind[df_ind['Conta'].str.startswith('01') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            desp = df_ind[df_ind['Conta'].str.startswith('02') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            lucro = rec + desp
            rent_val = (lucro/rec*100) if rec > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Faturamento", formatar_moeda_br(rec))
            c2.metric("Despesa", formatar_moeda_br(desp))
            c3.metric("Lucro Líquido", formatar_moeda_br(lucro), delta=f"{rent_val:.1f}% Rentabilidade")
            
            st.divider()
            df_chart = df_ind[(df_ind['Nivel'] == 2) & (df_ind['Conta'].isin(['01', '02']))].copy()
            df_melted = df_chart.melt(id_vars=['Descrição'], value_vars=meses_exibir, var_name='Mês', value_name='Valor')
            fig_evol = px.bar(df_melted, x='Mês', y=df_melted['Valor'].abs(), color='Descrição', barmode='group',
                            color_discrete_map={'RECEITAS': '#22c55e', 'DESPESAS': '#ef4444'}, text_auto='.2s', title="Evolução Mensal")
            df_lucro_line = df_ind[df_ind['Nivel'] == 1].melt(value_vars=meses_exibir, var_name='Mês', value_name='Lucro')
            fig_evol.add_trace(go.Scatter(x=df_lucro_line['Mês'], y=df_lucro_line['Lucro'], name='LUCRO LÍQUIDO', line=dict(color='#1e40af', width=3)))
            st.plotly_chart(fig_evol, use_container_width=True)

            col_top3, col_top4 = st.columns(2)
            with col_top3:
                st.write("### 📉 Maiores Grupos (Nível 3)")
                df_pizza3 = gerar_dados_pizza(df_ind, 3)
                fig_p3 = px.pie(df_pizza3, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
                st.plotly_chart(fig_p3, use_container_width=True)
            with col_top4:
                st.write("### 🔍 Maiores Detalhes (Nível 4)")
                df_pizza4 = gerar_dados_pizza(df_ind, 4)
                fig_p4 = px.pie(df_pizza4, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrRd)
                st.plotly_chart(fig_p4, use_container_width=True)

            st.divider()
            st.write("### 📊 Composição sobre Receita Líquida")
            df_perc = df_ind[df_ind['Nivel'] == 2].copy()
            df_perc['% s/ Receita'] = df_perc.apply(lambda x: (abs(x['ACUMULADO'])/rec*100) if rec > 0 else 0, axis=1)
            fig_bar_perc = px.bar(df_perc[df_perc['Conta'] != '01'], x='Descrição', y='% s/ Receita', text_auto='.1f', 
                                 color='Descrição', title="Peso das Despesas sobre a Receita Líquida (%)", color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_bar_perc, use_container_width=True)

with aba4:
    st.subheader("🏢 Análise de Obras e Rateio Dinâmico")
    
    @st.cache_data(ttl=300)
    def carregar_logica_rateio():
        try:
            df_log = pd.DataFrame(spreadsheet.worksheet("Rateio").get_all_records())
            df_log.iloc[:, 0] = df_log.iloc[:, 0].astype(str).str.lower().str.strip()
            return df_log
        except:
            st.warning("⚠️ Aba 'Rateio' não encontrada ou colunas inválidas.")
            return pd.DataFrame()

    df_rateio_config = carregar_logica_rateio()
    usar_rateio = st.toggle("🔄 Ativar Visão de Custo Real (Rateio Dinâmico)", value=False)
    
    if st.button("📊 Processar Obras", key="btn_aba4"):
        lista_dfs = []
        for aba_nome in [f"{m}_{a}" for a in [ano_sel] for m in meses_sel]:
            if aba_nome in abas_existentes:
                try:
                    df_m = pd.DataFrame(spreadsheet.worksheet(aba_nome).get_all_records())
                    if not df_m.empty: lista_dfs.append(df_m)
                except: pass
        
        if lista_dfs:
            df_all = pd.concat(lista_dfs, ignore_index=True)
            df_all['Valor_Final'] = pd.to_numeric(df_all['Valor_Final'], errors='coerce').fillna(0)
            res_cc_full = df_all.groupby('Centro de Custo').apply(lambda x: pd.Series({
                'Receitas': x[x['Conta_ID'].astype(str).str.startswith('01')]['Valor_Final'].sum(),
                'Despesa Direta': x[x['Conta_ID'].astype(str).str.startswith('02')]['Valor_Final'].sum(),
            })).reset_index()

            if usar_rateio and not df_rateio_config.empty:
                map_logica = dict(zip(df_rateio_config.iloc[:, 1], df_rateio_config.iloc[:, 0]))
                res_cc_full['Logica'] = res_cc_full['Centro de Custo'].map(map_logica).fillna('obra')
                bolo_rateio = res_cc_full[res_cc_full['Logica'] == 'rateio']['Despesa Direta'].sum()
                receptores_full = res_cc_full[res_cc_full['Logica'] == 'obra'].copy()
                total_desp_receptores = receptores_full['Despesa Direta'].sum()
                
                if abs(total_desp_receptores) > 0:
                    res_cc_full['Rateio Estrutura'] = 0.0
                    res_cc_full.loc[res_cc_full['Logica'] == 'obra', 'Rateio Estrutura'] = (res_cc_full['Despesa Direta'] / total_desp_receptores) * bolo_rateio
                else: res_cc_full['Rateio Estrutura'] = 0.0
                
                res_cc_final = res_cc_full[res_cc_full['Logica'] != 'rateio'].copy()
                res_cc_final['Resultado Real'] = res_cc_final['Receitas'] + res_cc_final['Despesa Direta'] + res_cc_final['Rateio Estrutura']
                cols_view = ['Centro de Custo', 'Receitas', 'Despesa Direta', 'Rateio Estrutura', 'Resultado Real']
            else:
                res_cc_final = res_cc_full.copy()
                res_cc_final['Resultado'] = res_cc_final['Receitas'] + res_cc_final['Despesa Direta']
                cols_view = ['Centro de Custo', 'Receitas', 'Despesa Direta', 'Resultado']

            st.dataframe(res_cc_final[cols_view].style.format({c: formatar_moeda_br for c in cols_view[1:]}), use_container_width=True)
with aba5:
    st.subheader("⚖️ Comparativo de Períodos")
    ocultar_aba5 = st.checkbox("🚫 Ocultar sem Movimento", value=False, key="ocultar_aba5")
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        aa = st.multiselect("Anos A", [2026, 2025, 2027, 2024], key="aa_c")
        ma = st.multiselect("Meses A", ordem_meses, default=ordem_meses, key="ma_c")
    with c_p2:
        ab = st.multiselect("Anos B", [2026, 2025, 2027, 2024], key="ab_c")
        mb = st.multiselect("Meses B", ordem_meses, default=ordem_meses, key="mb_c")
        
    if st.button("🔄 Comparar", key="btn_aba5"):
        df_base_c = carregar_aba_base().copy()
        if not df_base_c.empty:
            df_base_c.columns = [str(c).strip() for c in df_base_c.columns]
            df_base_c = df_base_c.rename(columns={df_base_c.columns[0]: 'Conta', df_base_c.columns[1]: 'Descrição', df_base_c.columns[2]: 'Nivel'})
            df_base_c['Conta'] = df_base_c.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)
            
            def calc_per(anos, meses):
                map_p = {}
                for aba in [f"{m}_{a}" for a in anos for m in meses]:
                    if aba in abas_existentes:
                        try:
                            df_m = pd.DataFrame(spreadsheet.worksheet(aba).get_all_records())
                            df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                            parciais = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                            for k,v in parciais.items(): map_p[k] = map_p.get(k,0)+v
                        except: pass
                return map_p
                
            m_a, m_b = calc_per(aa, ma), calc_per(ab, mb)
            df_base_c['PERÍODO A'] = df_base_c['Conta'].map(m_a).fillna(0)
            df_base_c['PERÍODO B'] = df_base_c['Conta'].map(m_b).fillna(0)
            
            for n in [3, 2, 1]:
                for idx, row in df_base_c[df_base_c['Nivel'] == n].iterrows():
                    pref = str(row['Conta']).strip() + "."
                    df_base_c.at[idx, 'PERÍODO A'] = df_base_c[(df_base_c['Nivel'] == 4) & (df_base_c['Conta'].str.startswith(pref))]['PERÍODO A'].sum()
                    df_base_c.at[idx, 'PERÍODO B'] = df_base_c[(df_base_c['Nivel'] == 4) & (df_base_c['Conta'].str.startswith(pref))]['PERÍODO B'].sum()
                    
            df_base_c['DIFERENÇA'] = df_base_c['PERÍODO B'] - df_base_c['PERÍODO A']
            df_base_c['VAR %'] = df_base_c.apply(lambda x: (x['DIFERENÇA']/abs(x['PERÍODO A'])*100) if x['PERÍODO A'] != 0 else 0, axis=1)
            
            if ocultar_aba5: df_base_c = filtrar_linhas_zeradas(df_base_c, ['PERÍODO A', 'PERÍODO B'])
            
            def style_comp(row):
                if row['Nivel'] == 1: return ['background-color: #334155; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)
            st.dataframe(df_base_c[['Nivel', 'Conta', 'Descrição', 'PERÍODO A', 'PERÍODO B', 'DIFERENÇA', 'VAR %']].style.apply(style_comp, axis=1).format({'PERÍODO A': formatar_moeda_br, 'PERÍODO B': formatar_moeda_br, 'DIFERENÇA': formatar_moeda_br, 'VAR %': formatar_pct}), use_container_width=True, height=700)

with aba6:
    st.subheader("⚠️ Central de Alertas Preventivos")
    if abas_existentes:
        abas_sort = sorted([a for a in abas_existentes if '_' in a], key=lambda x: (int(x.split('_')[1]), meses_lista.index(x.split('_')[0])), reverse=True)
        if len(abas_sort) >= 2:
            mes_atual_aba = abas_sort[0]
            meses_historico = abas_sort[1:4]
            st.write(f"**Analisando:** {mes_atual_aba} vs Média de ({', '.join(meses_historico)})")
            df_base_alert = carregar_aba_base().copy()
            if not df_base_alert.empty:
                df_base_alert.columns = [str(c).strip() for c in df_base_alert.columns]
                df_base_alert = df_base_alert.rename(columns={df_base_alert.columns[0]: 'Conta', df_base_alert.columns[1]: 'Descrição', df_base_alert.columns[2]: 'Nivel'})
                df_base_alert['Conta'] = df_base_alert.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)
                
                def get_vals(lista_abas):
                    map_v = {}
                    for a in lista_abas:
                        try:
                            df_m = pd.DataFrame(spreadsheet.worksheet(a).get_all_records())
                            df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                            parciais = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                            for k,v in parciais.items(): map_v[k] = map_v.get(k,0)+v
                        except: pass
                    return map_v
                
                v_at, v_hi = get_vals([mes_atual_aba]), get_vals(meses_historico)
                df_base_alert['Atual'] = df_base_alert['Conta'].map(v_at).fillna(0)
                df_base_alert['Media_Hist'] = df_base_alert['Conta'].map(v_hi).fillna(0) / len(meses_historico)
                
                alertas = df_base_alert[(df_base_alert['Nivel'] == 3) & (df_base_alert['Conta'].str.startswith('02'))].copy()
                alertas['Desvio'] = alertas['Atual'] - alertas['Media_Hist']
                estouros = alertas[alertas['Desvio'] < -100].sort_values(by='Desvio')
                
                if not estouros.empty:
                    for idx, row in estouros.iterrows():
                        with st.expander(f"🚨 Alerta: {row['Descrição']} - Estouro de {formatar_moeda_br(row['Desvio'])}"):
                            c1, c2, c3 = st.columns(3)
                            c1.metric("Gasto Atual", formatar_moeda_br(row['Atual']))
                            c2.metric("Média Histórica", formatar_moeda_br(row['Media_Hist']))
                            p_estouro = (abs(row['Atual'])/abs(row['Media_Hist'])-1)*100 if row['Media_Hist'] != 0 else 0
                            c3.metric("Aumento %", f"{p_estouro:.1f}%", delta_color="inverse")
                else: st.success("✅ Tudo sob controle nos alertas dinâmicos.")

with aba7:
    st.subheader("📉 Curva ABC de Despesas (Nível 4)")
    if st.button("🔍 Gerar Curva ABC", key="btn_aba7"):
        df_abc, _ = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_abc is not None:
            df_analitico = df_abc[(df_abc['Nivel'] == 4) & (df_abc['Conta'].str.startswith('02'))].copy()
            df_analitico['Valor_Abs'] = df_analitico['ACUMULADO'].abs()
            df_analitico = df_analitico[df_analitico['Valor_Abs'] > 0].sort_values(by='Valor_Abs', ascending=False)
            
            total_geral = df_analitico['Valor_Abs'].sum()
            if total_geral > 0:
                df_analitico['% Individual'] = (df_analitico['Valor_Abs'] / total_geral) * 100
                df_analitico['% Acumulado'] = df_analitico['% Individual'].cumsum()
                df_analitico['Classe'] = df_analitico['% Acumulado'].apply(lambda x: 'A' if x <= 80.1 else ('B' if x <= 95.1 else 'C'))
                
                c_a, c_b, c_c = st.columns(3)
                res_a, res_b, res_c = df_analitico[df_analitico['Classe'] == 'A'], df_analitico[df_analitico['Classe'] == 'B'], df_analitico[df_analitico['Classe'] == 'C']
                
                for col, classe, df_cl, color, bcolor in zip([c_a, c_b, c_c], ['A', 'B', 'C'], [res_a, res_b, res_c], ['#ef4444', '#f59e0b', '#22c55e'], ['#fee2e2', '#fef3c7', '#dcfce7']):
                    col.markdown(f"<div style='background-color: {bcolor}; padding: 20px; border-radius: 10px; border-left: 5px solid {color};'><h3 style='color: {color}; margin-top:0;'>CLASSE {classe}</h3><p style='font-size: 24px; font-weight: bold; margin-bottom:0;'>{formatar_moeda_br(-df_cl['Valor_Abs'].sum())}</p><p>{len(df_cl)} contas</p></div>", unsafe_allow_html=True)
                
                st.divider()
                fig_pareto = go.Figure()
                fig_pareto.add_trace(go.Bar(x=df_analitico['Descrição'], y=df_analitico['Valor_Abs'], name="Gasto Individual", marker_color='#334155'))
                fig_pareto.add_trace(go.Scatter(x=df_analitico['Descrição'], y=df_analitico['% Acumulado'], name="% Acumulado", yaxis="y2", line=dict(color="#ef4444", width=3)))
                fig_pareto.update_layout(title="Gráfico de Pareto", yaxis=dict(title="Valor R$"), yaxis2=dict(title="%", overlaying="y", side="right", range=[0, 105]), showlegend=False)
                st.plotly_chart(fig_pareto, use_container_width=True)
                
                with st.expander("🔥 Detalhamento ABC"):
                    st.dataframe(df_analitico[['Classe', 'Conta', 'Descrição', 'Valor_Abs', '% Individual', '% Acumulado']].style.format({'Valor_Abs': formatar_moeda_br, '% Individual': '{:.1f}%', '% Acumulado': '{:.1f}%'}), use_container_width=True)
