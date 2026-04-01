# STATUS DO SCRIPT: v16.0 - OBRAS INDEPENDENTES + BLINDAGEM TOTAL | DATA: 01/04/2026
import google.generativeai as genai
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

@st.cache_resource(ttl=3600)
def abrir_planilha(key):
    for tentativa in range(3):
        try:
            if client:
                return client.open_by_key(key)
        except Exception as e:
            if "quota" in str(e).lower() or "429" in str(e):
                time.sleep(10)
                continue
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
        prefix = str(row['Conta']).strip() + "."
        filhos = df[(df['Nivel'] == 4) & (df['Conta'].str.startswith(prefix))]
        if not filhos.empty and filhos['zerado'].all(): remover_indices.add(idx)
    return df.drop(index=list(remover_indices)).drop(columns=['zerado'])

@st.cache_data(ttl=600) 
def listar_abas_existentes():
    try: return [w.title for w in spreadsheet.worksheets()]
    except: return []
@st.cache_data(ttl=300)
def carregar_logica_rateio():
    try:
        df_log = pd.DataFrame(spreadsheet.worksheet("Rateio").get_all_records())
        df_log.iloc[:, 0] = df_log.iloc[:, 0].astype(str).str.lower().str.strip()
        return df_log
    except:
        return pd.DataFrame()
st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs(["📥 Carga", "📈 Relatório", "🎯 Indicadores", "🏢 Obras", "⚖️ Comparativo", "⚠️ Alertas", "📉 Curva ABC"])

with aba1:
    col_m, col_a = st.columns(2)
    meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    with col_m: m_ref = st.selectbox("Mês", meses_lista)
    with col_a: a_ref = st.selectbox("Ano", [2026, 2025, 2027, 2024])
    arq = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])
    
    if arq and st.button("🚀 Salvar Período"):
        df_carga = pd.read_excel(arq)
        df_carga.columns = [str(c).strip() for c in df_carga.columns]
        
        if 'Data Baixa' in df_carga.columns:
            df_carga['Data Baixa'] = pd.to_datetime(df_carga['Data Baixa'], errors='coerce')
            mes_num = meses_lista.index(m_ref) + 1
            ultimo_dia = calendar.monthrange(a_ref, mes_num)[1]
            data_inicio = datetime(a_ref, mes_num, 1)
            data_fim = datetime(a_ref, mes_num, ultimo_dia)
            fora_do_periodo = df_carga[(df_carga['Data Baixa'] < data_inicio) | (df_carga['Data Baixa'] > data_fim)]
            if not fora_do_periodo.empty:
                st.error(f"❌ CARGA ABORTADA: Datas fora de {m_ref}/{a_ref} detectadas.")
                st.stop()

        if 'Histórico' in df_carga.columns:
            total_antes = len(df_carga)
            df_carga = df_carga[~df_carga['Histórico'].astype(str).str.contains('baixa vinculo', case=False, na=False)]
            removidos = total_antes - len(df_carga)
            if removidos > 0:
                st.warning(f"ℹ️ {removidos} lançamentos de 'baixa vinculo' foram ignorados nesta carga.")

        # Validação da Base
        ws_base_val = spreadsheet.worksheet("Base")
        df_base_check = pd.DataFrame(ws_base_val.get_all_records())
        contas_base = set(df_base_check.iloc[:, 0].astype(str).str.strip().unique())
        
        df_carga['Conta_ID'] = df_carga['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        contas_carga = set(df_carga['Conta_ID'].unique())
        contas_faltantes = contas_carga - contas_base
        
        if contas_faltantes:
            st.error("⚠️ ERRO: Contas de Resultado novas detectadas. Cadastre na aba 'Base'.")
            st.write(list(contas_faltantes))
            st.stop()

        df_carga['Valor_Final'] = df_carga.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{m_ref}_{a_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        
        # Blindagem JSON Python 3.13
        dados_upload = [df_carga.columns.tolist()] + df_carga.fillna('').astype(str).values.tolist()
        ws.update(dados_upload)
        
        st.cache_data.clear()
        st.success(f"✅ Dados de {m_ref}/{a_ref} salvos! APP atualizado.")

# --- FILTROS SIDEBAR (BI MENSAL) ---
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
        time.sleep(2)
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
            
            df_m['Conta_ID'] = df_m['Conta_ID'].astype(str).str.strip()
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
    st.subheader("🎯 Indicadores de Gestão")
    if st.button("📈 Ver Dashboard Completo", key="btn_aba3_completo"):
        df_ind, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_ind is not None:
            # 1. MÉTRICAS DE TOPO
            rec = df_ind[df_ind['Conta'].str.startswith('01') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            desp = df_ind[df_ind['Conta'].str.startswith('02') & (df_ind['Nivel'] == 2)]['ACUMULADO'].sum()
            lucro = rec + desp
            rent_val = (lucro/rec*100) if rec > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Faturamento Total", formatar_moeda_br(rec))
            c2.metric("Despesa Total", formatar_moeda_br(desp))
            c3.metric("Lucro Líquido", formatar_moeda_br(lucro), delta=f"{rent_val:.1f}% Rentabilidade")
            
            st.divider()
            
            # 2. EVOLUÇÃO MENSAL
            df_chart = df_ind[(df_ind['Nivel'] == 2) & (df_ind['Conta'].isin(['01', '02']))].copy()
            df_melted = df_chart.melt(id_vars=['Descrição'], value_vars=meses_exibir, var_name='Mês', value_name='Valor')
            fig_evol = px.bar(df_melted, x='Mês', y=df_melted['Valor'].abs(), color='Descrição', barmode='group',
                            color_discrete_map={'RECEITAS': '#22c55e', 'DESPESAS': '#ef4444'}, text_auto='.2s', title="Evolução Mensal (R$)")
            st.plotly_chart(fig_evol, use_container_width=True)

            st.divider()

            # 3. AS DUAS ROSCAS E OS TOP 10 DETALHADOS
            col_top_n3, col_top_n4 = st.columns(2)
            
            with col_top_n3:
                st.write("### 📉 Maiores Grupos (Nível 3)")
                df_pizza3 = gerar_dados_pizza(df_ind, 3) # Agrupa o resto em "Outros"
                fig_p3 = px.pie(df_pizza3, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
                st.plotly_chart(fig_p3, use_container_width=True)
                
                st.write("**Top 10 Gastos por Grupo:**")
                # Lista real dos 10 maiores sem o agrupamento "Outros" para conferência
                top10_n3 = df_ind[(df_ind['Nivel'] == 3) & (df_ind['ACUMULADO'] < 0)].copy()
                top10_n3['Gasto'] = top10_n3['ACUMULADO'].abs()
                top10_n3 = top10_n3.sort_values(by='Gasto', ascending=False).head(10)
                st.table(top10_n3[['Descrição', 'ACUMULADO']].rename(columns={'ACUMULADO': 'Valor'}).style.format({'Valor': formatar_moeda_br}))

            with col_top_n4:
                st.write("### 🔍 Maiores Detalhes (Nível 4)")
                df_pizza4 = gerar_dados_pizza(df_ind, 4)
                fig_p4 = px.pie(df_pizza4, values='Abs_Acumulado', names='Descrição', hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrRd)
                st.plotly_chart(fig_p4, use_container_width=True)
                
                st.write("**Top 10 Contas Analíticas:**")
                top10_n4 = df_ind[(df_ind['Nivel'] == 4) & (df_ind['ACUMULADO'] < 0)].copy()
                top10_n4['Gasto'] = top10_n4['ACUMULADO'].abs()
                top10_n4 = top10_n4.sort_values(by='Gasto', ascending=False).head(10)
                st.table(top10_n4[['Descrição', 'ACUMULADO']].rename(columns={'ACUMULADO': 'Valor'}).style.format({'Valor': formatar_moeda_br}))

            st.divider()

            # 4. TABELA DE COMPOSIÇÃO SOBRE RECEITA (ANÁLISE VERTICAL)
            st.write("### 📊 Composição das Despesas s/ Receita Líquida")
            df_perc = df_ind[df_ind['Nivel'] == 2].copy()
            df_perc['% s/ Receita'] = df_perc.apply(lambda x: (abs(x['ACUMULADO'])/rec*100) if rec > 0 else 0, axis=1)
            
            # Filtra apenas o que é despesa para a tabela de composição
            df_comp_view = df_perc[df_perc['Conta'] != '01'].sort_values(by='% s/ Receita', ascending=False)
            
            fig_bar_perc = px.bar(df_comp_view, x='Descrição', y='% s/ Receita', text_auto='.1f', 
                                 color='Descrição', title="Impacto das Despesas (%)", color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_bar_perc, use_container_width=True)
            
            st.write("**Detalhamento da Composição:**")
            st.dataframe(df_comp_view[['Descrição', 'ACUMULADO', '% s/ Receita']].style.format({'ACUMULADO': formatar_moeda_br, '% s/ Receita': '{:.1f}%'}), use_container_width=True)
with aba4:
    st.subheader("🏢 Análise de Obras e Rateio Dinâmico")
    
    # 1. FILTROS INDEPENDENTES (Aba Obras não olha para a Sidebar para datas)
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        anos_obras_sel = st.multiselect("Anos da Obra (Acumulado)", [2024, 2025, 2026, 2027], default=[ano_sel], key="anos_obra_v16")
    with col_f2:
        meses_obras_sel = st.multiselect("Meses da Obra (Acumulado)", meses_lista, default=meses_lista, key="meses_obra_v16")
    
    usar_rateio = st.toggle("🔄 Ativar Visão de Custo Real (Rateio Dinâmico)", value=False)
    
    if st.button("📊 Processar Obras Acumuladas", key="btn_aba4_v16"):
        lista_dfs = []
        # BUSCA GLOBAL: Pega TUDO dos meses/anos selecionados (Sem filtrar Centro de Custo ainda!)
        for a_obra in anos_obras_sel:
            for m_obra in meses_obras_sel:
                aba_nome = f"{m_obra}_{a_obra}"
                if aba_nome in abas_existentes:
                    try:
                        df_m = pd.DataFrame(spreadsheet.worksheet(aba_nome).get_all_records())
                        if not df_m.empty: 
                            lista_dfs.append(df_m)
                    except: pass
        
        if lista_dfs:
            # Consolida a base bruta
            df_all = pd.concat(lista_dfs, ignore_index=True)
            df_all['Valor_Final'] = pd.to_numeric(df_all['Valor_Final'], errors='coerce').fillna(0)
            df_all['Conta_ID'] = df_all['Conta_ID'].astype(str).str.strip()

            # 2. CÁLCULO DO RATEIO SOBRE A BASE TOTAL (Regra do 100%)
            res_cc_full = df_all.groupby('Centro de Custo').apply(lambda x: pd.Series({
                'Receitas': x[x['Conta_ID'].str.startswith('01')]['Valor_Final'].sum(),
                'Despesa Direta': x[x['Conta_ID'].str.startswith('02')]['Valor_Final'].sum(),
            })).reset_index()

            if usar_rateio:
                df_rateio_config = carregar_logica_rateio() 
                if not df_rateio_config.empty:
                    # Classifica quem é obra, quem é rateio
                    map_logica = dict(zip(df_rateio_config.iloc[:, 1], df_rateio_config.iloc[:, 0]))
                    res_cc_full['Logica'] = res_cc_full['Centro de Custo'].map(map_logica).fillna('obra')
                    
                    # Soma o bolo do rateio e a despesa das obras (Mesa Cheia)
                    bolo_rateio = res_cc_full[res_cc_full['Logica'] == 'rateio']['Despesa Direta'].sum()
                    receptores = res_cc_full[res_cc_full['Logica'] == 'obra'].copy()
                    total_desp_receptores = receptores['Despesa Direta'].sum()
                    
                    # Distribui o rateio proporcionalmente
                    if abs(total_desp_receptores) > 0:
                        res_cc_full['Rateio Estrutura'] = 0.0
                        res_cc_full.loc[res_cc_full['Logica'] == 'obra', 'Rateio Estrutura'] = (res_cc_full['Despesa Direta'] / total_desp_receptores) * bolo_rateio
                    else:
                        res_cc_full['Rateio Estrutura'] = 0.0
                    
                    res_cc_final = res_cc_full[res_cc_full['Logica'] != 'rateio'].copy()
                    res_cc_final['Resultado Real'] = res_cc_final['Receitas'] + res_cc_final['Despesa Direta'] + res_cc_final.get('Rateio Estrutura', 0)
                    cols_v = ['Centro de Custo', 'Receitas', 'Despesa Direta', 'Rateio Estrutura', 'Resultado Real']
                else: 
                    res_cc_final, cols_v = res_cc_full, ['Centro de Custo', 'Receitas', 'Despesa Direta']
            else:
                res_cc_final = res_cc_full.copy()
                res_cc_final['Resultado'] = res_cc_final['Receitas'] + res_cc_final['Despesa Direta']
                cols_v = ['Centro de Custo', 'Receitas', 'Despesa Direta', 'Resultado']

            # 3. FILTRO DE EXIBIÇÃO (SÓ AGORA!)
            # Se você selecionou obras específicas na sidebar, elas serão filtradas aqui após o cálculo
            if "Todos" not in cc_sel and cc_sel:
                res_cc_final = res_cc_final[res_cc_final['Centro de Custo'].isin(cc_sel)]

            # Formatação e Totais
            res_cc_final = res_cc_final.sort_values(by=cols_v[-1])
            somas = res_cc_final[cols_v[1:]].sum()
            linha_t = pd.DataFrame([['TOTAL CONSOLIDADO (FILTRADO)'] + somas.tolist()], columns=cols_v)
            res_cc_final = pd.concat([linha_t, res_cc_final], ignore_index=True)

            st.dataframe(res_cc_final[cols_v].style.format({c: formatar_moeda_br for c in cols_v[1:]}), use_container_width=True)
            
            buffer_cc = io.BytesIO()
            with pd.ExcelWriter(buffer_cc, engine='openpyxl') as writer: res_cc_final.to_excel(writer, index=False)
            st.download_button(label="📥 Exportar Obras (Excel)", data=buffer_cc.getvalue(), file_name="Obras_CustoReal.xlsx")
        else:
            st.warning("Sem dados para o período selecionado.")
# --- FUNÇÕES DE SUPORTE ADICIONAIS ---

with aba5:
    st.subheader("⚖️ Comparativo de Períodos Independente")
    ocultar_aba5 = st.checkbox("🚫 Ocultar sem Movimento", value=False, key="ocultar_aba5_v16_5")
    
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        aa = st.multiselect("Anos A", [2024, 2025, 2026, 2027], key="aa_v16_5")
        ma = st.multiselect("Meses A", meses_lista, default=meses_lista, key="ma_v16_5")
    with c_p2:
        ab = st.multiselect("Anos B", [2024, 2025, 2026, 2027], key="ab_v16_5")
        mb = st.multiselect("Meses B", meses_lista, default=meses_lista, key="mb_v16_5")
        
    if st.button("🔄 Executar Comparativo", key="btn_aba5_v16_5"):
        # Limpeza de cache para evitar o erro do "reboot"
        st.cache_data.clear()
        
        df_base_c = carregar_aba_base().copy()
        if not df_base_c.empty:
            df_base_c.columns = [str(c).strip() for c in df_base_c.columns]
            df_base_c = df_base_c.rename(columns={df_base_c.columns[0]: 'Conta', df_base_c.columns[1]: 'Descrição', df_base_c.columns[2]: 'Nivel'})
            
            # BLINDAGEM DE HIERARQUIA: Garante que 1.01 vire 01.01 para somar os filhos
            df_base_c['Conta'] = df_base_c.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str).str.strip()
            
            def calc_soberano(anos_alvo, meses_alvo):
                map_res = {}
                abas_desejadas = [f"{m}_{a}" for a in anos_alvo for m in meses_alvo]
                for aba_nome in abas_desejadas:
                    if aba_nome in abas_existentes:
                        try:
                            df_m = pd.DataFrame(spreadsheet.worksheet(aba_nome).get_all_records())
                            df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                            if "Todos" not in cc_sel and cc_sel:
                                if 'Centro de Custo' in df_m.columns:
                                    df_m = df_m[df_m['Centro de Custo'].isin(cc_sel)]
                            
                            # Mantendo a conta exata do Excel para casar com o Nível 4 da Base
                            df_m['ID_TEXTO'] = df_m['Conta_ID'].astype(str).str.strip()
                            somas = df_m.groupby('ID_TEXTO')['Valor_Final'].sum().to_dict()
                            for conta, valor in somas.items():
                                map_res[conta] = map_res.get(conta, 0) + valor
                        except: pass
                return map_res

            dados_a = calc_soberano(aa, ma)
            dados_b = calc_soberano(ab, mb)
            
            # 1. Primeiro carregamos os valores no Nível 4 (Analítico)
            df_base_c['PERÍODO A'] = df_base_c['Conta'].map(dados_a).fillna(0)
            df_base_c['PERÍODO B'] = df_base_c['Conta'].map(dados_b).fillna(0)
            
            # 2. Somamos os Níveis 3 e 2 (Baseados nos filhos analíticos)
            for n in [3, 2]:
                for idx, row in df_base_c[df_base_c['Nivel'] == n].iterrows():
                    pref = str(row['Conta']).strip() + "."
                    df_base_c.at[idx, 'PERÍODO A'] = df_base_c[(df_base_c['Nivel'] == 4) & (df_base_c['Conta'].str.startswith(pref))]['PERÍODO A'].sum()
                    df_base_c.at[idx, 'PERÍODO B'] = df_base_c[(df_base_c['Nivel'] == 4) & (df_base_c['Conta'].str.startswith(pref))]['PERÍODO B'].sum()

            # 3. CORREÇÃO DO NÍVEL 1: Ele soma os totais dos Níveis 2 (Receitas + Despesas)
            for idx, row in df_base_c[df_base_c['Nivel'] == 1].iterrows():
                df_base_c.at[idx, 'PERÍODO A'] = df_base_c[df_base_c['Nivel'] == 2]['PERÍODO A'].sum()
                df_base_c.at[idx, 'PERÍODO B'] = df_base_c[df_base_c['Nivel'] == 2]['PERÍODO B'].sum()

            df_base_c['DIFERENÇA'] = df_base_c['PERÍODO B'] - df_base_c['PERÍODO A']
            df_base_c['VAR %'] = df_base_c.apply(lambda x: (x['DIFERENÇA']/abs(x['PERÍODO A'])*100) if x['PERÍODO A'] != 0 else 0, axis=1)
            
            if ocultar_aba5:
                df_base_c = filtrar_linhas_zeradas(df_base_c, ['PERÍODO A', 'PERÍODO B'])
            
            def style_comp(row):
                if row['Nivel'] == 1: return ['background-color: #334155; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)
            
            st.dataframe(df_base_c[['Nivel', 'Conta', 'Descrição', 'PERÍODO A', 'PERÍODO B', 'DIFERENÇA', 'VAR %']].style.apply(style_comp, axis=1).format({
                'PERÍODO A': formatar_moeda_br, 
                'PERÍODO B': formatar_moeda_br, 
                'DIFERENÇA': formatar_moeda_br, 
                'VAR %': formatar_pct
            }), use_container_width=True, height=750)
with aba6:
    st.subheader("⚠️ Central de Alertas Preventivos")
    if abas_existentes:
        abas_sort = sorted([a for a in abas_existentes if '_' in a], key=lambda x: (int(x.split('_')[1]), meses_lista.index(x.split('_')[0])), reverse=True)
        if len(abas_sort) >= 2:
            st.write(f"**Analisando:** {abas_sort[0]} vs Média de ({', '.join(abas_sort[1:4])})")
            df_base_alert = carregar_aba_base().copy()
            df_base_alert.columns = [str(c).strip() for c in df_base_alert.columns]
            df_base_alert = df_base_alert.rename(columns={df_base_alert.columns[0]: 'Conta', df_base_alert.columns[1]: 'Descrição', df_base_alert.columns[2]: 'Nivel'})
            df_base_alert['Conta'] = df_base_alert.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)
            
            def get_vals_alert(lista):
                mv = {}
                for a in lista:
                    try:
                        df_m = pd.DataFrame(spreadsheet.worksheet(a).get_all_records())
                        df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                        px = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                        for k,v in px.items(): mv[str(k).strip()] = mv.get(str(k).strip(),0)+v
                    except: pass
                return mv
            
            v_at = get_vals_alert([abas_sort[0]])
            v_hi = get_vals_alert(abas_sort[1:4])
            df_base_alert['Atual'] = df_base_alert['Conta'].map(v_at).fillna(0)
            df_base_alert['Media'] = df_base_alert['Conta'].map(v_hi).fillna(0) / 3
            
            alertas = df_base_alert[(df_base_alert['Nivel'] == 3) & (df_base_alert['Conta'].str.startswith('02'))].copy()
            alertas['Desvio'] = alertas['Atual'] - alertas['Media']
            estouros = alertas[alertas['Desvio'] < -100].sort_values(by='Desvio')
            
            if not estouros.empty:
                for idx, row in estouros.iterrows():
                    with st.expander(f"🚨 {row['Descrição']} - Estouro de {formatar_moeda_br(row['Desvio'])}"):
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Atual", formatar_moeda_br(row['Atual']))
                        c2.metric("Média Histórica", formatar_moeda_br(row['Media']))
                        p = (abs(row['Atual'])/abs(row['Media'])-1)*100 if row['Media'] != 0 else 0
                        c3.metric("Aumento %", f"{p:.1f}%", delta_color="inverse")
            else: st.success("✅ Tudo sob controle.")

with aba7:
    st.subheader("📉 Curva ABC de Despesas (Nível 4)")
    if st.button("🔍 Gerar Curva ABC", key="btn_aba7_final"):
        df_abc, _ = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_abc is not None:
            df_an = df_abc[(df_abc['Nivel'] == 4) & (df_abc['Conta'].str.startswith('02'))].copy()
            df_an['Valor_Abs'] = df_an['ACUMULADO'].abs()
            df_an = df_an[df_an['Valor_Abs'] > 0].sort_values(by='Valor_Abs', ascending=False)
            tot = df_an['Valor_Abs'].sum()
            if tot > 0:
                df_an['% Individual'] = (df_an['Valor_Abs'] / tot) * 100
                df_an['% Acumulado'] = df_an['% Individual'].cumsum()
                df_an['Classe'] = df_an['% Acumulado'].apply(lambda x: 'A' if x <= 80.1 else ('B' if x <= 95.1 else 'C'))
                
                c_a, c_b, c_c = st.columns(3)
                r_a, r_b, r_c = df_an[df_an['Classe'] == 'A'], df_an[df_an['Classe'] == 'B'], df_an[df_an['Classe'] == 'C']
                
                for col, cl, dcl, color, bcolor in zip([c_a, c_b, c_c], ['A', 'B', 'C'], [r_a, r_b, r_c], ['#ef4444', '#f59e0b', '#22c55e'], ['#fee2e2', '#fef3c7', '#dcfce7']):
                    col.markdown(f"<div style='background-color: {bcolor}; padding: 20px; border-radius: 10px; border-left: 5px solid {color};'><h3 style='color: {color}; margin-top:0;'>CLASSE {cl}</h3><p style='font-size: 24px; font-weight: bold; margin-bottom:0;'>{formatar_moeda_br(-dcl['Valor_Abs'].sum())}</p><p>{len(dcl)} contas</p></div>", unsafe_allow_html=True)
                
                st.divider()
                fig_p = go.Figure()
                fig_p.add_trace(go.Bar(x=df_an['Descrição'], y=df_an['Valor_Abs'], name="Gasto", marker_color='#334155'))
                fig_p.add_trace(go.Scatter(x=df_an['Descrição'], y=df_an['% Acumulado'], name="%", yaxis="y2", line=dict(color="#ef4444", width=3)))
                fig_p.update_layout(title="Pareto", yaxis=dict(title="R$"), yaxis2=dict(title="%", overlaying="y", side="right", range=[0, 105]), showlegend=False)
                st.plotly_chart(fig_p, use_container_width=True)
                
                with st.expander("🔴 DETALHAR CLASSE A"):
                    st.dataframe(r_a[['Conta', 'Descrição', 'Valor_Abs', '% Individual', '% Acumulado']].style.format({'Valor_Abs': formatar_moeda_br, '% Individual': '{:.1f}%', '% Acumulado': '{:.1f}%'}), use_container_width=True)
                with st.expander("🟡 DETALHAR CLASSE B"):
                    st.dataframe(r_b[['Conta', 'Descrição', 'Valor_Abs', '% Individual', '% Acumulado']].style.format({'Valor_Abs': formatar_moeda_br, '% Individual': '{:.1f}%', '% Acumulado': '{:.1f}%'}), use_container_width=True)
                with st.expander("🟢 DETALHAR CLASSE C"):
                    st.dataframe(r_c[['Conta', 'Descrição', 'Valor_Abs', '% Individual', '% Acumulado']].style.format({'Valor_Abs': formatar_moeda_br, '% Individual': '{:.1f}%', '% Acumulado': '{:.1f}%'}), use_container_width=True)
