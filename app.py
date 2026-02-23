import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO VISUAL ---
st.set_page_config(page_title="Status Marcenaria - Gestão Financeira", layout="wide")

# Estilo para deixar a tabela elegante
st.markdown("""
    <style>
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 5px; }
    h1 { color: #1e3a8a; }
    </style>
    """, unsafe_allow_html=True)

# --- CONEXÃO ---
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

st.title("📊 BI Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Importar Dados", "📈 Demonstrativo Consolidado"])

# --- ABA 1: CARGA ---
with aba1:
    col_mes, col_ano = st.columns(2)
    with col_mes: mes_ref = st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col_ano: ano_ref = st.selectbox("Ano", [2026, 2025, 2027])
    
    arquivo = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])
    
    if arquivo and st.button("Executar Carga"):
        df = pd.read_excel(arquivo)
        # Padroniza colunas do sistema
        df.columns = [str(c).strip() for c in df.columns]
        df['Conta_ID'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
        df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
        
        nome_aba = f"{mes_ref}_{ano_ref}"
        try:
            ws = spreadsheet.worksheet(nome_aba)
            ws.clear()
        except:
            ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="20")
        
        ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
        st.success(f"✅ Dados de {nome_aba} carregados!")

# --- ABA 2: RELATÓRIO ---
with aba2:
    ano_sel = st.sidebar.selectbox("Ano do Relatório", [2026, 2025, 2027])
    
    if st.button("🔄 Gerar Demonstrativo Completo"):
        with st.spinner("Calculando indicadores..."):
            # 1. Carrega e Limpa a Base
            df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
            df_base.columns = [str(c).strip() for c in df_base.columns]
            # Renomeia para garantir que o código funcione independente do espaço no nome original
            df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
            df_base['Conta'] = df_base['Conta'].astype(str).str.strip()

            # 2. Busca meses disponíveis
            abas = [w.title for w in spreadsheet.worksheets() if f"_{ano_sel}" in w.title]
            meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_ativos = [m for m in meses_lista if f"{m}_{ano_sel}" in abas]

            if not meses_ativos:
                st.warning(f"Nenhum dado encontrado para {ano_sel}")
                st.stop()

            # 3. Consolida valores
            for m in meses_ativos:
                df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano_sel}").get_all_records())
                df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                resumo = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                df_base[m] = df_base['Conta'].map(resumo).fillna(0)

                # SOMA HIERÁRQUICA (Níveis 4 -> 1)
                for n in [3, 2, 1]:
                    for idx in df_base[df_base['Nivel'] == n].index:
                        pref = df_base.at[idx, 'Conta']
                        filhos = df_base[df_base['Conta'].str.startswith(pref + ".")]
                        if not filhos.empty:
                            df_base.at[idx, m] = filhos[m].sum()

            # 4. Colunas de Indicadores
            df_base['ACUMULADO'] = df_base[meses_ativos].sum(axis=1)
            df_base['MÉDIA'] = df_base[meses_ativos].mean(axis=1)

            # 5. Formatação Visual
            def style_report(row):
                if row['Nivel'] == 1: return ['background-color: #1e3a8a; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #f1f5f9; font-weight: bold'] * len(row)
                return [''] * len(row)

            # Reordenar colunas conforme solicitado
            colunas_finais = ['Nivel', 'Conta', 'Descrição', 'MÉDIA', 'ACUMULADO'] + meses_ativos
            df_display = df_base[colunas_finais]

            st.dataframe(
                df_display.style.apply(style_report, axis=1)
                .format({col: "R$ {:,.2f}" for col in colunas_finais if col not in ['Nivel', 'Conta', 'Descrição']})
                .applymap(lambda x: 'color: #e11d48' if isinstance(x, (int, float)) and x < 0 else 'color: #16a34a' if isinstance(x, (int, float)) and x > 0 else '', 
                          subset=['MÉDIA', 'ACUMULADO'] + meses_ativos),
                use_container_width=True, height=800
            )

            # Exportação
            csv = df_final.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Baixar Relatório (CSV)", csv, f"Relatorio_{ano_filtro}.csv", "text/csv")
