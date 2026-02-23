import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Status Marcenaria - BI Financeiro", layout="wide")

# Estilização CSS para deixar o relatório mais limpo
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stDataFrame { background-color: white; border-radius: 10px; padding: 10px; }
    </style>
    """, unsafe_allow_html=True)

# --- CONEXÃO GOOGLE SHEETS ---
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_creds():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(info, scopes=scope)
    except Exception as e:
        st.error(f"⚠️ Erro nos Segredos do Streamlit: {e}")
        return None

creds = get_creds()
if creds:
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1qNqW6ybPR1Ge9TqJvB7hYJVLst8RDYce40ZEsMPoe4Q")
else:
    st.stop()

st.title("📊 BI Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Upload de Dados", "📈 Relatório Consolidado"])

# --- ABA 1: CARGA DE DADOS ---
with aba1:
    st.info("Utilize esta aba para subir o fechamento mensal. O sistema substituirá dados se o período já existir.")
    col_a, col_b = st.columns(2)
    with col_a:
        mes_carga = st.selectbox("Mês de Referência", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
    with col_b:
        ano_carga = st.selectbox("Ano de Referência", [2026, 2025, 2027], index=0)

    arquivo = st.file_uploader("Arraste o Excel do sistema aqui", type=["xlsx"])
    
    if arquivo and st.button("🚀 Processar e Salvar no Google"):
        with st.spinner("Limpando e enviando dados..."):
            df = pd.read_excel(arquivo)
            
            # Limpeza Crítica de Dados
            df.columns = [str(c).strip() for c in df.columns]
            # Extração do código da conta (ex: 01.01.001) - Trata como string sempre!
            df['Conta_ID'] = df['C. Resultado'].astype(str).str.split(' ').str[0].str.strip()
            # Conversão de Valor
            df['Valor Baixado'] = pd.to_numeric(df['Valor Baixado'], errors='coerce').fillna(0)
            # Regra de Sinal: P = Negativo, R = Positivo
            df['Valor_Final'] = df.apply(lambda x: x['Valor Baixado'] * -1 if str(x['Pag/Rec']).strip().upper() == 'P' else x['Valor Baixado'], axis=1)
            
            nome_aba = f"{mes_carga}_{ano_carga}"
            try:
                try:
                    ws = spreadsheet.worksheet(nome_aba)
                    ws.clear()
                except:
                    ws = spreadsheet.add_worksheet(title=nome_aba, rows="2000", cols="30")
                
                # Salva apenas colunas úteis para o relatório ser rápido
                colunas_finais = ['Conta_ID', 'Valor_Final', 'Data Baixa', 'Histórico', 'Pag/Rec']
                df_save = df[colunas_finais].astype(str)
                ws.update([df_save.columns.values.tolist()] + df_save.values.tolist())
                st.success(f"✅ Período {nome_aba} atualizado com sucesso!")
            except Exception as e:
                st.error(f"Falha na gravação: {e}")

# --- ABA 2: RELATÓRIO DE INDICADORES ---
with aba2:
    ano_filtro = st.sidebar.selectbox("Filtrar Ano", [2026, 2025, 2027])
    
    if st.button("🔄 Gerar Demonstrativo Completo"):
        with st.spinner("Consolidando meses e calculando níveis..."):
            # 1. Carrega a Base Estrutural
            try:
                base_ws = spreadsheet.worksheet("Base")
                df_report = pd.DataFrame(base_ws.get_all_records())
                df_report.columns = [c.strip() for c in df_report.columns]
                # Garante que 'Conta' na base seja string padronizada
                df_report['Conta'] = df_report['Conta'].astype(str).str.strip()
            except:
                st.error("Erro: Aba 'Base' não encontrada ou colunas incorretas.")
                st.stop()

            # 2. Localiza todas as abas do ano selecionado
            all_worksheets = spreadsheet.worksheets()
            abas_ano = [ws.title for ws in all_worksheets if f"_{ano_filtro}" in ws.title]
            
            if not abas_ano:
                st.warning(f"Nenhum dado encontrado para o ano {ano_filtro}.")
                st.stop()

            # 3. Consolida valores por mês
            ordem_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_encontrados = []

            for mes_ref in ordem_meses:
                nome_aba = f"{mes_ref}_{ano_filtro}"
                if nome_aba in abas_ano:
                    meses_encontrados.append(mes_ref)
                    mov_ws = spreadsheet.worksheet(nome_aba)
                    df_mov = pd.DataFrame(mov_ws.get_all_records())
                    df_mov['Valor_Final'] = pd.to_numeric(df_mov['Valor_Final'], errors='coerce').fillna(0)
                    
                    # Soma por conta
                    resumo = df_mov.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                    df_report[mes_ref] = df_report['Conta'].map(resumo).fillna(0)

            # 4. Cálculo Hierárquico (O coração do relatório)
            for m in meses_encontrados:
                # Soma de baixo para cima: Nível 4 -> 3 -> 2 -> 1
                for n in [3, 2, 1]:
                    indices = df_report[df_report['Nivel'] == n].index
                    for idx in indices:
                        prefixo = df_report.at[idx, 'Conta']
                        # Soma tudo que começa com o prefixo da conta pai
                        filhos = df_report[df_report['Conta'].str.startswith(prefixo + ".")]
                        if not filhos.empty:
                            df_report.at[idx, m] = filhos[m].sum()

            # 5. Adiciona Média e Acumulado
            df_report['TOTAL'] = df_report[meses_encontrados].sum(axis=1)
            df_report['MEDIA'] = df_report[meses_encontrados].mean(axis=1)

            # --- FORMATAÇÃO VISUAL ---
            def highlight_levels(row):
                if row['Nivel'] == 1: return ['background-color: #d1d5db; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #e5e7eb; font-weight: bold'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #f3f4f6'] * len(row)
                return [''] * len(row)

            # Formatação de Moeda e Cores para Valores
            format_dict = {m: "R$ {:,.2f}" for m in meses_encontrados}
            format_dict.update({"TOTAL": "R$ {:,.2f}", "MEDIA": "R$ {:,.2f}"})

            # Exibição Final
            st.subheader(f"Demonstrativo Consolidado - {ano_filtro}")
            
            df_final = df_report[['Nivel', 'Conta', 'Descrição ', 'MEDIA', 'TOTAL'] + meses_encontrados]
            
            st.dataframe(
                df_final.style.apply(highlight_levels, axis=1)
                .format(format_dict)
                .applymap(lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else 'color: green' if isinstance(x, (int, float)) and x > 0 else '', 
                          subset=meses_encontrados + ['TOTAL', 'MEDIA']),
                use_container_width=True,
                height=700
            )

            # Exportação
            csv = df_final.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Baixar Relatório (CSV)", csv, f"Relatorio_{ano_filtro}.csv", "text/csv")
