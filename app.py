import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

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
    if nivel == 2 and len(v) == 1: return v.zfill(2)
    return v

# --- FORMATAÇÃO BRASILEIRA ---
def formatar_moeda_br(val):
    if not isinstance(val, (int, float)): return val
    valor_abs = abs(val)
    f = f"{valor_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({f})" if val < 0 else f

st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2 = st.tabs(["📥 Carga de Dados", "📈 Relatório Consolidado"])

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

with aba2:
    ano_sel = st.sidebar.selectbox("Ano de Análise", [2026, 2025, 2027])
    
    if st.button("📊 Gerar Relatório de Níveis"):
        with st.spinner("Somando Níveis..."):
            # Carregar Base de Estrutura
            df_base = pd.DataFrame(spreadsheet.worksheet("Base").get_all_records())
            df_base.columns = [str(c).strip() for c in df_base.columns]
            df_base = df_base.rename(columns={df_base.columns[0]: 'Conta', df_base.columns[1]: 'Descrição', df_base.columns[2]: 'Nivel'})
            
            # Forçar Conta a ser String para não perder o "0" (ex: 01)
            df_base['Conta'] = df_base.apply(lambda x: limpar_conta_blindado(x['Conta'], x['Nivel']), axis=1).astype(str)

            abas = [w.title for w in spreadsheet.worksheets() if f"_{ano_sel}" in w.title]
            meses_exibir = [m for m in ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"] if f"{m}_{ano_sel}" in abas]

            if not meses_exibir:
                st.warning("Sem dados para este ano.")
                st.stop()

            for m in meses_exibir:
                df_m = pd.DataFrame(spreadsheet.worksheet(f"{m}_{ano_sel}").get_all_records())
                df_m['Valor_Final'] = pd.to_numeric(df_m['Valor_Final'], errors='coerce').fillna(0)
                
                # Agregação por ID de conta da carga
                mapeamento = df_m.groupby('Conta_ID')['Valor_Final'].sum().to_dict()
                df_base[m] = 0.0
                
                # PASSO 1: Nível 4 recebe os valores diretos
                df_base.loc[df_base['Nivel'] == 4, m] = df_base['Conta'].map(mapeamento).fillna(0)

                # CORREÇÃO DA LÓGICA DE SOMA (Hierarquia 4 -> 3 -> 2)
                for n in [3, 2]:
                    for idx, row in df_base[df_base['Nivel'] == n].iterrows():
                        # O segredo é que o filho deve começar com o código exato do pai + um ponto ou sequência
                        pref = str(row['Conta']).strip()
                        # Busca todos os níveis 4 (folhas) que pertencem a esta árvore
                        total = df_base[(df_base['Nivel'] == 4) & (df_base['Conta'].str.startswith(pref))][m].sum()
                        df_base.at[idx, m] = total
                
                # Passo 4: Nível 1 (Resultado Final) - Soma de tudo que é Nível 2
                for idx, row in df_base[df_base['Nivel'] == 1].iterrows():
                    # Resultado = Receitas + Despesas (considerando que despesas são negativas)
                    df_base.at[idx, m] = df_base[df_base['Nivel'] == 2][m].sum()

            # Cálculos Finais
            df_base['ACUMULADO'] = df_base[meses_exibir].sum(axis=1)
            df_base['MÉDIA'] = df_base[meses_exibir].mean(axis=1)

            # Estilização
            def style_rows(row):
                if row['Nivel'] == 1: return ['background-color: #1e40af; color: white; font-weight: bold'] * len(row)
                if row['Nivel'] == 2: return ['background-color: #cbd5e1; font-weight: bold; color: black'] * len(row)
                if row['Nivel'] == 3: return ['background-color: #D1EAFF; font-weight: bold; color: black'] * len(row)
                return [''] * len(row)

            cols = ['Nivel', 'Conta', 'Descrição', 'MÉDIA', 'ACUMULADO'] + meses_exibir
            st.dataframe(
                df_base[cols].style.apply(style_rows, axis=1)
                .format({c: formatar_moeda_br for c in cols if c not in ['Nivel', 'Conta', 'Descrição']})
                .applymap(lambda x: 'color: #D10000' if isinstance(x, (int, float)) and x < 0 else 'color: #008000' if isinstance(x, (int, float)) and x > 0 else '', subset=['MÉDIA', 'ACUMULADO'] + meses_exibir),
                use_container_width=True, height=800
            )
