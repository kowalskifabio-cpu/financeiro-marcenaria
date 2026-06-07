# STATUS DO SCRIPT: v17.0 - SUPABASE FULL | DATA: 2026-05-04
# App financeiro migrado para Supabase, sem dependência operacional de Google Sheets.

import io
import calendar
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from supabase import create_client

# =========================
# CONFIGURAÇÃO GERAL
# =========================
st.set_page_config(page_title="Status Marcenaria - BI Financeiro", layout="wide")

MESES_LISTA = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

MAPA_MESES = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}

MAPA_MESES_INV = {v: k for k, v in MAPA_MESES.items()}
ANOS_PADRAO = [2026, 2025, 2027, 2024]

# =========================
# CONEXÃO SUPABASE
# =========================
def mostrar_erro(contexto, erro):
    st.error(f"❌ {contexto}: {type(erro).__name__} - {erro}")

@st.cache_resource
def get_supabase_client():
    if "supabase" not in st.secrets:
        st.error("❌ Bloco [supabase] não encontrado nos Secrets do Streamlit.")
        st.stop()

    url = st.secrets["supabase"].get("url")
    key = st.secrets["supabase"].get("key")

    if not url or not key:
        st.error("❌ Secrets do Supabase incompletos. Informe url e key.")
        st.stop()

    return create_client(url, key)

supabase_client = get_supabase_client()

# =========================
# FUNÇÕES UTILITÁRIAS
# =========================
def limpar_conta_blindado(valor, nivel):
    v = str(valor).strip()

    if "/" in v or "-" in v:
        v = v.replace("/", ".").replace("-", ".")
        partes = v.split(".")
        if len(partes) >= 3:
            ano_corrigido = "001" if "2001" in partes[2] else partes[2][-3:]
            return f"{partes[1].zfill(2)}.{partes[0].zfill(2)}.{ano_corrigido}"

    if nivel == 3 and "." in v:
        p = v.split(".")
        if len(p) >= 2:
            p0, p1 = p[0].zfill(2), p[1]
            v = f"{p0}.{p1}0" if len(p1) == 1 else f"{p0}.{p1}"

    if nivel in [2, 3] and not v.startswith("0") and (len(v) == 1 or ("." in v and len(v.split(".")[0]) == 1)):
        v = "0" + v

    return v


def formatar_moeda_br(val):
    try:
        val = float(val)
    except Exception:
        return val

    valor_abs = abs(val)
    f = f"{valor_abs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({f})" if val < 0 else f


def formatar_pct(val):
    try:
        return f"{float(val):.1f}%"
    except Exception:
        return val


def filtrar_linhas_zeradas(df, colunas_valores):
    df = df.copy()
    colunas_validas = [c for c in colunas_valores if c in df.columns]

    if not colunas_validas:
        return df

    df["zerado"] = df[colunas_validas].abs().sum(axis=1) == 0
    remover_indices = set(df[(df["Nivel"] == 4) & (df["zerado"])].index)

    for idx, row in df[df["Nivel"] == 3].iterrows():
        prefix = str(row["Conta"]).strip() + "."
        filhos = df[(df["Nivel"] == 4) & (df["Conta"].astype(str).str.startswith(prefix))]
        if not filhos.empty and filhos["zerado"].all():
            remover_indices.add(idx)

    return df.drop(index=list(remover_indices)).drop(columns=["zerado"])


def supabase_fetch_all(table_name, columns="*"):
    """Busca todos os registros paginando de 1000 em 1000."""
    todos = []
    inicio = 0
    passo = 1000

    while True:
        resposta = (
            supabase_client
            .table(table_name)
            .select(columns)
            .range(inicio, inicio + passo - 1)
            .execute()
        )
        lote = resposta.data or []
        todos.extend(lote)

        if len(lote) < passo:
            break

        inicio += passo

    return todos


def normalizar_movimentos(df):
    """Converte nomes do Supabase para o padrão interno antigo do app."""
    if df.empty:
        return df

    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    rename_map = {
        "data": "Data",
        "ano": "Ano",
        "mes": "Mes",
        "conta_id": "Conta_ID",
        "centro_custo": "Centro de Custo",
        "valor": "Valor_Final",
    }
    df = df.rename(columns=rename_map)

    if "Valor_Final" in df.columns:
        df["Valor_Final"] = pd.to_numeric(df["Valor_Final"], errors="coerce").fillna(0.0)

    if "Conta_ID" in df.columns:
        df["Conta_ID"] = df["Conta_ID"].astype(str).str.strip()

    if "Centro de Custo" in df.columns:
        df["Centro de Custo"] = df["Centro de Custo"].astype(str).str.strip()

    if "Mes" in df.columns:
        df["Mes"] = pd.to_numeric(df["Mes"], errors="coerce").astype("Int64")

    if "Ano" in df.columns:
        df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce").astype("Int64")

    return df


def montar_abas_existentes_supabase(df_mov):
    if df_mov.empty or "Ano" not in df_mov.columns or "Mes" not in df_mov.columns:
        return []

    pares = df_mov[["Ano", "Mes"]].dropna().drop_duplicates()
    abas = []

    for _, row in pares.iterrows():
        ano = int(row["Ano"])
        mes = int(row["Mes"])
        if mes in MAPA_MESES_INV:
            abas.append(f"{MAPA_MESES_INV[mes]}_{ano}")

    return sorted(
        abas,
        key=lambda x: (int(x.split("_")[1]), MAPA_MESES[x.split("_")[0]])
    )

# =========================
# LEITURAS SUPABASE
# =========================
@st.cache_data(ttl=600)
def carregar_aba_base():
    try:
        dados = supabase_fetch_all("plano_contas")
        df = pd.DataFrame(dados)

        if df.empty:
            return pd.DataFrame()

        df.columns = [str(c).lower().strip() for c in df.columns]
        df = df.rename(columns={
            "conta_id": "Conta",
            "descricao": "Descrição",
            "nivel": "Nivel",
            "classificacao": "Classificacao"
        })

        colunas = ["Conta", "Descrição", "Nivel", "Classificacao"]
        faltantes = [c for c in colunas if c not in df.columns]
        if faltantes:
            st.error(f"❌ plano_contas sem colunas esperadas: {faltantes}")
            return pd.DataFrame()

        df = df[colunas].copy()
        df["Nivel"] = pd.to_numeric(df["Nivel"], errors="coerce")
        df = df.dropna(subset=["Conta", "Nivel"]).copy()
        df["Nivel"] = df["Nivel"].astype(int)
        df["Conta"] = df.apply(lambda x: limpar_conta_blindado(x["Conta"], x["Nivel"]), axis=1).astype(str).str.strip()
        def chave_ordem_conta(conta):
                    partes = str(conta).split(".")
                    return tuple(int(p) if p.isdigit() else 0 for p in partes)
        
        df["ordem_conta"] = df["Conta"].apply(chave_ordem_conta)
        df = df.sort_values(by="ordem_conta").drop(columns=["ordem_conta"]).reset_index(drop=True)
        df["Classificacao"] = (
            df["Classificacao"]
            .fillna("operacional")
            .astype(str)
            .str.lower()
            .str.strip()
        )

        return df

    except Exception as e:
        mostrar_erro("Erro ao ler plano_contas no Supabase", e)
        return pd.DataFrame()


@st.cache_data(ttl=600)
def carregar_logica_rateio():
    try:
        dados = supabase_fetch_all("rateio_config")
        df = pd.DataFrame(dados)

        if df.empty:
            return pd.DataFrame()

        df.columns = [str(c).lower().strip() for c in df.columns]
        df = df.rename(columns={
            "centro_custo": "Centro de Custo",
            "logica": "Logica"
        })

        colunas = ["Logica", "Centro de Custo"]
        faltantes = [c for c in colunas if c not in df.columns]
        if faltantes:
            st.error(f"❌ rateio_config sem colunas esperadas: {faltantes}")
            return pd.DataFrame()

        df = df[colunas].copy()
        df["Logica"] = df["Logica"].astype(str).str.lower().str.strip()
        df["Centro de Custo"] = df["Centro de Custo"].astype(str).str.strip()
        df = df[df["Logica"].isin(["obra", "rateio", "fora"])]
        df = df[df["Centro de Custo"] != ""]

        return df

    except Exception as e:
        mostrar_erro("Erro ao ler rateio_config no Supabase", e)
        return pd.DataFrame()


@st.cache_data(ttl=600)
def carregar_movimentos_periodo(ano, meses_numeros):
    """
    Lê movimentos do Supabase paginando de 1000 em 1000.
    Importante: o Supabase/PostgREST costuma limitar retorno por página.
    Sem paginação, o app lê só parte do mês e os totais ficam muito abaixo.
    """
    try:
        if not meses_numeros:
            return pd.DataFrame()

        dados = []
        passo = 1000

        for mes_num in meses_numeros:
            inicio = 0

            while True:
                resposta = (
                    supabase_client
                    .table("movimentos_financeiros")
                    .select("*")
                    .eq("ano", int(ano))
                    .eq("mes", str(int(mes_num)))
                    .range(inicio, inicio + passo - 1)
                    .execute()
                )

                lote = resposta.data or []
                dados.extend(lote)

                if len(lote) < passo:
                    break

                inicio += passo

        df = pd.DataFrame(dados)
        return normalizar_movimentos(df)

    except Exception as e:
        mostrar_erro("Erro ao ler movimentos_financeiros no Supabase", e)
        return pd.DataFrame()


@st.cache_data(ttl=600)
def carregar_aba_mensal(nome_aba):
    try:
        mes_nome, ano_txt = nome_aba.split("_")
        mes_num = MAPA_MESES[mes_nome]
        return carregar_movimentos_periodo(int(ano_txt), [mes_num])
    except Exception as e:
        mostrar_erro(f"Erro ao carregar período {nome_aba}", e)
        return pd.DataFrame()


@st.cache_data(ttl=600)
def carregar_todos_movimentos():
    try:
        dados = supabase_fetch_all("movimentos_financeiros")
        df = pd.DataFrame(dados)
        return normalizar_movimentos(df)
    except Exception as e:
        mostrar_erro("Erro ao ler todos os movimentos", e)
        return pd.DataFrame()


def obter_centros_custo(df_mov):
    if df_mov.empty or "Centro de Custo" not in df_mov.columns:
        return []
    return sorted([c for c in df_mov["Centro de Custo"].dropna().astype(str).unique().tolist() if c.strip()])
    
def cadastrar_centros_custo_automaticamente(df_mov_supabase):
    df_rateio = carregar_logica_rateio()

    centros_existentes = set()

    if not df_rateio.empty:
        centros_existentes = set(
            df_rateio["Centro de Custo"]
            .astype(str)
            .str.strip()
        )

    centros_importados = sorted(
        set(
            df_mov_supabase["centro_custo"]
            .astype(str)
            .str.strip()
        ) - centros_existentes
    )

    centros_importados = [
        c for c in centros_importados
        if c and c.lower() != "nan"
    ]

    if not centros_importados:
        return []

    novos_registros = [
        {
            "centro_custo": centro,
            "logica": "obra"
        }
        for centro in centros_importados
    ]

    tamanho_lote = 500

    for i in range(0, len(novos_registros), tamanho_lote):
        supabase_client.table("rateio_config").insert(
            novos_registros[i:i + tamanho_lote]
        ).execute()

    st.cache_data.clear()

    return centros_importados
# =========================
# PROCESSAMENTO PRINCIPAL
# =========================
def processar_bi(ano, meses, filtros_cc):
    if not meses:
        return None, []

    df_base = carregar_aba_base().copy()

    if df_base.empty:
        st.warning("Plano de contas não encontrado no Supabase.")
        return None, []

    meses_numeros = [MAPA_MESES[m] for m in meses if m in MAPA_MESES]
    df_mov = carregar_movimentos_periodo(ano, meses_numeros)

    for m in meses:
        df_base[m] = 0.0

    if not df_mov.empty:
        if "Todos" not in filtros_cc and filtros_cc:
            df_mov = df_mov[df_mov["Centro de Custo"].isin(filtros_cc)]

        for m in meses:
            mes_num = MAPA_MESES[m]
            df_m = df_mov[df_mov["Mes"] == mes_num].copy()

            if df_m.empty:
                continue

            df_m["Conta_ID"] = df_m["Conta_ID"].astype(str).str.strip()
            mapeamento = df_m.groupby("Conta_ID")["Valor_Final"].sum().to_dict()

            # Zera o mês
            df_base[m] = 0.0

            # 1) Primeiro joga valor exatamente no nível que existir
            df_base[m] = df_base["Conta"].map(mapeamento).fillna(0.0)

            # 2) Soma níveis superiores de baixo para cima: 5 -> 4 -> 3 -> 2
            niveis_existentes = sorted(df_base["Nivel"].dropna().unique(), reverse=True)

            for n in niveis_existentes:
                if n <= 1:
                    continue

                nivel_pai = n - 1

                for idx, row in df_base[df_base["Nivel"] == nivel_pai].iterrows():
                    pref = str(row["Conta"]).strip() + "."
                    total_filhos = df_base[
                        (df_base["Nivel"] == n) &
                        (df_base["Conta"].astype(str).str.startswith(pref))
                    ][m].sum()

                    if total_filhos != 0:
                        df_base.at[idx, m] = total_filhos

            # 3) Nível 1 soma os níveis 2
            for idx, _ in df_base[df_base["Nivel"] == 1].iterrows():
                df_base.at[idx, m] = df_base[df_base["Nivel"] == 2][m].sum()

    df_base["ACUMULADO"] = df_base[meses].sum(axis=1)
    df_base["MÉDIA"] = df_base[meses].mean(axis=1)

    return df_base, meses

def gerar_dados_pizza(df, nivel, limite=10):
    dados = df[(df["Nivel"] == nivel) & (df["ACUMULADO"] < 0)].copy()
    dados["Abs_Acumulado"] = dados["ACUMULADO"].abs()
    dados = dados.sort_values(by="Abs_Acumulado", ascending=False)

    if len(dados) > limite:
        principais = dados.head(limite).copy()
        outros_val = dados.iloc[limite:]["Abs_Acumulado"].sum()
        outros_df = pd.DataFrame({"Descrição": ["OUTRAS DESPESAS"], "Abs_Acumulado": [outros_val]})
        return pd.concat([principais, outros_df], ignore_index=True)

    return dados


def obter_movimentos_por_anos_meses(anos, meses):
    lista = []
    for ano in anos:
        meses_num = [MAPA_MESES[m] for m in meses if m in MAPA_MESES]
        df = carregar_movimentos_periodo(int(ano), meses_num)
        if not df.empty:
            lista.append(df)
    return pd.concat(lista, ignore_index=True) if lista else pd.DataFrame()

# =========================
# CARGA EXCEL -> SUPABASE
# =========================
def preparar_movimentos_para_supabase(df_carga, ano_ref, mes_ref_nome):
    df = df_carga.copy()
    df.columns = [str(c).strip() for c in df.columns]

    obrigatorias = ["Data Baixa", "Valor Baixado", "Pag/Rec", "C. Resultado", "Centro de Custo"]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas obrigatórias ausentes no Excel: {faltantes}")

    mes_num = MAPA_MESES[mes_ref_nome]
    data_inicio = datetime(int(ano_ref), mes_num, 1)
    data_fim = datetime(int(ano_ref), mes_num, calendar.monthrange(int(ano_ref), mes_num)[1])

    df["Data Baixa"] = pd.to_datetime(df["Data Baixa"], errors="coerce")
    fora = df[(df["Data Baixa"] < data_inicio) | (df["Data Baixa"] > data_fim) | (df["Data Baixa"].isna())]
    if not fora.empty:
        raise ValueError(f"Carga abortada: existem {len(fora)} linhas fora de {mes_ref_nome}/{ano_ref} pela Data Baixa.")

    if "Histórico" in df.columns:
        df = df[~df["Histórico"].astype(str).str.contains("baixa vinculo", case=False, na=False)].copy()

    df["valor"] = df.apply(
        lambda x: float(x["Valor Baixado"]) * -1 if str(x["Pag/Rec"]).strip().upper() == "P" else float(x["Valor Baixado"]),
        axis=1
    )

    df_out = pd.DataFrame({
        "data": df["Data Baixa"].dt.strftime("%Y-%m-%d"),
        "ano": int(ano_ref),
        "mes": str(mes_num),
        "conta_id": df["C. Resultado"].astype(str).str.split(" ").str[0].str.strip(),
        "centro_custo": df["Centro de Custo"].astype(str).str.strip(),
        "valor": df["valor"]
    })

    return df_out


def validar_importacao(df_mov_supabase):
    df_base = carregar_aba_base()
    df_rateio = carregar_logica_rateio()

    contas_base = set(df_base["Conta"].astype(str).str.strip()) if not df_base.empty else set()
    ccs_rateio = set(df_rateio["Centro de Custo"].astype(str).str.strip()) if not df_rateio.empty else set()

    contas_faltantes = sorted(set(df_mov_supabase["conta_id"].astype(str).str.strip()) - contas_base)
    centros_faltantes = sorted(set(df_mov_supabase["centro_custo"].astype(str).str.strip()) - ccs_rateio)

    return contas_faltantes, centros_faltantes


def inserir_movimentos_com_sobrescrita(df_mov_supabase, ano, mes_num):
    # Sobrescreve o mês importado.
    supabase_client.table("movimentos_financeiros").delete().eq("ano", int(ano)).eq("mes", str(int(mes_num))).execute()

    registros = df_mov_supabase.to_dict(orient="records")
    tamanho_lote = 500
    for i in range(0, len(registros), tamanho_lote):
        supabase_client.table("movimentos_financeiros").insert(registros[i:i + tamanho_lote]).execute()

# =========================
# INTERFACE
# =========================
st.title("📊 Gestor Financeiro - Status Marcenaria")

aba1, aba2, aba3, aba4, aba5, aba6, aba7, aba8, aba9, aba10, aba11 = st.tabs([
    "📥 Carga", "📈 Relatório", "🎯 Indicadores", "🏢 Obras", "⚖️ Comparativo",
    "⚠️ Alertas", "📉 Curva ABC", "🤖 Analista IA", "🧾 Composição da Obra",
    "⚙️ Configurações",
    "📊 Resultado Operacional"
])

# Sidebar baseada no Supabase
st.sidebar.header("Filtros de Análise")
df_mov_todos = carregar_todos_movimentos()
abas_existentes = montar_abas_existentes_supabase(df_mov_todos)

anos_disponiveis = sorted(df_mov_todos["Ano"].dropna().astype(int).unique().tolist(), reverse=True) if not df_mov_todos.empty and "Ano" in df_mov_todos.columns else ANOS_PADRAO
ano_sel = st.sidebar.selectbox("Ano de Referência", anos_disponiveis, index=0)

meses_disponiveis = [m for m in MESES_LISTA if f"{m}_{ano_sel}" in abas_existentes]
if not meses_disponiveis:
    meses_disponiveis = MESES_LISTA

meses_sel = st.sidebar.multiselect("Meses (Filtro Geral)", meses_disponiveis, default=meses_disponiveis)
lista_cc = obter_centros_custo(df_mov_todos)
cc_sel = st.sidebar.multiselect("Centros de Custo", ["Todos"] + lista_cc, default=["Todos"])
niveis_sel = st.sidebar.multiselect("Níveis", [1, 2, 3, 4], default=[1, 2, 3, 4])

with aba1:
    st.subheader("📥 Carga de Dados no Supabase")
    col_m, col_a = st.columns(2)

    with col_m:
        m_ref = st.selectbox("Mês", MESES_LISTA)
    with col_a:
        a_ref = st.selectbox("Ano", ANOS_PADRAO)

    arq = st.file_uploader("Subir Excel do Sistema", type=["xlsx"])

    if arq and st.button("🚀 Salvar Período no Supabase"):
        try:
            df_carga = pd.read_excel(arq)
            df_mov_import = preparar_movimentos_para_supabase(df_carga, a_ref, m_ref)

            contas_faltantes, centros_faltantes = validar_importacao(df_mov_import)
            centros_criados = cadastrar_centros_custo_automaticamente(df_mov_import)

            if centros_criados:
                st.info(
                    f"ℹ️ {len(centros_criados)} centros cadastrados automaticamente como OBRA."
                )
            
            # recarrega após inserir
            contas_faltantes, centros_faltantes = validar_importacao(df_mov_import)
            
            if contas_faltantes:
                st.warning(f"⚠️ Contas não encontradas no plano de contas: {len(contas_faltantes)}")
                st.dataframe(pd.DataFrame({"conta_id": contas_faltantes}), use_container_width=True)

            if centros_faltantes:
                st.warning(f"⚠️ Centros de custo sem configuração de rateio: {len(centros_faltantes)}")
                st.dataframe(pd.DataFrame({"centro_custo": centros_faltantes}), use_container_width=True)

            mes_num = MAPA_MESES[m_ref]
            inserir_movimentos_com_sobrescrita(df_mov_import, a_ref, mes_num)
            st.cache_data.clear()
            st.success(f"✅ {len(df_mov_import)} lançamentos de {m_ref}/{a_ref} gravados no Supabase com sobrescrita do mês.")

        except Exception as e:
            mostrar_erro("Erro na importação", e)

with aba2:
    st.markdown(
        """<style>.stDataFrame div[data-testid="stHorizontalScrollContainer"] { transform: rotateX(180deg); } .stDataFrame div[data-testid="stHorizontalScrollContainer"] > div { transform: rotateX(180deg); }</style>""",
        unsafe_allow_html=True
    )
    ocultar_vazios_aba2 = st.checkbox("🚫 Ocultar Contas sem Movimento", value=False, key="ocultar_aba2")

    if st.button("📊 Gerar Relatório Filtrado", key="btn_aba2"):
        df_res, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)

        if df_res is None:
            st.error("❌ Não foi possível gerar o relatório.")
        else:
            if ocultar_vazios_aba2:
                df_res = filtrar_linhas_zeradas(df_res, meses_exibir + ["ACUMULADO"])

            df_visual = df_res[df_res["Nivel"].isin(niveis_sel)].copy()
            cols_export = ["Nivel", "Conta", "Descrição"] + meses_exibir + ["MÉDIA", "ACUMULADO"]

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_visual[cols_export].to_excel(writer, index=False, sheet_name="Consolidado")

            st.download_button(
                label="📥 Exportar Relatório (Excel)",
                data=buffer.getvalue(),
                file_name=f"Relatorio_{ano_sel}.xlsx"
            )

            def style_rows(row):
                if row["Nivel"] == 1:
                    return ["background-color: #334155; color: white; font-weight: bold"] * len(row)
                if row["Nivel"] == 2:
                    return ["background-color: #cbd5e1; font-weight: bold; color: black"] * len(row)
                if row["Nivel"] == 3:
                    return ["background-color: #D1EAFF; font-weight: bold; color: black"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_visual[cols_export].style.apply(style_rows, axis=1).format({
                    c: formatar_moeda_br for c in cols_export if c not in ["Nivel", "Conta", "Descrição"]
                }),
                use_container_width=True,
                height=800
            )

with aba3:
    st.subheader("🎯 Indicadores de Gestão")

    if st.button("📈 Ver Dashboard Completo", key="btn_aba3_completo"):
        df_ind, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)

        if df_ind is not None:
            # 1. MÉTRICAS DE TOPO
            rec = df_ind[
                (df_ind["Conta"].astype(str).str.startswith("01")) &
                (df_ind["Nivel"] == 2)
            ]["ACUMULADO"].sum()

            desp = df_ind[
                (df_ind["Conta"].astype(str).str.startswith("02")) &
                (df_ind["Nivel"] == 2)
            ]["ACUMULADO"].sum()

            lucro = rec + desp
            rent_val = (lucro / rec * 100) if rec > 0 else 0

            c1, c2, c3 = st.columns(3)
            c1.metric("Faturamento Total", formatar_moeda_br(rec))
            c2.metric("Despesa Total", formatar_moeda_br(desp))
            c3.metric("Lucro Líquido", formatar_moeda_br(lucro), delta=f"{rent_val:.1f}% Rentabilidade")

            st.divider()

            # 2. EVOLUÇÃO MENSAL
            df_chart = df_ind[
                (df_ind["Nivel"] == 2) &
                (df_ind["Conta"].isin(["01", "02"]))
            ].copy()

            df_melted = df_chart.melt(
                id_vars=["Descrição"],
                value_vars=meses_exibir,
                var_name="Mês",
                value_name="Valor"
            )

            fig_evol = px.bar(
                df_melted,
                x="Mês",
                y=df_melted["Valor"].abs(),
                color="Descrição",
                barmode="group",
                color_discrete_map={
                    "RECEITAS": "#22c55e",
                    "DESPESAS": "#ef4444"
                },
                text_auto=".2s",
                title="Evolução Mensal (R$)"
            )

            st.plotly_chart(fig_evol, use_container_width=True)

            st.divider()

            # 3. ROSCAS + TOP 10
            col_top_n3, col_top_n4 = st.columns(2)

            with col_top_n3:
                st.write("### 📉 Maiores Grupos (Nível 3)")

                df_pizza3 = gerar_dados_pizza(df_ind, 3)

                fig_p3 = px.pie(
                    df_pizza3,
                    values="Abs_Acumulado",
                    names="Descrição",
                    hole=0.4,
                    color_discrete_sequence=px.colors.sequential.RdBu
                )

                st.plotly_chart(fig_p3, use_container_width=True)

                st.write("**Top 10 Gastos por Grupo:**")

                top10_n3 = df_ind[
                    (df_ind["Nivel"] == 3) &
                    (df_ind["ACUMULADO"] < 0)
                ].copy()

                top10_n3["Gasto"] = top10_n3["ACUMULADO"].abs()
                top10_n3 = top10_n3.sort_values(by="Gasto", ascending=False).head(10)

                st.table(
                    top10_n3[["Descrição", "ACUMULADO"]]
                    .rename(columns={"ACUMULADO": "Valor"})
                    .style.format({"Valor": formatar_moeda_br})
                )

            with col_top_n4:
                st.write("### 🔍 Maiores Detalhes (Nível 4)")

                df_pizza4 = gerar_dados_pizza(df_ind, 4)

                fig_p4 = px.pie(
                    df_pizza4,
                    values="Abs_Acumulado",
                    names="Descrição",
                    hole=0.4,
                    color_discrete_sequence=px.colors.sequential.YlOrRd
                )

                st.plotly_chart(fig_p4, use_container_width=True)

                st.write("**Top 10 Contas Analíticas:**")

                top10_n4 = df_ind[
                    (df_ind["Nivel"] == 4) &
                    (df_ind["ACUMULADO"] < 0)
                ].copy()

                top10_n4["Gasto"] = top10_n4["ACUMULADO"].abs()
                top10_n4 = top10_n4.sort_values(by="Gasto", ascending=False).head(10)

                st.table(
                    top10_n4[["Descrição", "ACUMULADO"]]
                    .rename(columns={"ACUMULADO": "Valor"})
                    .style.format({"Valor": formatar_moeda_br})
                )

            st.divider()

            # 4. COMPOSIÇÃO DAS DESPESAS SOBRE RECEITA
            st.write("### 📊 Composição das Despesas s/ Receita Líquida")

            df_perc = df_ind[df_ind["Nivel"] == 2].copy()

            df_perc["% s/ Receita"] = df_perc.apply(
                lambda x: (abs(x["ACUMULADO"]) / rec * 100) if rec > 0 else 0,
                axis=1
            )

            df_comp_view = df_perc[df_perc["Conta"] != "01"].sort_values(
                by="% s/ Receita",
                ascending=False
            )

            fig_bar_perc = px.bar(
                df_comp_view,
                x="Descrição",
                y="% s/ Receita",
                text_auto=".1f",
                color="Descrição",
                title="Impacto das Despesas (%)",
                color_discrete_sequence=px.colors.qualitative.Pastel
            )

            st.plotly_chart(fig_bar_perc, use_container_width=True)

            st.write("**Detalhamento da Composição:**")

            st.dataframe(
                df_comp_view[["Descrição", "ACUMULADO", "% s/ Receita"]]
                .style.format({
                    "ACUMULADO": formatar_moeda_br,
                    "% s/ Receita": "{:.1f}%"
                }),
                use_container_width=True
            )

with aba4:
    st.subheader("🏢 Análise de Obras e Rateio Dinâmico")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        anos_obras_sel = st.multiselect("Anos da Obra (Acumulado)", ANOS_PADRAO, default=[ano_sel], key="anos_obra_v17")
    with col_f2:
        meses_obras_sel = st.multiselect("Meses da Obra (Acumulado)", MESES_LISTA, default=MESES_LISTA, key="meses_obra_v17")

    usar_rateio = st.toggle("🔄 Ativar Visão de Custo Real (Rateio Dinâmico)", value=False)

    if st.button("📊 Processar Obras Acumuladas", key="btn_aba4_v17"):
        df_all = obter_movimentos_por_anos_meses(anos_obras_sel, meses_obras_sel)

        if df_all.empty:
            st.warning("Sem dados para o período selecionado.")
        else:
            res_cc_full = df_all.groupby("Centro de Custo").apply(
                lambda x: pd.Series({
                    "Receitas": x[x["Conta_ID"].astype(str).str.startswith("01")]["Valor_Final"].sum(),
                    "Despesa Direta": x[x["Conta_ID"].astype(str).str.startswith("02")]["Valor_Final"].sum(),
                })
            ).reset_index()

            if usar_rateio:
                df_rateio_config = carregar_logica_rateio()
                if df_rateio_config.empty:
                    st.info("ℹ️ Não foi possível ler rateio_config.")
                    st.stop()

                mapa_logica = dict(zip(df_rateio_config["Centro de Custo"], df_rateio_config["Logica"]))
                res_cc_full["Logica"] = res_cc_full["Centro de Custo"].astype(str).str.strip().map(mapa_logica).fillna("obra")
                bolo_rateio = res_cc_full.loc[res_cc_full["Logica"] == "rateio", "Despesa Direta"].sum()
                res_cc_full["Rateio Estrutura"] = 0.0

                idx_obras = (res_cc_full["Logica"] == "obra") & (res_cc_full["Despesa Direta"] != 0)
                total_desp_receptores = res_cc_full.loc[idx_obras, "Despesa Direta"].sum()

                if abs(total_desp_receptores) > 0:
                    res_cc_full.loc[idx_obras, "Rateio Estrutura"] = (
                        res_cc_full.loc[idx_obras, "Despesa Direta"] / total_desp_receptores
                    ) * bolo_rateio

                res_cc_final = res_cc_full[res_cc_full["Logica"] == "obra"].copy()
                res_cc_final["Resultado Real"] = res_cc_final["Receitas"] + res_cc_final["Despesa Direta"] + res_cc_final["Rateio Estrutura"]
                cols_v = ["Centro de Custo", "Receitas", "Despesa Direta", "Rateio Estrutura", "Resultado Real"]
            else:
                res_cc_final = res_cc_full.copy()
                res_cc_final["Resultado"] = res_cc_final["Receitas"] + res_cc_final["Despesa Direta"]
                cols_v = ["Centro de Custo", "Receitas", "Despesa Direta", "Resultado"]

            if "Todos" not in cc_sel and cc_sel:
                res_cc_final = res_cc_final[res_cc_final["Centro de Custo"].isin(cc_sel)]

            res_cc_final = res_cc_final.sort_values(by=cols_v[-1])
            somas = res_cc_final[cols_v[1:]].sum()
            linha_t = pd.DataFrame([["TOTAL CONSOLIDADO (FILTRADO)"] + somas.tolist()], columns=cols_v)
            res_cc_final = pd.concat([linha_t, res_cc_final], ignore_index=True)

            st.dataframe(res_cc_final[cols_v].style.format({c: formatar_moeda_br for c in cols_v[1:]}), use_container_width=True)

            buffer_cc = io.BytesIO()
            with pd.ExcelWriter(buffer_cc, engine="openpyxl") as writer:
                res_cc_final.to_excel(writer, index=False)

            st.download_button("📥 Exportar Obras (Excel)", data=buffer_cc.getvalue(), file_name="Obras_CustoReal.xlsx")

with aba5:
    st.subheader("⚖️ Comparativo de Períodos Independente")
    ocultar_aba5 = st.checkbox("🚫 Ocultar sem Movimento", value=False, key="ocultar_aba5_v17")

    c_p1, c_p2 = st.columns(2)
    with c_p1:
        aa = st.multiselect("Anos A", ANOS_PADRAO, key="aa_v17")
        ma = st.multiselect("Meses A", MESES_LISTA, default=MESES_LISTA, key="ma_v17")
    with c_p2:
        ab = st.multiselect("Anos B", ANOS_PADRAO, key="ab_v17")
        mb = st.multiselect("Meses B", MESES_LISTA, default=MESES_LISTA, key="mb_v17")

    if st.button("🔄 Executar Comparativo", key="btn_aba5_v17"):
        df_base_c = carregar_aba_base().copy()
        if df_base_c.empty:
            st.warning("Plano de contas vazio.")
            st.stop()

        df_base_c["Conta"] = df_base_c.apply(lambda x: limpar_conta_blindado(x["Conta"], x["Nivel"]), axis=1).astype(str).str.strip()

        def calc_soberano(anos_alvo, meses_alvo):
            map_res = {}
            df = obter_movimentos_por_anos_meses(anos_alvo, meses_alvo)
            if df.empty:
                return map_res
            if "Todos" not in cc_sel and cc_sel:
                df = df[df["Centro de Custo"].isin(cc_sel)]
            somas = df.groupby("Conta_ID")["Valor_Final"].sum().to_dict()
            for conta, valor in somas.items():
                map_res[str(conta).strip()] = map_res.get(str(conta).strip(), 0) + valor
            return map_res

        dados_a = calc_soberano(aa, ma)
        dados_b = calc_soberano(ab, mb)

        df_base_c["PERÍODO A"] = df_base_c["Conta"].map(dados_a).fillna(0)
        df_base_c["PERÍODO B"] = df_base_c["Conta"].map(dados_b).fillna(0)

        for n in [3, 2]:
            for idx, row in df_base_c[df_base_c["Nivel"] == n].iterrows():
                pref = str(row["Conta"]).strip() + "."
                df_base_c.at[idx, "PERÍODO A"] = df_base_c[(df_base_c["Nivel"] == 4) & (df_base_c["Conta"].str.startswith(pref))]["PERÍODO A"].sum()
                df_base_c.at[idx, "PERÍODO B"] = df_base_c[(df_base_c["Nivel"] == 4) & (df_base_c["Conta"].str.startswith(pref))]["PERÍODO B"].sum()

        for idx, _ in df_base_c[df_base_c["Nivel"] == 1].iterrows():
            df_base_c.at[idx, "PERÍODO A"] = df_base_c[df_base_c["Nivel"] == 2]["PERÍODO A"].sum()
            df_base_c.at[idx, "PERÍODO B"] = df_base_c[df_base_c["Nivel"] == 2]["PERÍODO B"].sum()

        df_base_c["DIFERENÇA"] = df_base_c["PERÍODO B"] - df_base_c["PERÍODO A"]
        df_base_c["VAR %"] = df_base_c.apply(lambda x: (x["DIFERENÇA"] / abs(x["PERÍODO A"]) * 100) if x["PERÍODO A"] != 0 else 0, axis=1)

        if ocultar_aba5:
            df_base_c = filtrar_linhas_zeradas(df_base_c, ["PERÍODO A", "PERÍODO B"])

        st.dataframe(df_base_c[["Nivel", "Conta", "Descrição", "PERÍODO A", "PERÍODO B", "DIFERENÇA", "VAR %"]].style.format({
            "PERÍODO A": formatar_moeda_br,
            "PERÍODO B": formatar_moeda_br,
            "DIFERENÇA": formatar_moeda_br,
            "VAR %": formatar_pct
        }), use_container_width=True, height=750)

with aba6:
    st.subheader("⚠️ Central de Alertas Preventivos")
    st.info("Nesta versão Supabase, os alertas serão recalibrados após validação do relatório e obras.")

with aba7:
    st.subheader("📉 Curva ABC de Despesas (Nível 4)")
    if st.button("🔍 Gerar Curva ABC", key="btn_aba7_final"):
        df_abc, _ = processar_bi(ano_sel, meses_sel, cc_sel)
        if df_abc is not None:
            df_an = df_abc[(df_abc["Nivel"] == 4) & (df_abc["Conta"].astype(str).str.startswith("02"))].copy()
            df_an["Valor_Abs"] = df_an["ACUMULADO"].abs()
            df_an = df_an[df_an["Valor_Abs"] > 0].sort_values(by="Valor_Abs", ascending=False)
            tot = df_an["Valor_Abs"].sum()

            if tot > 0:
                df_an["% Individual"] = (df_an["Valor_Abs"] / tot) * 100
                df_an["% Acumulado"] = df_an["% Individual"].cumsum()
                df_an["Classe"] = df_an["% Acumulado"].apply(lambda x: "A" if x <= 80.1 else ("B" if x <= 95.1 else "C"))

                fig_p = go.Figure()
                fig_p.add_trace(go.Bar(x=df_an["Descrição"], y=df_an["Valor_Abs"], name="Gasto"))
                fig_p.add_trace(go.Scatter(x=df_an["Descrição"], y=df_an["% Acumulado"], name="%", yaxis="y2"))
                fig_p.update_layout(title="Pareto", yaxis=dict(title="R$"), yaxis2=dict(title="%", overlaying="y", side="right", range=[0, 105]))
                st.plotly_chart(fig_p, use_container_width=True)
                st.dataframe(df_an[["Conta", "Descrição", "Valor_Abs", "% Individual", "% Acumulado", "Classe"]].style.format({
                    "Valor_Abs": formatar_moeda_br,
                    "% Individual": "{:.1f}%",
                    "% Acumulado": "{:.1f}%"
                }), use_container_width=True)
            else:
                st.info("Sem despesas para gerar Curva ABC.")

with aba8:
    st.subheader("🤖 Analista IA")
    st.info("A análise por IA será religada depois da estabilização da base Supabase.")

with aba9:
    st.subheader("🧾 Composição da Obra")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        anos_comp_sel = st.multiselect("Anos da Obra", ANOS_PADRAO, default=[ano_sel], key="anos_comp_obra_v17")
    with col_f2:
        meses_comp_sel = st.multiselect("Meses da Obra", MESES_LISTA, default=MESES_LISTA, key="meses_comp_obra_v17")

    obras_sel = [c for c in cc_sel if c != "Todos"]

    if not obras_sel:
        st.info("Selecione ao menos uma obra no filtro lateral de Centro de Custo.")
    
    else:
        st.write(f"📍 Obras selecionadas no filtro lateral: **{len(obras_sel)}**")
    
        if st.button("📊 Processar Composição da Obra", key="btn_comp_obra_v17"):
            df_rateio = carregar_logica_rateio()
            if df_rateio.empty:
                st.info("ℹ️ Não foi possível ler rateio_config.")
                st.stop()
    
            df_all = obter_movimentos_por_anos_meses(anos_comp_sel, meses_comp_sel)
            if df_all.empty:
                st.warning("Nenhum dado encontrado para o período selecionado.")
                st.stop()
    
            df_sel = df_all[
                (df_all["Centro de Custo"].isin(obras_sel)) &
                (df_all["Conta_ID"].astype(str).str.startswith("01") | df_all["Conta_ID"].astype(str).str.startswith("02"))
            ].copy()
    
            if df_sel.empty:
                st.warning("As obras selecionadas não possuem lançamentos no período informado.")
                st.stop()
    
            direto = df_sel.groupby("Conta_ID")["Valor_Final"].sum()
            direto_desp = direto[direto.index.astype(str).str.startswith("02")].copy()
    
            mapa_logica = dict(zip(df_rateio["Centro de Custo"], df_rateio["Logica"]))
            res_cc_full = df_all.groupby("Centro de Custo").apply(
                lambda x: pd.Series({
                    "Receitas": x[x["Conta_ID"].astype(str).str.startswith("01")]["Valor_Final"].sum(),
                    "Despesa Direta": x[x["Conta_ID"].astype(str).str.startswith("02")]["Valor_Final"].sum(),
                })
            ).reset_index()
    
            res_cc_full["Logica"] = res_cc_full["Centro de Custo"].astype(str).str.strip().map(mapa_logica).fillna("obra")
            bolo_rateio = res_cc_full.loc[res_cc_full["Logica"] == "rateio", "Despesa Direta"].sum()
            idx_obras = (res_cc_full["Logica"] == "obra") & (res_cc_full["Despesa Direta"] != 0)
            total_desp_receptores = res_cc_full.loc[idx_obras, "Despesa Direta"].sum()
    
            rateio_recebido_conjunto = 0.0
            if abs(total_desp_receptores) > 0:
                desp_direta_conjunto = res_cc_full[res_cc_full["Centro de Custo"].isin(obras_sel)]["Despesa Direta"].sum()
                rateio_recebido_conjunto = (desp_direta_conjunto / total_desp_receptores) * bolo_rateio
    
            rateado = pd.Series(0.0, index=direto.index)
            total_desp_direta = direto_desp.sum()
            if abs(total_desp_direta) > 0 and not direto_desp.empty:
                proporcao_desp = direto_desp / total_desp_direta
                rateado_desp = proporcao_desp * rateio_recebido_conjunto
                rateado.loc[rateado_desp.index] = rateado_desp
    
            final = direto + rateado
    
            df_base_comp = carregar_aba_base().copy()
            mapa_desc = dict(zip(df_base_comp["Conta"], df_base_comp["Descrição"])) if not df_base_comp.empty else {}
    
            df_final = pd.DataFrame({
                "Categoria": direto.index,
                "Descrição": [mapa_desc.get(conta, conta) for conta in direto.index],
                "Direto": direto.values,
                "Rateado": rateado.values,
                "Final": final.values
            }).sort_values(by="Categoria")
    
            total_row = pd.DataFrame([{
                "Categoria": "TOTAL",
                "Descrição": "",
                "Direto": df_final["Direto"].sum(),
                "Rateado": df_final["Rateado"].sum(),
                "Final": df_final["Final"].sum()
            }])
    
            df_final = pd.concat([df_final, total_row], ignore_index=True)
    
            st.dataframe(df_final.style.format({
                "Direto": formatar_moeda_br,
                "Rateado": formatar_moeda_br,
                "Final": formatar_moeda_br
            }), use_container_width=True, height=700)
    
            buffer_comp = io.BytesIO()
            with pd.ExcelWriter(buffer_comp, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Composicao_Obra")
    
            st.download_button("📥 Exportar Composição da Obra (Excel)", data=buffer_comp.getvalue(), file_name="Composicao_Obra_Consolidada.xlsx")

with aba10:
    st.subheader("⚙️ Configurações")

    tab_pc, tab_rateio = st.tabs([
        "📚 Plano de Contas",
        "🏢 Centros de Custo / Rateio"
    ])

    with tab_pc:
        st.write("### 📚 Plano de Contas")

        df_pc_raw = pd.DataFrame(supabase_fetch_all("plano_contas"))

        if df_pc_raw.empty:
            st.warning("Plano de contas vazio.")
        else:
            df_pc_raw = df_pc_raw.sort_values(by="conta_id").reset_index(drop=True)

            df_pc_editado = st.data_editor(
                df_pc_raw,
                use_container_width=True,
                height=600,
                num_rows="dynamic",
                column_config={
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                    "conta_id": st.column_config.TextColumn("Conta"),
                    "descricao": st.column_config.TextColumn("Descrição"),
                    "nivel": st.column_config.NumberColumn("Nível", min_value=1, max_value=5, step=1)
                }
            )

            if st.button("💾 Salvar Plano de Contas"):
                for _, row in df_pc_editado.iterrows():
                    conta = str(row.get("conta_id", "")).strip()
                    descricao = str(row.get("descricao", "")).strip().upper()
                    nivel = row.get("nivel", None)

                    if not conta or not descricao or pd.isna(nivel):
                        continue

                    if pd.isna(row.get("id")):
                        supabase_client.table("plano_contas").insert({
                            "conta_id": conta,
                            "descricao": descricao,
                            "nivel": int(nivel)
                        }).execute()
                    else:
                        supabase_client.table("plano_contas").update({
                            "conta_id": conta,
                            "descricao": descricao,
                            "nivel": int(nivel)
                        }).eq("id", int(row["id"])).execute()

                st.cache_data.clear()
                st.success("Plano de contas salvo com sucesso.")

    with tab_rateio:
        st.write("### 🏢 Centros de Custo / Rateio")

        df_rateio_raw = pd.DataFrame(supabase_fetch_all("rateio_config"))

        if df_rateio_raw.empty:
            st.warning("Nenhum centro de custo cadastrado.")
        else:
            df_rateio_raw = df_rateio_raw.sort_values(by="centro_custo").reset_index(drop=True)

            df_rateio_editado = st.data_editor(
                df_rateio_raw,
                use_container_width=True,
                height=650,
                num_rows="dynamic",
                column_config={
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                    "centro_custo": st.column_config.TextColumn("Centro de Custo"),
                    "logica": st.column_config.SelectboxColumn(
                        "Lógica",
                        options=["obra", "rateio", "fora"],
                        required=True
                    )
                }
            )

            if st.button("💾 Salvar Centros de Custo / Rateio"):
                for _, row in df_rateio_editado.iterrows():
                    centro = str(row.get("centro_custo", "")).strip()
                    logica = str(row.get("logica", "")).strip().lower()

                    if not centro or logica not in ["obra", "rateio", "fora"]:
                        continue

                    if pd.isna(row.get("id")):
                        supabase_client.table("rateio_config").insert({
                            "centro_custo": centro,
                            "logica": logica
                        }).execute()
                    else:
                        supabase_client.table("rateio_config").update({
                            "centro_custo": centro,
                            "logica": logica
                        }).eq("id", int(row["id"])).execute()

                st.cache_data.clear()
                st.success("Centros de custo atualizados com sucesso.")
                
with aba11:
    st.subheader("📊 Resultado Operacional / Não Operacional")

    filtro_classificacao = st.radio(
        "Escolha a visão",
        ["operacional", "nao_operacional", "todos"],
        horizontal=True
    )

    ocultar_vazios_aba11 = st.checkbox(
        "🚫 Ocultar Contas sem Movimento",
        value=False,
        key="ocultar_aba11"
    )

    if st.button("📊 Gerar Relatório", key="btn_aba11_resultado_operacional"):
        df_res, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)

        if df_res is None:
            st.error("❌ Não foi possível gerar o relatório.")
        else:
            colunas_valores = meses_exibir + ["MÉDIA", "ACUMULADO"]

            if filtro_classificacao != "todos":
                mask_manter = df_res["Classificacao"] == filtro_classificacao

                for col in colunas_valores:
                    if col in df_res.columns:
                        df_res.loc[~mask_manter, col] = 0.0

                # Recalcula níveis superiores depois do filtro
                for col in meses_exibir:
                    for n in sorted(df_res["Nivel"].dropna().unique(), reverse=True):
                        if n <= 1:
                            continue

                        nivel_pai = n - 1

                        for idx, row in df_res[df_res["Nivel"] == nivel_pai].iterrows():
                            pref = str(row["Conta"]).strip() + "."
                            total_filhos = df_res[
                                (df_res["Nivel"] == n) &
                                (df_res["Conta"].astype(str).str.startswith(pref))
                            ][col].sum()

                            df_res.at[idx, col] = total_filhos

                    for idx, _ in df_res[df_res["Nivel"] == 1].iterrows():
                        df_res.at[idx, col] = df_res[df_res["Nivel"] == 2][col].sum()

                df_res["ACUMULADO"] = df_res[meses_exibir].sum(axis=1)
                df_res["MÉDIA"] = df_res[meses_exibir].mean(axis=1)

            if ocultar_vazios_aba11:
                df_res = filtrar_linhas_zeradas(df_res, meses_exibir + ["ACUMULADO"])

            df_visual = df_res[df_res["Nivel"].isin(niveis_sel)].copy()

            cols_export = [
                "Nivel", "Conta", "Descrição", "Classificacao"
            ] + meses_exibir + ["MÉDIA", "ACUMULADO"]

            def style_rows_operacional(row):
                if row["Nivel"] == 1:
                    return ["background-color: #334155; color: white; font-weight: bold"] * len(row)
                if row["Nivel"] == 2:
                    return ["background-color: #cbd5e1; font-weight: bold; color: black"] * len(row)
                if row["Nivel"] == 3:
                    return ["background-color: #D1EAFF; font-weight: bold; color: black"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_visual[cols_export]
                .style
                .apply(style_rows_operacional, axis=1)
                .format({
                    c: formatar_moeda_br
                    for c in cols_export
                    if c not in ["Nivel", "Conta", "Descrição", "Classificacao"]
                }),
                use_container_width=True,
                height=800
            )
                
