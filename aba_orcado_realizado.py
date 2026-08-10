import io

import pandas as pd
import streamlit as st

from servico_orcamento import (
    MESES_NUMERO_NOME,
    carregar_itens_orcamento,
    carregar_orcamentos,
)


MESES = list(MESES_NUMERO_NOME.values())


def _moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    sinal = "-" if valor < 0 else ""

    numero = (
        f"{abs(valor):,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    return f"{sinal}R$ {numero}"


def _percentual(valor):
    try:
        return f"{float(valor):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def _classificar_desvio(conta, orcado, realizado):
    """
    Interpreta o desvio gerencialmente.

    Receita:
        realizado maior que orçado = favorável

    Despesa:
        realizado menos negativo que orçado = favorável
        realizado mais negativo que orçado = desfavorável
    """

    conta = str(conta).strip()

    if conta.startswith("01"):
        if realizado > orcado:
            return "Favorável"
        elif realizado < orcado:
            return "Desfavorável"
        return "Dentro do orçamento"

    if conta.startswith("02"):
        if realizado > orcado:
            return "Favorável"
        elif realizado < orcado:
            return "Desfavorável"
        return "Dentro do orçamento"

    return "Neutro"


def _montar_orcado_mensal(
    df_itens,
    meses_selecionados
):
    """
    Transforma os itens do orçamento em uma tabela
    com uma linha por conta.
    """

    if df_itens is None or df_itens.empty:
        return pd.DataFrame()

    df = df_itens.copy()

    df["conta_id"] = (
        df["conta_id"]
        .astype(str)
        .str.strip()
    )

    df["mes"] = pd.to_numeric(
        df["mes"],
        errors="coerce"
    )

    df["valor_orcado"] = pd.to_numeric(
        df["valor_orcado"],
        errors="coerce"
    ).fillna(0.0)

    df = df.dropna(
        subset=["mes"]
    ).copy()

    df["mes"] = df["mes"].astype(int)

    numero_para_nome = MESES_NUMERO_NOME

    df["Mes_Nome"] = (
        df["mes"]
        .map(numero_para_nome)
    )

    df = df[
        df["Mes_Nome"].isin(
            meses_selecionados
        )
    ].copy()

    tabela = df.pivot_table(
        index="conta_id",
        columns="Mes_Nome",
        values="valor_orcado",
        aggfunc="sum",
        fill_value=0.0
    )

    for mes in meses_selecionados:
        if mes not in tabela.columns:
            tabela[mes] = 0.0

    tabela = tabela[
        meses_selecionados
    ].copy()

    tabela["Orçado"] = tabela[
        meses_selecionados
    ].sum(axis=1)

    return tabela.reset_index()


def _montar_realizado(
    df_realizado,
    meses_selecionados
):
    """
    Usa o resultado já produzido pela função processar_bi
    do sistema financeiro.
    """

    if df_realizado is None or df_realizado.empty:
        return pd.DataFrame()

    df = df_realizado.copy()

    df["Conta"] = (
        df["Conta"]
        .astype(str)
        .str.strip()
    )

    for mes in meses_selecionados:
        if mes not in df.columns:
            df[mes] = 0.0

        df[mes] = pd.to_numeric(
            df[mes],
            errors="coerce"
        ).fillna(0.0)

    # Nesta primeira versão comparamos nas contas analíticas.
    df = df[
        df["Nivel"] >= 4
    ].copy()

    df["Realizado"] = df[
        meses_selecionados
    ].sum(axis=1)

    return df


def _montar_comparativo(
    df_orcado,
    df_realizado,
    df_plano,
    meses_selecionados
):
    """
    Junta orçamento e realizado pela conta contábil.
    """

    if df_plano is None or df_plano.empty:
        return pd.DataFrame()

    plano = df_plano.copy()

    plano["Conta"] = (
        plano["Conta"]
        .astype(str)
        .str.strip()
    )

    plano["Descrição"] = (
        plano["Descrição"]
        .astype(str)
        .str.strip()
    )

    plano = plano[
        plano["Nivel"] >= 4
    ][
        [
            "Conta",
            "Descrição",
            "Nivel",
            "Classificacao"
        ]
    ].copy()

    orcado = df_orcado.copy()

    if orcado.empty:
        orcado = pd.DataFrame(
            columns=[
                "conta_id",
                "Orçado"
            ] + meses_selecionados
        )

    orcado = orcado.rename(
        columns={
            "conta_id": "Conta"
        }
    )

    realizado = df_realizado.copy()

    colunas_realizado = [
        "Conta",
        "Realizado"
    ] + meses_selecionados

    if realizado.empty:
        realizado = pd.DataFrame(
            columns=colunas_realizado
        )
    else:
        realizado = realizado[
            colunas_realizado
        ].copy()

    comparativo = plano.merge(
        orcado,
        on="Conta",
        how="left"
    )

    comparativo = comparativo.merge(
        realizado,
        on="Conta",
        how="left",
        suffixes=("_Orçado", "_Realizado")
    )

    comparativo["Orçado"] = pd.to_numeric(
        comparativo.get("Orçado", 0),
        errors="coerce"
    ).fillna(0.0)

    comparativo["Realizado"] = pd.to_numeric(
        comparativo.get("Realizado", 0),
        errors="coerce"
    ).fillna(0.0)

    comparativo["Desvio R$"] = (
        comparativo["Realizado"]
        - comparativo["Orçado"]
    )

    comparativo["Desvio %"] = 0.0

    mask_orcado = (
        comparativo["Orçado"] != 0
    )

    comparativo.loc[
        mask_orcado,
        "Desvio %"
    ] = (
        comparativo.loc[
            mask_orcado,
            "Desvio R$"
        ]
        /
        comparativo.loc[
            mask_orcado,
            "Orçado"
        ].abs()
        * 100
    )

    comparativo["Status"] = comparativo.apply(
        lambda row: _classificar_desvio(
            row["Conta"],
            row["Orçado"],
            row["Realizado"]
        ),
        axis=1
    )

    return comparativo


def _gerar_excel(
    df,
    ano,
    versao
):
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Orçado x Realizado"
        )

        ws = writer.sheets[
            "Orçado x Realizado"
        ]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for coluna in ws.columns:
            letra = coluna[0].column_letter

            maior = 0

            for celula in coluna:
                valor = (
                    ""
                    if celula.value is None
                    else str(celula.value)
                )

                maior = max(
                    maior,
                    len(valor)
                )

            ws.column_dimensions[
                letra
            ].width = min(
                maior + 3,
                40
            )

    nome = (
        f"Orcado_x_Realizado_"
        f"{ano}_Versao_{versao}.xlsx"
    )

    return buffer.getvalue(), nome


def render_aba_orcado_realizado(
    supabase_client,
    carregar_aba_base,
    processar_bi,
    ano_sel,
    meses_sel,
    cc_sel
):
    st.subheader(
        "📊 Orçado × Realizado"
    )

    st.caption(
        "Comparação entre o orçamento aprovado e "
        "os movimentos financeiros efetivamente realizados."
    )

    # =====================================================
    # ORÇAMENTOS DISPONÍVEIS
    # =====================================================

    df_orcamentos = carregar_orcamentos(
        supabase_client
    )

    if df_orcamentos.empty:
        st.warning(
            "Nenhum orçamento cadastrado."
        )
        return

    # Para controle gerencial, priorizamos versões aprovadas
    # ou bloqueadas.
    df_validos = df_orcamentos[
        df_orcamentos["status"].isin(
            [
                "aprovado",
                "bloqueado"
            ]
        )
    ].copy()

    if df_validos.empty:
        st.warning(
            "Não existe orçamento aprovado ou bloqueado "
            "para comparação."
        )
        return

    df_validos["rotulo"] = (
        df_validos["ano"].astype(str)
        + " — Versão "
        + df_validos["versao"].astype(str)
        + " — "
        + df_validos["status"]
        .astype(str)
        .str.replace(
            "_",
            " ",
            regex=False
        )
        .str.title()
    )

    rotulo = st.selectbox(
        "Orçamento para comparação",
        options=df_validos[
            "rotulo"
        ].tolist(),
        key="orcado_realizado_orcamento"
    )

    linha_orcamento = df_validos[
        df_validos["rotulo"]
        == rotulo
    ].iloc[0]

    orcamento_id = int(
        linha_orcamento["id"]
    )

    ano_orcamento = int(
        linha_orcamento["ano"]
    )

    versao = int(
        linha_orcamento["versao"]
    )

    # =====================================================
    # FILTROS PRÓPRIOS DA ABA
    # =====================================================

    meses_disponiveis = MESES

    meses_comparar = st.multiselect(
        "Meses para comparar",
        options=meses_disponiveis,
        default=[
            mes
            for mes in meses_sel
            if mes in meses_disponiveis
        ],
        key="meses_orcado_realizado"
    )

    if not meses_comparar:
        st.info(
            "Selecione ao menos um mês."
        )
        return

    filtro_classificacao = st.radio(
        "Classificação",
        options=[
            "todos",
            "operacional",
            "nao_operacional"
        ],
        format_func=lambda valor: {
            "todos": "Todos",
            "operacional": "Operacional",
            "nao_operacional": "Não Operacional"
        }[valor],
        horizontal=True,
        key="classificacao_orcado_realizado"
    )

    ocultar_zerados = st.checkbox(
        "Ocultar contas sem orçamento e sem realizado",
        value=True,
        key="ocultar_zerados_orcado_realizado"
    )

    # =====================================================
    # PROCESSAMENTO
    # =====================================================

    if not st.button(
        "📊 Gerar comparativo",
        key="btn_gerar_orcado_realizado"
    ):
        return

    with st.spinner(
        "Montando Orçado × Realizado..."
    ):

        df_itens = carregar_itens_orcamento(
            supabase_client=supabase_client,
            orcamento_id=orcamento_id
        )

        df_orcado = _montar_orcado_mensal(
            df_itens=df_itens,
            meses_selecionados=meses_comparar
        )

        # Realizado usa exatamente a mesma função
        # utilizada pelo relatório financeiro principal.
        df_bi, meses_processados = processar_bi(
            ano_orcamento,
            meses_comparar,
            cc_sel
        )

        if df_bi is None:
            st.warning(
                "Não foi possível carregar o realizado."
            )
            return

        df_realizado = _montar_realizado(
            df_realizado=df_bi,
            meses_selecionados=meses_comparar
        )

        df_plano = carregar_aba_base().copy()

        comparativo = _montar_comparativo(
            df_orcado=df_orcado,
            df_realizado=df_realizado,
            df_plano=df_plano,
            meses_selecionados=meses_comparar
        )

        if comparativo.empty:
            st.warning(
                "Não há dados para comparação."
            )
            return

        if filtro_classificacao != "todos":
            comparativo = comparativo[
                comparativo[
                    "Classificacao"
                ]
                .fillna("operacional")
                .astype(str)
                .str.lower()
                .str.strip()
                == filtro_classificacao
            ].copy()

        if ocultar_zerados:
            comparativo = comparativo[
                (
                    comparativo["Orçado"] != 0
                )
                |
                (
                    comparativo["Realizado"] != 0
                )
            ].copy()

    # =====================================================
    # RESUMO
    # =====================================================

    total_orcado = comparativo[
        "Orçado"
    ].sum()

    total_realizado = comparativo[
        "Realizado"
    ].sum()

    total_desvio = (
        total_realizado
        - total_orcado
    )

    percentual_total = (
        total_desvio
        / abs(total_orcado)
        * 100
        if total_orcado != 0
        else 0.0
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Orçado",
        _moeda(total_orcado)
    )

    c2.metric(
        "Realizado",
        _moeda(total_realizado)
    )

    c3.metric(
        "Desvio",
        _moeda(total_desvio)
    )

    c4.metric(
        "Desvio %",
        _percentual(
            percentual_total
        )
    )

    # =====================================================
    # TABELA
    # =====================================================

    st.write(
        "### Comparativo por conta"
    )

    colunas_exibir = [
        "Conta",
        "Descrição",
        "Classificacao",
        "Orçado",
        "Realizado",
        "Desvio R$",
        "Desvio %",
        "Status"
    ]

    def estilo_status(row):
        if row["Status"] == "Favorável":
            return [
                "background-color: #dcfce7"
            ] * len(row)

        if row["Status"] == "Desfavorável":
            return [
                "background-color: #fee2e2"
            ] * len(row)

        return [""] * len(row)

    st.dataframe(
        comparativo[
            colunas_exibir
        ]
        .style
        .apply(
            estilo_status,
            axis=1
        )
        .format({
            "Orçado": _moeda,
            "Realizado": _moeda,
            "Desvio R$": _moeda,
            "Desvio %": _percentual
        }),
        use_container_width=True,
        height=750
    )

    # =====================================================
    # EXPORTAÇÃO
    # =====================================================

    excel, nome_excel = _gerar_excel(
        comparativo[
            colunas_exibir
        ],
        ano=ano_orcamento,
        versao=versao
    )

    st.download_button(
        "📥 Exportar Orçado × Realizado",
        data=excel,
        file_name=nome_excel,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
