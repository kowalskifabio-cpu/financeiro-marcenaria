import io

import pandas as pd
import plotly.express as px
import streamlit as st

from servico_orcamento import (
    MESES_NUMERO_NOME,
    carregar_itens_orcamento,
    carregar_orcamentos,
)

from servico_orcado_realizado import (
    calcular_forecast,
    consolidar_hierarquia,
    maiores_desvios,
    montar_comparativo_gerencial,
    montar_orcado_analitico,
    montar_realizado_analitico,
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
        numero = (
            f"{float(valor):,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
        return f"{numero}%"
    except Exception:
        return "0,00%"


def _style_rows(row):
    nivel = row.get("Nivel", 0)

    if nivel == 1:
        return [
            "background-color: #334155; color: white; font-weight: bold"
        ] * len(row)

    if nivel == 2:
        return [
            "background-color: #cbd5e1; color: black; font-weight: bold"
        ] * len(row)

    if nivel == 3:
        return [
            "background-color: #D1EAFF; color: black; font-weight: bold"
        ] * len(row)

    return [""] * len(row)


def _gerar_excel(
    df,
    ano,
    versao,
    visao
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
                42
            )

        nomes_monetarios = {
            "Orçado",
            "Realizado",
            "Desvio R$",
            "Forecast"
        }

        cabecalhos = {
            celula.value: celula.column
            for celula in ws[1]
        }

        for nome in nomes_monetarios:
            if nome not in cabecalhos:
                continue

            coluna_num = cabecalhos[nome]

            for linha in range(
                2,
                ws.max_row + 1
            ):
                ws.cell(
                    row=linha,
                    column=coluna_num
                ).number_format = (
                    'R$ #,##0.00;[Red]-R$ #,##0.00'
                )

        if "Desvio %" in cabecalhos:
            coluna_num = cabecalhos["Desvio %"]

            for linha in range(
                2,
                ws.max_row + 1
            ):
                ws.cell(
                    row=linha,
                    column=coluna_num
                ).number_format = '0.00"%"'

    nome_arquivo = (
        f"Orcado_x_Realizado_{ano}_"
        f"Versao_{versao}_{visao}.xlsx"
    )

    return buffer.getvalue(), nome_arquivo


def _filtrar_classificacao(
    df,
    filtro_classificacao
):
    """
    Filtra pelas contas finais da classificação escolhida e
    reconstrói toda a hierarquia: pais, avós e resultado.

    Assim uma visão Diretoria, por exemplo, não mistura
    valores Operacionais nos totais das contas-mãe.
    """

    if df is None or df.empty:
        return df

    if filtro_classificacao == "todos":
        return df.copy()

    resultado = df.copy()

    resultado["Conta"] = (
        resultado["Conta"]
        .astype(str)
        .str.strip()
    )

    resultado["Classificacao"] = (
        resultado["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    contas = resultado["Conta"].tolist()

    def eh_folha(conta):
        prefixo = str(conta).strip() + "."

        return not any(
            str(outra).startswith(prefixo)
            for outra in contas
            if str(outra) != str(conta)
        )

    resultado["_EhFolha"] = (
        resultado["Conta"]
        .apply(eh_folha)
    )

    colunas_valores = [
        coluna
        for coluna in [
            "Orçado",
            "Realizado",
            "Forecast"
        ]
        if coluna in resultado.columns
    ]

    # Mantém valores somente nas folhas pertencentes
    # à classificação selecionada.
    manter_valor = (
        resultado["_EhFolha"]
        &
        (
            resultado["Classificacao"]
            == filtro_classificacao
        )
    )

    for coluna in colunas_valores:
        resultado[coluna] = pd.to_numeric(
            resultado[coluna],
            errors="coerce"
        ).fillna(0.0)

        resultado.loc[
            ~manter_valor,
            coluna
        ] = 0.0

    # Recalcula novamente toda a árvore.
    resultado = consolidar_hierarquia(
        resultado,
        colunas_valores
    )

    # Recalcula desvios depois da consolidação.
    resultado["Desvio R$"] = (
        resultado["Realizado"]
        - resultado["Orçado"]
    )

    resultado["Desvio %"] = 0.0

    mask_orcado = (
        resultado["Orçado"] != 0
    )

    resultado.loc[
        mask_orcado,
        "Desvio %"
    ] = (
        resultado.loc[
            mask_orcado,
            "Desvio R$"
        ]
        /
        resultado.loc[
            mask_orcado,
            "Orçado"
        ].abs()
        * 100
    )

    resultado = resultado.drop(
        columns=["_EhFolha"],
        errors="ignore"
    )

    return resultado


def _calcular_resumo_executivo(df):
    if df is None or df.empty:
        return {
            "orcado": 0.0,
            "realizado": 0.0,
            "desvio": 0.0,
            "desvio_pct": 0.0
        }

    # Para evitar dupla contagem da hierarquia,
    # o resumo utiliza apenas o nível 1.
    nivel_1 = df[
        df["Nivel"] == 1
    ].copy()

    if nivel_1.empty:
        nivel_1 = df[
            df["Nivel"] == df["Nivel"].min()
        ].copy()

    orcado = pd.to_numeric(
        nivel_1["Orçado"],
        errors="coerce"
    ).fillna(0.0).sum()

    realizado = pd.to_numeric(
        nivel_1["Realizado"],
        errors="coerce"
    ).fillna(0.0).sum()

    desvio = realizado - orcado

    desvio_pct = (
        desvio / abs(orcado) * 100
        if orcado != 0
        else 0.0
    )

    return {
        "orcado": float(orcado),
        "realizado": float(realizado),
        "desvio": float(desvio),
        "desvio_pct": float(desvio_pct)
    }


def render_aba_orcado_realizado(
    supabase_client,
    carregar_aba_base,
    processar_bi,
    ano_sel,
    meses_sel,
    cc_sel
):
    st.subheader("📊 Orçado × Realizado")

    st.caption(
        "Visão gerencial do orçamento aprovado comparado "
        "ao realizado financeiro."
    )

    # =====================================================
    # ORÇAMENTO
    # =====================================================

    df_orcamentos = carregar_orcamentos(
        supabase_client
    )

    if df_orcamentos.empty:
        st.warning(
            "Nenhum orçamento cadastrado."
        )
        return

    df_validos = df_orcamentos[
        df_orcamentos["status"].isin(
            ["aprovado", "bloqueado"]
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
        + " — "
        + df_validos["nome"].astype(str)
        + " — Versão "
        + df_validos["versao"].astype(str)
        + " — "
        + df_validos["status"]
        .astype(str)
        .str.replace("_", " ", regex=False)
        .str.title()
    )

    rotulo = st.selectbox(
        "Orçamento para comparação",
        options=df_validos["rotulo"].tolist(),
        key="orcado_realizado_v2_orcamento"
    )

    linha_orcamento = df_validos[
        df_validos["rotulo"] == rotulo
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
    # FILTROS
    # =====================================================

    col_visao, col_classificacao = st.columns(2)

    with col_visao:
        visao = st.radio(
            "Visão",
            options=[
                "acumulado",
                "mensal",
                "forecast"
            ],
            format_func=lambda valor: {
                "acumulado": "Acumulado",
                "mensal": "Mensal",
                "forecast": "Forecast"
            }[valor],
            horizontal=True,
            key="visao_orcado_realizado_v2"
        )

    with col_classificacao:
        filtro_classificacao = st.radio(
            "Classificação",
            options=[
                "todos",
                "operacional",
                "nao_operacional",
                "diretoria",
                "diretoria_investimentos"
            ],
            format_func=lambda valor: {
                "todos": "Todos",
                "operacional": "Operacional",
                "nao_operacional": "Não Operacional",
                "diretoria": "Diretoria",
                "diretoria_investimentos": "Diretoria Investimentos"
            }[valor],
            horizontal=True,
            key="classificacao_orcado_realizado_v2"
        )

    niveis_sel = st.multiselect(
        "Níveis do plano de contas",
        options=[1, 2, 3, 4],
        default=[1, 2, 3, 4],
        key="niveis_orcado_realizado_v2"
    )

            
    ocultar_zerados = st.checkbox(
        "Ocultar contas sem orçamento e sem realizado",
        value=True,
        key="ocultar_zerados_orcado_realizado_v2"
    )

    if visao == "mensal":
        mes_unico = st.selectbox(
            "Mês",
            options=MESES,
            index=(
                MESES.index(meses_sel[-1])
                if meses_sel
                and meses_sel[-1] in MESES
                else 0
            ),
            key="mes_unico_orcado_realizado_v2"
        )

        meses_comparar = [mes_unico]

    else:
        meses_comparar = st.multiselect(
            "Meses realizados para análise",
            options=MESES,
            default=[
                mes
                for mes in meses_sel
                if mes in MESES
            ],
            key="meses_orcado_realizado_v2"
        )

    if not meses_comparar:
        st.info(
            "Selecione ao menos um mês."
        )
        return

    # =====================================================
    # FORECAST
    # =====================================================

    mes_fechamento = None

    if visao == "forecast":
        mes_fechamento = st.selectbox(
            "Último mês considerado realizado",
            options=MESES,
            index=(
                max(
                    0,
                    min(
                        len(MESES) - 1,
                        max(
                            [
                                MESES.index(m)
                                for m in meses_comparar
                                if m in MESES
                            ],
                            default=0
                        )
                    )
                )
            ),
            key="mes_fechamento_forecast_v2"
        )

        indice_fechamento = MESES.index(
            mes_fechamento
        )

        meses_realizados = MESES[
            :indice_fechamento + 1
        ]

        meses_futuros = MESES[
            indice_fechamento + 1:
        ]

        meses_comparar = meses_realizados

    # =====================================================
    # PROCESSAMENTO
    # =====================================================

    if not st.button(
        "📊 Gerar análise gerencial",
        key="btn_gerar_orcado_realizado_v2"
    ):
        return

    with st.spinner(
        "Processando Orçado × Realizado..."
    ):

        df_itens = carregar_itens_orcamento(
            supabase_client=supabase_client,
            orcamento_id=orcamento_id
        )

        df_plano = carregar_aba_base().copy()

        df_orcado = montar_orcado_analitico(
            df_itens=df_itens,
            meses_selecionados=meses_comparar
        )

        df_bi, _ = processar_bi(
            ano_orcamento,
            meses_comparar,
            cc_sel
        )

        if df_bi is None:
            st.warning(
                "Não foi possível carregar o realizado."
            )
            return

        df_realizado = montar_realizado_analitico(
            df_bi=df_bi,
            meses_selecionados=meses_comparar
        )

        comparativo = montar_comparativo_gerencial(
            df_plano=df_plano,
            df_orcado=df_orcado,
            df_realizado=df_realizado,
            meses_selecionados=meses_comparar
        )

        if comparativo.empty:
            st.warning(
                "Não há dados para comparação."
            )
            return

        comparativo = _filtrar_classificacao(
            comparativo,
            filtro_classificacao
        )

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

        comparativo_visual = comparativo[
            comparativo["Nivel"].isin(
                niveis_sel
            )
        ].copy()

    # =====================================================
    # RESUMO EXECUTIVO
    # =====================================================

    resumo = _calcular_resumo_executivo(
        comparativo
    )

    forecast_info = None

    if visao == "forecast":
        forecast_info = calcular_forecast(
            df_itens=df_itens,
            df_bi=df_bi,
            meses_realizados=meses_realizados,
            meses_futuros=meses_futuros
        )

    if visao == "forecast":
        c1, c2, c3, c4, c5 = st.columns(5)
    else:
        c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Orçado",
        _moeda(resumo["orcado"])
    )

    c2.metric(
        "Realizado",
        _moeda(resumo["realizado"])
    )

    c3.metric(
        "Desvio",
        _moeda(resumo["desvio"])
    )

    c4.metric(
        "Desvio %",
        _percentual(
            resumo["desvio_pct"]
        )
    )

    if visao == "forecast":
        c5.metric(
            "Forecast anual",
            _moeda(
                forecast_info[
                    "forecast"
                ]
            )
        )

        st.caption(
            "Forecast = realizado até "
            f"{mes_fechamento} + orçamento dos meses futuros."
        )

    # =====================================================
    # DRE GERENCIAL
    # =====================================================

    st.divider()
    st.write("### DRE Gerencial — Orçado × Realizado")

    colunas_exibir = [
        "Nivel",
        "Conta",
        "Descrição",
        "Classificacao",
        "Orçado",
        "Realizado",
        "Desvio R$",
        "Desvio %",
        "Status"
    ]

    if visao == "forecast":
        # Forecast por conta analítica:
        # realizado no período + orçamento futuro da mesma conta.
        df_orcado_futuro = montar_orcado_analitico(
            df_itens=df_itens,
            meses_selecionados=meses_futuros
        )

        mapa_futuro = {}

        if not df_orcado_futuro.empty:
            mapa_futuro = dict(
                zip(
                    df_orcado_futuro["conta_id"].astype(str),
                    df_orcado_futuro["Orçado"]
                )
            )

        comparativo_forecast = comparativo.copy()

        comparativo_forecast["Forecast"] = 0.0

        mask_analitica = (
            comparativo_forecast["Nivel"] >= 4
        )

        comparativo_forecast.loc[
            mask_analitica,
            "Forecast"
        ] = (
            comparativo_forecast.loc[
                mask_analitica,
                "Realizado"
            ]
            +
            comparativo_forecast.loc[
                mask_analitica,
                "Conta"
            ]
            .astype(str)
            .map(mapa_futuro)
            .fillna(0.0)
        )

        from servico_orcado_realizado import consolidar_hierarquia

        comparativo_forecast = consolidar_hierarquia(
            comparativo_forecast,
            ["Forecast"]
        )

        comparativo_visual = comparativo_forecast[
            comparativo_forecast[
                "Nivel"
            ].isin(niveis_sel)
        ].copy()

        if ocultar_zerados:
            comparativo_visual = comparativo_visual[
                (
                    comparativo_visual["Orçado"] != 0
                )
                |
                (
                    comparativo_visual["Realizado"] != 0
                )
                |
                (
                    comparativo_visual["Forecast"] != 0
                )
            ].copy()

        colunas_exibir.insert(
            -1,
            "Forecast"
        )

    formato = {
        "Orçado": _moeda,
        "Realizado": _moeda,
        "Desvio R$": _moeda,
        "Desvio %": _percentual
    }

    if "Forecast" in colunas_exibir:
        formato["Forecast"] = _moeda

    st.dataframe(
        comparativo_visual[
            colunas_exibir
        ]
        .style
        .apply(
            _style_rows,
            axis=1
        )
        .format(formato),
        use_container_width=True,
        height=800
    )

    # =====================================================
    # MAIORES DESVIOS
    # =====================================================

    st.divider()
    st.write("### Maiores desvios")

    top_desvios = maiores_desvios(
        comparativo,
        quantidade=10
    )

    if top_desvios.empty:
        st.info(
            "Não existem desvios relevantes no período."
        )
    else:
        top_desvios = top_desvios.copy()

        top_desvios["Desvio Absoluto"] = (
            top_desvios["Desvio R$"].abs()
        )

        fig = px.bar(
            top_desvios.sort_values(
                "Desvio Absoluto",
                ascending=True
            ),
            x="Desvio Absoluto",
            y="Descrição",
            orientation="h",
            text="Desvio R$",
            title="Top 10 desvios por valor absoluto"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        col_desfavoravel, col_favoravel = st.columns(2)

        with col_desfavoravel:
            st.write(
                "#### 🔴 Desvios desfavoráveis"
            )

            df_desfavoravel = (
                top_desvios[
                    top_desvios[
                        "Status"
                    ] == "Desfavorável"
                ]
                .sort_values(
                    "Desvio Absoluto",
                    ascending=False
                )
            )

            if df_desfavoravel.empty:
                st.info(
                    "Nenhum desvio desfavorável."
                )
            else:
                st.dataframe(
                    df_desfavoravel[
                        [
                            "Conta",
                            "Descrição",
                            "Orçado",
                            "Realizado",
                            "Desvio R$",
                            "Desvio %"
                        ]
                    ]
                    .style
                    .format({
                        "Orçado": _moeda,
                        "Realizado": _moeda,
                        "Desvio R$": _moeda,
                        "Desvio %": _percentual
                    }),
                    use_container_width=True,
                    hide_index=True
                )

        with col_favoravel:
            st.write(
                "#### 🟢 Desvios favoráveis"
            )

            df_favoravel = (
                top_desvios[
                    top_desvios[
                        "Status"
                    ] == "Favorável"
                ]
                .sort_values(
                    "Desvio Absoluto",
                    ascending=False
                )
            )

            if df_favoravel.empty:
                st.info(
                    "Nenhum desvio favorável."
                )
            else:
                st.dataframe(
                    df_favoravel[
                        [
                            "Conta",
                            "Descrição",
                            "Orçado",
                            "Realizado",
                            "Desvio R$",
                            "Desvio %"
                        ]
                    ]
                    .style
                    .format({
                        "Orçado": _moeda,
                        "Realizado": _moeda,
                        "Desvio R$": _moeda,
                        "Desvio %": _percentual
                    }),
                    use_container_width=True,
                    hide_index=True
                )

    # =====================================================
    # EXPORTAÇÃO
    # =====================================================

    st.divider()

    excel, nome_excel = _gerar_excel(
        comparativo_visual[
            colunas_exibir
        ],
        ano=ano_orcamento,
        versao=versao,
        visao=visao
    )

    st.download_button(
        "📥 Exportar análise para Excel",
        data=excel,
        file_name=nome_excel,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
