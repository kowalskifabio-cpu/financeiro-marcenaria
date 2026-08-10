import streamlit as st
import pandas as pd

from servico_controladoria import (
    calcular_resumo_financeiro,
    calcular_receita_mensal,
    calcular_resultado_mensal,
    calcular_margem_mensal,
    top_contas_despesa,
    top_contas_receita,
    gerar_alertas_financeiros,
    calcular_resultado_por_centro_custo,
    ranking_melhores_obras,
    ranking_piores_obras,
    calcular_resumo_obras
)


def render_aba_painel_executivo(
    ano_sel,
    meses_sel,
    cc_sel,
    processar_bi,
    formatar_moeda_br,
    obter_movimentos_por_anos_meses,
    carregar_logica_rateio
):
    
    st.title("🎯 Painel Executivo")

    df_bi, meses_processados = processar_bi(
        ano_sel,
        meses_sel,
        cc_sel
    )

    if df_bi is None or df_bi.empty:
        st.warning("Sem dados.")
        return

    resumo = calcular_resumo_financeiro(
        df_bi,
        meses_sel
    )

    col1,col2,col3,col4 = st.columns(4)

    col1.metric(
        "Receita",
        formatar_moeda_br(resumo["receita"])
    )

    col2.metric(
        "Despesas",
        formatar_moeda_br(resumo["despesas"])
    )

    col3.metric(
        "Resultado",
        formatar_moeda_br(resumo["resultado"])
    )

    col4.metric(
        "Margem",
        f'{resumo["margem"]:.2f}%'
    )

    st.divider()

    st.subheader("Receita Mensal")

    receita = calcular_receita_mensal(
        df_bi,
        meses_sel
    )

    if not receita.empty:
        receita = receita.set_index("Mês")
        st.line_chart(receita)

    st.subheader("Resultado Mensal")

    resultado = calcular_resultado_mensal(
        df_bi,
        meses_sel
    )

    if not resultado.empty:
        resultado = resultado.set_index("Mês")
        st.bar_chart(resultado)

    st.subheader("Margem Mensal")

    margem = calcular_margem_mensal(
        df_bi,
        meses_sel
    )

    if not margem.empty:
        margem = margem.set_index("Mês")
        st.line_chart(margem)

    st.divider()

    colA,colB = st.columns(2)

    with colA:

        st.subheader("Top Receitas")

        top = top_contas_receita(
            df_bi,
            meses_sel,
            10
        )

        if not top.empty:
            st.dataframe(
                top[
                    [
                        "Conta",
                        "Descrição",
                        "ACUMULADO_CONTROLADORIA"
                    ]
                ],
                use_container_width=True
            )

    with colB:

        st.subheader("Top Despesas")

        top = top_contas_despesa(
            df_bi,
            meses_sel,
            10
        )

        if not top.empty:
            st.dataframe(
                top[
                    [
                        "Conta",
                        "Descrição",
                        "ACUMULADO_CONTROLADORIA"
                    ]
                ],
                use_container_width=True
            )

    st.divider()

    st.subheader("Radar do Diretor")

    alertas = gerar_alertas_financeiros(
        df_bi,
        meses_sel
    )

    if len(alertas)==0:

        st.success(
            "Nenhum alerta encontrado."
        )

    else:

        for alerta in alertas:

            if alerta["nivel"]=="critico":

                st.error(
                    f'🔴 {alerta["titulo"]}\n\n{alerta["mensagem"]}'
                )

            elif alerta["nivel"]=="atencao":

                st.warning(
                    f'🟠 {alerta["titulo"]}\n\n{alerta["mensagem"]}'
                )

            else:

                st.info(
                    f'🔵 {alerta["titulo"]}\n\n{alerta["mensagem"]}'
                )

    st.divider()

    st.subheader("🏢 Performance das CTRs")

    usar_rateio_ctr = st.toggle(
        "Considerar rateio da estrutura",
        value=False,
        key="usar_rateio_painel_executivo",
        help=(
            "Quando ativado, os custos dos centros classificados "
            "como rateio são distribuídos proporcionalmente entre as obras."
        )
    )

    df_movimentos_ctr = obter_movimentos_por_anos_meses(
        [ano_sel],
        meses_sel
    )

    if df_movimentos_ctr is None or df_movimentos_ctr.empty:
        st.info(
            "Não há movimentos suficientes para calcular "
            "a performance das CTRs."
        )

    else:
        df_rateio_ctr = carregar_logica_rateio()

        df_resultado_ctr = calcular_resultado_por_centro_custo(
            df_movimentos=df_movimentos_ctr,
            df_rateio_config=df_rateio_ctr,
            usar_rateio=usar_rateio_ctr
        )

        if df_resultado_ctr.empty:
            st.info(
                "Nenhuma CTR classificada como obra foi encontrada."
            )

        else:
            resumo_ctr = calcular_resumo_obras(
                df_resultado_ctr
            )

            c_ctr1, c_ctr2, c_ctr3, c_ctr4 = st.columns(4)

            c_ctr1.metric(
                "CTRs analisadas",
                resumo_ctr["quantidade_obras"]
            )

            c_ctr2.metric(
                "CTRs lucrativas",
                resumo_ctr["obras_lucrativas"]
            )

            c_ctr3.metric(
                "CTRs deficitárias",
                resumo_ctr["obras_deficitarias"]
            )

            c_ctr4.metric(
                "Margem da carteira",
                f'{resumo_ctr["margem"]:.2f}%'
            )

            st.caption(
                (
                    "Resultado considerando custos diretos e "
                    "rateio de estrutura."
                    if usar_rateio_ctr
                    else
                    "Resultado considerando apenas os custos diretos das obras."
                )
            )

            melhores = ranking_melhores_obras(
                df_resultado_ctr,
                quantidade=10
            )

            piores = ranking_piores_obras(
                df_resultado_ctr,
                quantidade=10
            )

            col_melhores, col_piores = st.columns(2)

            with col_melhores:
                st.write("### 🟢 Melhores CTRs")

                if melhores.empty:
                    st.info(
                        "Nenhuma CTR disponível."
                    )
                else:
                    st.dataframe(
                        melhores[
                            [
                                "Centro de Custo",
                                "Receita",
                                "Despesa Direta",
                                "Rateio Estrutura",
                                "Resultado",
                                "Margem %",
                                "Status"
                            ]
                        ]
                        .style
                        .format({
                            "Receita": formatar_moeda_br,
                            "Despesa Direta": formatar_moeda_br,
                            "Rateio Estrutura": formatar_moeda_br,
                            "Resultado": formatar_moeda_br,
                            "Margem %": "{:.2f}%"
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

            with col_piores:
                st.write("### 🔴 CTRs que exigem atenção")

                if piores.empty:
                    st.info(
                        "Nenhuma CTR disponível."
                    )
                else:
                    st.dataframe(
                        piores[
                            [
                                "Centro de Custo",
                                "Receita",
                                "Despesa Direta",
                                "Rateio Estrutura",
                                "Resultado",
                                "Margem %",
                                "Status"
                            ]
                        ]
                        .style
                        .format({
                            "Receita": formatar_moeda_br,
                            "Despesa Direta": formatar_moeda_br,
                            "Rateio Estrutura": formatar_moeda_br,
                            "Resultado": formatar_moeda_br,
                            "Margem %": "{:.2f}%"
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

            st.write("### Semáforo das CTRs")

            df_semaforo = df_resultado_ctr[
                [
                    "Centro de Custo",
                    "Resultado",
                    "Margem %",
                    "Status"
                ]
            ].copy()

            def estilo_semaforo(row):
                if row["Status"] == "Crítico":
                    return [
                        "background-color: #fee2e2"
                    ] * len(row)

                if row["Status"] == "Atenção":
                    return [
                        "background-color: #fef3c7"
                    ] * len(row)

                if row["Status"] == "Saudável":
                    return [
                        "background-color: #dcfce7"
                    ] * len(row)

                return [""] * len(row)

            st.dataframe(
                df_semaforo
                .style
                .apply(
                    estilo_semaforo,
                    axis=1
                )
                .format({
                    "Resultado": formatar_moeda_br,
                    "Margem %": "{:.2f}%"
                }),
                use_container_width=True,
                hide_index=True,
                height=500
            )
