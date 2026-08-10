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
    formatar_moeda_br
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
