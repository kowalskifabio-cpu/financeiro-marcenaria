import io
import pandas as pd
import streamlit as st


def render_aba_resultado_operacional(
    ano_sel,
    meses_sel,
    cc_sel,
    niveis_sel,
    processar_bi,
    filtrar_linhas_zeradas,
    formatar_moeda_br
):
    st.subheader("📊 Resultado Operacional / Não Operacional")

    filtro_classificacao = st.radio(
        "Escolha a visão",
        ["operacional", "nao_operacional", "todos"],
        horizontal=True
    )

    ocultar_vazios = st.checkbox(
        "🚫 Ocultar Contas sem Movimento",
        value=False,
        key="ocultar_resultado_operacional"
    )

    if not st.button("📊 Gerar Relatório", key="btn_resultado_operacional_novo"):
        return

    df_res, meses_exibir = processar_bi(ano_sel, meses_sel, cc_sel)

    if df_res is None or df_res.empty:
        st.error("❌ Não foi possível gerar o relatório.")
        return

    df_res = df_res.copy()

    df_res["Classificacao"] = (
        df_res["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    colunas_valores = meses_exibir + ["MÉDIA", "ACUMULADO"]

    if filtro_classificacao != "todos":
        # Zera somente as linhas que NÃO pertencem à classificação escolhida
        # Mantém a linha de resultado para ser recalculada depois
        mask_manter = (
            (df_res["Classificacao"] == filtro_classificacao) |
            (df_res["Classificacao"] == "resultado")
        )

        for col in colunas_valores:
            if col in df_res.columns:
                df_res.loc[~mask_manter, col] = 0.0

        # Recalcula a árvore de baixo para cima
        for col in meses_exibir:
            niveis = sorted(df_res["Nivel"].dropna().unique(), reverse=True)

            for n in niveis:
                if n <= 1:
                    continue

                nivel_pai = n - 1

                for idx, row in df_res[df_res["Nivel"] == nivel_pai].iterrows():
                    pref = str(row["Conta"]).strip() + "."
                    total_filhos = df_res[
                        (df_res["Nivel"] == n) &
                        (df_res["Conta"].astype(str).str.startswith(pref))
                    ][col].sum()

                    filhos_existem = df_res[
                        (df_res["Nivel"] == n) &
                        (df_res["Conta"].astype(str).str.startswith(pref))
                    ]

                    if not filhos_existem.empty:
                        df_res.at[idx, col] = total_filhos

            for idx, _ in df_res[df_res["Nivel"] == 1].iterrows():
                df_res.at[idx, col] = df_res[df_res["Nivel"] == 2][col].sum()

        df_res["ACUMULADO"] = df_res[meses_exibir].sum(axis=1)
        df_res["MÉDIA"] = df_res[meses_exibir].mean(axis=1)

    if ocultar_vazios:
        df_res = filtrar_linhas_zeradas(df_res, meses_exibir + ["ACUMULADO"])

    df_visual = df_res[df_res["Nivel"].isin(niveis_sel)].copy()

    cols_export = ["Nivel", "Conta", "Descrição", "Classificacao"] + meses_exibir + ["MÉDIA", "ACUMULADO"]

    def style_rows(row):
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
        .apply(style_rows, axis=1)
        .format({
            c: formatar_moeda_br
            for c in cols_export
            if c not in ["Nivel", "Conta", "Descrição", "Classificacao"]
        }),
        use_container_width=True,
        height=800
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_visual[cols_export].to_excel(writer, index=False, sheet_name="Resultado")

    st.download_button(
        label="📥 Exportar Resultado (Excel)",
        data=buffer.getvalue(),
        file_name=f"Resultado_{filtro_classificacao}_{ano_sel}.xlsx"
    )
