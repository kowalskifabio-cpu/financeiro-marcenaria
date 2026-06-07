import io
import pandas as pd
import streamlit as st


def render_aba_resultado_operacional(
    ano_sel,
    meses_sel,
    cc_sel,
    niveis_sel,
    MAPA_MESES,
    carregar_aba_base,
    carregar_movimentos_periodo,
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

    df_base = carregar_aba_base().copy()

    if df_base.empty:
        st.error("Plano de contas não encontrado.")
        return

    meses_numeros = [MAPA_MESES[m] for m in meses_sel if m in MAPA_MESES]
    df_mov = carregar_movimentos_periodo(ano_sel, meses_numeros)

    if df_mov.empty:
        st.warning("Sem movimentos para o período selecionado.")
        return

    df_base["Classificacao"] = (
        df_base["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    contas_permitidas = set(df_base["Conta"].astype(str).str.strip())

    if filtro_classificacao != "todos":
        contas_permitidas = set(
            df_base.loc[
                df_base["Classificacao"] == filtro_classificacao,
                "Conta"
            ].astype(str).str.strip()
        )

        df_mov = df_mov[
            df_mov["Conta_ID"].astype(str).str.strip().isin(contas_permitidas)
        ].copy()

    if "Todos" not in cc_sel and cc_sel:
        df_mov = df_mov[df_mov["Centro de Custo"].isin(cc_sel)].copy()

    for mes in meses_sel:
        df_base[mes] = 0.0

    for mes in meses_sel:
        mes_num = MAPA_MESES[mes]
        df_m = df_mov[df_mov["Mes"] == mes_num].copy()

        if df_m.empty:
            continue

        df_m["Conta_ID"] = df_m["Conta_ID"].astype(str).str.strip()
        mapa_valores = df_m.groupby("Conta_ID")["Valor_Final"].sum().to_dict()

        df_base[mes] = df_base["Conta"].astype(str).str.strip().map(mapa_valores).fillna(0.0)

        for n in sorted(df_base["Nivel"].dropna().unique(), reverse=True):
            if n <= 1:
                continue

            nivel_pai = n - 1

            for idx, row in df_base[df_base["Nivel"] == nivel_pai].iterrows():
                pref = str(row["Conta"]).strip() + "."
                total_filhos = df_base[
                    (df_base["Nivel"] == n) &
                    (df_base["Conta"].astype(str).str.startswith(pref))
                ][mes].sum()

                df_base.at[idx, mes] = total_filhos

        for idx, _ in df_base[df_base["Nivel"] == 1].iterrows():
            df_base.at[idx, mes] = df_base[df_base["Nivel"] == 2][mes].sum()

    df_base["ACUMULADO"] = df_base[meses_sel].sum(axis=1)
    df_base["MÉDIA"] = df_base[meses_sel].mean(axis=1)

    if ocultar_vazios:
        df_base = filtrar_linhas_zeradas(df_base, meses_sel + ["ACUMULADO"])

    df_visual = df_base[df_base["Nivel"].isin(niveis_sel)].copy()

    cols_export = ["Nivel", "Conta", "Descrição", "Classificacao"] + meses_sel + ["MÉDIA", "ACUMULADO"]

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
