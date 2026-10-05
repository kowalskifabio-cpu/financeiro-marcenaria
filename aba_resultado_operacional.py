import io

import pandas as pd
import streamlit as st

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


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
    st.subheader("📊 Resultado por Classificação")

    filtro_classificacao = st.radio(
        "Escolha a visão",
        [
            "operacional",
            "nao_operacional",
            "diretoria",
            "diretoria_investimentos",
            "todos"
        ],
        format_func=lambda valor: {
            "operacional": "Operacional",
            "nao_operacional": "Não Operacional",
            "diretoria": "Diretoria",
            "diretoria_investimentos": "Diretoria Investimentos",
            "todos": "Todos"
        }[valor],
        horizontal=True
    )

    ocultar_vazios = st.checkbox(
        "🚫 Ocultar Contas sem Movimento",
        value=True,
        key="ocultar_resultado_operacional"
    )

    if not st.button(
        "📊 Gerar Relatório",
        key="btn_resultado_operacional_novo"
    ):
        return

    # ==========================================================
    # 1. CARREGAMENTO
    # ==========================================================

    df_base = carregar_aba_base().copy()

    meses_numeros = [
        MAPA_MESES[m]
        for m in meses_sel
        if m in MAPA_MESES
    ]

    df_mov = carregar_movimentos_periodo(
        ano_sel,
        meses_numeros
    )

    if df_base.empty or df_mov.empty:
        st.warning("Sem dados para gerar o relatório.")
        return

    # ==========================================================
    # 2. PREPARAÇÃO DO PLANO DE CONTAS
    # ==========================================================

    df_base["Conta"] = (
        df_base["Conta"]
        .astype(str)
        .str.strip()
    )

    df_base["Classificacao"] = (
        df_base["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    mapa_class = dict(
        zip(
            df_base["Conta"],
            df_base["Classificacao"]
        )
    )

    # ==========================================================
    # 3. PREPARAÇÃO DOS MOVIMENTOS
    # ==========================================================

    df_mov["Conta_ID"] = (
        df_mov["Conta_ID"]
        .astype(str)
        .str.strip()
    )

    df_mov["Valor_Final"] = pd.to_numeric(
        df_mov["Valor_Final"],
        errors="coerce"
    ).fillna(0.0)

    def classificar_movimento(conta):
        conta = str(conta).strip()

        if conta in mapa_class:
            return mapa_class[conta]

        partes = conta.split(".")

        while len(partes) > 1:
            partes = partes[:-1]
            pai = ".".join(partes)

            if pai in mapa_class:
                return mapa_class[pai]

        return "fora_resultado"

    df_mov["Classificacao"] = (
        df_mov["Conta_ID"]
        .apply(classificar_movimento)
    )

    # ==========================================================
    # 4. FILTRO DE CENTRO DE CUSTO
    #
    # IMPORTANTE:
    # aplicamos o CC antes de gerar o resumo.
    # Assim resumo e relatório detalhado usam o mesmo universo.
    # ==========================================================

    if "Todos" not in cc_sel and cc_sel:
        df_mov = df_mov[
            df_mov["Centro de Custo"].isin(cc_sel)
        ].copy()

    # ==========================================================
    # 5. RESUMO EXECUTIVO POR CLASSIFICAÇÃO
    #
    # ESTE RESUMO NÃO OBEDECE AO FILTRO DE CLASSIFICAÇÃO.
    # SEMPRE MOSTRA AS QUATRO CLASSIFICAÇÕES + TOTAL.
    # ==========================================================

    classificacoes_resumo = [
        ("operacional", "Operacional"),
        ("nao_operacional", "Não Operacional"),
        ("diretoria", "Diretoria"),
        (
            "diretoria_investimentos",
            "Diretoria Investimentos"
        )
    ]

    linhas_resumo = []

    for codigo, descricao in classificacoes_resumo:

        linha = {
            "Descrição": descricao
        }

        df_class = df_mov[
            df_mov["Classificacao"] == codigo
        ].copy()

        for mes in meses_sel:

            mes_num = int(
                MAPA_MESES[mes]
            )

            if df_class.empty:
                valor_mes = 0.0
            else:
                valor_mes = (
                    df_class[
                        df_class["Mes"].astype(int)
                        == mes_num
                    ]["Valor_Final"]
                    .sum()
                )

            linha[mes] = float(valor_mes)

        linha["ACUMULADO"] = sum(
            linha[mes]
            for mes in meses_sel
        )

        if meses_sel:
            linha["MÉDIA"] = (
                linha["ACUMULADO"]
                / len(meses_sel)
            )
        else:
            linha["MÉDIA"] = 0.0

        linhas_resumo.append(linha)

    df_resumo = pd.DataFrame(
        linhas_resumo
    )

    # ==========================================================
    # 6. LINHA TOTAL DO RESUMO
    # ==========================================================

    linha_total = {
        "Descrição": "Total"
    }

    for mes in meses_sel:
        linha_total[mes] = (
            df_resumo[mes].sum()
        )

    linha_total["ACUMULADO"] = (
        df_resumo["ACUMULADO"].sum()
    )

    if meses_sel:
        linha_total["MÉDIA"] = (
            linha_total["ACUMULADO"]
            / len(meses_sel)
        )
    else:
        linha_total["MÉDIA"] = 0.0

    df_resumo = pd.concat(
        [
            df_resumo,
            pd.DataFrame([linha_total])
        ],
        ignore_index=True
    )

    # Ordem visual igual ao modelo solicitado
    cols_resumo = (
        ["Descrição"]
        + meses_sel
        + ["MÉDIA", "ACUMULADO"]
    )

    df_resumo = df_resumo[
        cols_resumo
    ].copy()

    # ==========================================================
    # 7. EXIBIÇÃO DO RESUMO
    # ==========================================================

    st.markdown(
        "### 📌 Resumo por Classificação"
    )

    st.caption(
        "Visão consolidada de todas as classificações, "
        "independentemente da visão selecionada abaixo."
    )

    def style_resumo(row):

        if row["Descrição"] == "Total":
            return [
                (
                    "background-color: #334155; "
                    "color: white; "
                    "font-weight: bold"
                )
            ] * len(row)

        return [
            (
                "background-color: #D1EAFF; "
                "color: black"
            )
        ] * len(row)

    st.dataframe(
        df_resumo
        .style
        .apply(
            style_resumo,
            axis=1
        )
        .format({
            c: formatar_moeda_br
            for c in cols_resumo
            if c != "Descrição"
        }),
        use_container_width=True,
        hide_index=True
    )
    
    # ==========================================================
    # EXPORTAÇÃO EXCLUSIVA DO RESUMO POR CLASSIFICAÇÃO
    # ==========================================================

    buffer_resumo = io.BytesIO()

    with pd.ExcelWriter(
        buffer_resumo,
        engine="openpyxl"
    ) as writer:

        df_resumo.to_excel(
            writer,
            index=False,
            sheet_name="Resumo por Classificacao"
        )

        ws_resumo = writer.sheets[
            "Resumo por Classificacao"
        ]

        # Cores
        cor_escura = "334155"
        cor_azul_claro = "D1EAFF"
        cor_cabecalho = "E2E8F0"
        cor_branca = "FFFFFF"

        borda_fina = Side(
            style="thin",
            color="D1D5DB"
        )

        # Cabeçalho
        for cell in ws_resumo[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor=cor_cabecalho
            )

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

            cell.border = Border(
                bottom=borda_fina
            )

        # Linhas do resumo
        for row in range(
            2,
            ws_resumo.max_row + 1
        ):

            descricao = ws_resumo.cell(
                row=row,
                column=1
            ).value

            if descricao == "Total":

                fill = PatternFill(
                    "solid",
                    fgColor=cor_escura
                )

                font = Font(
                    bold=True,
                    color=cor_branca
                )

            else:

                fill = PatternFill(
                    "solid",
                    fgColor=cor_azul_claro
                )

                font = Font(
                    color="000000"
                )

            for col in range(
                1,
                ws_resumo.max_column + 1
            ):

                cell = ws_resumo.cell(
                    row=row,
                    column=col
                )

                cell.fill = fill
                cell.font = font

                if col > 1:
                    cell.number_format = (
                        'R$ #,##0.00;'
                        '[Red]-R$ #,##0.00'
                    )

        # Congela primeira coluna e cabeçalho
        ws_resumo.freeze_panes = "B2"

        # Filtro
        ws_resumo.auto_filter.ref = (
            ws_resumo.dimensions
        )

        # Largura das colunas
        ws_resumo.column_dimensions[
            "A"
        ].width = 28

        for col in range(
            2,
            ws_resumo.max_column + 1
        ):
            ws_resumo.column_dimensions[
                get_column_letter(col)
            ].width = 16

    st.download_button(
        "📥 Exportar Resumo por Classificação (Excel)",
        data=buffer_resumo.getvalue(),
        file_name=f"Resumo_Classificacao_{ano_sel}.xlsx",
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="download_resumo_classificacao"
    )

    st.divider()
  
    # ==========================================================
    # 8. A PARTIR DAQUI:
    # O FILTRO DE CLASSIFICAÇÃO AFETA SOMENTE O RELATÓRIO
    # DETALHADO.
    # ==========================================================

    df_mov_detalhado = df_mov.copy()

    if filtro_classificacao != "todos":
        df_mov_detalhado = (
            df_mov_detalhado[
                df_mov_detalhado["Classificacao"]
                == filtro_classificacao
            ]
            .copy()
        )

    # ==========================================================
    # 9. MONTA RELATÓRIO DETALHADO
    # ==========================================================

    for mes in meses_sel:
        df_base[mes] = 0.0

    for mes in meses_sel:

        mes_num = int(
            MAPA_MESES[mes]
        )

        df_m = df_mov_detalhado[
            df_mov_detalhado["Mes"].astype(int)
            == mes_num
        ].copy()

        if df_m.empty:
            continue

        mapa_valores = (
            df_m
            .groupby("Conta_ID")["Valor_Final"]
            .sum()
            .to_dict()
        )

        df_base[mes] = (
            df_base["Conta"]
            .map(mapa_valores)
            .fillna(0.0)
        )

        # ------------------------------------------------------
        # CONSOLIDAÇÃO HIERÁRQUICA
        # ------------------------------------------------------

        for n in sorted(
            df_base["Nivel"]
            .dropna()
            .unique(),
            reverse=True
        ):

            if n <= 1:
                continue

            nivel_pai = n - 1

            for idx, row in df_base[
                df_base["Nivel"] == nivel_pai
            ].iterrows():

                pref = (
                    str(row["Conta"]).strip()
                    + "."
                )

                filhos = df_base[
                    (df_base["Nivel"] == n)
                    &
                    (
                        df_base["Conta"]
                        .astype(str)
                        .str.startswith(pref)
                    )
                ]

                total_filhos = (
                    filhos[mes].sum()
                )

                if total_filhos != 0:
                    df_base.at[
                        idx,
                        mes
                    ] = total_filhos

        # ------------------------------------------------------
        # RESULTADO NÍVEL 1
        # ------------------------------------------------------

        for idx, _ in df_base[
            df_base["Nivel"] == 1
        ].iterrows():

            df_base.at[
                idx,
                mes
            ] = (
                df_base[
                    df_base["Nivel"] == 2
                ][mes]
                .sum()
            )

    # ==========================================================
    # 10. MÉDIA E ACUMULADO
    # ==========================================================

    df_base["ACUMULADO"] = (
        df_base[meses_sel]
        .sum(axis=1)
    )

    df_base["MÉDIA"] = (
        df_base[meses_sel]
        .mean(axis=1)
    )

    # ==========================================================
    # 11. OCULTAR CONTAS ZERADAS
    # ==========================================================

    if ocultar_vazios:
        df_base = filtrar_linhas_zeradas(
            df_base,
            meses_sel + ["ACUMULADO"]
        )

    # ==========================================================
    # 12. NÍVEIS SELECIONADOS
    # ==========================================================

    df_visual = df_base[
        df_base["Nivel"].isin(
            niveis_sel
        )
    ].copy()

    cols_export = [
        "Nivel",
        "Conta",
        "Descrição",
        "Classificacao"
    ] + meses_sel + [
        "MÉDIA",
        "ACUMULADO"
    ]

    # ==========================================================
    # 13. ESTILO DO RELATÓRIO DETALHADO
    # ==========================================================

    def style_rows(row):

        if row["Nivel"] == 1:
            return [
                (
                    "background-color: #334155; "
                    "color: white; "
                    "font-weight: bold"
                )
            ] * len(row)

        if row["Nivel"] == 2:
            return [
                (
                    "background-color: #cbd5e1; "
                    "font-weight: bold; "
                    "color: black"
                )
            ] * len(row)

        if row["Nivel"] == 3:
            return [
                (
                    "background-color: #D1EAFF; "
                    "font-weight: bold; "
                    "color: black"
                )
            ] * len(row)

        return [""] * len(row)

    # ==========================================================
    # 14. EXIBIÇÃO DO RELATÓRIO DETALHADO
    # ==========================================================

    st.markdown(
        "### 📊 Resultado Detalhado"
    )

    st.dataframe(
        df_visual[cols_export]
        .style
        .apply(
            style_rows,
            axis=1
        )
        .format({
            c: formatar_moeda_br
            for c in cols_export
            if c not in [
                "Nivel",
                "Conta",
                "Descrição",
                "Classificacao"
            ]
        }),
        use_container_width=True,
        height=800
    )

    # ==========================================================
    # 15. NOME DO ARQUIVO
    # ==========================================================

    nome_tipo = {
        "operacional": "Operacional",
        "nao_operacional": "Nao_Operacional",
        "diretoria": "Diretoria",
        "diretoria_investimentos":
            "Diretoria_Investimentos",
        "todos": "Todos"
    }[filtro_classificacao]

    if (
        "Todos" in cc_sel
        or len(cc_sel) == 0
    ):
        sufixo = ""

    elif len(cc_sel) == 1:

        ctr = (
            cc_sel[0]
            .split("-")[0]
            .replace("/", "-")
        )

        sufixo = f"_{ctr}"

    else:
        sufixo = "_Filtrado"

    nome_arquivo = (
        f"Resultado_{nome_tipo}_"
        f"{ano_sel}{sufixo}.xlsx"
    )

    # ==========================================================
    # 16. EXPORTAÇÃO PARA EXCEL
    # ==========================================================

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        # ------------------------------------------------------
        # ABA 1 - RESUMO
        # ------------------------------------------------------

        df_resumo.to_excel(
            writer,
            index=False,
            sheet_name="Resumo por Classificacao"
        )

        ws_resumo = writer.sheets[
            "Resumo por Classificacao"
        ]

        # ------------------------------------------------------
        # ABA 2 - DETALHADO
        # ------------------------------------------------------

        df_export = (
            df_visual[
                cols_export
            ].copy()
        )

        df_export.to_excel(
            writer,
            index=False,
            sheet_name="Resultado Detalhado"
        )

        ws_detalhado = writer.sheets[
            "Resultado Detalhado"
        ]

        # ======================================================
        # 17. ESTILOS DO EXCEL
        # ======================================================

        cor_escura = "334155"
        cor_nivel_2 = "CBD5E1"
        cor_azul_claro = "D1EAFF"
        cor_cabecalho = "E2E8F0"
        cor_branca = "FFFFFF"

        borda_fina = Side(
            style="thin",
            color="D1D5DB"
        )

        # ------------------------------------------------------
        # FORMATA RESUMO
        # ------------------------------------------------------

        ws_resumo.freeze_panes = "B2"
        ws_resumo.auto_filter.ref = (
            ws_resumo.dimensions
        )

        for cell in ws_resumo[1]:

            cell.fill = PatternFill(
                "solid",
                fgColor=cor_cabecalho
            )

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

            cell.border = Border(
                bottom=borda_fina
            )

        for row in range(
            2,
            ws_resumo.max_row + 1
        ):

            descricao = (
                ws_resumo.cell(
                    row=row,
                    column=1
                ).value
            )

            if descricao == "Total":

                fill = PatternFill(
                    "solid",
                    fgColor=cor_escura
                )

                font = Font(
                    bold=True,
                    color=cor_branca
                )

            else:

                fill = PatternFill(
                    "solid",
                    fgColor=cor_azul_claro
                )

                font = Font(
                    color="000000"
                )

            for col in range(
                1,
                ws_resumo.max_column + 1
            ):

                cell = ws_resumo.cell(
                    row=row,
                    column=col
                )

                cell.fill = fill
                cell.font = font

                if col > 1:
                    cell.number_format = (
                        'R$ #,##0.00;'
                        '[Red]-R$ #,##0.00'
                    )

        # ------------------------------------------------------
        # LARGURA DO RESUMO
        # ------------------------------------------------------

        ws_resumo.column_dimensions[
            "A"
        ].width = 28

        for col in range(
            2,
            ws_resumo.max_column + 1
        ):
            ws_resumo.column_dimensions[
                get_column_letter(col)
            ].width = 16

        # ------------------------------------------------------
        # FORMATA DETALHADO
        # ------------------------------------------------------

        ws_detalhado.freeze_panes = "A2"
        ws_detalhado.auto_filter.ref = (
            ws_detalhado.dimensions
        )

        for cell in ws_detalhado[1]:

            cell.fill = PatternFill(
                "solid",
                fgColor=cor_cabecalho
            )

            cell.font = Font(
                bold=True
            )

            cell.border = Border(
                bottom=borda_fina
            )

        for row in range(
            2,
            ws_detalhado.max_row + 1
        ):

            nivel = (
                ws_detalhado.cell(
                    row=row,
                    column=1
                ).value
            )

            if nivel == 1:

                fill = PatternFill(
                    "solid",
                    fgColor=cor_escura
                )

                font = Font(
                    bold=True,
                    color=cor_branca
                )

            elif nivel == 2:

                fill = PatternFill(
                    "solid",
                    fgColor=cor_nivel_2
                )

                font = Font(
                    bold=True,
                    color="000000"
                )

            elif nivel == 3:

                fill = PatternFill(
                    "solid",
                    fgColor=cor_azul_claro
                )

                font = Font(
                    bold=True,
                    color="000000"
                )

            else:

                fill = PatternFill(
                    fill_type=None
                )

                font = Font(
                    color="000000"
                )

            for col in range(
                1,
                ws_detalhado.max_column + 1
            ):

                cell = ws_detalhado.cell(
                    row=row,
                    column=col
                )

                cell.fill = fill
                cell.font = font

                if col >= 5:
                    cell.number_format = (
                        'R$ #,##0.00;'
                        '[Red]-R$ #,##0.00'
                    )

        # ------------------------------------------------------
        # LARGURA AUTOMÁTICA DO DETALHADO
        # ------------------------------------------------------

        for coluna in ws_detalhado.columns:

            tamanho = 0
            letra = (
                coluna[0]
                .column_letter
            )

            for celula in coluna:

                try:
                    tamanho = max(
                        tamanho,
                        len(
                            str(
                                celula.value
                                if celula.value
                                is not None
                                else ""
                            )
                        )
                    )

                except Exception:
                    pass

            ws_detalhado.column_dimensions[
                letra
            ].width = min(
                tamanho + 3,
                45
            )

    # ==========================================================
    # 18. DOWNLOAD
    # ==========================================================

    st.download_button(
        "📥 Exportar Resultado Completo (Excel)",
        data=buffer.getvalue(),
        file_name=nome_arquivo,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
