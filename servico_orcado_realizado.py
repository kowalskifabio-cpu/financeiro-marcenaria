import pandas as pd

from servico_orcamento import MESES_NUMERO_NOME


MESES = list(MESES_NUMERO_NOME.values())


def classificar_desvio(conta, orcado, realizado):
    """
    Classificação gerencial do desvio.

    Receita:
        realizado > orçado = favorável

    Despesa:
        realizado menos negativo que orçado = favorável
    """

    conta = str(conta).strip()

    if realizado == orcado:
        return "Dentro do orçamento"

    if conta.startswith("01"):
        return (
            "Favorável"
            if realizado > orcado
            else "Desfavorável"
        )

    if conta.startswith("02"):
        return (
            "Favorável"
            if realizado > orcado
            else "Desfavorável"
        )

    return "Neutro"


def preparar_plano_contas(df_plano):
    if df_plano is None or df_plano.empty:
        return pd.DataFrame()

    df = df_plano.copy()

    df["Conta"] = (
        df["Conta"]
        .astype(str)
        .str.strip()
    )

    df["Descrição"] = (
        df["Descrição"]
        .astype(str)
        .str.strip()
    )

    df["Nivel"] = pd.to_numeric(
        df["Nivel"],
        errors="coerce"
    ).fillna(0).astype(int)

    df["Classificacao"] = (
        df["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    return df


def montar_orcado_analitico(
    df_itens,
    meses_selecionados
):
    """
    Cria uma linha por conta com valores orçados por mês.
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

    df["Mes_Nome"] = (
        df["mes"]
        .map(MESES_NUMERO_NOME)
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


def montar_realizado_analitico(
    df_bi,
    meses_selecionados
):
    """
    Extrai o realizado das contas analíticas.
    """

    if df_bi is None or df_bi.empty:
        return pd.DataFrame()

    df = df_bi.copy()

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

    df = df[
        df["Nivel"] >= 4
    ].copy()

    df["Realizado"] = df[
        meses_selecionados
    ].sum(axis=1)

    return df


def consolidar_hierarquia(
    df_base,
    colunas_valores
):
    """
    Consolida a hierarquia do plano de contas.

    Regras:
    - Nível 4 consolida no nível 3
    - Nível 3 consolida no nível 2
    - Nível 1 representa o RESULTADO
    - O RESULTADO é a soma das contas de nível 2
    """

    df = df_base.copy()

    df["Conta"] = (
        df["Conta"]
        .astype(str)
        .str.strip()
    )

    df["Nivel"] = pd.to_numeric(
        df["Nivel"],
        errors="coerce"
    ).fillna(0).astype(int)

    for coluna in colunas_valores:
        df[coluna] = pd.to_numeric(
            df[coluna],
            errors="coerce"
        ).fillna(0.0)

    # =====================================================
    # NÍVEL 4 -> NÍVEL 3
    # =====================================================
    for idx, row in df[
        df["Nivel"] == 3
    ].iterrows():

        prefixo = (
            str(row["Conta"]).strip()
            + "."
        )

        filhos = df[
            (df["Nivel"] == 4)
            &
            (
                df["Conta"]
                .astype(str)
                .str.startswith(prefixo)
            )
        ]

        if filhos.empty:
            continue

        for coluna in colunas_valores:
            df.at[idx, coluna] = (
                filhos[coluna].sum()
            )

    # =====================================================
    # NÍVEL 3 -> NÍVEL 2
    # =====================================================
    for idx, row in df[
        df["Nivel"] == 2
    ].iterrows():

        prefixo = (
            str(row["Conta"]).strip()
            + "."
        )

        filhos = df[
            (df["Nivel"] == 3)
            &
            (
                df["Conta"]
                .astype(str)
                .str.startswith(prefixo)
            )
        ]

        if filhos.empty:
            continue

        for coluna in colunas_valores:
            df.at[idx, coluna] = (
                filhos[coluna].sum()
            )

    # =====================================================
    # NÍVEL 1 -> RESULTADO
    # =====================================================
    # As despesas já são negativas.
    # Portanto:
    #
    # RECEITAS + DESPESAS = RESULTADO
    # =====================================================

    df_nivel_2 = df[
        df["Nivel"] == 2
    ].copy()

    for idx, _ in df[
        df["Nivel"] == 1
    ].iterrows():

        for coluna in colunas_valores:
            df.at[idx, coluna] = (
                df_nivel_2[coluna].sum()
            )

    return df

def montar_comparativo_gerencial(
    df_plano,
    df_orcado,
    df_realizado,
    meses_selecionados
):
    """
    Junta orçamento e realizado e consolida a hierarquia.
    """

    plano = preparar_plano_contas(
        df_plano
    )

    if plano.empty:
        return pd.DataFrame()

    comparativo = plano[
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

    col_realizado = [
        "Conta",
        "Realizado"
    ] + meses_selecionados

    if realizado.empty:
        realizado = pd.DataFrame(
            columns=col_realizado
        )
    else:
        realizado = realizado[
            col_realizado
        ].copy()

    comparativo = comparativo.merge(
        orcado,
        on="Conta",
        how="left"
    )

    comparativo = comparativo.merge(
        realizado,
        on="Conta",
        how="left",
        suffixes=(
            "_Orçado",
            "_Realizado"
        )
    )

    comparativo["Orçado"] = pd.to_numeric(
        comparativo.get(
            "Orçado",
            0
        ),
        errors="coerce"
    ).fillna(0.0)

    comparativo["Realizado"] = pd.to_numeric(
        comparativo.get(
            "Realizado",
            0
        ),
        errors="coerce"
    ).fillna(0.0)

    colunas_hierarquia = [
        "Orçado",
        "Realizado"
    ]

    comparativo = consolidar_hierarquia(
        comparativo,
        colunas_hierarquia
    )

    comparativo["Desvio R$"] = (
        comparativo["Realizado"]
        - comparativo["Orçado"]
    )

    comparativo["Desvio %"] = 0.0

    mask = (
        comparativo["Orçado"] != 0
    )

    comparativo.loc[
        mask,
        "Desvio %"
    ] = (
        comparativo.loc[
            mask,
            "Desvio R$"
        ]
        /
        comparativo.loc[
            mask,
            "Orçado"
        ].abs()
        * 100
    )

    comparativo["Status"] = (
        comparativo.apply(
            lambda row: classificar_desvio(
                row["Conta"],
                row["Orçado"],
                row["Realizado"]
            ),
            axis=1
        )
    )

    return comparativo


def calcular_forecast(
    df_itens,
    df_bi,
    meses_realizados,
    meses_futuros
):
    """
    Forecast simples e auditável:

    realizado até o mês
    +
    orçamento dos meses futuros.
    """

    realizado = 0.0
    futuro = 0.0

    if (
        df_bi is not None
        and not df_bi.empty
        and meses_realizados
    ):
        df_nivel_1 = df_bi[
            df_bi["Nivel"] == 1
        ].copy()

        for mes in meses_realizados:
            if mes in df_nivel_1.columns:
                realizado += (
                    pd.to_numeric(
                        df_nivel_1[mes],
                        errors="coerce"
                    )
                    .fillna(0.0)
                    .sum()
                )

    if (
        df_itens is not None
        and not df_itens.empty
        and meses_futuros
    ):
        df = df_itens.copy()

        df["mes"] = pd.to_numeric(
            df["mes"],
            errors="coerce"
        )

        df["valor_orcado"] = pd.to_numeric(
            df["valor_orcado"],
            errors="coerce"
        ).fillna(0.0)

        numeros_futuros = [
            numero
            for numero, nome
            in MESES_NUMERO_NOME.items()
            if nome in meses_futuros
        ]

        futuro = df[
            df["mes"].isin(
                numeros_futuros
            )
        ]["valor_orcado"].sum()

    return {
        "realizado_ate_periodo": float(
            realizado
        ),
        "orcado_futuro": float(
            futuro
        ),
        "forecast": float(
            realizado + futuro
        )
    }


def maiores_desvios(
    df_comparativo,
    quantidade=10
):
    """
    Retorna os maiores desvios analíticos por valor absoluto.
    """

    if (
        df_comparativo is None
        or df_comparativo.empty
    ):
        return pd.DataFrame()

    df = df_comparativo[
        df_comparativo["Nivel"] >= 4
    ].copy()

    df["Desvio Absoluto"] = (
        df["Desvio R$"].abs()
    )

    return (
        df[
            df["Desvio Absoluto"] > 0
        ]
        .sort_values(
            "Desvio Absoluto",
            ascending=False
        )
        .head(quantidade)
    )
