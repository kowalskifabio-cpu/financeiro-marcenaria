import pandas as pd


def _numero(valor):
    return pd.to_numeric(
        valor,
        errors="coerce"
    ).fillna(0.0)


def preparar_base_controladoria(
    df_bi,
    meses_selecionados
):
    """
    Recebe a saída consolidada do processar_bi()
    e prepara a base para indicadores gerenciais.
    """

    if df_bi is None or df_bi.empty:
        return pd.DataFrame()

    df = df_bi.copy()

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

    for mes in meses_selecionados:
        if mes not in df.columns:
            df[mes] = 0.0

        df[mes] = _numero(
            df[mes]
        )

    df["ACUMULADO_CONTROLADORIA"] = (
        df[meses_selecionados]
        .sum(axis=1)
    )

    return df


def calcular_resumo_financeiro(
    df_bi,
    meses_selecionados
):
    """
    Calcula os principais indicadores financeiros
    com base na hierarquia do plano de contas.

    Premissas:
    01 = receitas
    02 = despesas
    """

    df = preparar_base_controladoria(
        df_bi,
        meses_selecionados
    )

    if df.empty:
        return {
            "receita": 0.0,
            "despesas": 0.0,
            "resultado": 0.0,
            "margem": 0.0
        }

    # Usa nível 2 para evitar dupla contagem.
    nivel_2 = df[
        df["Nivel"] == 2
    ].copy()

    receita = (
        nivel_2[
            nivel_2["Conta"]
            .astype(str)
            .str.startswith("01")
        ]["ACUMULADO_CONTROLADORIA"]
        .sum()
    )

    despesas = (
        nivel_2[
            nivel_2["Conta"]
            .astype(str)
            .str.startswith("02")
        ]["ACUMULADO_CONTROLADORIA"]
        .sum()
    )

    resultado = receita + despesas

    margem = (
        resultado / receita * 100
        if receita != 0
        else 0.0
    )

    return {
        "receita": float(receita),
        "despesas": float(despesas),
        "resultado": float(resultado),
        "margem": float(margem)
    }


def calcular_receita_mensal(
    df_bi,
    meses_selecionados
):
    """
    Série mensal de receitas.
    """

    df = preparar_base_controladoria(
        df_bi,
        meses_selecionados
    )

    if df.empty:
        return pd.DataFrame()

    nivel_2 = df[
        df["Nivel"] == 2
    ].copy()

    registros = []

    for mes in meses_selecionados:
        valor = (
            nivel_2[
                nivel_2["Conta"]
                .astype(str)
                .str.startswith("01")
            ][mes]
            .sum()
        )

        registros.append({
            "Mês": mes,
            "Valor": float(valor)
        })

    return pd.DataFrame(registros)


def calcular_despesa_mensal(
    df_bi,
    meses_selecionados
):
    """
    Série mensal de despesas.
    """

    df = preparar_base_controladoria(
        df_bi,
        meses_selecionados
    )

    if df.empty:
        return pd.DataFrame()

    nivel_2 = df[
        df["Nivel"] == 2
    ].copy()

    registros = []

    for mes in meses_selecionados:
        valor = (
            nivel_2[
                nivel_2["Conta"]
                .astype(str)
                .str.startswith("02")
            ][mes]
            .sum()
        )

        registros.append({
            "Mês": mes,
            "Valor": float(valor)
        })

    return pd.DataFrame(registros)


def calcular_resultado_mensal(
    df_bi,
    meses_selecionados
):
    """
    Série mensal de resultado.
    """

    receitas = calcular_receita_mensal(
        df_bi,
        meses_selecionados
    )

    despesas = calcular_despesa_mensal(
        df_bi,
        meses_selecionados
    )

    if receitas.empty:
        return pd.DataFrame()

    resultado = receitas.merge(
        despesas,
        on="Mês",
        suffixes=(
            "_Receita",
            "_Despesa"
        )
    )

    resultado["Resultado"] = (
        resultado["Valor_Receita"]
        + resultado["Valor_Despesa"]
    )

    return resultado[
        [
            "Mês",
            "Resultado"
        ]
    ]


def calcular_margem_mensal(
    df_bi,
    meses_selecionados
):
    """
    Margem mensal em percentual.
    """

    receitas = calcular_receita_mensal(
        df_bi,
        meses_selecionados
    )

    resultado = calcular_resultado_mensal(
        df_bi,
        meses_selecionados
    )

    if receitas.empty:
        return pd.DataFrame()

    margem = receitas.merge(
        resultado,
        on="Mês"
    )

    margem["Margem"] = margem.apply(
        lambda row: (
            row["Resultado"]
            / row["Valor"]
            * 100
            if row["Valor"] != 0
            else 0.0
        ),
        axis=1
    )

    return margem[
        [
            "Mês",
            "Margem"
        ]
    ]


def top_contas_analiticas(
    df_bi,
    meses_selecionados,
    quantidade=10
):
    """
    Maiores contas analíticas por valor absoluto.
    """

    df = preparar_base_controladoria(
        df_bi,
        meses_selecionados
    )

    if df.empty:
        return pd.DataFrame()

    analiticas = df[
        df["Nivel"] >= 4
    ].copy()

    analiticas[
        "VALOR_ABSOLUTO"
    ] = (
        analiticas[
            "ACUMULADO_CONTROLADORIA"
        ].abs()
    )

    return (
        analiticas[
            analiticas[
                "VALOR_ABSOLUTO"
            ] > 0
        ]
        .sort_values(
            "VALOR_ABSOLUTO",
            ascending=False
        )
        .head(quantidade)
    )


def top_contas_receita(
    df_bi,
    meses_selecionados,
    quantidade=10
):
    """
    Maiores contas de receita.
    """

    df = top_contas_analiticas(
        df_bi,
        meses_selecionados,
        quantidade=999
    )

    if df.empty:
        return pd.DataFrame()

    return (
        df[
            df["Conta"]
            .astype(str)
            .str.startswith("01")
        ]
        .sort_values(
            "ACUMULADO_CONTROLADORIA",
            ascending=False
        )
        .head(quantidade)
    )


def top_contas_despesa(
    df_bi,
    meses_selecionados,
    quantidade=10
):
    """
    Maiores contas de despesa.
    """

    df = top_contas_analiticas(
        df_bi,
        meses_selecionados,
        quantidade=999
    )

    if df.empty:
        return pd.DataFrame()

    df = df[
        df["Conta"]
        .astype(str)
        .str.startswith("02")
    ].copy()

    df["ABS_DESPESA"] = (
        df[
            "ACUMULADO_CONTROLADORIA"
        ].abs()
    )

    return (
        df
        .sort_values(
            "ABS_DESPESA",
            ascending=False
        )
        .head(quantidade)
    )


def gerar_alertas_financeiros(
    df_bi,
    meses_selecionados
):
    """
    Gera alertas simples, objetivos e auditáveis.
    """

    alertas = []

    resumo = calcular_resumo_financeiro(
        df_bi,
        meses_selecionados
    )

    if resumo["resultado"] < 0:
        alertas.append({
            "nivel": "critico",
            "titulo": "Resultado negativo",
            "mensagem": (
                "O período selecionado apresenta "
                "prejuízo acumulado."
            )
        })

    if resumo["receita"] > 0:
        percentual_despesas = (
            abs(
                resumo["despesas"]
            )
            / resumo["receita"]
            * 100
        )

        if percentual_despesas > 90:
            alertas.append({
                "nivel": "atencao",
                "titulo": (
                    "Despesas consumindo grande "
                    "parte da receita"
                ),
                "mensagem": (
                    f"As despesas representam "
                    f"{percentual_despesas:.2f}% "
                    "das receitas."
                )
            })

    if resumo["margem"] < 5:
        alertas.append({
            "nivel": "atencao",
            "titulo": "Margem reduzida",
            "mensagem": (
                f"A margem acumulada está em "
                f"{resumo['margem']:.2f}%."
            )
        })

    top_despesas = top_contas_despesa(
        df_bi,
        meses_selecionados,
        quantidade=5
    )

    for _, row in top_despesas.iterrows():
        alertas.append({
            "nivel": "informacao",
            "titulo": str(
                row["Descrição"]
            ),
            "mensagem": (
                "Conta relevante no período: "
                f"{float(row['ACUMULADO_CONTROLADORIA']):.2f}"
            )
        })

    return alertas


def montar_contexto_diretoria(
    df_bi,
    meses_selecionados
):
    """
    Estrutura consolidada que poderá ser usada
    tanto pelo Painel Executivo quanto pelo Analista IA.
    """

    resumo = calcular_resumo_financeiro(
        df_bi,
        meses_selecionados
    )

    top_receitas = top_contas_receita(
        df_bi,
        meses_selecionados,
        quantidade=10
    )

    top_despesas = top_contas_despesa(
        df_bi,
        meses_selecionados,
        quantidade=10
    )

    alertas = gerar_alertas_financeiros(
        df_bi,
        meses_selecionados
    )

    def converter_linhas(df):
        if df is None or df.empty:
            return []

        registros = []

        for _, row in df.iterrows():
            registros.append({
                "conta": str(
                    row["Conta"]
                ),
                "descricao": str(
                    row["Descrição"]
                ),
                "valor": float(
                    row[
                        "ACUMULADO_CONTROLADORIA"
                    ]
                )
            })

        return registros

    return {
        "resumo": resumo,
        "top_receitas": converter_linhas(
            top_receitas
        ),
        "top_despesas": converter_linhas(
            top_despesas
        ),
        "alertas": alertas
    }
