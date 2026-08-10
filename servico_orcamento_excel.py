import io

import pandas as pd

from servico_orcamento import (
    MESES_NUMERO_NOME,
    MESES_NOME_NUMERO,
)


COLUNAS_MESES = list(MESES_NOME_NUMERO.keys())


def _normalizar_conta(valor):
    return str(valor).strip()


def _normalizar_colunas(df):
    df = df.copy()

    mapa = {
        str(col).strip(): col
        for col in df.columns
    }

    renomear = {}

    for nome_esperado in (
        ["Conta", "Descrição", "Nivel", "Classificacao"]
        + COLUNAS_MESES
        + ["Total Anual"]
    ):
        if nome_esperado in mapa:
            renomear[mapa[nome_esperado]] = nome_esperado

    return df.rename(columns=renomear)


def gerar_modelo_orcamento_excel(
    df_plano_contas,
    df_itens_existentes=None,
    ano=None,
    versao=None
):
    """
    Gera o modelo Excel do orçamento a partir do plano de contas atual.

    Somente contas analíticas (Nivel >= 4) entram no modelo.

    Se houver itens existentes, os valores já salvos serão preenchidos
    automaticamente no arquivo gerado.
    """

    if df_plano_contas is None or df_plano_contas.empty:
        raise ValueError(
            "Plano de contas não encontrado."
        )

    plano = df_plano_contas.copy()

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

    plano["Nivel"] = pd.to_numeric(
        plano["Nivel"],
        errors="coerce"
    ).fillna(0).astype(int)

    if "Classificacao" not in plano.columns:
        plano["Classificacao"] = "operacional"

    plano["Classificacao"] = (
        plano["Classificacao"]
        .fillna("operacional")
        .astype(str)
        .str.strip()
    )

    grade = plano[
        plano["Nivel"] >= 4
    ][
        [
            "Conta",
            "Descrição",
            "Nivel",
            "Classificacao"
        ]
    ].copy()

    for mes in COLUNAS_MESES:
        grade[mes] = pd.NA

    if (
        df_itens_existentes is not None
        and not df_itens_existentes.empty
    ):
        itens = df_itens_existentes.copy()

        itens["conta_id"] = (
            itens["conta_id"]
            .astype(str)
            .str.strip()
        )

        itens["mes"] = pd.to_numeric(
            itens["mes"],
            errors="coerce"
        )

        itens["valor_orcado"] = pd.to_numeric(
            itens["valor_orcado"],
            errors="coerce"
        )

        itens = itens.dropna(
            subset=[
                "mes",
                "valor_orcado"
            ]
        ).copy()

        itens["mes"] = itens["mes"].astype(int)

        tabela = itens.pivot_table(
            index="conta_id",
            columns="mes",
            values="valor_orcado",
            aggfunc="sum"
        )

        for numero_mes, nome_mes in MESES_NUMERO_NOME.items():
            if numero_mes in tabela.columns:
                mapa_valores = tabela[
                    numero_mes
                ].to_dict()

                grade[nome_mes] = (
                    grade["Conta"]
                    .map(mapa_valores)
                )

    totais = []

    for _, row in grade.iterrows():
        total = 0.0

        for mes in COLUNAS_MESES:
            valor = pd.to_numeric(
                row[mes],
                errors="coerce"
            )

            if not pd.isna(valor):
                total += float(valor)

        totais.append(total)

    grade["Total Anual"] = totais

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        grade.to_excel(
            writer,
            index=False,
            sheet_name="Orçamento"
        )

        ws = writer.sheets["Orçamento"]

        ws.freeze_panes = "E2"
        ws.auto_filter.ref = ws.dimensions

        colunas = list(grade.columns)

        for nome in COLUNAS_MESES + ["Total Anual"]:
            coluna_numero = (
                colunas.index(nome) + 1
            )

            for linha in range(
                2,
                ws.max_row + 1
            ):
                ws.cell(
                    row=linha,
                    column=coluna_numero
                ).number_format = (
                    'R$ #,##0.00;[Red]-R$ #,##0.00'
                )

        larguras = {
            "A": 16,
            "B": 42,
            "C": 10,
            "D": 24,
        }

        for letra, largura in larguras.items():
            ws.column_dimensions[
                letra
            ].width = largura

        for coluna in ws.iter_cols(
            min_col=5,
            max_col=ws.max_column
        ):
            letra = coluna[0].column_letter
            ws.column_dimensions[
                letra
            ].width = 15

        ws.sheet_view.showGridLines = True

        # Aba de instruções
        instrucoes = writer.book.create_sheet(
            "Instruções"
        )

        instrucoes["A1"] = "MODELO DE ORÇAMENTO BASE ZERO"
        instrucoes["A3"] = (
            "Preencha apenas as colunas de Janeiro a Dezembro."
        )
        instrucoes["A4"] = (
            "Não altere os códigos das contas."
        )
        instrucoes["A5"] = (
            "Célula vazia = não alterar valor já existente."
        )
        instrucoes["A6"] = (
            "Valor 0 = zerar deliberadamente aquele mês."
        )
        instrucoes["A7"] = (
            "Receitas devem ser positivas."
        )
        instrucoes["A8"] = (
            "Despesas devem ser negativas."
        )

        if ano is not None:
            instrucoes["A10"] = f"Ano do orçamento: {ano}"

        if versao is not None:
            instrucoes["A11"] = (
                f"Versão do orçamento: {versao}"
            )

        instrucoes.column_dimensions["A"].width = 70

    nome = "Modelo_Orcamento_OBZ"

    if ano is not None:
        nome += f"_{ano}"

    if versao is not None:
        nome += f"_Versao_{versao}"

    nome += ".xlsx"

    return buffer.getvalue(), nome


def ler_excel_orcamento(
    arquivo_excel
):
    """
    Lê a planilha Orçamento do arquivo enviado.
    """

    try:
        df = pd.read_excel(
            arquivo_excel,
            sheet_name="Orçamento",
            dtype={
                "Conta": str
            }
        )
    except ValueError:
        raise ValueError(
            "O arquivo precisa conter uma aba chamada 'Orçamento'."
        )
    except Exception as erro:
        raise ValueError(
            "Não foi possível ler o Excel: "
            f"{type(erro).__name__} — {erro}"
        )

    df = _normalizar_colunas(
        df
    )

    return df


def validar_excel_orcamento(
    df_excel,
    df_plano_contas
):
    """
    Valida estrutura, contas e conteúdos do Excel.

    Retorna:
        {
            "valido": bool,
            "erros": [...],
            "avisos": [...],
            "contas_reconhecidas": int,
            "contas_nao_encontradas": [...],
            "celulas_invalidas": [...],
            "df_validado": DataFrame
        }
    """

    erros = []
    avisos = []
    contas_nao_encontradas = []
    celulas_invalidas = []

    if df_excel is None or df_excel.empty:
        return {
            "valido": False,
            "erros": [
                "A planilha está vazia."
            ],
            "avisos": [],
            "contas_reconhecidas": 0,
            "contas_nao_encontradas": [],
            "celulas_invalidas": [],
            "df_validado": pd.DataFrame()
        }

    df = _normalizar_colunas(
        df_excel
    ).copy()

    obrigatorias = [
        "Conta",
        "Descrição"
    ] + COLUNAS_MESES

    faltantes = [
        coluna
        for coluna in obrigatorias
        if coluna not in df.columns
    ]

    if faltantes:
        erros.append(
            "Colunas obrigatórias ausentes: "
            + ", ".join(faltantes)
        )

        return {
            "valido": False,
            "erros": erros,
            "avisos": avisos,
            "contas_reconhecidas": 0,
            "contas_nao_encontradas": [],
            "celulas_invalidas": [],
            "df_validado": df
        }

    plano = df_plano_contas.copy()

    plano["Conta"] = (
        plano["Conta"]
        .astype(str)
        .str.strip()
    )

    plano["Nivel"] = pd.to_numeric(
        plano["Nivel"],
        errors="coerce"
    ).fillna(0).astype(int)

    plano_analitico = plano[
        plano["Nivel"] >= 4
    ].copy()

    contas_validas = set(
        plano_analitico["Conta"].tolist()
    )

    df["Conta"] = (
        df["Conta"]
        .astype(str)
        .str.strip()
    )

    # Remove linhas completamente vazias de conta
    df = df[
        df["Conta"].ne("")
        &
        df["Conta"].ne("nan")
    ].copy()

    duplicadas = (
        df[
            df["Conta"].duplicated(
                keep=False
            )
        ]["Conta"]
        .unique()
        .tolist()
    )

    if duplicadas:
        erros.append(
            "Existem contas duplicadas no arquivo: "
            + ", ".join(
                str(x)
                for x in duplicadas[:20]
            )
        )

    for conta in df["Conta"]:
        if conta not in contas_validas:
            contas_nao_encontradas.append(
                conta
            )

    contas_nao_encontradas = sorted(
        set(
            contas_nao_encontradas
        )
    )

    if contas_nao_encontradas:
        erros.append(
            f"{len(contas_nao_encontradas)} conta(s) "
            "não foram encontradas no plano de contas."
        )

    for indice, row in df.iterrows():
        conta = row["Conta"]

        for mes in COLUNAS_MESES:
            valor_original = row[mes]

            if pd.isna(
                valor_original
            ):
                continue

            if (
                isinstance(
                    valor_original,
                    str
                )
                and not valor_original.strip()
            ):
                continue

            valor_num = pd.to_numeric(
                valor_original,
                errors="coerce"
            )

            if pd.isna(
                valor_num
            ):
                celulas_invalidas.append({
                    "linha_excel": int(
                        indice + 2
                    ),
                    "conta": conta,
                    "mes": mes,
                    "valor": str(
                        valor_original
                    )
                })

    if celulas_invalidas:
        erros.append(
            f"{len(celulas_invalidas)} célula(s) "
            "possuem conteúdo não numérico."
        )

    contas_reconhecidas = len(
        [
            conta
            for conta in df["Conta"]
            if conta in contas_validas
        ]
    )

    if "Total Anual" in df.columns:
        avisos.append(
            "A coluna Total Anual é apenas informativa "
            "e será recalculada pelo sistema."
        )

    return {
        "valido": len(erros) == 0,
        "erros": erros,
        "avisos": avisos,
        "contas_reconhecidas": int(
            contas_reconhecidas
        ),
        "contas_nao_encontradas": (
            contas_nao_encontradas
        ),
        "celulas_invalidas": (
            celulas_invalidas
        ),
        "df_validado": df
    }


def gerar_previa_importacao(
    df_validado,
    df_itens_existentes=None
):
    """
    Gera métricas da importação antes de gravar.
    """

    if (
        df_validado is None
        or df_validado.empty
    ):
        return {
            "contas": 0,
            "celulas_preenchidas": 0,
            "receitas": 0.0,
            "despesas": 0.0,
            "resultado": 0.0,
            "margem": 0.0,
            "alteracoes": 0
        }

    df = df_validado.copy()

    existentes = {}

    if (
        df_itens_existentes is not None
        and not df_itens_existentes.empty
    ):
        itens = df_itens_existentes.copy()

        itens["conta_id"] = (
            itens["conta_id"]
            .astype(str)
            .str.strip()
        )

        itens["mes"] = pd.to_numeric(
            itens["mes"],
            errors="coerce"
        )

        itens["valor_orcado"] = pd.to_numeric(
            itens["valor_orcado"],
            errors="coerce"
        ).fillna(0.0)

        itens = itens.dropna(
            subset=["mes"]
        ).copy()

        itens["mes"] = itens["mes"].astype(int)

        for _, item in itens.iterrows():
            existentes[
                (
                    item["conta_id"],
                    int(item["mes"])
                )
            ] = float(
                item["valor_orcado"]
            )

    celulas_preenchidas = 0
    alteracoes = 0
    receitas = 0.0
    despesas = 0.0

    for _, row in df.iterrows():
        conta = _normalizar_conta(
            row["Conta"]
        )

        for mes_nome, mes_num in MESES_NOME_NUMERO.items():
            valor = row[mes_nome]

            if pd.isna(valor):
                continue

            if (
                isinstance(valor, str)
                and not valor.strip()
            ):
                continue

            valor_num = pd.to_numeric(
                valor,
                errors="coerce"
            )

            if pd.isna(valor_num):
                continue

            valor_num = float(
                valor_num
            )

            celulas_preenchidas += 1

            valor_existente = existentes.get(
                (
                    conta,
                    int(mes_num)
                )
            )

            if (
                valor_existente is None
                or float(valor_existente)
                != valor_num
            ):
                alteracoes += 1

            if conta.startswith("01"):
                receitas += valor_num

            elif conta.startswith("02"):
                despesas += valor_num

    resultado = receitas + despesas

    margem = (
        resultado / receitas * 100
        if receitas != 0
        else 0.0
    )

    return {
        "contas": int(
            len(df)
        ),
        "celulas_preenchidas": int(
            celulas_preenchidas
        ),
        "receitas": float(
            receitas
        ),
        "despesas": float(
            despesas
        ),
        "resultado": float(
            resultado
        ),
        "margem": float(
            margem
        ),
        "alteracoes": int(
            alteracoes
        )
    }


def _montar_registros_atualizacao(
    df_validado,
    orcamento_id,
    justificativa,
    responsavel
):
    """
    Modo incremental:
    somente células preenchidas viram registros.
    """

    registros = []

    for _, row in df_validado.iterrows():
        conta = _normalizar_conta(
            row["Conta"]
        )

        for mes_nome, mes_num in MESES_NOME_NUMERO.items():
            valor = row[mes_nome]

            if pd.isna(valor):
                continue

            if (
                isinstance(valor, str)
                and not valor.strip()
            ):
                continue

            valor_num = pd.to_numeric(
                valor,
                errors="coerce"
            )

            if pd.isna(valor_num):
                continue

            registros.append({
                "orcamento_id": int(
                    orcamento_id
                ),
                "conta_id": conta,
                "mes": int(
                    mes_num
                ),
                "valor_orcado": float(
                    valor_num
                ),
                "justificativa": (
                    justificativa
                ),
                "responsavel": (
                    responsavel
                ),
                "criado_por": (
                    "Administrador Master"
                )
            })

    return registros


def _montar_registros_substituicao(
    df_validado,
    orcamento_id,
    justificativa,
    responsavel
):
    """
    Modo substituição:
    todas as contas e meses são gravados.
    Célula vazia vira zero.
    """

    registros = []

    for _, row in df_validado.iterrows():
        conta = _normalizar_conta(
            row["Conta"]
        )

        for mes_nome, mes_num in MESES_NOME_NUMERO.items():
            valor = row[mes_nome]

            valor_num = pd.to_numeric(
                valor,
                errors="coerce"
            )

            if pd.isna(
                valor_num
            ):
                valor_num = 0.0

            registros.append({
                "orcamento_id": int(
                    orcamento_id
                ),
                "conta_id": conta,
                "mes": int(
                    mes_num
                ),
                "valor_orcado": float(
                    valor_num
                ),
                "justificativa": (
                    justificativa
                ),
                "responsavel": (
                    responsavel
                ),
                "criado_por": (
                    "Administrador Master"
                )
            })

    return registros


def importar_orcamento_excel(
    supabase_client,
    orcamento_id,
    df_validado,
    modo,
    justificativa,
    responsavel,
    status_orcamento
):
    """
    Grava o Excel no Supabase.

    Modos:
        atualizar
        substituir

    atualizar:
        apenas células preenchidas são gravadas.

    substituir:
        regrava todos os meses das contas do modelo,
        com vazios convertidos para zero.
        Só é permitido em orçamento rascunho.
    """

    justificativa = str(
        justificativa
    ).strip()

    responsavel = str(
        responsavel
    ).strip()

    if len(justificativa) < 3:
        raise ValueError(
            "Informe uma justificativa com ao menos "
            "três caracteres."
        )

    if not responsavel:
        responsavel = (
            "Administrador Master"
        )

    status = str(
        status_orcamento
    ).strip().lower()

    if status not in [
        "rascunho",
        "em_revisao"
    ]:
        raise ValueError(
            "Este orçamento não permite importação porque "
            f"está com status '{status}'."
        )

    if modo not in [
        "atualizar",
        "substituir"
    ]:
        raise ValueError(
            "Modo de importação inválido."
        )

    if (
        modo == "substituir"
        and status != "rascunho"
    ):
        raise ValueError(
            "A substituição integral só é permitida "
            "enquanto o orçamento estiver em rascunho."
        )

    if modo == "atualizar":
        registros = (
            _montar_registros_atualizacao(
                df_validado=df_validado,
                orcamento_id=orcamento_id,
                justificativa=justificativa,
                responsavel=responsavel
            )
        )

        acao_historico = (
            "Importação incremental via Excel"
        )

    else:
        registros = (
            _montar_registros_substituicao(
                df_validado=df_validado,
                orcamento_id=orcamento_id,
                justificativa=justificativa,
                responsavel=responsavel
            )
        )

        acao_historico = (
            "Substituição integral via Excel"
        )

    if not registros:
        raise ValueError(
            "Nenhum valor preenchido foi encontrado "
            "para importação."
        )

    tamanho_lote = 500

    for inicio in range(
        0,
        len(registros),
        tamanho_lote
    ):
        lote = registros[
            inicio:inicio + tamanho_lote
        ]

        resposta = (
            supabase_client
            .table("orcamento_itens")
            .upsert(
                lote,
                on_conflict=(
                    "orcamento_id,conta_id,mes"
                )
            )
            .execute()
        )

        if resposta.data is None:
            raise RuntimeError(
                "O Supabase não confirmou a gravação "
                "de um dos lotes."
            )

    (
        supabase_client
        .table("orcamento_historico")
        .insert({
            "orcamento_id": int(
                orcamento_id
            ),
            "acao": "alterado",
            "valor_novo": {
                "tipo": "importacao_excel",
                "modo": modo,
                "quantidade_registros": len(
                    registros
                )
            },
            "usuario": (
                "Administrador Master"
            ),
            "observacao": (
                f"{acao_historico}. "
                f"{justificativa}"
            )
        })
        .execute()
    )

    return {
        "registros_processados": int(
            len(registros)
        ),
        "modo": modo
    }
