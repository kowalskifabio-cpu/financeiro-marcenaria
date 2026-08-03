import pandas as pd


MESES_NUMERO_NOME = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

MESES_NOME_NUMERO = {
    nome: numero
    for numero, nome in MESES_NUMERO_NOME.items()
}


def carregar_orcamentos(supabase_client):
    """
    Carrega as versões de orçamento cadastradas.
    """

    resposta = (
        supabase_client
        .table("orcamentos")
        .select("*")
        .order("ano", desc=True)
        .order("versao", desc=True)
        .execute()
    )

    return pd.DataFrame(resposta.data or [])


def carregar_itens_orcamento(
    supabase_client,
    orcamento_id
):
    """
    Carrega os valores mensais de um orçamento.
    """

    resposta = (
        supabase_client
        .table("orcamento_itens")
        .select("*")
        .eq("orcamento_id", int(orcamento_id))
        .order("conta_id")
        .order("mes")
        .execute()
    )

    return pd.DataFrame(resposta.data or [])


def montar_grade_orcamento(
    df_plano_contas,
    df_itens
):
    """
    Monta uma linha por conta e uma coluna por mês.

    A grade inclui todas as contas analíticas do plano de contas,
    mesmo que ainda não tenham orçamento.
    """

    if df_plano_contas is None or df_plano_contas.empty:
        return pd.DataFrame()

    df_contas = df_plano_contas.copy()

    df_contas["Conta"] = (
        df_contas["Conta"]
        .astype(str)
        .str.strip()
    )

    df_contas["Descrição"] = (
        df_contas["Descrição"]
        .astype(str)
        .str.strip()
    )

    df_contas["Nivel"] = pd.to_numeric(
        df_contas["Nivel"],
        errors="coerce"
    ).fillna(0).astype(int)

    # Primeira versão: orçamento somente nas contas analíticas.
    df_contas = df_contas[
        df_contas["Nivel"] >= 4
    ].copy()

    grade = df_contas[
        [
            "Conta",
            "Descrição",
            "Nivel",
            "Classificacao"
        ]
    ].copy()

    for nome_mes in MESES_NOME_NUMERO:
        grade[nome_mes] = 0.0

    if df_itens is not None and not df_itens.empty:
        itens = df_itens.copy()

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

        tabela_meses = itens.pivot_table(
            index="conta_id",
            columns="mes",
            values="valor_orcado",
            aggfunc="sum",
            fill_value=0.0
        )

        for numero_mes, nome_mes in MESES_NUMERO_NOME.items():
            if numero_mes in tabela_meses.columns:
                mapa_mes = tabela_meses[
                    numero_mes
                ].to_dict()

                grade[nome_mes] = (
                    grade["Conta"]
                    .map(mapa_mes)
                    .fillna(0.0)
                )

    grade["Total Anual"] = grade[
        list(MESES_NOME_NUMERO.keys())
    ].sum(axis=1)

    return grade.reset_index(drop=True)


def validar_grade_orcamento(df_grade):
    """
    Confere se a grade possui as colunas necessárias.
    """

    colunas_obrigatorias = [
        "Conta",
        "Descrição"
    ] + list(MESES_NOME_NUMERO.keys())

    faltantes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df_grade.columns
    ]

    if faltantes:
        raise ValueError(
            "A grade do orçamento está sem as colunas: "
            + ", ".join(faltantes)
        )


def transformar_grade_em_registros(
    df_grade,
    orcamento_id,
    justificativa_padrao,
    responsavel
):
    """
    Converte a grade em registros mensais para o Supabase.

    Valores zerados também são retornados para permitir que um valor
    anteriormente cadastrado seja removido da versão atual.
    """

    validar_grade_orcamento(df_grade)

    registros = []

    for _, linha in df_grade.iterrows():
        conta_id = str(
            linha.get("Conta", "")
        ).strip()

        if not conta_id:
            continue

        for nome_mes, numero_mes in MESES_NOME_NUMERO.items():
            valor = pd.to_numeric(
                linha.get(nome_mes, 0.0),
                errors="coerce"
            )

            if pd.isna(valor):
                valor = 0.0

            registros.append({
                "orcamento_id": int(orcamento_id),
                "conta_id": conta_id,
                "mes": int(numero_mes),
                "valor_orcado": float(valor),
                "justificativa": justificativa_padrao,
                "responsavel": responsavel,
                "criado_por": "Administrador Master"
            })

    return registros


def salvar_grade_orcamento(
    supabase_client,
    orcamento_id,
    df_grade,
    justificativa_padrao,
    responsavel
):
    """
    Salva toda a grade por upsert.

    A chave de conflito é:
    orçamento + conta + mês.
    """

    justificativa = str(
        justificativa_padrao
    ).strip()

    responsavel_limpo = str(
        responsavel
    ).strip()

    if len(justificativa) < 3:
        raise ValueError(
            "Informe uma justificativa geral com ao menos "
            "três caracteres."
        )

    if not responsavel_limpo:
        responsavel_limpo = "Administrador Master"

    registros = transformar_grade_em_registros(
        df_grade=df_grade,
        orcamento_id=orcamento_id,
        justificativa_padrao=justificativa,
        responsavel=responsavel_limpo
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

        (
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

    (
        supabase_client
        .table("orcamento_historico")
        .insert({
            "orcamento_id": int(orcamento_id),
            "acao": "alterado",
            "valor_novo": {
                "quantidade_registros": len(registros),
                "tipo": "salvamento_grade"
            },
            "usuario": "Administrador Master",
            "observacao": justificativa
        })
        .execute()
    )

    return len(registros)


def replicar_valor_na_grade(
    df_grade,
    conta_id,
    mes_origem,
    mes_final
):
    """
    Replica o valor de um mês para os meses seguintes
    na mesma conta.
    """

    grade = df_grade.copy()

    conta_id = str(conta_id).strip()

    numero_origem = MESES_NOME_NUMERO[
        mes_origem
    ]

    numero_final = MESES_NOME_NUMERO[
        mes_final
    ]

    if numero_final < numero_origem:
        raise ValueError(
            "O mês final não pode ser anterior ao mês de origem."
        )

    linha_conta = (
        grade["Conta"].astype(str).str.strip()
        == conta_id
    )

    if not linha_conta.any():
        raise ValueError(
            "A conta selecionada não foi encontrada na grade."
        )

    valor_origem = grade.loc[
        linha_conta,
        mes_origem
    ].iloc[0]

    for numero_mes in range(
        numero_origem + 1,
        numero_final + 1
    ):
        nome_mes = MESES_NUMERO_NOME[
            numero_mes
        ]

        grade.loc[
            linha_conta,
            nome_mes
        ] = valor_origem

    grade["Total Anual"] = grade[
        list(MESES_NOME_NUMERO.keys())
    ].sum(axis=1)

    return grade


def calcular_resumo_orcamento(df_grade):
    """
    Calcula receitas, despesas, resultado e margem.
    """

    if df_grade is None or df_grade.empty:
        return {
            "receitas": 0.0,
            "despesas": 0.0,
            "resultado": 0.0,
            "margem": 0.0
        }

    total_por_conta = pd.to_numeric(
        df_grade["Total Anual"],
        errors="coerce"
    ).fillna(0.0)

    contas = (
        df_grade["Conta"]
        .astype(str)
        .str.strip()
    )

    receitas = total_por_conta[
        contas.str.startswith("01")
    ].sum()

    despesas = total_por_conta[
        contas.str.startswith("02")
    ].sum()

    resultado = receitas + despesas

    margem = (
        resultado / receitas * 100
        if receitas != 0
        else 0.0
    )

    return {
        "receitas": float(receitas),
        "despesas": float(despesas),
        "resultado": float(resultado),
        "margem": float(margem)
    }

def alterar_status_orcamento(
    supabase_client,
    orcamento_id,
    novo_status,
    usuario,
    observacao=""
):
    """
    Altera o status do orçamento e registra a ação no histórico.
    """

    status_validos = [
        "rascunho",
        "em_revisao",
        "aprovado",
        "bloqueado",
        "cancelado"
    ]

    if novo_status not in status_validos:
        raise ValueError(
            f"Status inválido: {novo_status}"
        )

    dados_atualizacao = {
        "status": novo_status
    }

    if novo_status == "aprovado":
        dados_atualizacao["aprovado_por"] = usuario
        dados_atualizacao["aprovado_em"] = "now()"

    if novo_status == "bloqueado":
        dados_atualizacao["bloqueado_por"] = usuario
        dados_atualizacao["bloqueado_em"] = "now()"

    (
        supabase_client
        .table("orcamentos")
        .update(dados_atualizacao)
        .eq("id", int(orcamento_id))
        .execute()
    )

    mapa_acao = {
        "em_revisao": "enviado_revisao",
        "aprovado": "aprovado",
        "bloqueado": "bloqueado",
        "rascunho": "desbloqueado",
        "cancelado": "cancelado"
    }

    (
        supabase_client
        .table("orcamento_historico")
        .insert({
            "orcamento_id": int(orcamento_id),
            "acao": mapa_acao[novo_status],
            "valor_novo": {
                "status": novo_status
            },
            "usuario": usuario,
            "observacao": observacao
        })
        .execute()
    )


def criar_nova_versao_orcamento(
    supabase_client,
    orcamento_id_origem,
    usuario,
    observacao=""
):
    """
    Cria uma nova versão copiando o orçamento e todos os itens.
    """

    resposta_origem = (
        supabase_client
        .table("orcamentos")
        .select("*")
        .eq("id", int(orcamento_id_origem))
        .single()
        .execute()
    )

    origem = resposta_origem.data

    if not origem:
        raise ValueError(
            "Orçamento de origem não encontrado."
        )

    ano = int(origem["ano"])
    nome = str(origem["nome"])

    resposta_versoes = (
        supabase_client
        .table("orcamentos")
        .select("versao")
        .eq("ano", ano)
        .eq("nome", nome)
        .order("versao", desc=True)
        .limit(1)
        .execute()
    )

    versoes = resposta_versoes.data or []
    nova_versao = (
        int(versoes[0]["versao"]) + 1
        if versoes
        else 1
    )

    resposta_novo = (
        supabase_client
        .table("orcamentos")
        .insert({
            "ano": ano,
            "nome": nome,
            "versao": nova_versao,
            "status": "rascunho",
            "observacao": observacao,
            "criado_por": usuario
        })
        .execute()
    )

    novo_orcamento = (
        resposta_novo.data or []
    )[0]

    novo_id = int(
        novo_orcamento["id"]
    )

    resposta_itens = (
        supabase_client
        .table("orcamento_itens")
        .select("*")
        .eq("orcamento_id", int(orcamento_id_origem))
        .execute()
    )

    itens_origem = resposta_itens.data or []

    novos_itens = []

    for item in itens_origem:
        novos_itens.append({
            "orcamento_id": novo_id,
            "conta_id": item["conta_id"],
            "mes": int(item["mes"]),
            "valor_orcado": float(item["valor_orcado"]),
            "justificativa": item.get(
                "justificativa",
                "Copiado da versão anterior."
            ),
            "responsavel": item.get(
                "responsavel",
                usuario
            ),
            "observacao": item.get("observacao"),
            "criado_por": usuario
        })

    tamanho_lote = 500

    for inicio in range(
        0,
        len(novos_itens),
        tamanho_lote
    ):
        lote = novos_itens[
            inicio:inicio + tamanho_lote
        ]

        (
            supabase_client
            .table("orcamento_itens")
            .insert(lote)
            .execute()
        )

    (
        supabase_client
        .table("orcamento_historico")
        .insert({
            "orcamento_id": novo_id,
            "acao": "criado",
            "valor_novo": {
                "versao": nova_versao,
                "orcamento_origem_id": int(
                    orcamento_id_origem
                )
            },
            "usuario": usuario,
            "observacao": observacao
        })
        .execute()
    )

    return novo_id, nova_versao
