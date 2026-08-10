import io

import pandas as pd
import streamlit as st

from servico_orcamento import (
    MESES_NUMERO_NOME,
    alterar_status_orcamento,
    carregar_itens_orcamento,
    carregar_orcamentos,
    calcular_resumo_orcamento,
    criar_nova_versao_orcamento,
    montar_grade_orcamento,
    replicar_valor_na_grade,
    salvar_grade_orcamento,
)

COLUNAS_MESES = list(MESES_NUMERO_NOME.values())


def _formatar_moeda(valor):
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


def _validar_master():
    """
    Mantém a proteção adicional da área de orçamento.
    """

    try:
        senha_correta = st.secrets.get("MASTER_PASSWORD")
    except Exception:
        senha_correta = None

    if not senha_correta:
        st.error(
            "MASTER_PASSWORD não encontrada nos Secrets do Streamlit."
        )
        return False

    if st.session_state.get(
        "master_obz_autenticado",
        False
    ):
        return True

    st.warning(
        "Esta área é restrita ao Administrador Master."
    )

    senha_digitada = st.text_input(
        "Senha do Administrador Master",
        type="password",
        key="senha_master_obz"
    )

    if st.button(
        "Entrar no orçamento",
        key="btn_entrar_master_obz"
    ):
        if senha_digitada == senha_correta:
            st.session_state[
                "master_obz_autenticado"
            ] = True

            st.rerun()

        else:
            st.error("Senha incorreta.")

    return False


def _montar_rotulo_orcamento(row):
    ano = int(row["ano"])
    nome = str(row["nome"])
    versao = int(row["versao"])
    status = str(row["status"]).replace("_", " ").title()

    return (
        f"{ano} — {nome} — "
        f"Versão {versao} — {status}"
    )


def _carregar_grade_na_sessao(
    supabase_client,
    carregar_aba_base,
    orcamento_id
):
    """
    Recarrega a grade diretamente do banco.
    """

    df_plano = carregar_aba_base().copy()

    df_itens = carregar_itens_orcamento(
        supabase_client=supabase_client,
        orcamento_id=orcamento_id
    )

    grade = montar_grade_orcamento(
        df_plano_contas=df_plano,
        df_itens=df_itens
    )

    st.session_state["grade_obz"] = grade

    st.session_state[
        "grade_obz_orcamento_id"
    ] = int(orcamento_id)

    st.session_state[
        "grade_obz_alterada"
    ] = False


def _garantir_grade_carregada(
    supabase_client,
    carregar_aba_base,
    orcamento_id
):
    """
    Carrega a grade quando o usuário abre outro orçamento
    ou quando ainda não existe grade na sessão.
    """

    id_carregado = st.session_state.get(
        "grade_obz_orcamento_id"
    )

    if (
        "grade_obz" not in st.session_state
        or id_carregado != int(orcamento_id)
    ):
        _carregar_grade_na_sessao(
            supabase_client=supabase_client,
            carregar_aba_base=carregar_aba_base,
            orcamento_id=orcamento_id
        )


def _calcular_total_anual(df):
    df = df.copy()

    for coluna in COLUNAS_MESES:
        df[coluna] = pd.to_numeric(
            df[coluna],
            errors="coerce"
        ).fillna(0.0)

    df["Total Anual"] = df[
        COLUNAS_MESES
    ].sum(axis=1)

    return df


def _gerar_excel_orcamento(
    df_grade,
    ano,
    versao
):
    buffer = io.BytesIO()

    colunas_exportar = [
        "Conta",
        "Descrição",
        "Nivel",
        "Classificacao"
    ] + COLUNAS_MESES + ["Total Anual"]

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:
        df_grade[
            colunas_exportar
        ].to_excel(
            writer,
            index=False,
            sheet_name="Orçamento"
        )

        planilha = writer.sheets["Orçamento"]

        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        for coluna in planilha.columns:
            letra = coluna[0].column_letter

            maior_tamanho = 0

            for celula in coluna:
                valor = (
                    ""
                    if celula.value is None
                    else str(celula.value)
                )

                maior_tamanho = max(
                    maior_tamanho,
                    len(valor)
                )

            planilha.column_dimensions[
                letra
            ].width = min(
                maior_tamanho + 3,
                42
            )

        colunas_monetarias = (
            COLUNAS_MESES + ["Total Anual"]
        )

        for nome_coluna in colunas_monetarias:
            numero_coluna = colunas_exportar.index(
                nome_coluna
            ) + 1

            for linha in range(
                2,
                planilha.max_row + 1
            ):
                planilha.cell(
                    row=linha,
                    column=numero_coluna
                ).number_format = (
                    'R$ #,##0.00;[Red]-R$ #,##0.00'
                )

    nome_arquivo = (
        f"Orcamento_OBZ_{ano}_"
        f"Versao_{versao}.xlsx"
    )

    return buffer.getvalue(), nome_arquivo


def render_aba_orcamento_obz(
    supabase_client,
    carregar_aba_base
):
    st.subheader("💰 Orçamento Base Zero")

    if not _validar_master():
        return

    col_sair, col_recarregar, col_espaco = (
        st.columns([1, 1, 3])
    )

    with col_sair:
        if st.button(
            "🔒 Sair da área Master",
            key="btn_sair_master_obz"
        ):
            st.session_state[
                "master_obz_autenticado"
            ] = False

            st.session_state.pop(
                "grade_obz",
                None
            )

            st.session_state.pop(
                "grade_obz_orcamento_id",
                None
            )

            st.rerun()

    df_orcamentos = carregar_orcamentos(
        supabase_client
    )

    if df_orcamentos.empty:
        st.warning(
            "Nenhum orçamento foi encontrado no Supabase."
        )
        return

    df_orcamentos = df_orcamentos.copy()

    df_orcamentos["rotulo"] = (
        df_orcamentos.apply(
            _montar_rotulo_orcamento,
            axis=1
        )
    )

    rotulo_selecionado = st.selectbox(
        "Orçamento",
        options=df_orcamentos[
            "rotulo"
        ].tolist(),
        key="orcamento_obz_grade_selecionado"
    )

    linha_orcamento = df_orcamentos[
        df_orcamentos["rotulo"]
        == rotulo_selecionado
    ].iloc[0]

    orcamento_id = int(
        linha_orcamento["id"]
    )

    ano_orcamento = int(
        linha_orcamento["ano"]
    )

    versao_orcamento = int(
        linha_orcamento["versao"]
    )

    status_orcamento = str(
        linha_orcamento["status"]
    ).strip().lower()

    pode_editar = status_orcamento in [
        "rascunho",
        "em_revisao"
    ]

    _garantir_grade_carregada(
        supabase_client=supabase_client,
        carregar_aba_base=carregar_aba_base,
        orcamento_id=orcamento_id
    )

    with col_recarregar:
        if st.button(
            "🔄 Recarregar",
            key="btn_recarregar_grade_obz",
            help=(
                "Descarta alterações ainda não salvas "
                "e recarrega os valores do Supabase."
            )
        ):
            _carregar_grade_na_sessao(
                supabase_client=supabase_client,
                carregar_aba_base=carregar_aba_base,
                orcamento_id=orcamento_id
            )

            st.rerun()

    grade_atual = st.session_state[
        "grade_obz"
    ].copy()

    grade_atual = _calcular_total_anual(
        grade_atual
    )

    resumo = calcular_resumo_orcamento(
        grade_atual
    )

    st.write("### Resumo anual")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Receitas orçadas",
        _formatar_moeda(
            resumo["receitas"]
        )
    )

    c2.metric(
        "Despesas orçadas",
        _formatar_moeda(
            resumo["despesas"]
        )
    )

    c3.metric(
        "Resultado orçado",
        _formatar_moeda(
            resumo["resultado"]
        )
    )

    c4.metric(
        "Margem orçada",
        f"{resumo['margem']:.2f}%"
    )

    c5, c6, c7 = st.columns(3)

    c5.metric(
        "Ano",
        ano_orcamento
    )

    c6.metric(
        "Versão",
        versao_orcamento
    )

    c7.metric(
        "Status",
        status_orcamento
        .replace("_", " ")
        .title()
    )

    if not pode_editar:
        st.warning(
            "Este orçamento está bloqueado para edição. "
            f"Status atual: {status_orcamento}."
        )

    st.divider()

    st.write("### Grade anual do orçamento")

    st.caption(
        "Edite diretamente os valores de janeiro a dezembro. "
        "Receitas devem ser positivas e despesas negativas."
    )

    colunas_desabilitadas = [
        "Conta",
        "Descrição",
        "Nivel",
        "Classificacao",
        "Total Anual"
    ]

    configuracao_colunas = {
        "Conta": st.column_config.TextColumn(
            "Conta",
            disabled=True,
            width="medium"
        ),
        "Descrição": st.column_config.TextColumn(
            "Descrição",
            disabled=True,
            width="large"
        ),
        "Nivel": st.column_config.NumberColumn(
            "Nível",
            disabled=True
        ),
        "Classificacao": st.column_config.TextColumn(
            "Classificação",
            disabled=True
        ),
        "Total Anual": st.column_config.NumberColumn(
            "Total Anual",
            disabled=True,
            format="R$ %.2f"
        )
    }

    for nome_mes in COLUNAS_MESES:
        configuracao_colunas[
            nome_mes
        ] = st.column_config.NumberColumn(
            nome_mes,
            format="R$ %.2f",
            step=100.0,
            disabled=not pode_editar
        )

    grade_editada = st.data_editor(
        grade_atual,
        use_container_width=True,
        height=700,
        hide_index=True,
        disabled=(
            True
            if not pode_editar
            else colunas_desabilitadas
        ),
        column_config=configuracao_colunas,
        key=(
            f"editor_grade_obz_{orcamento_id}"
        )
    )

    grade_editada = _calcular_total_anual(
        grade_editada
    )

    st.session_state[
        "grade_obz"
    ] = grade_editada

    st.divider()

    st.write("### Replicar valor para meses seguintes")

    st.caption(
        "Use esta função depois de preencher o mês de origem "
        "na grade acima."
    )

    opcoes_contas = (
        grade_editada["Conta"]
        .astype(str)
        .tolist()
    )

    mapa_rotulo_contas = dict(
        zip(
            grade_editada["Conta"].astype(str),
            (
                grade_editada["Conta"].astype(str)
                + " — "
                + grade_editada["Descrição"].astype(str)
            )
        )
    )

    col_conta, col_origem, col_final = (
        st.columns([2, 1, 1])
    )

    with col_conta:
        conta_replicar = st.selectbox(
            "Conta",
            options=opcoes_contas,
            format_func=lambda conta: (
                mapa_rotulo_contas.get(
                    conta,
                    conta
                )
            ),
            key="conta_replicar_grade_obz",
            disabled=not pode_editar
        )

    with col_origem:
        mes_origem = st.selectbox(
            "Mês de origem",
            options=COLUNAS_MESES,
            key="mes_origem_grade_obz",
            disabled=not pode_editar
        )

    numero_mes_origem = (
        COLUNAS_MESES.index(
            mes_origem
        )
    )

    meses_finais_possiveis = (
        COLUNAS_MESES[
            numero_mes_origem:
        ]
    )

    with col_final:
        mes_final = st.selectbox(
            "Replicar até",
            options=meses_finais_possiveis,
            index=(
                len(
                    meses_finais_possiveis
                ) - 1
            ),
            key="mes_final_grade_obz",
            disabled=not pode_editar
        )

    if st.button(
        "➡️ Replicar valor",
        key="btn_replicar_grade_obz",
        disabled=not pode_editar
    ):
        try:
            grade_replicada = (
                replicar_valor_na_grade(
                    df_grade=grade_editada,
                    conta_id=conta_replicar,
                    mes_origem=mes_origem,
                    mes_final=mes_final
                )
            )

            st.session_state[
                "grade_obz"
            ] = grade_replicada

            st.success(
                f"Valor de {mes_origem} replicado "
                f"até {mes_final}."
            )

            st.rerun()

        except Exception as erro:
            st.error(
                "Não foi possível replicar o valor: "
                f"{type(erro).__name__} — {erro}"
            )

    st.divider()

    st.write("### Salvar orçamento")

    justificativa_geral = st.text_area(
        "Justificativa geral desta revisão",
        placeholder=(
            "Exemplo: revisão do orçamento mensal "
            "com base nas premissas aprovadas pela diretoria."
        ),
        key="justificativa_salvar_grade_obz",
        disabled=not pode_editar
    )

    responsavel = st.text_input(
        "Responsável pela revisão",
        value="Administrador Master",
        key="responsavel_salvar_grade_obz",
        disabled=not pode_editar
    )

    col_salvar, col_exportar = st.columns(2)

    with col_salvar:
        if st.button(
            "💾 Salvar alterações",
            key="btn_salvar_grade_obz",
            use_container_width=True,
            disabled=not pode_editar
        ):
            try:
                quantidade = (
                    salvar_grade_orcamento(
                        supabase_client=(
                            supabase_client
                        ),
                        orcamento_id=(
                            orcamento_id
                        ),
                        df_grade=(
                            grade_editada
                        ),
                        justificativa_padrao=(
                            justificativa_geral
                        ),
                        responsavel=(
                            responsavel
                        )
                    )
                )

                _carregar_grade_na_sessao(
                    supabase_client=(
                        supabase_client
                    ),
                    carregar_aba_base=(
                        carregar_aba_base
                    ),
                    orcamento_id=(
                        orcamento_id
                    )
                )

                st.success(
                    "Orçamento salvo com sucesso. "
                    f"{quantidade} registros mensais "
                    "foram processados."
                )

            except Exception as erro:
                st.error(
                    "Erro ao salvar a grade: "
                    f"{type(erro).__name__} — {erro}"
                )

    with col_exportar:
        arquivo_excel, nome_arquivo = (
            _gerar_excel_orcamento(
                df_grade=grade_editada,
                ano=ano_orcamento,
                versao=versao_orcamento
            )
        )

        st.download_button(
            "📥 Exportar orçamento para Excel",
            data=arquivo_excel,
            file_name=nome_arquivo,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    st.divider()

    st.write("### Fluxo de aprovação")

    observacao_status = st.text_area(
        "Observação para alteração de status",
        placeholder=(
            "Explique o motivo da revisão, aprovação, "
            "bloqueio ou criação de nova versão."
        ),
        key="observacao_status_obz"
    )

    usuario_master = st.session_state.get(
        "admin_usuario",
        "Administrador Master"
    )

    col_revisao, col_aprovar, col_bloquear, col_nova_versao = st.columns(4)

    with col_revisao:
        if st.button(
            "📤 Enviar para revisão",
            key="btn_enviar_revisao_obz",
            use_container_width=True,
            disabled=(status_orcamento != "rascunho")
        ):
            try:
                alterar_status_orcamento(
                    supabase_client=supabase_client,
                    orcamento_id=orcamento_id,
                    novo_status="em_revisao",
                    usuario=usuario_master,
                    observacao=observacao_status.strip()
                )

                st.session_state.pop("grade_obz", None)
                st.session_state.pop("grade_obz_orcamento_id", None)

                st.success("Orçamento enviado para revisão.")
                st.rerun()

            except Exception as erro:
                st.error(
                    "Erro ao enviar para revisão: "
                    f"{type(erro).__name__} — {erro}"
                )

    with col_aprovar:
        if st.button(
            "✅ Aprovar",
            key="btn_aprovar_obz",
            use_container_width=True,
            disabled=(status_orcamento != "em_revisao")
        ):
            try:
                alterar_status_orcamento(
                    supabase_client=supabase_client,
                    orcamento_id=orcamento_id,
                    novo_status="aprovado",
                    usuario=usuario_master,
                    observacao=observacao_status.strip()
                )

                st.session_state.pop("grade_obz", None)
                st.session_state.pop("grade_obz_orcamento_id", None)

                st.success("Orçamento aprovado.")
                st.rerun()

            except Exception as erro:
                st.error(
                    "Erro ao aprovar: "
                    f"{type(erro).__name__} — {erro}"
                )

    with col_bloquear:
        if st.button(
            "🔒 Bloquear",
            key="btn_bloquear_obz",
            use_container_width=True,
            disabled=(status_orcamento != "aprovado")
        ):
            try:
                alterar_status_orcamento(
                    supabase_client=supabase_client,
                    orcamento_id=orcamento_id,
                    novo_status="bloqueado",
                    usuario=usuario_master,
                    observacao=observacao_status.strip()
                )

                st.session_state.pop("grade_obz", None)
                st.session_state.pop("grade_obz_orcamento_id", None)

                st.success("Orçamento bloqueado.")
                st.rerun()

            except Exception as erro:
                st.error(
                    "Erro ao bloquear: "
                    f"{type(erro).__name__} — {erro}"
                )

    with col_nova_versao:
        if st.button(
            "🧾 Criar nova versão",
            key="btn_nova_versao_obz",
            use_container_width=True,
            disabled=(status_orcamento not in ["aprovado", "bloqueado"])
        ):
            try:
                novo_id, nova_versao = criar_nova_versao_orcamento(
                    supabase_client=supabase_client,
                    orcamento_id_origem=orcamento_id,
                    usuario=usuario_master,
                    observacao=observacao_status.strip()
                )

                st.session_state.pop("grade_obz", None)
                st.session_state.pop("grade_obz_orcamento_id", None)
                st.session_state.pop("orcamento_obz_grade_selecionado", None)

                st.success(
                    "Nova versão criada com sucesso. "
                    f"Versão {nova_versao}."
                )
                st.rerun()

            except Exception as erro:
                st.error(
                    "Erro ao criar nova versão: "
                    f"{type(erro).__name__} — {erro}"
                )
