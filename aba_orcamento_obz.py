import pandas as pd
import streamlit as st


MESES = {
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


def _validar_master():
    """
    Protege a aba de orçamento com a senha cadastrada
    nos Secrets do Streamlit.
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

    if st.session_state.get("master_obz_autenticado", False):
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
            st.session_state["master_obz_autenticado"] = True
            st.success("Acesso liberado.")
            st.rerun()
        else:
            st.error("Senha incorreta.")

    return False


def _carregar_orcamentos(supabase_client):
    resposta = (
        supabase_client
        .table("orcamentos")
        .select("*")
        .order("ano", desc=True)
        .order("versao", desc=True)
        .execute()
    )

    return pd.DataFrame(resposta.data or [])


def _carregar_contas(carregar_aba_base):
    df_contas = carregar_aba_base().copy()

    if df_contas.empty:
        return pd.DataFrame()

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

    return df_contas


def render_aba_orcamento_obz(
    supabase_client,
    carregar_aba_base
):
    st.subheader("💰 Orçamento Base Zero")

    if not _validar_master():
        return

    col_sair, col_espaco = st.columns([1, 4])

    with col_sair:
        if st.button(
            "🔒 Sair da área Master",
            key="btn_sair_master_obz"
        ):
            st.session_state["master_obz_autenticado"] = False
            st.rerun()

    df_orcamentos = _carregar_orcamentos(
        supabase_client
    )

    if df_orcamentos.empty:
        st.warning(
            "Nenhum orçamento foi encontrado no Supabase."
        )
        return

    df_orcamentos["rotulo"] = (
        df_orcamentos["ano"].astype(str)
        + " — "
        + df_orcamentos["nome"].astype(str)
        + " — Versão "
        + df_orcamentos["versao"].astype(str)
        + " — "
        + df_orcamentos["status"].astype(str)
    )

    rotulo_selecionado = st.selectbox(
        "Orçamento",
        options=df_orcamentos["rotulo"].tolist(),
        key="orcamento_obz_selecionado"
    )

    orcamento = df_orcamentos[
        df_orcamentos["rotulo"] == rotulo_selecionado
    ].iloc[0]

    orcamento_id = int(orcamento["id"])
    status = str(orcamento["status"]).strip().lower()

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Ano",
        int(orcamento["ano"])
    )

    c2.metric(
        "Versão",
        int(orcamento["versao"])
    )

    c3.metric(
        "Status",
        status.replace("_", " ").title()
    )

    if status not in ["rascunho", "em_revisao"]:
        st.warning(
            "Este orçamento não permite alterações porque "
            f"está com status: {status}."
        )
        return

    st.divider()
    st.write("### Incluir valor orçado")

    df_contas = _carregar_contas(
        carregar_aba_base
    )

    if df_contas.empty:
        st.error(
            "Não foi possível carregar o plano de contas."
        )
        return

    # Inicialmente usamos apenas contas analíticas.
    df_analiticas = df_contas[
        df_contas["Nivel"] >= 4
    ].copy()

    df_analiticas["rotulo"] = (
        df_analiticas["Conta"]
        + " — "
        + df_analiticas["Descrição"]
    )

    conta_rotulo = st.selectbox(
        "Conta contábil",
        options=df_analiticas["rotulo"].tolist(),
        key="conta_orcamento_obz"
    )

    conta_id = conta_rotulo.split(" — ")[0].strip()

    col_mes, col_valor = st.columns(2)

    with col_mes:
        mes_inicial = st.selectbox(
            "Mês inicial",
            options=list(MESES.keys()),
            format_func=lambda numero: MESES[numero],
            key="mes_inicial_orcamento_obz"
        )

    with col_valor:
        valor_orcado = st.number_input(
            "Valor mensal",
            value=0.0,
            step=100.0,
            format="%.2f",
            key="valor_orcamento_obz"
        )

    replicar = st.checkbox(
        "Replicar este valor para os meses seguintes",
        value=False,
        key="replicar_orcamento_obz"
    )

    mes_final = mes_inicial

    if replicar:
        mes_final = st.selectbox(
            "Replicar até",
            options=[
                mes
                for mes in MESES.keys()
                if mes >= mes_inicial
            ],
            index=len([
                mes
                for mes in MESES.keys()
                if mes >= mes_inicial
            ]) - 1,
            format_func=lambda numero: MESES[numero],
            key="mes_final_orcamento_obz"
        )

    justificativa = st.text_area(
        "Justificativa obrigatória",
        placeholder=(
            "Explique por que este valor é necessário "
            "no orçamento base zero."
        ),
        key="justificativa_orcamento_obz"
    )

    responsavel = st.text_input(
        "Responsável",
        value="Administrador Master",
        key="responsavel_orcamento_obz"
    )

    if st.button(
        "💾 Salvar valor orçado",
        key="btn_salvar_item_orcamento_obz"
    ):
        justificativa_limpa = justificativa.strip()

        if valor_orcado == 0:
            st.error(
                "Informe um valor diferente de zero."
            )
            return

        if len(justificativa_limpa) < 3:
            st.error(
                "Informe uma justificativa com ao menos "
                "três caracteres."
            )
            return

        meses_gravar = list(
            range(mes_inicial, mes_final + 1)
        )

        registros = []

        for mes in meses_gravar:
            registros.append({
                "orcamento_id": orcamento_id,
                "conta_id": conta_id,
                "mes": int(mes),
                "valor_orcado": float(valor_orcado),
                "justificativa": justificativa_limpa,
                "responsavel": responsavel.strip(),
                "criado_por": "Administrador Master"
            })

        try:
            for registro in registros:
                (
                    supabase_client
                    .table("orcamento_itens")
                    .upsert(
                        registro,
                        on_conflict=(
                            "orcamento_id,conta_id,mes"
                        )
                    )
                    .execute()
                )

            acao = (
                "replicado"
                if len(registros) > 1
                else "incluido"
            )

            (
                supabase_client
                .table("orcamento_historico")
                .insert({
                    "orcamento_id": orcamento_id,
                    "acao": acao,
                    "valor_novo": {
                        "conta_id": conta_id,
                        "mes_inicial": mes_inicial,
                        "mes_final": mes_final,
                        "valor_orcado": float(valor_orcado)
                    },
                    "usuario": "Administrador Master",
                    "observacao": justificativa_limpa
                })
                .execute()
            )

            st.success(
                f"Valor salvo para {len(registros)} mês(es)."
            )

        except Exception as erro:
            st.error(
                "Erro ao salvar o orçamento: "
                f"{type(erro).__name__} — {erro}"
            )

    st.divider()
    st.write("### Valores já cadastrados")

    try:
        resposta_itens = (
            supabase_client
            .table("orcamento_itens")
            .select("*")
            .eq("orcamento_id", orcamento_id)
            .order("conta_id")
            .order("mes")
            .execute()
        )

        df_itens = pd.DataFrame(
            resposta_itens.data or []
        )

        if df_itens.empty:
            st.info(
                "Este orçamento ainda não possui valores."
            )
            return

        mapa_descricao = dict(
            zip(
                df_contas["Conta"],
                df_contas["Descrição"]
            )
        )

        df_itens["Descrição"] = (
            df_itens["conta_id"]
            .map(mapa_descricao)
            .fillna("")
        )

        df_itens["Mês"] = (
            df_itens["mes"]
            .map(MESES)
        )

        df_visual = df_itens[[
            "conta_id",
            "Descrição",
            "Mês",
            "valor_orcado",
            "justificativa",
            "responsavel"
        ]].copy()

        df_visual = df_visual.rename(columns={
            "conta_id": "Conta",
            "valor_orcado": "Valor Orçado",
            "justificativa": "Justificativa",
            "responsavel": "Responsável"
        })

        st.dataframe(
            df_visual.style.format({
                "Valor Orçado": (
                    "R$ {:,.2f}"
                )
            }),
            use_container_width=True,
            height=500
        )

    except Exception as erro:
        st.error(
            "Erro ao carregar os itens do orçamento: "
            f"{type(erro).__name__} — {erro}"
        )
