import json
from urllib import error, request

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURAÇÃO DO GOOGLE GEMINI
# ============================================================

GEMINI_MODEL = "gemini-2.5-flash"

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)


def _obter_gemini_api_key():
    """
    Busca a chave do Gemini nos Secrets do Streamlit.
    """
    try:
        return st.secrets.get("GEMINI_API_KEY")
    except Exception:
        return None


def _formatar_moeda(valor):
    """
    Formata valores no padrão brasileiro.
    """
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


def _preparar_contexto_financeiro(
    df,
    ano_sel,
    meses_processados,
    cc_sel
):
    """
    Prepara somente os dados consolidados necessários para a IA.

    Não envia todos os lançamentos financeiros.
    Isso reduz custo, consumo da API e risco de exposição desnecessária.
    """

    df = df.copy()

    df["ACUMULADO"] = pd.to_numeric(
        df["ACUMULADO"],
        errors="coerce"
    ).fillna(0.0)

    df["MÉDIA"] = pd.to_numeric(
        df["MÉDIA"],
        errors="coerce"
    ).fillna(0.0)

    df_relevante = df[df["ACUMULADO"] != 0].copy()

    df_relevante["VALOR_ABS"] = (
        df_relevante["ACUMULADO"].abs()
    )

    linhas_nivel_1 = (
        df_relevante[
            df_relevante["Nivel"] == 1
        ]
        .copy()
    )

    linhas_nivel_2 = (
        df_relevante[
            df_relevante["Nivel"] == 2
        ]
        .copy()
    )

    maiores_contas = (
        df_relevante[
            df_relevante["Nivel"].isin([3, 4])
        ]
        .sort_values(
            "VALOR_ABS",
            ascending=False
        )
        .head(20)
    )

    def montar_linhas(dados):
        linhas = []

        for _, row in dados.iterrows():
            linhas.append(
                {
                    "conta": str(
                        row.get("Conta", "")
                    ),
                    "descricao": str(
                        row.get("Descrição", "")
                    ),
                    "nivel": int(
                        row.get("Nivel", 0)
                    ),
                    "acumulado": float(
                        row.get("ACUMULADO", 0.0)
                    ),
                    "media": float(
                        row.get("MÉDIA", 0.0)
                    ),
                }
            )

        return linhas

    contexto = {
        "ano": int(ano_sel),
        "meses": list(meses_processados),
        "centros_de_custo": list(cc_sel or []),
        "total_linhas_dataframe": int(len(df)),
        "linhas_com_movimento": int(len(df_relevante)),
        "visao_nivel_1": montar_linhas(linhas_nivel_1),
        "visao_nivel_2": montar_linhas(linhas_nivel_2),
        "maiores_contas_por_valor_absoluto": montar_linhas(
            maiores_contas
        ),
    }

    return contexto, linhas_nivel_1, maiores_contas


def _montar_prompt(contexto):
    """
    Define o perfil e a estrutura obrigatória da análise.
    """

    return f"""
Você atua como consultor empresarial sênior, controller e diretor
financeiro experiente, com domínio de finanças, controladoria,
gestão empresarial, rentabilidade, custos e tomada de decisão.

Analise exclusivamente os números existentes no JSON abaixo.

REGRAS OBRIGATÓRIAS:

- Não invente números.
- Não invente percentuais.
- Não invente metas, benchmarks ou comparações.
- Não presuma fatos que não estejam presentes nos dados.
- Quando faltar informação, declare claramente que não há dados suficientes.
- Diferencie fato, interpretação e recomendação.
- Use linguagem profissional, executiva e direta.
- Fale com o diretor da empresa.
- Use princípios de comunicação consultiva e PNL sem manipulação.
- Utilize perguntas estratégicas para provocar reflexão e decisão.
- Não suavize riscos relevantes.
- Não faça recomendações genéricas.
- Relacione cada recomendação aos números apresentados.

ESTRUTURA OBRIGATÓRIA:

## 1. Diagnóstico executivo

Apresente uma leitura objetiva do resultado do período.

## 2. Pontos críticos

Identifique os principais riscos, distorções, concentrações de despesas
e fatos que exigem atenção da diretoria.

## 3. Oportunidades

Apresente oportunidades reais de melhoria de margem, caixa,
custos, produtividade e gestão.

## 4. Recomendações práticas

Liste ações concretas, priorizadas em:

- ação imediata;
- ação para os próximos 30 dias;
- ação estrutural.

## 5. Perguntas estratégicas para o diretor

Faça perguntas consultivas que ajudem o diretor a avaliar decisões,
prioridades, riscos e oportunidades.

## 6. Plano de ação sugerido

Apresente uma tabela com:

- prioridade;
- ação;
- justificativa;
- responsável sugerido;
- prazo recomendado;
- indicador de acompanhamento.

DADOS FINANCEIROS:

{json.dumps(contexto, ensure_ascii=False, indent=2)}
""".strip()


def _extrair_texto_gemini(corpo):
    """
    Extrai o texto da resposta do Gemini com validação.
    """

    candidatos = corpo.get("candidates", [])

    if not candidatos:
        feedback = corpo.get("promptFeedback", {})

        raise ValueError(
            "O Gemini não retornou uma resposta. "
            f"Detalhes: {feedback}"
        )

    conteudo = candidatos[0].get("content", {})
    partes = conteudo.get("parts", [])

    textos = []

    for parte in partes:
        texto = parte.get("text")

        if texto:
            textos.append(texto)

    if not textos:
        raise ValueError(
            "O Gemini respondeu, mas não retornou texto."
        )

    return "\n".join(textos)


def _chamar_gemini(api_key, prompt):
    """
    Envia a solicitação ao Google Gemini usando REST.
    """

    payload = {
        "systemInstruction": {
            "parts": [
                {
                    "text": (
                        "Você gera análises financeiras e gerenciais "
                        "objetivas, não inventa números e sempre separa "
                        "fatos, interpretações e recomendações."
                    )
                }
            ]
        },
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
            "topP": 0.9,
            "maxOutputTokens": 4096
        }
    }

    dados = json.dumps(
        payload,
        ensure_ascii=False
    ).encode("utf-8")

    requisicao = request.Request(
        GEMINI_URL,
        data=dados,
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )

    with request.urlopen(
        requisicao,
        timeout=90
    ) as resposta:

        corpo = json.loads(
            resposta.read().decode("utf-8")
        )

    return _extrair_texto_gemini(corpo)


def render_aba_analista_ia(
    ano_sel,
    meses_sel,
    cc_sel,
    processar_bi
):
    """
    Renderiza a aba Analista IA.
    """

    st.subheader("🤖 Analista IA")

    api_key = _obter_gemini_api_key()

    if not api_key:
        st.warning(
            "GEMINI_API_KEY não encontrada nos Secrets do Streamlit. "
            "Cadastre a chave do Google Gemini para habilitar a análise."
        )
        return

    if not meses_sel:
        st.info(
            "Selecione ao menos um mês no filtro lateral "
            "para gerar a análise."
        )
        return

    st.info(
        "A análise utiliza somente informações financeiras "
        "consolidadas. Os lançamentos individuais não são enviados."
    )

    if st.button(
        "Gerar análise executiva",
        key="btn_analista_ia"
    ):
        with st.spinner(
            "Processando dados financeiros e gerando análise..."
        ):
            df_bi, meses_processados = processar_bi(
                ano_sel,
                meses_sel,
                cc_sel
            )

            if df_bi is None or df_bi.empty:
                st.warning(
                    "Não há dados financeiros para o período selecionado."
                )
                return

            (
                contexto,
                linhas_nivel_1,
                maiores_contas,
            ) = _preparar_contexto_financeiro(
                df_bi,
                ano_sel,
                meses_processados,
                cc_sel,
            )

            st.write("### Base numérica usada pela IA")

            st.caption(
                "A análise usa somente os números consolidados "
                "pela função processar_bi."
            )

            if not linhas_nivel_1.empty:
                visao = linhas_nivel_1[
                    [
                        "Conta",
                        "Descrição",
                        "ACUMULADO",
                        "MÉDIA",
                    ]
                ].copy()

                st.dataframe(
                    visao.style.format(
                        {
                            "ACUMULADO": _formatar_moeda,
                            "MÉDIA": _formatar_moeda,
                        }
                    ),
                    use_container_width=True,
                )

            with st.expander(
                "Maiores contas enviadas para análise"
            ):
                if maiores_contas.empty:
                    st.info(
                        "Nenhuma conta com movimento no período."
                    )
                else:
                    st.dataframe(
                        maiores_contas[
                            [
                                "Nivel",
                                "Conta",
                                "Descrição",
                                "ACUMULADO",
                                "MÉDIA",
                            ]
                        ]
                        .style.format(
                            {
                                "ACUMULADO": _formatar_moeda,
                                "MÉDIA": _formatar_moeda,
                            }
                        ),
                        use_container_width=True,
                    )

            try:
                analise = _chamar_gemini(
                    api_key,
                    _montar_prompt(contexto)
                )

            except error.HTTPError as exc:
                detalhe = exc.read().decode(
                    "utf-8",
                    errors="ignore"
                )

                if exc.code == 429:
                    st.error(
                        "O limite temporário da camada gratuita do "
                        "Google Gemini foi atingido. Aguarde alguns "
                        "minutos e tente novamente."
                    )
                elif exc.code == 400:
                    st.error(
                        "O Google Gemini recusou a solicitação. "
                        f"Detalhes técnicos: {detalhe}"
                    )
                elif exc.code == 403:
                    st.error(
                        "A chave do Google Gemini não tem autorização "
                        "para usar essa API. Verifique a chave no "
                        "Google AI Studio."
                    )
                else:
                    st.error(
                        f"Erro ao chamar o Google Gemini: "
                        f"HTTP {exc.code}. {detalhe}"
                    )

                return

            except error.URLError as exc:
                st.error(
                    "Não foi possível conectar ao Google Gemini. "
                    f"Detalhes: {exc.reason}"
                )
                return

            except Exception as exc:
                st.error(
                    "Erro ao gerar análise por IA: "
                    f"{type(exc).__name__} - {exc}"
                )
                return

            st.write("### Análise executiva")

            st.markdown(analise)
