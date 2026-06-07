import json
from urllib import error, request

import pandas as pd
import streamlit as st


OPENAI_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-4o-mini"


def _obter_openai_api_key():
    try:
        return st.secrets.get("OPENAI_API_KEY")
    except Exception:
        return None


def _formatar_moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    sinal = "-" if valor < 0 else ""
    numero = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sinal}R$ {numero}"


def _preparar_contexto_financeiro(df, ano_sel, meses_processados, cc_sel):
    df = df.copy()
    df["ACUMULADO"] = pd.to_numeric(df["ACUMULADO"], errors="coerce").fillna(0.0)
    df["MÉDIA"] = pd.to_numeric(df["MÉDIA"], errors="coerce").fillna(0.0)

    df_relevante = df[df["ACUMULADO"] != 0].copy()
    df_relevante["VALOR_ABS"] = df_relevante["ACUMULADO"].abs()

    linhas_nivel_1 = df_relevante[df_relevante["Nivel"] == 1].copy()
    linhas_nivel_2 = df_relevante[df_relevante["Nivel"] == 2].copy()
    maiores_contas = (
        df_relevante[df_relevante["Nivel"].isin([3, 4])]
        .sort_values("VALOR_ABS", ascending=False)
        .head(20)
    )

    def montar_linhas(dados):
        linhas = []
        for _, row in dados.iterrows():
            linhas.append(
                {
                    "conta": str(row.get("Conta", "")),
                    "descricao": str(row.get("Descrição", "")),
                    "nivel": int(row.get("Nivel", 0)),
                    "acumulado": float(row.get("ACUMULADO", 0.0)),
                    "media": float(row.get("MÉDIA", 0.0)),
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
        "maiores_contas_por_valor_absoluto": montar_linhas(maiores_contas),
    }

    return contexto, linhas_nivel_1, maiores_contas


def _montar_prompt(contexto):
    return f"""
Você é um analista financeiro sênior para uma marcenaria.
Analise exclusivamente os números do JSON abaixo. Não invente números, percentuais,
metas, benchmarks, saldos ou comparativos que não estejam nos dados.

Quando citar valores, use apenas valores presentes no JSON ou somas/conclusões
diretamente derivadas desses valores. Se faltar algum dado, diga que não há dado
suficiente.

Responda em português do Brasil, com linguagem executiva e prática, usando
exatamente estas seções:

1. Diagnóstico executivo
2. Pontos críticos
3. Oportunidades
4. Recomendações práticas
5. Perguntas estratégicas para o diretor

Dados financeiros:
{json.dumps(contexto, ensure_ascii=False, indent=2)}
""".strip()


def _chamar_openai(api_key, prompt):
    payload = {
        "model": OPENAI_MODEL,
        "messages": [
            {
                "role": "system",
                "content": (
                    "Você gera análises financeiras objetivas e nunca cria números "
                    "fora dos dados fornecidos."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.2,
    }

    dados = json.dumps(payload).encode("utf-8")
    requisicao = request.Request(
        OPENAI_URL,
        data=dados,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    with request.urlopen(requisicao, timeout=60) as resposta:
        corpo = json.loads(resposta.read().decode("utf-8"))

    return corpo["choices"][0]["message"]["content"]


def render_aba_analista_ia(ano_sel, meses_sel, cc_sel, processar_bi):
    st.subheader("🤖 Analista IA")

    api_key = _obter_openai_api_key()
    if not api_key:
        st.warning(
            "OPENAI_API_KEY não encontrada nos Secrets do Streamlit. "
            "Cadastre essa chave para habilitar a análise por IA."
        )
        return

    if not meses_sel:
        st.info("Selecione ao menos um mês no filtro lateral para gerar a análise.")
        return

    if st.button("Gerar análise executiva", key="btn_analista_ia"):
        with st.spinner("Processando dados financeiros e gerando análise..."):
            df_bi, meses_processados = processar_bi(ano_sel, meses_sel, cc_sel)

            if df_bi is None or df_bi.empty:
                st.warning("Não há dados financeiros para o período selecionado.")
                return

            contexto, linhas_nivel_1, maiores_contas = _preparar_contexto_financeiro(
                df_bi,
                ano_sel,
                meses_processados,
                cc_sel,
            )

            st.write("### Base numérica usada pela IA")
            st.caption("A análise abaixo usa somente os números consolidados por processar_bi.")

            if not linhas_nivel_1.empty:
                visao = linhas_nivel_1[["Conta", "Descrição", "ACUMULADO", "MÉDIA"]].copy()
                st.dataframe(
                    visao.style.format(
                        {
                            "ACUMULADO": _formatar_moeda,
                            "MÉDIA": _formatar_moeda,
                        }
                    ),
                    use_container_width=True,
                )

            with st.expander("Maiores contas enviadas para análise"):
                if maiores_contas.empty:
                    st.info("Nenhuma conta com movimento no período.")
                else:
                    st.dataframe(
                        maiores_contas[["Nivel", "Conta", "Descrição", "ACUMULADO", "MÉDIA"]]
                        .style.format(
                            {
                                "ACUMULADO": _formatar_moeda,
                                "MÉDIA": _formatar_moeda,
                            }
                        ),
                        use_container_width=True,
                    )

            try:
                analise = _chamar_openai(api_key, _montar_prompt(contexto))
            except error.HTTPError as exc:
                detalhe = exc.read().decode("utf-8", errors="ignore")
                st.error(f"Erro ao chamar a OpenAI: HTTP {exc.code}. {detalhe}")
                return
            except Exception as exc:
                st.error(f"Erro ao gerar análise por IA: {type(exc).__name__} - {exc}")
                return

            st.write("### Análise")
            st.markdown(analise)
