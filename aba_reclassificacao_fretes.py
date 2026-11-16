"""Reclassificacao gerencial de fretes — sem alterar movimentos_financeiros."""
import io
from datetime import date
from decimal import Decimal, InvalidOperation

import pandas as pd
import streamlit as st
from supabase import create_client

ORIGEM = '01.01.002'
DESTINO = '01.01.019'
COLUNAS = ['Data do lançamento', 'CTR', 'Centro de Custo', 'Valor do Frete (R$)']


def cliente_admin():
    """A credencial privilegiada permanece no servidor Streamlit."""
    url = st.secrets['supabase']['url']
    key = st.secrets['supabase_admin']['service_role_key']
    if not url or not key:
        raise ValueError('Configuração administrativa incompleta.')
    return create_client(url, key)


def _paginar(consulta, passo=1000):
    saida = []
    inicio = 0
    while True:
        lote = consulta.range(inicio, inicio + passo - 1).execute().data or []
        saida.extend(lote)
        if len(lote) < passo:
            break
        inicio += passo
    return saida


def movimentos_gerenciais(ano=None, meses=None):
    """Somente lotes ativos; dois ajustes por item, soma total zero."""
    admin = cliente_admin()
    q = admin.table('reclassificacoes_lotes').select('id,ano,mes').eq('status', 'ativo')
    if ano is not None:
        q = q.eq('ano', int(ano))
    if meses is not None:
        numeros = [int(x) for x in meses]
        if not numeros:
            return pd.DataFrame()
        q = q.in_('mes', numeros)
    lotes = _paginar(q)
    if not lotes:
        return pd.DataFrame()
    ids = [x['id'] for x in lotes]
    itens = []
    for i in range(0, len(ids), 100):
        q = admin.table('reclassificacoes_itens').select('data_lancamento,ano,mes,centro_custo,conta_origem,conta_destino,valor').in_('lote_id', ids[i:i+100])
        itens.extend(_paginar(q))
    registros = []
    for x in itens:
        v = float(x['valor'])
        base = {'Data': x['data_lancamento'], 'Ano': int(x['ano']), 'Mes': int(x['mes']), 'Centro de Custo': x['centro_custo']}
        registros.append({**base, 'Conta_ID': x['conta_origem'], 'Valor_Final': -v})
        registros.append({**base, 'Conta_ID': x['conta_destino'], 'Valor_Final': v})
    return pd.DataFrame(registros)


def combinar_movimentos(df_original, ano=None, meses=None):
    """Preserva o formato interno usado por todas as abas existentes."""
    ajustes = movimentos_gerenciais(ano, meses)
    if ajustes.empty:
        return df_original
    if df_original.empty:
        return ajustes
    return pd.concat([df_original, ajustes], ignore_index=True)


def preparar_planilha(arquivo, ano, mes):
    df = pd.read_excel(io.BytesIO(arquivo), sheet_name=0, dtype={'CTR': str, 'Centro de Custo': str})
    df.columns = [str(c).strip() for c in df.columns]
    faltantes = [c for c in COLUNAS if c not in df.columns]
    if faltantes:
        raise ValueError('Colunas ausentes: ' + ', '.join(faltantes))
    df = df[COLUNAS].dropna(how='all').copy()
    if df.empty:
        raise ValueError('A planilha está vazia.')
    erros = []
    itens = []
    for pos, row in enumerate(df.to_dict('records'), start=2):
        try:
            dt = pd.to_datetime(row[COLUNAS[0]], errors='raise', dayfirst=True)
            if pd.isna(dt):
                raise ValueError('data vazia')
            data = dt.date()
            if data.year != ano or data.month != mes:
                raise ValueError('data fora do mês selecionado')
            ctr = str(row['CTR']).strip() if pd.notna(row['CTR']) else ''
            cc = str(row['Centro de Custo']).strip() if pd.notna(row['Centro de Custo']) else ''
            if not ctr or not cc:
                raise ValueError('CTR ou Centro de Custo vazio')
            bruto = row[COLUNAS[3]]
            if isinstance(bruto, str):
                bruto = bruto.strip().replace('R$', '').replace(' ', '')
                if ',' in bruto:
                    bruto = bruto.replace('.', '').replace(',', '.')
            valor = Decimal(str(bruto))
            if not valor.is_finite() or valor <= 0 or valor != valor.quantize(Decimal('0.01')):
                raise ValueError('valor inválido ou com mais de duas casas decimais')
            itens.append({'data_lancamento': data.isoformat(), 'ctr': ctr, 'centro_custo': cc, 'valor': str(valor.quantize(Decimal('0.01')))})
        except (ValueError, TypeError, InvalidOperation, OverflowError) as e:
            erros.append(f'Linha {pos}: {e}')
    if erros:
        raise ValueError('Problemas encontrados:\n' + '\n'.join(erros[:30]))
    return itens


def conferir_itens(itens, ano, mes, supabase_client):
    """Prévia: checagem de receita original por CC e de cadastro de CC."""
    receitas = _paginar(supabase_client.table('movimentos_financeiros').select('centro_custo,valor').eq('ano', ano).eq('mes', str(mes)).eq('conta_id', ORIGEM))
    rateios = _paginar(supabase_client.table('rateio_config').select('centro_custo'))
    cadastrados = {str(x['centro_custo']).strip() for x in rateios}
    disponivel = {}
    for x in receitas:
        cc = str(x['centro_custo']).strip()
        disponivel[cc] = disponivel.get(cc, Decimal('0')) + Decimal(str(x['valor']))
    solicitado = {}
    for x in itens:
        cc = x['centro_custo']
        solicitado[cc] = solicitado.get(cc, Decimal('0')) + Decimal(x['valor'])
    linhas = []
    valido = True
    for cc, v in sorted(solicitado.items()):
        saldo = disponivel.get(cc, Decimal('0'))
        cadastrado = cc in cadastrados
        suficiente = v <= saldo
        valido = valido and cadastrado and suficiente
        linhas.append({'Centro de Custo': cc, 'Frete solicitado (R$)': float(v), 'Receita disponível (R$)': float(saldo), 'Cadastrado': cadastrado, 'Receita suficiente': suficiente})
    return pd.DataFrame(linhas), valido


def render_aba_reclassificacao_fretes(ano_sel, mapa_meses, supabase_client):
    st.subheader('Reclassificação gerencial de fretes')
    st.caption('Transfere receita de 01.01.002 para 01.01.019 sem alterar os lançamentos originais. Cada mês possui um único lote ativo.')
    nomes = list(mapa_meses.keys())
    c1, c2 = st.columns(2)
    with c1:
        ano = st.number_input('Ano da reclassificação', min_value=2020, max_value=2100, value=int(ano_sel), step=1, key='reclass_ano')
    with c2:
        mes_nome = st.selectbox('Mês da reclassificação', nomes, key='reclass_mes')
    mes = mapa_meses[mes_nome]
    arq = st.file_uploader('Planilha de fretes (.xlsx)', type=['xlsx'], key='reclass_arquivo')
    if arq is None:
        st.info('Selecione uma planilha para conferir os lançamentos antes de importar.')
        return
    try:
        conteudo = arq.getvalue()
        itens = preparar_planilha(conteudo, int(ano), int(mes))
        previa, valido = conferir_itens(itens, int(ano), int(mes), supabase_client)
        total = sum((Decimal(x['valor']) for x in itens), Decimal('0'))
        st.metric('Fretes na planilha', f'R$ {total:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
        st.write(f'{len(itens)} lançamentos em {len(previa)} Centros de Custo.')
        st.dataframe(previa, use_container_width=True)
        if not valido:
            st.error('Importação bloqueada: existem Centros de Custo não cadastrados ou receita insuficiente. Corrija a planilha ou os dados de origem.')
            return
        st.warning('A confirmação substituirá o lote ativo deste mês, se houver. O histórico anterior será preservado.')
        confirmar = st.checkbox('Conferi o mês, os Centros de Custo e os valores.', key='reclass_confirmar')
        if st.button('Confirmar reclassificação', disabled=not confirmar, type='primary', key='reclass_gravar'):
            try:
                admin = cliente_admin()
                resposta = admin.rpc('importar_reclassificacao_fretes', {
                    'p_ano': int(ano), 'p_mes': int(mes), 'p_nome_arquivo': arq.name, 'p_itens': itens
                }).execute()
                st.cache_data.clear()
                resultado = resposta.data or {}
                st.success(f"Lote {resultado.get('lote_id', '')} gravado. {len(itens)} fretes reclassificados; total R$ {total:,.2f}.")
            except Exception:
                st.error('Falha ao gravar o lote. Nenhuma alteração parcial deve ser aplicada. Verifique os logs do servidor e as validações do banco.')
    except Exception as e:
        st.error(f'Não foi possível validar a planilha: {e}')
