"""Reclassificação gerencial e relatório de fretes — sem alterar movimentos_financeiros."""
import io
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
        q = admin.table('reclassificacoes_itens').select(
            'data_lancamento,ano,mes,centro_custo,conta_origem,conta_destino,valor'
        ).in_('lote_id', ids[i:i+100])
        itens.extend(_paginar(q))
    registros = []
    for x in itens:
        v = float(x['valor'])
        base = {'Data': x['data_lancamento'], 'Ano': int(x['ano']),
                'Mes': int(x['mes']), 'Centro de Custo': x['centro_custo']}
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
    df = pd.read_excel(io.BytesIO(arquivo), sheet_name=0,
                       dtype={'CTR': str, 'Centro de Custo': str})
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
            itens.append({'data_lancamento': data.isoformat(), 'ctr': ctr,
                          'centro_custo': cc, 'valor': str(valor.quantize(Decimal('0.01')))})
        except (ValueError, TypeError, InvalidOperation, OverflowError) as e:
            erros.append(f'Linha {pos}: {e}')
    if erros:
        raise ValueError('Problemas encontrados:\n' + '\n'.join(erros[:30]))
    return itens


def conferir_itens(itens, ano, mes, supabase_client):
    """Prévia: verifica apenas cadastro de Centros de Custo, sem consultar receita."""
    rateios = _paginar(supabase_client.table('rateio_config').select('centro_custo'))
    cadastrados = {str(x['centro_custo']).strip() for x in rateios}
    solicitado = {}
    for x in itens:
        cc = x['centro_custo']
        solicitado[cc] = solicitado.get(cc, Decimal('0')) + Decimal(x['valor'])
    linhas = []
    valido = True
    for cc, v in sorted(solicitado.items()):
        cadastrado = cc in cadastrados
        valido = valido and cadastrado
        linhas.append({'Centro de Custo': cc,
                       'Frete solicitado (R$)': float(v), 'Cadastrado': cadastrado})
    return pd.DataFrame(linhas), valido


def _moeda(valor):
    return 'R$ ' + f'{float(valor):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def _aba_importacao(ano_sel, mapa_meses, supabase_client):
    st.subheader('Reclassificação gerencial de fretes')
    st.caption('Transfere receita de 01.01.002 para 01.01.019 sem alterar os lançamentos originais. Cada mês possui um único lote ativo.')
    nomes = list(mapa_meses.keys())
    c1, c2 = st.columns(2)
    with c1:
        ano = st.number_input('Ano da reclassificação', min_value=2020,
                              max_value=2100, value=int(ano_sel), step=1, key='reclass_ano')
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
        st.metric('Fretes na planilha', _moeda(total))
        st.write(f'{len(itens)} lançamentos em {len(previa)} Centros de Custo.')
        st.dataframe(previa, use_container_width=True)
        if not valido:
            st.error('Importação bloqueada: existem Centros de Custo não cadastrados. Corrija a planilha ou cadastre os Centros de Custo.')
            return
        st.warning('A confirmação substituirá o lote ativo deste mês, se houver. O histórico anterior será preservado.')
        confirmar = st.checkbox('Conferi o mês, os Centros de Custo e os valores.', key='reclass_confirmar')
        if st.button('Confirmar reclassificação', disabled=not confirmar,
                     type='primary', key='reclass_gravar'):
            try:
                admin = cliente_admin()
                resposta = admin.rpc('importar_reclassificacao_fretes', {
                    'p_ano': int(ano), 'p_mes': int(mes),
                    'p_nome_arquivo': arq.name, 'p_itens': itens
                }).execute()
                st.cache_data.clear()
                resultado = resposta.data or {}
                st.success(f"Lote {resultado.get('lote_id', '')} gravado. {len(itens)} fretes reclassificados; total {_moeda(total)}.")
            except Exception:
                st.error('Falha ao gravar o lote. Nenhuma alteração parcial deve ser aplicada. Verifique os logs do servidor e as validações do banco.')
    except Exception as e:
        st.error(f'Não foi possível validar a planilha: {e}')


def _ler_movimentos_frete(ano, meses, supabase_client):
    """Usa o mesmo regime de caixa e os mesmos ajustes gerenciais do BI."""
    registros = []
    for mes in meses:
        q = (supabase_client.table('movimentos_financeiros')
             .select('data,ano,mes,conta_id,centro_custo,valor')
             .eq('ano', int(ano)).eq('mes', str(int(mes))))
        registros.extend(_paginar(q))
    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.rename(columns={'data': 'Data', 'ano': 'Ano', 'mes': 'Mes',
                                'conta_id': 'Conta_ID', 'centro_custo': 'Centro de Custo',
                                'valor': 'Valor_Final'})
        df['Mes'] = pd.to_numeric(df['Mes'], errors='coerce')
        df['Valor_Final'] = pd.to_numeric(df['Valor_Final'], errors='coerce').fillna(0.0)
        df['Conta_ID'] = df['Conta_ID'].astype(str).str.strip()
        df['Centro de Custo'] = df['Centro de Custo'].astype(str).str.strip()
    return combinar_movimentos(df, ano, meses)


def _montar_relatorio(ano, meses, centros, supabase_client):
    plano = pd.DataFrame(_paginar(
        supabase_client.table('plano_contas').select(
            'conta_id,descricao,nivel,classificacao,classificacao_gerencial')))
    if plano.empty:
        raise ValueError('Plano de Contas vazio.')
    obrigatorias = {'conta_id', 'descricao', 'nivel', 'classificacao', 'classificacao_gerencial'}
    if not obrigatorias.issubset(plano.columns):
        raise ValueError('Plano de Contas sem os campos necessários à classificação gerencial.')
    plano['conta_id'] = plano['conta_id'].astype(str).str.strip()
    plano['classificacao_gerencial'] = plano['classificacao_gerencial'].fillna('').astype(str).str.lower().str.strip()
    plano['classificacao'] = plano['classificacao'].fillna('').astype(str).str.lower().str.strip()
    plano['nivel'] = pd.to_numeric(plano['nivel'], errors='coerce')
    contas = plano[(plano['classificacao_gerencial'] == 'operacional_frete') &
                   (plano['nivel'] == 4)].copy()
    if contas.empty:
        raise ValueError('Nenhuma conta analítica (nível 4) classificada como operacional_frete.')
    fora_operacional = contas[contas['classificacao'] != 'operacional']
    if not fora_operacional.empty:
        raise ValueError('Há contas operacional_frete cuja classificação principal não é operacional. Corrija o Plano de Contas antes de gerar o relatório.')
    df = _ler_movimentos_frete(ano, meses, supabase_client)
    if df.empty:
        return pd.DataFrame(), contas, pd.DataFrame()
    df = df[df['Conta_ID'].isin(contas['conta_id'])].copy()
    if centros:
        df = df[df['Centro de Custo'].isin(centros)]
    if df.empty:
        return df, contas, pd.DataFrame()
    df = df.merge(contas[['conta_id', 'descricao']], left_on='Conta_ID',
                  right_on='conta_id', how='left', validate='many_to_one')
    df['Tipo'] = df['Conta_ID'].apply(lambda c: 'Receita' if c.startswith('01.') else 'Despesa' if c.startswith('02.') else 'Outro')
    if (df['Tipo'] == 'Outro').any():
        raise ValueError('Há contas operacional_frete fora dos grupos 01 (receita) e 02 (despesa).')
    # Valores das despesas são negativos no BI. Resultado = receita + despesa.
    mensal = df.groupby(['Mes', 'Tipo'], as_index=False)['Valor_Final'].sum()
    mensal = mensal.pivot(index='Mes', columns='Tipo', values='Valor_Final').fillna(0.0)
    for col in ['Receita', 'Despesa']:
        if col not in mensal.columns:
            mensal[col] = 0.0
    mensal = mensal.reindex(meses, fill_value=0.0)
    mensal['Resultado'] = mensal['Receita'] + mensal['Despesa']
    mensal['Margem (%)'] = mensal.apply(
        lambda r: r['Resultado'] / r['Receita'] * 100 if r['Receita'] > 0 else 0.0, axis=1)
    mensal.index.name = 'Mes'
    return df, contas, mensal.reset_index()


def _montar_dre_fretes(df, meses, mapa_meses):
    """DRE gerencial: contas analíticas sem somar linhas de grupo novamente."""
    nomes = {v: k for k, v in mapa_meses.items()}
    colunas_meses = [nomes[m] for m in meses]
    base = (df.groupby(['Conta_ID', 'descricao', 'Mes'], as_index=False)['Valor_Final']
            .sum())
    pivot = base.pivot_table(index=['Conta_ID', 'descricao'], columns='Mes',
                             values='Valor_Final', aggfunc='sum', fill_value=0.0)
    pivot = pivot.reindex(columns=meses, fill_value=0.0)
    pivot.columns = colunas_meses
    linhas = []

    def adicionar(tipo, conta, descricao, valores):
        registro = {'Tipo': tipo, 'Conta': conta, 'Descrição': descricao}
        registro.update({col: float(valores.get(col, 0.0)) for col in colunas_meses})
        registro['ACUMULADO'] = sum(registro[col] for col in colunas_meses)
        linhas.append(registro)

    receitas = pivot[pivot.index.get_level_values(0).str.startswith('01.')]
    despesas = pivot[pivot.index.get_level_values(0).str.startswith('02.')]
    total_receita = receitas.sum(axis=0)
    total_despesa = despesas.sum(axis=0)
    adicionar('total', 'R', 'RECEITA DE FRETES', total_receita)
    for (conta, descricao), valores in receitas.sort_index().iterrows():
        adicionar('conta', conta, descricao, valores)
    adicionar('total', 'D', 'DESPESAS DA FROTA', total_despesa)
    # Grupos por prefixo contábil, mantendo a hierarquia da conta.
    for prefixo in sorted({'.'.join(conta.split('.')[:2]) for conta, _ in despesas.index}):
        grupo = despesas[despesas.index.get_level_values(0).str.startswith(prefixo + '.')]
        adicionar('grupo', prefixo, 'FROTA' if prefixo == '02.07' else f'GRUPO {prefixo}', grupo.sum(axis=0))
        for (conta, descricao), valores in grupo.sort_index().iterrows():
            adicionar('conta', conta, descricao, valores)
    resultado = total_receita.add(total_despesa, fill_value=0.0)
    adicionar('resultado', 'RF', 'RESULTADO GERENCIAL DE FRETES', resultado)
    margens = {col: (float(resultado.get(col, 0)) / float(total_receita.get(col, 0)) * 100)
               if float(total_receita.get(col, 0)) > 0 else 0.0 for col in colunas_meses}
    adicionar('margem', '%', 'MARGEM SOBRE RECEITA (%)', margens)
    dre = pd.DataFrame(linhas)
    # Acumulado de percentuais não deve ser somado.
    receita_acum = float(total_receita.sum())
    dre.loc[dre['Tipo'] == 'margem', 'ACUMULADO'] = (
        float(resultado.sum()) / receita_acum * 100 if receita_acum > 0 else 0.0)
    return dre, colunas_meses


def _estilizar_dre_fretes(linha):
    tipo = linha['Tipo']
    if tipo == 'resultado':
        return ['background-color: #334155; color: white; font-weight: bold'] * len(linha)
    if tipo == 'total':
        return ['background-color: #cbd5e1; color: black; font-weight: bold'] * len(linha)
    if tipo == 'grupo':
        return ['background-color: #d1eaff; color: black; font-weight: bold'] * len(linha)
    if tipo == 'margem':
        return ['background-color: #e2e8f0; color: black; font-weight: bold'] * len(linha)
    return [''] * len(linha)


def _aba_relatorio(ano_sel, mapa_meses, supabase_client):
    st.subheader('Resultado gerencial de fretes')
    st.caption('Somente contas analíticas com classificacao_gerencial = operacional_frete. '
               'O resultado operacional consolidado permanece inalterado.')
    nomes = list(mapa_meses.keys())
    c1, c2 = st.columns([1, 2])
    with c1:
        ano = st.number_input('Ano do relatório', min_value=2020, max_value=2100,
                              value=int(ano_sel), step=1, key='frete_rel_ano')
    with c2:
        meses_nomes = st.multiselect('Meses do relatório', nomes, default=nomes,
                                     key='frete_rel_meses')
    try:
        centros_cadastrados = _paginar(supabase_client.table('rateio_config').select('centro_custo'))
        centros_opcoes = sorted({str(x['centro_custo']).strip() for x in centros_cadastrados
                                if x.get('centro_custo')})
    except Exception as e:
        st.error(f'Não foi possível carregar os Centros de Custo: {e}')
        return
    centros = st.multiselect('Centro de Custo (vazio = todos)', centros_opcoes,
                             key='frete_rel_centros')
    gerar = st.button('Gerar relatório de fretes', type='primary', key='frete_rel_gerar')
    filtros = (int(ano), tuple(sorted(meses_nomes)), tuple(sorted(centros)))
    if gerar:
        if not meses_nomes:
            st.warning('Selecione pelo menos um mês.')
            return
        meses = sorted({int(mapa_meses[n]) for n in meses_nomes})
        try:
            df, contas, mensal = _montar_relatorio(int(ano), meses, centros, supabase_client)
            st.session_state['frete_rel_resultado'] = (filtros, df, mensal)
        except Exception as e:
            st.session_state.pop('frete_rel_resultado', None)
            st.error(f'Não foi possível gerar o relatório: {e}')
            return
    resultado_salvo = st.session_state.get('frete_rel_resultado')
    if resultado_salvo is None or resultado_salvo[0] != filtros:
        st.info('Selecione os filtros e clique em Gerar relatório de fretes.')
        return
    _, df, mensal = resultado_salvo
    if df.empty:
        st.info('Não há lançamentos nas contas classificadas como operacional_frete para os filtros selecionados.')
        return
    receita = float(df.loc[df['Tipo'] == 'Receita', 'Valor_Final'].sum())
    despesas = float(df.loc[df['Tipo'] == 'Despesa', 'Valor_Final'].sum())
    resultado = receita + despesas
    margem = (resultado / receita * 100) if receita > 0 else 0.0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Receitas de fretes', _moeda(receita))
    c2.metric('Despesas da Frota', _moeda(abs(despesas)))
    c3.metric('Resultado de fretes', _moeda(resultado))
    c4.metric('Margem de fretes', f'{margem:.1f}%'.replace('.', ','))
    st.markdown('#### DRE gerencial de fretes')
    dre, colunas_dre = _montar_dre_fretes(df, sorted(mensal['Mes'].astype(int).tolist()), mapa_meses)
    dre_visivel = dre[['Tipo', 'Conta', 'Descrição'] + colunas_dre + ['ACUMULADO']].copy()
    formatacoes = {col: _moeda for col in colunas_dre + ['ACUMULADO']}
    for col in colunas_dre + ['ACUMULADO']:
        formatacoes[col] = lambda v, col=col: _moeda(v)
    # Percentuais usam formato próprio, sem somar percentuais mensais no acumulado.
    def formatar_celula(valor, tipo):
        return (f'{valor:.1f}%'.replace('.', ',') if tipo == 'margem' else _moeda(valor))
    dre_formatada = dre_visivel.copy()
    for col in colunas_dre + ['ACUMULADO']:
        dre_formatada[col] = [formatar_celula(v, t) for v, t in zip(dre_visivel[col], dre_visivel['Tipo'])]
    st.dataframe(dre_formatada.style.apply(_estilizar_dre_fretes, axis=1),
                 column_order=['Conta', 'Descrição'] + colunas_dre + ['ACUMULADO'],
                 hide_index=True, use_container_width=True, height=min(850, 95 + len(dre) * 35))
    st.caption('DRE gerencial em regime de caixa. Valores das despesas negativos; '
               'contas analíticas aparecem uma vez e as linhas de grupo são subtotais.')
    st.markdown('#### Evolução mensal')
    nomes_por_numero = {v: k for k, v in mapa_meses.items()}
    mensal['Mês'] = mensal['Mes'].map(nomes_por_numero)
    st.dataframe(mensal[['Mês', 'Receita', 'Despesa', 'Resultado', 'Margem (%)']].style.format({
        'Receita': _moeda, 'Despesa': _moeda, 'Resultado': _moeda,
        'Margem (%)': lambda x: f'{x:.1f}%'.replace('.', ',')
    }), use_container_width=True, hide_index=True)
    st.bar_chart(mensal.set_index('Mês')[['Receita', 'Resultado']])
    st.markdown('#### Detalhamento por conta')
    por_conta = (df.groupby(['Conta_ID', 'descricao', 'Tipo'], as_index=False)['Valor_Final']
                 .sum().sort_values('Conta_ID'))
    por_conta = por_conta.rename(columns={'Conta_ID': 'Conta', 'descricao': 'Descrição',
                                          'Valor_Final': 'Valor (R$)'})
    st.dataframe(por_conta.style.format({'Valor (R$)': _moeda}),
                 use_container_width=True, hide_index=True)
    st.markdown('#### Resultado por Centro de Custo')
    por_centro = df.pivot_table(index='Centro de Custo', columns='Tipo',
                                values='Valor_Final', aggfunc='sum', fill_value=0.0)
    for col in ['Receita', 'Despesa']:
        if col not in por_centro.columns:
            por_centro[col] = 0.0
    por_centro['Resultado'] = por_centro['Receita'] + por_centro['Despesa']
    por_centro = por_centro.reset_index().sort_values('Resultado', ascending=False)
    st.dataframe(por_centro[['Centro de Custo', 'Receita', 'Despesa', 'Resultado']].style.format({
        'Receita': _moeda, 'Despesa': _moeda, 'Resultado': _moeda
    }), use_container_width=True, hide_index=True)
    st.caption('Despesas exibidas com sinal negativo; resultado = receitas + despesas. '
               'Não há rateio adicional de despesas da Frota entre obras.')
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        dre.drop(columns=['Tipo']).to_excel(writer, sheet_name='DRE Fretes', index=False)
        mensal[['Mês', 'Receita', 'Despesa', 'Resultado', 'Margem (%)']].to_excel(
            writer, sheet_name='Resumo Mensal', index=False)
        por_conta.to_excel(writer, sheet_name='Por Conta', index=False)
        por_centro[['Centro de Custo', 'Receita', 'Despesa', 'Resultado']].to_excel(
            writer, sheet_name='Por Centro de Custo', index=False)
        df[['Data', 'Ano', 'Mes', 'Centro de Custo', 'Conta_ID', 'descricao',
            'Tipo', 'Valor_Final']].to_excel(writer, sheet_name='Lancamentos', index=False)
    st.download_button('Exportar relatório de fretes (Excel)', data=buffer.getvalue(),
                       file_name=f'Relatorio_Fretes_{int(ano)}.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       key='frete_rel_excel')


def render_aba_reclassificacao_fretes(ano_sel, mapa_meses, supabase_client):
    tab_importar, tab_relatorio = st.tabs(['Reclassificação mensal', 'Relatório de fretes'])
    with tab_importar:
        _aba_importacao(ano_sel, mapa_meses, supabase_client)
    with tab_relatorio:
        _aba_relatorio(ano_sel, mapa_meses, supabase_client)
