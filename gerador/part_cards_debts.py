"""CARTÕES DE CRÉDITO e DÍVIDAS E FINANCIAMENTOS (com simulador de quitação)."""
from openpyxl.chart import LineChart, Reference, Series
from common import *
from part_base import N_CART, dyn

D0, D1 = 9, 16          # linhas das dívidas
N_DIV = 8
HOR = 120               # meses simulados


def fat(card, f, extra=''):
    return f'SUMIFS(L_VParc,L_Conta,{card},L_Tipo,"Despesa"{extra},L_FIni,"<="&({f}),L_FFim,">="&({f}))'


def venc(card_k, f_expr):
    """data de vencimento da fatura de índice f (card_k = posição do cartão)."""
    vi = f'(({f_expr})+IF(INDEX(Cartoes_Venc,{card_k})>INDEX(Cartoes_Fech,{card_k}),0,1))'
    return f'DATE(INT(({vi}-1)/12),MOD({vi}-1,12)+1,INDEX(Cartoes_Venc,{card_k}))'


def build_cartoes(ctx):
    sh = Sh(ctx, 'CARTÕES DE CRÉDITO')
    ws = sh.ws
    S = 'CARTÕES DE CRÉDITO'
    sh.header('CARTÕES DE CRÉDITO', 'Tudo calculado a partir de LANÇAMENTOS (compras com a coluna "Conta ou Cartão" = cartão) '
              'e do cadastro de cartões em CONFIGURAÇÕES. Compras a partir do dia de fechamento caem na fatura seguinte.',
              ncols=15)
    sh.widths({'B': 24, 'C': 14, 'D': 14, 'E': 14, 'F': 10, 'G': 15, 'H': 15, 'I': 15, 'J': 13, 'K': 13, 'L': 12, 'M': 14,
               'N': 28, 'O': 14})
    sh.put('B6', '1. Situação de cada cartão', 'sec')
    heads = ['Cartão', 'Limite', 'Usado (comprometido)', 'Disponível', '% do limite', 'Fatura fechada a pagar (não paga)',
             'Fatura atual (aberta)', 'Próxima fatura', 'Vencimento da fatura fechada', 'Vencimento da fatura atual',
             'Melhor dia de compra', 'Gasto no mês selecionado', 'Alerta', 'Idx fatura atual']
    for i, h in enumerate(heads):
        sh.put(f'{CL(2 + i)}8', h, 'hdr')
    ws.row_dimensions[8].height = 46
    for k in range(1, N_CART + 1):
        r = 8 + k
        card = f'$B{r}'
        sh.put(f'B{r}', f'=IF(INDEX(Cartoes_Nome,{k})="","",INDEX(Cartoes_Nome,{k}))', 'calc', bold=True)
        sh.put(f'C{r}', f'=IF($B{r}="",0,N(INDEX(Cartoes_Limite,{k})))', 'calc', BRL)
        sh.put(f'O{r}', f'=IF($B{r}="",0,YEAR(Cfg_Hoje)*12+MONTH(Cfg_Hoje)+IF(DAY(Cfg_Hoje)>=INDEX(Cartoes_Fech,{k}),1,0))', 'calc', INT)
        sh.put(f'G{r}', (f'=IF($B{r}="",0,IF(AND(J{r}>=Cfg_Hoje,SUMIFS(L_VParc,L_Tipo,"Transferência",L_Status,"Pago",L_Desc,"*"&{card}&"*",'
                         f'L_IIni,YEAR(J{r})*12+MONTH(J{r}))>=0),{fat(card, f"O{r}-1")},0))'), 'calc', BRL)
        sh.put(f'H{r}', f'=IF($B{r}="",0,{fat(card, f"O{r}")})', 'calc', BRL)
        sh.put(f'I{r}', f'=IF($B{r}="",0,{fat(card, f"O{r}+1")})', 'calc', BRL)
        sh.put(f'J{r}', f'=IF($B{r}="","",{venc(k, f"O{r}-1")})', 'calc', DATA, h='center')
        sh.put(f'K{r}', f'=IF($B{r}="","",{venc(k, f"O{r}")})', 'calc', DATA, h='center')
        sh.put(f'D{r}', f'=IF($B{r}="",0,SUMIFS(L_Aberto,L_Conta,{card})+G{r})', 'calc', BRL)
        sh.put(f'E{r}', f'=IF($B{r}="",0,C{r}-D{r})', 'calc', BRL)
        sh.put(f'F{r}', f'=IF(C{r}=0,0,D{r}/C{r})', 'calc', PCT)
        sh.put(f'L{r}', f'=IF($B{r}="","","dia "&INDEX(Cartoes_Fech,{k}))', 'calc', h='center')
        sh.put(f'M{r}', f'=IF($B{r}="",0,SUMIFS(L_VParc,L_Conta,{card},L_Tipo,"Despesa",L_IIni,"<="&Sel_Idx,L_IFim,">="&Sel_Idx))', 'calc', BRL)
        sh.put(f'N{r}', f'=IF($B{r}="","",IF(F{r}>Cfg_LimCartao,"⚠ Uso acima de "&ROUND(Cfg_LimCartao*100,0)&"% do limite","✔ Uso saudável"))', 'calc')
        sh.cf_formula(f'B{r}:O{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    r1 = 8 + N_CART
    sh.put(f'B{r1 + 1}', 'Total', 'calc', bold=True)
    for L in 'CDEGHIM':
        sh.put(f'{L}{r1 + 1}', f'=SUM({L}9:{L}{r1})', 'calc', BRL, bold=True)
    sh.put(f'F{r1 + 1}', f'=IF(C{r1 + 1}=0,0,D{r1 + 1}/C{r1 + 1})', 'calc', PCT, bold=True)
    for L, n in zip('BCDEFGHIJKLMN', ['Nome', 'Limite', 'Usado', 'Disp', 'Pct', 'FatAnt', 'FatAtual', 'FatProx', 'VencAnt',
                                        'VencAtual', 'Melhor', 'GastoMes', 'Alerta']):
        ctx.name(f'Card_{n}', S, f'{L}9:{L}{r1}')
    ctx.name('Card_FatAntTot', S, f'G{r1 + 1}')
    ctx.name('Card_FatAtualTot', S, f'H{r1 + 1}')
    ctx.name('Card_UsadoTot', S, f'D{r1 + 1}')
    sh.cf_formula(f'F9:F{r1}', f'AND($B9<>"",F9>=0.9)', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'F9:F{r1}', f'AND($B9<>"",F9>Cfg_LimCartao,F9<0.9)', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'F9:F{r1}', f'AND($B9<>"",F9<=Cfg_LimCartao)', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'N9:N{r1}', f'LEFT(N9,1)="⚠"', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'J9:K{r1}', f'AND(ISNUMBER(J9),J9-Cfg_Hoje<=7,J9>=Cfg_Hoje)', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'J9:K{r1}', f'AND(ISNUMBER(J9),J9<Cfg_Hoje)', None, COR['mudo'])
    sh.put(f'B{r1 + 2}', 'Usado = parcelas a vencer da fatura atual em diante + fatura fechada ainda não vencida. '
           'Melhor dia de compra = dia de fechamento (compras a partir dele só entram na fatura seguinte).', 'note')
    # ---- parcelas futuras mês a mês
    p = r1 + 4
    sh.put(f'B{p}', '2. Parcelas futuras comprometidas, mês a mês (fatura de cada mês)', 'sec')
    sh.put(f'B{p + 1}', 'Cartão', 'hdr')
    for j in range(12):
        L = CL(3 + j)
        sh.put(f'{L}{p + 1}', f'={ml_label(f"(YEAR(Cfg_Hoje)*12+MONTH(Cfg_Hoje)+{j})")}', 'hdr')
    sh.put(f'O{p + 1}', 'Total', 'hdr')
    for k in range(1, N_CART + 1):
        r = p + 1 + k
        sh.put(f'B{r}', f'=B{8 + k}', 'calc', bold=True)
        for j in range(12):
            L = CL(3 + j)
            sh.put(f'{L}{r}', f'=IF($B{r}="",0,{fat(f"$B{r}", f"YEAR(Cfg_Hoje)*12+MONTH(Cfg_Hoje)+{j}")})', 'calc', BRL)
        sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRL, bold=True)
        sh.cf_formula(f'B{r}:O{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    rt = p + 2 + N_CART
    sh.put(f'B{rt}', 'Total do mês', 'calc', bold=True)
    for j in range(13):
        L = CL(3 + j)
        sh.put(f'{L}{rt}', f'=SUM({L}{p + 2}:{L}{rt - 1})', 'calc', BRL, bold=True)
    sh.put(f'B{rt + 1}', 'Cada coluna mostra a fatura do mês indicado (fatura que fecha naquele mês), a partir do mês de hoje, com as parcelas '
           'já lançadas em LANÇAMENTOS.', 'note')
    ctx.name('Card_Matriz', S, f'C{p + 2}:N{rt - 1}')
    # ---- comparativo
    q = rt + 3
    sh.put(f'B{q}', '3. Comparativo de gastos entre cartões (mês selecionado)', 'sec')
    for i, h in enumerate(['Cartão', 'Gasto no mês', '% do total', 'Ranking']):
        sh.put(f'{CL(2 + i)}{q + 1}', h, 'hdr')
    for k in range(1, N_CART + 1):
        r = q + 1 + k
        sh.put(f'B{r}', f'=B{8 + k}', 'calc')
        sh.put(f'C{r}', f'=M{8 + k}', 'calc', BRL)
        sh.put(f'D{r}', f'=IFERROR(C{r}/$M${r1 + 1},0)', 'calc', PCT)
        sh.put(f'E{r}', f'=IF(OR($B{r}="",C{r}=0),"",RANK(C{r},$C${q + 2}:$C${q + 1 + N_CART})+COUNTIF($C${q + 2}:C{r},C{r})-1)', 'calc', INT, h='center')
        sh.cf_formula(f'B{r}:E{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    sh.cf(f'C{q + 2}:C{q + 1 + N_CART}', DataBarRule(start_type='num', start_value=0, end_type='max', color=COR['vermelho']))
    sh.finish(freeze='C9')
    return sh


# ======================================================================
def build_dividas(ctx):
    sh = Sh(ctx, 'DÍVIDAS E FINANCIAMENTOS')
    ws = sh.ws
    S = 'DÍVIDAS E FINANCIAMENTOS'
    sh.header('DÍVIDAS E FINANCIAMENTOS', 'Cadastre suas dívidas (células azuis). A planilha calcula quitação, tabelas de amortização '
              '(Price e SAC) e simula Bola de Neve x Avalanche com um valor extra por mês.', ncols=16)
    sh.widths({'B': 28, 'C': 24, 'D': 15, 'E': 15, 'F': 14, 'G': 13, 'H': 13, 'I': 11, 'J': 11, 'K': 13, 'L': 11, 'M': 13,
               'N': 15, 'O': 11, 'P': 22})
    sh.put('B5', '⬇ digite nas colunas azuis (B a K)', 'note', color=COR['azul'], bold=True)
    sh.put('B6', '1. Cadastro de dívidas', 'sec')
    heads = ['Credor / dívida', 'Tipo', 'Valor original', 'Saldo devedor', 'Taxa de juros (mês)', 'Parcela', 'Nº de parcelas',
             'Parcelas pagas', 'Sistema', 'Status', 'Parcelas restantes', 'Quitação prevista', 'Juros restantes (est.)',
             '% já pago', 'Situação']
    for i, h in enumerate(heads):
        sh.put(f'{CL(2 + i)}8', h, 'hdr')
    ws.row_dimensions[8].height = 46
    fm_in = [None, None, BRLZ, BRLZ, '0.00%', BRLZ, INT, INT, None, None]
    for r in range(D0, D1 + 1):
        for L, fm in zip('BCDEFGHIJK', fm_in):
            sh.put(f'{L}{r}', None, 'in', fm)
        sh.put(f'L{r}', f'=IF($B{r}="",0,MAX(0,N(H{r})-N(I{r})))', 'calc', INT)
        sh.put(f'M{r}', f'=IF(OR($B{r}="",L{r}=0,K{r}="Quitada"),"",EDATE(Cfg_Hoje,L{r}))', 'calc', MESANO, h='center')
        sh.put(f'N{r}', f'=IF($B{r}="",0,MAX(0,N(G{r})*L{r}-N(E{r})))', 'calc', BRL)
        sh.put(f'O{r}', f'=IF(N(D{r})=0,0,MAX(0,1-N(E{r})/D{r}))', 'calc', PCT)
        sh.put(f'P{r}', f'=IF($B{r}="","",IF(K{r}="Quitada","✔ Quitada",IF(K{r}="Atrasada","⚠ Atrasada",IF(N(F{r})>0.03,"⚠ Juros altos","Em dia"))))', 'calc')
    sh.dv(f'C{D0}:C{D1}', '=Lista_TipoDivida')
    sh.dv(f'J{D0}:J{D1}', '=Lista_Sistema')
    sh.dv(f'K{D0}:K{D1}', '=Lista_StatusDivida')
    for L, n in zip('BCDEFGHIJKLMNOP', ['Nome', 'Tipo', 'Orig', 'Saldo', 'Taxa', 'Parc', 'NTot', 'Pagas', 'Sist', 'Status',
                                         'Rest', 'Quit', 'Juros', 'PctPago', 'Sit']):
        ctx.name(f'Div_{n}', S, f'{L}{D0}:{L}{D1}')
    ctx.name_formula('Lista_Dividas', dyn(S, 'B', D0, D1))
    ctx.name('Exemplo_Dividas', S, f'B{D0}:K{D1}')
    ctx.clear_ranges.append((S, f'B{D0}:K{D1}', 'Dívidas de exemplo'))
    sh.cf_formula(f'P{D0}:P{D1}', f'LEFT(P{D0},1)="⚠"', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'P{D0}:P{D1}', f'LEFT(P{D0},1)="✔"', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'P{D0}:P{D1}', f'P{D0}="Em dia"', None, COR['verde'])
    sh.cf_formula(f'K{D0}:K{D1}', f'K{D0}="Atrasada"', COR['vermelho_l'], COR['vermelho'], True)
    t = D1 + 1
    sh.put(f'B{t}', 'Total (dívidas não quitadas)', 'calc', bold=True)
    sh.put(f'D{t}', f'=SUMIFS(D{D0}:D{D1},K{D0}:K{D1},"<>Quitada")', 'calc', BRL, bold=True)
    sh.put(f'E{t}', f'=SUMIFS(E{D0}:E{D1},K{D0}:K{D1},"<>Quitada")', 'calc', BRL, bold=True)
    sh.put(f'G{t}', f'=SUMIFS(G{D0}:G{D1},K{D0}:K{D1},"<>Quitada")', 'calc', BRL, bold=True)
    sh.put(f'N{t}', f'=SUMIFS(N{D0}:N{D1},K{D0}:K{D1},"<>Quitada")', 'calc', BRL, bold=True)
    ctx.name('Div_Total', S, f'E{t}')
    ctx.name('Div_ParcTot', S, f'G{t}')
    # ---- indicadores
    i0 = t + 2
    sh.put(f'B{i0}', '2. Indicador de endividamento', 'sec')
    ind = [('Renda mensal considerada', '=IF(Cfg_Renda>0,Cfg_Renda,K_Med3Rec)', BRL, 'Div_Renda'),
           ('Dívidas totais ÷ renda anual', '=IFERROR(Div_Total/(Div_Renda*12),0)', PCT, 'Div_PctDiv'),
           ('Parcelas mensais ÷ renda mensal', '=IFERROR(Div_ParcTot/Div_Renda,0)', PCT, 'Div_PctParc'),
           ('Limite de alerta', '=IF(Cfg_LimDiv=0,0.3,Cfg_LimDiv)', PCT, 'Div_Lim')]
    for i, (lbl, fx, fm, n) in enumerate(ind):
        r = i0 + 1 + i
        sh.put(f'B{r}', lbl, 'lbl')
        sh.put(f'D{r}', fx, 'calc', fm, bold=True)
        ctx.name(n, S, f'D{r}')
    sh.put(f'E{i0 + 2}', '=IF(Div_PctDiv>Div_Lim,"⚠ Acima do limite","✔ Ok")', 'calc')
    sh.put(f'E{i0 + 3}', '=IF(Div_PctParc>Div_Lim,"⚠ Acima do limite","✔ Ok")', 'calc')
    sh.cf_formula(f'E{i0 + 2}:E{i0 + 3}', f'LEFT(E{i0 + 2},1)="⚠"', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'E{i0 + 2}:E{i0 + 3}', f'LEFT(E{i0 + 2},1)="✔"', COR['verde_l'], COR['verde'], True)
    # ---- simulador
    s0 = i0 + 7
    sh.put(f'B{s0}', '3. Simulador de quitação: Bola de Neve x Avalanche', 'sec')
    sh.put(f'B{s0 + 1}', 'Quanto posso pagar a MAIS por mês (R$):', 'lblb')
    sh.put(f'D{s0 + 1}', None, 'in', BRLZ)
    ctx.name('Sim_Extra', S, f'D{s0 + 1}')
    ctx.name('Exemplo_SimExtra', S, f'D{s0 + 1}')
    sh.put(f'E{s0 + 1}', '← digite aqui. Bola de Neve: quita primeiro o MENOR saldo. Avalanche: primeiro os MAIORES juros. '
           'As parcelas liberadas rolam para a próxima dívida.', 'note')
    heads2 = ['Método', 'Meses para quitar tudo', 'Data de quitação', 'Total pago', 'Juros pagos', 'Economia de juros vs. mínimo',
              'Meses economizados']
    for i, h in enumerate(heads2):
        sh.put(f'{CL(2 + i)}{s0 + 3}', h, 'hdr')
    ws.row_dimensions[s0 + 3].height = 46
    meths = [('Pagamento mínimo', 'Min'), ('Bola de Neve (menor saldo primeiro)', 'Snow'), ('Avalanche (maior juros primeiro)', 'Ava')]
    for i, (lbl, key) in enumerate(meths):
        r = s0 + 4 + i
        sh.put(f'B{r}', lbl, 'calc', bold=True)
        sh.put(f'C{r}', f'=Sim{key}_Meses', 'calc', INT, h='center')
        sh.put(f'D{r}', f'=IF(ISNUMBER(C{r}),IF(C{r}=0,"—",EDATE(Cfg_Hoje,C{r})),"não quita em "&{HOR // 12}&" anos")', 'calc', MESANO, h='center')
        sh.put(f'E{r}', f'=Sim{key}_Pago', 'calc', BRL)
        sh.put(f'F{r}', f'=Sim{key}_Juros', 'calc', BRL)
        sh.put(f'G{r}', f'=SimMin_Juros-Sim{key}_Juros', 'calc', BRL, bold=True)
        sh.put(f'H{r}', f'=IF(AND(ISNUMBER(SimMin_Meses),ISNUMBER(C{r})),SimMin_Meses-C{r},"—")', 'calc', INT, h='center')
    sh.cf_formula(f'G{s0 + 5}:G{s0 + 6}', f'G{s0 + 5}>0', COR['verde_l'], COR['verde'], True)
    # quitação por dívida
    d0 = s0 + 8
    for i, h in enumerate(['Dívida', 'Saldo devedor', 'Quitação: mínimo', 'Quitação: Bola de Neve', 'Quitação: Avalanche',
                           'Ordem na Bola de Neve', 'Ordem na Avalanche']):
        sh.put(f'{CL(2 + i)}{d0}', h, 'hdr')
    ws.row_dimensions[d0].height = 32
    for d in range(1, N_DIV + 1):
        r = d0 + d
        m = D0 + d - 1
        sh.put(f'B{r}', f'=IF($B{m}="","",$B{m})', 'calc')
        sh.put(f'C{r}', f'=IF(OR($B{m}="",$K{m}="Quitada"),0,N($E{m}))', 'calc', BRL)
        for L, key, rk in (('D', 'Min', None), ('E', 'Snow', 'G'), ('F', 'Ava', 'H')):
            pos = f'{d}' if key == 'Min' else f'{rk}{r}'
            sh.put(f'{L}{r}', f'=IF(C{r}<=0,"",IFERROR(IF(INDEX(Sim{key}_MesesPos,{pos})>{HOR},"> {HOR // 12} anos",'
                   f'EDATE(Cfg_Hoje,INDEX(Sim{key}_MesesPos,{pos}))),""))', 'calc', MESANO, h='center')
        sh.put(f'G{r}', f'=IF(C{r}<=0,"",INDEX(Sim_RankSnow,{d}))', 'calc', INT, h='center')
        sh.put(f'H{r}', f'=IF(C{r}<=0,"",INDEX(Sim_RankAva,{d}))', 'calc', INT, h='center')
        sh.cf_formula(f'B{r}:H{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    # ---- gráfico de saldo devedor total
    ch = LineChart()
    ch.title = 'Saldo devedor total ao longo do tempo (R$)'
    ch.height, ch.width = 8.5, 20
    ch.y_axis.title = 'R$'
    ch.x_axis.title = 'Meses a partir de hoje'
    sim = ctx.wb['SIM_DÍVIDAS']
    for col, nm in ((ctx.sim_cols['Min'], 'Pagamento mínimo'), (ctx.sim_cols['Snow'], 'Bola de Neve'), (ctx.sim_cols['Ava'], 'Avalanche')):
        ref = Reference(sim, min_col=col, min_row=ctx.sim_chart_row0, max_row=ctx.sim_chart_row0 + 72)
        se = Series(ref, title=nm)
        se.smooth = False
        ch.series.append(se)
    ch.set_categories(Reference(sim, min_col=ctx.sim_cols['m'], min_row=ctx.sim_chart_row0, max_row=ctx.sim_chart_row0 + 72))
    ch.legend.position = 'b'
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblSkip = 6
    ch.visible_cells_only = False
    colors = [COR['vermelho'], COR['azul'], COR['verde']]
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.line.solidFill = c
        s.graphicalProperties.line.width = 28000
    ws.add_chart(ch, f'B{d0 + N_DIV + 3}')
    # ---- tabelas de amortização
    a0 = d0 + N_DIV + 3 + 19
    sh.put(f'B{a0}', '4. Tabela de amortização da dívida escolhida (Price x SAC)', 'sec')
    sh.put(f'B{a0 + 1}', 'Escolha a dívida:', 'lblb')
    sh.put(f'C{a0 + 1}', None, 'in')
    sh.dv(f'C{a0 + 1}', '=Lista_Dividas')
    ctx.name('Am_Sel', S, f'C{a0 + 1}')
    ctx.name('Exemplo_AmSel', S, f'C{a0 + 1}')
    sh.put(f'D{a0 + 1}', '← digite aqui (vazio = 1ª dívida). Usa saldo devedor, taxa e parcelas restantes da dívida (até 120).', 'note')
    sh.put(f'B{a0 + 2}', 'Saldo devedor:', 'lbl')
    sh.put(f'C{a0 + 2}', '=INDEX(Div_Saldo,Am_Idx)', 'calc', BRL)
    sh.put(f'D{a0 + 2}', 'Taxa mensal:', 'lbl', h='right')
    sh.put(f'E{a0 + 2}', '=INDEX(Div_Taxa,Am_Idx)', 'calc', '0.00%')
    sh.put(f'F{a0 + 2}', 'Parcelas:', 'lbl', h='right')
    sh.put(f'G{a0 + 2}', '=MIN(120,INDEX(Div_Rest,Am_Idx))', 'calc', INT)
    ctx.name('Am_Idx', S, f'I{a0 + 2}')
    sh.put(f'H{a0 + 2}', 'nº:', 'lbl', h='right')
    sh.put(f'I{a0 + 2}', '=IFERROR(MATCH(Am_Sel,Div_Nome,0),1)', 'calc', INT)
    ctx.name('Am_Saldo', S, f'C{a0 + 2}')
    ctx.name('Am_Taxa', S, f'E{a0 + 2}')
    ctx.name('Am_N', S, f'G{a0 + 2}')
    h0 = a0 + 4
    sh.put(f'C{h0 - 1}', 'PRICE (parcelas fixas)', 'lblb', color=COR['azul'])
    sh.put(f'H{h0 - 1}', 'SAC (amortização constante)', 'lblb', color=COR['roxo'])
    for L, t in zip('BCDEF', ['Nº', 'Parcela', 'Juros', 'Amortização', 'Saldo']):
        sh.put(f'{L}{h0}', t, 'hdr')
    for L, t in zip('GHIJK', ['', 'Parcela', 'Juros', 'Amortização', 'Saldo']):
        if t:
            sh.put(f'{L}{h0}', t, 'hdr')
    first = h0 + 1
    sh.put(f'B{h0 + 121}', 'Total', 'calc', bold=True)
    for k in range(1, 121):
        r = h0 + k
        prevF = 'Am_Saldo' if k == 1 else f'F{r - 1}'
        prevK = 'Am_Saldo' if k == 1 else f'K{r - 1}'
        sh.put(f'B{r}', f'=IF({k}<=Am_N,{k},"")', 'calc', INT, h='center')
        sh.put(f'C{r}', f'=IF(B{r}="","",IFERROR(PMT(Am_Taxa,Am_N,-Am_Saldo),Am_Saldo/Am_N))', 'calc', BRL)
        sh.put(f'D{r}', f'=IF(B{r}="","",{prevF}*Am_Taxa)', 'calc', BRL)
        sh.put(f'E{r}', f'=IF(B{r}="","",C{r}-D{r})', 'calc', BRL)
        sh.put(f'F{r}', f'=IF(B{r}="","",MAX(0,{prevF}-E{r}))', 'calc', BRL)
        sh.put(f'I{r}', f'=IF(B{r}="","",{prevK}*Am_Taxa)', 'calc', BRL)
        sh.put(f'J{r}', f'=IF(B{r}="","",Am_Saldo/Am_N)', 'calc', BRL)
        sh.put(f'H{r}', f'=IF(B{r}="","",I{r}+J{r})', 'calc', BRL)
        sh.put(f'K{r}', f'=IF(B{r}="","",MAX(0,{prevK}-J{r}))', 'calc', BRL)
    last = h0 + 120
    for L in 'CDEHIJ':
        sh.put(f'{L}{h0 + 121}', f'=SUM({L}{first}:{L}{last})', 'calc', BRL, bold=True)
    sh.put(f'B{h0 + 122}', 'Comparação: juros totais no Price menos no SAC (economia do SAC):', 'note')
    sh.put(f'H{h0 + 122}', f'=D{h0 + 121}-I{h0 + 121}', 'calc', BRL, bold=True)
    sh.finish(freeze='C9')
    return sh


# ======================================================================
def build_sim_dividas(ctx):
    """Motor da simulação (mínimo / bola de neve / avalanche) — aba oculta SIM_DÍVIDAS."""
    sh = Sh(ctx, 'SIM_DÍVIDAS')
    ws = sh.ws
    sh.header('SIM_DÍVIDAS (apoio)', 'Motor do simulador de quitação. Não precisa mexer. Esta aba pode ficar oculta.', ncols=10)
    # parâmetros por dívida (ordem original)
    sh.put('B6', 'Parâmetros (ordem de cadastro)', 'sec')
    for i, h in enumerate(['d', 'Saldo', 'Taxa', 'Parcela', 'Chave bola de neve', 'Ordem BN', 'Chave avalanche', 'Ordem AV']):
        sh.put(f'{CL(2 + i)}7', h, 'hdr')
    for d in range(1, N_DIV + 1):
        r = 7 + d
        m = D0 + d - 1
        sh.put(f'B{r}', d, 'calc')
        sh.put(f'C{r}', f"=IF(OR(INDEX(Div_Nome,{d})=\"\",INDEX(Div_Status,{d})=\"Quitada\"),0,N(INDEX(Div_Saldo,{d})))", 'calc', BRLZ)
        sh.put(f'D{r}', f'=N(INDEX(Div_Taxa,{d}))', 'calc', '0.00%')
        sh.put(f'E{r}', f'=IF(C{r}>0,N(INDEX(Div_Parc,{d})),0)', 'calc', BRLZ)
        sh.put(f'F{r}', f'=IF(C{r}>0,C{r}+B{r}/1000000,1000000000000+B{r})', 'calc', '0.000000')
        sh.put(f'G{r}', f'=SUMPRODUCT(--($F$8:$F$15<F{r}))+1', 'calc', INT)
        sh.put(f'H{r}', f'=IF(C{r}>0,1-D{r}+B{r}/1000000000,1000000000000+B{r})', 'calc', '0.000000000')
        sh.put(f'I{r}', f'=SUMPRODUCT(--($H$8:$H$15<H{r}))+1', 'calc', INT)
    ctx.name('Sim_RankSnow', 'SIM_DÍVIDAS', 'G8:G15')
    ctx.name('Sim_RankAva', 'SIM_DÍVIDAS', 'I8:I15')
    sh.put('B17', 'Soma das parcelas (dívidas ativas)', 'lbl')
    sh.put('E17', '=SUM(E8:E15)', 'calc', BRLZ)
    sh.put('B18', 'Total dos saldos', 'lbl')
    sh.put('C18', '=SUM(C8:C15)', 'calc', BRLZ)
    ctx.name('Sim_SomaParc', 'SIM_DÍVIDAS', 'E17')
    ctx.name('Sim_SomaSaldo', 'SIM_DÍVIDAS', 'C18')
    ctx.sim_cols = {}
    blocks = [('Min', 'Mínimo (sem extra, sem rolagem)', 'G'), ('Snow', 'Bola de Neve', 'G'), ('Ava', 'Avalanche', 'I')]
    row0 = 22
    for key, titulo, rankcol in blocks:
        sh.put(f'B{row0}', f'Método: {titulo}', 'sec')
        # parâmetros por posição p
        ph = row0 + 1
        sh.put(f'B{ph}', 'Posição p', 'lbl')
        sh.put(f'B{ph + 1}', 'Saldo inicial', 'lbl')
        sh.put(f'B{ph + 2}', 'Taxa mensal', 'lbl')
        sh.put(f'B{ph + 3}', 'Parcela', 'lbl')
        sh.put(f'B{ph + 4}', 'Meses até quitar', 'lbl')
        # colunas: B=mês, C=pool, D=min total, E..L bb, M..T ex, U..AB ba, AC=saldo total, AD=pago no mês
        BB = lambda p: CL(4 + p)          # E.. (p=1..8)
        EX = lambda p: CL(12 + p)         # M..
        BA = lambda p: CL(20 + p)         # U..
        for p in range(1, N_DIV + 1):
            c = CL(4 + p)   # coluna de parâmetros alinhada às colunas bb
            sh.put(f'{c}{ph}', p, 'calc', INT)
            ri = f'MATCH({p},{"Sim_RankAva" if rankcol == "I" else "Sim_RankSnow"},0)' if key != 'Min' else f'{p}'
            sh.put(f'{c}{ph + 1}', f'=INDEX($C$8:$C$15,{ri})', 'calc', BRLZ)
            sh.put(f'{c}{ph + 2}', f'=INDEX($D$8:$D$15,{ri})', 'calc', '0.00%')
            sh.put(f'{c}{ph + 3}', f'=INDEX($E$8:$E$15,{ri})', 'calc', BRLZ)
        hr = ph + 6
        heads = ['Mês', 'Extra disponível', 'Mínimos do mês'] + [f'Saldo+juros p{p}' for p in range(1, 9)] + \
                [f'Extra p{p}' for p in range(1, 9)] + [f'Saldo após p{p}' for p in range(1, 9)] + ['Saldo total', 'Pago no mês']
        for i, h in enumerate(heads):
            sh.put(f'{CL(2 + i)}{hr}', h, 'hdr', size=9)
        ws.row_dimensions[hr].height = 30
        r0 = hr + 1
        for m in range(0, HOR + 1):
            r = r0 + m
            sh.put(f'B{r}', m, 'calc', INT)
            if m == 0:
                for p in range(1, 9):
                    sh.put(f'{BA(p)}{r}', f'={BB(p)}${ph + 1}', 'calc', BRLZ)
                sh.put(f'AC{r}', f'=SUM(U{r}:AB{r})', 'calc', BRLZ)
                continue
            for p in range(1, 9):
                sh.put(f'{BB(p)}{r}', f'={BA(p)}{r - 1}*(1+{BB(p)}${ph + 2})', 'calc', BRLZ)
            mins = '+'.join(f'MIN({BB(p)}${ph + 3},{BB(p)}{r})' for p in range(1, 9))
            sh.put(f'D{r}', f'={mins}', 'calc', BRLZ)
            if key == 'Min':
                sh.put(f'C{r}', 0, 'calc', BRLZ)
            else:
                sh.put(f'C{r}', f'=MAX(0,Sim_SomaParc+N(Sim_Extra)-D{r})', 'calc', BRLZ)
            for p in range(1, 9):
                prev = '0' if p == 1 else f'SUM(${EX(1)}{r}:{EX(p - 1)}{r})'
                sh.put(f'{EX(p)}{r}', f'=MIN({BB(p)}{r}-MIN({BB(p)}${ph + 3},{BB(p)}{r}),MAX(0,$C{r}-{prev}))', 'calc', BRLZ)
                sh.put(f'{BA(p)}{r}', f'=IF({BB(p)}{r}-MIN({BB(p)}${ph + 3},{BB(p)}{r})-{EX(p)}{r}<0.005,0,'
                       f'{BB(p)}{r}-MIN({BB(p)}${ph + 3},{BB(p)}{r})-{EX(p)}{r})', 'calc', BRLZ)
            sh.put(f'AC{r}', f'=SUM(U{r}:AB{r})', 'calc', BRLZ)
            sh.put(f'AD{r}', f'=D{r}+SUM(M{r}:T{r})', 'calc', BRLZ)
        rl = r0 + HOR
        # meses até quitar por posição
        for p in range(1, 9):
            c = CL(4 + p)
            sh.put(f'{c}{ph + 4}', f'=IF({c}{ph + 1}<=0,0,IF({BA(p)}{rl}>0.005,{HOR + 1},COUNTIF({BA(p)}{r0 + 1}:{BA(p)}{rl},">0.005")+1))', 'calc', INT)
        ctx.name(f'Sim{key}_MesesPos', 'SIM_DÍVIDAS', f'E{ph + 4}:L{ph + 4}')
        sh.put(f'B{ph + 4}', 'Meses até quitar', 'lbl')
        ctx.name_formula(f'Sim{key}_Meses',
                         f"IF(MAX('SIM_DÍVIDAS'!$E${ph + 4}:$L${ph + 4})>{HOR},\"não quita\",MAX('SIM_DÍVIDAS'!$E${ph + 4}:$L${ph + 4}))")
        ctx.name_formula(f'Sim{key}_Pago', f"SUM('SIM_DÍVIDAS'!$AD${r0 + 1}:$AD${rl})")
        ctx.name_formula(f'Sim{key}_Juros',
                         f"MAX(0,SUM('SIM_DÍVIDAS'!$AD${r0 + 1}:$AD${rl})+'SIM_DÍVIDAS'!$AC${rl}-Sim_SomaSaldo)")
        ctx.sim_cols[key] = 29      # AC
        ctx.sim_chart_row0 = r0
        ctx.sim_cols['m'] = 2
        # cada bloco tem posição própria do gráfico: guardamos a linha inicial por método
        setattr(ctx, f'sim_row0_{key}', r0)
        row0 = r0 + HOR + 6
    ws.column_dimensions['B'].width = 24
    for c in range(3, 31):
        ws.column_dimensions[CL(c)].width = 13
    sh.finish(freeze='C8', protect=True)
    # série para gráfico (mínimo/bola/avalanche lado a lado)
    gc = 33
    sh.put(f'{CL(gc)}7', 'Mês', 'hdr')
    sh.put(f'{CL(gc + 1)}7', 'Mínimo', 'hdr')
    sh.put(f'{CL(gc + 2)}7', 'Bola de Neve', 'hdr')
    sh.put(f'{CL(gc + 3)}7', 'Avalanche', 'hdr')
    for m in range(0, 73):
        r = 8 + m
        sh.put(f'{CL(gc)}{r}', m, 'calc', INT)
        for j, key in enumerate(('Min', 'Snow', 'Ava')):
            sh.put(f'{CL(gc + 1 + j)}{r}', f'=AC{getattr(ctx, f"sim_row0_{key}") + m}', 'calc', BRLZ)
    ctx.sim_cols = {'m': gc, 'Min': gc + 1, 'Snow': gc + 2, 'Ava': gc + 3}
    ctx.sim_chart_row0 = 8
