"""ORÇAMENTO e FLUXO DE CAIXA."""
from common import *
from part_base import MESES_AB, N_CATD, N_CATR


def cat_sum(tipo, catref, idx):
    return SM(tipo, idx, f',L_Cat,{catref}')


def build_orcamento(ctx):
    sh = Sh(ctx, 'ORÇAMENTO')
    ws = sh.ws
    sh.header('ORÇAMENTO', 'Planeje por categoria e por mês (células azuis) e compare com o realizado vindo de LANÇAMENTOS. '
              'Verde até 80%, amarelo de 80 a 100%, vermelho acima de 100%. O mês analisado é o escolhido no DASHBOARD.',
              ncols=15)
    sh.widths({'B': 30, **{CL(c): 13.5 for c in range(3, 15)}, 'K': 15, 'O': 15})
    sh.put('B5', 'Mês analisado:', 'lblb')
    sh.put('C5', '=Sel_Label', 'calc', bold=True)
    sh.merge('C5:D5')
    sh.put('F5', 'Orçamento base zero:', 'lblb')
    sh.merge('F5:G5')
    sh.put('H5', '=IF(Cfg_BaseZero="Sim","SIM","Não")', 'calc', bold=True, h='center')
    sh.put('I5', 'Para copiar o orçamento do mês anterior: selecione a coluna do mês anterior (células azuis), Ctrl+C, '
           'clique no mês desejado e Ctrl+V.', 'note')
    # ------------------------------------------------ 1. planejamento
    sh.put('B7', '1. Planejamento mensal  ⬇ digite aqui (valores planejados por categoria)', 'sec')
    sh.put('B8', 'Categoria', 'hdr')
    for k, m in enumerate(MESES_AB):
        sh.put(f'{CL(3 + k)}8', m, 'hdr')
    sh.put('O8', 'Total no ano', 'hdr')
    sh.put('B9', 'RECEITAS', 'lblb', color=COR['verde'])
    r_r0 = 10
    for k in range(1, N_CATR + 1):
        r = r_r0 + k - 1
        sh.put(f'B{r}', f'=IF(INDEX(CatR_Nome,{k})="","",INDEX(CatR_Nome,{k}))', 'calc')
        sh.block(f'C{r}:N{r}', 'in', BRL)
        sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRL)
    r_r1 = r_r0 + N_CATR - 1
    r_tr = r_r1 + 1
    sh.put(f'B{r_tr}', 'Total de receitas planejadas', 'calc', bold=True)
    for k in range(13):
        L = CL(3 + k)
        sh.put(f'{L}{r_tr}', f'=SUM({L}{r_r0}:{L}{r_r1})', 'calc', BRL, bold=True)
    r = r_tr + 1
    sh.put(f'B{r}', 'DESPESAS', 'lblb', color=COR['vermelho'])
    r_d0 = r + 1
    for k in range(1, N_CATD + 1):
        rr = r_d0 + k - 1
        sh.put(f'B{rr}', f'=IF(INDEX(CatD_Nome,{k})="","",INDEX(CatD_Nome,{k}))', 'calc')
        sh.block(f'C{rr}:N{rr}', 'in', BRL)
        sh.put(f'O{rr}', f'=SUM(C{rr}:N{rr})', 'calc', BRL)
    r_d1 = r_d0 + N_CATD - 1
    r_td = r_d1 + 1
    sh.put(f'B{r_td}', 'Total de despesas planejadas', 'calc', bold=True)
    for k in range(13):
        L = CL(3 + k)
        sh.put(f'{L}{r_td}', f'=SUM({L}{r_d0}:{L}{r_d1})', 'calc', BRL, bold=True)
    r_sb = r_td + 1
    sh.put(f'B{r_sb}', 'Sobra planejada', 'calc', bold=True)
    sh.put(f'B{r_sb + 1}', 'Situação (base zero)', 'calc')
    for k in range(13):
        L = CL(3 + k)
        sh.put(f'{L}{r_sb}', f'={L}{r_tr}-{L}{r_td}', 'calc', BRLZ, bold=True)
        if k < 12:
            sh.put(f'{L}{r_sb + 1}', f'=IF(Cfg_BaseZero<>"Sim","",IF(ROUND({L}{r_sb},2)=0,"✔ zerado",'
                   f'IF({L}{r_sb}>0,"alocar "&FIXED({L}{r_sb},0),"estourou")))', 'calc', h='center', size=10)
    sh.cf_formula(f'C{r_sb}:O{r_sb}', f'C{r_sb}<0', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'C{r_sb}:O{r_sb}', f'C{r_sb}>=0', None, COR['verde'], True)
    ctx.name('Orc_GridR', 'ORÇAMENTO', f'C{r_r0}:N{r_r1}')
    ctx.name('Orc_GridD', 'ORÇAMENTO', f'C{r_d0}:N{r_d1}')
    ctx.name('Orc_TotRec', 'ORÇAMENTO', f'C{r_tr}:N{r_tr}')
    ctx.name('Orc_TotDesp', 'ORÇAMENTO', f'C{r_td}:N{r_td}')
    ctx.clear_ranges.append(('ORÇAMENTO', f'C{r_r0}:N{r_r1}', 'Orçamento de receitas'))
    ctx.clear_ranges.append(('ORÇAMENTO', f'C{r_d0}:N{r_d1}', 'Orçamento de despesas'))
    ctx.name('Exemplo_OrcReceitas', 'ORÇAMENTO', f'C{r_r0}:N{r_r1}')
    ctx.name('Exemplo_OrcDespesas', 'ORÇAMENTO', f'C{r_d0}:N{r_d1}')
    ctx.orc_rows = dict(r0=r_r0, r1=r_r1, d0=r_d0, d1=r_d1)
    # blank slots apagados
    for rr in list(range(r_r0, r_r1 + 1)) + list(range(r_d0, r_d1 + 1)):
        sh.cf_formula(f'B{rr}:O{rr}', f'$B{rr}=""', COR['cinza'], COR['cinza'])
    # ------------------------------------------------ 2. análise do mês
    a = r_sb + 4
    sh.put(f'B{a}', '2. Planejado x realizado no mês analisado', 'sec')
    heads = ['Categoria', 'Planejado', 'Realizado', 'Diferença (plan. − real.)', '% usado', 'Média dos 3 meses anteriores',
             'Sugestão de orçamento', 'Mês anterior', 'Mesmo mês, ano anterior', 'Situação']
    for i, h in enumerate(heads):
        sh.put(f'{CL(2 + i)}{a + 1}', h, 'hdr')
    ws.row_dimensions[a + 1].height = 46

    def analise(r0, n, tipo, grid0, cat_nome, nm, titulo, fmt_d):
        sh.put(f'B{r0}', titulo, 'lblb', color=COR['verde'] if tipo == 'Receita' else COR['vermelho'])
        first = r0 + 1
        for k in range(1, n + 1):
            r = first + k - 1
            g = grid0 + k - 1
            sh.put(f'B{r}', f'=B{g}', 'calc')
            sh.put(f'C{r}', f'=IF($B{r}="",0,INDEX($C{g}:$N{g},Sel_Mes))', 'calc', BRL)
            sh.put(f'D{r}', f'=IF($B{r}="",0,{cat_sum(tipo, f"$B{r}", "Sel_Idx")})', 'calc', BRL)
            if tipo == 'Despesa':
                sh.put(f'E{r}', f'=C{r}-D{r}', 'calc', BRL)
            else:
                sh.put(f'E{r}', f'=D{r}-C{r}', 'calc', BRL)
            sh.put(f'F{r}', f'=IF(C{r}=0,0,D{r}/C{r})', 'calc', PCT)
            m3 = '+'.join(cat_sum(tipo, f'$B{r}', f'Sel_Idx-{j}') for j in (1, 2, 3))
            sh.put(f'G{r}', f'=IF($B{r}="",0,({m3})/3)', 'calc', BRL)
            sh.put(f'H{r}', f'=IF(G{r}=0,0,ROUNDUP(G{r}/10,0)*10)', 'calc', BRL)
            sh.put(f'I{r}', f'=IF($B{r}="",0,{cat_sum(tipo, f"$B{r}", "Sel_Idx-1")})', 'calc', BRL)
            sh.put(f'J{r}', f'=IF($B{r}="",0,{cat_sum(tipo, f"$B{r}", "Sel_Idx-12")})', 'calc', BRL)
            if tipo == 'Despesa':
                sh.put(f'K{r}', f'=IF($B{r}="","",IF(C{r}=0,IF(D{r}>0,"Sem orçamento","—"),IF(F{r}>1,"Estourou",'
                       f'IF(F{r}>=0.8,"Atenção","Dentro"))))', 'calc', h='center')
            else:
                sh.put(f'K{r}', f'=IF($B{r}="","",IF(C{r}=0,IF(D{r}>0,"Sem meta","—"),IF(F{r}>=1,"Meta atingida",'
                       f'IF(F{r}>=0.8,"Quase lá","Abaixo"))))', 'calc', h='center')
        last = first + n - 1
        tot = last + 1
        sh.put(f'B{tot}', 'Total', 'calc', bold=True)
        for L in 'CDEGHIJ':
            sh.put(f'{L}{tot}', f'=SUM({L}{first}:{L}{last})', 'calc', BRL, bold=True)
        sh.put(f'F{tot}', f'=IF(C{tot}=0,0,D{tot}/C{tot})', 'calc', PCT, bold=True)
        sh.put(f'K{tot}', None, 'calc')
        for L, nn in zip('BCDEFGHIJK', ['Cat', 'Plan', 'Real', 'Dif', 'Pct', 'Med3', 'Sug', 'Ant', 'AA', 'Sit']):
            ctx.name(f'{nm}_{nn}', 'ORÇAMENTO', f'{L}{first}:{L}{last}')
        for L, nn in zip('CDEFGHIJ', ['Plan', 'Real', 'Dif', 'Pct', 'Med3', 'Sug', 'Ant', 'AA']):
            ctx.name(f'{nm}_Tot{nn}', 'ORÇAMENTO', f'{L}{tot}')
        for rr in range(first, last + 1):
            sh.cf_formula(f'B{rr}:K{rr}', f'$B{rr}=""', COR['cinza'], COR['cinza'])
        return first, last, tot

    rf, rl, rt = analise(a + 2, N_CATR, 'Receita', r_r0, None, 'OrcR', 'RECEITAS', None)
    sh.cf_formula(f'F{rf}:F{rl}', f'AND($B{rf}<>"",$C{rf}>0,F{rf}>=1)', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'F{rf}:F{rl}', f'AND($B{rf}<>"",$C{rf}>0,F{rf}>=0.8,F{rf}<1)', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'F{rf}:F{rl}', f'AND($B{rf}<>"",$C{rf}>0,F{rf}<0.8)', COR['vermelho_l'], COR['vermelho'], True)
    df, dl, dt = analise(rt + 2, N_CATD, 'Despesa', r_d0, None, 'Orc', 'DESPESAS', None)
    for rng_, k in ((f'F{df}:F{dt}', df), (f'K{df}:K{dl}', df)):
        pass
    sh.cf_formula(f'F{df}:F{dt}', f'AND($C{df}>0,F{df}<0.8)', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'F{df}:F{dt}', f'AND($C{df}>0,F{df}>=0.8,F{df}<=1)', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'F{df}:F{dt}', f'AND($C{df}>0,F{df}>1)', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'K{df}:K{dl}', f'K{df}="Dentro"', None, COR['verde'], True)
    sh.cf_formula(f'K{df}:K{dl}', f'K{df}="Atenção"', None, '9A5B00', True)
    sh.cf_formula(f'K{df}:K{dl}', f'OR(K{df}="Estourou",K{df}="Sem orçamento")', None, COR['vermelho'], True)
    sh.cf_formula(f'E{df}:E{dt}', f'E{df}<0', None, COR['vermelho'], True)
    sh.cf_formula(f'K{rf}:K{rl}', f'K{rf}="Meta atingida"', None, COR['verde'], True)
    # dentro do orçamento (para a nota de saúde)
    ctx.name_formula('Orc_QtdPlan', 'COUNTIF(Orc_Plan,">0")')
    # ------------------------------------------------ 3. 50/30/20
    b = dt + 3
    sh.put(f'B{b}', '3. Regra 50/30/20 no mês analisado (sobre a receita realizada)', 'sec')
    for i, h in enumerate(['Grupo', 'Ideal %', 'Ideal R$', 'Real R$', 'Real %', 'Diferença (p.p.)', 'Situação']):
        sh.put(f'{CL(2 + i)}{b + 1}', h, 'hdr')
    grupos = [('Necessidades', 0.5), ('Desejos', 0.3), ('Poupança', 0.2)]
    for i, (g, p) in enumerate(grupos):
        r = b + 2 + i
        sh.put(f'B{r}', g, 'calc')
        sh.put(f'C{r}', p, 'calc', PCT)
        sh.put(f'D{r}', f'=C{r}*K_Rec', 'calc', BRL)
        if g == 'Poupança':
            sh.put(f'E{r}', f'=K_Rec-E{b + 2}-E{b + 3}', 'calc', BRLZ)
        else:
            sh.put(f'E{r}', f'=SUMIF(CatD_Regra,B{r},Orc_Real)', 'calc', BRL)
        sh.put(f'F{r}', f'=IFERROR(E{r}/K_Rec,0)', 'calc', PCT)
        sh.put(f'G{r}', f'=(F{r}-C{r})*100', 'calc', '0.0;-0.0;0.0')
        if g == 'Poupança':
            sh.put(f'H{r}', f'=IF(K_Rec=0,"—",IF(F{r}>=C{r},"✔ Ok","Abaixo do ideal"))', 'calc', h='center')
        else:
            sh.put(f'H{r}', f'=IF(K_Rec=0,"—",IF(F{r}<=C{r},"✔ Ok","Acima do ideal"))', 'calc', h='center')
    ctx.name('R50_Grupo', 'ORÇAMENTO', f'B{b + 2}:B{b + 4}')
    ctx.name('R50_Ideal', 'ORÇAMENTO', f'C{b + 2}:C{b + 4}')
    ctx.name('R50_Real', 'ORÇAMENTO', f'F{b + 2}:F{b + 4}')
    ctx.name('R50_RealRS', 'ORÇAMENTO', f'E{b + 2}:E{b + 4}')
    sh.cf_formula(f'H{b + 2}:H{b + 4}', f'LEFT(H{b + 2},1)="✔"', None, COR['verde'], True)
    sh.cf_formula(f'H{b + 2}:H{b + 4}', f'OR(H{b + 2}="Abaixo do ideal",H{b + 2}="Acima do ideal")', None, COR['vermelho'], True)
    sh.put(f'B{b + 6}', 'Necessidades = moradia, saúde, alimentação... · Desejos = lazer, assinaturas... · Poupança = o que sobra '
           '(receitas − necessidades − desejos). A regra de cada categoria é definida em CONFIGURAÇÕES.', 'note')
    sh.finish(freeze='C9')
    return sh


# ======================================================================
def build_fluxo(ctx):
    sh = Sh(ctx, 'FLUXO DE CAIXA')
    ws = sh.ws
    sh.header('FLUXO DE CAIXA', 'Matriz de 12 meses do ano selecionado no DASHBOARD: receitas e despesas por categoria, '
              'saldo, média, maior e menor mês, realizado x previsto e sazonalidade. Tudo calculado, nada a digitar.',
              ncols=19)
    sh.widths({'B': 34, **{CL(c): 13.5 for c in range(3, 15)}, 'O': 15, 'P': 14, 'Q': 14, 'R': 14, 'S': 10})
    sh.put('B5', 'Ano:', 'lblb')
    sh.put('C5', '=Sel_Ano', 'calc', INT, bold=True)
    # linha oculta com índices dos meses
    for k in range(12):
        sh.put(f'{CL(3 + k)}7', f'=Sel_Ano*12+{k + 1}', 'calc', INT)
    ws.row_dimensions[7].hidden = True
    sh.put('B8', 'Categoria', 'hdr')
    for k, m in enumerate(MESES_AB):
        sh.put(f'{CL(3 + k)}8', m, 'hdr')
    for L, t in zip('OPQRS', ['Total', 'Média mensal', 'Maior mês', 'Menor mês', 'Mês do maior']):
        sh.put(f'{L}8', t, 'hdr')
    ctx.name('Flux_Mes', 'FLUXO DE CAIXA', 'C8:N8')

    def linhas(r0, n, tipo, cats, nm, titulo, cor):
        sh.put(f'B{r0}', titulo, 'lblb', color=cor)
        first = r0 + 1
        for k in range(1, n + 1):
            r = first + k - 1
            sh.put(f'B{r}', f'=IF(INDEX({cats},{k})="","",INDEX({cats},{k}))', 'calc')
            for m in range(12):
                L = CL(3 + m)
                sh.put(f'{L}{r}', f'=IF($B{r}="",0,{SM(tipo, f"{L}$7", f",L_Cat,$B{r}")})', 'calc', BRL)
            sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRL)
            sh.put(f'P{r}', f'=IFERROR(SUMIF(C{r}:N{r},"<>0")/COUNTIF(C{r}:N{r},"<>0"),0)', 'calc', BRL)
            sh.put(f'Q{r}', f'=MAX(C{r}:N{r})', 'calc', BRL)
            sh.put(f'R{r}', f'=IFERROR(SMALL(C{r}:N{r},COUNTIF(C{r}:N{r},0)+1),0)', 'calc', BRL)
            sh.put(f'S{r}', f'=IF(Q{r}=0,"",INDEX(Lista_MesesAb,MATCH(Q{r},C{r}:N{r},0)))', 'calc', h='center')
            sh.cf_formula(f'B{r}:S{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
        last = first + n - 1
        t = last + 1
        sh.put(f'B{t}', 'Total de ' + titulo.lower(), 'calc', bold=True)
        for m in range(13):
            L = CL(3 + m)
            sh.put(f'{L}{t}', f'=SUM({L}{first}:{L}{last})', 'calc', BRL, bold=True)
        sh.put(f'P{t}', f'=IFERROR(SUMIF(C{t}:N{t},"<>0")/COUNTIF(C{t}:N{t},"<>0"),0)', 'calc', BRL, bold=True)
        sh.put(f'Q{t}', f'=MAX(C{t}:N{t})', 'calc', BRL, bold=True)
        sh.put(f'R{t}', f'=MIN(C{t}:N{t})', 'calc', BRL, bold=True)
        sh.put(f'S{t}', f'=IF(Q{t}=0,"",INDEX(Lista_MesesAb,MATCH(Q{t},C{t}:N{t},0)))', 'calc', h='center')
        ctx.name(f'Flux_{nm}', 'FLUXO DE CAIXA', f'C{t}:N{t}')
        ctx.name(f'Flux_{nm}Tot', 'FLUXO DE CAIXA', f'O{t}')
        return first, last, t

    rf, rl, rt = linhas(9, N_CATR, 'Receita', 'CatR_Nome', 'Rec', 'RECEITAS', COR['verde'])
    df, dl, dt = linhas(rt + 1, N_CATD, 'Despesa', 'CatD_Nome', 'Desp', 'DESPESAS', COR['vermelho'])
    r = dt + 1
    sh.put(f'B{r}', 'Aportes / investimentos', 'calc', bold=True, color=COR['azul'])
    for m in range(12):
        L = CL(3 + m)
        sh.put(f'{L}{r}', f'={SM("Investimento", f"{L}$7")}', 'calc', BRL)
    sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRL, bold=True)
    sh.put(f'P{r}', f'=IFERROR(SUMIF(C{r}:N{r},"<>0")/COUNTIF(C{r}:N{r},"<>0"),0)', 'calc', BRL)
    sh.put(f'Q{r}', f'=MAX(C{r}:N{r})', 'calc', BRL)
    sh.put(f'R{r}', f'=MIN(C{r}:N{r})', 'calc', BRL)
    ctx.name('Flux_Apo', 'FLUXO DE CAIXA', f'C{r}:N{r}')
    r_apo = r
    r += 2
    sh.put(f'B{r}', 'SALDO DO MÊS (receitas − despesas)', 'calc', bold=True)
    for m in range(12):
        L = CL(3 + m)
        sh.put(f'{L}{r}', f'={L}{rt}-{L}{dt}', 'calc', BRLZ, bold=True)
    sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRLZ, bold=True)
    sh.put(f'P{r}', f'=AVERAGE(C{r}:N{r})', 'calc', BRLZ)
    sh.put(f'Q{r}', f'=MAX(C{r}:N{r})', 'calc', BRLZ)
    sh.put(f'R{r}', f'=MIN(C{r}:N{r})', 'calc', BRLZ)
    sh.put(f'S{r}', f'=INDEX(Lista_MesesAb,MATCH(Q{r},C{r}:N{r},0))', 'calc', h='center')
    r_sd = r
    sh.cf_formula(f'C{r}:R{r}', f'C{r}<0', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'C{r}:R{r}', f'C{r}>=0', COR['verde_l'], COR['verde'], True)
    r += 1
    sh.put(f'B{r}', 'Saldo acumulado no ano', 'calc', bold=True)
    for m in range(12):
        L = CL(3 + m)
        sh.put(f'{L}{r}', f'=SUM($C{r_sd}:{L}{r_sd})', 'calc', BRLZ, bold=True)
    sh.cf_formula(f'C{r}:N{r}', f'C{r}<0', None, COR['vermelho'], True)
    ctx.name('Flux_Saldo', 'FLUXO DE CAIXA', f'C{r_sd}:N{r_sd}')
    ctx.name('Flux_Acum', 'FLUXO DE CAIXA', f'C{r}:N{r}')
    r += 1
    sh.put(f'B{r}', 'Saldo após aportes (saldo − investimentos)', 'calc')
    for m in range(12):
        L = CL(3 + m)
        sh.put(f'{L}{r}', f'={L}{r_sd}-{L}{r_apo}', 'calc', BRLZ)
    sh.put(f'O{r}', f'=SUM(C{r}:N{r})', 'calc', BRLZ)
    # ---- realizado x previsto
    r += 2
    sh.put(f'B{r}', 'Realizado x previsto (previsto = orçamento planejado em ORÇAMENTO)', 'sec')
    r += 1
    sh.put(f'B{r}', 'Item', 'hdr')
    for k, m in enumerate(MESES_AB):
        sh.put(f'{CL(3 + k)}{r}', m, 'hdr')
    sh.put(f'O{r}', 'Total', 'hdr')
    linhas_rp = [('Receitas previstas', lambda k: f'=INDEX(Orc_TotRec,{k + 1})', BRL),
                 ('Receitas realizadas', lambda k: f'={CL(3 + k)}{rt}', BRL),
                 ('Diferença receitas', None, BRLZ),
                 ('Despesas previstas', lambda k: f'=INDEX(Orc_TotDesp,{k + 1})', BRL),
                 ('Despesas realizadas', lambda k: f'={CL(3 + k)}{dt}', BRL),
                 ('Diferença despesas (prev. − real.)', None, BRLZ),
                 ('Saldo previsto', None, BRLZ), ('Saldo realizado', None, BRLZ)]
    base = r + 1
    for i, (lbl, fn, fm) in enumerate(linhas_rp):
        rr = base + i
        sh.put(f'B{rr}', lbl, 'calc', bold=(i in (2, 5, 6, 7)))
        for k in range(12):
            L = CL(3 + k)
            if fn:
                f = fn(k)
            elif i == 2:
                f = f'={L}{base + 1}-{L}{base}'
            elif i == 5:
                f = f'={L}{base + 3}-{L}{base + 4}'
            elif i == 6:
                f = f'={L}{base}-{L}{base + 3}'
            else:
                f = f'={L}{base + 1}-{L}{base + 4}'
            sh.put(f'{L}{rr}', f, 'calc', fm, bold=(i in (2, 5, 6, 7)))
        sh.put(f'O{rr}', f'=SUM(C{rr}:N{rr})', 'calc', fm, bold=True)
    sh.cf_formula(f'C{base + 2}:O{base + 2}', f'C{base + 2}<0', None, COR['vermelho'], True)
    sh.cf_formula(f'C{base + 5}:O{base + 5}', f'C{base + 5}<0', None, COR['vermelho'], True)
    sh.cf_formula(f'C{base + 2}:O{base + 2}', f'C{base + 2}>=0', None, COR['verde'], True)
    sh.cf_formula(f'C{base + 5}:O{base + 5}', f'C{base + 5}>=0', None, COR['verde'], True)
    # ---- sazonalidade
    r = base + len(linhas_rp) + 1
    sh.put(f'B{r}', 'Sazonalidade: meses mais caros do ano', 'sec')
    r += 1
    sh.put(f'B{r}', 'Item', 'hdr')
    for k, m in enumerate(MESES_AB):
        sh.put(f'{CL(3 + k)}{r}', m, 'hdr')
    sh.put(f'B{r + 1}', 'Índice (despesa do mês ÷ média do ano)', 'calc')
    sh.put(f'B{r + 2}', 'Ranking (1 = mais caro)', 'calc')
    for k in range(12):
        L = CL(3 + k)
        sh.put(f'{L}{r + 1}', f'=IFERROR({L}{dt}/AVERAGEIF($C{dt}:$N{dt},"<>0"),0)', 'calc', '0.00')
        sh.put(f'{L}{r + 2}', f'=IF({L}{dt}=0,"",RANK({L}{dt},$C{dt}:$N{dt})+COUNTIF($C{dt}:{L}{dt},{L}{dt})-1)', 'calc', INT, h='center')
    sh.cf(f'C{r + 1}:N{r + 1}', ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='F2A93B'))
    sh.put(f'B{r + 3}', '3 meses mais caros:', 'calc', bold=True)
    for i in range(3):
        L = CL(3 + i * 3)
        sh.put(f'{L}{r + 3}', f'=IFERROR(INDEX(Lista_MesesAb,MATCH({i + 1},$C{r + 2}:$N{r + 2},0))&" ("&FIXED(INDEX($C{dt}:$N{dt},MATCH({i + 1},$C{r + 2}:$N{r + 2},0)),0)&")","")', 'calc')
        sh.merge(f'{L}{r + 3}:{CL(4 + i * 3)}{r + 3}')
    sh.finish(freeze='C9')
    return sh
