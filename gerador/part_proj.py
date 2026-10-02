"""PROJEÇÕES E CENÁRIOS, CALENDÁRIO DE CONTAS e HISTÓRICO E ANÁLISES."""
from openpyxl.chart import LineChart, Reference, Series
from common import *
from part_base import MESES_AB, N_CATD, N_CATR, N_TAG, N_FORMA, N_METAS, N_REC, REC
from part_cards_debts import fat, venc, N_DIV
from part_metas import M0, M1

NP = 24     # meses projetados


def build_calc_extra(ctx):
    """Médias históricas por categoria (3 meses fechados antes do mês atual) — base das projeções."""
    c = ctx.calc
    sh, ws = c.sh, c.ws
    sh.put('Q5', 'Média 3 meses fechados por categoria de despesa', 'sec')
    for i, h in enumerate(['Categoria', 'Natureza', 'Média 3m', 'Chave rank (mês sel.)', 'Valor mês sel.']):
        sh.put(f'{CL(17 + i)}6', h, 'hdr')
    for k in range(1, N_CATD + 1):
        r = 6 + k
        sh.put(f'Q{r}', f'=IF(INDEX(CatD_Nome,{k})="","",INDEX(CatD_Nome,{k}))', 'calc')
        sh.put(f'R{r}', f'=INDEX(CatD_Nat,{k})&""', 'calc')
        m3 = '+'.join(SM('Despesa', f'Cfg_Idx-{j}', f',L_Cat,$Q{r}') for j in (1, 2, 3))
        sh.put(f'S{r}', f'=IF(Q{r}="",0,({m3})/3)', 'calc', BRLZ)
        sh.put(f'T{r}', f'=IF(INDEX(Orc_Real,{k})>0,INDEX(Orc_Real,{k})+ROW()/1000000000,0)', 'calc', '0.000000000')
        sh.put(f'U{r}', f'=INDEX(Orc_Real,{k})', 'calc', BRLZ)
    r1 = 6 + N_CATD
    ctx.name('Cat_Med3', 'CÁLCULOS', f'S7:S{r1}')
    ctx.name('Cat_Nat', 'CÁLCULOS', f'R7:R{r1}')
    ctx.name('Cat_Chave', 'CÁLCULOS', f'T7:T{r1}')
    # ranking top-8 + outros para o gráfico de rosca (mês selecionado)
    sh.put('W5', 'Ranking de despesas do mês (para gráficos)', 'sec')
    for i, h in enumerate(['Posição', 'Categoria', 'Valor']):
        sh.put(f'{CL(23 + i)}6', h, 'hdr')
    for p in range(1, 10):
        r = 6 + p
        sh.put(f'W{r}', p, 'calc')
        if p <= 8:
            sh.put(f'Y{r}', f'=IFERROR(INDEX(U$7:U${r1},MATCH(LARGE(T$7:T${r1},{p}),T$7:T${r1},0)),0)', 'calc', BRLZ)
            sh.put(f'X{r}', f'=IF(Y{r}>0,INDEX(Q$7:Q${r1},MATCH(LARGE(T$7:T${r1},{p}),T$7:T${r1},0)),"—")', 'calc')
        else:
            sh.put(f'X{r}', 'Outras', 'calc')
            sh.put(f'Y{r}', f'=MAX(0,K_Desp-SUM(Y7:Y14))', 'calc', BRLZ)
    ctx.name('Rk_Nome', 'CÁLCULOS', 'X7:X15')
    ctx.name('Rk_Val', 'CÁLCULOS', 'Y7:Y15')
    # ranking orçado x realizado (top 10 por max(planejado, realizado))
    sh.put('AA5', 'Orçado x realizado (top 10)', 'sec')
    for i, h in enumerate(['Chave', 'Posição', 'Categoria', 'Orçado', 'Realizado']):
        sh.put(f'{CL(27 + i)}6', h, 'hdr')
    for k in range(1, N_CATD + 1):
        r = 6 + k
        sh.put(f'AA{r}', f'=IF(OR(INDEX(Orc_Plan,{k})>0,INDEX(Orc_Real,{k})>0),MAX(INDEX(Orc_Plan,{k}),INDEX(Orc_Real,{k}))+ROW()/1000000000,0)', 'calc', '0.000000000')
    for p in range(1, 11):
        r = 6 + p
        sh.put(f'AB{r}', p, 'calc')
        m = f'MATCH(LARGE(AA$7:AA${r1},{p}),AA$7:AA${r1},0)'
        sh.put(f'AC{r}', f'=IFERROR(IF(LARGE(AA$7:AA${r1},{p})>0,INDEX(Q$7:Q${r1},{m}),"—"),"—")', 'calc')
        sh.put(f'AD{r}', f'=IFERROR(IF(LARGE(AA$7:AA${r1},{p})>0,INDEX(Orc_Plan,{m}),0),0)', 'calc', BRLZ)
        sh.put(f'AE{r}', f'=IFERROR(IF(LARGE(AA$7:AA${r1},{p})>0,INDEX(Orc_Real,{m}),0),0)', 'calc', BRLZ)
    ctx.name('OR_Nome', 'CÁLCULOS', 'AC7:AC16')
    ctx.name('OR_Orc', 'CÁLCULOS', 'AD7:AD16')
    ctx.name('OR_Real', 'CÁLCULOS', 'AE7:AE16')
    for L in ['Q', 'R', 'S', 'T', 'U', 'W', 'X', 'Y', 'AA', 'AB', 'AC', 'AD', 'AE']:
        ws.column_dimensions[L].width = 16
    c.sec('Médias para projeções')
    c.val('K_Med3Desp', f'=({SM("Despesa", "Cfg_Idx-1")}+{SM("Despesa", "Cfg_Idx-2")}+{SM("Despesa", "Cfg_Idx-3")})/3',
          'Média mensal de despesas (3 meses fechados)', BRLZ)
    c.val('K_Med3Rec', f'=({SM("Receita", "Cfg_Idx-1")}+{SM("Receita", "Cfg_Idx-2")}+{SM("Receita", "Cfg_Idx-3")})/3',
          'Média mensal de receitas (3 meses fechados)', BRLZ)
    c.val('K_VarMed', '=SUMIF(Cat_Nat,"Variável",Cat_Med3)', 'Média mensal das despesas variáveis', BRLZ)
    c.val('K_FixaMed', '=SUMIF(Cat_Nat,"Fixa",Cat_Med3)', 'Média mensal das despesas fixas', BRLZ)
    c.val('K_RendaBase', '=IF(Cfg_Renda>0,Cfg_Renda,K_Med3Rec)', 'Renda mensal considerada', BRLZ)
    c.val('K_SaldoLivre', '=MAX(0,K_RendaBase-K_Med3Desp)', 'Saldo livre mensal (renda − despesas médias)', BRLZ)


def build_proj(ctx):
    sh = Sh(ctx, 'PROJEÇÕES E CENÁRIOS')
    ws = sh.ws
    S = 'PROJEÇÕES E CENÁRIOS'
    sh.header('PROJEÇÕES E CENÁRIOS', 'Projeção de 12 e 24 meses a partir do mês atual: recorrentes, parcelas e lançamentos futuros, média '
              'histórica das despesas variáveis e inflação. Três cenários editáveis e um simulador "E se?".', ncols=28)
    sh.widths({'B': 40, 'C': 15, 'D': 15, **{CL(c): 12.5 for c in range(5, 5 + NP)}})
    sh.put('B6', '1. Cenários (edite os percentuais)', 'sec')
    for i, h in enumerate(['Cenário', 'Δ Despesas', 'Δ Receitas']):
        sh.put(f'{CL(2 + i)}7', h, 'hdr')
    cen = [('Realista', 0.0, 0.0), ('Otimista', -0.10, 0.05), ('Pessimista', 0.10, -0.10)]
    for i, (n, d, rcv) in enumerate(cen):
        r = 8 + i
        sh.put(f'B{r}', n, 'calc', bold=True)
        sh.put(f'C{r}', d, 'in', '0.0%')
        sh.put(f'D{r}', rcv, 'in', '0.0%')
    ctx.name('Cen_Desp', S, 'C8:C10')
    ctx.name('Cen_Rec', S, 'D8:D10')
    sh.put('E8', '← digite aqui. Realista = sem ajuste. Valores padrão: Otimista −10% despesas/+5% receitas; Pessimista +10%/−10%.', 'note')
    # ---- base
    b = 12
    sh.put(f'B{b}', '2. Projeção base (cenário Realista)', 'sec')
    sh.put(f'B{b + 1}', 'Item', 'hdr')
    sh.put(f'C{b + 1}', 'Total 12 meses', 'hdr')
    sh.put(f'D{b + 1}', 'Total 24 meses', 'hdr')
    for k in range(NP):
        L = CL(5 + k)
        sh.put(f'{L}{b + 1}', f'={ml_label(f"{L}{b + 2}")}', 'hdr')
        sh.put(f'{L}{b + 2}', f'=Cfg_Idx+{k + 1}', 'calc', INT, size=9, color=COR['mudo'])
    sh.put(f'B{b + 2}', 'Índice do mês (apoio)', 'note')
    ws.row_dimensions[b + 2].hidden = True
    rows = [
        ('Receitas recorrentes', lambda L, k: f'={REC("Receita", f"{L}${b + 2}")}', BRL),
        ('Receitas já lançadas/agendadas (LANÇAMENTOS)', lambda L, k: f'={SM("Receita", f"{L}${b + 2}")}', BRL),
        ('RECEITAS PROJETADAS', lambda L, k: f'=IF({L}{b + 3}+{L}{b + 4}=0,K_RendaBase,{L}{b + 3}+{L}{b + 4})', BRL),
        ('Despesas fixas (recorrentes; ou média histórica se não houver)',
         lambda L, k: f'=IF({REC("Despesa", f"{L}${b + 2}")}=0,K_FixaMed,{REC("Despesa", f"{L}${b + 2}")})*(1+Cfg_Infl)^({k + 1}/12)', BRL),
        ('Parcelas futuras e lançamentos agendados', lambda L, k: f'={SM("Despesa", f"{L}${b + 2}")}', BRL),
        ('Despesas variáveis (média histórica corrigida pela inflação)', lambda L, k: f'=K_VarMed*(1+Cfg_Infl)^({k + 1}/12)', BRL),
        ('DESPESAS PROJETADAS', lambda L, k: f'={L}{b + 6}+{L}{b + 7}+{L}{b + 8}', BRL),
        ('Aportes recorrentes (investimentos)', lambda L, k: f'={REC("Investimento", f"{L}${b + 2}")}', BRL),
        ('SALDO DO MÊS (receitas − despesas)', lambda L, k: f'={L}{b + 5}-{L}{b + 9}', BRLZ),
    ]
    for i, (lbl, fn, fm) in enumerate(rows):
        r = b + 3 + i
        bold = lbl.isupper() or lbl.startswith(('RECEITAS', 'DESPESAS', 'SALDO'))
        sh.put(f'B{r}', lbl, 'calc', bold=bold)
        for k in range(NP):
            L = CL(5 + k)
            sh.put(f'{L}{r}', fn(L, k), 'calc', fm, bold=bold)
        sh.put(f'C{r}', f'=SUM(E{r}:P{r})', 'calc', fm, bold=True)
        sh.put(f'D{r}', f'=SUM(E{r}:AB{r})', 'calc', fm, bold=True)
    r_rec, r_desp, r_apo, r_sal = b + 5, b + 9, b + 10, b + 11
    ctx.name('Proj_Idx', S, f'E{b + 2}:AB{b + 2}')
    sh.cf_formula(f'E{r_sal}:AB{r_sal}', f'E{r_sal}<0', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'E{r_sal}:AB{r_sal}', f'E{r_sal}>=0', None, COR['verde'], True)
    ctx.name('Proj_Rec', S, f'E{r_rec}:AB{r_rec}')
    ctx.name('Proj_Desp', S, f'E{r_desp}:AB{r_desp}')
    ctx.name('Proj_Apo', S, f'E{r_apo}:AB{r_apo}')
    # ---- cenários
    c0 = r_sal + 3
    sh.put(f'B{c0}', '3. Os três cenários lado a lado', 'sec')
    sh.put(f'B{c0 + 1}', 'Item', 'hdr')
    sh.put(f'C{c0 + 1}', 'Total 12 meses', 'hdr')
    sh.put(f'D{c0 + 1}', 'Total 24 meses', 'hdr')
    for k in range(NP):
        L = CL(5 + k)
        sh.put(f'{L}{c0 + 1}', f'={L}{b + 1}', 'hdr')
    cur = c0 + 2
    ctx.sc_rows = {}
    for si, (n, _, _) in enumerate(cen):
        sh.put(f'B{cur}', n.upper(), 'lblb', color=[COR['azul'], COR['verde'], COR['vermelho']][si])
        names = ['Receitas', 'Despesas', 'Saldo do mês', 'Saldo acumulado', 'Saldo após aportes (acumulado)']
        for i, nm in enumerate(names):
            r = cur + 1 + i
            sh.put(f'B{r}', f'  {nm}', 'calc', bold=(i in (2, 3)))
            for k in range(NP):
                L = CL(5 + k)
                if i == 0:
                    f = f'={L}${r_rec}*(1+INDEX(Cen_Rec,{si + 1}))'
                elif i == 1:
                    f = f'={L}${r_desp}*(1+INDEX(Cen_Desp,{si + 1}))'
                elif i == 2:
                    f = f'={L}{cur + 1}-{L}{cur + 2}'
                elif i == 3:
                    f = f'=SUM($E{cur + 3}:{L}{cur + 3})'
                else:
                    f = f'=SUM($E{cur + 3}:{L}{cur + 3})-SUM($E${r_apo}:{L}${r_apo})'
                sh.put(f'{L}{r}', f, 'calc', BRLZ, bold=(i in (2, 3)))
            if i in (0, 1, 2):
                sh.put(f'C{r}', f'=SUM(E{r}:P{r})', 'calc', BRLZ, bold=True)
                sh.put(f'D{r}', f'=SUM(E{r}:AB{r})', 'calc', BRLZ, bold=True)
            else:
                sh.put(f'C{r}', f'=P{r}', 'calc', BRLZ, bold=True)
                sh.put(f'D{r}', f'=AB{r}', 'calc', BRLZ, bold=True)
        ctx.sc_rows[si] = cur
        ctx.name(f'Sc{si + 1}_Acum', S, f'E{cur + 4}:P{cur + 4}')
        ctx.name(f'Sc{si + 1}_Saldo12', S, f'C{cur + 3}')
        ctx.name(f'Sc{si + 1}_Saldo24', S, f'D{cur + 3}')
        ctx.name(f'Sc{si + 1}_Rec12', S, f'C{cur + 1}')
        ctx.name(f'Sc{si + 1}_Desp12', S, f'C{cur + 2}')
        sh.cf_formula(f'E{cur + 3}:AB{cur + 5}', f'E{cur + 3}<0', None, COR['vermelho'], True)
        cur += 7
    ctx.name('Proj_Labels', S, f'E{b + 1}:P{b + 1}')
    # ---- simulador E se?
    s0 = cur + 1
    sh.put(f'B{s0}', '4. Simulador "E se?"  ⬇ digite nos campos azuis', 'sec')
    campos = [('E se eu cortar (R$/mês) ...', BRLZ, 'Es_Corte'), ('... da categoria', None, 'Es_Cat'),
              ('E se eu aumentar o aporte em (R$/mês)', BRLZ, 'Es_Aporte'), ('E se minha renda cair (%)', '0.0%', 'Es_Renda')]
    for i, (lb, fm, n) in enumerate(campos):
        r = s0 + 1 + i
        sh.put(f'B{r}', lb, 'lbl')
        sh.put(f'C{r}', None, 'in', fm)
        ctx.name(n, S, f'C{r}')
    sh.dv(f'C{s0 + 2}', '=Lista_CatD')
    ctx.name('Exemplo_Simulador', S, f'C{s0 + 1}:C{s0 + 4}')
    ctx.clear_ranges.append((S, f'C{s0 + 1}:C{s0 + 4}', 'Campos do simulador'))
    out_r = s0 + 6
    sh.put(f'B{out_r}', 'Resultado (12 meses)', 'hdr')
    sh.put(f'C{out_r}', 'Valor', 'hdr')
    res = [
        ('Saldo anual projetado (cenário Realista)', '=Sc1_Saldo12', BRLZ),
        ('Corte considerado (limitado ao gasto médio da categoria) × 12', '=12*MIN(N(Es_Corte),IFERROR(INDEX(Cat_Med3,MATCH(Es_Cat,CatD_Nome,0)),N(Es_Corte)))', BRLZ),
        ('Efeito de aumentar o aporte (reduz o saldo livre) × 12', '=-12*N(Es_Aporte)', BRLZ),
        ('Efeito da queda de renda × 12', '=-N(Es_Renda)*Sc1_Rec12', BRLZ),
        ('SALDO ANUAL SIMULADO', f'=SUM(C{out_r + 1}:C{out_r + 4})', BRLZ),
        ('Diferença no saldo anual', f'=C{out_r + 5}-C{out_r + 1}', BRLZ),
        ('Patrimônio extra investido em 12 meses (aporte extra + rendimento CDI)',
         '=IF(N(Es_Aporte)=0,0,N(Es_Aporte)*(((1+Cfg_CDI)^(1/12))^12-1)/((1+Cfg_CDI)^(1/12)-1))', BRLZ),
        ('Valor extra mensal destinado às metas (corte + aporte − queda de renda)',
         f'=C{out_r + 2}/12-C{out_r + 3}/12*0+(-C{out_r + 3})/12*-1+C{out_r + 4}/12', BRLZ),
    ]
    # linha 8 simplificada abaixo
    for i, (lb, fx, fm) in enumerate(res):
        r = out_r + 1 + i
        sh.put(f'B{r}', lb, 'calc', bold=(i in (4, 5)))
        sh.put(f'C{r}', fx, 'calc', fm, bold=(i in (4, 5)))
    # extra mensal para metas = corte/12 + aporte − renda caída/mês
    ws[f'C{out_r + 8}'] = f'=C{out_r + 2}/12+N(Es_Aporte)-N(Es_Renda)*K_RendaBase'
    ctx.name('Es_Extra', S, f'C{out_r + 8}')
    sh.cf_formula(f'C{out_r + 5}:C{out_r + 6}', f'C{out_r + 5}<0', COR['vermelho_l'], COR['vermelho'], True)
    sh.cf_formula(f'C{out_r + 5}:C{out_r + 6}', f'C{out_r + 5}>=0', COR['verde_l'], COR['verde'], True)
    # efeito nas metas
    g0 = out_r + 11
    sh.put(f'B{g0 - 1}', 'Efeito na data das metas (o valor extra mensal é dividido igualmente entre as metas em andamento)', 'sec')
    for i, h in enumerate(['Meta', 'Ritmo atual (R$/mês)', 'Ritmo simulado (R$/mês)', 'Conclusão no ritmo atual', 'Conclusão simulada',
                           'Diferença (meses)']):
        sh.put(f'{CL(2 + i)}{g0}', h, 'hdr')
    ws.row_dimensions[g0].height = 32
    nat = 'MAX(1,COUNTIF(Meta_Falta,">0"))'
    for i in range(N_METAS):
        r = g0 + 1 + i
        m = M0 + i
        sh.put(f'B{r}', f"=INDEX(Meta_Nome,{i + 1})&\"\"", 'calc')
        sh.put(f'C{r}', f'=IF(B{r}="",0,INDEX(Meta_Ritmo,{i + 1}))', 'calc', BRL)
        sh.put(f'D{r}', f'=IF(OR(B{r}="",INDEX(Meta_Falta,{i + 1})<=0),0,MAX(0,C{r}+Es_Extra/{nat}))', 'calc', BRL)
        sh.put(f'E{r}', f'=IF(B{r}="","",IF(INDEX(Meta_Falta,{i + 1})<=0,"Concluída",IF(C{r}<=0,"sem ritmo",EDATE(Cfg_Hoje,ROUNDUP(INDEX(Meta_Falta,{i + 1})/C{r},0)))))', 'calc', MESANO, h='center')
        sh.put(f'F{r}', f'=IF(B{r}="","",IF(INDEX(Meta_Falta,{i + 1})<=0,"Concluída",IF(D{r}<=0,"sem ritmo",EDATE(Cfg_Hoje,ROUNDUP(INDEX(Meta_Falta,{i + 1})/D{r},0)))))', 'calc', MESANO, h='center')
        sh.put(f'G{r}', f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),(YEAR(E{r})-YEAR(F{r}))*12+MONTH(E{r})-MONTH(F{r}),"")', 'calc', '+0;-0;0', h='center')
        sh.cf_formula(f'B{r}:G{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    sh.cf_formula(f'G{g0 + 1}:G{g0 + N_METAS}', f'AND(ISNUMBER(G{g0 + 1}),G{g0 + 1}>0)', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'G{g0 + 1}:G{g0 + N_METAS}', f'AND(ISNUMBER(G{g0 + 1}),G{g0 + 1}<0)', COR['vermelho_l'], COR['vermelho'], True)
    sh.put(f'B{g0 + N_METAS + 1}', 'Diferença positiva = a meta chega quantos meses ANTES; negativa = depois.', 'note')
    sh.finish(freeze='C9')
    return sh


# ======================================================================
def build_calendario(ctx):
    sh = Sh(ctx, 'CALENDÁRIO DE CONTAS')
    ws = sh.ws
    S = 'CALENDÁRIO DE CONTAS'
    sh.header('CALENDÁRIO DE CONTAS', 'Vencimentos do mês selecionado no DASHBOARD (contas, faturas e parcelas). Amarelo = vence em até '
              '7 dias · Vermelho = atrasado · Verde = pago. Nada a digitar: tudo vem de LANÇAMENTOS, CARTÕES e RECORRENTES.', ncols=9)
    sh.widths({'B': 15, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 20, 'I': 12, 'J': 14})
    sh.put('B5', 'Mês:', 'lblb')
    sh.put('C5', '=Sel_Label', 'calc', bold=True)
    # ---- resumo próximos vencimentos
    sh.put('B7', '1. Total a pagar nos próximos dias (a partir de hoje)', 'sec')
    for i, h in enumerate(['Período', 'Total a pagar', 'Atrasado (já vencido)']):
        sh.put(f'{CL(2 + i)}8', h, 'hdr')
    # tabelas auxiliares (mais abaixo): faturas próximas e recorrentes próximas
    # layout: grid em 15-26, lista em 30+, aux em colunas L..R
    F0 = 12   # linha inicial da tabela aux de faturas (colunas L..O)
    ws.column_dimensions['K'].width = 3
    for L_ in 'KLMNOPQR':
        ws.column_dimensions[L_].hidden = True
    for L, w in zip('LMNOPQR', [24, 14, 13, 14, 14, 14, 12]):
        ws.column_dimensions[L].width = w
    sh.put('L7', 'Apoio: faturas dos próximos meses', 'lblb')
    for i, h in enumerate(['Cartão', 'Vencimento', 'Valor da fatura']):
        sh.put(f'{CL(12 + i)}8', h, 'hdr')
    rr = 9
    for mo in range(3):
        for k in range(1, 9):
            r = rr + (mo * 8) + k - 1
            # fatura cujo vencimento cai no mês (mês de hoje + mo)
            midx = f'(YEAR(Cfg_Hoje)*12+MONTH(Cfg_Hoje)+{mo})'
            f = f'({midx}-IF(INDEX(Cartoes_Venc,{k})>INDEX(Cartoes_Fech,{k}),0,1))'
            sh.put(f'L{r}', f'=IF(INDEX(Cartoes_Nome,{k})="","",INDEX(Cartoes_Nome,{k}))', 'calc')
            sh.put(f'M{r}', f'=IF(L{r}="","",DATE(INT(({midx}-1)/12),MOD({midx}-1,12)+1,INDEX(Cartoes_Venc,{k})))', 'calc', DATA)
            sh.put(f'N{r}', f'=IF(L{r}="",0,IF(SUMIFS(L_VParc,L_Tipo,"Transferência",L_Status,"Pago",L_Desc,"*"&L{r}&"*",L_IIni,{midx})<0,0,{fat(f"L{r}", f)}))', 'calc', BRL)
    f_last = rr + 24 - 1
    ctx.name('CalF_Data', S, f'M{rr}:M{f_last}')
    ctx.name('CalF_Valor', S, f'N{rr}:N{f_last}')
    # recorrentes: próxima ocorrência
    sh.put('P7', 'Apoio: recorrentes (próxima data)', 'lblb')
    for i, h in enumerate(['Próxima data', 'Valor']):
        sh.put(f'{CL(16 + i)}8', h, 'hdr')
    for k in range(N_REC):
        r = 9 + k
        d1 = f'DATE(YEAR(Cfg_Hoje),MONTH(Cfg_Hoje),MIN(INDEX(Rec_Dia,{k + 1}),28))'
        d2 = f'DATE(YEAR(Cfg_Hoje),MONTH(Cfg_Hoje)+1,MIN(INDEX(Rec_Dia,{k + 1}),28))'
        dd = f'IF({d1}>=Cfg_Hoje,{d1},{d2})'
        idx = f'(YEAR({dd})*12+MONTH({dd}))'
        sh.put(f'P{r}', f'=IF(OR(INDEX(Rec_Tipo,{k + 1})<>"Despesa",INDEX(Rec_Ativo,{k + 1})<>"Sim",INDEX(Rec_IIni,{k + 1})=0,'
               f'N(INDEX(Rec_Dia,{k + 1}))=0),"",IF(AND({idx}>=INDEX(Rec_IIni,{k + 1}),{idx}<=INDEX(Rec_IFim,{k + 1}),'
               f'OR(INDEX(Rec_Freq,{k + 1})<>"Anual",MONTH({dd})=INDEX(Rec_MesRef,{k + 1}))),{dd},""))', 'calc', DATA)
        sh.put(f'Q{r}', f'=IF(P{r}="",0,N(INDEX(Rec_Valor,{k + 1})))', 'calc', BRL)
    ctx.name('CalR_Data', S, f'P9:P{8 + N_REC}')
    ctx.name('CalR_Valor', S, f'Q9:Q{8 + N_REC}')
    janelas = [('Próximos 7 dias', 7), ('Próximos 15 dias', 15), ('Próximos 30 dias', 30)]
    for i, (lb, n) in enumerate(janelas):
        r = 9 + i
        lim = f'Cfg_Hoje+{n}'
        sh.put(f'B{r}', lb, 'calc', bold=True)
        sh.put(f'C{r}', (f'=SUMIFS(L_VParc,L_Tipo,"Despesa",L_Status,"<>Pago",L_FIni,0,L_Data,">="&Cfg_Hoje,L_Data,"<="&({lim}))'
                         f'+SUMIFS(CalF_Valor,CalF_Data,">="&Cfg_Hoje,CalF_Data,"<="&({lim}))'
                         f'+SUMIFS(CalR_Valor,CalR_Data,">="&Cfg_Hoje,CalR_Data,"<="&({lim}))'), 'calc', BRL, bold=True)
    sh.put('D9', '=SUMIFS(L_VParc,L_Tipo,"Despesa",L_Status,"<>Pago",L_FIni,0,L_Data,"<"&Cfg_Hoje)', 'calc', BRL, bold=True)
    sh.cf_formula('D9', 'D9>0', COR['vermelho_l'], COR['vermelho'], True)
    ctx.name('Cal_7', S, 'C9')
    ctx.name('Cal_15', S, 'C10')
    ctx.name('Cal_30', S, 'C11')
    ctx.name('Cal_Atraso', S, 'D9')
    sh.put('B12', 'Inclui: lançamentos Pendente/Agendado (exceto compras no cartão), faturas de cartão a vencer e recorrentes previstos.', 'note')
    # ---- grade do mês
    g = 14
    sh.put(f'B{g}', '2. Calendário do mês (valor a pagar por dia)', 'sec')
    for i, d in enumerate(['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']):
        sh.put(f'{CL(2 + i)}{g + 1}', d, 'hdr')
    wd = 'WEEKDAY(Sel_Ini)'
    # tabela de faturas do mês selecionado (aux, no final)
    L0 = g + 1 + 18 + 4        # início da lista
    fat0 = L0 + 44 + 3         # faturas do mês
    for w in range(6):
        r_d = g + 2 + w * 3
        r_v, r_p = r_d + 1, r_d + 2
        ws.row_dimensions[r_p].hidden = True
        ws.row_dimensions[r_v].height = 20
        for c in range(1, 8):
            col = CL(1 + c)
            dnum = f'({w * 7 + c}-({wd}-1))'
            sh.put(f'{col}{r_d}', f'=IF(AND({dnum}>=1,{dnum}<=Sel_Dias),{dnum},"")', 'calc', INT, bold=True, h='left', size=10)
            base = ('L_Tipo,"Despesa",L_FIni,0,L_IIni,"<="&Sel_Idx,L_IFim,">="&Sel_Idx,L_Dia,' + f'{col}{r_d}')
            rec = f'IF(Sel_Idx>Cfg_Idx,{REC("Despesa", "Sel_Idx", f",Rec_Dia,{col}{r_d}")},0)'
            tot = f'SUMIFS(L_VParc,{base})+{rec}+SUMIFS($D${fat0 + 1}:$D${fat0 + 8},$C${fat0 + 1}:$C${fat0 + 8},{col}{r_d})'
            pend = (f'SUMIFS(L_VParc,{base},L_Status,"<>Pago")+{rec}+'
                    f'SUMIFS($F${fat0 + 1}:$F${fat0 + 8},$C${fat0 + 1}:$C${fat0 + 8},{col}{r_d})')
            sh.put(f'{col}{r_v}', f'=IF({col}{r_d}="","",{tot})', 'calc', BRL0, h='right', size=10)
            sh.put(f'{col}{r_p}', f'=IF({col}{r_d}="",0,{pend})', 'calc', BRL0, size=8)
            dt = f'DATE(Sel_Ano,Sel_Mes,{col}${r_d})'
            rng_ = f'{col}{r_d}:{col}{r_v}'
            sh.cf_formula(rng_, f'{col}${r_d}=""', COR['cinza'], COR['cinza'])
            sh.cf_formula(rng_, f'AND(N({col}${r_p})>0,{dt}<Cfg_Hoje)', COR['vermelho_l'], COR['vermelho'])
            sh.cf_formula(rng_, f'AND(N({col}${r_p})>0,{dt}>=Cfg_Hoje,{dt}-Cfg_Hoje<=7)', COR['amarelo_l'], '9A5B00')
            sh.cf_formula(rng_, f'AND(N({col}${r_p})>0,{dt}>Cfg_Hoje+7)', COR['azul_l'], COR['azul'])
            sh.cf_formula(rng_, f'AND(N({col}${r_v})>0,N({col}${r_p})=0)', COR['verde_l'], COR['verde'])
    sh.put(f'B{g + 20}', 'Cores: verde = tudo pago · azul = a pagar (mais de 7 dias) · amarelo = vence em até 7 dias · vermelho = atrasado. '
           'O valor do dia soma contas, parcelas, recorrentes futuros e faturas de cartão.', 'note')
    # ---- lista do mês
    sh.put(f'B{L0 - 1}', '3. Contas e parcelas do mês (em ordem de dia)', 'sec')
    for i, h in enumerate(['Vencimento', 'Descrição', 'Categoria', 'Conta ou Cartão', 'Valor', 'Status', 'Situação', 'Linha (apoio)']):
        sh.put(f'{CL(2 + i)}{L0}', h, 'hdr')
    for k in range(1, 41):
        r = L0 + k
        ix = f'$I{r}'
        sh.put(f'I{r}', f'=IFERROR(MATCH({k},L_CalOrdem,0),0)', 'calc', INT)
        sh.put(f'B{r}', f'=IF({ix}=0,"",DATE(Sel_Ano,Sel_Mes,INDEX(L_Dia,{ix})))', 'calc', DATA, h='center')
        sh.put(f'C{r}', f'=IF({ix}=0,"",INDEX(L_Desc,{ix}))', 'calc')
        sh.put(f'D{r}', f'=IF({ix}=0,"",INDEX(L_Cat,{ix})&"")', 'calc')
        sh.put(f'E{r}', f'=IF({ix}=0,"",INDEX(L_Conta,{ix})&"")', 'calc')
        sh.put(f'F{r}', f'=IF({ix}=0,"",INDEX(L_VParc,{ix}))', 'calc', BRLZ)
        sh.put(f'G{r}', f'=IF({ix}=0,"",INDEX(L_Status,{ix})&"")', 'calc', h='center')
        sh.put(f'H{r}', (f'=IF({ix}=0,"",IF(INDEX(L_FIni,{ix})>0,"Na fatura do cartão",IF(G{r}="Pago","✔ Pago",IF(B{r}<Cfg_Hoje,"⚠ Atrasado",IF(B{r}-Cfg_Hoje<=7,'
                         f'"Vence em "&(B{r}-Cfg_Hoje)&" dia(s)","No prazo")))))'), 'calc', h='center')
        sh.cf_formula(f'B{r}:H{r}', f'$H{r}="⚠ Atrasado"', COR['vermelho_l'], COR['vermelho'])
        sh.cf_formula(f'B{r}:H{r}', f'LEFT($H{r},5)="Vence"', COR['amarelo_l'], '9A5B00')
        sh.cf_formula(f'B{r}:H{r}', f'$H{r}="✔ Pago"', COR['verde_l'], COR['verde'])
    ws.column_dimensions['I'].width = 12
    # compactar: C (descrição) larga
    ws.column_dimensions['C'].width = 26
    sh.put(f'B{L0 + 41}', 'Mostra até 40 itens do mês (lançamentos de despesa e parcelas). Compras no cartão aparecem na fatura abaixo.', 'note')
    # ---- faturas do mês selecionado
    sh.put(f'B{fat0 - 1}', '4. Faturas de cartão com vencimento no mês', 'sec')
    for i, h in enumerate(['Cartão', 'Dia do vencimento', 'Valor da fatura', 'Situação', 'A pagar', 'Paga? (apoio)']):
        sh.put(f'{CL(2 + i)}{fat0}', h, 'hdr')
    for k in range(1, 9):
        r = fat0 + k
        f = f'(Sel_Idx-IF(INDEX(Cartoes_Venc,{k})>INDEX(Cartoes_Fech,{k}),0,1))'
        sh.put(f'B{r}', f'=IF(INDEX(Cartoes_Nome,{k})="","",INDEX(Cartoes_Nome,{k}))', 'calc')
        sh.put(f'C{r}', f'=IF(B{r}="","",INDEX(Cartoes_Venc,{k}))', 'calc', INT, h='center')
        sh.put(f'D{r}', f'=IF(B{r}="",0,{fat(f"B{r}", f)})', 'calc', BRL)
        sh.put(f'G{r}', f'=IF(OR(B{r}="",D{r}=0),0,IF(SUMIFS(L_VParc,L_Tipo,"Transferência",L_Status,"Pago",L_Desc,"*"&B{r}&"*",L_IIni,Sel_Idx)<0,1,0))', 'calc', INT, size=8, color=COR['mudo'])
        sh.put(f'F{r}', f'=IF(G{r}=1,0,D{r})', 'calc', BRL)
        sh.put(f'E{r}', (f'=IF(OR(B{r}="",D{r}=0),"",IF(G{r}=1,"✔ Paga",IF(DATE(Sel_Ano,Sel_Mes,C{r})<Cfg_Hoje,"⚠ Vencida",'
                         f'IF(DATE(Sel_Ano,Sel_Mes,C{r})-Cfg_Hoje<=7,"Vence em "&(DATE(Sel_Ano,Sel_Mes,C{r})-Cfg_Hoje)&" dia(s)","No prazo"))))'), 'calc', h='center')
        sh.cf_formula(f'B{r}:F{r}', f'$E{r}="⚠ Vencida"', COR['vermelho_l'], COR['vermelho'])
        sh.cf_formula(f'B{r}:F{r}', f'LEFT($E{r},5)="Vence"', COR['amarelo_l'], '9A5B00')
        sh.cf_formula(f'B{r}:F{r}', f'$E{r}="✔ Paga"', COR['verde_l'], COR['verde'])
    sh.put(f'B{fat0 + 9}', 'A fatura é considerada PAGA quando existe uma Transferência negativa no mês do vencimento com o nome do cartão na descrição '
           '(ex.: "Pagamento fatura Cartão Roxinho").', 'note')
    sh.finish(freeze='A7')
    return sh


# ======================================================================
def build_analises(ctx):
    sh = Sh(ctx, 'HISTÓRICO E ANÁLISES')
    ws = sh.ws
    S = 'HISTÓRICO E ANÁLISES'
    sh.header('HISTÓRICO E ANÁLISES', 'Onde o dinheiro foi: maiores gastos do mês, categorias, evolução de 12 meses, formas de pagamento, '
              'dias da semana, tags e assinaturas. Tudo automático, para o mês/ano escolhido no DASHBOARD.', ncols=14)
    sh.widths({'B': 32, 'C': 14, 'D': 14, 'E': 14, 'F': 14, **{CL(c): 12.5 for c in range(7, 15)}})
    sh.put('B5', 'Período:', 'lblb')
    sh.put('C5', '=Sel_Label', 'calc', bold=True)
    # 1. top 10
    sh.put('B7', '1. Os 10 maiores gastos do mês', 'sec')
    for i, h in enumerate(['Descrição', 'Categoria', 'Data da compra', 'Valor do mês', '% das despesas']):
        sh.put(f'{CL(2 + i)}8', h, 'hdr')
    for k in range(1, 11):
        r = 8 + k
        v = f'LARGE(L_DespSel,{k})'
        ix = f'IFERROR(MATCH({v},L_DespSel,0),0)'
        sh.put(f'G{r}', f'=IF({v}>=1,{ix},0)', 'calc', INT, size=9, color=COR['mudo'])
        sh.put(f'B{r}', f'=IF($G{r}=0,"",INDEX(L_Desc,$G{r}))', 'calc')
        sh.put(f'C{r}', f'=IF($G{r}=0,"",INDEX(L_Cat,$G{r}))', 'calc')
        sh.put(f'D{r}', f'=IF($G{r}=0,"",INDEX(L_Data,$G{r}))', 'calc', DATA, h='center')
        sh.put(f'E{r}', f'=IF($G{r}=0,"",INDEX(L_VParc,$G{r}))', 'calc', BRL)
        sh.put(f'F{r}', f'=IF($G{r}=0,"",IFERROR(E{r}/K_Desp,0))', 'calc', PCT)
    sh.put('G8', 'linha (apoio)', 'hdr', size=9)
    sh.cf(f'E9:E18', DataBarRule(start_type='num', start_value=0, end_type='max', color=COR['vermelho']))
    # 2. maiores categorias
    sh.put('B21', '2. Maiores categorias de despesa do mês', 'sec')
    for i, h in enumerate(['Categoria', 'Valor', '% das despesas', 'Planejado', '% do planejado']):
        sh.put(f'{CL(2 + i)}22', h, 'hdr')
    for p in range(1, 9):
        r = 22 + p
        sh.put(f'B{r}', f'=INDEX(Rk_Nome,{p})', 'calc')
        sh.put(f'C{r}', f'=INDEX(Rk_Val,{p})', 'calc', BRL)
        sh.put(f'D{r}', f'=IFERROR(C{r}/K_Desp,0)', 'calc', PCT)
        sh.put(f'E{r}', f'=IFERROR(INDEX(Orc_Plan,MATCH(B{r},CatD_Nome,0)),0)', 'calc', BRL)
        sh.put(f'F{r}', f'=IF(E{r}=0,0,C{r}/E{r})', 'calc', PCT)
    sh.cf(f'C23:C30', DataBarRule(start_type='num', start_value=0, end_type='max', color=COR['azul']))
    # 3. evolução 12 meses
    e = 33
    sh.put(f'B{e}', '3. Evolução de cada categoria nos últimos 12 meses (termina no mês selecionado)', 'sec')
    sh.put(f'B{e + 1}', 'Categoria', 'hdr')
    for j in range(12):
        L = CL(3 + j)
        sh.put(f'{L}{e + 1}', f'={ml_label(f"(Sel_Idx-{11 - j})")}', 'hdr')
    sh.put(f'O{e + 1}', 'Média', 'hdr')
    sh.put(f'P{e + 1}', 'Tendência', 'hdr')
    for k in range(1, N_CATD + 1):
        r = e + 1 + k
        sh.put(f'B{r}', f'=IF(INDEX(CatD_Nome,{k})="","",INDEX(CatD_Nome,{k}))', 'calc')
        for j in range(12):
            L = CL(3 + j)
            sh.put(f'{L}{r}', f'=IF($B{r}="",0,{SM("Despesa", f"Sel_Idx-{11 - j}", f",L_Cat,$B{r}")})', 'calc', BRL0)
        sh.put(f'O{r}', f'=AVERAGE(C{r}:N{r})', 'calc', BRL0)
        sh.put(f'P{r}', f'=IF($B{r}="","",IF(O{r}=0,"—",IF(N{r}>O{r}*1.1,"▲ subindo",IF(N{r}<O{r}*0.9,"▼ caindo","▬ estável"))))', 'calc', h='center')
        sh.cf(f'C{r}:N{r}', ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='F2A93B'))
        sh.cf_formula(f'B{r}:P{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    sh.cf_formula(f'P{e + 2}:P{e + 1 + N_CATD}', f'LEFT(P{e + 2},1)="▲"', None, COR['vermelho'], True)
    sh.cf_formula(f'P{e + 2}:P{e + 1 + N_CATD}', f'LEFT(P{e + 2},1)="▼"', None, COR['verde'], True)
    sh.widths({'O': 13, 'P': 14})
    # 4. forma de pagamento
    f0 = e + 1 + N_CATD + 3
    sh.put(f'B{f0}', '4. Gastos por forma de pagamento (mês selecionado)', 'sec')
    for i, h in enumerate(['Forma de pagamento', 'Valor', '% do total']):
        sh.put(f'{CL(2 + i)}{f0 + 1}', h, 'hdr')
    for k in range(1, N_FORMA + 1):
        r = f0 + 1 + k
        sh.put(f'B{r}', f'=IF(INDEX(Forma_Nome,{k})="","",INDEX(Forma_Nome,{k}))', 'calc')
        sh.put(f'C{r}', f'=IF(B{r}="",0,{SM("Despesa", "Sel_Idx", f",L_Forma,$B{r}")})', 'calc', BRL)
        sh.put(f'D{r}', f'=IFERROR(C{r}/K_Desp,0)', 'calc', PCT)
        sh.cf_formula(f'B{r}:D{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    r = f0 + 2 + N_FORMA
    sh.put(f'B{r}', 'Sem forma informada', 'calc')
    sh.put(f'C{r}', f'=K_Desp-SUM(C{f0 + 2}:C{r - 1})', 'calc', BRL)
    sh.put(f'D{r}', f'=IFERROR(C{r}/K_Desp,0)', 'calc', PCT)
    sh.cf(f'C{f0 + 2}:C{r}', DataBarRule(start_type='num', start_value=0, end_type='max', color=COR['azul']))
    # 5. dia da semana (ao lado)
    sh.put(f'F{f0}', '5. Gastos por dia da semana (mês)', 'sec')
    for i, h in enumerate(['Dia', 'Valor', '% do total']):
        sh.put(f'{CL(6 + i)}{f0 + 1}', h, 'hdr')
    for d in range(1, 8):
        r = f0 + 1 + d
        sh.put(f'F{r}', f'=INDEX(Lista_DiaSem,{d})', 'calc')
        sh.put(f'G{r}', f'={SM("Despesa", "Sel_Idx", f",L_Sem,{d}")}', 'calc', BRL)
        sh.put(f'H{r}', f'=IFERROR(G{r}/K_Desp,0)', 'calc', PCT)
    sh.cf(f'G{f0 + 2}:G{f0 + 8}', DataBarRule(start_type='num', start_value=0, end_type='max', color=COR['amarelo']))
    # 6. tags
    t0 = f0 + 2 + N_FORMA + 3
    sh.put(f'B{t0}', '6. Gastos por tag (ano selecionado)', 'sec')
    for i, h in enumerate(['Tag', 'Valor no ano', 'Nº de lançamentos']):
        sh.put(f'{CL(2 + i)}{t0 + 1}', h, 'hdr')
    for k in range(1, N_TAG + 1):
        r = t0 + 1 + k
        sh.put(f'B{r}', f'=IF(INDEX(Tag_Nome,{k})="","",INDEX(Tag_Nome,{k}))', 'calc')
        sh.put(f'C{r}', f'=IF(B{r}="",0,SUMIFS(L_VParc,L_Tipo,"Despesa",L_Tag,B{r},L_Ano,Sel_Ano))', 'calc', BRL)
        sh.put(f'D{r}', f'=IF(B{r}="",0,COUNTIFS(L_Tag,B{r},L_Ano,Sel_Ano))', 'calc', INT)
        sh.cf_formula(f'B{r}:D{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    # 7. assinaturas
    a0 = t0 + N_TAG + 4
    sh.put(f'B{a0}', '7. Assinaturas ativas (RECORRENTES com categoria "Assinaturas")', 'sec')
    for i, h in enumerate(['Assinatura', 'Valor', 'Frequência', 'Custo anual']):
        sh.put(f'{CL(2 + i)}{a0 + 1}', h, 'hdr')
    for k in range(1, 16):
        r = a0 + 1 + k
        ix = f'IFERROR(MATCH({k},Rec_AssinOrd,0),0)'
        sh.put(f'G{r}', f'={ix}', 'calc', INT, size=9, color=COR['mudo'])
        sh.put(f'B{r}', f'=IF($G{r}=0,"",INDEX(Rec_Desc,$G{r}))', 'calc')
        sh.put(f'C{r}', f'=IF($G{r}=0,"",INDEX(Rec_Valor,$G{r}))', 'calc', BRL)
        sh.put(f'D{r}', f'=IF($G{r}=0,"",INDEX(Rec_Freq,$G{r})&"")', 'calc', h='center')
        sh.put(f'E{r}', f'=IF($G{r}=0,"",INDEX(Rec_Anual,$G{r}))', 'calc', BRL)
    sh.put(f'G{a0 + 1}', 'linha (apoio)', 'hdr', size=9)
    r = a0 + 17
    sh.put(f'B{r}', 'CUSTO ANUAL TOTAL DAS ASSINATURAS', 'calc', bold=True)
    sh.put(f'E{r}', f'=SUMIFS(Rec_Anual,Rec_Assin,1)', 'calc', BRL, bold=True)
    sh.put(f'B{r + 1}', 'Custo mensal equivalente', 'calc')
    sh.put(f'E{r + 1}', f'=E{r}/12', 'calc', BRL)
    ctx.name('Assin_Anual', S, f'E{r}')
    ctx.name('Assin_Mensal', S, f'E{r + 1}')
    sh.finish(freeze='C7')
    return sh
