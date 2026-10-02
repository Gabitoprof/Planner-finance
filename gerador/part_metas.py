"""METAS E OBJETIVOS (aba central)."""
from common import *
from part_base import N_METAS, N_CATD, CAT_META, META_COR, dyn

M0 = 9                  # primeira linha das metas
M1 = M0 + N_METAS - 1


def build_metas(ctx):
    sh = Sh(ctx, 'METAS E OBJETIVOS')
    ws = sh.ws
    S = 'METAS E OBJETIVOS'
    sh.header('METAS E OBJETIVOS', 'Cadastre até 15 metas. O valor guardado soma automaticamente os aportes de LANÇAMENTOS '
              '(Tipo = Investimento, coluna "Meta vinculada"). Veja status, ritmo, prazo, planejador reverso e prioridades.',
              ncols=19)
    sh.widths({'B': 28, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 12, 'H': 12, 'I': 22, 'J': 14, 'K': 15, 'L': 15,
               'M': 11, 'N': 10, 'O': 15, 'P': 15, 'Q': 13, 'R': 13, 'S': 24})
    sh.put('B5', '⬇ digite nas colunas azuis (B a I)', 'note', color=COR['azul'], bold=True)
    sh.put('J5', 'calculado automaticamente ⬇', 'note')
    sh.put('B6', '1. Cadastro e acompanhamento das metas', 'sec')
    heads = ['Meta', 'Categoria', 'Prioridade', 'Valor alvo', 'Já guardado (inicial)', 'Data inicial', 'Prazo final',
             'Onde está guardado', 'Aportes lançados', 'Guardado total', 'Falta', '% concluído', 'Meses restantes',
             'Necessário por mês', 'Ritmo real (média 3 meses)', 'Conclusão estimada', 'Status', 'Progresso']
    for i, h in enumerate(heads):
        sh.put(f'{CL(2 + i)}8', h, 'hdr')
    ws.row_dimensions[8].height = 46
    for r in range(M0, M1 + 1):
        for L, fm in zip('BCDEFGHI', [None, None, None, BRLZ, BRLZ, DATA, DATA, None]):
            sh.put(f'{L}{r}', None, 'in', fm)
        f = {
            'J': f'=IF($B{r}="",0,SUMIFS(L_VParc,L_Meta,$B{r},L_Tipo,"Investimento",L_Status,"<>Agendado",L_Status,"<>Pendente",L_Data,"<="&Cfg_Hoje))',
            'K': f'=IF($B{r}="",0,N($F{r})+J{r})',
            'L': f'=IF($B{r}="",0,MAX(0,N($E{r})-K{r}))',
            'M': f'=IF(N($E{r})=0,0,MIN(1,K{r}/$E{r}))',
            'N': f'=IF(OR($B{r}="",NOT(ISNUMBER($H{r}))),0,MAX(0,(YEAR($H{r})-YEAR(Cfg_Hoje))*12+MONTH($H{r})-MONTH(Cfg_Hoje)))',
            'O': f'=IF(OR($B{r}="",L{r}=0),0,IF(N{r}>0,L{r}/N{r},L{r}))',
            'P': (f'=IF($B{r}="",0,SUMIFS(L_VParc,L_Meta,$B{r},L_Tipo,"Investimento",L_IIni,">="&(Cfg_Idx-3),'
                  f'L_IIni,"<="&(Cfg_Idx-1))/3)'),
            'Q': f'=IF(OR($B{r}="",N($E{r})=0),"",IF(L{r}=0,"Concluída",IF(P{r}<=0,"sem ritmo",EDATE(Cfg_Hoje,ROUNDUP(L{r}/P{r},0)))))',
            'R': (f'=IF(OR($B{r}="",N($E{r})=0),"",IF(L{r}=0,"Concluída",IF(AND(ISNUMBER($H{r}),Cfg_Hoje>$H{r}),"Atrasada",'
                  f'IF(NOT(ISNUMBER($H{r})),IF(P{r}>0,"No prazo","Atenção"),IF(P{r}>=O{r},"No prazo",'
                  f'IF(P{r}>=0.7*O{r},"Atenção","Atrasada"))))))'),
            'S': f'=IF($B{r}="","",REPT("█",ROUND(M{r}*20,0))&REPT("░",20-ROUND(M{r}*20,0)))',
        }
        fm = {'J': BRL, 'K': BRL, 'L': BRL, 'M': PCT, 'N': INT, 'O': BRL, 'P': BRL, 'Q': MESANO}
        for L, fx in f.items():
            sh.put(f'{L}{r}', fx, 'calc', fm.get(L), h=('center' if L in 'QRS' else None))
    nm = dict(B='Nome', C='Cat', D='Prio', E='Alvo', F='Ini0', G='DataIni', H='Prazo', I='Onde', J='Aportes', K='Guardado',
              L='Falta', M='Pct', N='MesesRest', O='Nec', P='Ritmo', Q='Estim', R='Status', S='Barra')
    for L, n in nm.items():
        ctx.name(f'Meta_{n}', S, f'{L}{M0}:{L}{M1}')
    ctx.name_formula('Lista_MetasNomes', dyn(S, 'B', M0, M1))
    ctx.name('Exemplo_Metas', S, f'B{M0}:I{M1}')
    ctx.clear_ranges.append((S, f'B{M0}:I{M1}', 'Metas de exemplo'))
    sh.dv(f'C{M0}:C{M1}', '=Lista_CatMeta')
    sh.dv(f'D{M0}:D{M1}', '=Lista_Prioridade')
    # cores por categoria / status
    for c in CAT_META:
        sh.cf_formula(f'B{M0}:B{M1}', f'AND($B{M0}<>"",$C{M0}="{c}")', META_COR[c], 'FFFFFF', True)
    sh.cf_formula(f'R{M0}:R{M1}', f'R{M0}="Concluída"', COR['verde'], 'FFFFFF', True)
    sh.cf_formula(f'R{M0}:R{M1}', f'R{M0}="No prazo"', COR['azul'], 'FFFFFF', True)
    sh.cf_formula(f'R{M0}:R{M1}', f'R{M0}="Atenção"', COR['amarelo'], '1F2933', True)
    sh.cf_formula(f'R{M0}:R{M1}', f'R{M0}="Atrasada"', COR['vermelho'], 'FFFFFF', True)
    for st, cr in (('Concluída', COR['verde']), ('No prazo', COR['azul']), ('Atenção', COR['amarelo']),
                   ('Atrasada', COR['vermelho'])):
        sh.cf_formula(f'S{M0}:S{M1}', f'$R{M0}="{st}"', None, cr, True)
    for r in range(M0, M1 + 1):
        pass
    # ---------------- textos dinâmicos
    t0 = M1 + 3
    sh.put(f'B{t0 - 1}', '2. O que cada meta está dizendo', 'sec')
    for i in range(N_METAS):
        r = M0 + i
        rr = t0 + i
        diff = f'((YEAR(Q{r})-YEAR(H{r}))*12+MONTH(Q{r})-MONTH(H{r}))'
        txt = (f'=IF($B{r}="","",IF(L{r}=0,"🎉 "&$B{r}&": meta concluída! Você guardou R$ "&FIXED(K{r},0)&".",'
               f'$B{r}&": faltam R$ "&FIXED(L{r},0)&". "&IF(P{r}<=0,"Sem aportes nos últimos 3 meses. ",'
               f'"Guardando R$ "&FIXED(P{r},0)&"/mês você chega em "&{my_label(f"Q{r}")}&'
               f'IF(NOT(ISNUMBER(H{r})),". ",IF({diff}>0,", "&{diff}&" mes(es) depois do prazo. ",IF({diff}<0,", "&-{diff}&" mes(es) antes do prazo. ",", no prazo. "))))&'
               f'"Para cumprir o prazo, guarde R$ "&FIXED(O{r},0)&"/mês."))')
        sh.put(f'B{rr}', txt, 'calc', wrap=True, v_align='center')
        sh.merge(f'B{rr}:L{rr}')
        sh.put(f'M{rr}', f'=IF($B{r}="","",R{r})', 'calc', h='center', bold=True)
        ws.row_dimensions[rr].height = 32
    ctx.name('Meta_Texto', S, f'B{t0}:B{t0 + N_METAS - 1}')
    ctx.name('Meta_TextoStatus', S, f'M{t0}:M{t0 + N_METAS - 1}')
    for st, cr in (('Concluída', COR['verde']), ('No prazo', COR['azul']), ('Atenção', COR['amarelo']),
                   ('Atrasada', COR['vermelho'])):
        sh.cf_formula(f'M{t0}:M{t0 + N_METAS - 1}', f'M{t0}="{st}"', cr, ('1F2933' if st == 'Atenção' else 'FFFFFF'), True)
    # ---------------- planejador reverso
    p = t0 + N_METAS + 1
    sh.put(f'B{p}', '3. Planejador reverso: "Quero ter R$ X até a data Y"', 'sec')
    inp = [('Quero ter (R$)', BRLZ, 'Pl_X'), ('Até a data', DATA, 'Pl_Data'), ('Já tenho para isso (R$)', BRLZ, 'Pl_PV'),
           ('Rendimento mensal esperado', '0.00%', 'Pl_Taxa')]
    for i, (lbl, fm, n) in enumerate(inp):
        r = p + 1 + i
        sh.put(f'B{r}', lbl, 'lbl')
        sh.put(f'C{r}', None, 'in', fm)
        ctx.name(n, S, f'C{r}')
    ws[f'C{p + 4}'] = '=(1+Cfg_CDI)^(1/12)-1'
    sh.put(f'D{p + 4}', '← sugestão: CDI de CONFIGURAÇÕES (pode digitar outra taxa)', 'note')
    ctx.name('Exemplo_Planejador', S, f'C{p + 1}:C{p + 3}')
    out = [
        ('Meses até a data', '=IF(ISNUMBER(Pl_Data),MAX(1,(YEAR(Pl_Data)-YEAR(Cfg_Hoje))*12+MONTH(Pl_Data)-MONTH(Cfg_Hoje)),0)', INT, 'Pl_N'),
        ('Guardar por mês SEM rendimento', '=IF(Pl_N=0,0,MAX(0,N(Pl_X)-N(Pl_PV))/Pl_N)', BRL, 'Pl_SemRend'),
        ('Guardar por mês COM rendimento', '=IF(Pl_N=0,0,MAX(0,-PMT(N(Pl_Taxa),Pl_N,-N(Pl_PV),N(Pl_X))))', BRL, 'Pl_ComRend'),
        ('Seu saldo livre mensal (renda − despesas médias)', '=K_SaldoLivre', BRL, 'Pl_Livre'),
        ('Cabe no saldo livre?', '=IF(Pl_N=0,"Informe valor e data",IF(Pl_ComRend<=Pl_Livre,"✔ Cabe: sobram R$ "&FIXED(Pl_Livre-Pl_ComRend,0)&" por mês","✘ Não cabe: faltam R$ "&FIXED(Pl_ComRend-Pl_Livre,0)&" por mês"))', None, 'Pl_Cabe'),
        ('Prazo mínimo viável com seu saldo livre (meses)', '=IF(OR(Pl_N=0,Pl_Livre<=0),"—",IFERROR(ROUNDUP(NPER(N(Pl_Taxa),-Pl_Livre,-N(Pl_PV),N(Pl_X)),0),"—"))', INT, 'Pl_Prazo'),
        ('…ou seja, data mínima viável', '=IF(ISNUMBER(Pl_Prazo),EDATE(Cfg_Hoje,Pl_Prazo),"—")', MESANO, 'Pl_DataMin'),
        ('O que cortar por mês para cumprir o prazo', '=IF(Pl_N=0,0,MAX(0,Pl_ComRend-Pl_Livre))', BRL, 'Pl_Cortar'),
    ]
    for i, (lbl, fx, fm, n) in enumerate(out):
        r = p + 6 + i
        sh.put(f'B{r}', lbl, 'lbl')
        sh.put(f'C{r}', fx, 'calc', fm, h='right', bold=(i in (1, 2, 4)))
        ctx.name(n, S, f'C{r}')
    sh.merge(f'C{p + 10}:G{p + 10}')
    sh.cf_formula(f'C{p + 10}', f'LEFT(C{p + 10},1)="✔"', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'C{p + 10}', f'LEFT(C{p + 10},1)="✘"', COR['vermelho_l'], COR['vermelho'], True)
    # ---------------- priorização
    q = p + 6 + len(out) + 2
    sh.put(f'B{q}', '4. Priorização: como distribuir o saldo livre mensal entre as metas', 'sec')
    sh.put(f'B{q + 1}', 'Saldo livre mensal disponível:', 'lblb')
    sh.put(f'D{q + 1}', '=K_SaldoLivre', 'calc', BRL, bold=True)
    for i, h in enumerate(['Meta', 'Prioridade', 'Prazo', 'Necessário/mês', 'Ordem', 'Aporte sugerido/mês', '% do necessário']):
        sh.put(f'{CL(2 + i)}{q + 2}', h, 'hdr')
    k0 = q + 3
    k1 = k0 + N_METAS - 1
    for i in range(N_METAS):
        r = k0 + i
        m = M0 + i
        sh.put(f'B{r}', f'=IF($B{m}="","",$B{m})', 'calc')
        sh.put(f'C{r}', f'=IF($B{m}="","",$D{m})', 'calc', h='center')
        sh.put(f'D{r}', f'=IF(OR($B{m}="",NOT(ISNUMBER($H{m}))),"",$H{m})', 'calc', DATA)
        sh.put(f'E{r}', f'=IF($B{m}="",0,O{m})', 'calc', BRL)
        sh.put(f'H{r}', (f'=IF(OR($B{m}="",O{m}=0),"",IF($D{m}="Alta",1,IF($D{m}="Média",2,3))*1000000+IF(ISNUMBER($H{m}),$H{m},999999)'
                         f'+ROW()/1000)'), 'calc', NUM1)
        sh.put(f'F{r}', f'=IF(H{r}="","",COUNTIF($H${k0}:$H${k1},"<"&H{r})+1)', 'calc', INT, h='center')
        sh.put(f'G{r}', f'=IF(F{r}="",0,MIN(E{r},MAX(0,$D${q + 1}-SUMIFS($E${k0}:$E${k1},$F${k0}:$F${k1},"<"&F{r}))))', 'calc', BRL)
        sh.put(f'I{r}', f'=IF(E{r}=0,0,G{r}/E{r})', 'calc', PCT)
        sh.cf_formula(f'B{r}:I{r}', f'$B{r}=""', COR['cinza'], COR['cinza'])
    sh.put(f'H{q + 2}', 'Chave (apoio)', 'hdr')
    r = k1 + 1
    sh.put(f'B{r}', 'Total', 'calc', bold=True)
    sh.put(f'E{r}', f'=SUM(E{k0}:E{k1})', 'calc', BRL, bold=True)
    sh.put(f'G{r}', f'=SUM(G{k0}:G{k1})', 'calc', BRL, bold=True)
    sh.put(f'I{r}', f'=IF(E{r}=0,0,G{r}/E{r})', 'calc', PCT, bold=True)
    ctx.name('Prio_Aporte', S, f'G{k0}:G{k1}')
    ctx.name('Prio_Ordem', S, f'F{k0}:F{k1}')
    sh.cf_formula(f'I{k0}:I{r}', f'AND($B{k0}<>"",I{k0}>=1)', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'I{k0}:I{r}', f'AND($B{k0}<>"",I{k0}<1,I{k0}>0)', COR['amarelo_l'], '9A5B00', True)
    sh.cf_formula(f'I{k0}:I{r}', f'AND($B{k0}<>"",I{k0}=0)', COR['vermelho_l'], COR['vermelho'], True)
    # ---------------- metas de comportamento
    c0 = r + 3
    sh.put(f'B{c0}', '5. Metas mensais de comportamento (mês selecionado no DASHBOARD)', 'sec')
    for i, h in enumerate(['Indicador', 'Meta', 'Realizado', '% atingido', 'Situação']):
        sh.put(f'{CL(2 + i)}{c0 + 1}', h, 'hdr')
    rows = [
        ('Poupança (% da receita)', '=IF(Cfg_MetaPoup=0,"",Cfg_MetaPoup)', '=K_Poup', PCT, 'min'),
        ('Gasto máximo total (R$)', '=IF(Cfg_MaxGasto=0,"",Cfg_MaxGasto)', '=K_Desp', BRL, 'max'),
        ('Aporte mensal (R$)', '=IF(Cfg_MetaAporte=0,"",Cfg_MetaAporte)', '=K_Apo', BRL, 'min'),
        ('Patrimônio líquido ao fim do ano (R$)', '=IF(Cfg_MetaPatr=0,"",Cfg_MetaPatr)', '=Pat_Liquido', BRL, 'min'),
    ]
    b0 = c0 + 2
    for i, (lbl, meta, real, fm, mode) in enumerate(rows):
        r = b0 + i
        sh.put(f'B{r}', lbl, 'calc')
        sh.put(f'C{r}', meta, 'calc', fm)
        sh.put(f'D{r}', real, 'calc', fm)
        sh.put(f'E{r}', f'=IF(C{r}="","",IFERROR(D{r}/C{r},0))', 'calc', PCT)
        if mode == 'min':
            sh.put(f'F{r}', f'=IF(C{r}="","—",IF(D{r}>=C{r},"✔","✘"))', 'calc', h='center', bold=True)
        else:
            sh.put(f'F{r}', f'=IF(C{r}="","—",IF(D{r}<=C{r},"✔","✘"))', 'calc', h='center', bold=True)
    for k in range(1, N_CATD + 1):
        r = b0 + 4 + k - 1
        sh.put(f'B{r}', f'=IF(INDEX(CatD_Nome,{k})="","","Gasto máx. "&INDEX(CatD_Nome,{k}))', 'calc')
        sh.put(f'C{r}', f'=IF(N(INDEX(CatD_Max,{k}))=0,"",INDEX(CatD_Max,{k}))', 'calc', BRL)
        sh.put(f'D{r}', f'=INDEX(Orc_Real,{k})', 'calc', BRL)
        sh.put(f'E{r}', f'=IF(C{r}="","",IFERROR(D{r}/C{r},0))', 'calc', PCT)
        sh.put(f'F{r}', f'=IF(B{r}="","",IF(C{r}="","—",IF(D{r}<=C{r},"✔","✘")))', 'calc', h='center', bold=True)
    b1 = b0 + 4 + N_CATD - 1
    sh.cf_formula(f'F{b0}:F{b1}', f'F{b0}="✔"', COR['verde_l'], COR['verde'], True)
    sh.cf_formula(f'F{b0}:F{b1}', f'F{b0}="✘"', COR['vermelho_l'], COR['vermelho'], True)
    ctx.name('Comp_Check', S, f'F{b0}:F{b1}')
    # ---------------- conquistas
    z = b1 + 3
    sh.put(f'B{z}', '6. Painel de conquistas', 'sec')
    cards = [('Meses seguidos cumprindo a meta de poupança', '=Conq_Streak', INT, 'Conq_StreakV'),
             ('Metas concluídas', '=COUNTIF(Meta_Status,"Concluída")', INT, 'Conq_Concl'),
             ('Pontuação de disciplina (0 a 100)', '=Conq_Disc', INT, 'Conq_DiscV')]
    for i, (lbl, fx, fm, n) in enumerate(cards):
        r = z + 1 + i
        sh.put(f'B{r}', lbl, 'lbl')
        sh.put(f'C{r}', fx, 'calc', fm, bold=True, size=14, h='center')
        ctx.name(n, S, f'C{r}')
    sh.put(f'D{z + 3}', '40% meses dentro do orçamento + 30% aportes feitos + 30% metas no prazo (últimos 6 meses)', 'note')
    sh.cf(f'C{z + 3}', CellIsRule(operator='greaterThanOrEqual', formula=['70'], fill=PatternFill('solid', start_color=COR['verde_l'], end_color=COR['verde_l'])))
    sh.cf(f'C{z + 3}', CellIsRule(operator='between', formula=['40', '69.99'], fill=PatternFill('solid', start_color=COR['amarelo_l'], end_color=COR['amarelo_l'])))
    sh.cf(f'C{z + 3}', CellIsRule(operator='lessThan', formula=['40'], fill=PatternFill('solid', start_color=COR['vermelho_l'], end_color=COR['vermelho_l'])))
    sh.finish(freeze='C9')
    return sh


def build_hist24(ctx):
    """Histórico de 24 meses (apoio) para sequência de poupança e disciplina — fica em CÁLCULOS."""
    c = ctx.calc
    ws, sh = c.ws, c.sh
    sh.put('F5', 'Histórico de 24 meses terminando no mês selecionado', 'sec')
    heads = ['j', 'Índice', 'Receitas', 'Despesas', 'Aportes', 'Despesa planejada', 'Poupou?', 'Dentro do orçamento?',
             'Aportou?', 'Sequência']
    for i, h in enumerate(heads):
        sh.put(f'{CL(6 + i)}6', h, 'hdr')
    for j in range(24):
        r = 7 + j
        sh.put(f'F{r}', j, 'calc')
        sh.put(f'G{r}', f'=Sel_Idx-F{r}', 'calc')
        sh.put(f'H{r}', f'={SM("Receita", f"G{r}")}', 'calc', BRLZ)
        sh.put(f'I{r}', f'={SM("Despesa", f"G{r}")}', 'calc', BRLZ)
        sh.put(f'J{r}', f'={SM("Investimento", f"G{r}")}', 'calc', BRLZ)
        sh.put(f'K{r}', f'=INDEX(Orc_TotDesp,MOD(G{r}-1,12)+1)', 'calc', BRLZ)
        sh.put(f'L{r}', f'=IF(H{r}>0,IF((H{r}-I{r})/H{r}>=Cfg_MetaPoup,1,0),0)', 'calc')
        sh.put(f'M{r}', f'=IF(AND(K{r}>0,I{r}>0,I{r}<=K{r}),1,0)', 'calc')
        sh.put(f'N{r}', f'=IF(J{r}>0,1,0)', 'calc')
        sh.put(f'O{r}', '=L7' if j == 0 else f'=IF(AND(O{r - 1}={j},L{r}=1),{j + 1},O{r - 1})', 'calc')
    for L in 'FGHIJKLMNO':
        ws.column_dimensions[L].width = 13
    ctx.name('Hist24_Rec', 'CÁLCULOS', 'H7:H30')
    ctx.name('Hist24_Desp', 'CÁLCULOS', 'I7:I30')
    ctx.name('Hist24_Apo', 'CÁLCULOS', 'J7:J30')
    c.sec('Conquistas')
    c.val('Conq_Streak', '=O30', 'Meses seguidos cumprindo a meta de poupança', INT)
    c.val('Conq_Orc6', '=SUM(M7:M12)', 'Meses (de 6) dentro do orçamento', INT)
    c.val('Conq_Apo6', '=SUM(N7:N12)', 'Meses (de 6) com aporte', INT)
    c.val('Conq_Prazo', '=IF(COUNTA(Meta_Nome)=0,1,(COUNTIF(Meta_Status,"No prazo")+COUNTIF(Meta_Status,"Concluída"))/'
          'MAX(1,COUNTIF(Meta_Status,"No prazo")+COUNTIF(Meta_Status,"Concluída")+COUNTIF(Meta_Status,"Atenção")+COUNTIF(Meta_Status,"Atrasada")))',
          'Fração de metas no prazo', PCT)
    c.val('Conq_Disc', '=ROUND(100*(0.4*Conq_Orc6/6+0.3*Conq_Apo6/6+0.3*Conq_Prazo),0)', 'Pontuação de disciplina', INT)
