"""CÁLCULOS: valores intermediários nomeados (seleção de mês, indicadores, nota de saúde...)."""
from common import *


class Calc:
    def __init__(self, ctx):
        self.ctx = ctx
        self.sh = Sh(ctx, 'CÁLCULOS')
        self.ws = self.sh.ws
        self.r = 6
        self.sh.header('CÁLCULOS (apoio)', 'Valores intermediários usados pelo Dashboard e pelas demais abas. '
                       'Não precisa mexer. Esta aba pode ficar oculta.', ncols=8)
        self.sh.widths({'B': 44, 'C': 20, 'D': 16})

    def sec(self, titulo):
        self.r += 1
        self.sh.put(f'B{self.r}', titulo, 'sec')
        self.r += 1

    def val(self, nome, formula, label, fmt=None):
        self.sh.put(f'B{self.r}', label, 'lbl')
        self.sh.put(f'C{self.r}', formula, 'calc', fmt, h='right')
        self.ctx.name(nome, 'CÁLCULOS', f'C{self.r}')
        self.r += 1


def build_calc_core(ctx):
    c = Calc(ctx)
    c.sec('Período de referência')
    c.val('Cfg_Ano', '=IF(N(Cfg_AnoIn)=0,YEAR(Cfg_Hoje),Cfg_AnoIn)', 'Ano de referência (efetivo)', INT)
    c.val('Cfg_Mes', '=IF(Cfg_MesIn="",MONTH(Cfg_Hoje),IFERROR(MATCH(Cfg_MesIn,Lista_Meses,0),MONTH(Cfg_Hoje)))',
          'Mês atual (1-12)', INT)
    c.val('Cfg_Idx', '=Cfg_Ano*12+Cfg_Mes', 'Índice do mês atual (ano*12+mês)', INT)
    c.val('Sel_Mes', '=IF(Dash_Mes="",Cfg_Mes,IFERROR(MATCH(Dash_Mes,Lista_Meses,0),Cfg_Mes))',
          'Mês selecionado no Dashboard (1-12)', INT)
    c.val('Sel_Ano', '=IF(N(Dash_Ano)=0,Cfg_Ano,Dash_Ano)', 'Ano selecionado no Dashboard', INT)
    c.val('Sel_Idx', '=Sel_Ano*12+Sel_Mes', 'Índice do mês selecionado', INT)
    c.val('Sel_Ini', '=DATE(Sel_Ano,Sel_Mes,1)', 'Primeiro dia do mês selecionado', DATA)
    c.val('Sel_Fim', '=DATE(Sel_Ano,Sel_Mes+1,1)-1', 'Último dia do mês selecionado', DATA)
    c.val('Sel_Dias', '=DAY(Sel_Fim)', 'Dias do mês selecionado', INT)
    c.val('Sel_MesNome', '=INDEX(Lista_Meses,Sel_Mes)', 'Nome do mês selecionado')
    c.val('Sel_Label', '=Sel_MesNome&"/"&Sel_Ano', 'Rótulo do período')
    c.val('Cfg_Label', '=INDEX(Lista_Meses,Cfg_Mes)&"/"&Cfg_Ano', 'Rótulo do mês atual')
    ctx.calc = c
    return c
