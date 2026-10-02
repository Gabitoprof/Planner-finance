import openpyxl, subprocess, shutil, datetime as dt, sys, glob, os
SRC = '/home/user/Planner-finance/Planilha_Financeira_Avancada.xlsx'
T = '/home/user/Planner-finance/.teste'
SK = glob.glob('/root/.claude/skills/synced/*/xlsx')[0]

def recalc(path):
    r = subprocess.run(['python3', f'{SK}/scripts/recalc.py', path, '580'], capture_output=True, text=True, cwd=T)
    return r.stdout[:300]

def names(wb):
    return wb.defined_names

def val(wbv, name):
    d = wbv.defined_names[name]
    sheet, ref = list(d.destinations)[0]
    return wbv[sheet][ref.replace('$', '')].value

def snapshot(path):
    wbv = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for n in ['K_Rec', 'K_Desp', 'K_Saldo', 'K_Poup', 'K_Apo', 'H_Nota', 'Sel_Label', 'Orc_TotPlan', 'Orc_TotReal', 'Div_Total', 'Pat_Liquido', 'Card_UsadoTot']:
        out[n] = val(wbv, n)
    out['alert_txt'] = wbv['DASHBOARD']['B34'].value
    return wbv, out

def first_empty_row(ws, r0=7):
    r = r0
    while ws.cell(row=r, column=1).value is not None:
        r += 1
    return r

base = snapshot(f'{T}/t.xlsx')[1]
print('BASE', base)

# 1) trocar mês/ano
wb = openpyxl.load_workbook(SRC)
wb['DASHBOARD']['C6'] = 'Agosto'
wb.save(f'{T}/t1.xlsx'); print(recalc(f'{T}/t1.xlsx')[:80])
v1, s1 = snapshot(f'{T}/t1.xlsx'); print('AGOSTO', s1)
wb = openpyxl.load_workbook(SRC)
wb['DASHBOARD']['C6'] = 'Dezembro'; wb['DASHBOARD']['E6'] = 2025
wb.save(f'{T}/t1b.xlsx'); recalc(f'{T}/t1b.xlsx')
print('DEZ/2025', snapshot(f'{T}/t1b.xlsx')[1])

# 2) novo lançamento (Lazer 500 no cartão em set/2026)
wb = openpyxl.load_workbook(SRC)
ws = wb['LANÇAMENTOS']
r = first_empty_row(ws)
for c, v in enumerate([dt.date(2026, 9, 20), 'Teste novo lançamento', 'Despesa', 'Lazer', None, 'Cartão Azul Gold', 'Crédito', None, None, 'Pendente', None, None, 'TESTE', 500], 1):
    if v is not None: ws.cell(row=r, column=c, value=v)
wb.save(f'{T}/t2.xlsx'); recalc(f'{T}/t2.xlsx')
v2, s2 = snapshot(f'{T}/t2.xlsx'); print('NOVO LANC', s2)
wsF = v2['FLUXO DE CAIXA']
for row in wsF.iter_rows(min_row=10, max_row=45, min_col=2, max_col=15):
    if row[0].value == 'Lazer': print('Fluxo Lazer set:', row[8].value)
wsO = v2['ORÇAMENTO']
for row in wsO.iter_rows(min_row=58, max_row=80, min_col=2, max_col=6):
    if row[0].value == 'Lazer': print('Orçamento Lazer:', [c.value for c in row])
dsh = v2['DASHBOARD']
print('Alertas:', [dsh.cell(row=r_, column=2).value for r_ in range(26, 34)])

# 3) nova categoria
wb = openpyxl.load_workbook(SRC)
cfg = wb['CONFIGURAÇÕES']
# achar bloco de categorias de despesa: primeira célula vazia após 'Outros' na coluna B depois de 'Moradia'
rows = [r for r in range(1, 140) if cfg.cell(row=r, column=2).value == 'Moradia']
r0 = rows[0]
slot = r0 + 13
assert cfg.cell(row=slot, column=2).value is None
cfg.cell(row=slot, column=2, value='Cursos Online'); cfg.cell(row=slot, column=3, value='Variável'); cfg.cell(row=slot, column=4, value='Não essencial')
cfg.cell(row=slot, column=5, value='Desejos'); cfg.cell(row=slot, column=6, value=200)
orc = wb['ORÇAMENTO']
for r in range(1, 60):
    if orc.cell(row=r, column=2).value == 'DESPESAS': ro = r + 14; break
orc.cell(row=ro, column=2)  # linha do slot 14 (planejamento)
orc.cell(row=ro, column=3 + 8, value=200)  # Set
ws = wb['LANÇAMENTOS']
r = first_empty_row(ws)
for c, v in enumerate([dt.date(2026, 9, 12), 'Curso de dados', 'Despesa', 'Cursos Online', None, 'Conta Corrente Banco Azul', 'Pix', None, None, 'Pago', None, None, 'TESTE', 150], 1):
    if v is not None: ws.cell(row=r, column=c, value=v)
wb.save(f'{T}/t3.xlsx'); print(recalc(f'{T}/t3.xlsx')[:80])
v3, s3 = snapshot(f'{T}/t3.xlsx'); print('NOVA CAT', s3)
wsO = v3['ORÇAMENTO']
for row in wsO.iter_rows(min_row=40, max_row=110, min_col=2, max_col=11):
    if row[0].value == 'Cursos Online': print('Orçamento:', [round(c.value, 2) if isinstance(c.value, float) else c.value for c in row])
wsF = v3['FLUXO DE CAIXA']
for row in wsF.iter_rows(min_row=10, max_row=45, min_col=2, max_col=15):
    if row[0].value == 'Cursos Online': print('Fluxo set:', row[8].value)
tb = v3['TABELAS']
print('Lista cat inclui nova?', any(c.value == 'Cursos Online' for col in tb.iter_cols(min_row=7, max_row=60) for c in col))
cx = v3['CÁLCULOS']
print('Ranking:', [(cx.cell(row=r_, column=24).value, cx.cell(row=r_, column=25).value) for r_ in range(7, 16)])
