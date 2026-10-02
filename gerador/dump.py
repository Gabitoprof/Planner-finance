import sys, openpyxl
wb = openpyxl.load_workbook('/home/user/Planner-finance/.teste/t.xlsx', data_only=True)
def show(sheet, rng, w=14):
    ws = wb[sheet]
    for row in ws[rng]:
        vals = []
        for c in row:
            v = c.value
            if isinstance(v, float): v = round(v, 2)
            if v is None: v = ''
            vals.append(str(v)[:w].ljust(w))
        print(' '.join(vals))
if __name__ == '__main__':
    show(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 14)
