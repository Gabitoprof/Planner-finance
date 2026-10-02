"""Injeta os valores calculados (cache) pelo LibreOffice no arquivo gerado pelo openpyxl,
preservando 100% da formatação/gráficos originais. O Excel continua recalculando ao abrir."""
import re, sys, zipfile, datetime
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.utils.datetime import to_excel

CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"((?: s="[0-9]+")?)><f>(.*?)</f><v ?/></c>', re.S)


def sheet_paths(zin):
    wb = zin.read('xl/workbook.xml').decode()
    rels = zin.read('xl/_rels/workbook.xml.rels').decode()
    rid2t = dict(re.findall(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
    rid2t.update({a: b for b, a in re.findall(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    out = {}
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
        name = m.group(1).replace('&amp;', '&')
        t = rid2t[m.group(2)]
        out['xl/' + t if not t.startswith('/') else t[1:]] = name
    return out


def inject(src, calc, out):
    wbv = openpyxl.load_workbook(calc, data_only=True)
    zin = zipfile.ZipFile(src)
    paths = sheet_paths(zin)
    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    total = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in paths:
            ws = wbv[paths[item.filename]]
            txt = data.decode('utf-8')

            def rep(m):
                nonlocal total
                ref, s, f = m.group(1), m.group(2), m.group(3)
                v = ws[ref].value
                total += 1
                if isinstance(v, bool):
                    return f'<c r="{ref}"{s} t="b"><f>{f}</f><v>{int(v)}</v></c>'
                if isinstance(v, (int, float)):
                    return f'<c r="{ref}"{s}><f>{f}</f><v>{v!r}</v></c>'
                if isinstance(v, (datetime.datetime, datetime.date)):
                    return f'<c r="{ref}"{s}><f>{f}</f><v>{float(to_excel(v))!r}</v></c>'
                return f'<c r="{ref}"{s} t="str"><f>{f}</f><v>{escape("" if v is None else str(v))}</v></c>'
            txt = CELL.sub(rep, txt)
            data = txt.encode('utf-8')
        zout.writestr(item, data)
    zout.close()
    return total


if __name__ == '__main__':
    print(inject(sys.argv[1], sys.argv[2], sys.argv[3]))
