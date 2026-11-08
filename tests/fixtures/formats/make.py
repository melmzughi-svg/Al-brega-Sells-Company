#!/usr/bin/env python3
"""Builds one sample workbook in every import format the system reads (tests/browser/formats.js uses them).
The source is a flat OpenDocument file written here with typed cells (dates, numbers, Arabic text, a code with a
leading zero); LibreOffice converts it to the other formats. WPS Spreadsheets writes .et/.ett in the same binary
(OLE2/BIFF8) layout as .xls, and newer WPS versions can also save them in the xlsx layout, so both are produced.
Run: python3 tests/fixtures/formats/make.py   (needs LibreOffice: soffice)"""
import datetime, json, os, shutil, subprocess, tempfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
HEAD = ['التاريخ', 'رقم العميل', 'اسم الزبون', 'رقم المنتج', 'اسم المنتج', 'وحدة القياس', 'النوع', 'الكمية', 'القيمة', 'الكمية المسترجعة', 'القيمة المسترجعة']
PRODUCTS = [('16291', 'غاز البترول المسال', 'اسطوانة', 1.25), ('33500', 'بنزين 95', 'لتر', 0.15), ('44347', 'وقود الديزل', 'لتر', 0.14)]
PLACES = ['النور', 'الأمل', 'الوحدة', 'السلام', 'الفجر', 'البركة', 'الريان', 'الهدى']
N = 300  # enough distinct strings for the .xls shared-string table to run over CONTINUE records

def rows():
    out = []; s = 7
    for i in range(N):
        s = (s * 16807) % 2147483647
        d = datetime.date(2026, 7, 1) + datetime.timedelta(days=i % 62)
        p = PRODUCTS[i % 3]
        q = 20 + s % 400 if p[0] == '16291' else 3000 + s % 30000
        v = round(q * p[3], 2)
        ret = i % 37 == 0
        out.append([d, '0%d' % (110019 + (i % 40) * 37), 'محطة وقود %s %d' % (PLACES[i % 8], i + 1), p[0], p[1], p[2], 'CH', q, v, 5 if ret else None, round(5 * p[3], 2) if ret else None])
    return out

def cell(v):
    if v is None: return '<table:table-cell/>'
    if isinstance(v, datetime.date):
        return '<table:table-cell table:style-name="d" office:value-type="date" office:date-value="%s"><text:p>%s</text:p></table:table-cell>' % (v.isoformat(), v.isoformat())
    if isinstance(v, (int, float)):
        return '<table:table-cell office:value-type="float" office:value="%s"><text:p>%s</text:p></table:table-cell>' % (v, v)
    return '<table:table-cell office:value-type="string"><text:p>%s</text:p></table:table-cell>' % escape(v)

def fods(data):
    tr = lambda r: '<table:table-row>' + ''.join(cell(v) for v in r) + '</table:table-row>'
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
            'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
            'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
            'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">'
            '<office:automatic-styles><number:date-style style:name="N1"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>'
            '<style:style style:name="d" style:family="table-cell" style:data-style-name="N1"/></office:automatic-styles>'
            '<office:body><office:spreadsheet><table:table table:name="SMS055_Report-1">' + tr(HEAD) + ''.join(tr(r) for r in data) +
            '</table:table><table:table table:name="ملاحظات"><table:table-row>' + cell('ملاحظة') + '</table:table-row><table:table-row>' + cell('ورقة ثانية') + '</table:table-row></table:table>'
            '</office:spreadsheet></office:body></office:document>')

def convert(src, fmt, out_name, workdir):
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', fmt, '--outdir', workdir, src], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    produced = os.path.join(workdir, os.path.splitext(os.path.basename(src))[0] + '.' + fmt.split(':')[0])
    shutil.move(produced, os.path.join(HERE, out_name))

def xlsb(data, path):
    """A small Excel binary workbook written record by record (LibreOffice cannot save .xlsb)."""
    import struct, zipfile
    def rec(t, body=b''):
        h = bytes([t]) if t < 128 else bytes([(t & 127) | 128, t >> 7])
        n = len(body); sz = b''
        while True:
            b = n & 127; n >>= 7
            sz += bytes([b | (128 if n else 0)])
            if not n: break
        return h + sz + body
    ws = lambda s: struct.pack('<I', len(s)) + s.encode('utf-16-le')
    sst, idx = [], {}
    def si(s):
        if s not in idx: idx[s] = len(sst); sst.append(s)
        return idx[s]
    body = rec(129) + rec(145)
    for r, row in enumerate([HEAD] + data):
        body += rec(0, struct.pack('<IIHHBI', r, 0, 300, 0, 0, 0))
        for c, v in enumerate(row):
            if v is None: continue
            if isinstance(v, datetime.date):
                serial = (v - datetime.date(1899, 12, 30)).days
                body += rec(2, struct.pack('<III', c, 1, (serial << 2) | 2))
            elif isinstance(v, int):
                body += rec(2, struct.pack('<III', c, 0, (v << 2) | 2))
            elif isinstance(v, float):
                body += rec(5, struct.pack('<IId', c, 0, v))
            else:
                body += rec(7, struct.pack('<III', c, 0, si(v)))
    body += rec(146) + rec(130)
    sstb = rec(159, struct.pack('<II', len(sst), len(sst))) + b''.join(rec(19, b'\x00' + ws(x)) for x in sst) + rec(160)
    xf = lambda fmt: rec(47, struct.pack('<HHHHHHBBH', 0xFFFF, fmt, 0, 0, 0, 0, 0, 0, 0))
    sty = rec(278) + rec(617, struct.pack('<I', 2)) + xf(0) + xf(14) + rec(618) + rec(279)
    wb = rec(131) + rec(143) + rec(156, struct.pack('<II', 0, 1) + ws('rId1') + ws('SMS055_Report-1')) + rec(144) + rec(132)
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet"/><Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/><Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/></Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + R + 'officeDocument" Target="xl/workbook.bin"/></Relationships>')
        z.writestr('xl/_rels/workbook.bin.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + R + 'worksheet" Target="worksheets/sheet1.bin"/><Relationship Id="rId2" Type="' + R + 'sharedStrings" Target="sharedStrings.bin"/><Relationship Id="rId3" Type="' + R + 'styles" Target="styles.bin"/></Relationships>')
        z.writestr('xl/workbook.bin', wb); z.writestr('xl/worksheets/sheet1.bin', body); z.writestr('xl/sharedStrings.bin', sstb); z.writestr('xl/styles.bin', sty)

def main():
    data = rows()
    with tempfile.TemporaryDirectory() as w:
        src = os.path.join(w, 'sample.fods')
        open(src, 'w', encoding='utf-8').write(fods(data))
        shutil.copy(src, os.path.join(HERE, 'sample.fods'))
        convert(src, 'xlsx', 'sample.xlsx', w)
        convert(src, 'xls', 'sample.xls', w)
        convert(src, 'ods', 'sample.ods', w)
        convert(src, 'xml:MS Excel 2003 XML', 'sample.xml', w)
        convert(src, 'html:HTML (StarCalc)', 'sample.html', w)
        convert(src, 'xlsm:Calc MS Excel 2007 VBA XML', 'sample.xlsm', w)
    xlsb(data, os.path.join(HERE, 'sample.xlsb'))
    shutil.copy(os.path.join(HERE, 'sample.xls'), os.path.join(HERE, 'sample.et'))      # WPS Spreadsheets (binary layout)
    shutil.copy(os.path.join(HERE, 'sample.xls'), os.path.join(HERE, 'sample.ett'))     # WPS template
    shutil.copy(os.path.join(HERE, 'sample.xlsx'), os.path.join(HERE, 'sample-xlsx-layout.et'))  # WPS .et saved in the xlsx layout
    fmt = lambda v: '' if v is None else v.strftime('%d/%m/%Y') if isinstance(v, datetime.date) else str(v)
    def delim(sep, quote_all=False):
        q = lambda s: '"' + s.replace('"', '""') + '"' if (quote_all or sep in s or '"' in s) else s
        return '\r\n'.join(sep.join(q(fmt(v)) for v in r) for r in [HEAD] + data) + '\r\n'
    open(os.path.join(HERE, 'sample.csv'), 'w', encoding='utf-8-sig', newline='').write(delim(','))
    open(os.path.join(HERE, 'sample-semicolon.csv'), 'w', encoding='utf-8', newline='').write(delim(';', True))
    open(os.path.join(HERE, 'sample-1256.csv'), 'w', encoding='cp1256', newline='').write(delim(','))
    open(os.path.join(HERE, 'sample.tsv'), 'w', encoding='utf-8', newline='').write(delim('\t'))
    open(os.path.join(HERE, 'sample-utf16.txt'), 'w', encoding='utf-16', newline='').write(delim('\t'))   # WPS/Excel «Unicode text»
    open(os.path.join(HERE, 'sample-pipe.txt'), 'w', encoding='utf-8', newline='').write(delim('|'))
    js = [{k: (fmt(v) if isinstance(v, datetime.date) else v) for k, v in zip(HEAD, r) if v is not None} for r in data]
    open(os.path.join(HERE, 'sample.json'), 'w', encoding='utf-8').write(json.dumps(js, ensure_ascii=False))
    expected = {'rows': N, 'head': HEAD, 'value': round(sum(r[8] for r in data), 2), 'qty': sum(r[7] for r in data),
                'first': {'date': data[0][0].isoformat(), 'customer': data[0][1], 'name': data[0][2]},
                'last': {'date': data[-1][0].isoformat(), 'name': data[-1][2]}}
    open(os.path.join(HERE, 'expected.json'), 'w', encoding='utf-8').write(json.dumps(expected, ensure_ascii=False, indent=1))

if __name__ == '__main__':
    main()
