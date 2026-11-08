/* Builds the PDF samples for tests/browser/formats.js from the same 300 rows as the other formats (sample.json):
     sample-calc.pdf   the report as a right-to-left spreadsheet printed to PDF by LibreOffice Calc (borders, number
                       format #,##0.00, dates dd/mm/yyyy, header row repeated on every page, «Page n» footer)
     sample-print.pdf  the report as an Arabic web page printed by Chrome (repeated table header, running title and
                       «صفحة n من m» footer), followed on a new page by a second table with its own header
     sample-scan.pdf   a page that is only a picture of the report (what a scanner or a phone photo gives)
     sample-locked.pdf the first page of sample-print.pdf protected with a password
   Run from the repo root: node tests/fixtures/formats/make-pdf.js   (needs LibreOffice: soffice, and Playwright) */
const fs = require('fs'); const path = require('path'); const os = require('os'); const { execFileSync } = require('child_process');
const { chromium } = require('playwright');
const HERE = __dirname;
const EXP = JSON.parse(fs.readFileSync(path.join(HERE, 'expected.json'), 'utf8'));
const rows = JSON.parse(fs.readFileSync(path.join(HERE, 'sample.json'), 'utf8'));
const HEAD = EXP.head;
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const iso = s => { const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : s; };
const money = v => v == null ? '' : Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const qty = v => v == null ? '' : Number(v).toLocaleString('en-US');

function fods() {
  const cell = (k, v) => {
    if (v == null) return '<table:table-cell table:style-name="b"/>';
    if (k === 'التاريخ') return `<table:table-cell table:style-name="dt" office:value-type="date" office:date-value="${iso(v)}"><text:p>${v}</text:p></table:table-cell>`;
    if (typeof v === 'number') return `<table:table-cell table:style-name="${/القيمة/.test(k) ? 'mo' : 'nu'}" office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
    return `<table:table-cell table:style-name="b" office:value-type="string"><text:p>${esc(v)}</text:p></table:table-cell>`;
  };
  const head = '<table:table-row>' + HEAD.map(h => `<table:table-cell table:style-name="h" office:value-type="string"><text:p>${esc(h)}</text:p></table:table-cell>`).join('') + '</table:table-row>';
  const body = rows.map(r => '<table:table-row>' + HEAD.map(k => cell(k, r[k])).join('') + '</table:table-row>').join('');
  const widths = ['2.3cm', '2.1cm', '4.6cm', '1.9cm', '3.8cm', '2.1cm', '1.3cm', '1.9cm', '2.4cm', '2.6cm', '2.6cm'];
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:automatic-styles>
 <number:date-style style:name="N1"><number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/></number:date-style>
 <number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
 <number:number-style style:name="N3"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
 <style:style style:name="ta" style:family="table" style:master-page-name="pm"><style:table-properties table:display="true" style:writing-mode="rl-tb"/></style:style>
 ${widths.map((w, i) => `<style:style style:name="co${i}" style:family="table-column"><style:table-column-properties style:column-width="${w}"/></style:style>`).join('')}
 <style:style style:name="b" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000"/></style:style>
 <style:style style:name="h" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000" fo:background-color="#dddddd"/><style:text-properties fo:font-weight="bold"/></style:style>
 <style:style style:name="dt" style:family="table-cell" style:data-style-name="N1"><style:table-cell-properties fo:border="0.5pt solid #000000"/></style:style>
 <style:style style:name="mo" style:family="table-cell" style:data-style-name="N2"><style:table-cell-properties fo:border="0.5pt solid #000000"/></style:style>
 <style:style style:name="nu" style:family="table-cell" style:data-style-name="N3"><style:table-cell-properties fo:border="0.5pt solid #000000"/></style:style>
 <style:page-layout style:name="pl"><style:page-layout-properties fo:page-width="29.7cm" fo:page-height="21cm" style:print-orientation="landscape" fo:margin="1cm" style:scale-to-X="1" style:scale-to-Y="0"/></style:page-layout>
</office:automatic-styles>
<office:master-styles><style:master-page style:name="pm" style:page-layout-name="pl"><style:header style:display="false"/><style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer></style:master-page></office:master-styles>
<office:body><office:spreadsheet><table:table table:name="SMS055_Report-1" table:style-name="ta">${widths.map((_, i) => `<table:table-column table:style-name="co${i}"/>`).join('')}
<table:table-header-rows>${head}</table:table-header-rows>${body}</table:table></office:spreadsheet></office:body></office:document>`;
}

function html() {
  const td = (k, v) => v == null ? '<td></td>' : typeof v === 'number' ? `<td class="n">${/القيمة/.test(k) ? money(v) : qty(v)}</td>` : `<td>${esc(v)}</td>`;
  return `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>
  body{font-family:"DejaVu Sans",Arial,sans-serif;font-size:9px;margin:0}
  h1{font-size:14px;margin:0 0 6px} table{border-collapse:collapse;width:100%} th,td{border:1px solid #444;padding:2px 4px;text-align:right;white-space:nowrap}
  th{background:#e6e6e6} td.n{text-align:left;font-variant-numeric:tabular-nums} tr{break-inside:avoid} .pb{break-before:page}
  </style></head><body><h1>تقرير المبيعات التفصيلي — الفترة من 01/07/2026 إلى 31/08/2026</h1>
  <table><thead><tr>${HEAD.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>
  ${rows.map(r => '<tr>' + HEAD.map(k => td(k, r[k])).join('') + '</tr>').join('')}</tbody></table>
  <div class="pb"></div><h1>ملخص حسب المنتج</h1>
  <table style="width:60%"><thead><tr><th>رقم المنتج</th><th>اسم المنتج</th><th>عدد الحركات</th><th>الكمية</th><th>القيمة</th></tr></thead><tbody>
  ${['16291', '33500', '44347'].map(p => { const rs = rows.filter(r => r['رقم المنتج'] === p); return `<tr><td>${p}</td><td>${esc(rs[0]['اسم المنتج'])}</td><td class="n">${rs.length}</td><td class="n">${qty(rs.reduce((a, r) => a + r['الكمية'], 0))}</td><td class="n">${money(rs.reduce((a, r) => a + r['القيمة'], 0))}</td></tr>`; }).join('')}
  </tbody></table></body></html>`;
}

(async () => {
  const w = fs.mkdtempSync(path.join(os.tmpdir(), 'pdfmk-'));
  fs.writeFileSync(path.join(w, 'sample-calc.fods'), fods());
  execFileSync('soffice', ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', w, path.join(w, 'sample-calc.fods')], { stdio: 'ignore' });
  fs.copyFileSync(path.join(w, 'sample-calc.pdf'), path.join(HERE, 'sample-calc.pdf'));
  const browser = await chromium.launch(); const page = await browser.newPage();
  await page.setContent(html());
  await page.pdf({ path: path.join(HERE, 'sample-print.pdf'), format: 'A4', landscape: true, margin: { top: '14mm', bottom: '14mm', left: '8mm', right: '8mm' }, displayHeaderFooter: true,
    headerTemplate: '<div dir="rtl" style="font-size:8px;width:100%;text-align:center">الإدارة العامة للمبيعات — تقرير SMS055</div>',
    footerTemplate: '<div dir="rtl" style="font-size:8px;width:100%;text-align:center">صفحة <span class="pageNumber"></span> من <span class="totalPages"></span></div>' });
  // the scanned copy: a picture of the first page of the printed report, nothing else
  await page.setViewportSize({ width: 1123, height: 794 }); await page.setContent(html()); const png = await page.screenshot({ type: 'jpeg', quality: 55 });
  await page.setContent(`<html><body style="margin:0"><img style="width:100%" src="data:image/jpeg;base64,${png.toString('base64')}"></body></html>`);
  await page.pdf({ path: path.join(HERE, 'sample-scan.pdf'), format: 'A4', landscape: true, printBackground: true, margin: { top: 0, bottom: 0, left: 0, right: 0 } });
  await browser.close(); fs.rmSync(w, { recursive: true, force: true });
  // the first page of the printed report, protected with a password (needs Python with pypdf)
  execFileSync('python3', ['-c', "import pypdf,sys;w=pypdf.PdfWriter(clone_from=sys.argv[1]);[w.remove_page(len(w.pages)-1) for _ in range(len(w.pages)-1)];w.encrypt(user_password='brega',owner_password='brega-owner');w.write(sys.argv[2])",
    path.join(HERE, 'sample-print.pdf'), path.join(HERE, 'sample-locked.pdf')]);
  for (const f of ['sample-calc.pdf', 'sample-print.pdf', 'sample-scan.pdf', 'sample-locked.pdf']) console.log(f, fs.statSync(path.join(HERE, f)).size);
})();
