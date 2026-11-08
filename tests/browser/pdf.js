/* Import from PDF: the shared reader (fileimport.js) rebuilds the table from the text layer of a printed report.
   - the 300-row sample printed by LibreOffice Calc (right-to-left sheet, borders, #,##0.00) and by Chrome (Arabic page,
     repeated header, running title and «صفحة n من m» footer, then a second table on its own page): same rows, totals,
     dates, codes with a leading zero and Arabic names as the Excel file, and the second table becomes a second sheet;
   - the SMS055, sales-analysis, transfers and receipts dashboards and the profit-margin import read a PDF printed the way
     Excel prints these reports (left-to-right sheet, no borders, header fill, dates as 9/1/2026 0:00, long numbers as
     1.30E+13) end to end;
   - a scanned PDF (picture only) goes to the text recogniser and its review (see ocr.js); a password-protected PDF and a
     damaged PDF get clear Arabic messages;
   - pdf.js is loaded only when a PDF is chosen. Fixtures: tests/fixtures/formats (make-pdf.js rebuilds the PDFs). */
const { open } = require('./harness');
const fs = require('fs'); const path = require('path');
const DIR = path.join(__dirname, '..', 'fixtures', 'formats');
const EXP = JSON.parse(fs.readFileSync(path.join(DIR, 'expected.json'), 'utf8'));
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' } };
const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
const closures = fs.readFileSync(path.join(__dirname, 'closures.js'), 'utf8');
const MAKE = closures.slice(closures.indexOf('const MAKE = `') + 14, closures.indexOf('`;', closures.indexOf('const MAKE = `')));
const out = {};
/* a report printed the way Excel prints it: left-to-right sheet, gridlines off, header row filled, numbers right-aligned */
const excelPrint = rows => {
  const head = Object.keys(rows[0]); const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const cell = v => typeof v === 'number' ? (v >= 1e11 ? v.toExponential(2).toUpperCase().replace('E+', 'E+') : String(v)) : esc(v == null ? '' : v);
  return `<!doctype html><html lang="ar"><head><meta charset="utf-8"><style>body{margin:0;font:8px "DejaVu Sans",Arial,sans-serif}table{border-collapse:collapse}
   th{background:#1f3864;color:#fff;padding:1px 3px;text-align:center;white-space:nowrap}td{padding:1px 3px;white-space:nowrap}td.n{text-align:right}thead{display:table-header-group}</style></head>
   <body><table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => '<tr>' + head.map(h => `<td${typeof r[h] === 'number' ? ' class="n"' : ''}>${cell(r[h])}</td>`).join('') + '</tr>').join('')}</tbody></table></body></html>`;
};
(async () => {
  const { browser, ctx, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data });
  const loaded = () => page.evaluate(() => !!window.pdfjsLib);
  out.lazy = !(await loaded());
  const read = (name, buf) => page.evaluate(async ({ name, b64 }) => {
    const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    try { return { sheets: (await FileImport.readGrids(new File([u8], name))).map(s => ({ name: s.name, grid: s.grid })) }; } catch (e) { return { err: e.message }; }
  }, { name, b64: buf.toString('base64') });
  // the sample report: same content as the Excel file
  for (const f of ['sample-calc.pdf', 'sample-print.pdf']) {
    const r = await read(f, fs.readFileSync(path.join(DIR, f)));
    if (r.err) { out[f] = r.err; continue; }
    const s = r.sheets[0]; const hi = s.grid.findIndex(row => row[0] === 'التاريخ'); const head = s.grid[hi] || [];
    const body = s.grid.slice(hi + 1).filter(row => row.some(x => String(x).trim()));
    const iso = v => { const m = String(v).match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : v; };
    const val = Math.round(body.reduce((a, row) => a + (+row[8] || 0), 0) * 100) / 100, qty = body.reduce((a, row) => a + (+row[7] || 0), 0);
    const ret = body.filter(row => row[9] !== '' && row[9] != null);
    const res = { head: JSON.stringify(head) === JSON.stringify(EXP.head), rows: body.length === EXP.rows, value: Math.abs(val - EXP.value) < 0.01, qty: qty === EXP.qty,
      numbers: body.every(row => typeof row[7] === 'number' && typeof row[8] === 'number'), firstDate: iso(body[0][0]) === EXP.first.date, lastDate: iso(body[body.length - 1][0]) === EXP.last.date,
      code: body[0][1] === EXP.first.customer, arabic: body[0][2] === EXP.first.name && body[body.length - 1][2] === EXP.last.name,
      ligatures: body.some(row => row[4] === 'غاز البترول المسال') && body.some(row => row[5] === 'لتر') && body.some(row => /السلام/.test(row[2])),
      returns: ret.length === 9 && ret.every(row => row[9] === 5), noFooters: !s.grid.some(row => row.some(x => /صفحة|^Page/.test(String(x)))), repeatedHeaderOnce: s.grid.filter(row => row[0] === 'التاريخ').length === 1 };
    if (f === 'sample-print.pdf') { const t = r.sheets[1]; res.secondSheet = r.sheets.length === 2 && !!t && t.grid.some(row => row[0] === 'رقم المنتج' && row[4] === 'القيمة') && t.grid.some(row => row[0] === 16291 && row[1] === 'غاز البترول المسال' && row[2] === 100);
      res.noRunningTitle = !r.sheets.some(x => x.grid.some(row => row.some(c => /الإدارة العامة/.test(String(c))))); }
    const bad = Object.keys(res).filter(k => !res[k]); out[f] = bad.length ? 'mismatch: ' + bad.join(',') + ' ' + JSON.stringify(body[0]) : true;
  }
  out.loadedOnDemand = await loaded();
  // clear messages
  const msg = async (name, buf) => (await read(name, buf)).err || 'read';
  /* a scan goes to the text recogniser (tests/browser/ocr.js covers it); here: it starts, and cancelling imports nothing */
  const pending = read('scan.pdf', fs.readFileSync(path.join(DIR, 'sample-scan.pdf')));
  out.scanned = await page.waitForSelector('.ocr-dlg', { timeout: 30000 }).then(() => page.click('.ocr-dlg [data-ocr="cancel"]')).then(async () => /أُلغي/.test((await pending).err || ''), () => false);
  out.locked = /محمي بكلمة مرور/.test(await msg('locked.pdf', fs.readFileSync(path.join(DIR, 'sample-locked.pdf'))));
  out.damaged = /تالف/.test(await msg('broken.pdf', Buffer.from('%PDF-1.7\n1 0 obj << /Type /Catalog >>\n%%EOF garbage')));
  out.listNamesPdf = await page.evaluate(() => /PDF/.test(FileImport.LIST) && /\.pdf\b/.test(FileImport.ACCEPT));
  // end to end through the dashboards, from PDFs printed like the real Excel reports
  await page.evaluate(MAKE);
  const printer = await ctx.newPage();
  const pdfOf = async rows => { await printer.setContent(excelPrint(rows)); return printer.pdf({ format: 'A4', landscape: true, margin: { top: '10mm', bottom: '10mm', left: '6mm', right: '6mm' } }); };
  for (const [id, kind, months, file] of [['sales', 'sales', [[2026, 9, 30]], 'Sales_Report_20261105.pdf'], ['transfers', 'transfers', [[2026, 9, 30]], 'Transfers_Report_20261105.pdf'], ['receipts', 'receipts', [[2026, 9, 30]], 'Receipts_Report_20261105.pdf']]) {
    const rows = await page.evaluate(({ kind, months }) => __cxBook(kind, months, 11, 0), { kind, months });
    await page.evaluate(v => go(v), id); await page.waitForTimeout(600);
    await page.setInputFiles('#cx-file-' + id, { name: file, mimeType: 'application/pdf', buffer: await pdfOf(rows) }); await page.waitForTimeout(2500);
    out['dash:' + id] = await page.evaluate(n => { const d = document.getElementById('dlg'); if (!d.open) return 'no dialog'; const ok = /مراجعة ملف/.test(d.textContent) && [...d.querySelectorAll('.sm-cols-map span')].every(s => s.classList.contains('ok'));
      const lines = (d.querySelector('.sm-mini b') || {}).textContent || ''; return ok && lines.replace(/[^\d]/g, '') === String(n) ? true : d.textContent.slice(0, 300); }, rows.length);
    await page.evaluate(() => dlg.close());
  }
  // SMS055 from the Calc print and from an Excel-style print of the same rows
  await page.evaluate(() => go('sms055')); await page.waitForTimeout(500);
  const smsRows = JSON.parse(fs.readFileSync(path.join(DIR, 'sample.json'), 'utf8')).map(r => Object.assign({}, r, { 'التاريخ': r['التاريخ'].replace(/^(\d\d)\/(\d\d)\/(\d{4})$/, (m, d, mo, y) => `${+mo}/${+d}/${y} 0:00`) }));
  for (const [f, buf] of [['SMS055_Report.pdf', fs.readFileSync(path.join(DIR, 'sample-calc.pdf'))], ['SMS055_Report (3).pdf', await pdfOf(smsRows)]]) {
    await page.setInputFiles('#sm-file', { name: f, mimeType: 'application/pdf', buffer: buf }); await page.waitForTimeout(2500);
    out['sms:' + f] = await page.evaluate(() => dlg.open && /2026\/07\/01 – 2026\/08\/31/.test(dlg.textContent) && document.querySelectorAll('.sm-cols-map .ok').length >= 11 || (dlg.open ? dlg.textContent.slice(0, 300) : 'no dialog'));
    await page.evaluate(() => dlg.close());
  }
  // profit margin: a sales-analysis PDF opens the sales review
  const salesRows = await page.evaluate(() => __cxBook('sales', [[2026, 9, 30]], 5, 0));
  await page.evaluate(() => go('margin')); await page.waitForTimeout(700);
  await page.setInputFiles('#pm-file', { name: 'Sales_Report.pdf', mimeType: 'application/pdf', buffer: await pdfOf(salesRows) }); await page.waitForTimeout(2500);
  out.margin = await page.evaluate(() => dlg.open && /مراجعة ملف تحليل المبيعات/.test(dlg.textContent) || (dlg.open ? dlg.textContent.slice(0, 300) : 'no dialog'));
  await page.evaluate(() => dlg.close());
  // every dashboard import input offers PDF
  const accepts = {};
  for (const v of ['sms055', 'sales', 'transfers', 'receipts', 'debts', 'margin', 'settings']) {
    await page.evaluate(v => go(v), v); await page.waitForTimeout(500);
    accepts[v] = await page.evaluate(() => { const i = [...document.querySelectorAll('#view input[type=file]')].filter(x => x.accept !== '.json,application/json'); return i.length > 0 && i.every(x => /\.pdf\b/.test(x.accept)); });
  }
  out.accepts = Object.values(accepts).every(Boolean) || accepts;
  out.errors = errors.filter(e => !/favicon|Failed to load resource/.test(e));
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.entries(out).every(([k, v]) => k === 'errors' ? v.length === 0 : v === true);
  console.log(ok ? 'PDF OK' : 'PDF FAILED');
  await browser.close(); process.exit(ok ? 0 : 1);
})();
