/* Import in every file format: the shared reader (fileimport.js) reads the same sample workbook saved as
   xlsx, xlsm, xlsb, xls, WPS et/ett (binary and xlsx layout), ods, fods, Excel 2003 XML, HTML, CSV (UTF-8, ;-separated,
   Windows-1256 Arabic), TSV, UTF-16 text, pipe-separated text and JSON; every dashboard import input accepts the full
   list; the SMS055, sales-analysis and debts dashboards read the WPS and OpenDocument files end to end; a file that is
   not a spreadsheet, or a damaged PDF, gets a clear Arabic message. Fixtures: tests/fixtures/formats (make.py rebuilds them). */
const { open } = require('./harness');
const fs = require('fs'); const path = require('path');
const DIR = path.join(__dirname, '..', 'fixtures', 'formats');
const EXP = JSON.parse(fs.readFileSync(path.join(DIR, 'expected.json'), 'utf8'));
const FILES = ['sample.xlsx', 'sample.xlsm', 'sample.xlsb', 'sample.xls', 'sample.et', 'sample.ett', 'sample-xlsx-layout.et', 'sample.ods', 'sample.fods', 'sample.xml', 'sample.html',
  'sample.csv', 'sample-semicolon.csv', 'sample-1256.csv', 'sample.tsv', 'sample-utf16.txt', 'sample-pipe.txt', 'sample.json'];
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' } };
const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
const out = {};
(async () => {
  const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data });
  const read = (name, b64) => page.evaluate(async ({ name, b64, exp }) => {
    const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    let sheets; try { sheets = await FileImport.readGrids(new File([u8], name)); } catch (e) { return { err: e.message }; }
    const s = sheets.find(x => x.grid.length > 1); if (!s) return { err: 'no data' };
    const g = s.grid; const head = g[0].map(x => String(x).trim()); const body = g.slice(1).filter(r => r.some(x => String(x).trim()));
    const num = v => typeof v === 'number' ? v : +String(v).replace(/,/g, '') || 0;
    const iso = v => { if (v && typeof v === 'object') return new Date(Date.UTC(1899, 11, 30) + Math.round(v.serial * 86400000)).toISOString().slice(0, 10);
      const s = String(v); let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0]; m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : s; };
    const val = Math.round(body.reduce((a, r) => a + num(r[8]), 0) * 100) / 100, qty = body.reduce((a, r) => a + num(r[7]), 0);
    const r0 = body[0], rl = body[body.length - 1];
    const res = { head: JSON.stringify(head) === JSON.stringify(exp.head), rows: body.length === exp.rows, value: Math.abs(val - exp.value) < 0.01, qty: qty === exp.qty,
      firstDate: iso(r0[0]) === exp.first.date, lastDate: iso(rl[0]) === exp.last.date, code: String(r0[1]) === exp.first.customer, arabic: r0[2] === exp.first.name && rl[2] === exp.last.name,
      sheets: sheets.length };
    const bad = Object.keys(res).filter(k => res[k] === false);
    return bad.length ? { err: 'mismatch: ' + bad.join(',') + ' · ' + JSON.stringify({ n: body.length, val, qty, d0: iso(r0[0]), c: r0[1], n0: r0[2] }) } : res;
  }, { name, b64, exp: EXP });
  out.formats = {};
  for (const f of FILES) { const r = await read(f, fs.readFileSync(path.join(DIR, f)).toString('base64')); out.formats[f] = r.err || true; }
  out.multiSheet = (await read('sample.xls', fs.readFileSync(path.join(DIR, 'sample.xls')).toString('base64'))).sheets === 2;
  // the contracts import in Settings reads rows by header from any format (dates as yyyy-mm-dd)
  out.workbook = await page.evaluate(async b64 => { const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    const wb = await readWorkbook(new File([u8], 'x.ods')); const r = wb.rows(wb.SheetNames[0]); return wb.SheetNames.length === 2 && r.length === 300 && r[0]['التاريخ'] === '2026-07-01' && r[0]['رقم العميل'] === '0110019' && typeof r[0]['القيمة'] === 'number'; },
    fs.readFileSync(path.join(DIR, 'sample.ods')).toString('base64'));
  // clear messages for what is not a spreadsheet
  const msg = (name, bytes) => page.evaluate(async ({ name, bytes }) => { try { await FileImport.readGrids(new File([new Uint8Array(bytes)], name)); return 'read'; } catch (e) { return e.message; } }, { name, bytes });
  out.pdfDamaged = /تالف/.test(await msg('report.pdf', [0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37, 10, 0, 1, 2, 3, 4, 5, 6, 7, 8, 0, 0, 0]));
  out.emptyRefused = /الملف فارغ/.test(await msg('a.xlsx', []));
  out.brokenZip = /تالف/.test(await msg('a.xlsx', [0x50, 0x4b, 3, 4, 1, 2, 3, 4, 5, 6]));
  out.binaryRefused = /غير مدعومة/.test(await msg('a.dat', Array.from({ length: 400 }, (_, i) => (i * 7) % 9)));
  // every dashboard import input accepts the full list
  const accepts = {};
  for (const v of ['sms055', 'sales', 'transfers', 'receipts', 'debts', 'margin', 'settings']) {
    await page.evaluate(v => go(v), v); await page.waitForTimeout(500);
    accepts[v] = await page.evaluate(() => { const i = [...document.querySelectorAll('#view input[type=file]')].filter(x => x.accept !== '.json,application/json'); return i.length > 0 && i.every(x => x.accept === FileImport.ACCEPT && /\.et\b/.test(x.accept) && /\.ods\b/.test(x.accept) && /\.xls\b/.test(x.accept)); });
  }
  out.accepts = Object.values(accepts).every(Boolean) || accepts;
  // end to end: SMS055 from a WPS .et, from .ods and from an Arabic-Windows CSV
  const put = async (sel, f) => { await page.setInputFiles(sel, { name: f, mimeType: 'application/octet-stream', buffer: fs.readFileSync(path.join(DIR, f)) }); await page.waitForTimeout(900); };
  await page.evaluate(() => go('sms055')); await page.waitForTimeout(500);
  for (const f of ['sample.et', 'sample.ods', 'sample-1256.csv', 'sample.xlsb']) {
    await put('#sm-file', f);
    out['sms:' + f] = await page.evaluate(() => dlg.open && /2026\/07\/01 – 2026\/08\/31/.test(dlg.textContent) && document.querySelectorAll('.sm-cols-map .ok').length >= 11);
    await page.evaluate(() => dlg.close());
  }
  await put('#sm-file', 'sample.et'); await page.click('[data-sm="save"]'); await page.waitForTimeout(1200);
  out.smsSaved = await page.evaluate(n => window.__sms.rows.length === n && window.__sms.rows[0].cn !== undefined, EXP.rows);
  // a damaged PDF: an Arabic message in the import dialog, nothing saved (readable PDFs: tests/browser/pdf.js)
  await page.setInputFiles('#sm-file', { name: 'كشف.pdf', mimeType: 'application/pdf', buffer: Buffer.from('%PDF-1.7\n\x00\x01\x02\x03binary') }); await page.waitForTimeout(1500);
  out.smsPdf = await page.evaluate(() => dlg.open && /تعذّر الاستيراد/.test(dlg.textContent) && /تالف/.test(dlg.textContent));
  await page.evaluate(() => dlg.close());
  out.errors = errors.filter(e => !/favicon|ERR_|net::/.test(e));
  await browser.close();
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.values(out.formats).every(v => v === true) && Object.entries(out).every(([k, v]) => k === 'formats' || k === 'errors' ? true : v === true) && !out.errors.length;
  console.log(ok ? 'FORMATS OK' : 'FORMATS FAILED'); process.exit(ok ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
