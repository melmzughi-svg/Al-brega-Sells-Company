/* Import from scanned files and pictures (ocrimport.js):
   - the text recogniser (vendor/tesseract) is loaded only when a scan or a picture is chosen;
   - a phone-like photo of a ruled Arabic report (tests/fixtures/ocr/scan-report.jpg: rotated, blurred, grainy JPEG)
     opens the review window; no figure (number, code, date) that was read wrongly is left unmarked: every wrong one
     is coloured and has to be checked (a wrong letter in an Arabic name may be unmarked, at most 1 in 10 names), «Approve» stays disabled until every marked cell is checked and the agreement box is ticked, and
     after the reviewer corrects the marked cells the table equals the original rows exactly (numbers as numbers);
   - the same picture inside a PDF with no text layer (scan-report.pdf) opens the same review; cancelling it imports
     nothing and shows no error;
   - end to end: a photo of the sales-analysis report printed the way Excel prints it (no borders) goes through the
     review into the dashboard's own import preview with every column recognised and the right number of rows;
   - every dashboard import input accepts pictures. Fixtures: tests/fixtures/ocr (make-ocr.js rebuilds them). */
const { open } = require('./harness');
const fs = require('fs'); const path = require('path');
const DIR = path.join(__dirname, '..', 'fixtures', 'ocr');
const EXP = JSON.parse(fs.readFileSync(path.join(DIR, 'scan-expected.json'), 'utf8'));
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' } };
const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
const closures = fs.readFileSync(path.join(__dirname, 'closures.js'), 'utf8');
const MAKE = closures.slice(closures.indexOf('const MAKE = `') + 14, closures.indexOf('`;', closures.indexOf('const MAKE = `')));
const out = {};
const T0 = Date.now(); const log = (...a) => process.env.OCR_LOG && console.error(((Date.now() - T0) / 1000).toFixed(1) + 's', ...a);
const shown = (k, v) => typeof v === 'number' ? (/القيمة/.test(k) ? v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : v.toLocaleString('en-US')) : String(v == null ? '' : v);
/* the reviewer: in the review window, every marked cell is compared with the original — corrected when it differs,
   confirmed with ✓ when it is right. Returns how many cells were wrong, and how many of those were not marked. */
const REVIEW = `window.__review=(truth,headKey)=>{const d=document.querySelector('.ocr-review');if(!d)return {err:'no review'};
  const trs=[...d.querySelectorAll('tbody tr:not(.ocr-colrow)')];const hi=typeof headKey==='number'?headKey:trs.findIndex(tr=>[...tr.querySelectorAll('input')].some(i=>i.value.includes(headKey)));
  let wrong=0,unmarked=[],textWrong=[],fixed=0,confirmed=0,text=0;const fig=v=>/\\d/.test(v)&&!/[\\u0600-\\u06FF]/.test(v);const norm=(s,f)=>f?String(s).replace(/\\s+/g,''):String(s).replace(/\\s+/g,' ').trim();
  truth.forEach((row,i)=>{const tr=trs[hi+1+i];if(!tr){wrong+=row.length;unmarked.push('missing row '+i);return}const tds=[...tr.querySelectorAll('td.ocr-c')];
    row.forEach((v,j)=>{const td=tds[j];if(!td){wrong++;unmarked.push('missing cell '+i+','+j);return}const inp=td.querySelector('input');const marked=td.classList.contains('lo')||td.classList.contains('mid');
      const f=fig(v);if(!f)text++;
      if(norm(inp.value,f)!==norm(v,f)){wrong++;if(!marked)(f?unmarked:textWrong).push(i+','+j+': '+v+' → '+inp.value);inp.focus();inp.value=v;inp.dispatchEvent(new Event('input',{bubbles:true}));fixed++}
      else if(marked){td.querySelector('.ocr-ok').click();confirmed++}})});
  /* the header and anything else still marked: checked by hand too */
  d.querySelectorAll('.ocr-ok').forEach(b=>{b.click();confirmed++});
  return {wrong,unmarked,textWrong,textOk:text?Math.round(100*(text-textWrong.length)/text):100,fixed,confirmed,headerRow:hi}}`;
(async () => {
  const { browser, ctx, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data });
  out.lazy = await page.evaluate(() => !window.Tesseract);
  await page.evaluate(REVIEW);
  const start = (name, buf, type) => page.evaluate(({ name, b64, type }) => {
    const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    window.__res = null; FileImport.readGrids(new File([u8], name, { type })).then(s => { window.__res = { sheets: s }; }, e => { window.__res = { err: e.message, cancelled: !!e.cancelled }; });
  }, { name, b64: buf.toString('base64'), type });
  const reviewOpen = () => page.waitForSelector('.ocr-review', { timeout: 180000 }).then(() => true, () => false);
  // 1. the ruled report photographed with a phone
  await start('scan-report.jpg', fs.readFileSync(path.join(DIR, 'scan-report.jpg')), 'image/jpeg');
  out.reviewOpens = await reviewOpen(); log('review 1', out.reviewOpens);
  out.loadedOnDemand = await page.evaluate(() => !!window.Tesseract);
  const truth = EXP.rows.map(r => r.map((v, j) => shown(EXP.head[j], v)));
  out.approveLocked = await page.evaluate(() => { const b = document.querySelector('.ocr-review [data-ocr="ok"]'); return b.disabled && /بقيت \d+ خلية/.test(document.querySelector('.ocr-left').textContent); });
  out.confidenceShown = await page.evaluate(() => [...document.querySelectorAll('.ocr-review td.ocr-c input')].filter(i => i.value).every(i => /الثقة \d+%/.test(i.title)));
  const rv = await page.evaluate(({ truth }) => __review(truth, 'رقم العميل'), { truth });
  out.review = rv;
  out.noWrongFigureUnmarked = !rv.err && rv.unmarked.length === 0 && rv.textOk >= 90;
  out.needsAgreement = await page.evaluate(() => document.querySelector('.ocr-review [data-ocr="ok"]').disabled);
  await page.evaluate(() => { const a = document.querySelector('.ocr-review [data-ocr="agree"]'); a.checked = true; a.dispatchEvent(new Event('change')); });
  out.approveUnlocked = await page.evaluate(() => !document.querySelector('.ocr-review [data-ocr="ok"]').disabled);
  await page.click('.ocr-review [data-ocr="ok"]');
  await page.waitForFunction(() => window.__res, null, { timeout: 30000 });
  out.table = await page.evaluate(({ exp }) => { const r = window.__res; if (r.err) return r.err; const g = r.sheets[0].grid; const hi = g.findIndex(row => row.includes('رقم العميل'));
    const body = g.slice(hi + 1, hi + 1 + exp.rows.length); const bad = [];
    exp.rows.forEach((row, i) => row.forEach((v, j) => { const got = (body[i] || [])[j]; if (got !== v && String(got) !== String(v)) bad.push(i + ',' + j + ' ' + v + ' ≠ ' + got); }));
    const nums = body.every(row => typeof row[6] === 'number' && typeof row[7] === 'number');
    return bad.length ? bad.slice(0, 6).join(' | ') : nums && !!r.sheets[0].ocr ? true : 'numbers not numbers'; }, { exp: EXP });
  log('table', out.table, JSON.stringify(rv).slice(0, 300));
  // 2. the same picture as a PDF without a text layer; cancel imports nothing
  await start('scan-report.pdf', fs.readFileSync(path.join(DIR, 'scan-report.pdf')), 'application/pdf');
  out.pdfReviewOpens = await reviewOpen() && await page.evaluate(() => document.querySelectorAll('.ocr-review tbody tr').length >= 24);
  await page.click('.ocr-review [data-ocr="cancel"]');
  await page.waitForFunction(() => window.__res, null, { timeout: 10000 });
  log('pdf review', out.pdfReviewOpens);
  out.cancel = await page.evaluate(() => window.__res.cancelled === true && !document.querySelector('.ocr-review'));
  // 3. end to end: a photo of the sales-analysis report printed by Excel (no borders), through the sales dashboard
  await page.evaluate(MAKE);
  const rows = await page.evaluate(() => __cxBook('sales', [[2026, 9, 30]], 11, 24));
  const head = Object.keys(rows[0]); const shownHead = head.slice().reverse(); // the review shows the page's columns right to left, as the Arabic window reads
  const photo = await ctx.newPage({ viewport: { width: 1500, height: 900 }, deviceScaleFactor: 2 });
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  await photo.setContent(`<!doctype html><html lang="ar"><head><meta charset="utf-8"><style>body{margin:20px;font:15px "DejaVu Sans",Arial,sans-serif;background:#fff;color:#111}table{border-collapse:collapse}
    th{background:#1f3864;color:#fff;padding:4px 7px;white-space:nowrap}td{padding:4px 7px;white-space:nowrap}td.n{text-align:right}</style></head>
    <body><table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => '<tr>' + head.map(h => `<td${typeof r[h] === 'number' ? ' class="n"' : ''}>${esc(r[h])}</td>`).join('') + '</tr>').join('')}</tbody></table></body></html>`);
  const shot = await photo.screenshot({ fullPage: true, type: 'jpeg', quality: 80 }); await photo.close();
  await page.evaluate(() => go('sales')); await page.waitForTimeout(600);
  await page.setInputFiles('#cx-file-sales', { name: 'Sales_Report_20261105.jpg', mimeType: 'image/jpeg', buffer: shot });
  out.salesReviewOpens = await reviewOpen(); log('sales review', out.salesReviewOpens);
  const truth2 = rows.map(r => shownHead.map(h => String(r[h])));
  const rv2 = await page.evaluate(({ truth }) => __review(truth, 0), { truth: truth2 });
  out.salesReview = { wrong: rv2.wrong, unmarked: rv2.unmarked && rv2.unmarked.slice(0, 8), textWrong: rv2.textWrong, textOk: rv2.textOk, headerRow: rv2.headerRow };
  out.salesNoWrongFigureUnmarked = !rv2.err && rv2.unmarked.length === 0 && rv2.textOk >= 90;
  /* the header row too: whatever the reviewer saw, the column names are set to the report's own */
  await page.evaluate(({ head }) => { const d = document.querySelector('.ocr-review'); const tr = d.querySelector('tbody tr');
    if (tr) [...tr.querySelectorAll('input')].forEach((i, j) => { if (head[j] != null && i.value !== head[j]) { i.value = head[j]; i.dispatchEvent(new Event('input', { bubbles: true })); } });
    d.querySelectorAll('.ocr-ok').forEach(b => b.click()); const a = d.querySelector('[data-ocr="agree"]'); a.checked = true; a.dispatchEvent(new Event('change')); }, { head: shownHead });
  await page.click('.ocr-review [data-ocr="ok"]'); await page.waitForTimeout(1500);
  out.salesDashboard = await page.evaluate(n => { const d = document.getElementById('dlg'); if (!d.open) return 'no dialog';
    const ok = /مراجعة ملف/.test(d.textContent) && [...d.querySelectorAll('.sm-cols-map span')].every(s => s.classList.contains('ok'));
    const lines = (d.querySelector('.sm-mini b') || {}).textContent || ''; return ok && lines.replace(/[^\d]/g, '') === String(n) ? true : d.textContent.slice(0, 300); }, rows.length);
  await page.evaluate(() => dlg.close());
  // 4. a photo of a written contract (running Arabic text with a few numbers) is not a table: a clear message, no review
  const para = ['تم الاتفاق بين الطرف الأول شركة البريقة لتسويق النفط والطرف الثاني على أن يورد الطرف الأول الغاز المسال وفق الكميات المتفق عليها في هذا العقد.',
    'يلتزم الطرف الثاني بسداد قيمة كل شحنة خلال 30 يومًا من تاريخ الاستلام، وفي حال التأخير تطبق غرامة قدرها 2% عن كل شهر.',
    'مدة هذا العقد سنة واحدة تبدأ من 01/01/2026 وتتجدد تلقائيًا ما لم يخطر أحد الطرفين الآخر كتابيًا قبل انتهائها بستين يومًا.',
    'تخضع جميع النزاعات الناشئة عن هذا العقد للقانون الليبي، ويحاول الطرفان حلها وديًا قبل اللجوء إلى المحاكم المختصة.',
    'حرر هذا العقد من نسختين أصليتين بيد كل طرف نسخة للعمل بموجبها، وقد وقع عليه الطرفان بعد قراءته والموافقة على جميع بنوده.'];
  const doc = await ctx.newPage({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 2 });
  await doc.setContent(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>body{margin:50px;font:17px/1.9 "DejaVu Sans",Arial,sans-serif;background:#fff;color:#111}h2{text-align:center}</style></head>
    <body><h2>عقد بيع غاز مسال</h2>${para.map((t, i) => `<p><b>المادة ${i + 1}:</b> ${t}</p>`).join('')}</body></html>`);
  const letter = await doc.screenshot({ fullPage: true, type: 'jpeg', quality: 75 }); await doc.close();
  await start('contract.jpg', letter, 'image/jpeg');
  await page.waitForFunction(() => window.__res, null, { timeout: 180000 }).catch(() => {});
  out.contractNotTable = await page.evaluate(() => { const r = window.__res; return !!r && !r.sheets && /ليس جدول أرقام/.test(r.err) && /العقود/.test(r.err) && !document.querySelector('.ocr-review') ? true : JSON.stringify(r).slice(0, 200); });
  log('contract', out.contractNotTable);
  // 5. every dashboard import input accepts pictures
  const accepts = {};
  for (const v of ['sms055', 'sales', 'transfers', 'receipts', 'debts', 'margin']) {
    await page.evaluate(v => go(v), v); await page.waitForTimeout(500);
    accepts[v] = await page.evaluate(() => { const i = [...document.querySelectorAll('#view input[type=file]')].filter(x => x.accept !== '.json,application/json'); return i.length > 0 && i.every(x => /\.jpg\b/.test(x.accept) && /\.png\b/.test(x.accept)); });
  }
  out.accepts = Object.values(accepts).every(Boolean) || accepts;
  out.errors = errors.filter(e => !/favicon|Failed to load resource/.test(e));
  console.log(JSON.stringify(out, null, 1));
  const info = ['review', 'salesReview'];
  const ok = Object.entries(out).every(([k, v]) => info.includes(k) || (k === 'errors' ? v.length === 0 : v === true));
  console.log(ok ? 'OCR OK' : 'OCR FAILED');
  await browser.close(); process.exit(ok ? 0 : 1);
})();
