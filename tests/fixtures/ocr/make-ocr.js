/* Builds the scanned samples for tests/browser/ocr.js from the first rows of the shared 300-row sample (sample.json):
     scan-report.jpg  a phone-like photo of the report printed the way Excel prints it: slightly rotated, blurred,
                      grainy and saved as JPEG
     scan-report.pdf  the same picture inside a PDF with no text layer (what a scanner gives)
     scan-expected.json  the rows that are on the page
   Run from the repo root: node tests/fixtures/ocr/make-ocr.js   (needs Playwright) */
const fs = require('fs'); const path = require('path');
const { chromium } = require('playwright');
const HERE = __dirname;
const rows = JSON.parse(fs.readFileSync(path.join(HERE, '..', 'formats', 'sample.json'), 'utf8')).slice(0, 24);
const HEAD = ['التاريخ', 'رقم العميل', 'اسم الزبون', 'رقم المنتج', 'اسم المنتج', 'وحدة القياس', 'الكمية', 'القيمة'];
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
const money = v => Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const show = (k, v) => v == null ? '' : k === 'القيمة' ? money(v) : k === 'الكمية' ? Number(v).toLocaleString('en-US') : String(v);
const html = `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>
 body{margin:28px;font:15px "DejaVu Sans",Arial,sans-serif;color:#111;background:#fff}
 h1{font-size:18px;margin:0 0 4px}p{margin:0 0 12px;font-size:13px}
 table{border-collapse:collapse}th{background:#d9d9d9;padding:5px 9px;border:1px solid #444;white-space:nowrap}
 td{padding:5px 9px;border:1px solid #444;white-space:nowrap}td.n{text-align:left;direction:ltr}</style></head><body>
 <h1>تقرير المبيعات</h1><p>الفترة من 01/07/2026 إلى 31/08/2026</p>
 <table><thead><tr>${HEAD.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => '<tr>' + HEAD.map(k => `<td${/الكمية|القيمة/.test(k) ? ' class="n"' : ''}>${esc(show(k, r[k]))}</td>`).join('') + '</tr>').join('')}</tbody></table></body></html>`;
(async () => {
  const browser = await chromium.launch(); const page = await browser.newPage({ viewport: { width: 1100, height: 1000 }, deviceScaleFactor: 2 });
  await page.setContent(html); const box = await page.locator('body').boundingBox();
  const png = await page.screenshot({ fullPage: true });
  /* rotate, blur, add grain, save as JPEG: a photo of the printed page */
  const jpg = await page.evaluate(async b64 => {
    const img = new Image(); img.src = 'data:image/png;base64,' + b64; await img.decode();
    const c = document.createElement('canvas'); c.width = img.width; c.height = img.height; const x = c.getContext('2d');
    x.fillStyle = '#f4f1ea'; x.fillRect(0, 0, c.width, c.height); x.translate(c.width / 2, c.height / 2); x.rotate(0.6 * Math.PI / 180); x.filter = 'blur(0.8px) contrast(0.9)';
    x.drawImage(img, -img.width / 2, -img.height / 2); x.setTransform(1, 0, 0, 1, 0, 0); x.filter = 'none';
    const d = x.getImageData(0, 0, c.width, c.height); let s = 7; const rnd = () => (s = (s * 16807) % 2147483647) / 2147483647;
    for (let i = 0; i < d.data.length; i += 4) { const n = (rnd() - 0.5) * 38; for (let k = 0; k < 3; k++) d.data[i + k] = Math.max(0, Math.min(255, d.data[i + k] + n)); }
    x.putImageData(d, 0, 0); return c.toDataURL('image/jpeg', 0.72).split(',')[1];
  }, png.toString('base64'));
  fs.writeFileSync(path.join(HERE, 'scan-report.jpg'), Buffer.from(jpg, 'base64'));
  const pg = await browser.newPage();
  await pg.setContent(`<html><body style="margin:0"><img style="display:block;margin:auto;max-width:100%;max-height:200mm" src="data:image/jpeg;base64,${jpg}"></body></html>`);
  await pg.waitForTimeout(200);
  fs.writeFileSync(path.join(HERE, 'scan-report.pdf'), await pg.pdf({ format: 'A4', landscape: true, printBackground: true, margin: { top: '0', bottom: '0', left: '0', right: '0' } }));
  fs.writeFileSync(path.join(HERE, 'scan-expected.json'), JSON.stringify({ head: HEAD, rows: rows.map(r => HEAD.map(k => r[k] == null ? '' : r[k])) }, null, 1));
  await browser.close(); console.log('ok', box);
})();
