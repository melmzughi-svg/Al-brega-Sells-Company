/* Profit margin dashboard: Sales_Report import from the margin page (saved with the sales-analysis data), quantities in m³,
   prices edited on the board with an effective date (history, log), two price lists (local LYD / foreign currency × rate),
   profit and margin match a hand calculation, price list import, simulator and calculator, foreign customer toggle,
   permissions (another department cannot read; a viewer cannot edit), English + dark. */
const { open } = require('./harness');
const path = require('path');
const SHOTS = process.env.SHOTS || '';
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' }, 'v@test.ly': { role: 'viewer', name: 'V', unit: 'ca' }, 'w@test.ly': { role: 'editor', name: 'W', unit: 'w' } };
const MAKE = `(function(){
  const P=[['16291','غاز البترول المسال','اسطوانة'],['33500','بنزين ممتاز 95','لتر'],['44347','وقود الديزل','لتر']];
  const C=[];for(let i=0;i<20;i++)C.push([String(120000+i*37),'محطة وقود '+(i+1)]);
  let s=7;const rnd=()=>(s=(s*16807)%2147483647)/2147483647;
  window.__pmBook=(months,seed)=>{s=seed||7;const rows=[];months.forEach(([y,m])=>{const mm=String(m).padStart(2,'0');
    C.forEach((c,ci)=>P.forEach((p,pi)=>{if(rnd()<.3)return;rows.push({'المنطقة رقم':1,'المنطقة':'المنطقة الغربية','الشهر':y+'-'+mm,'رقم المستودع':String(11+ci%2),'المستودع':ci%2?'مستودع طرابلس':'مستودع الزاوية','رقم الحساب':c[0],'اسم الزبون':c[1],'TRNSFR_THRU':2,'رمز المنتج':p[0],'اسم المنتج':p[1],'PACK_CD':1,'الوحدة':p[2],'إجمالي الكمية':Math.round(pi?5000+rnd()*90000:20+rnd()*500)})}))});return rows};
  window.__pmXlsx=async(name,sheet,rows)=>{const b=buildXlsx([[sheet,rows]]);const u8=new Uint8Array(await b.arrayBuffer());let s='';for(let i=0;i<u8.length;i+=8192)s+=String.fromCharCode.apply(null,u8.subarray(i,i+8192));return btoa(s)};
})()`;
async function put(page, name, b64) { await page.setInputFiles('#pm-file', { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(b64, 'base64') }); await page.waitForTimeout(900); }
const out = {};
/* these checks cover the quick mode (prices and the sales report); the full Brega model is tests/browser/margin-econ.js */
const quick = page => page.evaluate(() => { const b = document.querySelector('[data-pm-tab="prices"]'); if (b) b.click(); });
const near = (a, b) => Math.abs(a - b) <= Math.max(0.02, Math.abs(b) * 1e-6);
(async () => {
  const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
  { const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, theme: process.env.THEME || 'light' });
    await page.evaluate(MAKE);
    out.inApps = await page.evaluate(() => !!window.__appIcon && VIEWS.some(v => v.id === 'margin') && ORG_TASKS.ca.includes('margin'));
    // the company's product types: densities per type, kerosene kept, marine fuel added, service types never converted
    out.productTypes = await page.evaluate(() => { const F = window.FuelUnits, t = (n, u) => F.toM3(10, u || 'طن', '', n);
      const ok = (n, d) => { const x = t(n); return !!x && x.d === d; };
      return ok('بيع البنزين', 740) && ok('بيع وقود الديزل', 835) && ok('بيع منتج غاز مسال', 540) && ok('بيع منتج الزيت الثقيل', 970) && ok('بيع وقود السفن', 890) && ok('كيروسين', 800)
        && ['تقديم خدمات', 'صيانة', 'نقل وتشغيل', 'توريد منتجات نفطية'].every(n => t(n, 'لتر') === null) && t('وقود الديزل', 'لتر').q === 0.01; });
    await page.evaluate(() => go('margin')); await page.waitForTimeout(700); await quick(page);
    out.emptyState = await page.evaluate(() => !!document.querySelector('.sm-empty [data-pm="import"]'));
    out.powerBiImport = await page.evaluate(() => !!document.querySelector('#view [data-pm="import"] + .bi-btn[data-bi="imp"]'));
    // Sales_Report through the margin page
    const f = await page.evaluate(async () => { const rows = __pmBook([[2026, 8], [2026, 9]], 11); return { b64: await __pmXlsx('Sales_Report', 'Sales_Report', rows), n: rows.length }; });
    await put(page, 'Sales_Report_20261105.xlsx', f.b64);
    out.salesReview = await page.evaluate(() => /مراجعة ملف تحليل المبيعات/.test(document.getElementById('dlg').textContent));
    await page.click('#dlg [data-pm="save-sales"]'); await page.waitForTimeout(1500);
    out.savedToSales = await page.evaluate(n => window.__cx.ST.sales.rows.length === n && window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'sales_report' && l[3].kind === 'batch'), f.n);
    out.unpricedBanner = await page.evaluate(() => /بلا تكلفة أو سعر/.test(document.querySelector('#view .banner.warn').textContent));
    out.m3Units = await page.evaluate(() => /م³/.test(document.querySelector('.pm-units').textContent) && /اسطوانة/.test(document.querySelector('.pm-units').textContent));
    // edit prices on the board (diesel + gasoline per m³, LPG per cylinder), effective from Aug 1st
    await page.click('#view [data-pm="edit"]'); await page.waitForTimeout(400);
    const set = async (p, k, v) => page.fill(`[data-pm-in="${p}"][data-k="${k}"]`, String(v));
    await set('44347', 'cost', 120); await set('44347', 'loc', 150); await set('44347', 'fx', 40);
    await set('33500', 'cost', 130); await set('33500', 'loc', 150); await set('33500', 'fx', 45);
    await set('16291', 'cost', 8); await set('16291', 'loc', 10); await set('16291', 'fx', 3);
    await page.fill('#pm-eff', '2026-08-01'); await page.fill('#pm-rate', '4.8'); await page.fill('#pm-target', '15'); await page.fill('#pm-min', '5'); await page.fill('#pm-why', 'قرار رقم 12');
    await page.click('#view [data-pm="save-prices"]'); await page.waitForTimeout(900);
    out.savedPrices = await page.evaluate(() => { const l = window.__FBLOG.filter(x => x[0] === 'set' && x[1] === 'margin' && x[2] === 'prices').pop(); return !!l && l[3].products.length === 3 && l[3].rates[0].rate === 4.8 && l[3].updatedBy === 'ca@test.ly' && l[3].products.every(p => p.h[0].from === '2026-08-01' && p.h[0].why === 'قرار رقم 12'); });
    out.priceLog = await page.evaluate(() => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'log' && l[3].col === 'margin' && /وقود الديزل/.test(l[3].summary) && /150/.test(l[3].summary)));
    out.noBanner = await page.evaluate(() => !document.querySelector('#view .banner.warn'));
    // hand calculation, all local: diesel/gasoline litres → m³
    const calc = await page.evaluate(() => { const P = { '44347': [120, 150], '33500': [130, 150], '16291': [8, 10] }; let rev = 0, cogs = 0;
      window.__cx.ST.sales.rows.forEach(r => { const q = r.u === 'لتر' ? r.q / 1000 : r.q; rev += q * P[r.p][1]; cogs += q * P[r.p][0]; }); return { rev, cogs, gp: rev - cogs }; });
    const A = await page.evaluate(() => { const a = window.__pm.agg(window.__pm.rowsAll()); return { rev: a.rev, cogs: a.cogs, gp: a.gp, m: a.m }; });
    out.profitMatches = near(A.gp, calc.gp) && near(A.rev, calc.rev);
    out.marginMatches = near(A.m, calc.gp * 100 / calc.rev);
    out.kpis = await page.evaluate(() => document.querySelectorAll('#view .an-kpis .kpi').length === 8);
    out.reading = await page.evaluate(() => document.querySelectorAll('#view .sm-ins li').length >= 4);
    out.chart = await page.evaluate(() => document.querySelectorAll('#view .sm-col').length === 2);
    out.powerBiExport = await page.evaluate(() => !!document.querySelector('#view [data-pm="export"] + .bi-btn[data-bi="exp"]'));
    // foreign customer: priced at fx × rate
    const c0 = await page.evaluate(() => document.querySelector('#view [data-pm-fx]').dataset.pmFx);
    await page.click(`#view [data-pm-fx="${c0}"]`); await page.waitForTimeout(800);
    out.foreignSaved = await page.evaluate(c => { const l = window.__FBLOG.filter(x => x[0] === 'set' && x[1] === 'margin').pop(); return l[3].foreign.includes(c); }, c0);
    const calc2 = await page.evaluate(c => { const P = { '44347': [120, 150, 40], '33500': [130, 150, 45], '16291': [8, 10, 3] }; let gp = 0, fx = 0;
      window.__cx.ST.sales.rows.forEach(r => { const q = r.u === 'لتر' ? r.q / 1000 : r.q; const f = r.c === c; const pr = f ? P[r.p][2] * 4.8 : P[r.p][1]; gp += q * (pr - P[r.p][0]); if (f) fx += q * P[r.p][2]; }); return { gp, fx }; }, c0);
    const A2 = await page.evaluate(() => { const a = window.__pm.agg(window.__pm.rowsAll()); return { gp: a.gp, fx: a.fxRev }; });
    out.foreignMath = near(A2.gp, calc2.gp) && near(A2.fx, calc2.fx) && calc2.fx > 0;
    // a later price (Sep 1st) applies to September only
    await page.click('#view [data-pm="edit"]'); await page.waitForTimeout(300);
    await set('44347', 'loc', 180); await page.fill('#pm-eff', '2026-09-01'); await page.click('#view [data-pm="save-prices"]'); await page.waitForTimeout(800);
    out.effective = await page.evaluate(() => { const r = window.__pm.rowsAll(); const a = r.find(x => x.p === '44347' && x.d.startsWith('2026-08')), b = r.find(x => x.p === '44347' && x.d.startsWith('2026-09')); return +window.__pm.priceAt('44347', a.d).loc === 150 && +window.__pm.priceAt('44347', b.d).loc === 180; });
    await page.click('#view [data-pm-hist="44347"]'); await page.waitForTimeout(300);
    out.history = await page.evaluate(() => document.querySelectorAll('#dlg tbody tr').length === 2); await page.click('#dlg [data-act="close-dlg"]');
    // simulator + calculator
    await page.fill('#pm-sim-chg', '10'); await page.waitForTimeout(200);
    out.simulator = await page.evaluate(() => document.querySelectorAll('#pm-sim-out > div').length === 8 && /\+/.test(document.getElementById('pm-sim-out').textContent));
    await page.fill('#pm-c-cost', '100'); await page.fill('#pm-c-m', '20'); await page.fill('#pm-c-price', '125'); await page.waitForTimeout(200);
    out.calculator = await page.evaluate(() => { const t = document.getElementById('pm-calc-out').textContent; return /125/.test(t) && /25%/.test(t) && /20%/.test(t); });
    // price list import
    const pl = await page.evaluate(async () => __pmXlsx('prices', 'قائمة الأسعار', [{ 'رمز المنتج': '16291', 'اسم المنتج': 'غاز البترول المسال', 'تكلفة الوحدة': 8.5, 'السعر المحلي': 11, 'سعر الأجانب': 3.2, 'تاريخ السريان': '2026-10-01' }]));
    await put(page, 'prices.xlsx', pl);
    out.priceReview = await page.evaluate(() => /مراجعة قائمة الأسعار/.test(document.getElementById('dlg').textContent));
    await page.click('#dlg [data-pm="save-pl"]'); await page.waitForTimeout(800);
    out.priceImported = await page.evaluate(() => { const l = window.__FBLOG.filter(x => x[0] === 'set' && x[1] === 'margin').pop(); const p = l[3].products.find(x => x.p === '16291'); return p.h.length === 2 && p.h[1].from === '2026-10-01' && p.h[1].loc === 11; });
    // export
    const dl = page.waitForEvent('download').then(d => d.suggestedFilename());
    await page.click('#view [data-pm="export"]'); out.exportFile = /^Profit-margin-.*\.xlsx$/.test(await dl);
    out.orgCard = await page.evaluate(() => /هامش الربح/.test(recPanel('ca')));
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'margin-dashboard-' + (process.env.THEME || 'light') + '.png'), fullPage: true });
    out.errors = errors.filter(e => !/favicon|manifest|fonts/.test(e));
    await browser.close(); }
  // a viewer of the department reads but cannot edit; another department cannot read
  const saved = { products: [{ p: '44347', pn: 'وقود الديزل', u: 'م³', h: [{ from: '2026-01-01', cost: 120, loc: 150, fx: 40 }] }], rates: [{ from: '2026-01-01', rate: 4.8 }], cur: 'USD', target: 15, min: 5, foreign: [] };
  { const { browser, page } = await open({ user: { email: 'v@test.ly', displayName: 'V' }, data: Object.assign({}, data, { margin: { prices: saved } }) });
    await page.evaluate(() => go('margin')); await page.waitForTimeout(700); await quick(page);
    out.viewerNoEdit = await page.evaluate(() => !document.querySelector('#view [data-pm="edit"]') && /150/.test(document.getElementById('pm-board').textContent));
    await browser.close(); }
  { const { browser, page } = await open({ user: { email: 'w@test.ly', displayName: 'W' }, data, fail: { margin: { n: 5, code: 'permission-denied' } } });
    await page.evaluate(() => go('margin')); await page.waitForTimeout(700); await quick(page);
    out.otherDenied = await page.evaluate(() => /هذه اللوحة لإدارة الشؤون التجارية/.test(document.getElementById('view').textContent));
    // no read permission: a profit file can still be imported and analysed on this device, without a save button
    out.deniedLocal = await page.evaluate(async () => {
      const wb = [['الاستلامات والتكلفة', [{ 'تاريخ الاستلام': '2026-08-05', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'الكمية المستلمة': 2000000, 'قيمة الشراء': 1000000 }]],
        ['حركة الكميات', [{ 'الشهر': '2026-08', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'رصيد أول المدة': 0, 'الوارد': 2000000, 'المبيعات': 1800000, 'الرصيد الفعلي (الجرد)': 199000 }]],
        ['المبيعات', [{ 'الشهر': '2026-08', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'فئة الزبون': 'محطات وقود', 'الكمية المباعة': 1800000 }]]];
      const hasImport = !!document.querySelector('#view [data-pm="import"]');
      const f = new File([buildXlsx(wb)], 'Brega_Profit.xlsx'); await window.__pmx.tryImport(await FileImport.readGrids(f), f);
      const noSave = !document.querySelector('#dlg [data-px="save"]'); document.querySelector('#dlg [data-px="view-local"]').click();
      await new Promise(r => setTimeout(r, 700)); const v = document.getElementById('view').innerText;
      return hasImport && noSave && /بانتظار أسعار المنتجات/.test(v) && /عرض مؤقت/.test(v); });
    await browser.close(); }
  // administrator refused because «admin via Google» is on and the session used a password: the page says so
  { const { browser, page } = await open({ user: { email: 'a@test.ly', displayName: 'A' }, data: Object.assign({}, data, { roles: Object.assign({}, data.roles, { 'a@test.ly': { role: 'admin', name: 'A' } }), config: { security: { adminGoogle: true } } }), fail: { margin: { n: 5, code: 'permission-denied' } } });
    await page.evaluate(() => go('margin')); await page.waitForTimeout(900);
    out.adminPasswordReason = await page.evaluate(() => /دخلت بكلمة المرور/.test((document.querySelector('.px-why') || {}).textContent || ''));
    await browser.close(); }
  // stored XSS: payloads in the saved prices and in the imported names render as text
  { const X = t => `"'><img src=x data-p="${t}" onerror="window.__xss=(window.__xss||0)+1">`;
    const bad = { products: [{ p: '44347', pn: X('pn'), u: X('u'), h: [{ from: '2026-01-01', cost: 120, loc: 150, fx: 40, by: X('by'), why: X('why') }] }], rates: [{ from: '2026-01-01', rate: 4.8 }], cur: X('cur'), target: 15, min: 5, foreign: [] };
    const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data: Object.assign({}, data, { margin: { prices: bad } }) });
    await page.evaluate(MAKE); await page.evaluate(() => go('margin')); await page.waitForTimeout(600); await quick(page);
    const f = await page.evaluate(async x => { const rows = __pmBook([[2026, 9]], 3).map((r, i) => Object.assign(r, { 'اسم الزبون': x + i, 'المستودع': x, 'اسم المنتج': i % 2 ? x : r['اسم المنتج'] })); return __pmXlsx('Sales_Report', 'Sales_Report', rows); }, X('row'));
    await put(page, 'Sales_Report.xlsx', f); await page.click('#dlg [data-pm="view-local"]'); await page.waitForTimeout(600);
    await page.click('#view [data-pm-hist="44347"]').catch(() => {}); await page.waitForTimeout(300); await page.click('#dlg [data-act="close-dlg"]').catch(() => {});
    await page.click('#view tr[data-pm-cust]').catch(() => {}); await page.waitForTimeout(300);
    out.noXss = await page.evaluate(() => !window.__xss && !document.querySelector('img[src="x"]'));
    await browser.close(); }
  // English + dark
  { const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data: Object.assign({}, data, { margin: { prices: saved } }), lang: 'en', theme: 'dark' });
    await page.evaluate(MAKE); await page.evaluate(() => go('margin')); await page.waitForTimeout(700); await quick(page);
    const f = await page.evaluate(async () => __pmXlsx('Sales_Report', 'Sales_Report', __pmBook([[2026, 9]], 5)));
    await put(page, 'Sales_Report.xlsx', f); await page.click('#dlg [data-pm="save-sales"]'); await page.waitForTimeout(1200);
    out.english = await page.evaluate(() => { const t = document.getElementById('view').textContent; return /Profit margin dashboard/.test(t) && /Gross margin/.test(t) && /Pricing calculator/.test(t); });
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'margin-dashboard-en-dark.png'), fullPage: true });
    await browser.close(); }
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.entries(out).every(([k, v]) => Array.isArray(v) ? v.length === 0 : v === true);
  console.log(ok ? 'MARGIN OK' : 'MARGIN FAILED'); process.exit(ok ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
