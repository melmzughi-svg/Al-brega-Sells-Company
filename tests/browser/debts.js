/* Debts dashboard: the template (sheets and headers), reading a filled template back, every KPI with the requested formula
   checked to the dinar on a small portfolio (total receivables, overdue ratio, DSO, collection ratio, credit-limit use,
   customers over the limit, guarantee cover, top 10), the six aging buckets, government debts kept apart from the
   private sector, collection actions and settlements, the customer statement, saving / replacing a period, the monthly
   balances statement still read, a wrong file refused, another department (view only), a large portfolio, English + dark. */
const { open } = require('./harness');
const path = require('path');
const { small, large } = require('./debts-sample');
const SHOTS = process.env.SHOTS || '';
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' }, 'w@test.ly': { role: 'editor', name: 'W', unit: 'w' } };
const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };

/* the old monthly balances statement (BPMC-FIN-TOC-F38): summary sheet + one region sheet with customer accounts */
function ledger() {
  const head = ['اسم الزبون', 'الرصيد في', '2026-07-31', 'المسحوبات', 'التسويات المدينة', 'التحصيلات', 'التسويات الدائنة', 'الرصيد في', '2026-08-31'];
  const S = [[], ['', '', 'ملخص أرصدة حسابات المدينون التجاريون', '', '', '', '', 'رقم الوثيقة', 'BPMC-FIN-TOC-F38'], ['اولا : ديون حسابات بالعملة المحلية'], head, ['', 'مدين', 'دائن', '', '', '', '', 'مدين', 'دائن'],
    ['شركات التوزيع غ', 300000, 0, 80000, 0, -80000, 0, 300000, 0], ['شركات التوزيع ش', 160000, 0, 40000, 0, -60000, 0, 140000, 0], ['اجمالي ديون حسابات بالعملة المحلية', 460000, 0, 120000, 0, -140000, 0, 440000, 0]];
  const W = [[], ['', '', 'حسابات العملة المحلية المنطقة الغربية'], ['ديون حسابات شركات التوزيع " المنطقة الغربية "'],
    ['اسم الزبون', 'رقم الحساب', 'الرصيد في', '2026-07-31', 'المسحوبات ', 'التسويات المدينة', 'التحصيلات', '', 'الرصيد في', '2026-08-31'],
    ['', '', 'مدين', 'دائن', '', '', '', '', 'مدين', 'دائن'],
    ['ليبيا للنفط المشتركة', '160725', 200000, '0.00', 50000, 0, -50000, '', 200000, 0], ['شرارة ذهبية طرابلس', '161500', 100000, '0.00', 30000, 0, -30000, '', 100000, 0]];
  return [['ملخص المحلية والأجنبية', S], ['المنطقة الغربية', W]];
}
const MAKE = `window.__grid=async(sheets)=>{const L=i=>{let s='';i++;while(i){const r=(i-1)%26;s=String.fromCharCode(65+r)+s;i=Math.floor((i-1)/26)}return s};
  const b=buildXlsx(sheets.map(([n,rows])=>[n,rows.map(r=>{const o={};for(let i=0;i<10;i++)o[L(i)]='';r.forEach((v,i)=>o[L(i)]=v);return o})]));
  const u8=new Uint8Array(await b.arrayBuffer());let s='';for(let i=0;i<u8.length;i+=8192)s+=String.fromCharCode.apply(null,u8.subarray(i,i+8192));return btoa(s)};
 window.__book=async(D,opt={})=>{const T=window.__debts.TPL;const sheets=[];
  T.forEach(sp=>{const rows=D[sp.k];if(!rows)return;const name=(opt.names&&opt.names[sp.k])||sp.ar;
    sheets.push([name,rows.map(r=>{const o={};sp.cols.forEach(c=>{const h=(opt.heads&&opt.heads[c[1]])||c[1];o[h]=r[c[0]]==null?'':r[c[0]]});return o})])});
  if(opt.shuffle)sheets.reverse();
  const b=buildXlsx(sheets);const u8=new Uint8Array(await b.arrayBuffer());let s='';for(let i=0;i<u8.length;i+=8192)s+=String.fromCharCode.apply(null,u8.subarray(i,i+8192));return btoa(s)};`;
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
async function put(page, b64, name) { await page.setInputFiles('#db-file', { name, mimeType: XLSX, buffer: Buffer.from(b64, 'base64') }); await page.waitForTimeout(900); }
const book = (page, D, name, opt) => page.evaluate(([D, o]) => __book(D, o), [D, opt || {}]).then(b => put(page, b, name));
const grid = (page, sheets, name) => page.evaluate(s => __grid(s), sheets).then(b => put(page, b, name));
const txt = (page, sel) => page.evaluate(s => (document.querySelector(s) || {}).textContent || '', sel);
const kpis = page => page.evaluate(() => Object.fromEntries([...document.querySelectorAll('#view .db-kpi')].map(k => [k.querySelector('.l').textContent, k.querySelector('.v').textContent + ' | ' + k.textContent])));
const out = {}; process.on("SIGTERM",()=>{console.log("PARTIAL",JSON.stringify(out));process.exit(2)});
(async () => {
  { const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, theme: process.env.THEME || 'light' });
    await page.evaluate(MAKE); await page.evaluate(() => go('debts')); await page.waitForTimeout(600);
    out.emptyState = await page.evaluate(() => !!document.querySelector('.sm-empty [data-dbt="import"]') && !!document.querySelector('#view [data-dbt="template"]'));
    out.powerBiNext = await page.evaluate(() => !!document.querySelector('#view [data-dbt="import"] + .bi-btn[data-bi="imp"]'));
    out.newIcon = await page.evaluate(() => /M8 17.5v-2.5/.test(I.debts));
    // the template: read-me + column guide + 12 data sheets, every requested field present
    const T = await page.evaluate(() => { const s = window.__debts.templateSheets(); return s.map(([n, r]) => [n, Object.keys(r[0] || {})]); });
    const heads = Object.fromEntries(T);
    out.templateSheets = T.length === 14 && ['الفترة', 'الزبائن', 'الضمانات', 'الفواتير', 'التحصيلات', 'الصكوك المرتجعة', 'إجراءات التحصيل', 'الجدولة', 'المخصصات والشطب', 'الإشعارات', 'مقاصة الدعم', 'أعمار الديون'].every(n => heads[n]);
    out.templateFields = ['كود الزبون', 'اسم الزبون', 'الفئة', 'المنطقة', 'المستودع', 'حالة الحساب', 'سقف الائتمان', 'مدة السداد (يوم)', 'رقم العقد', 'تاريخ العقد'].every(h => heads['الزبائن'].includes(h))
      && ['نوع الضمان', 'قيمة الضمان', 'تاريخ الانتهاء'].every(h => heads['الضمانات'].includes(h))
      && ['رقم الفاتورة', 'تاريخ الفاتورة', 'المنتج', 'الكمية', 'قيمة الفاتورة', 'تاريخ الاستحقاق', 'المبلغ المسدد', 'المتبقي'].every(h => heads['الفواتير'].includes(h))
      && ['تاريخ الدفعة', 'مبلغ الدفعة', 'طريقة الدفع', 'الفواتير المسددة'].every(h => heads['التحصيلات'].includes(h))
      && ['غير مستحق بعد', '1 – 30 يومًا', '31 – 60 يومًا', '61 – 90 يومًا', '91 – 180 يومًا', 'أكثر من 180 يومًا'].every(h => heads['أعمار الديون'].includes(h));
    await page.evaluate(() => { const s = S.downloads.save; window.__saved = []; S.downloads.save = a => { window.__saved.push(a.filename); return s(a); }; });
    await page.click('#view [data-dbt="template"]'); await page.waitForTimeout(500);
    out.templateDownload = await page.evaluate(() => window.__saved.includes('Debts-data-template.xlsx'));
    // a filled template
    await book(page, small(), 'بيانات الديون 10-2026.xlsx');
    const dlg = await txt(page, '#dlg');
    out.review = /مراجعة الملف/.test(dlg) && /الفترة 2026\/10\/01 – 2026\/10\/31/.test(dlg) && /الفواتير: 7 صف/.test(dlg) && /الزبائن: 4 صف/.test(dlg) && /645 ألف/.test(dlg);
    out.noFalseWarning = !/يختلف عن مجموع المتبقي/.test(dlg) && !/المسدد \+ المتبقي/.test(dlg);
    out.parsed = await page.evaluate(() => { const P = window.__debts.ST.pending; const c = k => P.cust.find(x => x.a === k);
      return P.T === 'ar' && c('C1').cat === 'station' && c('C2').cat === 'private' && c('C3').cat === 'power' && c('C4').cat === 'gov' && c('C2').st === 'stopped'
        && P.pay.find(p => p.a === 'C3').m === 'off' && P.pay.find(p => p.a === 'C1').m === 'chq' && P.guar[0].gt === 'lg' && P.guar[1].gt === 'cert' && P.bnc[0].rec === false
        && P.act.filter(a => a.t === 'legal').length === 1 && P.prov.find(x => x.t === 'wo').v === 2000 && P.note[0].t === 'cr' && P.off.filter(o => o.st === 'pend').length === 1; });
    await page.click('#dlg [data-dbt="save"]'); await page.waitForTimeout(1500);
    out.savedBatch = await page.evaluate(() => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'debts_report' && l[3].kind === 'batch' && l[3].sheet === 'ar' && l[3].from === '2026-10-01' && l[3].to === '2026-10-31' && l[3].value === 645000 && l[3].parts === 1));
    out.logged = await page.evaluate(() => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'log' && /^استيراد بيانات الديون/.test(l[3].action)));
    out.hero = /^645 ألف$/.test(await txt(page, '.sm-hero-v b')) && /القطاع الخاص 135 ألف/.test(await txt(page, '.db-sub')) && /الجهات الحكومية 510 ألف/.test(await txt(page, '.db-sub'));
    // private sector (default): C1 105,000 + C2 30,000
    let K = await kpis(page);
    out.kpiCount = Object.keys(K).length === 8;
    out.total = /^135 ألف/.test(K['إجمالي الذمم']) && /مجموع أرصدة الزبائن/.test(K['إجمالي الذمم']);
    out.overdue = /^55\.6%/.test(K['نسبة الديون المتأخرة']) && /المتأخر ÷ إجمالي الذمم/.test(K['نسبة الديون المتأخرة']);               // 75,000 / 135,000
    out.dso = /^52 ?يوم/.test(K['متوسط فترة التحصيل (DSO)']) && /\(الذمم ÷ المبيعات الآجلة\) × 31 يوم/.test(K['متوسط فترة التحصيل (DSO)']); // 135,000 / 80,000 × 31
    out.collection = /^21\.1%/.test(K['نسبة التحصيل']) && /محصّل 20 ألف من 95 ألف/.test(K['نسبة التحصيل']);                              // (25,000 − 5,000 returned) / (20,000 + 75,000)
    out.utilisation = /^90%/.test(K['استخدام سقف الائتمان']);                                                                                 // 135,000 / 150,000
    out.overLimit = /^1 /.test(K['الزبائن المتجاوزون للسقف']) && /التجاوز 5,000/.test(K['الزبائن المتجاوزون للسقف']);
    out.cover = /^59\.3%/.test(K['تغطية الضمانات']) && /دون غطاء 55 ألف/.test(K['تغطية الضمانات']);                                        // 80,000 valid / 135,000 (C2's cheque expired)
    out.top10 = /^100%/.test(K['تركّز أكبر 10 مدينين']);
    out.aging = await page.evaluate(() => { const r = [...document.querySelectorAll('#view .db-ageT tbody tr')].map(t => [...t.children].map(x => x.textContent.trim()));
      return r.length === 6 && r[0][1] === 'لم يحل موعد سداده' && r[0][2] === '60,000' && r[3][2] === '45,000' && r[4][2] === '30,000' && r[5][1] === 'مشكوك في تحصيله' && r[5][2] === '0'; });
    out.sectorPanel = await page.evaluate(() => { const s = [...document.querySelectorAll('#view .an-grid > section')].find(x => /القطاع الخاص والجهات الحكومية/.test(x.textContent)); return !!s && s.querySelectorAll('.db-sector > div').length === 2; });
    out.top10Panel = await page.evaluate(() => [...document.querySelectorAll('#view .db-top .r')].map(r => r.dataset.dbAcct).join() === 'C1,C2');
    out.actions = await page.evaluate(() => { const s = [...document.querySelectorAll('#view .an-grid > section')].find(x => /إجراءات التحصيل والتسويات/.test(x.textContent)); if (!s) return false; const t = s.textContent;
      return /محال للشؤون القانونية · 30 ألف/.test(t) && /25%/.test(t) && /متعثر/.test(t) && /15,000|15 ألف/.test(t) && /2,000/.test(t); });
    out.noGovPanelInPrivate = await page.evaluate(() => ![...document.querySelectorAll('#view .an-grid > section')].some(x => /ديون الجهات الحكومية/.test(x.querySelector('h3').textContent)));
    out.reading = await page.evaluate(() => { const t = [...document.querySelectorAll('#view .sm-ins li')].map(l => l.textContent).join('\n'); return /تجاوز سقف الائتمان/.test(t) && /صكوك مرتجعة/.test(t) && /مفصولة عن هذه المؤشرات/.test(t) && /منتهي الصلاحية/.test(t); });
    out.prio = await page.evaluate(() => { const r = [...document.querySelectorAll('#view .an-grid > section')].find(x => /أولويات المتابعة/.test(x.textContent)); return !!r && /متابعة القضية لدى الشؤون القانونية/.test(r.textContent) && /إنذار كتابي وإيقاف البيع الآجل/.test(r.textContent); });
    // government sector
    await page.click('#view [data-db-sec="gov"]'); await page.waitForTimeout(400);
    K = await kpis(page);
    out.govKpis = /^510 ألف/.test(K['إجمالي الذمم']) && /^80\.4%/.test(K['نسبة الديون المتأخرة']) && /^10\.9%/.test(K['نسبة التحصيل']);
    out.govPanel = await page.evaluate(() => { const s = [...document.querySelectorAll('#view .an-grid > section')].find(x => /ديون الجهات الحكومية/.test(x.querySelector('h3').textContent)); if (!s) return false; const t = s.textContent;
      return /الكهرباء/.test(t) && /الوزارات/.test(t) && /900 ألف/.test(t) && /مقاصّة منفذة/.test(t) && /300 ألف/.test(t) && /210 ألف/.test(t); }); // 510k − 300k pending
    await page.click('#view [data-db-sec="all"]'); await page.waitForTimeout(400);
    out.mixWarning = await page.evaluate(() => /تعرض المؤشرات القطاعين معًا/.test(document.getElementById('view').textContent));
    await page.click('#view [data-db-sec="priv"]'); await page.waitForTimeout(300);
    // customer statement
    await page.click('#view .db-top .r[data-db-acct="C1"]'); await page.waitForTimeout(300);
    out.statement = await page.evaluate(() => { const d = document.getElementById('dlg'); const t = d.textContent; return d.open && /محطة النور للوقود \(C1\)/.test(t) && /105,000/.test(t) && /105%/.test(t) && /الفواتير \(2\)/.test(t) && /عدم كفاية الرصيد/.test(t) && /خطاب ضمان مصرفي/.test(t) && /إشعار دائن/.test(t) && d.classList.contains('db-wide'); });
    await page.click('#dlg [data-act="close-dlg"]'); await page.waitForTimeout(200);
    out.wideCleared = await page.evaluate(() => !document.getElementById('dlg').classList.contains('db-wide'));
    // filters and sorting
    await page.selectOption('#db-st', 'stopped'); await page.waitForTimeout(400);
    out.statusFilter = /^30,000|^30 ألف/.test((await kpis(page))['إجمالي الذمم']);
    await page.click('#view [data-dbt="clear"]'); await page.waitForTimeout(300);
    await page.click('#view th[data-db-sort="name"]'); await page.waitForTimeout(300);
    out.sort = await page.evaluate(() => { const s = [...document.querySelectorAll('#view .an-grid > section')].pop(); return [...s.querySelectorAll('tbody tr')].map(r => r.dataset.dbAcct).join() === 'C2,C1'; });
    // export
    const dl = page.waitForEvent('download').then(d => d.suggestedFilename());
    await page.click('#view [data-dbt="export"]'); const ef = await dl; out.exportFile = /^(تحليل-الديون|Debts-analysis)-2026-10-31\.xlsx$/.test(ef) || ef;
    // the same period again, with other sheet names, other headers and sheets in another order → replaces the saved one
    const D2 = small(); D2.inv[0].rem = 50000; D2.inv[0].paid = 10000;
    await book(page, D2, 'بيانات الديون 10-2026 معدلة.xlsx', { shuffle: true, names: { cust: 'بيانات الزبائن', inv: 'Invoices' }, heads: { 'كود الزبون': 'رقم الزبون', 'قيمة الفاتورة': 'القيمة', 'سقف الائتمان': 'الحد الائتماني' } });
    out.aliases = await page.evaluate(() => { const P = window.__debts.ST.pending; return P.inv.length === 7 && P.cust.find(c => c.a === 'C1').lim === 100000; });
    out.overlapWarn = await page.evaluate(() => !!document.getElementById('db-replace') && document.getElementById('db-replace').checked);
    await page.click('#dlg [data-dbt="save"]'); await page.waitForTimeout(1500);
    out.replaced = await page.evaluate(() => window.__debts.ST.batches.length === 1 && /معدلة/.test(window.__debts.ST.batches[0].file) && window.__FBLOG.some(l => l[0] === 'delete' && l[1] === 'debts_report'));
    out.replacedValue = /^125 ألف/.test((await kpis(page))['إجمالي الذمم']);
    // a large realistic month: a second period, trend panel, no errors
    const big = large(); big.per = [{ from: '2026-11-01', to: '2026-11-30' }];
    await book(page, big, 'بيانات الديون 11-2026.xlsx'); await page.click('#dlg [data-dbt="save"]'); await page.waitForTimeout(2000);
    out.largeShown = await page.evaluate(() => document.getElementById('db-at').value === '2026-11-30' && document.querySelectorAll('#db-at option').length === 2 && document.querySelectorAll('#view .db-top .r').length === 10);
    out.trend = await page.evaluate(() => [...document.querySelectorAll('#view .an-grid > section')].some(x => /تطوّر الذمم/.test(x.textContent)) && document.querySelectorAll('#view .sm-col').length === 2);
    if (SHOTS) { await page.selectOption('#db-at', '2026-11-30'); await page.waitForTimeout(400); await page.screenshot({ path: path.join(SHOTS, 'debts-dashboard-' + (process.env.THEME || 'light') + '.png'), fullPage: true }); }
    // the monthly balances statement is still read (balances, no aging)
    await grid(page, ledger(), 'مديونية كاملة 8-2026.xlsx');
    out.ledger = await page.evaluate(() => { const P = window.__debts.ST.pending; return !!P && P.src === 'ledger' && P.to === '2026-08-31' && P.cust.length === 2 && P.cust[0].bal === 200000 && P.cust[0].w === 50000; })
      && /لا يحوي الفواتير/.test(await txt(page, '#dlg'));
    await page.click('#dlg [data-dbt="view-local"]'); await page.waitForTimeout(500);
    out.ledgerView = /^300 ألف$/.test(await txt(page, '.sm-hero-v b')) && await page.evaluate(() => /لا تتوفر أعمار/.test(document.getElementById('view').textContent));
    await page.click('#view [data-dbt="drop-local"]'); await page.waitForTimeout(300);
    // a wrong file is refused
    await grid(page, [['Sheet1', [['التاريخ', 'الكمية', 'المنتج'], ['2026-08-01', 5, 'بنزين']]]], 'other.xlsx');
    out.refused = /تعذّر الاستيراد/.test(await txt(page, '#dlg')) && /قالب بيانات الديون/.test(await txt(page, '#dlg'));
    await page.click('#dlg [data-act="close-dlg"]'); await page.waitForTimeout(200);
    // placement
    out.orgTask = await page.evaluate(() => ORG_TASKS.ca.includes('debts') && VIEWS.some(v => v.id === 'debts' && v.group === 'مهام إدارة الشؤون التجارية والتحصيلات'));
    await page.evaluate(() => go('org')); await page.waitForTimeout(600);
    out.orgCard = await page.evaluate(() => !!document.querySelector('#view [data-org-go="debts"]'));
    await page.evaluate(() => go('home')); await page.waitForTimeout(600);
    out.homeTile = await page.evaluate(() => !!document.querySelector('.hm-app[data-hm="debts"]'));
    out.errors = errors.filter(e => !/favicon|manifest|fonts/.test(e));
    await browser.close(); }
  // another department: cannot save, can view locally
  { const { browser, page } = await open({ user: { email: 'w@test.ly', displayName: 'W' }, data, theme: 'light' });
    await page.evaluate(MAKE); await page.evaluate(() => go('debts')); await page.waitForTimeout(700);
    await book(page, small(), 'debts.xlsx');
    out.otherNoSave = await page.evaluate(() => !document.querySelector('#dlg [data-dbt="save"]') && !!document.querySelector('#dlg [data-dbt="view-local"]'));
    await page.click('#dlg [data-dbt="view-local"]'); await page.waitForTimeout(500);
    out.otherLocal = await page.evaluate(() => !!document.querySelector('#view .banner.info') && document.querySelectorAll('#view .db-kpi').length === 8);
    await browser.close(); }
  // English + dark
  { const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, lang: 'en', theme: 'dark' });
    await page.evaluate(MAKE); await page.evaluate(() => go('debts')); await page.waitForTimeout(700);
    await book(page, large(), 'debts-10-2026.xlsx'); await page.click('#dlg [data-dbt="save"]'); await page.waitForTimeout(1500);
    out.english = await page.evaluate(() => { const t = document.getElementById('view').textContent; return /Debts dashboard/.test(t) && /Days sales outstanding/.test(t) && /Import Excel file/.test(t) && /Debt aging/.test(t) && /Over 180 days/.test(t) && /Credit-limit use/.test(t); });
    out.englishMenu = await page.evaluate(() => { const b = document.querySelector('#o-menu button[data-hm="debts"]'); return !!b && /Debts/.test(b.textContent); });
    out.englishErrors = errors.filter(e => !/favicon|manifest|fonts/.test(e));
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'debts-dashboard-en-dark.png'), fullPage: true });
    await browser.close(); }
  // screenshots: a realistic October in Arabic, light theme
  if (SHOTS) { const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, theme: 'light' });
    await page.evaluate(MAKE); await page.evaluate(() => go('debts')); await page.waitForTimeout(700);
    await book(page, large(), 'بيانات الديون 10-2026.xlsx'); await page.screenshot({ path: path.join(SHOTS, 'debts-import-review-ar.png') });
    await page.click('#dlg [data-dbt="save"]'); await page.waitForTimeout(1500);
    await page.screenshot({ path: path.join(SHOTS, 'debts-dashboard-ar.png'), fullPage: true });
    await page.click('#view [data-db-sec="gov"]'); await page.waitForTimeout(400); await page.screenshot({ path: path.join(SHOTS, 'debts-government-ar.png'), fullPage: true });
    await page.click('#view [data-db-sec="priv"]'); await page.waitForTimeout(300); await page.click('#view .db-top .r'); await page.waitForTimeout(400); await page.screenshot({ path: path.join(SHOTS, 'debts-customer-ar.png') });
    await browser.close(); }
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.values(out).every(v => Array.isArray(v) ? v.length === 0 : v === true);
  console.log(ok ? 'DEBTS OK' : 'DEBTS FAILED'); process.exit(ok ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
