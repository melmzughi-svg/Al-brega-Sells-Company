/* Profit margin, full model on Brega's workbook: the template downloads, a ten-sheet workbook imports through the margin
   page, quantities are converted to m³ at 15 °C (litres with the temperature factor, tonnes by density), and every KPI
   matches a hand calculation: net revenue = qty × price − discounts + subsidy, COGS = qty × weighted average landed cost,
   gross profit after losses and depot costs, loss ratio, official vs actual margin, subsidy due vs collected, net profit
   after operating expenses, top customers with overdue. Also: gross-only label without expenses, view without saving,
   a viewer cannot save, English + dark, no XSS from the file. */
const { open } = require('./harness');
const path = require('path');
const SHOTS = process.env.SHOTS || '';
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' }, 'v@test.ly': { role: 'viewer', name: 'V', unit: 'ca' } };
const BOOK = `(function(){
  window.near=(a,b,e)=>Math.abs(a-b)<=(e||0.01);
  const sheets=(opex,evil)=>{const S=[];
    S.push(['المنتجات',[{'رمز المنتج':'44347','اسم المنتج':evil?'<img src=x onerror="window.__xss=1">ديزل':'ديزل','الفئة':'وقود','وحدة القياس':'لتر','الكثافة عند 15°م':835},
      {'رمز المنتج':'33500','اسم المنتج':'بنزين','الفئة':'وقود','وحدة القياس':'لتر','الكثافة عند 15°م':740},{'رمز المنتج':'16291','اسم المنتج':'غاز الطهي (LPG)','الفئة':'غاز','وحدة القياس':'طن','الكثافة عند 15°م':540}]]);
    S.push(['هيكل السعر',[{'رمز المنتج':'44347','تاريخ السريان':'2026-01-01','وحدة السعر':'لتر','تكلفة المنتج':0.10,'النقل':0.01,'هامش توزيع البريقة':0.02,'هامش المحطة':0.015,'الرسوم':0.005,'السعر النهائي':0.15}]]);
    S.push(['أسعار الفئات',[{'رمز المنتج':'44347','فئة الزبون':'محطة وقود','تاريخ السريان':'2026-01-01','وحدة السعر':'لتر','السعر المحلي':0.15,'سعر الأجانب':0.1,'العملة':'USD','سعر الصرف':5,'نسبة الخصم %':0},
      {'رمز المنتج':'44347','فئة الزبون':'جهة حكومية','تاريخ السريان':'2026-01-01','وحدة السعر':'لتر','السعر المحلي':0.16}]]);
    S.push(['الاستلامات والتكلفة',[{'الشهر':'2026-09','المستودع':'طرابلس','المنطقة':'الغربية','رمز المنتج':'44347','المصدر':'المؤسسة الوطنية للنفط','رقم الشحنة':'SH-1','الكمية':100,'الوحدة':'م³','قيمة الشراء':9000,'الشحن البحري':500,'التأمين':100,'رسوم الموانئ':200,'الجمارك':200},
      {'الشهر':'2026-09','المستودع':'بنغازي','المنطقة':'الشرقية','رمز المنتج':'33500','الكمية':10000,'الوحدة':'لتر','درجة الحرارة':30,'قيمة الشراء':5000},
      {'الشهر':'2026-09','المستودع':'بنغازي','المنطقة':'الشرقية','رمز المنتج':'16291','الكمية':54,'الوحدة':'طن','درجة الحرارة':35,'قيمة الشراء':3000}]]);
    S.push(['تكاليف المستودعات',[{'الشهر':'2026-09','المستودع':'طرابلس','المنطقة':'الغربية','النقل البري':300,'التخزين والمناولة':150}]]);
    S.push(['الكميات والمخزون',[{'الشهر':'2026-09','المستودع':'طرابلس','المنطقة':'الغربية','رمز المنتج':'44347','الوحدة':'م³','مخزون أول المدة':0,'الكمية المستلمة':100,'الكمية المباعة':90,'تحويلات واردة':0,'تحويلات صادرة':0,'مخزون آخر المدة':9,'الفاقد المسموح %':0.5}]]);
    S.push(['المبيعات حسب الزبون',[{'الشهر':'2026-09','المستودع':'طرابلس','المنطقة':'الغربية','رمز الزبون':'C1','اسم الزبون':'محطة النصر','فئة الزبون':'محطات','رمز المنتج':'44347','الكمية':60000,'الوحدة':'لتر'},
      {'الشهر':'2026-09','المستودع':'طرابلس','المنطقة':'الغربية','رمز الزبون':'C2','اسم الزبون':evil?'<script>window.__xss=1</script>بلدية':'بلدية طرابلس','فئة الزبون':'جهة حكومية','رمز المنتج':'44347','الكمية':30,'الوحدة':'م³','قيمة البيع':4950,'الخصومات':50}]]);
    S.push(['الدعم الحكومي',[{'الشهر':'2026-09','رمز المنتج':'44347','وحدة الدعم':'لتر','الدعم للوحدة':0.05,'المستحق':4500,'المحصل':3000,'تاريخ التحصيل':'2026-10-10'}]]);
    if(opex)S.push(['المصروفات التشغيلية',[{'الشهر':'2026-09','البند':'رواتب','المبلغ':600},{'الشهر':'2026-09','البند':'كهرباء','المبلغ':300}]]);
    S.push(['الذمم',[{'الشهر':'2026-09','رمز الزبون':'C2','اسم الزبون':'بلدية طرابلس','المبيعات الآجلة':5000,'المتأخر':2000,'أيام التأخر':75}]]);
    return S};
  window.__pxBook=async(opex,evil)=>{const b=buildXlsx(sheets(opex,evil));const u8=new Uint8Array(await b.arrayBuffer());let s='';for(let i=0;i<u8.length;i+=8192)s+=String.fromCharCode.apply(null,u8.subarray(i,i+8192));return btoa(s)};
})()`;
async function put(page, name, b64) { await page.setInputFiles('#pm-file', { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(b64, 'base64') }); await page.waitForTimeout(900); }
const out = {};
const near = (a, b, e) => Math.abs(a - b) <= (e || 0.01);
(async () => {
  const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
  let saved = null;
  { const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, theme: process.env.THEME || 'light' });
    await page.evaluate(BOOK);
    await page.evaluate(() => go('margin')); await page.waitForTimeout(800);
    out.tabs = await page.evaluate(() => !!document.querySelector('[data-pm-tab="full"][aria-selected="true"]') && !!document.querySelector('[data-pm-tab="prices"]'));
    out.onboarding = await page.evaluate(() => !!document.querySelector('.px-empty [data-px="template"]') && document.querySelectorAll('.px-cov > div').length === 10);
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'margin-full-empty.png'), fullPage: true });
    const dl = page.waitForEvent('download', { timeout: 5000 }).then(d => d.suggestedFilename()).catch(() => '');
    await page.click('.px-empty [data-px="template"]'); out.template = (await dl) === 'Brega-profit-margin-template.xlsx';
    // the workbook without operating expenses: gross only
    await put(page, 'Brega-margin-2026-09.xlsx', await page.evaluate(() => __pxBook(false)));
    out.review = await page.evaluate(() => /مراجعة ملف بيانات هامش الربح/.test(document.getElementById('dlg').textContent) && /9/.test(document.querySelector('#dlg .sm-mini b').textContent));
    out.conversions = await page.evaluate(() => { const P = window.__pmx.state.pending, r = P.sheets.recv; const g = r.find(x => x.p === '33500'), l = r.find(x => x.p === '16291');
      return near(g.q, 9.8151, 1e-6) && g.vf === 0.98151 && near(l.q, 54000 / (540 - 1.1), 1e-6) && !l.vf && near(P.sheets.sales[0].q, 60, 1e-9) && P.sheets.sales[0].seg === 'station' && P.sheets.sales[1].seg === 'gov' && near(P.sheets.struct[0].price, 150, 1e-9) && near(P.sheets.sub[0].rate, 50, 1e-9); });
    await page.click('#dlg [data-px="save"]'); await page.waitForTimeout(1200);
    out.saved = await page.evaluate(() => { const L = window.__FBLOG.filter(l => l[0] === 'set' && l[1] === 'margin'); return L.some(l => l[2] === 'econ_ref') && L.some(l => l[2] === 'econ_2026-09' && l[3].sp === 1 && !l[3].sales) && L.some(l => l[2] === 'econs_2026-09_0' && l[3].sales.length === 2)
      && window.__FBLOG.some(l => l[1] === 'log' && l[3] && /استيراد بيانات هامش الربح/.test(l[3].action)); });
    saved = await page.evaluate(() => { const o = {}; window.__FBLOG.filter(l => l[0] === 'set' && l[1] === 'margin').forEach(l => { o[l[2]] = l[3]; }); return o; });
    out.grossOnly = await page.evaluate(() => /ربح إجمالي فقط/.test(document.querySelector('.sm-hero').textContent) && /الربح الإجمالي فقط/.test(document.getElementById('view').textContent));
    out.math = await page.evaluate(() => { const X = window.__pmx, B = X.build(), A = X.sum(B.lines.filter(x => x.m === '2026-09' && x.p === '44347'));
      const c = id => B.lines.filter(x => x.c === id).reduce((a, x) => a + x.gross - x.disc + x.sub - x.cogs - x.loss - x.dist, 0);
      return { gross: near(A.gross, 13950), disc: near(A.disc, 50), sub: near(A.sub, 4500), nr: near(A.nr, 18400), cogs: near(A.cogs, 9000), loss: near(A.loss, 100), dist: near(A.dist, 450), gp: near(A.gp, 8850),
        m: near(A.m, 8850 * 100 / 18400, 1e-6), pm3: near(A.pm3, 98.3333, 1e-3), off: near(A.offM3, 20, 1e-9), c1: near(c('C1'), 5633.333, 1e-2), c2: near(c('C2'), 3216.667, 1e-2) }; });
    out.kpis = await page.evaluate(() => { const t = document.querySelector('.kpis').textContent; return /18,400|18٬400|18 ألف/.test(t) && /48\.1%/.test(t) && /1%/.test(t) && /67%|66\.7%/.test(t) && /20 \/ 98/.test(t); });
    out.views = await page.evaluate(() => { const v = document.getElementById('view'); return !!v.querySelector('.px-wf') && !!v.querySelector('.px-ps') && !!v.querySelector('.px-line, .px-two') && /ترتيب المنتجات حسب الربحية/.test(v.textContent)
      && /مقارنة المستودعات/.test(v.textContent) && /أكثر 10 زبائن ربحًا وأقلهم/.test(v.textContent) && /متأخر/.test(v.querySelector('.px-two').textContent) && /المعادلات/.test(v.textContent); });
    out.lossOver = await page.evaluate(() => /تجاوز/.test(document.getElementById('view').textContent));
    out.subsidyRead = await page.evaluate(() => /الدعم المستحق/.test(document.querySelector('.sm-ins').textContent));
    // group by region, sort by margin, filter a category
    await page.click('[data-px-grp="rg"]'); await page.waitForTimeout(300);
    out.regions = await page.evaluate(() => /مقارنة المناطق/.test(document.getElementById('view').textContent) && /الغربية/.test(document.getElementById('view').textContent));
    await page.selectOption('#px-seg', 'gov'); await page.waitForTimeout(300);
    out.segFilter = await page.evaluate(() => { const X = window.__pmx, B = X.build(); const A = X.sum(X.fl(B, '2026-09', '2026-09')); return near(A.gross, 4950) && near(A.gp, 3216.667, 0.01); });
    await page.click('[data-px="clear"]'); await page.waitForTimeout(300);
    // re-import with operating expenses: net profit
    await put(page, 'Brega-margin-2026-09.xlsx', await page.evaluate(() => __pxBook(true)));
    await page.click('#dlg [data-px="save"]'); await page.waitForTimeout(1200);
    out.net = await page.evaluate(() => { const X = window.__pmx, B = X.build(), A = X.sum(X.fl(B, '2026-01', '2026-12')); return near(A.opex, 900) && near(A.np, 7950) && /ربح صافٍ بعد المصروفات/.test(document.querySelector('.sm-hero').textContent); });
    out.exportBtn = await page.evaluate(() => !!document.querySelector('#view [data-px="export"]'));
    const dl2 = page.waitForEvent('download', { timeout: 5000 }).then(d => d.suggestedFilename()).catch(() => '');
    await page.click('#view [data-px="export"]'); out.exported = /^Profit-margin-full-2026-09_2026-09\.xlsx$/.test(await dl2);
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'margin-full.png'), fullPage: true });
    // the quick tab still works
    await page.click('[data-pm-tab="prices"]'); await page.waitForTimeout(400);
    out.quickTab = await page.evaluate(() => !!document.querySelector('[data-pm-tab="prices"][aria-selected="true"]') && !!document.querySelector('#pm-board, .sm-empty'));
    await browser.close(); }
  // a viewer sees the saved data, cannot save; XSS from the file is inert
  { const { browser, page } = await open({ user: { email: 'v@test.ly', displayName: 'V' }, data: Object.assign({}, data, { margin: saved }) });
    await page.evaluate(BOOK); await page.evaluate(() => go('margin')); await page.waitForTimeout(900);
    out.viewerSees = await page.evaluate(() => !!document.querySelector('.px-wf') && /طرابلس/.test(document.getElementById('view').textContent));
    await put(page, 'evil.xlsx', await page.evaluate(() => __pxBook(true, true)));
    out.viewerNoSave = await page.evaluate(() => !document.querySelector('#dlg [data-px="save"]') && !!document.querySelector('#dlg [data-px="view-local"]'));
    await page.click('#dlg [data-px="view-local"]'); await page.waitForTimeout(600);
    out.localView = await page.evaluate(() => /عرض مؤقت/.test(document.getElementById('view').textContent));
    out.noXss = await page.evaluate(() => !window.__xss && !document.querySelector('#view img[src="x"]') && /&lt;|<img/.test(document.getElementById('view').innerHTML) === true);
    await browser.close(); }
  // English + dark
  { const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data: Object.assign({}, data, { margin: saved }), lang: 'en', theme: 'dark' });
    await page.evaluate(() => go('margin')); await page.waitForTimeout(900);
    out.english = await page.evaluate(() => /Products ranked by profitability/.test(document.getElementById('view').textContent) && /Net profit after operating expenses|Gross profit only/.test(document.querySelector('.sm-hero').textContent));
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'margin-full-en-dark.png'), fullPage: true });
    await browser.close(); }
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.values(out).every(v => v && typeof v === 'object' ? Object.values(v).every(x => x === true) : v === true);
  console.log(ok ? 'MARGIN-ECON OK' : 'MARGIN-ECON FAILED'); process.exit(ok ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
