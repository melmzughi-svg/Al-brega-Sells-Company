/* Closing-reports dashboards (sales / transfers / receipts): each imports an Excel export shaped like the real report
   (header names as in Sales_Report / Transfers_Report / Receipts_Report), saves it in parts, KPIs match the file,
   unit handling, previous-period comparison, entity detail, export, org-page placement, refusal of a wrong file, English. */
const { open } = require('./harness');
const fs = require('fs'); const path = require('path');
const SHOTS = process.env.SHOTS || '';
const roles = { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' }, 'w@test.ly': { role: 'editor', name: 'W', unit: 'w' } };
const data = { roles, config: { main: { t1: 90, t2: 60, t3: 30 } } };
const MAKE = `(function(){
  const P=[['16291','غاز البترول المسال','اسطوانة','87'],['33500','بنزين ممتاز 95','سالب/لتر','1'],['44347','وقود الديزل','سالب/لتر','1']];
  const WH=[['11','وقود/مستودع الزاوية'],['18','وقود/مستودع طرابلس النفطي'],['30','وقود/مستودع مصراتة']];
  const C=[];for(let i=0;i<30;i++)C.push([String(120000+i*37),'محطة وقود '+['النور','الأمل','الوحدة','السلام','الفجر','البركة'][i%6]+' '+(i+1)]);
  let s=7;const rnd=()=>(s=(s*16807)%2147483647)/2147483647;
  window.__cxBook=(kind,months,seed,drop)=>{s=seed||7;const rows=[];
    months.forEach(([y,m,nd])=>{const mm=String(m).padStart(2,'0');
      if(kind==='sales'){C.forEach((c,ci)=>{if(drop&&ci<drop)return;P.forEach((p,pi)=>{if(rnd()<.4)return;const w=WH[ci%3];
        rows.push({'المنطقة رقم':1,'المنطقة':'المنطقة الغربية والجنوبية','الشهر':y+'-'+mm,'رقم المستودع':w[0],'المستودع':w[1],'رقم الحساب':c[0],'اسم الزبون':c[1],'TRNSFR_THRU':2,'رمز المنتج':p[0],'اسم المنتج':p[1],'PACK_CD':p[3],'الوحدة':p[2],'إجمالي الكمية':Math.round(pi?3000+rnd()*120000:50+rnd()*2000)})})});return}
      for(let d=1;d<=nd;d++){const ds=m+'/'+d+'/'+y+' 0:00';
        if(kind==='transfers'){WH.forEach((w,wi)=>{if(drop&&wi<drop)return;P.forEach((p,pi)=>{if(rnd()<.5)return;
          const ton=pi===2;rows.push({'رقم المستودع':w[0],'اسم المستودع':w[1],'رقم المخزن':String(wi*10+pi+1),'وصف المخزن':'خزان '+(pi+1)+' '+p[1],'رمز المنتج':p[0],'اسم المنتج':p[1],'PRD_IND':1,'PACK_CD':p[3],'وحدة القياس':ton?'طن':p[2],'التاريخ':ds,'نوع الحركة':rnd()<.7?21:35,'وصف الحركة':rnd()<.7?'المبيعات':'الواردات','رقم المستند':1.01083e12+d*1000+wi,'ITEM_NO':1,'الكمية السابقة':0,'الكمية':ton?Math.round((2+rnd()*30)*100)/100:Math.round(pi?2000+rnd()*40000:30+rnd()*800),'PROCESS_COUNTER':0})})});}
        else{for(let k=0;k<4;k++){const sp=['مصفاة الزاوية','شركة ليبيا للنقل','شركة البحر للشحن'][(k+d)%3];if(drop&&(k+d)%3<drop)continue;const p=P[k===0?0:1+(k%2)];const arr=m+'/'+Math.min(nd,d+1+(k%2))+'/'+y+' 0:00';
          const ton=p[0]==='44347';rows.push({'رقم المستند':1.3e13+d*100+k,'التاريخ':ds,'رقم المستودع':WH[k%3][0],'اسم المستودع':WH[k%3][1],'تاريخ المغادرة':ds,'تاريخ الوصول':arr,'رقم الإذن':String(1000+d*10+k),'نوع الاستلام':k%3===0?'محلي':'خارجي','رقم المورد':String(10+(k+d)%3),'اسم المورد':sp,'رمز المنتج':p[0],'المنتج':p[1],'وحدة القياس':ton?'طن':p[2],'الكمية':ton?Math.round((20+rnd()*300)*100)/100:Math.round(k%3?20000+rnd()*400000:100+rnd()*1500)})}}}});
    return rows};
  window.__cxXlsx=async(name,rows)=>{const b=buildXlsx([[name,rows]]);const u8=new Uint8Array(await b.arrayBuffer());let s='';for(let i=0;i<u8.length;i+=8192)s+=String.fromCharCode.apply(null,u8.subarray(i,i+8192));return btoa(s)};
})()`;
async function upload(page, id, kind, months, name, seed, drop) {
  const r = await page.evaluate(async ({ kind, months, seed, drop, name }) => { const rows = __cxBook(kind, months, seed, drop); const q = r => +r['الكمية'] || +r['إجمالي الكمية'] || 0, un = r => r['الوحدة'] || r['وحدة القياس']; const m3 = rows.reduce((a, r) => a + (/لتر/.test(un(r)) ? q(r) / 1000 : /طن/.test(un(r)) ? q(r) * 1000 / (835 - 1.1) : 0), 0); const tons = rows.filter(r => /طن/.test(un(r))).length; return { b64: await __cxXlsx(name.replace(/\.xlsx$/, ''), rows), n: rows.length, m3, tons }; }, { kind, months, seed, drop, name });
  await page.setInputFiles('#cx-file-' + id, { name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(r.b64, 'base64') });
  await page.waitForTimeout(900); return r;
}
const out = {};
const num = s => +String(s).replace(/[^\d.-]/g, '');
(async () => {
  { const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, theme: process.env.THEME || 'light' });
    await page.evaluate(MAKE);
    for (const [id, kind, months, prev, file] of [['sales', 'sales', [[2026, 8, 31], [2026, 9, 30]], [[2026, 7, 31]], 'Sales_Report_20261105.xlsx'], ['transfers', 'transfers', [[2026, 9, 30]], [[2026, 8, 31]], 'Transfers_Report_20261105.xlsx'], ['receipts', 'receipts', [[2026, 9, 30]], [[2026, 8, 31]], 'Receipts_Report_20261105.xlsx']]) {
      const o = out[id] = {};
      await page.evaluate(v => go(v), id); await page.waitForTimeout(600);
      o.emptyState = await page.evaluate(() => !!document.querySelector('.sm-empty [data-cx="import"]'));
      o.powerBiNext = await page.evaluate(() => !!document.querySelector('#view [data-cx="import"] + .bi-btn[data-bi="imp"]'));
      // current period file
      const f1 = await upload(page, id, kind, months, file, 11, id === 'sales' ? 0 : 1); // transfers/receipts: the first warehouse / supplier is missing from the current file
      o.review = await page.evaluate(() => /مراجعة ملف/.test(document.getElementById('dlg').textContent) && [...document.querySelectorAll('#dlg .sm-cols-map span')].every(s => s.classList.contains('ok')));
      o.reviewLines = (await page.evaluate(() => document.querySelector('#dlg .sm-mini b').textContent)).replace(/[^\d]/g, '') === String(f1.n);
      await page.click('#dlg [data-cx="save"]'); await page.waitForTimeout(1500);
      const col = { sales: 'sales_report', transfers: 'transfers_report', receipts: 'receipts_report' }[id];
      o.savedBatch = await page.evaluate(c => window.__FBLOG.some(l => l[0] === 'set' && l[1] === c && l[3].kind === 'batch' && l[3].rows > 0 && l[3].parts >= 1), col);
      o.savedParts = await page.evaluate(c => { const b = window.__FBLOG.find(l => l[0] === 'set' && l[1] === c && l[3].kind === 'batch'); const p = window.__FBLOG.filter(l => l[0] === 'set' && l[1] === c && l[3].kind === 'part' && l[3].batch === b[2]); return p.length === b[3].parts && p.every(x => JSON.parse(x[3].d).R.length > 0); }, col);
      o.logged = await page.evaluate(c => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'log' && /^استيراد /.test(l[3].action) && l[3].col === c), col);
      o.rowsLoaded = await page.evaluate(({ id, n }) => window.__cx.ST[id].rows.length === n, { id, n: f1.n });
      // quantities in cubic metres: litres ÷ 1000 and diesel tonnes ÷ 0.835 share one unit (م³); the hero total equals the file
      o.unitM3 = await page.evaluate(() => /م³/.test(document.querySelector('.sm-hero-v span').textContent));
      o.heroTotal = await page.evaluate(() => { const t = document.querySelector('.sm-hero-v b').textContent; return t; });
      o.heroMatches = Math.abs(num(o.heroTotal) * (/مليون/.test(o.heroTotal) ? 1e6 : /ألف/.test(o.heroTotal) ? 1e3 : 1) - f1.m3) / f1.m3 < 0.06;
      o.unitNote = await page.evaluate(tons => { const n = document.querySelector('#view .sm-note'); return !!n && /لتر ÷ 1000/.test(n.textContent) && (!tons || /طن × 1000 ÷ \(835\.0 − 1\.1\)/.test(n.textContent)); }, f1.tons);
      o.rawKept = await page.evaluate(id => window.__cx.ST[id].rows.every(r => r.u !== 'م³') && window.__cx.data(id).some(r => r.u === 'م³' && r.u0 && r.q0 != null), id);
      o.kpis = await page.evaluate(() => document.querySelectorAll('#view .an-kpis .kpi').length >= 5);
      o.chart = await page.evaluate(() => document.querySelectorAll('#view .sm-col').length > 0);
      o.reading = await page.evaluate(() => document.querySelectorAll('#view .sm-ins li').length >= 3);
      o.products = await page.evaluate(() => document.querySelectorAll('#view .sm-pc').length === 2); // the two litre products
      o.groups = await page.evaluate(() => document.querySelectorAll('#view .an-grid > section').length >= 6);
      o.table = await page.evaluate(() => document.querySelectorAll('#view tr[data-cx-ent]').length > 0);
      o.noPrev = await page.evaluate(() => !document.querySelector('.sm-delta'));
      // switch the unit to cylinders: one product, hero changes
      const hasUnit = await page.evaluate(() => !!document.getElementById('cx-unit'));
      if (hasUnit) { await page.selectOption('#cx-unit', { index: 0 }); await page.waitForTimeout(400); const u0 = await page.evaluate(() => document.getElementById('cx-unit').value); if (!/م³/.test(u0)) { o.unitSwitch = await page.evaluate(() => document.querySelectorAll('#view .sm-pc').length === 1 && /اسطوانة/.test(document.querySelector('.sm-hero-v span').textContent)); } else { await page.selectOption('#cx-unit', { index: 1 }); await page.waitForTimeout(400); o.unitSwitch = await page.evaluate(() => document.querySelectorAll('#view .sm-pc').length === 1); } }
      else o.unitSwitch = true;
      await page.click('#view [data-cx="clear"]'); await page.waitForTimeout(400);
      // entity detail
      await page.click('#view tr[data-cx-ent]'); await page.waitForTimeout(300);
      o.detail = await page.evaluate(() => document.getElementById('dlg').open && /حسب المنتج/.test(document.getElementById('dlg').textContent));
      await page.click('#dlg [data-act="close-dlg"]'); await page.waitForTimeout(200);
      // export
      const dl = page.waitForEvent('download').then(d => d.suggestedFilename());
      await page.click('#view [data-cx="export"]'); const fnm = await dl; o.exportFile = new RegExp('^' + kind.replace(/^./, c => c.toUpperCase()) + '_Report-analysis-.*\\.xlsx$').test(fnm) || fnm;
      // previous-period file (fewer parties) → delta and "stopped" list appear
      await upload(page, id, kind, prev, file.replace('20261105', '20261005'), 5, id === 'sales' ? 3 : 0);
      o.noOverlap = await page.evaluate(() => !document.getElementById('cx-replace'));
      await page.click('#dlg [data-cx="save"]'); await page.waitForTimeout(1500);
      await page.click('#view [data-cx-pre="' + (id === 'sales' ? 'month' : 'month') + '"]'); await page.waitForTimeout(500);
      o.delta = await page.evaluate(() => !!document.querySelector('.sm-delta'));
      o.stopped = await page.evaluate(() => { const s = [...document.querySelectorAll('#view .an-grid > section')].find(x => /توقّف|بلا حركة|لم يورّدوا/.test(x.textContent)); return !!s && s.querySelectorAll('.an-row').length > 0; });
      o.files = await page.evaluate(id => window.__cx.ST[id].batches.length === 2, id);
      // overlapping re-import replaces
      await upload(page, id, kind, months, file.replace('.xlsx', '-v2.xlsx'), 9);
      o.overlapWarn = await page.evaluate(() => !!document.getElementById('cx-replace') && document.getElementById('cx-replace').checked);
      await page.click('#dlg [data-cx="save"]'); await page.waitForTimeout(1500);
      o.replaced = await page.evaluate(({ id, c }) => window.__cx.ST[id].batches.length === 2 && window.__cx.ST[id].batches.some(b => /-v2\.xlsx$/.test(b.file)) && window.__FBLOG.some(l => l[0] === 'delete' && l[1] === c), { id, c: col });
      if (SHOTS) { await page.click('#view [data-cx-pre="all"]'); await page.waitForTimeout(500); await page.screenshot({ path: path.join(SHOTS, id + '-dashboard-' + (process.env.THEME || 'light') + '.png'), fullPage: true }); }
      // a wrong file (another report) is refused with a clear message
      const other = kind === 'sales' ? 'receipts' : 'sales';
      await upload(page, id, other, [[2026, 9, 30]], 'other.xlsx', 3);
      o.refused = await page.evaluate(() => /تعذّر الاستيراد/.test(document.getElementById('dlg').textContent) && /لم نجد أعمدة/.test(document.getElementById('dlg').textContent));
      await page.click('#dlg [data-act="close-dlg"]'); await page.waitForTimeout(200);
    }
    // densities: the settings panel is for the administrator only; a config/units override is applied live
    await page.evaluate(() => go('settings')); await page.waitForTimeout(500);
    out.noPanelForEditor = await page.evaluate(() => !document.getElementById('fu-panel'));
    // placement: org page card, duties shortcut, home tiles, app menu
    await page.evaluate(() => go('org')); await page.waitForTimeout(600);
    out.orgTask = await page.evaluate(() => ['sales', 'transfers', 'receipts'].every(id => ORG_TASKS.ca.includes(id)));
    out.views = await page.evaluate(() => ['sales', 'transfers', 'receipts'].every(id => VIEWS.some(v => v.id === id && v.group === 'مهام إدارة الشؤون التجارية والتحصيلات')));
    out.orgCard = await page.evaluate(() => { const v = document.getElementById('view'); return ['sales', 'transfers', 'receipts'].every(id => !!v.querySelector('[data-org-go="' + id + '"]')); });
    await page.evaluate(() => go('home')); await page.waitForTimeout(600);
    out.homeTiles = await page.evaluate(() => ['sales', 'transfers', 'receipts'].every(id => !!document.querySelector('.hm-app[data-hm="' + id + '"]')));
    out.errors = errors.filter(e => !/favicon|manifest|fonts/.test(e));
    await browser.close(); }
  // another department: cannot save, can view locally
  { const { browser, page } = await open({ user: { email: 'w@test.ly', displayName: 'W' }, data, theme: 'light' });
    await page.evaluate(MAKE); await page.evaluate(() => go('transfers')); await page.waitForTimeout(700);
    await upload(page, 'transfers', 'transfers', [[2026, 9, 30]], 'Transfers_Report.xlsx', 4);
    out.otherNoSave = await page.evaluate(() => !document.querySelector('#dlg [data-cx="save"]') && !!document.querySelector('#dlg [data-cx="view-local"]'));
    await page.click('#dlg [data-cx="view-local"]'); await page.waitForTimeout(500);
    out.otherLocal = await page.evaluate(() => !!document.querySelector('#view .banner.info') && document.querySelectorAll('#view .sm-pc').length > 0);
    await browser.close(); }
  // administrator: densities panel, override changes the conversion
  { const { browser, page } = await open({ user: { email: 'admin@test.ly', displayName: 'A' }, data: { roles: Object.assign({ 'admin@test.ly': { role: 'admin', name: 'Admin' } }, roles), config: { main: { t1: 90, t2: 60, t3: 30 } } }, theme: 'light' });
    await page.evaluate(MAKE); await page.evaluate(() => go('transfers')); await page.waitForTimeout(700);
    const f = await upload(page, 'transfers', 'transfers', [[2026, 9, 30]], 'Transfers_Report.xlsx', 4); await page.click('#dlg [data-cx="save"]'); await page.waitForTimeout(1200);
    const before = await page.evaluate(() => window.__cx.data('transfers').filter(r => r.p === '44347').reduce((a, r) => a + r.q, 0));
    await page.evaluate(() => go('settings')); await page.waitForTimeout(500);
    out.adminPanel = await page.evaluate(() => !!document.getElementById('fu-panel') && !!document.querySelector('#fu-panel [data-fu-d="44347"]') && document.querySelector('#fu-panel [data-fu-d="44347"]').value === '835');
    await page.fill('#fu-panel [data-fu-d="44347"]', '850'); await page.click('#fu-panel [data-fu="save"]'); await page.waitForTimeout(800);
    out.densitySaved = await page.evaluate(() => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'config' && l[2] === 'units' && l[3].densities['44347'] === 850));
    await page.evaluate(() => { FuelUnits.state.over['44347'] = 850; FuelUnits.state.ver++; go('transfers'); }); await page.waitForTimeout(500);
    const after = await page.evaluate(() => window.__cx.data('transfers').filter(r => r.p === '44347').reduce((a, r) => a + r.q, 0));
    out.densityApplied = Math.abs(after / before - (835 - 1.1) / (850 - 1.1)) < 0.001 && f.tons > 0;
    // temperature: general setting and per-type setting saved; litres corrected to 15 °C, tonnes and LPG not
    await page.evaluate(() => go('settings')); await page.waitForTimeout(500);
    await page.fill('#fu-temp', '30'); await page.fill('#fu-panel [data-fu-t="@diesel"]', '35'); await page.click('#fu-panel [data-fu="save"]'); await page.waitForTimeout(800);
    out.tempSaved = await page.evaluate(() => window.__FBLOG.some(l => l[0] === 'set' && l[1] === 'config' && l[2] === 'units' && l[3].temp === 30 && l[3].temps['@diesel'] === 35));
    out.tempBad = await page.evaluate(async () => { const i = document.getElementById('fu-temp'); i.value = '200'; document.querySelector('#fu-panel [data-fu="save"]').click(); await new Promise(r => setTimeout(r, 300)); const m = /−18/.test(document.getElementById('fu-panel').textContent); i.value = '30'; return m; });
    out.vcf = await page.evaluate(() => { const F = window.FuelUnits; F.state.over = {}; F.state.temp = 30; F.state.temps = { '@diesel': 35 };
      const r = (q, u, c, n, t) => F.toM3(q, u, c, n, t);
      const ok = [F.vcf(835, 30) === 0.98716, F.vcf(740, 30) === 0.98151, F.vcf(835, 15) === 1,
        r(1000, 'لتر', '44347', 'بيع وقود الديزل', 30).q === 0.98716, r(1000, 'لتر', '44347', 'بيع وقود الديزل').vcf === 0.98285,
        r(1000, 'لتر', '33500', 'بيع البنزين').t === 30, r(1000, 'لتر', '16291', 'بيع منتج غاز مسال', 30).q === 1,
        Math.abs(r(1, 'طن', '44347', 'بيع وقود الديزل', 30).q - 1000 / (835 - 1.1)) < 1e-9, r(5, 'لتر', '', 'صيانة', 30) === null];
      const rows = F.normalize([{ p: '44347', pn: 'ديزل', q: 2000, u: 'لتر', tc: 30 }]); const note = F.explain(rows);
      F.state.temp = 15; F.state.temps = {}; return ok.every(Boolean) && Math.abs(rows[0].q - 1.97432) < 1e-6 && /0\.98716/.test(note) && /54B/.test(note) && /30°م/.test(note); });
    await browser.close(); }
  // English + dark
  { const { browser, page } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data, lang: 'en', theme: 'dark' });
    await page.evaluate(MAKE); await page.evaluate(() => go('receipts')); await page.waitForTimeout(700);
    await upload(page, 'receipts', 'receipts', [[2026, 9, 30]], 'Receipts_Report.xlsx', 9); await page.click('#dlg [data-cx="save"]'); await page.waitForTimeout(1200);
    out.english = await page.evaluate(() => /Receipts dashboard/.test(document.getElementById('view').textContent) && /Average transit/.test(document.getElementById('view').textContent) && /Import receipts file/.test(document.getElementById('view').textContent));
    out.englishMenu = await page.evaluate(() => { const b = document.querySelector('#o-menu button[data-hm="receipts"]'); return !!b && /Receipts/.test(b.textContent); });
    if (SHOTS) await page.screenshot({ path: path.join(SHOTS, 'receipts-dashboard-en-dark.png'), fullPage: true });
    await page.evaluate(() => go('sales')); await page.waitForTimeout(500);
    out.englishSales = await page.evaluate(() => /Sales analysis dashboard/.test(document.getElementById('view').textContent));
    await browser.close(); }
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.entries(out).every(([k, v]) => Array.isArray(v) ? v.length === 0 : v && typeof v === 'object' ? Object.entries(v).every(([kk, vv]) => kk === 'heroTotal' ? true : vv === true) : v === true);
  console.log(ok ? 'CLOSURES OK' : 'CLOSURES FAILED'); process.exit(ok ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
