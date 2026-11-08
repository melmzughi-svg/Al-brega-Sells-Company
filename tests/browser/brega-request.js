/* Brega's own data-request workbooks (profit and receivables) are read as they are: their sheet and column names,
   the grey sample row skipped, density in kg/litre, percentages stored as fractions, road transport and storage on the
   receipts sheet, write-offs / provisions / subsidy offsets on one notes sheet, returned cheques flagged on the
   collections sheet, and the aging date taken from the instructions sheet. */
const { open } = require('./harness');
(async () => {
  const { browser, page, errors } = await open({ user: { email: 'ca@test.ly', displayName: 'CA' }, data: { roles: { 'ca@test.ly': { role: 'editor', name: 'CA', unit: 'ca' } } } });
  const out = await page.evaluate(async () => {
    const o = {}; const file = (sheets, name) => new File([buildXlsx(sheets)], name);
    const S = '(مثال) بنزين';
    const profit = [
      ['المنتجات', [{ 'م': '', 'المنتج': S, 'وحدة الأساس (لتر قياسي / طن / م3)': 'لتر قياسي', 'الكثافة عند 15°م (كغم/لتر)': 0.74 }, { 'م': 1, 'المنتج': 'بنزين', 'وحدة الأساس (لتر قياسي / طن / م3)': 'لتر قياسي', 'الكثافة عند 15°م (كغم/لتر)': 0.74, 'رمز المنتج': '33500', 'الفئة': 'وقود' }]],
      ['هيكل الأسعار', [{ 'المنتج': 'بنزين', 'تاريخ السريان': '2026-01-01', 'الوحدة': 'لتر', 'تكلفة المنتج (مصفاة / استيراد)': 0.6, 'النقل': 0.02, 'هامش توزيع البريقة': 0.012, 'هامش المحطة': 0.01, 'رسوم للدولة': 0, 'السعر الرسمي المعتمد': 0.15 }]],
      ['أسعار البيع', [{ 'المنتج': 'بنزين', 'فئة الزبون': 'محطات وقود', 'سعر البيع للوحدة': 0.15, 'تاريخ السريان': '2026-01-01', 'نسبة الخصم أو العمولة': 0.05 }]],
      ['الاستلامات والتكلفة', [{ 'تاريخ الاستلام': '2026-01-05', 'المنتج': 'بنزين', 'المستودع': 'مستودع طرابلس', 'رقم الشحنة / الفاتورة': 'INV-1', 'الكمية المستلمة (وحدة الأساس)': 1000000, 'قيمة الشراء': 600000, 'رسوم الموانئ والجمارك': 1000, 'النقل البري': 15000, 'التخزين والمناولة': 5000, 'تكاليف مباشرة أخرى': 500 }]],
      ['حركة الكميات', [{ 'الشهر': '2026-01', 'المنتج': S, 'المستودع': 'x', 'رصيد أول المدة': 1 }, { 'الشهر': '2026-01', 'المنتج': 'بنزين', 'المستودع': 'مستودع طرابلس', 'رصيد أول المدة': 200000, 'الوارد': 1000000, 'تحويلات واردة': 0, 'تحويلات صادرة': 50000, 'المبيعات': 900000, 'الرصيد الفعلي (الجرد)': 248000, 'نسبة الفاقد المسموح بها': 0.002 }]],
      ['المبيعات', [{ 'الشهر': '2026-01', 'المنتج': 'بنزين', 'المستودع': 'مستودع طرابلس', 'فئة الزبون': 'محطات وقود', 'الكمية المباعة (وحدة الأساس)': 900000, 'قيمة المبيعات قبل الخصم': 135000, 'الخصومات والعمولات': 0.5 }]],
      ['الدعم', [{ 'الشهر': '2026-01', 'المنتج': 'بنزين', 'الدعم للوحدة': 0.49, 'الدعم المستحق': 441000, 'المحصَّل فعلًا': 300000 }]],
      ['المصروفات التشغيلية', [{ 'الشهر': '2026-01', 'المستودع / الإدارة': 'مستودع طرابلس', 'بند المصروف': 'رواتب', 'المبلغ': 120000 }]]];
    const pf = file(profit, 'Brega_Profit_Data_Request.xlsx'); const g = await FileImport.readGrids(pf);
    o.profitRead = await window.__pmx.tryImport(g, pf); const P = window.__pmx.state.pending; const sh = P ? P.sheets : {};
    o.profitSheets = Object.keys(sh).sort().join(',') === 'dep,opex,prods,recv,sales,segs,stock,struct,sub';
    o.sampleSkipped = (sh.prods || []).length === 1 && (sh.stock || []).length === 1;
    o.density = sh.prods && sh.prods[0].d15 === 740 && sh.prods[0].p === '33500';
    o.fractionPct = sh.segs && sh.segs[0].disc === 5 && sh.stock && Math.abs(sh.stock[0].al - 0.2) < 1e-9;
    o.salesDiscAmount = sh.sales && sh.sales[0].disc === 0.5;
    o.priceStruct = sh.struct && sh.struct[0].price === 150 && sh.struct[0].fees === 0;
    o.receipts = sh.recv && sh.recv[0].m === '2026-01' && sh.recv[0].q === 1000 && sh.recv[0].port === 1000 && sh.recv[0].cus === 500;
    o.depFromReceipts = sh.dep && sh.dep[0].road === 15000 && sh.dep[0].stor === 5000 && sh.dep[0].w === 'مستودع طرابلس';
    o.stockCols = sh.stock && sh.stock[0].rc === 1000 && sh.stock[0].sd === 900 && sh.stock[0].cl === 248;
    o.subOpex = sh.sub && sh.sub[0].due === 441000 && sh.sub[0].col === 300000 && sh.opex && sh.opex[0].k === 'رواتب' && sh.opex[0].w === 'مستودع طرابلس';
    try { document.getElementById('dlg').close(); } catch (e) {} window.__pmx.state.pending = null;
    const recv = [
      ['التعليمات', [{ 'البند': 'تاريخ قياس أعمار الديون', 'القيمة': '2026-10-31' }]],
      ['الزبائن', [{ 'كود الزبون': 'مثال-001', 'اسم الزبون': '(مثال) شركة ...', 'الفئة': 'شركات خاصة', 'حالة الحساب': 'نشط', 'سقف الائتمان': 1, 'مدة السداد (يوم)': 30 }, { 'كود الزبون': 'C1', 'اسم الزبون': 'شركة الأفق', 'الفئة': 'شركات خاصة', 'حالة الحساب': 'نشط', 'سقف الائتمان': 500000, 'مدة السداد (يوم)': 30 }]],
      ['الفواتير الآجلة', [{ 'رقم الفاتورة': 'F1', 'تاريخ الفاتورة': '2026-09-01', 'كود الزبون': 'C1', 'قيمة الفاتورة': 15000, 'تاريخ الاستحقاق': '2026-10-01', 'المسدد والتسويات': 6000, 'المتبقي': 9000 }]],
      ['التحصيلات', [{ 'تاريخ الدفعة': '2026-09-10', 'كود الزبون': 'C1', 'رقم الفاتورة المسددة': 'F1', 'المبلغ': 6000, 'طريقة الدفع': 'صك', 'رقم الصك / الإيصال': 'CH1', 'صك مرتجع؟': 'لا' }, { 'تاريخ الدفعة': '2026-10-02', 'كود الزبون': 'C1', 'رقم الفاتورة المسددة': 'F1', 'المبلغ': 4000, 'طريقة الدفع': 'صك', 'رقم الصك / الإيصال': 'CH2', 'صك مرتجع؟': 'نعم' }]],
      ['الإشعارات والتسويات', [{ 'التاريخ': '2026-09-20', 'كود الزبون': 'C1', 'رقم الفاتورة': 'F1', 'النوع': 'إشعار دائن', 'المبلغ': 500, 'السبب': 'فرق كمية' }, { 'التاريخ': '2026-10-30', 'كود الزبون': 'C1', 'رقم الفاتورة': 'F1', 'النوع': 'شطب دين معدوم', 'المبلغ': 700 },
        { 'التاريخ': '2026-10-30', 'كود الزبون': 'C1', 'النوع': 'مخصص ديون مشكوك فيها', 'المبلغ': 900 }, { 'التاريخ': '2026-10-30', 'كود الزبون': 'C1', 'النوع': 'مقاصّة', 'المبلغ': 2000, 'الدعم المستحق للشركة': 5000, 'حالة المقاصة': 'قيد الإجراء' }]]];
    const D = await window.__debts.parseFile(file(recv, 'Brega_Receivables_Data_Request.xlsx'));
    o.agingDate = D.to === '2026-10-31' && D.from === '2026-10-01';
    o.sampleCustomer = D.cust.length === 1 && D.cust[0].a === 'C1';
    o.bounced = D.pay.length === 2 && D.bnc.length === 1 && D.bnc[0].no === 'CH2' && D.bnc[0].v === 4000;
    o.notesRouted = D.note.length === 1 && D.note[0].t === 'cr' && D.prov.length === 2 && D.prov.some(x => x.t === 'wo' && x.v === 700) && D.prov.some(x => x.t === 'pr' && x.v === 900)
      && D.off.length === 1 && D.off[0].st === 'pend' && D.off[0].sub === 5000;
    o.templateClean = (() => { const t = window.__debts.templateSheets(); const n = t.find(x => x[0] === 'الإشعارات'); const p = t.find(x => x[0] === 'التحصيلات'); return !('حالة المقاصة' in n[1][0]) && !('صك مرتجع؟' in p[1][0]); })();
    // the monthly statement: a group that is only on the summary sheet (ministry-funded sector) is added from it
    const head = ['اسم الزبون', 'الرصيد في', '2026-07-31', 'المسحوبات', 'التسويات المدينة', 'التحصيلات', 'التسويات الدائنة', 'الرصيد في', '2026-08-31'];
    const SUM = [[], ['', '', 'ملخص أرصدة حسابات المدينون التجاريون', '', '', '', '', 'رقم الوثيقة', 'BPMC-FIN-TOC-F38'], ['اولا : ديون حسابات بالعملة المحلية'], ['كما هي في دفاترنا الرئيسية ليوم', '2026-08-31'], head, ['', 'مدين', 'دائن', '', '', '', '', 'مدين', 'دائن'],
      [' شركة الكهرباء غ', 1000000, 0, 50000, 0, 0, 0, 1050000, 0], [' شركة الكهرباء ش', 500000, 0, 0, 0, 0, 0, 500000, 0], [' محطات التحلية غ', 100000, 0, 0, 0, 0, 0, 100000, 0], [' الممولة من وزارة المالية', 1600000, 0, 50000, 0, 0, 0, 1650000, 0],
      ['شركات التوزيع غ', 300000, 0, 80000, 0, -80000, 0, 300000, 0], ['شركات التوزيع ش', 0, 0, 10, 0, -10, 0, 0, 0], [' شركات توزيع ', 300000, 0, 80010, 0, -80010, 0, 300000, 0], ['اجمالي ديون حسابات بالعملة المحلية', 1900000, 0, 130010, 0, -80010, 0, 1950000, 0]];
    const WR = [[], ['', '', 'حسابات العملة المحلية المنطقة الغربية'], ['ديون حسابات شركات التوزيع " المنطقة الغربية "'],
      ['اسم الزبون', 'رقم الحساب', 'الرصيد في', '2026-07-31', 'المسحوبات ', 'التسويات المدينة', 'التحصيلات', '', 'الرصيد في', '2026-08-31'], ['', '', 'مدين', 'دائن', '', '', '', '', 'مدين', 'دائن'],
      ['ليبيا للنفط المشتركة', '160725', 200000, '0.00', 50000, 0, -50000, '', 200000, 0], ['شرارة ذهبية طرابلس', '161500', 100000, '0.00', 30000, 0, -30000, '', 100000, 0]];
    const col = i => String.fromCharCode(65 + i);
    const asObj = rows => rows.map(r => { const x = {}; for (let i = 0; i < 10; i++) x[col(i)] = ''; r.forEach((v, i) => x[col(i)] = v); return x; });
    const LD = await window.__debts.parseFile(file([['ملخص المحلية والأجنبية', asObj(SUM)], ['المنطقة الغربية', asObj(WR)]], 'مديونية كاملة 8-2026.xlsx'));
    const tot = LD.cust.reduce((a, c) => a + c.bal, 0);
    o.ledgerSummaryOnly = LD.src === 'ledger' && LD.cust.length === 5 && Math.abs(tot - 1950000) < 0.01 && LD.cust.some(c => c.a === 'شركة الكهرباء غ') && LD.cust.find(c => /التحلية/.test(c.n)).cat === 'gov' && LD.cust.filter(c => /الكهرباء/.test(c.n)).every(c => c.cat === 'power') && LD.warn.some(w => /أُخذت من الخلاصة/.test(w[1]));
    // profit margin before Brega sends prices: an «awaiting prices» page, never a loss equal to the cost
    { const wb = [['الاستلامات والتكلفة', [{ 'تاريخ الاستلام': '2026-08-05', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'الكمية المستلمة': 2000000, 'قيمة الشراء': 1000000 }]],
        ['حركة الكميات', [{ 'الشهر': '2026-08', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'رصيد أول المدة': 0, 'الوارد': 2000000, 'المبيعات': 1800000, 'الرصيد الفعلي (الجرد)': 199000 }]],
        ['المبيعات', [{ 'الشهر': '2026-08', 'المنتج': 'ديزل', 'المستودع': 'مستودع طرابلس', 'فئة الزبون': 'محطات وقود', 'الكمية المباعة': 1800000, 'رمز الزبون': '120170' }]]];
      const f = file(wb, 'Brega_Profit.xlsx'); await window.__pmx.tryImport(await FileImport.readGrids(f), f);
      document.querySelector('#dlg [data-px="view-local"]').click(); go('margin'); await new Promise(r => setTimeout(r, 900));
      const v = document.getElementById('view').innerText;
      o.awaitingPrices = /بانتظار أسعار المنتجات من شركة البريقة/.test(v) && /أسعار الفئات/.test(v) && /تكلفة البضاعة لكل م³/.test(v) && !/-1\.\d+ مليون|NaN/.test(v); }
    o.ledgerNoCollRatio = window.__debts.analyze(LD, { sec: 'all' }).collR === null;
    return o;
  });
  out.noErrors = errors.length === 0; if (errors.length) console.log(errors);
  console.log(JSON.stringify(out, null, 1));
  const ok = Object.values(out).every(v => v === true);
  console.log(ok ? 'BREGA-REQUEST OK' : 'BREGA-REQUEST FAILED'); await browser.close(); process.exit(ok ? 0 : 1);
})();
