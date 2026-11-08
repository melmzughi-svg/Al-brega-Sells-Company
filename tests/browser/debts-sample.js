/* Synthetic data for the debts dashboard, keyed by the template's sheet and column keys (window.__debts.TPL).
   small(): a hand-made portfolio whose KPIs are checked to the dinar in debts.js.
   large(): a realistic month (about 60 customers, 900 invoices) for screenshots and a smoke run. All names are invented. */
function small() {
  return {
    per: [{ from: '2026-10-01', to: '2026-10-31' }],
    cust: [
      { a: 'C1', n: 'محطة النور للوقود', cat: 'محطة وقود', rg: 'المنطقة الغربية', wh: 'مستودع طرابلس', st: 'نشط', lim: 100000, terms: 30, ctr: 'ع/2026/115', ctrd: '2026-01-15' },
      { a: 'C2', n: 'شركة الأفق للنقل', cat: 'شركة خاصة', rg: 'المنطقة الشرقية', wh: 'مستودع بنغازي', st: 'موقوف', lim: 50000, terms: 60 },
      { a: 'C3', n: 'الشركة العامة للكهرباء', cat: 'كهرباء', rg: 'المنطقة الغربية', wh: 'مستودع طرابلس', st: 'نشط', lim: 0, terms: 90 },
      { a: 'C4', n: 'وزارة الصحة', cat: 'جهة حكومية', pub: 'وزارة', rg: 'المنطقة الغربية', st: 'نشط', lim: 0, terms: 60 }],
    guar: [{ a: 'C1', gt: 'خطاب ضمان مصرفي', ref: 'LG-44', bank: 'مصرف الجمهورية', v: 80000, exp: '2027-03-01' },
      { a: 'C2', gt: 'صك مصدق', ref: 'CQ-9', v: 10000, exp: '2026-09-30' }],
    inv: [
      { a: 'C1', no: 'I1', d: '2026-10-10', p: 'وقود الديزل', q: 400000, u: 'لتر', v: 60000, due: '2026-11-09', paid: 0, rem: 60000 },
      { a: 'C1', no: 'I2', d: '2026-08-01', p: 'البنزين', q: 330000, u: 'لتر', v: 50000, due: '2026-08-31', paid: 5000, rem: 45000 },
      { a: 'C2', no: 'I3', d: '2026-04-01', p: 'وقود الديزل', v: 30000, due: '2026-05-31', paid: 0, rem: 30000 },
      { a: 'C2', no: 'I4', d: '2026-10-05', p: 'وقود الديزل', v: 20000, due: '2026-10-20', paid: 20000, rem: 0 },
      { a: 'C3', no: 'I5', d: '2026-03-01', p: 'الزيت الثقيل', v: 400000, due: '2026-03-31', paid: 0, rem: 400000 },
      { a: 'C3', no: 'I6', d: '2026-10-15', p: 'الزيت الثقيل', v: 100000, due: '2026-11-14', paid: 0, rem: 100000 },
      { a: 'C4', no: 'I7', d: '2026-09-20', p: 'وقود الديزل', v: 10000, due: '2026-10-20', paid: 0, rem: 10000 }],
    pay: [{ a: 'C1', d: '2026-10-12', v: 5000, m: 'صك', ref: '004512', inv: 'I2' },
      { a: 'C2', d: '2026-10-18', v: 20000, m: 'تحويل', ref: 'TR-1', inv: 'I4' },
      { a: 'C3', d: '2026-10-25', v: 50000, m: 'مقاصة مع الدولة', ref: 'OFF-1' }],
    bnc: [{ a: 'C1', no: '004512', d: '2026-10-28', v: 5000, bank: 'مصرف الوحدة', why: 'عدم كفاية الرصيد', rec: 'لا' }],
    act: [{ a: 'C2', d: '2026-10-03', t: 'إنذار', ref: 'م/77', v: 30000 }, { a: 'C2', d: '2026-10-20', t: 'إحالة للشؤون القانونية', ref: 'ق/12' }, { a: 'C1', d: '2026-10-15', t: 'مطالبة', v: 45000 }],
    sch: [{ a: 'C2', ref: 'ج/9', d: '2026-06-01', v: 12000, n: 12, first: '2026-07-01', every: 1, paid: 1000 }],
    prov: [{ a: 'C2', t: 'مخصص', d: '2026-10-31', v: 15000, note: 'قرار لجنة الديون 4/2026' }, { a: 'C2', t: 'شطب', d: '2026-10-31', v: 2000, inv: 'I0' }],
    note: [{ a: 'C1', t: 'دائن', no: 'CN-1', d: '2026-10-15', inv: 'I2', v: 200, why: 'خصم كمية' }],
    off: [{ a: 'C3', d: '2026-10-25', sub: 900000, v: 50000, st: 'منفذة', ref: 'كتاب وزارة المالية 3321' }, { a: 'C3', d: '2026-10-30', v: 300000, st: 'قيد الإجراء' }]
  };
}
function large() {
  let s = 20261031; const rnd = () => { s = (s + 0x6D2B79F5) | 0; let t = Math.imul(s ^ (s >>> 15), 1 | s); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; const pick = a => a[Math.floor(rnd() * a.length)];
  const R = ['المنطقة الغربية', 'المنطقة الشرقية', 'المنطقة الجنوبية', 'المنطقة الوسطى'];
  const WH = { 'المنطقة الغربية': ['مستودع طرابلس', 'مستودع الزاوية', 'مستودع مصراتة'], 'المنطقة الشرقية': ['مستودع بنغازي', 'مستودع طبرق'], 'المنطقة الجنوبية': ['مستودع سبها'], 'المنطقة الوسطى': ['مستودع سرت'] };
  const towns = ['النور', 'الفجر', 'السلام', 'الوحدة', 'الريان', 'الأمل', 'الهلال', 'المدينة', 'الشاطئ', 'القلعة', 'النخيل', 'الربيع', 'الواحة', 'الصحراء', 'الجبل', 'الساحل', 'الزيتون', 'الكرامة', 'المروج', 'البحر'];
  const kinds = [['محطة وقود', n => 'محطة ' + n + ' للوقود', 28, [150000, 600000], [30]],
    ['شركة خاصة', n => 'شركة ' + n + ' للنقل والخدمات', 10, [200000, 1500000], [30, 60]],
    ['صناعة', n => 'مصنع ' + n + ' للإسمنت', 5, [1000000, 5000000], [60, 90]],
    ['طيران', n => 'شركة طيران ' + n, 3, [2000000, 8000000], [30, 60]],
    ['كهرباء', n => 'محطة كهرباء ' + n, 5, [0, 0], [90]],
    ['جهة حكومية', n => pick(['وزارة الصحة', 'وزارة التعليم', 'وزارة الداخلية', 'بلدية ' + n, 'جهاز الإسعاف', 'بلدية ' + n + ' المركز']), 9, [0, 0], [60, 90]]];
  const prods = [['وقود الديزل', 'لتر', 0.15], ['البنزين', 'لتر', 0.15], ['الزيت الثقيل', 'طن', 120], ['الكيروسين', 'لتر', 0.13]];
  const out = { per: [{ from: '2026-10-01', to: '2026-10-31' }], cust: [], guar: [], inv: [], pay: [], bnc: [], act: [], sch: [], prov: [], note: [], off: [] };
  const iso = t => new Date(t).toISOString().slice(0, 10); const END = Date.UTC(2026, 9, 31);
  let k = 1, ino = 1, used = new Set();
  kinds.forEach(([cat, mk, count, lim, terms]) => {
    for (let i = 0; i < count; i++) {
      let name, tries = 0; do { name = mk(pick(towns)); } while (used.has(name) && ++tries < 50); if (used.has(name)) name += ' ' + k; used.add(name);
      const a = 'C' + String(1000 + k++); const rg = pick(R); const gov = cat === 'كهرباء' || cat === 'جهة حكومية';
      const st = gov ? 'نشط' : rnd() < 0.08 ? 'متعثر' : rnd() < 0.08 ? 'موقوف' : rnd() < 0.04 ? 'نزاع قانوني' : 'نشط';
      const L = lim[1] ? Math.round((lim[0] + rnd() * (lim[1] - lim[0])) / 10000) * 10000 : 0; const T = pick(terms);
      out.cust.push({ a, n: name, cat, pub: cat === 'جهة حكومية' ? (/وزارة/.test(name) ? 'وزارة' : /بلدية/.test(name) ? 'بلدية' : 'جهة عامة أخرى') : cat === 'كهرباء' ? 'الكهرباء' : '', rg, wh: pick(WH[rg]), st, lim: L, terms: T, ctr: 'ع/2026/' + (100 + k), ctrd: '2026-0' + (1 + Math.floor(rnd() * 9)) + '-1' + Math.floor(rnd() * 9) });
      if (!gov && rnd() < 0.7) out.guar.push({ a, gt: pick(['خطاب ضمان مصرفي', 'خطاب ضمان مصرفي', 'صك مصدق', 'رهن']), ref: 'G-' + k, bank: pick(['مصرف الجمهورية', 'مصرف الوحدة', 'مصرف التجارة والتنمية', 'المصرف التجاري الوطني']), v: Math.round(L * (0.3 + rnd() * 0.7) / 1000) * 1000, exp: iso(END + (rnd() * 400 - 60) * 864e5) });
      const late = gov ? 0.75 : st === 'متعثر' || st === 'نزاع قانوني' ? 0.85 : 0.25;
      const n = gov ? 6 + Math.floor(rnd() * 8) : cat === 'محطة وقود' ? 8 + Math.floor(rnd() * 14) : 10 + Math.floor(rnd() * 14);
      const heavy = !gov && (st === "متعثر" || rnd() < 0.12) ? 2.2 : 1; const scale = (gov ? (800000 + rnd() * 6e6) / n : L ? L * (0.5 + rnd() * 0.9) / n : 100000 / n) * heavy;
      for (let j = 0; j < n; j++) {
        const age = rnd() < late ? Math.floor(rnd() * (gov ? 420 : 260)) : Math.floor(rnd() * 40);
        const d = iso(END - (age + Math.floor(rnd() * 5)) * 864e5); const due = iso(Date.parse(d) + T * 864e5);
        const [p, u, price] = pick(prods); const v = Math.round(scale * (0.3 + rnd()) / 10) * 10; const q = Math.round(v / price);
        const paidShare = Date.parse(due) > END ? (rnd() < 0.3 ? rnd() * 0.5 : 0) : rnd() < (gov ? 0.25 : 0.55) ? 1 : rnd() * (gov ? 0.3 : 0.6);
        const paid = Math.round(v * paidShare); const no = 'F-26-' + String(ino++).padStart(5, '0');
        out.inv.push({ a, no, d, p, q, u, v, due, paid, rem: v - paid });
        if (paid > 0 && rnd() < 0.8) { const pd = iso(Math.min(END, Date.parse(due) - 864e5 * Math.floor(rnd() * 20))); if (pd >= '2026-10-01') out.pay.push({ a, d: pd, v: paid, m: gov ? pick(['تحويل', 'مقاصة مع الدولة', 'مقاصة مع الدولة']) : pick(['صك', 'صك', 'تحويل', 'نقدا']), ref: 'R-' + ino, inv: no }); }
      }
      if (!gov && rnd() < 0.12) out.bnc.push({ a, no: String(400000 + k), d: iso(END - Math.floor(rnd() * 30) * 864e5), v: Math.round(scale * 0.2), bank: pick(['مصرف الوحدة', 'مصرف الجمهورية']), why: 'عدم كفاية الرصيد', rec: rnd() < 0.3 ? 'نعم' : 'لا' });
      if (late > 0.5 || rnd() < 0.3) out.act.push({ a, d: iso(END - Math.floor(rnd() * 28) * 864e5), t: gov ? 'مطالبة' : st === 'نزاع قانوني' ? 'إحالة للشؤون القانونية' : pick(['مطالبة', 'مطالبة', 'إنذار']), ref: 'م/2026/' + k });
      if (!gov && (st === 'متعثر' || rnd() < 0.06)) { const v = Math.round(scale * 2 / 1000) * 1000; out.sch.push({ a, ref: 'ج/2026/' + k, d: '2026-05-01', v, n: 12, first: '2026-06-01', every: 1, paid: Math.round(v / 12 * (1 + Math.floor(rnd() * 5))) }); }
      if (!gov && (st === 'متعثر' || st === 'نزاع قانوني')) out.prov.push({ a, t: 'مخصص', d: '2026-10-31', v: Math.round(scale * 1.5), note: 'قرار لجنة الديون 4/2026' });
      if (!gov && rnd() < 0.15) out.note.push({ a, t: rnd() < 0.8 ? 'دائن' : 'مدين', no: 'N-' + k, d: iso(END - Math.floor(rnd() * 30) * 864e5), v: Math.round(scale * 0.02), why: 'خصم كمية' });
      if (gov && rnd() < 0.6) out.off.push({ a, d: iso(END - Math.floor(rnd() * 30) * 864e5), sub: Math.round(scale * 3), v: Math.round(scale * (0.5 + rnd())), st: rnd() < 0.5 ? 'منفذة' : 'قيد الإجراء', ref: 'كتاب وزارة المالية ' + (3000 + k) });
    }
  });
  out.prov.push({ a: out.cust[3].a, t: 'شطب', d: '2026-10-20', v: 12000, note: 'قرار شطب 2/2026' });
  return out;
}
module.exports = { small, large };
