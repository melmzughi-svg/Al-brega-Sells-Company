/* ===== لوحة الديون والمبيعات الآجلة (إدارة الشؤون التجارية والتحصيلات) =====
   Credit sales and overdue debts, customer by customer, built on the SMS055 page pattern (sms055.js styles, readers, charts).
   Source: one workbook from the company's accounts, filled from the template this page downloads («قالب بيانات الديون»):
     الفترة            period start / end (the aging date is the period end)
     الزبائن           customer master: code, name, category, region, warehouse, account status, credit limit, payment terms,
                       contract number / date, public-body type (for government customers)
     الضمانات          guarantees: type (bank letter of guarantee, certified cheque, mortgage), value, expiry
     الفواتير          credit sales invoice by invoice: number, date, product, quantity, value, due date, paid, remaining
     التحصيلات         payments: date, amount, method (cash, cheque, transfer, offset with the state), invoices settled
     الصكوك المرتجعة    returned cheques
     إجراءات التحصيل   claims, warnings and referrals to Legal Affairs
     الجدولة           rescheduling agreements: amount, instalments, instalments paid
     المخصصات والشطب   provision for doubtful debts and written-off debts
     الإشعارات         credit and debit notes (discounts and adjustments after the invoice)
     مقاصة الدعم       offsets between government debts and the subsidy owed to the company
     أعمار الديون      optional: customer balances already split into the six aging buckets, when invoices are not supplied
   Aging buckets at the period end: not yet due, 1–30, 31–60, 61–90, 91–180, over 180 days.
   Government bodies (category gov or power) are kept apart from the private sector: the KPIs are computed per sector.
   The monthly «أرصدة حسابات المدينون والدائنون التجاريون» statement (BPMC-FIN-TOC-F38) is still read: it gives customer
   balances, withdrawals and collections (no invoices, so no aging).
   Storage: Firestore collection `debts_report`, one `batch` document per imported snapshot (sheet = ar, to = period end)
   and `part` documents holding its JSON (same rights as sms055, see firestore.rules). */
"use strict";
(function(){
const L=window.__smsLib;if(!L){console.warn('debts: sms055 library missing');return}
const {hn,AD,num,serialIso,colChart,bars,stack,pct,fdi,mName,me}=L;
const readGrids=f=>(window.FileImport&&FileImport.readGrids?FileImport.readGrids(f):L.readGrids(f));
const ACCEPT=()=>(window.FileImport&&FileImport.ACCEPT)||'.xlsx,.xls,.csv,.htm,.html';
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const ID='debts',COL='debts_report',UNIT='ca',CHUNK=800000,DEPT=['إدارة الشؤون التجارية والتحصيلات','Commercial Affairs & Collections'];
const LABEL=['لوحة الديون','Debts dashboard'];
const TPL_NAME='Debts-data-template.xlsx';
/* amounts: 6.4 مليار / 218.6 مليون / 52 ألف */
const big=v=>{if(v==null||isNaN(v))return'—';const a=Math.abs(v);if(a>=1e9)return String(+(v/1e9).toFixed(a>=1e10?1:2))+tr(' مليار',' bn');if(a>=1e6)return String(+(v/1e6).toFixed(a>=1e8?0:1))+tr(' مليون',' M');if(a>=1e4)return fn(Math.round(v/1e3))+tr(' ألف',' K');return fn(Math.round(v))};
const n0=v=>fn(Math.round(v||0));
const txt=v=>String(v==null?'':(v&&typeof v==='object'&&'serial' in v?serialIso(v.serial):v)).trim();
const H=v=>hn(txt(v));
const near=(a,b)=>Math.abs(a-b)<=Math.max(1,Math.abs(b)*1e-6);
const r2=v=>Math.round((+v||0)*100)/100;
const r3=v=>Math.round((+v||0)*1000)/1000;
const MONTHS=[['يناير','اي النار','كانون الثاني'],['فبراير','النوار','شباط'],['مارس','الربيع','اذار'],['ابريل','الطير','نيسان'],['مايو','الماء','ايار'],['يونيو','الصيف','حزيران'],
  ['يوليو','ناصر','تموز'],['اغسطس','هانيبال','اب'],['سبتمبر','الفاتح','ايلول'],['اكتوبر','التمور','تشرين الاول'],['نوفمبر','الحرث','تشرين الثاني'],['ديسمبر','الكانون','كانون الاول']];
const monthEnd=(y,m)=>new Date(Date.UTC(y,m,0)).toISOString().slice(0,10);
const mLabel=d=>d?mName(+d.slice(5,7)-1)+' '+d.slice(0,4):'—';
const DAY=864e5;
const dnum=d=>d?Date.UTC(+d.slice(0,4),+d.slice(5,7)-1,+d.slice(8,10))/DAY:NaN;
const ddiff=(a,b)=>Math.round(dnum(a)-dnum(b));
const addDays=(d,n)=>new Date((dnum(d)+n)*DAY).toISOString().slice(0,10);
const addMonths=(d,n)=>{const y=+d.slice(0,4),m=+d.slice(5,7)-1+n,day=+d.slice(8,10);const last=new Date(Date.UTC(y,m+1,0)).getUTCDate();return new Date(Date.UTC(y,m,Math.min(day,last))).toISOString().slice(0,10)};
const P1=v=>v==null||!isFinite(v)?'—':(Math.round(v*10)/10)+'%';

/* ---------- shared customer categories (the profit-margin dashboard uses the same keys) ---------- */
if(!window.CUSTOMER_CATS){
  window.CUSTOMER_CATS=[{k:'station',ar:'محطة وقود',en:'Fuel station',sector:'private'},{k:'gov',ar:'جهة حكومية',en:'Government body',sector:'gov'},
    {k:'private',ar:'شركة خاصة',en:'Private company',sector:'private'},{k:'aviation',ar:'طيران',en:'Aviation',sector:'private'},
    {k:'industry',ar:'صناعة',en:'Industry',sector:'private'},{k:'power',ar:'كهرباء',en:'Electricity',sector:'gov'}];
  window.custCat=function(s){const t=hn(String(s||''));if(!t)return '';
    const k=window.CUSTOMER_CATS.find(c=>c.k===t.toLowerCase()||hn(c.ar)===t||c.en.toLowerCase()===t.toLowerCase());if(k)return k.k;
    if(/كهرب|electric|power|gecol/i.test(t))return 'power';
    if(/طيران|مطار|جويه|الخطوط|aviation|airline|airport|jet/i.test(t))return 'aviation';
    if(/محط|station|توزيع|وكيل/i.test(t))return 'station';
    if(/حكوم|وزار|بلدي|عام[هة]|هيئ|جهاز|مصلح|ممول|government|ministry|municipal|public/i.test(t))return 'gov';
    if(/صناع|مصنع|مصانع|اسمنت|حديد|صلب|industr|factory|cement|steel/i.test(t))return 'industry';
    if(/شرك|خاص|private|company|co\b/i.test(t))return 'private';
    return ''};
}
const CATS=()=>window.CUSTOMER_CATS;
const catOf=k=>CATS().find(c=>c.k===k);
const catName=k=>{const c=catOf(k);return c?tr(c.ar,c.en):tr('غير مصنّف','Unclassified')};
const isGov=k=>{const c=catOf(k);return !!c&&c.sector==='gov'};
/* public bodies: electricity, ministries, municipalities, other public bodies */
const PUB={power:['الكهرباء','Electricity'],min:['الوزارات','Ministries'],mun:['البلديات','Municipalities'],oth:['جهات عامة أخرى','Other public bodies']};
function pubType(c){if(c.cat==='power')return 'power';const t=hn((c.pub||'')+' '+c.n);if(/كهرب/.test(t))return 'power';if(/وزار|ministry/i.test(t))return 'min';if(/بلدي|municipal/i.test(t))return 'mun';return 'oth'}

/* ---------- vocabularies ---------- */
const STATUS={active:['نشط','Active','ok'],stopped:['موقوف','Suspended','warn'],default:['متعثر','Defaulting','crit'],legal:['نزاع قانوني','Legal dispute','crit']};
function statusOf(s){const t=hn(String(s||''));if(!t)return 'active';if(/نزاع|قانون|قضا|محكم|legal|dispute/i.test(t))return 'legal';if(/متعثر|default|bad/i.test(t))return 'default';if(/موقوف|ايقاف|معلق|suspend|stop|block/i.test(t))return 'stopped';return 'active'}
const GTYPE={lg:['خطاب ضمان مصرفي','Bank letter of guarantee'],cert:['صك مصدّق','Certified cheque'],mort:['رهن','Mortgage'],oth:['ضمان آخر','Other guarantee']};
function gtypeOf(s){const t=hn(String(s||''));if(/خطاب|مصرف|بنك|letter|lg|bank/i.test(t))return 'lg';if(/صك|شيك|مصدق|cheque|check/i.test(t))return 'cert';if(/رهن|عقار|mortgage|pledge/i.test(t))return 'mort';return 'oth'}
const METHOD={cash:['نقدًا','Cash'],chq:['صك','Cheque'],trf:['تحويل','Transfer'],off:['مقاصّة مع الدولة','Offset with the state']};
function methodOf(s){const t=hn(String(s||''));if(/مقاص|تسويه مع الدوله|دعم|offset|netting/i.test(t))return 'off';if(/صك|شيك|cheque|check/i.test(t))return 'chq';if(/تحويل|حواله|مصرف|بنك|transfer|wire|bank/i.test(t))return 'trf';return 'cash'}
const ACT={claim:['مطالبة','Claim'],warn:['إنذار','Warning'],legal:['إحالة للشؤون القانونية','Referral to Legal Affairs'],visit:['زيارة / اتصال','Visit / call']};
function actOf(s){const t=hn(String(s||''));if(/احال|قانون|قضا|legal/i.test(t))return 'legal';if(/انذار|warning|notice/i.test(t))return 'warn';if(/زيار|اتصال|visit|call/i.test(t))return 'visit';return 'claim'}
const NOTE={cr:['إشعار دائن','Credit note'],dr:['إشعار مدين','Debit note']};
function noteOf(s){const t=hn(String(s||''));return /مدين|debit/i.test(t)?'dr':'cr'}
const yes=s=>/^(نعم|اي|ايوه|y|yes|true|1|تم|✓)$/i.test(hn(String(s||'')));
/* aging buckets at the period end */
const BUCKETS=[{k:'b0',ar:'غير مستحق بعد',en:'Not yet due',m:['لم يحل موعد سداده','not due yet'],lo:-Infinity,hi:0},
  {k:'b1',ar:'1 – 30 يومًا',en:'1–30 days',m:['متأخر قليلًا','slightly late'],lo:1,hi:30},
  {k:'b2',ar:'31 – 60 يومًا',en:'31–60 days',m:['يحتاج متابعة','needs follow-up'],lo:31,hi:60},
  {k:'b3',ar:'61 – 90 يومًا',en:'61–90 days',m:['متأخر','late'],lo:61,hi:90},
  {k:'b4',ar:'91 – 180 يومًا',en:'91–180 days',m:['خطر','at risk'],lo:91,hi:180},
  {k:'b5',ar:'أكثر من 180 يومًا',en:'Over 180 days',m:['مشكوك في تحصيله','doubtful'],lo:181,hi:Infinity}];
const BK=BUCKETS.map(b=>b.k);
const bucketOf=d=>d<=0?'b0':d<=30?'b1':d<=60?'b2':d<=90?'b3':d<=180?'b4':'b5';
const bName=k=>{const b=BUCKETS.find(x=>x.k===k);return tr(b.ar,b.en)};

/* ---------- reading the ledger workbook (BPMC-FIN-TOC-F38) ---------- */
/* a ledger table header: two «الرصيد في» columns (opening, closing) and «المسحوبات»; the next row says مدين / دائن */
function ledgerHeader(row,sub){const h=row.map(H);const bal=[];h.forEach((x,i)=>{if(/^الرصيد( في| السابق| الافتتاحي| الحالي| الختامي)?$/.test(x)||/^الرصيد في/.test(x))bal.push(i)});
  const w=h.findIndex(x=>/^المسحوبات/.test(x));if(bal.length<2||w<0)return null;const s=(sub||[]).map(H);const f=re=>h.findIndex(x=>re.test(x));
  const m={name:f(/^(اسم الزبون|اسم العميل|البيان|الزبون|العميل|الجهه)/),acct:f(/^رقم (الحساب|العميل|الزبون)/),ob:bal[0],cb:bal[bal.length-1],w,da:f(/^التسويات المدينه/),c:f(/^التحصيلات/),ca:f(/^التسويات الدا[يئ]نه/)};
  if(m.name<0)m.name=0;const cred=(lo,hi)=>{for(let i=lo+1;i<hi;i++)if(/دا[يئ]ن/.test(s[i]||''))return i;return -1};
  m.oc=cred(m.ob,w);if(m.oc<0&&m.ob+1<w&&m.ob+1!==m.acct)m.oc=m.ob+1;
  m.cc=cred(m.cb,m.cb+4);if(m.cc<0){for(let i=m.cb+1;i<Math.min(row.length,m.cb+3);i++)if(txt(row[i])){m.cc=i;break}}
  const ds=[];row.forEach((v,i)=>{const d=cellDate(v);if(d)ds.push([i,d])});
  m.prev=(ds.find(([i])=>i>m.ob&&i<=m.ob+2)||[])[1]||'';m.asOf=(ds.find(([i])=>i>m.cb&&i<=m.cb+2)||[])[1]||'';
  m.sub=s.some(x=>/مدين|دا[يئ]ن/.test(x));return m}
const VALS=['ob','oc','w','da','c','ca','cb','cc'];
function rowVals(m,r){const g=i=>i!=null&&i>=0?num(r[i]):0;
  return {ob:g(m.ob),oc:Math.abs(g(m.oc)),w:g(m.w),da:g(m.da),c:Math.abs(g(m.c)),ca:Math.abs(g(m.ca)),cb:g(m.cb),cc:Math.abs(g(m.cc))}}
const isNum=x=>typeof x==='number'||(typeof x==='string'&&/^\(?-?[\d,.٠-٩]+\)?$/.test(AD(x).trim())&&/\d/.test(AD(x)));
function cleanSec(s){return String(s).replace(/["“”«»]/g,' ').replace(/^\s*(تابع\s+)?ديون\s+(حسابات\s+)?/,'').replace(/(بالعملة|العملة)\s+(المحلية|الأجنبية|الاجنبية)/g,'')
  .replace(/المنطقة\s+(الغربية|الشرقية)/g,'').replace(/\s+/g,' ').trim()||String(s).trim()}
/* summary rows → categories: a row whose name ends with غ / ش is a category of that region; a row that equals the sum of the
   rows before it is their sector subtotal; «إجمالي ديون …» is the stated grand total of the currency */
function groupify(rows){const cats=[],stated={};let P=[];
  const flush=()=>{P.forEach(p=>{p.g=p.g||p.n;cats.push(p)});P=[]};
  rows.forEach(x=>{const hx=H(x.n);
    if(/^اجمال/.test(hx)&&/ديون/.test(hx)){flush();stated[x.cur]=x;return}
    const m=x.n.match(/^(.*\S)\s+(غ|ش)$/);if(m){x.n=m[1].trim();x.rg=m[2]==='غ'?'W':'E';P.push(x);return}
    const sum=k=>P.reduce((s,p)=>s+p[k],0);
    if(P.length>1&&P.every(p=>p.cur===x.cur)&&near(sum('cb'),x.cb)&&near(sum('w'),x.w)&&near(sum('ob'),x.ob)){const g=/^اجمال/.test(hx)?P[0].n:x.n.trim();P.forEach(p=>p.g=g);flush();return}
    if(/^اجمال/.test(hx))return;
    x.rg='';P.push(x)});
  flush();return {cats,stated}}
function parseLedger(sheets,file){const custs=[],sums=[],dates=[];const seen=new Set();
  for(const sh of sheets){const g=sh.grid,sn=H(sh.name);const sheetRg=/الشرقيه/.test(sn)?'E':/الغربيه/.test(sn)?'W':'';
    let cur='L',sec='',rg=sheetRg,hdr=null;const srows=[];
    for(let i=0;i<g.length;i++){const r=g[i]||[];const filled=r.filter(x=>txt(x)!=='');if(!filled.length)continue;
      const h=ledgerHeader(r,g[i+1]);if(h){hdr=h;if(h.prev)dates.push(h.prev);if(h.asOf)dates.push(h.asOf);if(h.sub)i++;continue}
      const numeric=r.some((x,j)=>isNum(x)&&(!hdr||j!==hdr.acct));
      if(!numeric){const t=H(filled.map(txt).join(' '));if(/الاجنبيه/.test(t))cur='F';else if(/المحليه/.test(t))cur='L';
        if(/ليوم/.test(t))r.forEach(v=>{const d=cellDate(v);if(d)dates.push(d)});
        if(/ديون/.test(t)&&!/رقم الوثيقه/.test(t)){sec=cleanSec(filled.map(txt).join(' '));rg=/الشرقيه/.test(t)?'E':/الغربيه/.test(t)?'W':sheetRg;if(/الاجنبيه/.test(t))cur='F';else if(/المحليه/.test(t))cur='L'}
        continue}
      if(!hdr)continue;const name=txt(r[hdr.name]);if(!name||/مدين|دا[يئ]ن/.test(H(name)))continue;const v=rowVals(hdr,r);
      if(hdr.acct>=0){if(/^اجمال/.test(H(name)))continue;const a=txt(r[hdr.acct]);if(!a)continue;if(VALS.every(k=>!v[k]))continue;
        const key=a+'|'+cur+'|'+rg;if(seen.has(key))continue;seen.add(key);custs.push(Object.assign({a,n:name.replace(/\s+/g,' '),rg,sec,cur},v))}
      else srows.push(Object.assign({n:name.replace(/\s+/g,' ').trim(),cur},v))}
    if(srows.length)sums.push(Object.assign(groupify(srows),{sheet:sh.name}))}
  /* the summary sheet with the most region categories is the base; a sectors-only sheet gives the sector names */
  sums.sort((a,b)=>b.cats.filter(c=>c.rg).length-a.cats.filter(c=>c.rg).length||b.cats.length-a.cats.length);
  let cats=[],stated={},sheet='';if(sums.length){cats=sums[0].cats;stated=Object.assign({},...sums.slice().reverse().map(s=>s.stated));sheet=sums[0].sheet;
    const names=sums.slice(1).flatMap(s=>s.cats.filter(c=>c.g===c.n&&!c.rg));
    [...new Set(cats.map(c=>c.g))].forEach(gn=>{const gs=cats.filter(c=>c.g===gn);const cb=gs.reduce((s,c)=>s+c.cb,0),w=gs.reduce((s,c)=>s+c.w,0);
      const hit=names.find(x=>x.cur===gs[0].cur&&near(x.cb,cb)&&near(x.w,w));if(hit)gs.forEach(c=>c.g=hit.n.trim())})}
  else if(custs.length){const m=new Map();custs.forEach(c=>{const k=c.sec+'|'+c.rg+'|'+c.cur;let x=m.get(k);if(!x)m.set(k,x={n:c.sec||tr('حسابات أخرى','Other accounts'),g:c.sec||tr('حسابات أخرى','Other accounts'),rg:c.rg,cur:c.cur,ob:0,oc:0,w:0,da:0,c:0,ca:0,cb:0,cc:0});VALS.forEach(k2=>x[k2]+=c[k2])});cats=[...m.values()];sheet=tr('أوراق المناطق','region sheets')}
  if(!cats.length&&!custs.length)return null;
  dates.sort();let asOf=dates[dates.length-1]||'',prev=dates.length>1?dates[0]:'';
  if(!asOf){const x=AD(file||'').match(/(?:^|\D)(\d{1,2})\s*[-_.\/ ]\s*(20\d\d)(?:\D|$)/);if(x)asOf=monthEnd(+x[2],+x[1])}
  if(!asOf)throw new Error(tr('لم نجد تاريخ الأرصدة («كما هي في دفاترنا الرئيسية ليوم …») في الملف، ولا شهرًا في اسمه.','The balance date («as in our books on …») was not found in the file, nor a month in its name.'));
  if(!prev||prev>=asOf)prev=new Date(Date.UTC(+asOf.slice(0,4),+asOf.slice(5,7)-1,0)).toISOString().slice(0,10);
  const rnd=o=>{VALS.forEach(k=>o[k]=r3(o[k]));return o};
  return {T:'ledger',asOf,prev,sheet,cats:cats.map(c=>rnd({n:c.n,g:c.g,rg:c.rg||'',cur:c.cur,ob:c.ob,oc:c.oc,w:c.w,da:c.da,c:c.c,ca:c.ca,cb:c.cb,cc:c.cc})),
    custs:custs.map(c=>rnd(c)),stated:Object.fromEntries(Object.entries(stated).map(([k,v])=>[k,rnd(Object.fromEntries(VALS.map(x=>[x,v[x]])))]))}}


/* ---------- the import template: one definition used to write the template and to read it back ---------- */
/* column: [key, Arabic header, English header, type (s text · n number · d date), required, example, note (ar), aliases] */
const CODE=['a','كود الزبون','Customer code','s',1,'C-1001','رقم الزبون كما في حسابات الشركة؛ يربط كل الأوراق ببعضها.',['رقم الزبون','رقم الحساب','الكود','رمز الزبون','رقم العميل','كود العميل','customer code','customer no','account no','code']];
const CNAME=['n','اسم الزبون','Customer name','s',0,'محطة النور للوقود','اختياري في الأوراق الأخرى؛ يُستعمل إن لم يكن الزبون في ورقة الزبائن.',['الزبون','اسم العميل','العميل','الجهه','customer','customer name','name']];
const TPL=[
 {k:'per',ar:'الفترة',en:'Period',names:['الفتره','period','الفتره الماليه'],desc:'صف واحد: بداية الفترة ونهايتها. أعمار الديون تُحسب في تاريخ نهاية الفترة.',cols:[
   ['from','بداية الفترة','Period start','d',1,'2026-10-01','أول يوم في الفترة.',['من','من تاريخ','بدايه الفتره','from','start']],
   ['to','نهاية الفترة','Period end','d',1,'2026-10-31','تاريخ آخر الفترة؛ تاريخ احتساب الأعمار.',['الى','إلى','حتى','الى تاريخ','نهايه الفتره','كما في','تاريخ الاعمار','to','end','as of']]]},
 {k:'cust',ar:'الزبائن',en:'Customers',names:['الزبائن','بيانات الزبائن','العملاء','customers'],desc:'بيانات الزبون الأساسية وشروط الائتمان: صف لكل زبون.',cols:[
   CODE,['n','اسم الزبون','Customer name','s',1,'محطة النور للوقود','',CNAME[7]],
   ['cat','الفئة','Category','s',1,'محطة وقود','محطة وقود · جهة حكومية · شركة خاصة · طيران · صناعة · كهرباء',['فئه الزبون','التصنيف','النوع','category','segment']],
   ['pub','نوع الجهة العامة','Public body type','s',0,'','للجهات الحكومية فقط: الكهرباء · وزارة · بلدية · جهة عامة أخرى',['الجهه العامه','نوع الجهه','public body']],
   ['rg','المنطقة','Region','s',0,'المنطقة الغربية','',['region','الاقليم']],
   ['wh','المستودع','Warehouse','s',0,'مستودع طرابلس','المستودع الذي يُورِّد منه.',['مستودع التوريد','المستودع المورد','warehouse','depot']],
   ['st','حالة الحساب','Account status','s',1,'نشط','نشط · موقوف · متعثر · نزاع قانوني',['الحاله','حاله الزبون','status']],
   ['lim','سقف الائتمان','Credit limit','n',1,500000,'أقصى مبلغ مسموح به للزبون بالدينار.',['الحد الائتماني','حد الائتمان','السقف الائتماني','credit limit','limit']],
   ['terms','مدة السداد (يوم)','Payment terms (days)','n',1,30,'30 أو 60 أو 90 يومًا.',['مده السداد','فتره السداد','مده الائتمان','payment terms','terms']],
   ['ctr','رقم العقد','Contract number','s',0,'ع/2026/115','رقم العقد أو الاتفاقية.',['رقم الاتفاقيه','رقم العقد او الاتفاقيه','العقد','contract','contract no']],
   ['ctrd','تاريخ العقد','Contract date','d',0,'2026-01-15','',['تاريخ الاتفاقيه','contract date']],
   ['bal','الرصيد في نهاية الفترة','Closing balance','n',0,'','اختياري: للمطابقة مع مجموع المتبقي من الفواتير.',['الرصيد','رصيد الزبون','الرصيد الحالي','الرصيد الختامي','balance','closing balance']]]},
 {k:'guar',ar:'الضمانات',en:'Guarantees',names:['الضمانات','guarantees'],desc:'صف لكل ضمان.',cols:[CODE,
   ['gt','نوع الضمان','Guarantee type','s',1,'خطاب ضمان مصرفي','خطاب ضمان مصرفي · صك مصدّق · رهن',['النوع','type','guarantee type']],
   ['ref','رقم الضمان','Guarantee number','s',0,'LG-2026-044','',['رقم الخطاب','المرجع','reference','ref']],
   ['bank','المصرف','Bank','s',0,'مصرف الجمهورية','',['البنك','الجهه المصدره','bank']],
   ['v','قيمة الضمان','Guarantee value','n',1,300000,'',['القيمه','المبلغ','value','amount']],
   ['iss','تاريخ الإصدار','Issue date','d',0,'2026-01-20','',['issue date']],
   ['exp','تاريخ الانتهاء','Expiry date','d',1,'2027-01-19','',['تاريخ انتهاء الضمان','صالح حتى','expiry','expiry date']]]},
 {k:'inv',ar:'الفواتير',en:'Invoices',names:['الفواتير','الفواتير الاجله','المبيعات الاجله','حركه المبيعات الاجله','invoices','credit sales'],desc:'حركة المبيعات الآجلة فاتورة بفاتورة.',cols:[CODE,CNAME,
   ['no','رقم الفاتورة','Invoice number','s',1,'F-26-10-0457','',['الفاتوره','invoice','invoice no']],
   ['d','تاريخ الفاتورة','Invoice date','d',1,'2026-10-04','',['التاريخ','invoice date','date']],
   ['p','المنتج','Product','s',0,'وقود الديزل','',['الصنف','نوع المنتج','product']],
   ['q','الكمية','Quantity','n',0,36000,'',['quantity','qty']],
   ['u','الوحدة','Unit','s',0,'لتر','',['unit']],
   ['v','قيمة الفاتورة','Invoice value','n',1,5400,'بالدينار.',['القيمه','المبلغ','اجمالي الفاتوره','value','amount']],
   ['due','تاريخ الاستحقاق','Due date','d',1,'2026-11-03','إن تُرك فارغًا يُحسب من تاريخ الفاتورة ومدة السداد.',['الاستحقاق','due','due date']],
   ['paid','المبلغ المسدد','Paid','n',1,2000,'المسدد من الفاتورة حتى نهاية الفترة.',['المسدد','المسدد من الفاتوره','paid']],
   ['rem','المتبقي','Remaining','n',1,3400,'المتبقي من الفاتورة بعد السداد والإشعارات والشطب.',['المتبقي من الفاتوره','الرصيد المتبقي','الباقي','remaining','outstanding']]]},
 {k:'pay',ar:'التحصيلات',en:'Collections',names:['التحصيلات','الدفعات','المقبوضات','collections','payments'],desc:'صف لكل دفعة.',cols:[CODE,CNAME,
   ['d','تاريخ الدفعة','Payment date','d',1,'2026-10-20','',['التاريخ','تاريخ التحصيل','payment date','date']],
   ['v','مبلغ الدفعة','Amount','n',1,2000,'',['المبلغ','المبلغ المحصل','القيمه','amount']],
   ['m','طريقة الدفع','Payment method','s',1,'صك','نقدًا · صك · تحويل · مقاصّة مع الدولة',['طريقه السداد','method','payment method']],
   ['ref','رقم الإيصال','Receipt number','s',0,'R-8812','رقم الإيصال أو الصك أو الحوالة.',['رقم المستند','رقم الصك','المرجع','receipt','reference']],
   ['inv','الفواتير المسددة','Invoices settled','s',0,'F-26-10-0457','رقم فاتورة أو أكثر مفصولة بفاصلة.',['الفواتير','رقم الفاتوره','رقم الفاتوره المسدده','invoices']],
   ['bnc','صك مرتجع؟','Returned cheque?','s',0,'','نعم للصك المرتجع: لا يُحسب تحصيلًا.',['صك مرتجع','مرتجع','returned'],1]]},
 {k:'bnc',ar:'الصكوك المرتجعة',en:'Returned cheques',names:['الصكوك المرتجعه','الشيكات المرتجعه','returned cheques','bounced cheques'],desc:'صف لكل صك مرتجع.',cols:[CODE,
   ['no','رقم الصك','Cheque number','s',1,'004512','',['الصك','cheque','cheque no']],
   ['d','تاريخ الإرجاع','Return date','d',1,'2026-10-25','',['التاريخ','date']],
   ['v','قيمة الصك','Cheque value','n',1,1500,'',['القيمه','المبلغ','amount']],
   ['bank','المصرف','Bank','s',0,'مصرف الوحدة','',['البنك','bank']],
   ['why','سبب الإرجاع','Reason','s',0,'عدم كفاية الرصيد','',['السبب','reason']],
   ['rec','تم التحصيل لاحقًا','Recovered later','s',0,'لا','نعم أو لا.',['حصل لاحقا','recovered']]]},
 {k:'act',ar:'إجراءات التحصيل',en:'Collection actions',names:['اجراءات التحصيل','المطالبات','المطالبات والانذارات','collection actions'],desc:'المطالبات والإنذارات المرسلة والديون المحالة للشؤون القانونية.',cols:[CODE,
   ['d','التاريخ','Date','d',1,'2026-10-12','',['تاريخ الاجراء','date']],
   ['t','نوع الإجراء','Action type','s',1,'مطالبة','مطالبة · إنذار · إحالة للشؤون القانونية · زيارة / اتصال',['النوع','الاجراء','type','action']],
   ['ref','رقم المراسلة','Letter number','s',0,'م/2026/77','',['المرجع','reference']],
   ['v','المبلغ المطالب به','Amount claimed','n',0,3400,'',['المبلغ','amount']],
   ['note','ملاحظات','Notes','s',0,'','',['notes']]]},
 {k:'sch',ar:'الجدولة',en:'Rescheduling',names:['الجدوله','اتفاقيات الجدوله','rescheduling'],desc:'اتفاقيات جدولة الديون ومدى الالتزام بها.',cols:[CODE,
   ['ref','رقم الاتفاقية','Agreement number','s',0,'ج/2026/09','',['الاتفاقيه','رقم الاتفاق','agreement']],
   ['d','تاريخ الاتفاقية','Agreement date','d',0,'2026-06-01','',['تاريخ الاتفاق','date']],
   ['v','المبلغ المجدول','Rescheduled amount','n',1,120000,'',['المبلغ','amount']],
   ['n','عدد الأقساط','Instalments','n',1,12,'',['الاقساط','instalments']],
   ['inst','قيمة القسط','Instalment value','n',0,10000,'إن تُرك فارغًا = المبلغ ÷ عدد الأقساط.',['القسط','instalment']],
   ['first','تاريخ أول قسط','First instalment','d',1,'2026-07-01','',['اول قسط','first instalment']],
   ['every','كل كم شهر','Every (months)','n',0,1,'1 شهري، 3 ربع سنوي.',['الدوريه','every']],
   ['paid','المبلغ المسدد','Amount paid','n',1,30000,'المسدد من الاتفاقية حتى نهاية الفترة.',['المسدد','paid']]]},
 {k:'prov',ar:'المخصصات والشطب',en:'Provisions & write-offs',names:['المخصصات والشطب','المخصصات','الديون المعدومه','provisions'],desc:'مخصص الديون المشكوك فيها والديون المعدومة المشطوبة.',cols:[CODE,
   ['t','النوع','Type','s',1,'مخصص','مخصص · شطب',['type']],
   ['d','التاريخ','Date','d',1,'2026-10-31','',['date']],
   ['v','المبلغ','Amount','n',1,25000,'',['القيمه','amount']],
   ['inv','رقم الفاتورة','Invoice number','s',0,'','',['الفاتوره','invoice']],
   ['note','رقم القرار / السبب','Decision / reason','s',0,'قرار لجنة الديون 4/2026','',['السبب','القرار','note']]]},
 {k:'note',ar:'الإشعارات',en:'Credit & debit notes',names:['الاشعارات','الاشعارات والتسويات','الاشعارات الدائنه والمدينه','notes'],desc:'الإشعارات الدائنة والمدينة (الخصومات والتسويات اللاحقة للفاتورة).',cols:[CODE,
   ['t','نوع الإشعار','Note type','s',1,'دائن','دائن · مدين',['النوع','type']],
   ['no','رقم الإشعار','Note number','s',0,'CN-118','',['الاشعار','note no']],
   ['d','التاريخ','Date','d',1,'2026-10-15','',['date']],
   ['inv','رقم الفاتورة','Invoice number','s',0,'F-26-10-0457','',['الفاتوره','invoice']],
   ['v','المبلغ','Amount','n',1,200,'',['القيمه','amount']],
   ['why','السبب','Reason','s',0,'خصم كمية','',['reason']],
   ['ref','المرجع / رقم القرار','Reference','s',0,'','',['المرجع','رقم القرار','reference'],1],
   ['sub','الدعم المستحق للشركة','Subsidy owed to the company','n',0,'','لسطور المقاصّة فقط.',['الدعم المستحق','subsidy'],1],
   ['st','حالة المقاصة','Offset status','s',0,'','لسطور المقاصّة فقط: منفذة · قيد الإجراء.',['الحاله','status'],1]]},
 {k:'off',ar:'مقاصة الدعم',en:'Subsidy offsets',names:['مقاصه الدعم','المقاصه','مقاصه مع الدوله','subsidy offsets'],desc:'المقاصّة بين ديون الجهات الحكومية والدعم المستحق للشركة.',cols:[CODE,
   ['d','التاريخ','Date','d',1,'2026-10-30','',['date']],
   ['sub','الدعم المستحق للشركة','Subsidy owed to the company','n',0,900000,'',['الدعم المستحق','الدعم','subsidy']],
   ['v','قيمة المقاصة','Offset amount','n',1,600000,'',['المقاصه','المبلغ','amount']],
   ['st','الحالة','Status','s',1,'منفذة','منفذة · قيد الإجراء',['status']],
   ['ref','المرجع / القرار','Reference','s',0,'كتاب وزارة المالية 3321','',['المرجع','القرار','reference']]]},
 {k:'age',ar:'أعمار الديون',en:'Aging',names:['اعمار الديون','الاعمار','aging'],desc:'اختيارية: أرصدة الزبائن موزّعة على فئات العمر، إن لم تتوفر ورقة الفواتير.',cols:[CODE,CNAME,
   ['b0','غير مستحق بعد','Not yet due','n',1,0,'',['غير مستحق','لم يستحق','not due']],
   ['b1','1 – 30 يومًا','1–30 days','n',1,0,'',['1-30','1 30','من 1 الى 30']],
   ['b2','31 – 60 يومًا','31–60 days','n',1,0,'',['31-60','31 60','من 31 الى 60']],
   ['b3','61 – 90 يومًا','61–90 days','n',1,0,'',['61-90','61 90','من 61 الى 90']],
   ['b4','91 – 180 يومًا','91–180 days','n',1,0,'',['91-180','91 180','من 91 الى 180']],
   ['b5','أكثر من 180 يومًا','Over 180 days','n',1,0,'',['اكثر من 180','180+','over 180']]]}];
const TPLK=Object.fromEntries(TPL.map(s=>[s.k,s]));

/* header matching: ignore spaces, dashes, brackets and the word «يوم» */
const hk=s=>hn(txt(s)).toLowerCase().replace(/[()\[\]\-–—_:.\/%]/g,' ').replace(/(^|\s)(يوما|يوم|ايام|days?)(?=\s|$)/g,' ').replace(/\s+/g,' ').trim();
function headerMap(row,spec){const h=row.map(hk);const m={};let hits=0;
  spec.cols.forEach(c=>{const keys=[c[1],c[2],...(c[7]||[])].map(hk);let i=h.findIndex(x=>x&&keys.includes(x));
    if(i<0)i=h.findIndex(x=>x&&keys.some(k=>k.length>3&&x.startsWith(k)));
    if(i>=0&&!Object.values(m).includes(i)){m[c[0]]=i;hits++}});
  const req=spec.cols.filter(c=>c[4]&&c[0]!=='n');const missing=req.filter(c=>m[c[0]]==null);
  return {m,hits,missing,ok:missing.length===0||(spec.k!=='per'&&hits>=3&&missing.length<=Math.floor(req.length/3)&&m.a!=null)}}
function cellDate(v){if(v&&typeof v==='object'&&'serial' in v)return serialIso(v.serial);if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);if(typeof v==='number')return v>20000&&v<80000?serialIso(v):'';
  const s=AD(String(v||'')).trim();let x=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(x)return x[1]+'-'+x[2].padStart(2,'0')+'-'+x[3].padStart(2,'0');
  x=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);if(x){let a=+x[1],b=+x[2];const d=a>12?a:b,m=a>12?b:a;if(m<1||m>12||d<1||d>31)return '';return x[3]+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0')}return ''}
const isTotal=s=>/^(الاجمالي|اجمالي|المجموع|الاجمالي العام|total|grand total)/i.test(hn(s));
/* the grey sample row of a data-request workbook («مثال-001», «(مثال) شركة …») is never data */
const isSample=s=>/^\(?\s*مثال/.test(hn(String(s||'')));
/* read every sheet that looks like a template sheet: by its name first, then by its headers */
function readTemplate(sheets){const found={};const used=new Set();
  const tryRead=(sh,spec)=>{const g=sh.grid||[];for(let i=0;i<Math.min(g.length,15);i++){const hm=headerMap(g[i]||[],spec);if(!hm.ok)continue;
      const rows=[];for(let j=i+1;j<g.length;j++){const r=g[j]||[];if(!r.some(x=>txt(x)!==''))continue;const o={};
        spec.cols.forEach(c=>{const ix=hm.m[c[0]];if(ix==null)return;const v=r[ix];o[c[0]]=c[3]==='n'?(txt(v)===''?null:num(v)):c[3]==='d'?cellDate(v):txt(v)});
        if(spec.k!=='per'){if(!o.a||isTotal(o.a)||isTotal(o.n||'')||isSample(o.a)||isSample(o.n))continue}rows.push(o)}
      return {rows,sheet:sh.name,m:hm.m,hits:hm.hits,missing:hm.missing.map(c=>c[1])}}return null};
  TPL.forEach(spec=>{const names=[spec.ar,...spec.names].map(hk);const sh=sheets.find(s=>!used.has(s)&&names.includes(hk(s.name)));if(!sh)return;const r=tryRead(sh,spec);if(r){found[spec.k]=r;used.add(sh)}});
  sheets.forEach(sh=>{if(used.has(sh))return;let best=null;TPL.forEach(spec=>{if(found[spec.k])return;const r=tryRead(sh,spec);if(r&&(!best||r.hits>best.r.hits))best={spec,r}});
    if(best&&best.r.hits>=3){found[best.spec.k]=best.r;used.add(sh)}});
  return found}

/* a data-request workbook keeps the aging date on its instructions sheet: «تاريخ قياس أعمار الديون | 2026-10-31» */
function agingDate(sheets){for(const sh of sheets){const g=sh.grid||[];for(let i=0;i<Math.min(g.length,12);i++){const r=g[i]||[];
    const j=r.findIndex(x=>/تاريخ (قياس )?اعمار الديون|تاريخ القياس/.test(hn(txt(x))));if(j<0)continue;for(let k=j+1;k<r.length;k++){const d=cellDate(r[k]);if(d)return d}}}return ''}
function parseTemplate(sheets,file){const F=readTemplate(sheets);
  if(!F.inv&&!F.age&&!(F.cust&&F.cust.m.bal!=null))return null;
  const W=[];const R=k=>F[k]?F[k].rows:[];
  /* period */
  let from='',to='';const per=R('per').find(r=>r.to||r.from);if(per){from=per.from||'';to=per.to||''}
  if(!to){to=agingDate(sheets);if(to)W.push(['i',tr(`لا توجد ورقة «الفترة»؛ اعتُمد تاريخ قياس أعمار الديون ${fdi(to)} نهايةً للفترة.`,`No «Period» sheet; the aging date ${fdi(to)} was taken as the period end.`)])}
  if(!to){const ds=[...R('inv').map(r=>r.d),...R('pay').map(r=>r.d)].filter(Boolean).sort();const x=AD(file||'').match(/(?:^|\D)(\d{1,2})\s*[-_.\/ ]\s*(20\d\d)(?:\D|$)/);
    if(x)to=monthEnd(+x[2],+x[1]);else if(ds.length){const d=ds[ds.length-1];to=monthEnd(+d.slice(0,4),+d.slice(5,7))}
    if(to)W.push(['i',tr(`لا توجد ورقة «الفترة»؛ اعتُبر ${fdi(to)} نهاية الفترة.`,`No «Period» sheet; ${fdi(to)} was taken as the period end.`)])}
  if(!to)throw new Error(tr('أضف ورقة «الفترة» وفيها بداية الفترة ونهايتها (تاريخ احتساب أعمار الديون).','Add the «Period» sheet with the period start and end (the aging date).'));
  if(!from||from>to)from=to.slice(0,8)+'01';
  /* customers */
  const C=new Map();let noCat=0;
  R('cust').forEach(r=>{if(C.has(r.a)){W.push(['w',tr(`كود الزبون ${r.a} مكرر في ورقة الزبائن؛ أُخذ أول صف.`,`Customer code ${r.a} repeats; the first row was kept.`)]);return}
    const cat=window.custCat(r.cat)||window.custCat(r.n);if(!cat)noCat++;
    C.set(r.a,{a:r.a,n:(r.n||r.a).replace(/\s+/g,' '),cat,pub:r.pub||'',rg:r.rg||'',wh:r.wh||'',st:statusOf(r.st),lim:r2(r.lim),terms:Math.round(r.terms||0),ctr:r.ctr||'',ctrd:r.ctrd||'',bal:r.bal!=null?r2(r.bal):null})});
  const auto=new Set();
  const ensure=(a,n)=>{if(!C.has(a)){C.set(a,{a,n:(n||a).replace(/\s+/g,' '),cat:window.custCat(n),pub:'',rg:'',wh:'',st:'active',lim:0,terms:0,ctr:'',ctrd:'',bal:null});auto.add(a)}return C.get(a)};
  /* invoices */
  const inv=[];const seen=new Set();let dup=0,calcDue=0,mism=0,neg=0;
  R('inv').forEach(r=>{const c=ensure(r.a,r.n);if(!r.no&&!r.v)return;const key=r.a+'|'+r.no;if(r.no&&seen.has(key)){dup++;return}seen.add(key);
    let v=r2(r.v),paid=r2(r.paid),rem=r.rem!=null?r2(r.rem):r2(v-paid);if(r.paid==null)paid=r2(v-rem);
    if(rem<-0.01||rem>v+0.01)neg++;else if(r.paid!=null&&r.rem!=null&&Math.abs(v-paid-rem)>Math.max(1,v*0.001))mism++;
    let due=r.due;if(!due&&r.d){due=addDays(r.d,c.terms||30);calcDue++}
    inv.push({a:r.a,no:r.no||'',d:r.d||'',p:r.p||'',q:r3(r.q),u:r.u||'',v,due:due||'',paid,rem})});
  if(dup)W.push(['w',tr(`${dup} فاتورة مكررة (الزبون ورقم الفاتورة نفساهما) أُهملت.`,`${dup} duplicate invoices (same customer and number) were skipped.`)]);
  if(calcDue)W.push(['i',tr(`${calcDue} فاتورة بلا تاريخ استحقاق؛ حُسب من تاريخ الفاتورة ومدة سداد الزبون (أو 30 يومًا).`,`${calcDue} invoices had no due date; it was computed from the invoice date and the customer's terms (or 30 days).`)]);
  if(mism)W.push(['w',tr(`${mism} فاتورة: المسدد + المتبقي لا يساوي قيمة الفاتورة (قد تكون إشعارات أو شطب). اعتُمد «المتبقي».`,`${mism} invoices: paid + remaining ≠ value (notes or write-offs?). «Remaining» was used.`)]);
  if(neg)W.push(['w',tr(`${neg} فاتورة متبقيها سالب أو أكبر من قيمتها؛ راجعها في الملف.`,`${neg} invoices have a negative remaining or one above the value; check them.`)]);
  /* the other sheets */
  const invNos=new Set(inv.map(i=>i.no).filter(Boolean));let badRef=0;
  const pay=R('pay').filter(r=>r.v).map(r=>{ensure(r.a,r.n);const refs=String(r.inv||'').split(/[,،;؛\n]+/).map(s=>s.trim()).filter(Boolean);if(inv.length&&refs.some(x=>!invNos.has(x)))badRef++;
    return {a:r.a,d:r.d||'',v:r2(r.v),m:methodOf(r.m),ref:r.ref||'',inv:refs.join('، ')}});
  if(badRef)W.push(['w',tr(`${badRef} دفعة تذكر رقم فاتورة غير موجود في ورقة الفواتير.`,`${badRef} payments name an invoice that is not in the Invoices sheet.`)]);
  const guar=R('guar').filter(r=>r.v).map(r=>{ensure(r.a);return {a:r.a,gt:gtypeOf(r.gt),ref:r.ref||'',bank:r.bank||'',v:r2(r.v),iss:r.iss||'',exp:r.exp||''}});
  const bnc=R('bnc').filter(r=>r.v).map(r=>{ensure(r.a);return {a:r.a,no:r.no||'',d:r.d||'',v:r2(r.v),bank:r.bank||'',why:r.why||'',rec:yes(r.rec)}});
  /* a payment marked «صك مرتجع؟ = نعم» stays in the collections and becomes a returned cheque (collected = payments − returned), unless the cheque is already listed */
  {const known=new Set(bnc.map(b=>b.a+'|'+b.no));let n=0;R('pay').filter(r=>r.v&&yes(r.bnc)).forEach(r=>{const no=r.ref||'';if(no&&known.has(r.a+'|'+no))return;bnc.push({a:r.a,no,d:r.d||'',v:r2(r.v),bank:'',why:'',rec:false});n++});
    if(n)W.push(['i',tr(`${n} دفعة معلّمة «صك مرتجع» في التحصيلات؛ عُدّت صكوكًا مرتجعة ولا تدخل في المحصَّل.`,`${n} payments are marked as returned cheques; they are counted as returned and not as collected.`)])}
  const act=R('act').map(r=>{ensure(r.a);return {a:r.a,d:r.d||'',t:actOf(r.t),ref:r.ref||'',v:r2(r.v),note:r.note||''}});
  const sch=R('sch').filter(r=>r.v).map(r=>{ensure(r.a);const n=Math.max(1,Math.round(r.n||1));return {a:r.a,ref:r.ref||'',d:r.d||'',v:r2(r.v),n,inst:r2(r.inst||r.v/n),first:r.first||r.d||'',every:Math.max(1,Math.round(r.every||1)),paid:r2(r.paid)}});
  const NR=R('note').filter(r=>r.v||r.sub),kindN=r=>{const t=hn(String(r.t||''));return /مخصص|مشكوك|provision/i.test(t)?'pr':/شطب|معدوم|write/i.test(t)?'wo':/مقاص|offset|netting/i.test(t)?'off':'note'};
  const prov=[...R('prov'),...NR.filter(r=>['pr','wo'].includes(kindN(r))).map(r=>({a:r.a,t:kindN(r)==='wo'?'شطب':'مخصص',d:r.d,v:r.v,inv:r.inv,note:[r.why,r.ref].filter(Boolean).join(' · ')}))].filter(r=>r.v).map(r=>{ensure(r.a);return {a:r.a,t:/شطب|معدوم|write/i.test(hn(r.t))?'wo':'pr',d:r.d||'',v:r2(r.v),inv:r.inv||'',note:r.note||''}});
  const note=NR.filter(r=>r.v&&kindN(r)==='note').map(r=>{ensure(r.a);return {a:r.a,t:noteOf(r.t),no:r.no||'',d:r.d||'',inv:r.inv||'',v:r2(r.v),why:r.why||''}});
  const off=[...R('off'),...NR.filter(r=>kindN(r)==='off')].filter(r=>r.v||r.sub).map(r=>{ensure(r.a);return {a:r.a,d:r.d||'',sub:r2(r.sub),v:r2(r.v),st:/قيد|مقترح|انتظار|pending/i.test(hn(r.st))?'pend':'done',ref:r.ref||''}});
  const age=R('age').map(r=>{ensure(r.a,r.n);const o={a:r.a};BK.forEach(k=>o[k]=r2(r[k]));return o});
  if(auto.size)W.push(['w',tr(`${auto.size} زبونًا ورد في الأوراق دون أن يكون في ورقة «الزبائن»؛ أُضيف بلا سقف ائتمان ولا فئة.`,`${auto.size} customers appear in the sheets but not in «Customers»; added without a credit limit or category.`)]);
  const cust=[...C.values()];const nc=cust.filter(c=>!c.cat).length;
  if(nc)W.push(['w',tr(`${nc} زبونًا بلا فئة معروفة (محطة وقود، جهة حكومية، شركة خاصة، طيران، صناعة، كهرباء)؛ عُدّ من القطاع الخاص.`,`${nc} customers without a known category; counted in the private sector.`)]);
  /* reconciliation: the master balance against the invoices */
  if(inv.length&&cust.some(c=>c.bal!=null)){const s=new Map();inv.forEach(i=>s.set(i.a,(s.get(i.a)||0)+i.rem));const off2=cust.filter(c=>c.bal!=null&&!near(c.bal,s.get(c.a)||0));
    if(off2.length)W.push(['w',tr(`${off2.length} زبونًا رصيده في ورقة الزبائن يختلف عن مجموع المتبقي من فواتيره؛ اعتُمدت الفواتير.`,`${off2.length} customers: the master balance differs from their invoices' remaining total; invoices were used.`)]);
    else W.push(['o',tr('أرصدة ورقة الزبائن تطابق مجموع المتبقي من الفواتير.','Master balances match the invoices.')])}
  const sheetsRead=Object.entries(F).map(([k,v])=>({k,sheet:v.sheet,rows:v.rows.length,missing:v.missing}));
  return {T:'ar',v:2,src:'tpl',from,to,cust,guar,inv,pay,bnc,act,sch,prov,note,off,age,warn:W,sheets:sheetsRead}}

/* the monthly balances statement becomes customers with a balance and the month's withdrawals / collections (no aging) */
/* The region sheets list accounts one by one, but some groups exist only on the summary sheet (the ministry-funded
   sector: electricity, desalination, army, air ambulance). A summary group none of whose categories matches a section of the
   region sheets is added from the summary, so the total equals the statement's total. */
const LG_STOP=new Set(['شركه','شركات','الشركه','الشركات','العامه','عام','حسابات','ديون','قطاع','القطاع','م','ش','غ','و','من','في','ال']);
const lgTok=s=>new Set(hn(String(s||'')).replace(/[.()«»"\-–_\/]/g,' ').split(/\s+/).map(w=>w.replace(/^(ال|وال)/,'')).filter(w=>w.length>2&&!LG_STOP.has(w)));
function ledgerList(P){const sumCats=P.cats.map(c=>Object.assign({a:c.n+(c.rg==='W'?' غ':c.rg==='E'?' ش':''),sec:c.g},c));if(!(P.custs&&P.custs.length))return {list:sumCats,added:[]};
  const covered=new Set();const secs=new Map();P.custs.forEach(c=>secs.set(c.rg+'|'+c.cur+'|'+c.sec,c));
  secs.forEach(c=>{const t=lgTok(c.sec);let best=null,bn=0;P.cats.forEach(k=>{if(k.cur!==c.cur||(k.rg&&c.rg&&k.rg!==c.rg))return;const kt=lgTok(k.n);let n=0;t.forEach(w=>{if(kt.has(w))n++});if(n>bn){bn=n;best=k}});if(best)covered.add(best.cur+'|'+best.g)});
  const added=sumCats.filter(c=>!covered.has(c.cur+'|'+c.g)&&Math.abs((c.cb||0)-(c.cc||0))>0.0005);return {list:[...P.custs,...added],added}}
/* the section of a region sheet tells the category when the account name does not */
const secCat=s=>{const t=hn(String(s||''));if(/القطاع العام|قطاع عام|ممول/.test(t))return 'gov';if(/محطات|بيع مباشر/.test(t))return 'station';if(/طيران|خطوط|جوي/.test(t))return 'aviation';if(/نفطي|خدمي|توزيع|تزويد/.test(t))return 'private';return ''};
function fromLedger(P){const LL=ledgerList(P);const list=LL.list.filter(c=>c.cur!=='F');
  const RGN={W:tr('المنطقة الغربية','Western region'),E:tr('المنطقة الشرقية','Eastern region')};
  const cust=list.map(c=>({a:String(c.a),n:c.n,cat:/ممول|وزار/.test(hn(String(c.sec||'')))?(window.custCat(c.n)==='power'?'power':'gov'):window.custCat(c.n)||window.custCat(c.sec||'')||secCat(c.sec)||secCat(c.n),pub:'',rg:RGN[c.rg]||'',wh:'',st:'active',lim:0,terms:0,ctr:'',ctrd:'',
    bal:r2((c.cb||0)-(c.cc||0)),w:r2((c.w||0)+(c.da||0)),c:r2((c.c||0)+(c.ca||0)),ob:r2((c.ob||0)-(c.oc||0)),sec:c.sec||''}));
  const from=P.prev?addDays(P.prev,1):P.asOf.slice(0,8)+'01';
  return {T:'ar',v:2,src:'ledger',from,to:P.asOf,cust,guar:[],inv:[],pay:[],bnc:[],act:[],sch:[],prov:[],note:[],off:[],age:[],warn:[...(LL.added.length?[['i',tr(`${LL.added.filter(c=>c.cur!=='F').length} فئة غير مفصّلة في أوراق المناطق أُخذت من الخلاصة (${[...new Set(LL.added.map(c=>c.g))].join('، ')})، ليطابق الإجمالي الكشف.`,`${LL.added.filter(c=>c.cur!=='F').length} categories not itemised on the region sheets were taken from the summary so the total matches the statement.`)]]:[]),['i',tr('كشف الأرصدة الشهري لا يحوي الفواتير ولا تواريخ الاستحقاق، فلا تظهر أعمار الديون ولا السقوف والضمانات. استعمل قالب بيانات الديون للصورة الكاملة.','The monthly balances statement has no invoices or due dates, so aging, limits and guarantees are not shown. Use the debts data template for the full picture.')]],sheets:[{k:'ledger',sheet:P.sheet||'',rows:cust.length,missing:[]}],file:P.file}}

/* ---------- analysis ---------- */
const by=(arr,k)=>{const m=new Map();arr.forEach(x=>{const v=x[k];if(!m.has(v))m.set(v,[]);m.get(v).push(x)});return m};
const sum=(arr,f)=>arr.reduce((s,x)=>s+((typeof f==="function"?f(x):x[f])||0),0);
/* per-customer figures, once per snapshot */
function prep(P){if(P.__c)return P.__c;const to=P.to,from=P.from;const inP=d=>d&&d>=from&&d<=to;
  const I=by(P.inv,'a'),PY=by(P.pay,'a'),G=by(P.guar,'a'),B=by(P.bnc,'a'),A=by(P.act,'a'),S=by(P.sch,'a'),PR=by(P.prov,'a'),N=by(P.note,'a'),O=by(P.off,'a'),AG=new Map(P.age.map(x=>[x.a,x]));
  const days=Math.max(1,ddiff(to,from)+1);
  const cs=P.cust.map(c0=>{const c=Object.assign({},c0);c.gov=isGov(c.cat);c.pubT=c.gov?pubType(c):'';
    const inv=I.get(c.a)||[];const ag={b0:0,b1:0,b2:0,b3:0,b4:0,b5:0};let oldest=0;
    if(inv.length){inv.forEach(i=>{if(!(i.rem>0))return;const d=i.due?ddiff(to,i.due):0;ag[bucketOf(d)]+=i.rem;if(d>oldest)oldest=d});c.bal=sum(inv,'rem');c.aged=true}
    else if(AG.has(c.a)){const x=AG.get(c.a);BK.forEach(k=>ag[k]=x[k]||0);c.bal=sum(BK,k=>ag[k]);c.aged=true;oldest=ag.b5?181:ag.b4?91:ag.b3?61:ag.b2?31:ag.b1?1:0}
    else{c.bal=c.bal||0;c.aged=false}
    c.ag=ag;c.over=ag.b1+ag.b2+ag.b3+ag.b4+ag.b5;c.o90=ag.b4+ag.b5;c.oldest=oldest;
    c.nInv=inv.length;c.sales=inv.length?sum(inv.filter(i=>inP(i.d)),'v'):(c.w||0);
    const pays=PY.get(c.a)||[];const bn=B.get(c.a)||[];
    c.paid=pays.length?sum(pays.filter(p=>inP(p.d)),'v'):(c.c||0);c.bounced=sum(bn.filter(b=>inP(b.d)&&!b.rec),'v');c.coll=Math.max(0,c.paid-c.bounced);
    c.lastPay=pays.map(p=>p.d).filter(Boolean).sort().pop()||'';c.pays=pays;c.bnc=bn;
    const gs=G.get(c.a)||[];c.gs=gs;c.gv=sum(gs.filter(g=>!g.exp||g.exp>=to),'v');c.gExp=gs.filter(g=>g.exp&&g.exp<to);c.gSoon=gs.filter(g=>g.exp&&g.exp>=to&&ddiff(g.exp,to)<=60);
    c.acts=A.get(c.a)||[];c.legal=c.st==='legal'||c.acts.some(x=>x.t==='legal');
    c.sch=(S.get(c.a)||[]).map(schState.bind(null,to));
    const pr=PR.get(c.a)||[];c.prov=sum(pr.filter(x=>x.t==='pr'),'v');c.wo=sum(pr.filter(x=>x.t==='wo'),'v');c.woP=sum(pr.filter(x=>x.t==='wo'&&inP(x.d)),'v');c.prs=pr;
    const nt=N.get(c.a)||[];c.cn=sum(nt.filter(x=>x.t==='cr'&&inP(x.d)),'v');c.dn=sum(nt.filter(x=>x.t==='dr'&&inP(x.d)),'v');c.notes=nt;
    c.offs=O.get(c.a)||[];c.inv=inv;
    c.util=c.lim>0?c.bal*100/c.lim:null;c.overLim=c.lim>0&&c.bal>c.lim;c.cov=P.guar.length&&c.bal>0?c.gv*100/c.bal:null;
    c.dso=c.sales>0&&c.bal>0?c.bal/c.sales*days:null;
    c.step=nextStep(c);c.risk=c.ag.b5*4+c.ag.b4*3+c.ag.b3*2+c.ag.b2+c.ag.b1*0.3+(c.overLim?c.bal-c.lim:0)+(c.st==='default'?c.bal*0.5:0);
    return c});
  return P.__c={cs,days,map:new Map(cs.map(c=>[c.a,c]))}}
function schState(to,s){const dueN=!s.first||s.first>to?0:Math.min(s.n,Math.floor(monthsBetween(s.first,to)/s.every)+1);const dueAmt=Math.min(s.v,dueN*s.inst);
  const comp=dueAmt>0?Math.min(100,s.paid*100/dueAmt):100;const st=s.paid>=s.v-0.01?'done':comp>=95?'ok':comp>=50?'late':'broken';
  const next=dueN<s.n?addMonths(s.first,dueN*s.every):'';return Object.assign({},s,{dueN,dueAmt,comp,st,rem:Math.max(0,s.v-s.paid),next})}
function monthsBetween(a,b){let m=(+b.slice(0,4)-+a.slice(0,4))*12+(+b.slice(5,7)-+a.slice(5,7));if(+b.slice(8,10)<+a.slice(8,10))m--;return Math.max(0,m)}
const SCH={done:['مكتملة','Completed','ok'],ok:['ملتزم','On track','ok'],late:['متأخر','Behind','warn'],broken:['متعثر','Broken','crit']};
/* the next collection step, from the oldest overdue amount and the account state */
function nextStep(c){if(c.bal<=0)return '';
  if(c.legal)return tr('متابعة القضية لدى الشؤون القانونية','Follow the case with Legal Affairs');
  if(c.gov)return c.over>0?tr('مطالبة الجهة ووزارة المالية، أو مقاصّة مع الدعم','Claim from the body and the Ministry of Finance, or offset against the subsidy'):'';
  const o=c.oldest;
  if(o>180)return tr('إحالة للشؤون القانونية وتكوين مخصص','Refer to Legal Affairs and provide for it');
  if(o>90)return c.gv>0?tr('إنذار نهائي وتسييل الضمان','Final warning and call the guarantee'):tr('إنذار نهائي وإيقاف التوريد','Final warning and stop supply');
  if(o>60)return tr('إنذار كتابي وإيقاف البيع الآجل','Written warning and stop credit sales');
  if(c.overLim)return tr('إيقاف التوريد الآجل حتى النزول تحت السقف','Stop credit supply until below the limit');
  if(o>30)return tr('مطالبة كتابية','Written claim');
  if(o>0)return tr('تذكير هاتفي','Phone reminder');
  return ''}
/* filtered portfolio figures */
function analyze(P,u){const X=prep(P);const q=norm(u.q||'');
  const list=X.cs.filter(c=>(u.sec==='all'||(u.sec==='gov'?c.gov:!c.gov))&&(!u.cat||c.cat===u.cat)&&(!u.rg||c.rg===u.rg)&&(!u.wh||c.wh===u.wh)&&(!u.st||c.st===u.st)
    &&(!q||norm(c.n+' '+c.a+' '+(c.ctr||'')).includes(q)));
  const A={list,days:X.days};const pos=list.filter(c=>c.bal>0);
  A.total=sum(pos,'bal');A.credit=-sum(list.filter(c=>c.bal<0),'bal');A.n=list.length;A.nDebt=pos.length;
  A.aged=list.some(c=>c.aged);A.ag={};BK.forEach(k=>A.ag[k]=sum(list,c=>c.ag[k]));A.unaged=sum(pos.filter(c=>!c.aged),'bal');
  A.over=sum(list,'over');A.o90=sum(list,'o90');A.overR=A.total&&A.aged?A.over*100/A.total:null;
  A.hasSales=P.inv.length>0||P.src==='ledger';A.sales=sum(list,'sales');A.dso=A.hasSales&&A.sales>0?A.total/A.sales*X.days:null;
  A.hasColl=P.pay.length>0||P.src==='ledger';A.paid=sum(list,'paid');A.bounced=sum(list,'bounced');A.coll=sum(list,'coll');
  A.due=A.coll+A.over;/* the monthly statement has no due dates, so what was due (and the ratio) is unknown */A.collR=A.hasColl&&A.due>0&&P.src!=='ledger'?A.coll*100/A.due:null;
  const lim=list.filter(c=>c.lim>0);A.hasLim=lim.length>0;A.lim=sum(lim,'lim');A.limBal=sum(lim,c=>Math.max(0,c.bal));A.util=A.lim?A.limBal*100/A.lim:null;
  A.overLim=list.filter(c=>c.overLim).sort((a,b)=>(b.bal-b.lim)-(a.bal-a.lim));A.overLimV=sum(A.overLim,c=>c.bal-c.lim);
  A.hasG=list.some(c=>c.gs.length>0);A.gv=sum(list,'gv');A.cov=A.total?A.gv*100/A.total:null;A.gCovered=sum(pos,c=>Math.min(c.bal,c.gv));A.unsecured=A.total-A.gCovered;
  A.gExp=list.flatMap(c=>c.gExp.map(g=>Object.assign({c},g)));A.gSoon=list.flatMap(c=>c.gSoon.map(g=>Object.assign({c},g))).sort((a,b)=>a.exp.localeCompare(b.exp));
  A.top=pos.slice().sort((a,b)=>b.bal-a.bal).slice(0,10);A.top10=sum(A.top,'bal');
  A.byCat=[...by(pos,'cat').entries()].map(([k,l])=>({k,l,bal:sum(l,'bal'),over:sum(l,'over'),n:l.length})).sort((a,b)=>b.bal-a.bal);
  A.byRg=[...by(pos,'rg').entries()].map(([k,l])=>({k,l,bal:sum(l,'bal'),over:sum(l,'over'),n:l.length})).sort((a,b)=>b.bal-a.bal);
  A.byWh=[...by(pos,'wh').entries()].filter(([k])=>k).map(([k,l])=>({k,bal:sum(l,'bal'),over:sum(l,'over'),n:l.length})).sort((a,b)=>b.bal-a.bal);
  A.bySt=Object.keys(STATUS).map(k=>{const l=list.filter(c=>c.st===k);return {k,n:l.length,bal:sum(l,c=>Math.max(0,c.bal))}});
  const inP=d=>d&&d>=P.from&&d<=P.to;const ids=new Set(list.map(c=>c.a));const mine=x=>ids.has(x.a);
  A.byM=Object.keys(METHOD).map(k=>({k,v:sum(P.pay.filter(p=>mine(p)&&inP(p.d)&&p.m===k),'v')}));
  A.bnc=P.bnc.filter(b=>mine(b)&&inP(b.d));A.bncOpen=P.bnc.filter(b=>mine(b)&&!b.rec);
  A.acts=P.act.filter(x=>mine(x)&&inP(x.d));A.actN={};Object.keys(ACT).forEach(k=>A.actN[k]=A.acts.filter(x=>x.t===k).length);
  A.legal=list.filter(c=>c.legal&&c.bal>0);A.legalV=sum(A.legal,'bal');
  A.sch=list.flatMap(c=>c.sch.map(s=>Object.assign({c},s)));A.schV=sum(A.sch,'v');A.schRem=sum(A.sch,'rem');A.schDue=sum(A.sch,'dueAmt');A.schPaid=sum(A.sch,s=>Math.min(s.paid,s.dueAmt||s.paid));
  A.schComp=A.schDue>0?Math.min(100,A.schPaid*100/A.schDue):null;
  A.prov=sum(list,'prov');A.wo=sum(list,'wo');A.woP=sum(list,'woP');A.cn=sum(list,'cn');A.dn=sum(list,'dn');A.net=A.total-A.prov;A.provCov=A.o90?A.prov*100/A.o90:null;
  A.offs=list.flatMap(c=>c.offs.map(o=>Object.assign({c},o)));A.offDone=sum(A.offs.filter(o=>o.st==='done'),'v');A.offPend=sum(A.offs.filter(o=>o.st==='pend'),'v');A.sub=sum(A.offs,'sub');
  A.pub=Object.keys(PUB).map(k=>{const l=pos.filter(c=>c.gov&&c.pubT===k);return {k,n:l.length,bal:sum(l,'bal'),over:sum(l,'over')}}).filter(x=>x.n);
  A.prio=pos.filter(c=>c.step).sort((a,b)=>b.risk-a.risk||b.bal-a.bal);
  return A}
/* both sectors side by side, always from the whole portfolio (filters other than the sector still apply) */
function sectors(P,u){const g=analyze(P,Object.assign({},u,{sec:'gov'})),p=analyze(P,Object.assign({},u,{sec:'priv'}));return {g,p}}

/* ---------- styles (on top of the .sm-* classes of sms055.js) ---------- */
(function(){const st=document.createElement('style');st.textContent=`
.db-page .sm-hero{--hero-a:#3d0a18;--hero-b:#9b1c31;--hero-c:#5e1022}
.db-page{--b0:#1f9d6b;--b1:#7fb03a;--b2:#d9a520;--b3:#e07a2e;--b4:#d03f3f;--b5:#7d1428;--gv:#2a62c9;--pv:#9b1c31}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .db-page{--b0:#2bb37e;--b1:#93c24d;--b2:#e5b43a;--b3:#ec8c45;--b4:#e45a5a;--b5:#c4475f;--gv:#5b8ef0;--pv:#e0566f}}
:root[data-theme="dark"] .db-page{--b0:#2bb37e;--b1:#93c24d;--b2:#e5b43a;--b3:#ec8c45;--b4:#e45a5a;--b5:#c4475f;--gv:#5b8ef0;--pv:#e0566f}
.db-page .k-b0{background:var(--b0)}.db-page .k-b1{background:var(--b1)}.db-page .k-b2{background:var(--b2)}.db-page .k-b3{background:var(--b3)}.db-page .k-b4{background:var(--b4)}.db-page .k-b5{background:var(--b5)}
.db-page .k-gv{background:var(--gv)}.db-page .k-pv{background:var(--pv)}.db-page .k-ok{background:var(--ok)}.db-page .k-mu{background:var(--line)}
.db-sub{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12.5px;color:rgba(255,255,255,.85);margin-top:10px}.db-sub span{display:inline-flex;gap:6px;align-items:center}.db-sub b{color:#fff;font-variant-numeric:tabular-nums}
.db-sub i{width:9px;height:9px;border-radius:3px;display:inline-block;background:#fff}.db-sub i.g{background:#9cc0ff}.db-sub i.p{background:#ffb3c0}
.db-seg{display:inline-flex;border:1px solid var(--line);border-radius:10px;overflow:hidden;background:var(--surface);height:36px}
.db-seg button{border:0;background:none;color:var(--ink);font:inherit;font-size:12.5px;padding:0 14px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.db-seg button+button{border-inline-start:1px solid var(--line)}
.db-seg button[aria-pressed="true"]{background:var(--accent);color:var(--on-accent,#fff);font-weight:600}
.db-seg i{width:8px;height:8px;border-radius:50%;display:inline-block}
.db-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:14px}
.db-kpi{background:var(--surface);border:1px solid var(--line);border-radius:var(--r,14px);padding:13px 15px 12px;display:grid;gap:3px;position:relative;overflow:hidden;min-width:0}
.db-kpi::before{content:"";position:absolute;inset-inline-start:0;top:0;bottom:0;width:4px;background:var(--line)}
.db-kpi.ok::before{background:var(--ok)}.db-kpi.warn::before{background:var(--warn)}.db-kpi.crit::before{background:var(--crit)}
.db-kpi .l{font-size:12.5px;color:var(--muted);font-weight:600}
.db-kpi .v{font-family:var(--f-display);font-size:25px;line-height:1.25;font-variant-numeric:tabular-nums;color:var(--ink)}
.db-kpi .v small{font-size:13px;color:var(--muted);font-family:inherit;margin-inline-start:4px}
.db-kpi .f{font-size:11.5px;color:var(--muted);border-top:1px dashed var(--line);padding-top:5px;margin-top:3px;line-height:1.6}
.db-kpi .s{font-size:12px;color:var(--ink)}
.db-kpi .bar{height:6px;border-radius:99px;background:var(--line);overflow:hidden;margin-top:4px}.db-kpi .bar span{display:block;height:100%;border-radius:99px;background:var(--accent)}
.db-kpi.crit .bar span{background:var(--crit)}.db-kpi.warn .bar span{background:var(--warn)}.db-kpi.ok .bar span{background:var(--ok)}
.db-age{display:grid;gap:10px}
.db-age .sm-stack{height:22px}
.db-ageT td .sw{width:10px;height:10px;border-radius:3px;display:inline-block;margin-inline-end:6px;vertical-align:-1px}
.db-sector{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.db-sector>div{border:1px solid var(--line);border-radius:12px;padding:10px 12px;display:grid;gap:6px;border-top:4px solid var(--pv)}
.db-sector>div.g{border-top-color:var(--gv)}
.db-sector h4{margin:0;font-size:13.5px}.db-sector dl{margin:0;display:grid;grid-template-columns:1fr auto;gap:4px 10px;font-size:12.5px}
.db-sector dt{color:var(--muted)}.db-sector dd{margin:0;font-variant-numeric:tabular-nums;font-weight:600;text-align:end}
.db-top{display:grid;gap:7px}
.db-top .r{display:grid;grid-template-columns:22px minmax(110px,32%) 1fr 76px;gap:8px;align-items:center;font-size:12.5px;cursor:pointer;border-radius:8px;padding:2px 4px}
.db-top .r:hover,.db-top .r:focus-visible{background:var(--surface-2,var(--accent-soft));outline:none}
.db-top .rk{width:22px;height:22px;border-radius:50%;display:grid;place-items:center;font-size:11px;font-weight:700;background:var(--surface-2,var(--accent-soft))}
.db-top .nm{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.db-top .nm small{color:var(--muted)}
.db-top .tk{height:12px;display:flex;border-radius:4px;overflow:hidden;background:color-mix(in srgb,var(--line) 50%,transparent)}
.db-top .tk span{display:block;height:100%}.db-top .vv{text-align:end;font-variant-numeric:tabular-nums;font-weight:600}
.db-mini{display:grid;grid-template-columns:repeat(auto-fit,minmax(128px,1fr));gap:8px;margin:2px 0 10px}
.db-mini>div{border:1px solid var(--line);border-radius:10px;padding:8px 10px;display:grid;gap:1px}
.db-mini b{font-size:17px;font-variant-numeric:tabular-nums;font-family:var(--f-display)}.db-mini span{font-size:11.5px;color:var(--muted)}
.db-mini .crit b{color:var(--crit)}.db-mini .warn b{color:var(--warn)}.db-mini .ok b{color:var(--ok)}
.db-util{display:inline-grid;grid-template-columns:46px 1fr;gap:6px;align-items:center;min-width:110px}
.db-util i{display:block;height:6px;border-radius:9px;background:var(--line);overflow:hidden}.db-util i b{display:block;height:100%;background:var(--ok)}
.db-util.warn i b{background:var(--warn)}.db-util.crit i b{background:var(--crit)}
#dlg.db-wide{width:min(1100px,calc(100% - 24px))}
.db-h4{margin:10px 0 4px;font-size:12.5px;color:var(--muted);font-weight:600}
.db-note{font-size:12px;color:var(--muted);margin:6px 0 0;line-height:1.7}
.db-formula{font-size:12px;color:var(--muted);margin:-4px 0 10px}
.db-dl{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:6px 14px;font-size:12.5px;margin:4px 0 10px}
.db-dl div{display:flex;justify-content:space-between;gap:8px;border-bottom:1px dashed var(--line);padding:3px 0}.db-dl span{color:var(--muted)}
.db-dlg h4{margin:14px 0 6px;font-size:13.5px}
.db-tpl ol{margin:6px 0;padding-inline-start:20px;line-height:1.9;font-size:13px}
.db-chk{list-style:none;margin:6px 0;padding:0;display:grid;gap:4px;font-size:12.5px}.db-chk li{display:flex;gap:6px;align-items:flex-start;line-height:1.7}
.db-chk li::before{content:"•";font-weight:700;color:var(--muted)}.db-chk li.o::before{content:"✓";color:var(--ok)}.db-chk li.w::before{content:"!";color:var(--warn)}.db-chk li.i::before{content:"i";color:var(--info,var(--accent))}
.sm-tbl tr[data-db-acct]{cursor:pointer}
.sm-tbl th[data-db-sort]{cursor:pointer;white-space:nowrap}.sm-tbl th[data-db-sort][aria-sort]::after{content:" ▾";font-size:10px}.sm-tbl th[data-db-sort][aria-sort="ascending"]::after{content:" ▴"}
.sm-tbl td.crit{color:var(--crit);font-weight:600}.sm-tbl td.warn{color:var(--warn)}
.db-tag{display:inline-block;font-size:11px;padding:1px 7px;border-radius:999px;border:1px solid var(--line);color:var(--muted);margin-inline-start:4px;white-space:nowrap}
.db-tag.g{border-color:var(--gv);color:var(--gv)}
@media (max-width:1100px){.db-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:720px){.db-kpis{grid-template-columns:1fr 1fr;gap:8px}.db-kpi .v{font-size:20px}.db-sector{grid-template-columns:1fr}.db-top .r{grid-template-columns:22px minmax(80px,36%) 1fr 62px}.db-seg button{padding:0 9px}}
@media print{.db-seg,.sm-bar{display:none}.db-kpi,.panel{break-inside:avoid}}
`;document.head.appendChild(st)})();

/* ---------- reading a file ---------- */
async function parseFile(file){const sheets=await readGrids(file);
  let P=parseTemplate(sheets,file.name);
  if(!P){const lg=parseLedger(sheets,file.name);if(lg)P=fromLedger(lg)}
  if(!P)throw new Error(tr('لم نتعرّف على الملف. استعمل «قالب بيانات الديون» (زر «تنزيل القالب»): ورقة الزبائن مع ورقة الفواتير أو ورقة أعمار الديون على الأقل. ويُقرأ أيضًا كشف «أرصدة حسابات المدينون والدائنون التجاريون» الشهري.',
    'The file was not recognised. Use the debts data template («Download template»): the Customers sheet with the Invoices sheet or the Aging sheet at least. The monthly «trade debtors and creditors balances» statement is read too.'));
  P.file=file.name;return P}

/* ---------- storage ---------- */
const ST={loaded:false,loading:null,err:'',batches:[],snaps:[],local:null,pending:null,canImport:false};
async function load(force){if(ST.loading&&!force)return ST.loading;
  ST.loading=(async()=>{const FS=window.__fs&&window.__fs();if(!FS){ST.loaded=true;return}
    try{const q=await FS.collection(COL).get();const docs=[];q.forEach(d=>docs.push(Object.assign({id:d.id},d.data())));
      const parts={};docs.filter(d=>d.kind==='part').forEach(p=>{(parts[p.batch]=parts[p.batch]||[]).push(p)});
      const batches=docs.filter(d=>d.kind==='batch').sort((a,b)=>(a.to||'').localeCompare(b.to||''));const snaps=[];
      batches.forEach(b=>{const ps=(parts[b.id]||[]).sort((x,y)=>x.n-y.n);b.ok=ps.length===b.parts;if(!b.ok)return;
        try{let o=JSON.parse(ps.map(p=>p.d).join(''));if(o.T==='ledger')o=fromLedger(o);if(o.T!=='ar')return;o.b=b.id;o.file=b.file;snaps.push(o)}catch(e){b.ok=false}});
      ST.batches=batches;ST.snaps=snaps;ST.err=''}
    catch(e){ST.err=e&&e.code==='permission-denied'?'denied':(e&&e.message)||'error';ST.batches=[];ST.snaps=[]}
    if(ST.err==='denied'&&window.__denyPrep)await window.__denyPrep();
    ST.loaded=true;
    try{const u=await recMyUnit();const m=me();ST.canImport=m.role==='admin'||(!!S.canWrite&&m.role!=='viewer'&&(u===UNIT||u==='gm'))}catch(e){ST.canImport=false}
  })();return ST.loading}
const overlaps=P=>ST.batches.filter(b=>(b.sheet==='ar'||b.sheet==='ledger')&&b.to===P.to);
const nRows=P=>['cust','guar','inv','pay','bnc','act','sch','prov','note','off','age'].reduce((s,k)=>s+(P[k]||[]).length,0);
async function saveOne(P,replace){const FS=window.__fs(),m=me();const bid='b'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
  const s=JSON.stringify(P,(k,v)=>k==='__c'||k==='file'||k==='b'?undefined:v);const chunks=[];for(let i=0;i<s.length;i+=CHUNK)chunks.push(s.slice(i,i+CHUNK));
  if(chunks.length>200)throw new Error(tr('الملف أكبر من الحد المسموح للحفظ.','The file is larger than the save limit.'));
  for(let i=0;i<chunks.length;i+=4){const b=FS.batch();chunks.slice(i,i+4).forEach((d,j)=>b.set(FS.collection(COL).doc(bid+'_p'+(i+j)),{kind:'part',batch:bid,n:i+j,d,createdBy:m.email}));await b.commit()}
  const old=replace?overlaps(P):[];const X=prep(P);
  await FS.collection(COL).doc(bid).set({kind:'batch',file:String(P.file||'').slice(0,200),sheet:'ar',from:P.from,to:P.to,rows:nRows(P),parts:chunks.length,
    value:r2(sum(X.cs,c=>Math.max(0,c.bal))),returns:r2(sum(X.cs,'coll')),createdBy:m.email,by:String(m.name||m.email||'').slice(0,120),at:new Date().toISOString()});
  for(const o of old)await delBatch(o,true);
  try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'استيراد بيانات الديون',col:COL,id:bid,summary:`${P.file||''} · ${P.from} – ${P.to} · ${P.cust.length} زبون · ${P.inv.length} فاتورة`})}catch(e){}}
async function delBatch(b,quiet){const FS=window.__fs();const ops=[];for(let i=0;i<(b.parts||0);i++)ops.push(FS.collection(COL).doc(b.id+'_p'+i));
  for(let i=0;i<ops.length;i+=400){const bt=FS.batch();ops.slice(i,i+400).forEach(r=>bt.delete(r));await bt.commit()}
  await FS.collection(COL).doc(b.id).delete();
  if(!quiet){const m=me();try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'حذف ملف الديون',col:COL,id:b.id,summary:`${b.file||''} · ${b.to}`})}catch(e){}}}

/* ---------- page ---------- */
I.debts='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 3.5h9.5L19 8v12.5H5z"/><path d="M14.5 3.5V8H19"/><path d="M8 17.5v-2.5M11 17.5v-4.5M14 17.5v-6.5" stroke-width="2"/><circle cx="17.6" cy="17.6" r="3.6" fill="var(--surface,#fff)"/><path d="M17.6 15.8v1.9l1.2.8"/></svg>';
function UI(){return S.ui.debts||(S.ui.debts={at:'',sec:'priv',cat:'',rg:'',wh:'',st:'',q:'',sort:'bal',dir:'desc',lim:25})}
function snaps(){return ST.local?[ST.local]:ST.snaps.slice().sort((a,b)=>a.to.localeCompare(b.to))}
function cur(){const L1=snaps();if(!L1.length)return null;const u=UI();return L1.find(s=>s.to===u.at)||L1[L1.length-1]}
const U=()=>tr('د.ل','LYD');
const SECN={priv:['القطاع الخاص','Private sector'],gov:['الجهات الحكومية','Government bodies'],all:['كل القطاعات','All sectors']};
const secName=k=>tr(SECN[k][0],SECN[k][1]);
const stPill=k=>{const s=STATUS[k]||STATUS.active;return `<span class="pill p-${s[2]}">${esc(tr(s[0],s[1]))}</span>`};
const tone=(v,warn,crit,higherBad=true)=>v==null?'':higherBad?(v>=crit?'crit':v>=warn?'warn':'ok'):(v<=crit?'crit':v<=warn?'warn':'ok');
function kpi(l,v,f,s,cls,barPct){return `<div class="db-kpi ${cls||''}"><span class="l">${esc(l)}</span><span class="v">${v}</span>${s?`<span class="s">${s}</span>`:''}${barPct!=null?`<span class="bar"><span style="width:${Math.max(0,Math.min(100,barPct))}%"></span></span>`:''}<span class="f">${esc(f)}</span></div>`}
function utilCell(c){if(c.util==null)return '<span class="muted">—</span>';const t=c.util>100?'crit':c.util>=80?'warn':'';return `<span class="db-util ${t}"><span class="num">${Math.round(c.util)}%</span><i><b style="width:${Math.min(100,c.util)}%"></b></i></span>`}
function heroHtml(sub,v,vl,extra){return `<div class="sm-hero"><div><span class="sm-eyebrow">${esc(tr(DEPT[0],DEPT[1]))}</span>
  <h2>${esc(tr('متابعة المبيعات الآجلة والديون المتأخرة','Credit sales and overdue debts'))}</h2><p>${sub}</p>${extra||''}</div>${v==null?'':`<div class="sm-hero-v"><b>${v}</b><span>${esc(vl)}</span></div>`}</div>`}
function toolbar(){const has=snaps().length;return `<div class="toolbar sm-acts">
  ${ST.canImport||!ST.snaps.length||ST.err?`<button class="btn primary" data-dbt="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button>`:''}
  <button class="btn" data-dbt="template">${I.down}${esc(tr('تنزيل القالب','Download template'))}</button>
  ${has?`<button class="btn" data-dbt="export">${I.down}${esc(tr('تصدير التحليل Excel','Export analysis (Excel)'))}</button><button class="btn" data-dbt="print">${esc(tr('طباعة / PDF','Print / PDF'))}</button>`:''}
  ${ST.batches.length?`<button class="btn ghost" data-dbt="files">${esc(tr('الملفات المستوردة','Imported files'))} <span class="num">${ST.batches.length}</span></button>`:''}
  ${ST.local?`<button class="btn ghost" data-dbt="drop-local">${esc(tr('إغلاق العرض المؤقت','Close temporary view'))}</button>`:''}
  <input type="file" id="db-file" accept="${esc(ACCEPT())}" hidden></div>`}
function emptyView(head){const denied=ST.err==='denied';
  return `<div class="db-page">`+head+heroHtml(esc(tr('من بيانات شركة البريقة لكل زبون: الذمم وأعمارها، والتحصيلات ونسبتها، ومتوسط فترة التحصيل، واستخدام سقوف الائتمان وتغطية الضمانات، وأكبر المدينين، وإجراءات التحصيل والتسويات، مع فصل ديون الجهات الحكومية عن القطاع الخاص.','From the company data for every customer: receivables and their aging, collections and the collection ratio, DSO, credit-limit use and guarantee cover, the largest debtors and the collection actions, with government debts kept apart from the private sector.')),null)+
  `<div class="panel sm-empty db-tpl">${I.debts.replace('<svg','<svg style="width:52px;height:52px;color:var(--accent)"')}<h3>${esc(denied?tr('هذه اللوحة لإدارة الشؤون التجارية والتحصيلات','This dashboard belongs to Commercial Affairs & Collections'):tr('ابدأ بتعبئة قالب بيانات الديون','Start by filling the debts data template'))}</h3>
   <p>${esc(denied?tr('لا تملك صلاحية الاطلاع على البيانات المحفوظة. يمكنك استيراد ملف وتحليله مؤقتًا على جهازك دون حفظه.','You cannot view the saved data. You can still import a file and analyse it on this device without saving it.')+(window.__denyWhy?' '+window.__denyWhy():''):tr('نزّل القالب، واملأ أوراقه من حسابات الشركة (يكفي «الفترة» و«الزبائن» و«الفواتير» و«التحصيلات» للبدء، والباقي يُضاف متى توفّر)، ثم استورده هنا. يُقرأ أيضًا كشف أرصدة المدينين الشهري.','Download the template, fill its sheets from the company accounts (Period, Customers, Invoices and Collections are enough to start), then import it here. The monthly debtors balances statement is read too.'))}</p>
   <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:center"><button class="btn primary" data-dbt="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button><button class="btn" data-dbt="template-help">${esc(tr('ما المطلوب في القالب؟','What goes in the template?'))}</button></div></div></div>`}

function view(){
  if(!ST.loaded){load().then(()=>{if(S.view===ID)render()});return `<div class="db-page">`+heroHtml(esc(tr('جارٍ تحميل البيانات…','Loading data…')),null)+'<div class="panel muted">'+esc(tr('جارٍ التحميل…','Loading…'))+'</div></div>'}
  const head=`<div class="page-head" style="margin-bottom:8px"><div><h2 style="margin:0">${esc(tr('لوحة الديون','Debts dashboard'))}</h2></div>${toolbar()}</div>`;
  const P=cur();if(!P)return emptyView(head);
  const u=UI();u.at=P.to;const X=prep(P);const A=analyze(P,u);const SC=sectors(P,Object.assign({},u,{q:''}));
  const banner=ST.local?`<div class="banner info" style="margin-bottom:10px">${esc(tr(`عرض مؤقت للملف «${ST.local.file}» على هذا الجهاز فقط، ولم يُحفظ في المنظومة.`,`Temporary view of “${ST.local.file}” on this device only; not saved.`))}</div>`:'';
  const all=SC.g.total+SC.p.total;const allOver=SC.g.over+SC.p.over;
  const sub=esc(tr(`الفترة ${fdi(P.from)} – ${fdi(P.to)} · الأعمار في ${fdi(P.to)} · ${fn(X.cs.length)} زبونًا · ${fn(P.inv.length)} فاتورة · ${fn(P.pay.length)} دفعة`,`Period ${fdi(P.from)} – ${fdi(P.to)} · aging at ${fdi(P.to)} · ${fn(X.cs.length)} customers · ${fn(P.inv.length)} invoices · ${fn(P.pay.length)} payments`));
  const extra=`<div class="db-sub"><span><i class="p"></i>${esc(secName('priv'))} <b>${big(SC.p.total)}</b> (${pct(SC.p.total,all)}%)</span><span><i class="g"></i>${esc(secName('gov'))} <b>${big(SC.g.total)}</b> (${pct(SC.g.total,all)}%)</span>${SC.p.aged||SC.g.aged?`<span>${esc(tr('متأخر','Overdue'))} <b>${big(allOver)}</b> (${pct(allOver,all)}%)</span>`:''}</div>`;
  const L1=snaps();const cats=uniqK(X.cs.map(c=>c.cat)),rgs=uniq(X.cs.map(c=>c.rg)),whs=uniq(X.cs.map(c=>c.wh));
  const filtered=u.cat||u.rg||u.wh||u.st||u.q;
  const bar=`<div class="sm-bar">
    <label class="an-f">${esc(tr('الفترة','Period'))}<select id="db-at">${L1.slice().reverse().map(s=>`<option value="${esc(s.to)}" ${s.to===P.to?'selected':''}>${esc(mLabel(s.to))} (${fdi(s.from)} – ${fdi(s.to)})</option>`).join('')}</select></label>
    <div class="an-f">${esc(tr('القطاع','Sector'))}<div class="db-seg" role="group">${['priv','gov','all'].map(k=>`<button type="button" data-db-sec="${k}" aria-pressed="${u.sec===k}">${k!=='all'?`<i class="k-${k==='gov'?'gv':'pv'}"></i>`:''}${esc(secName(k))}</button>`).join('')}</div></div>
    ${cats.length>1?`<label class="an-f">${esc(tr('الفئة','Category'))}<select id="db-cat"><option value="">${esc(tr('كل الفئات','All categories'))}</option>${cats.map(k=>`<option value="${esc(k)}" ${u.cat===k?'selected':''}>${esc(catName(k))}</option>`).join('')}</select></label>`:''}
    ${rgs.length>1?`<label class="an-f">${esc(tr('المنطقة','Region'))}<select id="db-rg"><option value="">${esc(tr('كل المناطق','All regions'))}</option>${rgs.map(k=>`<option ${u.rg===k?'selected':''}>${esc(k)}</option>`).join('')}</select></label>`:''}
    ${whs.length>1?`<label class="an-f">${esc(tr('المستودع','Warehouse'))}<select id="db-wh"><option value="">${esc(tr('كل المستودعات','All warehouses'))}</option>${whs.map(k=>`<option ${u.wh===k?'selected':''}>${esc(k)}</option>`).join('')}</select></label>`:''}
    <label class="an-f">${esc(tr('حالة الحساب','Account status'))}<select id="db-st"><option value="">${esc(tr('كل الحالات','All'))}</option>${Object.keys(STATUS).map(k=>`<option value="${k}" ${u.st===k?'selected':''}>${esc(tr(STATUS[k][0],STATUS[k][1]))}</option>`).join('')}</select></label>
    <label class="an-f" style="flex:1;min-width:170px">${esc(tr('بحث','Search'))}<input type="search" id="db-q" value="${esc(u.q||'')}" placeholder="${esc(tr('اسم الزبون أو كوده أو رقم العقد','Customer name, code or contract'))}"></label>
    ${filtered?`<button class="btn sm ghost" data-dbt="clear">${esc(tr('مسح التصفية','Clear filters'))}</button>`:''}</div>`;
  const mix=u.sec==='all'&&SC.g.total&&SC.p.total?`<div class="banner" style="margin-bottom:12px">${esc(tr('تعرض المؤشرات القطاعين معًا. ديون الجهات الحكومية تتصرّف بشكل مختلف (الدعم والمقاصّة)، وخلطها مع القطاع الخاص يشوّه نسب التأخر والتحصيل؛ اقرأ كل قطاع منفصلًا.','These figures mix both sectors. Government debts behave differently (subsidy and offsets); mixing them with the private sector distorts the overdue and collection ratios. Read each sector on its own.'))}</div>`:'';
  if(!A.n)return `<div class="db-page">`+head+banner+heroHtml(sub,big(all),tr(`إجمالي الذمم · ${U()}`,`Total receivables · ${U()}`),extra)+bar+`<div class="panel empty"><b>${esc(tr('لا يوجد زبائن','No customers'))}</b> ${esc(tr('ضمن التصفية الحالية.','for the current filter.'))}</div></div>`;
  return `<div class="db-page">`+head+banner+heroHtml(sub,big(all),tr(`إجمالي الذمم · ${U()}`,`Total receivables · ${U()}`),extra)+bar+mix+kpiRow(A,P)+
   `<div class="an-grid sm-root">${agingPanel(A)}${sectorPanel(SC,P)}${topPanel(A)}${limitPanel(A)}${guarPanel(A)}${collPanel(A,P)}${mixPanel(A)}${actionsPanel(A)}${govPanel(A,u)}${trendPanel(L1,u)}${readingPanel(A,SC,P,u)}${prioPanel(A)}${tablePanel(A,u)}</div></div>`}
function uniqK(arr){return [...new Set(arr.filter(Boolean))].sort((a,b)=>catName(a).localeCompare(catName(b),'ar'))}

function kpiRow(A,P){const days=A.days;
  return `<div class="db-kpis">
   ${kpi(tr('إجمالي الذمم','Total receivables'),`${big(A.total)}<small>${esc(U())}</small>`,tr('مجموع أرصدة الزبائن','Sum of customer balances'),esc(tr(`${fn(A.nDebt)} زبونًا مدينًا`,`${fn(A.nDebt)} customers owe`))+(A.credit?` · <span class="muted">${esc(tr('أرصدة دائنة','credit balances'))} ${big(A.credit)}</span>`:''),'')}
   ${kpi(tr('نسبة الديون المتأخرة','Overdue ratio'),A.overR==null?'—':P1(A.overR),tr('المتأخر ÷ إجمالي الذمم','Overdue ÷ total receivables'),A.overR==null?esc(tr('تحتاج الفواتير أو ورقة الأعمار','needs invoices or the aging sheet')):esc(tr(`متأخر ${big(A.over)} · أكثر من 90 يومًا ${big(A.o90)}`,`overdue ${big(A.over)} · over 90 days ${big(A.o90)}`)),tone(A.overR,25,50),A.overR)}
   ${kpi(tr('متوسط فترة التحصيل (DSO)','Days sales outstanding (DSO)'),A.dso==null?'—':`${fn(Math.round(A.dso))}<small>${esc(tr('يوم','days'))}</small>`,tr(`(الذمم ÷ المبيعات الآجلة) × ${days} يوم`,`(receivables ÷ credit sales) × ${days} days`),A.dso==null?esc(tr('لا مبيعات آجلة في الفترة','no credit sales in the period')):esc(tr(`مبيعات آجلة ${big(A.sales)}`,`credit sales ${big(A.sales)}`)),tone(A.dso,60,120))}
   ${kpi(tr('نسبة التحصيل','Collection ratio'),A.collR==null?'—':P1(A.collR),tr('المحصَّل ÷ المستحق في الفترة','Collected ÷ due in the period'),A.collR==null?esc(P.src==='ledger'?tr(`محصّل ${big(A.coll)}؛ المستحق يحتاج الفواتير بتواريخ استحقاقها`,`Collected ${big(A.coll)}; what was due needs invoices with due dates`):tr('تحتاج ورقة التحصيلات','needs the collections sheet')):esc(tr(`محصّل ${big(A.coll)} من ${big(A.due)}`,`collected ${big(A.coll)} of ${big(A.due)}`)),tone(A.collR,80,50,false),A.collR)}
   ${kpi(tr('استخدام سقف الائتمان','Credit-limit use'),A.util==null?'—':P1(A.util),tr('رصيد الزبون ÷ سقفه','Customer balance ÷ limit'),A.util==null?esc(tr('لا سقوف في ورقة الزبائن','no limits in Customers')):esc(tr(`${big(A.limBal)} من سقوف ${big(A.lim)}`,`${big(A.limBal)} of ${big(A.lim)} limits`)),tone(A.util,80,100),A.util)}
   ${kpi(tr('الزبائن المتجاوزون للسقف','Customers over limit'),A.hasLim?fn(A.overLim.length):'—',tr('الرصيد أكبر من السقف','Balance above the limit'),A.hasLim?esc(tr(`التجاوز ${big(A.overLimV)} ${U()}`,`over by ${big(A.overLimV)} ${U()}`)):'',A.hasLim?(A.overLim.length?'crit':'ok'):'')}
   ${kpi(tr('تغطية الضمانات','Guarantee cover'),A.cov==null||!A.hasG?'—':P1(A.cov),tr('قيمة الضمان ÷ رصيد الزبون','Guarantee value ÷ customer balance'),A.hasG?esc(tr(`ضمانات سارية ${big(A.gv)} · دون غطاء ${big(A.unsecured)}`,`valid ${big(A.gv)} · unsecured ${big(A.unsecured)}`)):esc(tr('لا ضمانات لهؤلاء الزبائن','no guarantees for these customers')),A.hasG?tone(A.cov,50,25,false):'',A.hasG?A.cov:null)}
   ${kpi(tr('تركّز أكبر 10 مدينين','Top-10 concentration'),P1(A.total?A.top10*100/A.total:0),tr('أرصدة أكبر 10 ÷ إجمالي الذمم','Top-10 balances ÷ total'),esc(tr(`${big(A.top10)} ${U()}`,`${big(A.top10)} ${U()}`)),tone(A.total?A.top10*100/A.total:0,50,75),A.total?A.top10*100/A.total:0)}
  </div>`}

function agingPanel(A){if(!A.aged)return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('أعمار الديون','Debt aging'))}</h3><div class="empty" style="padding:18px">${esc(tr('لا تتوفر أعمار: الملف بلا فواتير بتواريخ استحقاق ولا ورقة «أعمار الديون».','No aging: the file has no invoices with due dates and no «Aging» sheet.'))}</div></section>`;
  const tot=sum(BK,k=>A.ag[k])||1;const cnt=k=>A.list.filter(c=>c.ag[k]>0).length;
  return `<section class="panel an-wide db-age"><h3 class="section-title">${esc(tr('أعمار الديون','Debt aging'))} <small>${esc(tr(`في تاريخ آخر الفترة · من تاريخ الاستحقاق · ${U()}`,`at the period end · from the due date · ${U()}`))}</small></h3>
   ${stack(BUCKETS.map(b=>({l:tr(b.ar,b.en),v:A.ag[b.k],k:'k-'+b.k})))}
   <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl db-ageT"><thead><tr><th class="nosort">${esc(tr('الفئة','Bucket'))}</th><th class="nosort">${esc(tr('المعنى','Meaning'))}</th><th class="nosort n">${esc(tr('المبلغ','Amount'))}</th><th class="nosort n">${esc(tr('النسبة','Share'))}</th><th class="nosort n">${esc(tr('الزبائن','Customers'))}</th></tr></thead>
   <tbody>${BUCKETS.map(b=>`<tr><td class="t-strong"><span class="sw k-${b.k}"></span>${esc(tr(b.ar,b.en))}</td><td>${esc(tr(b.m[0],b.m[1]))}</td><td class="n">${n0(A.ag[b.k])}</td><td class="n">${pct(A.ag[b.k],tot)}%</td><td class="n">${fn(cnt(b.k))}</td></tr>`).join('')}</tbody>
   <tfoot><tr><td class="t-strong">${esc(tr('الإجمالي','Total'))}</td><td>${esc(tr(`المتأخر ${pct(A.over,tot)}%`,`overdue ${pct(A.over,tot)}%`))}</td><td class="n"><b>${n0(tot)}</b></td><td class="n">100%</td><td class="n">${fn(A.nDebt)}</td></tr></tfoot></table></div>
   ${A.unaged>0?`<p class="db-note">${esc(tr(`${n0(A.unaged)} ${U()} من أرصدة زبائن بلا فواتير ولا أعمار لم تدخل في التوزيع.`,`${n0(A.unaged)} ${U()} of balances without invoices or aging is not distributed.`))}</p>`:''}</section>`}
function sectorPanel(SC,P){const g=SC.g,p=SC.p;if(!g.n||!p.n)return '';
  const col=(k,A)=>`<div class="${k==='gov'?'g':''}"><h4>${esc(secName(k))}</h4><dl>
    <dt>${esc(tr('إجمالي الذمم','Receivables'))}</dt><dd>${big(A.total)}</dd>
    <dt>${esc(tr('نسبة المتأخر','Overdue ratio'))}</dt><dd>${P1(A.overR)}</dd>
    <dt>${esc(tr('أكثر من 180 يومًا','Over 180 days'))}</dt><dd>${big(A.ag.b5)}</dd>
    <dt>DSO</dt><dd>${A.dso==null?'—':fn(Math.round(A.dso))+' '+esc(tr('يوم','d'))}</dd>
    <dt>${esc(tr('نسبة التحصيل','Collection ratio'))}</dt><dd>${P1(A.collR)}</dd>
    <dt>${esc(tr('عدد الزبائن المدينين','Customers owing'))}</dt><dd>${fn(A.nDebt)}</dd></dl></div>`;
  return `<section class="panel"><h3 class="section-title">${esc(tr('القطاع الخاص والجهات الحكومية','Private sector vs government'))} <small>${esc(tr('كل قطاع بمؤشراته','each sector on its own'))}</small></h3>
   ${stack([{l:secName('priv'),v:p.total,k:'k-pv'},{l:secName('gov'),v:g.total,k:'k-gv'}])}<div class="db-sector" style="margin-top:10px">${col('priv',p)}${col('gov',g)}</div></section>`}
function topPanel(A){if(!A.top.length)return '';const mx=Math.max(...A.top.map(c=>c.bal),1);
  return `<section class="panel"><h3 class="section-title">${esc(tr('أكبر 10 مدينين','Top 10 debtors'))} <small>${esc(tr('حسب الرصيد · الجزء الداكن متأخر · اضغط للتفاصيل','by balance · dark part overdue · click for details'))}</small></h3>
   <div class="db-top">${A.top.map((c,i)=>{const cu=Math.max(0,c.bal-c.over);return `<div class="r" data-db-acct="${esc(c.a)}" tabindex="0" role="button" data-tip="${esc(`${c.n} (${c.a}): ${n0(c.bal)} ${U()} · ${tr('متأخر','overdue')} ${n0(c.over)}`)}"><span class="rk">${i+1}</span><span class="nm">${esc(c.n)} <small>${esc(catName(c.cat))}</small></span>
     <span class="tk"><span class="k-b0" style="width:${cu*100/mx}%"></span><span class="k-b4" style="width:${c.over*100/mx}%"></span></span><span class="vv">${big(c.bal)}</span></div>`}).join('')}</div>
   <div class="sm-legend" style="margin-top:8px"><span><i class="k-b0"></i>${esc(tr('غير مستحق','Not due'))}</span><span><i class="k-b4"></i>${esc(tr('متأخر','Overdue'))}</span></div></section>`}
function limitPanel(A){if(!A.hasLim)return '';const L0=A.list.filter(c=>c.lim>0);const band=(lo,hi)=>L0.filter(c=>c.util>=lo&&c.util<hi);
  const B=[[tr('أقل من 50%','Under 50%'),band(-1e9,50),'k-b0'],[tr('50% – 80%','50–80%'),band(50,80),'k-b1'],[tr('80% – 100%','80–100%'),band(80,100.0001),'k-b2'],[tr('فوق السقف','Over the limit'),L0.filter(c=>c.overLim),'k-b4']];
  return `<section class="panel"><h3 class="section-title">${esc(tr('سقوف الائتمان','Credit limits'))} <small>${esc(tr('توزيع الزبائن حسب استخدام السقف','customers by limit use'))}</small></h3>
   <div class="sm-stack">${B.filter(b=>b[1].length).map(b=>`<span class="${b[2]}" style="width:${b[1].length*100/L0.length}%" data-tip="${esc(b[0]+': '+b[1].length)}"></span>`).join('')}</div>
   <div class="sm-legend">${B.map(b=>`<span><i class="${b[2]}"></i>${esc(b[0])} <b class="num">${fn(b[1].length)}</b></span>`).join('')}</div>
   ${A.overLim.length?`<div class="tablewrap" style="max-height:none;min-height:0;margin-top:10px"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الزبون','Customer'))}</th><th class="nosort n">${esc(tr('السقف','Limit'))}</th><th class="nosort n">${esc(tr('الرصيد','Balance'))}</th><th class="nosort n">${esc(tr('التجاوز','Over by'))}</th><th class="nosort">${esc(tr('الاستخدام','Use'))}</th></tr></thead>
    <tbody>${A.overLim.slice(0,8).map(c=>`<tr data-db-acct="${esc(c.a)}" tabindex="0"><td class="t-strong">${esc(c.n)}</td><td class="n">${n0(c.lim)}</td><td class="n">${n0(c.bal)}</td><td class="n crit">${n0(c.bal-c.lim)}</td><td>${utilCell(c)}</td></tr>`).join('')}</tbody></table></div>`:`<p class="db-note">${esc(tr('لا يوجد زبون فوق سقفه.','No customer is over the limit.'))}</p>`}</section>`}
function guarPanel(A){if(!A.hasG)return '';const gs=A.list.flatMap(c=>c.gs.filter(g=>!g.exp||g.exp>=cur().to));const byT=Object.keys(GTYPE).map(k=>({k,v:sum(gs.filter(g=>g.gt===k),'v')})).filter(x=>x.v);
  return `<section class="panel"><h3 class="section-title">${esc(tr('الضمانات','Guarantees'))} <small>${esc(tr('السارية في تاريخ آخر الفترة','valid at the period end'))}</small></h3>
   <div class="db-mini"><div class="${tone(A.cov,50,25,false)}"><b>${P1(A.cov)}</b><span>${esc(tr('التغطية','Cover'))}</span></div><div><b>${big(A.gCovered)}</b><span>${esc(tr('ذمم مغطاة','Covered'))}</span></div><div class="${A.unsecured>0?'warn':''}"><b>${big(A.unsecured)}</b><span>${esc(tr('دون غطاء','Unsecured'))}</span></div><div class="${A.gExp.length?'crit':''}"><b>${fn(A.gExp.length)}</b><span>${esc(tr('ضمان منتهٍ','Expired'))}</span></div></div>
   ${byT.length?stack(byT.map((x,i)=>({l:tr(GTYPE[x.k][0],GTYPE[x.k][1]),v:x.v,k:'sm-k'+(i+1)}))):''}
   ${A.gSoon.length||A.gExp.length?`<div class="tablewrap" style="max-height:none;min-height:0;margin-top:10px"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الزبون','Customer'))}</th><th class="nosort">${esc(tr('الضمان','Guarantee'))}</th><th class="nosort n">${esc(tr('القيمة','Value'))}</th><th class="nosort">${esc(tr('ينتهي','Expires'))}</th></tr></thead>
    <tbody>${[...A.gExp,...A.gSoon].slice(0,8).map(g=>`<tr data-db-acct="${esc(g.c.a)}" tabindex="0"><td class="t-strong">${esc(g.c.n)}</td><td>${esc(tr(GTYPE[g.gt][0],GTYPE[g.gt][1]))} ${g.ref?`<small class="muted">${esc(g.ref)}</small>`:''}</td><td class="n">${n0(g.v)}</td><td class="${g.exp<cur().to?'crit':'warn'}">${fdi(g.exp)} ${esc(g.exp<cur().to?tr('(منتهٍ)','(expired)'):tr(`(بعد ${ddiff(g.exp,cur().to)} يومًا)`,`(in ${ddiff(g.exp,cur().to)} days)`))}</td></tr>`).join('')}</tbody></table></div>`:''}</section>`}
function collPanel(A,P){if(!A.hasColl)return '';const ms=A.byM.filter(x=>x.v);
  return `<section class="panel"><h3 class="section-title">${esc(tr('التحصيلات','Collections'))} <small>${esc(tr(`في الفترة · ${U()}`,`in the period · ${U()}`))}</small></h3>
   <div class="db-mini"><div><b>${big(A.paid)}</b><span>${esc(tr('الدفعات','Payments'))}</span></div><div class="${A.bounced?'crit':''}"><b>${big(A.bounced)}</b><span>${esc(tr('صكوك مرتجعة','Returned cheques'))}</span></div><div class="ok"><b>${big(A.coll)}</b><span>${esc(tr('صافي المحصَّل','Net collected'))}</span></div><div class="${tone(A.collR,80,50,false)}"><b>${P1(A.collR)}</b><span>${esc(tr('نسبة التحصيل','Collection ratio'))}</span></div></div>
   ${ms.length?stack(ms.map((x,i)=>({l:tr(METHOD[x.k][0],METHOD[x.k][1]),v:x.v,k:x.k==='off'?'k-gv':'sm-k'+(i+1)}))):''}
   ${A.bncOpen.length?`<p class="db-note">${esc(tr(`صكوك مرتجعة لم تُحصَّل بعد: ${A.bncOpen.length} بقيمة ${n0(sum(A.bncOpen,'v'))} ${U()}.`,`Returned cheques not yet recovered: ${A.bncOpen.length} worth ${n0(sum(A.bncOpen,'v'))} ${U()}.`))}</p>`:''}
   <p class="db-note">${esc(tr('المستحق في الفترة = المحصَّل + ما بقي متأخرًا في آخر الفترة. الصكوك المرتجعة غير المحصّلة تُطرح من المحصَّل.','Due in the period = collected + what is still overdue at the period end. Returned cheques not recovered are deducted from collections.'))}</p></section>`}
function actionsPanel(A){const any=A.acts.length||A.legal.length||A.sch.length||A.prov||A.wo||A.cn||A.dn;if(!any)return '';
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('إجراءات التحصيل والتسويات','Collection actions and settlements'))}</h3>
   <div class="db-mini">
    <div><b>${fn(A.actN.claim)}</b><span>${esc(tr('مطالبات في الفترة','Claims in the period'))}</span></div>
    <div class="${A.actN.warn?'warn':''}"><b>${fn(A.actN.warn)}</b><span>${esc(tr('إنذارات في الفترة','Warnings in the period'))}</span></div>
    <div class="${A.legal.length?'crit':''}"><b>${fn(A.legal.length)}</b><span>${esc(tr(`محال للشؤون القانونية · ${big(A.legalV)}`,`with Legal Affairs · ${big(A.legalV)}`))}</span></div>
    <div><b>${big(A.schRem)}</b><span>${esc(tr(`متبقٍ من ${fn(A.sch.length)} جدولة`,`left on ${fn(A.sch.length)} reschedulings`))}</span></div>
    <div class="${tone(A.schComp,95,50,false)}"><b>${P1(A.schComp)}</b><span>${esc(tr('الالتزام بالأقساط','Instalment compliance'))}</span></div>
    <div><b>${big(A.prov)}</b><span>${esc(tr(`مخصص الديون المشكوك فيها${A.provCov!=null?` · يغطي ${Math.round(A.provCov)}% مما فوق 90 يومًا`:''}`,`Doubtful-debt provision${A.provCov!=null?` · covers ${Math.round(A.provCov)}% of 90+ days`:''}`))}</span></div>
    <div><b>${big(A.wo)}</b><span>${esc(tr(`ديون معدومة مشطوبة · منها في الفترة ${big(A.woP)}`,`Written off · in the period ${big(A.woP)}`))}</span></div>
    <div><b>${big(A.cn)} / ${big(A.dn)}</b><span>${esc(tr('إشعارات دائنة / مدينة في الفترة','Credit / debit notes in the period'))}</span></div>
    <div><b>${big(A.net)}</b><span>${esc(tr('صافي الذمم بعد المخصص','Net receivables after provision'))}</span></div>
   </div>
   ${A.sch.length?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الزبون','Customer'))}</th><th class="nosort">${esc(tr('الاتفاقية','Agreement'))}</th><th class="nosort n">${esc(tr('المبلغ','Amount'))}</th><th class="nosort n">${esc(tr('الأقساط المستحقة','Due instalments'))}</th><th class="nosort n">${esc(tr('المسدد','Paid'))}</th><th class="nosort n">${esc(tr('الالتزام','Compliance'))}</th><th class="nosort">${esc(tr('الحالة','State'))}</th><th class="nosort">${esc(tr('القسط القادم','Next'))}</th></tr></thead>
    <tbody>${A.sch.slice().sort((a,b)=>a.comp-b.comp).slice(0,10).map(s=>`<tr data-db-acct="${esc(s.c.a)}" tabindex="0"><td class="t-strong">${esc(s.c.n)}</td><td>${esc(s.ref||'—')}</td><td class="n">${n0(s.v)}</td><td class="n">${fn(s.dueN)} / ${fn(s.n)}</td><td class="n">${n0(s.paid)}</td><td class="n">${Math.round(s.comp)}%</td><td><span class="pill p-${SCH[s.st][2]}">${esc(tr(SCH[s.st][0],SCH[s.st][1]))}</span></td><td>${s.next?fdi(s.next):'—'}</td></tr>`).join('')}</tbody></table></div>`:''}</section>`}
function govPanel(A,u){if(u.sec==='priv')return '';const G=A.list.filter(c=>c.gov);if(!G.length)return '';const pub=A.pub;const gt=sum(pub,'bal');
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('ديون الجهات الحكومية','Government debts'))} <small>${esc(tr('الكهرباء والوزارات والبلديات · منفصلة عن القطاع الخاص','electricity, ministries, municipalities · kept apart from the private sector'))}</small></h3>
   <div class="db-mini"><div><b>${big(gt)}</b><span>${esc(tr('المستحق على الجهات العامة','Owed by public bodies'))}</span></div><div><b>${big(A.sub)}</b><span>${esc(tr('الدعم المستحق للشركة','Subsidy owed to the company'))}</span></div><div class="ok"><b>${big(A.offDone)}</b><span>${esc(tr('مقاصّة منفذة','Offsets executed'))}</span></div><div class="warn"><b>${big(A.offPend)}</b><span>${esc(tr('مقاصّة قيد الإجراء','Offsets pending'))}</span></div><div><b>${big(Math.max(0,gt-A.offPend))}</b><span>${esc(tr('الصافي بعد المقاصّة المنتظرة','Net after pending offsets'))}</span></div></div>
   ${stack(pub.map(x=>({l:tr(PUB[x.k][0],PUB[x.k][1]),v:x.bal,k:{power:'sm-k2',min:'sm-k1',mun:'sm-k3',oth:'sm-k4'}[x.k]})))}
   ${A.offs.length?`<div class="tablewrap" style="max-height:none;min-height:0;margin-top:10px"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الجهة','Body'))}</th><th class="nosort">${esc(tr('التاريخ','Date'))}</th><th class="nosort n">${esc(tr('الدعم المستحق','Subsidy owed'))}</th><th class="nosort n">${esc(tr('المقاصّة','Offset'))}</th><th class="nosort">${esc(tr('الحالة','Status'))}</th><th class="nosort">${esc(tr('المرجع','Reference'))}</th></tr></thead>
    <tbody>${A.offs.slice().sort((a,b)=>(b.d||'').localeCompare(a.d||'')).slice(0,10).map(o=>`<tr data-db-acct="${esc(o.c.a)}" tabindex="0"><td class="t-strong">${esc(o.c.n)}</td><td>${fdi(o.d)}</td><td class="n">${o.sub?n0(o.sub):'—'}</td><td class="n">${n0(o.v)}</td><td><span class="pill p-${o.st==='done'?'ok':'warn'}">${esc(o.st==='done'?tr('منفذة','Executed'):tr('قيد الإجراء','Pending'))}</span></td><td>${esc(o.ref||'—')}</td></tr>`).join('')}</tbody></table></div>`:''}
   <p class="db-note">${esc(tr('المقاصّة المنفذة التي سُجّلت دفعةً بطريقة «مقاصّة مع الدولة» في ورقة التحصيلات تدخل في نسبة التحصيل؛ أرقام هذه الورقة للمتابعة ولا تُضاف مرة ثانية.','Executed offsets recorded as «offset with the state» payments are part of the collection ratio; this sheet is for follow-up and is not added twice.'))}</p></section>`}
function mixPanel(A){if(A.byCat.length<2&&A.byRg.length<2&&A.byWh.length<2)return '';
  const B=(rows,name)=>bars(rows.slice(0,10).map(x=>({l:name(x.k),v:x.bal,tip:`${name(x.k)}: ${n0(x.bal)} ${U()} · ${tr('متأخر','overdue')} ${pct(x.over,x.bal)}% · ${x.n} ${tr('زبون','customers')}`})),{fmt:big});
  return `<section class="panel"><h3 class="section-title">${esc(tr('توزيع الذمم','Receivables mix'))} <small>${esc(tr('حسب الفئة والمنطقة والمستودع','by category, region and warehouse'))}</small></h3>
   ${A.byCat.length>1?`<h4 class="db-h4">${esc(tr('الفئة','Category'))}</h4>${B(A.byCat,catName)}`:''}${A.byRg.length>1?`<h4 class="db-h4">${esc(tr('المنطقة','Region'))}</h4>${B(A.byRg,k=>k||tr('غير محددة','Unspecified'))}`:''}${A.byWh.length>1?`<h4 class="db-h4">${esc(tr('المستودع','Warehouse'))}</h4>${B(A.byWh,k=>k)}`:''}</section>`}
function trendPanel(L1,u){if(L1.length<2)return '';const pts=L1.map(P=>{const A=analyze(P,u);return {P,A}});
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('تطوّر الذمم','Receivables trend'))} <small>${esc(secName(u.sec))} · ${esc(U())}</small></h3>
   ${colChart(pts.map(({P,A})=>({l:mLabel(P.to),y:A.total,tip:`${mLabel(P.to)}: ${n0(A.total)} ${U()}`})),tr('تطور الذمم','Receivables trend'))}
   <div class="tablewrap" style="max-height:none;min-height:0;margin-top:8px"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الفترة','Period'))}</th><th class="nosort n">${esc(tr('الذمم','Receivables'))}</th><th class="nosort n">${esc(tr('المتأخر','Overdue'))}</th><th class="nosort n">DSO</th><th class="nosort n">${esc(tr('نسبة التحصيل','Collection'))}</th><th class="nosort n">${esc(tr('فوق السقف','Over limit'))}</th></tr></thead>
   <tbody>${pts.slice().reverse().map(({P,A})=>`<tr><td>${esc(mLabel(P.to))}</td><td class="n">${n0(A.total)}</td><td class="n">${P1(A.overR)}</td><td class="n">${A.dso==null?'—':fn(Math.round(A.dso))}</td><td class="n">${P1(A.collR)}</td><td class="n">${A.hasLim?fn(A.overLim.length):'—'}</td></tr>`).join('')}</tbody></table></div></section>`}
function readingPanel(A,SC,P,u){const ins=[];const add=(t,ic,s)=>ins.push([t,ic,s]);
  add('',tr('Σ','Σ'),tr(`إجمالي ذمم ${secName(u.sec)} <b>${n0(A.total)}</b> ${U()} على <b>${fn(A.nDebt)}</b> زبونًا مدينًا في ${fdi(P.to)}.`,`${secName(u.sec)} receivables: <b>${n0(A.total)}</b> ${U()} across <b>${fn(A.nDebt)}</b> customers on ${fdi(P.to)}.`));
  if(A.overR!=null)add(A.overR>=50?'c':A.overR>=25?'w':'o','%',tr(`المتأخر <b>${P1(A.overR)}</b> من الذمم (${n0(A.over)} ${U()})، منه <b>${n0(A.o90)}</b> فوق 90 يومًا و<b>${n0(A.ag.b5)}</b> مشكوك في تحصيله (أكثر من 180 يومًا).`,`Overdue is <b>${P1(A.overR)}</b> of receivables (${n0(A.over)} ${U()}); <b>${n0(A.o90)}</b> is over 90 days and <b>${n0(A.ag.b5)}</b> is doubtful (over 180).`));
  if(A.dso!=null){const t=Math.round(A.dso);const terms=A.list.filter(c=>c.terms>0);const avgT=terms.length?Math.round(sum(terms,c=>c.terms*Math.max(0,c.bal))/Math.max(1,sum(terms,c=>Math.max(0,c.bal)))):0;
    add(avgT&&t>avgT*1.25?'w':'o','⏱',tr(`متوسط فترة التحصيل <b>${fn(t)}</b> يومًا${avgT?` مقابل مدة سداد متفق عليها متوسطها <b>${avgT}</b> يومًا`:''}.`,`DSO is <b>${fn(t)}</b> days${avgT?` against agreed terms averaging <b>${avgT}</b> days`:''}.`))}
  if(A.collR!=null)add(A.collR<50?'c':A.collR<80?'w':'o','↺',tr(`حُصّل <b>${P1(A.collR)}</b> من المستحق في الفترة (${n0(A.coll)} من ${n0(A.due)} ${U()}).`,`<b>${P1(A.collR)}</b> of what was due in the period was collected (${n0(A.coll)} of ${n0(A.due)} ${U()}).`));
  if(A.overLim.length)add('c','⚠',tr(`<b>${A.overLim.length}</b> زبونًا تجاوز سقف الائتمان بمجموع <b>${n0(A.overLimV)}</b> ${U()}؛ أكبرهم ${esc(A.overLim[0].n)} (${Math.round(A.overLim[0].util)}% من سقفه).`,`<b>${A.overLim.length}</b> customers are over their limit by <b>${n0(A.overLimV)}</b> ${U()}; the largest is ${esc(A.overLim[0].n)} (${Math.round(A.overLim[0].util)}% of the limit).`));
  if(A.hasG){if(A.gExp.length)add('c','⛨',tr(`<b>${A.gExp.length}</b> ضمانًا منتهي الصلاحية ما زال زبونه مدينًا؛ و<b>${A.gSoon.length}</b> ينتهي خلال 60 يومًا.`,`<b>${A.gExp.length}</b> guarantees have expired while the customer still owes; <b>${A.gSoon.length}</b> expire within 60 days.`));
    if(A.unsecured>0)add('w','⛨',tr(`<b>${n0(A.unsecured)}</b> ${U()} من الذمم بلا ضمان يغطيها (التغطية ${P1(A.cov)}).`,`<b>${n0(A.unsecured)}</b> ${U()} of receivables has no guarantee behind it (cover ${P1(A.cov)}).`))}
  if(A.total&&A.top.length){const s=A.top10*100/A.total;add(s>=75?'w':'','★',tr(`أكبر 10 مدينين يحملون <b>${P1(s)}</b> من الذمم؛ أولهم ${esc(A.top[0].n)} بـ ${n0(A.top[0].bal)} ${U()}.`,`The top 10 debtors hold <b>${P1(s)}</b> of receivables; first is ${esc(A.top[0].n)} with ${n0(A.top[0].bal)} ${U()}.`))}
  if(A.bounced)add('c','✗',tr(`صكوك مرتجعة في الفترة بقيمة <b>${n0(A.bounced)}</b> ${U()} (${A.bnc.length} صكًا).`,`Returned cheques in the period: <b>${n0(A.bounced)}</b> ${U()} (${A.bnc.length}).`));
  if(A.sch.length){const br=A.sch.filter(s=>s.st==='broken');if(br.length)add('c','⌛',tr(`<b>${br.length}</b> اتفاقية جدولة متعثرة (التزام أقل من 50%).`,`<b>${br.length}</b> rescheduling agreements are broken (under 50% compliance).`))}
  if(A.o90&&A.prov<A.o90*0.5&&P.prov.length)add('w','₪',tr(`المخصص (${n0(A.prov)}) يغطي ${Math.round(A.provCov)}% فقط من الديون فوق 90 يومًا.`,`The provision (${n0(A.prov)}) covers only ${Math.round(A.provCov)}% of debts over 90 days.`));
  if(u.sec!=='gov'&&SC.g.total)add('','⚖',tr(`ديون الجهات الحكومية (${n0(SC.g.total)} ${U()}) مفصولة عن هذه المؤشرات${A.offPend?`؛ مقاصّة قيد الإجراء ${n0(A.offPend)}`:''}.`,`Government debts (${n0(SC.g.total)} ${U()}) are kept out of these figures.`));
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('قراءة تحليلية','Reading'))} <small>${esc(tr('تُكتب تلقائيًا من البيانات المصفّاة','written from the filtered data'))}</small></h3><ul class="sm-ins">${ins.map(([t,ic,s])=>`<li><span class="ic ${t}">${esc(ic)}</span><span>${s}</span></li>`).join('')}</ul></section>`}
function prioPanel(A){if(!A.prio.length)return '';
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('أولويات المتابعة والتحصيل','Collection priorities'))} <small>${esc(tr('الأقدم والأكبر أولًا · الإجراء المقترح حسب عمر أقدم دين وحالة الحساب','oldest and largest first · step from the oldest debt and the account state'))}</small></h3>
   <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">#</th><th class="nosort">${esc(tr('الزبون','Customer'))}</th><th class="nosort">${esc(tr('الحالة','Status'))}</th><th class="nosort n">${esc(tr('الرصيد','Balance'))}</th><th class="nosort n">${esc(tr('المتأخر','Overdue'))}</th><th class="nosort n">${esc(tr('أكثر من 90 يومًا','90+ days'))}</th><th class="nosort n">${esc(tr('أقدم تأخر','Oldest'))}</th><th class="nosort">${esc(tr('الإجراء المقترح','Suggested step'))}</th></tr></thead>
   <tbody>${A.prio.slice(0,15).map((c,i)=>`<tr data-db-acct="${esc(c.a)}" tabindex="0"><td class="n">${i+1}</td><td class="t-strong">${esc(c.n)} <small class="muted num">${esc(c.a)}</small>${c.gov?`<span class="db-tag g">${esc(tr('حكومي','gov'))}</span>`:''}</td><td>${stPill(c.st)}</td><td class="n">${n0(c.bal)}</td><td class="n">${n0(c.over)}</td><td class="n ${c.o90?'crit':''}">${n0(c.o90)}</td><td class="n">${c.oldest>0?esc(tr(`${fn(c.oldest)} يومًا`,`${fn(c.oldest)} d`)):'—'}</td><td>${esc(c.step)}</td></tr>`).join('')}</tbody></table></div></section>`}
const SORT={name:c=>c.n,bal:c=>c.bal,lim:c=>c.lim,util:c=>c.util==null?-1:c.util,over:c=>c.over,o90:c=>c.o90,cov:c=>c.cov==null?-1:c.cov,dso:c=>c.dso==null?-1:c.dso,last:c=>c.lastPay||''};
function tablePanel(A,u){const K=SORT[u.sort]||SORT.bal;const rows=A.list.slice().sort((a,b)=>{const x=K(a),y=K(b);const r=typeof x==='string'?x.localeCompare(y,'ar'):x-y;return u.dir==='asc'?r:-r});
  const tb=rows.slice(0,u.lim);const th=(k,l,cls='')=>`<th class="${cls}" data-db-sort="${k}" ${u.sort===k?`aria-sort="${u.dir==='asc'?'ascending':'descending'}"`:''}>${esc(l)}</th>`;
  return `<section class="panel an-wide"><h3 class="section-title">${esc(tr('الزبائن','Customers'))} <small>${esc(tr(`${fn(rows.length)} زبونًا · اضغط على الزبون لكشف حسابه`,`${fn(rows.length)} customers · click a customer for the statement`))}</small></h3>
   <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">#</th>${th('name',tr('الزبون','Customer'))}<th class="nosort">${esc(tr('الفئة','Category'))}</th><th class="nosort">${esc(tr('الحالة','Status'))}</th>${th('lim',tr('السقف','Limit'),'n')}${th('bal',tr('الرصيد','Balance'),'n')}${th('util',tr('استخدام السقف','Limit use'))}${th('over',tr('المتأخر','Overdue'),'n')}${th('o90',tr('+90 يومًا','90+ days'),'n')}${th('cov',tr('تغطية الضمان','Cover'),'n')}${th('dso','DSO','n')}${th('last',tr('آخر دفعة','Last payment'))}</tr></thead>
   <tbody>${tb.map((c,i)=>`<tr data-db-acct="${esc(c.a)}" tabindex="0"><td class="n">${i+1}</td><td class="t-strong">${esc(c.n)} <small class="muted num">${esc(c.a)}</small></td><td>${esc(catName(c.cat))}</td><td>${stPill(c.st)}</td><td class="n">${c.lim?n0(c.lim):'—'}</td><td class="n"><b>${n0(c.bal)}</b></td><td>${utilCell(c)}</td><td class="n ${c.over>0?'warn':''}">${c.aged?n0(c.over):'—'}</td><td class="n ${c.o90>0?'crit':''}">${c.aged?n0(c.o90):'—'}</td><td class="n ${c.cov!=null&&c.cov<25&&c.bal>0?'warn':''}">${c.cov==null?'—':Math.round(c.cov)+'%'}</td><td class="n">${c.dso==null?'—':fn(Math.round(c.dso))}</td><td>${c.lastPay?fdi(c.lastPay):'—'}</td></tr>`).join('')}</tbody>
   <tfoot><tr><td></td><td class="t-strong">${esc(tr('الإجمالي','Total'))}</td><td></td><td></td><td class="n">${n0(A.lim)}</td><td class="n"><b>${n0(sum(rows,'bal'))}</b></td><td>${A.util==null?'':P1(A.util)}</td><td class="n">${n0(A.over)}</td><td class="n">${n0(A.o90)}</td><td class="n">${A.hasG?P1(A.cov):''}</td><td class="n">${A.dso==null?'':fn(Math.round(A.dso))}</td><td></td></tr></tfoot></table></div>
   ${rows.length>u.lim?`<div style="text-align:center;margin-top:8px"><button class="btn sm" data-dbt="more">${esc(tr(`عرض المزيد (${rows.length-u.lim} متبقٍ)`,`Show more (${rows.length-u.lim} left)`))}</button></div>`:''}</section>`}

/* ---------- customer statement ---------- */
function acctDetail(a){const P=cur();if(!P)return;const c=prep(P).map.get(a);if(!c)return;
  const T=(head,rows)=>`<div class="tablewrap" style="max-height:320px;min-height:0"><table class="sm-tbl"><thead><tr>${head.map(h=>`<th class="nosort${h[1]?' n':''}">${esc(h[0])}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
  const inv=c.inv.slice().sort((x,y)=>(y.rem>0)-(x.rem>0)||(x.due||'').localeCompare(y.due||''));
  const body=`<div class="db-dlg db-page"><div class="db-mini">
    <div><b>${n0(c.bal)}</b><span>${esc(tr(`الرصيد · ${U()}`,`Balance · ${U()}`))}</span></div>
    <div class="${c.over>0?'warn':''}"><b>${c.aged?n0(c.over):'—'}</b><span>${esc(tr('المتأخر','Overdue'))}</span></div>
    <div class="${c.overLim?'crit':''}"><b>${c.util==null?'—':Math.round(c.util)+'%'}</b><span>${esc(tr(`من سقف ${c.lim?n0(c.lim):'—'}`,`of limit ${c.lim?n0(c.lim):'—'}`))}</span></div>
    <div class="${c.cov!=null&&c.cov<25?'warn':''}"><b>${c.cov==null?'—':Math.round(c.cov)+'%'}</b><span>${esc(tr(`تغطية الضمان (${n0(c.gv)})`,`Guarantee cover (${n0(c.gv)})`))}</span></div>
    <div><b>${c.dso==null?'—':fn(Math.round(c.dso))}</b><span>DSO ${esc(tr('(يوم)','(days)'))}</span></div>
    <div><b>${c.lastPay?fdi(c.lastPay):'—'}</b><span>${esc(tr('آخر دفعة','Last payment'))}</span></div></div>
   <div class="db-dl">
    <div><span>${esc(tr('الكود','Code'))}</span><b class="num">${esc(c.a)}</b></div><div><span>${esc(tr('الفئة','Category'))}</span><b>${esc(catName(c.cat))}${c.gov?' · '+esc(tr(PUB[c.pubT][0],PUB[c.pubT][1])):''}</b></div>
    <div><span>${esc(tr('حالة الحساب','Status'))}</span>${stPill(c.st)}</div><div><span>${esc(tr('المنطقة','Region'))}</span><b>${esc(c.rg||'—')}</b></div>
    <div><span>${esc(tr('المستودع','Warehouse'))}</span><b>${esc(c.wh||'—')}</b></div><div><span>${esc(tr('مدة السداد','Terms'))}</span><b>${c.terms?esc(tr(`${c.terms} يومًا`,`${c.terms} days`)):'—'}</b></div>
    <div><span>${esc(tr('العقد','Contract'))}</span><b>${esc(c.ctr||'—')}${c.ctrd?' · '+fdi(c.ctrd):''}</b></div><div><span>${esc(tr('الإجراء المقترح','Suggested step'))}</span><b>${esc(c.step||'—')}</b></div></div>
   ${c.aged&&c.bal>0?`<h4>${esc(tr('أعمار الديون','Aging'))}</h4>${stack(BUCKETS.map(b=>({l:tr(b.ar,b.en),v:c.ag[b.k],k:'k-'+b.k})))}`:''}
   ${inv.length?`<h4>${esc(tr(`الفواتير (${inv.length})`,`Invoices (${inv.length})`))}</h4>`+T([[tr('رقم الفاتورة','Invoice')],[tr('التاريخ','Date')],[tr('المنتج','Product')],[tr('القيمة','Value'),1],[tr('الاستحقاق','Due')],[tr('المسدد','Paid'),1],[tr('المتبقي','Remaining'),1],[tr('التأخر','Late'),1]],
     inv.slice(0,200).map(i=>{const d=i.due&&i.rem>0?ddiff(P.to,i.due):0;return `<tr><td class="num">${esc(i.no)}</td><td>${fdi(i.d)}</td><td>${esc(i.p||'—')}${i.q?` <small class="muted num">${fn(i.q)} ${esc(i.u||'')}</small>`:''}</td><td class="n">${n0(i.v)}</td><td>${fdi(i.due)}</td><td class="n">${n0(i.paid)}</td><td class="n"><b>${n0(i.rem)}</b></td><td class="n ${d>90?'crit':d>0?'warn':''}">${d>0?fn(d):'—'}</td></tr>`})):''}
   ${c.pays.length?`<h4>${esc(tr(`التحصيلات (${c.pays.length})`,`Collections (${c.pays.length})`))}</h4>`+T([[tr('التاريخ','Date')],[tr('المبلغ','Amount'),1],[tr('الطريقة','Method')],[tr('الإيصال','Receipt')],[tr('الفواتير المسددة','Invoices')]],
     c.pays.slice().sort((x,y)=>(y.d||'').localeCompare(x.d||'')).slice(0,100).map(p=>`<tr><td>${fdi(p.d)}</td><td class="n">${n0(p.v)}</td><td>${esc(tr(METHOD[p.m][0],METHOD[p.m][1]))}</td><td>${esc(p.ref||'—')}</td><td>${esc(p.inv||'—')}</td></tr>`)):''}
   ${c.bnc.length?`<h4>${esc(tr('الصكوك المرتجعة','Returned cheques'))}</h4>`+T([[tr('رقم الصك','Cheque')],[tr('التاريخ','Date')],[tr('القيمة','Value'),1],[tr('المصرف','Bank')],[tr('السبب','Reason')],[tr('حُصّل لاحقًا','Recovered')]],
     c.bnc.map(b=>`<tr><td class="num">${esc(b.no)}</td><td>${fdi(b.d)}</td><td class="n">${n0(b.v)}</td><td>${esc(b.bank||'—')}</td><td>${esc(b.why||'—')}</td><td>${b.rec?esc(tr('نعم','Yes')):`<span class="pill p-crit">${esc(tr('لا','No'))}</span>`}</td></tr>`)):''}
   ${c.gs.length?`<h4>${esc(tr('الضمانات','Guarantees'))}</h4>`+T([[tr('النوع','Type')],[tr('الرقم','Number')],[tr('المصرف','Bank')],[tr('القيمة','Value'),1],[tr('ينتهي','Expires')]],
     c.gs.map(g=>`<tr><td>${esc(tr(GTYPE[g.gt][0],GTYPE[g.gt][1]))}</td><td>${esc(g.ref||'—')}</td><td>${esc(g.bank||'—')}</td><td class="n">${n0(g.v)}</td><td class="${g.exp&&g.exp<P.to?'crit':''}">${fdi(g.exp)}${g.exp&&g.exp<P.to?' '+esc(tr('(منتهٍ)','(expired)')):''}</td></tr>`)):''}
   ${c.acts.length?`<h4>${esc(tr('إجراءات التحصيل','Collection actions'))}</h4>`+T([[tr('التاريخ','Date')],[tr('الإجراء','Action')],[tr('المرجع','Reference')],[tr('المبلغ','Amount'),1],[tr('ملاحظات','Notes')]],
     c.acts.slice().sort((x,y)=>(y.d||'').localeCompare(x.d||'')).map(x=>`<tr><td>${fdi(x.d)}</td><td>${esc(tr(ACT[x.t][0],ACT[x.t][1]))}</td><td>${esc(x.ref||'—')}</td><td class="n">${x.v?n0(x.v):'—'}</td><td>${esc(x.note||'')}</td></tr>`)):''}
   ${c.sch.length?`<h4>${esc(tr('الجدولة','Rescheduling'))}</h4>`+T([[tr('الاتفاقية','Agreement')],[tr('المبلغ','Amount'),1],[tr('الأقساط المستحقة','Due'),1],[tr('المسدد','Paid'),1],[tr('الالتزام','Compliance'),1],[tr('الحالة','State')]],
     c.sch.map(s=>`<tr><td>${esc(s.ref||'—')}</td><td class="n">${n0(s.v)}</td><td class="n">${fn(s.dueN)} / ${fn(s.n)}</td><td class="n">${n0(s.paid)}</td><td class="n">${Math.round(s.comp)}%</td><td><span class="pill p-${SCH[s.st][2]}">${esc(tr(SCH[s.st][0],SCH[s.st][1]))}</span></td></tr>`)):''}
   ${c.notes.length||c.prs.length?`<h4>${esc(tr('التسويات','Settlements'))}</h4>`+T([[tr('النوع','Type')],[tr('التاريخ','Date')],[tr('الفاتورة','Invoice')],[tr('المبلغ','Amount'),1],[tr('البيان','Detail')]],
     [...c.notes.map(x=>`<tr><td>${esc(tr(NOTE[x.t][0],NOTE[x.t][1]))} ${x.no?`<small class="muted">${esc(x.no)}</small>`:''}</td><td>${fdi(x.d)}</td><td>${esc(x.inv||'—')}</td><td class="n">${n0(x.v)}</td><td>${esc(x.why||'')}</td></tr>`),
      ...c.prs.map(x=>`<tr><td>${esc(x.t==='wo'?tr('شطب دين معدوم','Write-off'):tr('مخصص ديون مشكوك فيها','Doubtful-debt provision'))}</td><td>${fdi(x.d)}</td><td>${esc(x.inv||'—')}</td><td class="n">${n0(x.v)}</td><td>${esc(x.note||'')}</td></tr>`)]):''}
   ${c.offs.length?`<h4>${esc(tr('مقاصّة الدعم','Subsidy offsets'))}</h4>`+T([[tr('التاريخ','Date')],[tr('الدعم المستحق','Subsidy'),1],[tr('المقاصّة','Offset'),1],[tr('الحالة','Status')],[tr('المرجع','Reference')]],
     c.offs.map(o=>`<tr><td>${fdi(o.d)}</td><td class="n">${o.sub?n0(o.sub):'—'}</td><td class="n">${n0(o.v)}</td><td>${esc(o.st==='done'?tr('منفذة','Executed'):tr('قيد الإجراء','Pending'))}</td><td>${esc(o.ref||'—')}</td></tr>`)):''}
   ${P.src==='ledger'?`<p class="db-note">${esc(tr(`من كشف الأرصدة الشهري: رصيد افتتاحي ${n0(c.ob)} · مسحوبات ${n0(c.w)} · تحصيلات ${n0(c.c)}.`,`From the monthly statement: opening ${n0(c.ob)} · withdrawals ${n0(c.w)} · collections ${n0(c.c)}.`))}</p>`:''}</div>`;
  modal(`${c.n} (${c.a})`,body,`<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`);
  const d=document.getElementById('dlg');if(d){d.classList.add('db-wide');d.addEventListener('close',()=>d.classList.remove('db-wide'),{once:true})}}

/* ---------- import ---------- */
async function importFile(file){toast(tr('جارٍ قراءة الملف…','Reading the file…'));let P;
  try{P=await parseFile(file)}catch(e){if(e&&e.cancelled)return;modal(tr('تعذّر الاستيراد','Import failed'),`<p style="line-height:1.8">${esc(e.message||String(e))}</p>`,`<button type="button" class="btn" data-dbt="template">${esc(tr('تنزيل القالب','Download template'))}</button><button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`);return}
  ST.pending=P;const ov=ST.local?[]:overlaps(P);const X=prep(P);const tot=sum(X.cs,c=>Math.max(0,c.bal));const gov=sum(X.cs.filter(c=>c.gov),c=>Math.max(0,c.bal));
  const nm=k=>k==='ledger'?tr('كشف أرصدة المدينين','Debtors balances statement'):tr(TPLK[k].ar,TPLK[k].en);
  modal(tr('مراجعة الملف قبل الحفظ','Review the file before saving'),`<p class="muted" style="margin:0;font-size:12.5px">${esc(tr('الملف: ','File: ')+file.name)} · ${esc(tr(`الفترة ${fdi(P.from)} – ${fdi(P.to)}`,`Period ${fdi(P.from)} – ${fdi(P.to)}`))}</p>
   <div class="sm-mini"><div><b class="num">${big(tot)}</b><span>${esc(tr('إجمالي الذمم (د.ل)','Total receivables (LYD)'))}</span></div><div><b class="num">${big(gov)}</b><span>${esc(tr('منها الجهات الحكومية','of which government'))}</span></div><div><b class="num">${fn(P.cust.length)}</b><span>${esc(tr('زبون','customers'))}</span></div><div><b class="num">${fn(P.inv.length)}</b><span>${esc(tr('فاتورة','invoices'))}</span></div><div><b class="num">${fn(P.pay.length)}</b><span>${esc(tr('دفعة','payments'))}</span></div></div>
   <h4 style="margin:10px 0 4px">${esc(tr('الأوراق المقروءة','Sheets read'))}</h4>
   <ul class="db-chk">${P.sheets.map(s=>`<li class="${s.missing.length?'w':'o'}">${esc(nm(s.k))}: ${esc(tr(`${fn(s.rows)} صف من ورقة «${s.sheet}»`,`${fn(s.rows)} rows from «${s.sheet}»`))}${s.missing.length?' · '+esc(tr('أعمدة ناقصة: ','missing columns: ')+s.missing.join('، ')):''}</li>`).join('')}
    ${P.src==='tpl'?TPL.filter(s=>!P.sheets.some(x=>x.k===s.k)&&s.k!=='age').map(s=>`<li>${esc(tr(`${s.ar}: غير موجودة (اختيارية)`,`${s.en}: not present (optional)`))}</li>`).join(''):''}</ul>
   ${P.warn.length?`<h4 style="margin:10px 0 4px">${esc(tr('ملاحظات الفحص','Checks'))}</h4><ul class="db-chk">${P.warn.map(([t,m])=>`<li class="${t}">${esc(m)}</li>`).join('')}</ul>`:''}
   ${ov.length?`<div class="banner warn" style="margin-top:8px">${esc(tr('يوجد ملف محفوظ للفترة نفسها:','A file for the same period is already saved:'))} ${ov.map(b=>esc(`${b.file} (${fdi(b.to)})`)).join('، ')}
     <label style="display:flex;gap:6px;align-items:center;margin-top:6px"><input type="checkbox" id="db-replace" checked> ${esc(tr('استبداله بهذا الملف (يمنع تكرار الأرقام)','Replace it with this file (avoids double counting)'))}</label></div>`:''}
   <p id="db-msg" class="muted" style="margin:8px 0 0;font-size:12.5px"></p>`,
   `${ST.canImport?`<button type="button" class="btn primary" data-dbt="save">${esc(tr('حفظ في المنظومة','Save to the system'))}</button>`:''}<button type="button" class="btn" data-dbt="view-local">${esc(tr('عرض دون حفظ','View without saving'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}

/* ---------- the template ---------- */
function templateSheets(){const guide=[];
  TPL.forEach(s=>{guide.push({'الورقة':s.ar,'العمود':'','مطلوب':'','الصيغة':'','القيم المسموحة / الشرح':s.desc,'مثال':''});
    s.cols.filter(c=>!c[8]).forEach(c=>guide.push({'الورقة':'','العمود':c[1],'مطلوب':c[4]?'نعم':'اختياري','الصيغة':{s:'نص',n:'رقم',d:'تاريخ (سنة-شهر-يوم)'}[c[3]],'القيم المسموحة / الشرح':c[6]||'','مثال':c[5]===''?'':c[5]}))});
  const notes=[{'#':1,'إرشاد':'لا تغيّر أسماء الأوراق ولا عناوين الأعمدة. يمكن حذف ورقة لا تتوفر بياناتها، أو ترك عمود اختياري فارغًا.'},
    {'#':2,'إرشاد':'كود الزبون هو الرابط بين كل الأوراق: يجب أن يكون هو نفسه في الزبائن والضمانات والفواتير والتحصيلات وبقية الأوراق.'},
    {'#':3,'إرشاد':'المبالغ بالدينار الليبي، أرقامًا دون رموز عملة. التواريخ بصيغة تاريخ Excel أو 2026-10-31.'},
    {'#':4,'إرشاد':'أعمار الديون تُحسب تلقائيًا من «تاريخ الاستحقاق» و«المتبقي» في كل فاتورة في تاريخ نهاية الفترة. ورقة «أعمار الديون» بديل عند عدم توفر الفواتير فقط.'},
    {'#':5,'إرشاد':'الجهات الحكومية: الفئة «جهة حكومية» أو «كهرباء»، و«نوع الجهة العامة»: الكهرباء أو وزارة أو بلدية. تُعرض مؤشراتها منفصلة عن القطاع الخاص.'},
    {'#':6,'إرشاد':'التحصيل بطريقة «مقاصّة مع الدولة» يُسجّل في التحصيلات كأي دفعة، وتفاصيل المقاصّة مع الدعم في ورقة «مقاصة الدعم» للمتابعة.'},
    {'#':7,'إرشاد':'ملف واحد لكل فترة. استيراد فترة محفوظة سابقًا يستبدلها.'}];
  const empty=s=>[Object.fromEntries(s.cols.filter(c=>!c[8]).map(c=>[c[1],'']))];
  return [['اقرأني',notes],['دليل الأعمدة',guide],...TPL.map(s=>[s.ar,s.k==='per'?[{'بداية الفترة':'','نهاية الفترة':''}]:empty(s)])]}
async function downloadTemplate(){try{const blob=buildXlsx(templateSheets());await S.downloads.save({filename:TPL_NAME,data:blob});toast(tr('تم تنزيل القالب','Template downloaded'))}catch(e){console.error(e);toast(tr('تعذّر تجهيز القالب','Could not build the template'))}}
function templateHelp(){modal(tr('قالب بيانات الديون','Debts data template'),`<div class="db-tpl"><p style="margin:0 0 6px;line-height:1.8">${esc(tr('ملف Excel واحد لكل فترة، بالأوراق التالية. كود الزبون يربط الأوراق ببعضها، وأعمار الديون تُحسب تلقائيًا في تاريخ نهاية الفترة.','One Excel file per period with these sheets. The customer code links the sheets; aging is computed at the period end.'))}</p>
  <ol>${TPL.map(s=>`<li><b>${esc(tr(s.ar,s.en))}</b>: ${esc(s.desc)} <small class="muted">${esc(s.cols.filter(c=>c[4]).map(c=>c[1]).join('، '))}</small></li>`).join('')}</ol>
  <p class="db-note">${esc(tr('الأعمدة الصغيرة الرمادية هي الأعمدة المطلوبة في كل ورقة؛ والشرح الكامل لكل عمود في ورقة «دليل الأعمدة» داخل القالب.','The grey columns are required; every column is explained in the «column guide» sheet inside the template.'))}</p></div>`,
  `<button type="button" class="btn primary" data-dbt="template">${I.down}${esc(tr('تنزيل القالب','Download template'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- export ---------- */
function exportXlsx(){const P=cur();if(!P)return;const u=UI();const A=analyze(P,u);const K=tr('المؤشر','Indicator'),V=tr('القيمة','Value'),F=tr('المعادلة','Formula');
  const kp=[[tr('القطاع','Sector'),secName(u.sec),''],[tr('الفترة','Period'),`${P.from} – ${P.to}`,''],[tr('إجمالي الذمم','Total receivables'),r2(A.total),tr('مجموع أرصدة الزبائن','sum of balances')],
    [tr('المتأخر','Overdue'),r2(A.over),''],[tr('نسبة الديون المتأخرة %','Overdue ratio %'),A.overR==null?'':r2(A.overR),tr('المتأخر ÷ إجمالي الذمم','overdue ÷ total')],
    [tr('المبيعات الآجلة في الفترة','Credit sales in the period'),r2(A.sales),''],[tr('متوسط فترة التحصيل DSO (يوم)','DSO (days)'),A.dso==null?'':Math.round(A.dso),tr(`(الذمم ÷ المبيعات الآجلة) × ${A.days}`,`(receivables ÷ credit sales) × ${A.days}`)],
    [tr('المحصَّل في الفترة','Collected'),r2(A.coll),tr('الدفعات − الصكوك المرتجعة','payments − returned cheques')],[tr('المستحق في الفترة','Due in the period'),r2(A.due),tr('المحصّل + المتأخر في آخر الفترة','collected + overdue at period end')],
    [tr('نسبة التحصيل %','Collection ratio %'),A.collR==null?'':r2(A.collR),tr('المحصَّل ÷ المستحق','collected ÷ due')],[tr('استخدام سقف الائتمان %','Credit-limit use %'),A.util==null?'':r2(A.util),tr('الرصيد ÷ السقف','balance ÷ limit')],
    [tr('الزبائن المتجاوزون للسقف','Customers over limit'),A.overLim.length,tr('الرصيد > السقف','balance > limit')],[tr('تغطية الضمانات %','Guarantee cover %'),A.cov==null?'':r2(A.cov),tr('قيمة الضمانات السارية ÷ الرصيد','valid guarantees ÷ balance')],
    [tr('تركّز أكبر 10 مدينين %','Top-10 concentration %'),A.total?r2(A.top10*100/A.total):0,''],[tr('مخصص الديون المشكوك فيها','Doubtful-debt provision'),r2(A.prov),''],[tr('الديون المعدومة المشطوبة','Written off'),r2(A.wo),'']]
    .map(([k,v,f])=>({[K]:k,[V]:v,[F]:f}));
  const ag=BUCKETS.map(b=>({[tr('الفئة','Bucket')]:tr(b.ar,b.en),[tr('المعنى','Meaning')]:tr(b.m[0],b.m[1]),[tr('المبلغ','Amount')]:r2(A.ag[b.k])}));
  const cust=A.list.map(c=>Object.assign({'كود الزبون':c.a,'اسم الزبون':c.n,'الفئة':catName(c.cat),'القطاع':c.gov?'حكومي':'خاص','المنطقة':c.rg,'المستودع':c.wh,'حالة الحساب':tr(STATUS[c.st][0],STATUS[c.st][1]),
    'سقف الائتمان':c.lim,'مدة السداد':c.terms,'الرصيد':r2(c.bal),'استخدام السقف %':c.util==null?'':r2(c.util),'المتأخر':r2(c.over)},Object.fromEntries(BUCKETS.map(b=>[b.ar,r2(c.ag[b.k])])),
    {'الضمانات السارية':r2(c.gv),'تغطية الضمان %':c.cov==null?'':r2(c.cov),'المبيعات الآجلة':r2(c.sales),'المحصَّل':r2(c.coll),'DSO':c.dso==null?'':Math.round(c.dso),'آخر دفعة':c.lastPay,'الإجراء المقترح':c.step}));
  const top=A.top.map((c,i)=>({'#':i+1,'كود الزبون':c.a,'اسم الزبون':c.n,'الرصيد':r2(c.bal),'المتأخر':r2(c.over),'الحصة %':A.total?r2(c.bal*100/A.total):0}));
  const ol=A.overLim.map(c=>({'كود الزبون':c.a,'اسم الزبون':c.n,'السقف':c.lim,'الرصيد':r2(c.bal),'التجاوز':r2(c.bal-c.lim),'الاستخدام %':r2(c.util)}));
  const gv=A.list.filter(c=>c.gov).map(c=>({'كود الزبون':c.a,'الجهة':c.n,'النوع':tr(PUB[c.pubT][0],PUB[c.pubT][1]),'الرصيد':r2(c.bal),'المتأخر':r2(c.over),'مقاصّة منفذة':r2(sum(c.offs.filter(o=>o.st==='done'),'v')),'مقاصّة قيد الإجراء':r2(sum(c.offs.filter(o=>o.st==='pend'),'v'))}));
  const sheets=[[tr('المؤشرات','KPIs'),kp],[tr('أعمار الديون','Aging'),ag],[tr('الزبائن','Customers'),cust],[tr('أكبر 10 مدينين','Top 10'),top]];
  if(ol.length)sheets.push([tr('متجاوزو السقف','Over limit'),ol]);if(gv.length)sheets.push([tr('الجهات الحكومية','Government'),gv]);
  saveXlsx(sheets,`Debts-analysis-${P.to}.xlsx`)}
function filesDlg(){modal(tr('ملفات الديون المستوردة','Imported debts files'),`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الملف','File'))}</th><th class="nosort">${esc(tr('النوع','Kind'))}</th><th class="nosort">${esc(tr('الفترة','Period'))}</th><th class="nosort n">${esc(tr('الذمم','Receivables'))}</th><th class="nosort">${esc(tr('بواسطة','By'))}</th>${ST.canImport?'<th class="nosort"></th>':''}</tr></thead>
  <tbody>${ST.batches.slice().reverse().map(b=>`<tr><td>${esc(b.file||'—')}${b.ok===false?` <span class="pill p-crit">${esc(tr('غير مكتمل','incomplete'))}</span>`:''}</td><td>${esc(b.sheet==='ar'?tr('بيانات الديون','Debts data'):b.sheet==='ledger'?tr('كشف الأرصدة','Balances statement'):tr('بيان قيمة المحروقات (لم يعد يُعرض)','Fuel value (no longer shown)'))}</td><td>${fdi(b.from)} – ${fdi(b.to)}</td><td class="n">${big(b.value)}</td><td>${esc(b.by||'')}</td>${ST.canImport?`<td><button type="button" class="btn sm ghost" data-db-del="${esc(b.id)}">${esc(tr('حذف','Delete'))}</button></td>`:''}</tr>`).join('')}</tbody></table></div>`,
  `<button type="button" class="btn" data-dbt="reload">${esc(tr('تحديث','Refresh'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- events ---------- */
function after(){const u=UI();
  const bind=(sel,f)=>{const e=document.getElementById(sel);if(e)e.addEventListener('change',()=>{f(e);u.lim=25;render()})};
  bind('db-at',e=>{u.at=e.value});bind('db-cat',e=>{u.cat=e.value});bind('db-rg',e=>{u.rg=e.value});bind('db-wh',e=>{u.wh=e.value});bind('db-st',e=>{u.st=e.value});
  const q=document.getElementById('db-q');if(q)q.addEventListener('input',debounce(()=>{u.q=q.value;u.lim=25;const p=q.selectionStart;render();const n=document.getElementById('db-q');if(n){n.focus();try{n.setSelectionRange(p,p)}catch(e){}}},250));
  const f=document.getElementById('db-file');if(f)f.addEventListener('change',()=>{const file=f.files&&f.files[0];f.value='';if(file)importFile(file)})}
document.addEventListener('click',async e=>{
  const sc=e.target.closest('[data-db-sec]');if(sc){const u=UI();u.sec=sc.dataset.dbSec;u.cat='';u.lim=25;render();return}
  const so=e.target.closest('[data-db-sort]');if(so&&S.view===ID){const u=UI(),k=so.dataset.dbSort;if(u.sort===k)u.dir=u.dir==='asc'?'desc':'asc';else{u.sort=k;u.dir=k==='name'?'asc':'desc'}render();return}
  const ac=e.target.closest('[data-db-acct]');if(ac&&document.getElementById('view').contains(ac)){acctDetail(ac.dataset.dbAcct);return}
  const del=e.target.closest('[data-db-del]');if(del){if(!del.classList.contains('armed')){del.classList.add('armed');del.textContent=tr('تأكيد الحذف','Confirm delete');setTimeout(()=>{del.classList.remove('armed');del.textContent=tr('حذف','Delete')},4000);return}
    const b=ST.batches.find(x=>x.id===del.dataset.dbDel);if(!b)return;del.disabled=true;try{await delBatch(b);toast(tr('تم حذف الملف','File deleted'));await load(true);filesDlg();render()}catch(err){toast(tr('تعذّر الحذف: ','Could not delete: ')+(err.message||''));del.disabled=false}return}
  const a=e.target.closest('[data-dbt]');if(!a)return;const act=a.dataset.dbt;
  if(act==='import'){const f=document.getElementById('db-file');if(f)f.click();return}
  if(act==='template'){downloadTemplate();return}
  if(act==='template-help'){templateHelp();return}
  if(act==='export'){exportXlsx();return}
  if(act==='print'){try{window.print()}catch(e2){}return}
  if(act==='files'){filesDlg();return}
  if(act==='more'){UI().lim+=50;render();return}
  if(act==='clear'){Object.assign(UI(),{cat:'',rg:'',wh:'',st:'',q:''});render();return}
  if(act==='drop-local'){ST.local=null;UI().at='';render();return}
  if(act==='reload'){a.disabled=true;await load(true);dlg.close();render();toast(tr('تم تحديث البيانات','Data refreshed'));return}
  if(act==='view-local'){ST.local=ST.pending;ST.pending=null;dlg.close();Object.assign(UI(),{at:'',cat:'',rg:'',wh:'',st:'',q:''});render();return}
  if(act==='save'){const P=ST.pending;if(!P)return;const msg=document.getElementById('db-msg');a.disabled=true;msg.style.color='';msg.textContent=tr('جارٍ الحفظ…','Saving…');
    try{const rep=document.getElementById('db-replace');await saveOne(P,!!(rep&&rep.checked));ST.pending=null;ST.local=null;await load(true);dlg.close();Object.assign(UI(),{at:'',cat:'',rg:'',wh:'',st:'',q:''});render();toast(tr('تم حفظ الملف','File saved'))}
    catch(err){a.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية الحفظ في هذه اللوحة','You are not allowed to save to this dashboard'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}}
});
document.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;const ac=e.target.closest&&e.target.closest('[data-db-acct]');if(ac&&document.getElementById('view').contains(ac)){e.preventDefault();acctDetail(ac.dataset.dbAcct)}});

/* ---------- navigation: a task of Commercial Affairs & Collections, next to SMS055 ---------- */
if(window.I18N&&I18N.add)I18N.add({[LABEL[0]]:LABEL[1]});
VIEWS.push({id:ID,label:LABEL[0],icon:'debts',group:'مهام إدارة الشؤون التجارية والتحصيلات',nav:false});
ORG_TASKS.ca=(ORG_TASKS.ca||[]).concat([ID]);
const __render=render;
render=function(){if(S.view===ID){recompute();document.getElementById('view').innerHTML=dbBanner()+view();after();return}__render()};
const __rp=recPanel;
recPanel=function(unit){const h=__rp(unit);if(unit!==UNIT)return h;let body;
  if(!ST.loaded){load().then(()=>{if(S.view==='org')render()});body=`<span class="muted">${esc(tr('جارٍ التحميل…','Loading…'))}</span>`}
  else{const L1=ST.snaps.slice().sort((a,b)=>a.to.localeCompare(b.to));const P=L1[L1.length-1];
    if(!P)body=`<span class="muted">${esc(ST.err==='denied'?tr('البيانات متاحة للإدارة والإدارة العامة.','Data is available to the department and the General Administration.'):tr('لم تُستورد بيانات الديون بعد.','No debts data imported yet.'))}</span>`;
    else{const SC=sectors(P,{sec:'priv',q:''});body=`<b class="num">${big(SC.p.total+SC.g.total)}</b> ${esc(U())} · ${esc(tr('الخاص','private'))} ${big(SC.p.total)}${SC.p.overR!=null?` (${esc(tr('متأخر','overdue'))} ${P1(SC.p.overR)})`:''} · ${esc(tr('الحكومي','government'))} ${big(SC.g.total)} <small class="muted">(${esc(mLabel(P.to))})</small>`}}
  return `<section class="panel"><h3 class="section-title">${esc(tr('لوحة الديون','Debts dashboard'))} <small>${esc(tr('الذمم وأعمارها والتحصيلات وسقوف الائتمان','receivables, aging, collections and credit limits'))}</small></h3><div style="font-size:13px">${body}</div>
   <div class="cx-links"><button type="button" class="btn sm primary" data-org-go="${ID}">${esc(tr('فتح لوحة الديون','Open the debts dashboard'))}</button></div></section>`+h};
const __vod=viewOrgDuties;
viewOrgDuties=function(el){__vod(el);const sec=el.querySelector('#od-'+UNIT+' .od-acts');if(sec&&!sec.querySelector(`[data-org-go="${ID}"]`))sec.insertAdjacentHTML('afterbegin',`<button type="button" class="btn sm" data-org-go="${ID}">${esc(tr('لوحة الديون','Debts dashboard'))}</button>`)};
window.__debts={ST,parseFile,parseTemplate,analyze,prep,snaps,templateSheets,TPL,BUCKETS};
})();
