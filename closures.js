/* ===== لوحات مركز تقارير الإقفال: تحليل المبيعات، حركة التحويلات، أذونات الاستلام (إدارة الشؤون التجارية والتحصيلات) =====
   Loaded after sms055.js and built on the same pattern: the department downloads one of the three Excel exports of the
   sales closing-reports centre (Sales_Report, Transfers_Report, Receipts_Report), imports it, and the page turns it into
   KPIs, charts, a written reading, a sortable table and an Excel export. One dashboard (view) per report:
     sales      ← Sales_Report     (المنطقة، الشهر، المستودع، رقم الحساب، اسم الزبون، TRNSFR_THRU، رمز المنتج، اسم المنتج، PACK_CD، الوحدة، إجمالي الكمية)
     transfers  ← Transfers_Report (المستودع، المخزن، رمز المنتج، اسم المنتج، PRD_IND، PACK_CD، وحدة القياس، التاريخ، نوع الحركة، وصف الحركة، رقم المستند، ITEM_NO، الكمية السابقة، الكمية، PROCESS_COUNTER)
     receipts   ← Receipts_Report  (رقم المستند، التاريخ، المستودع، تاريخ المغادرة، تاريخ الوصول، رقم الإذن، نوع الاستلام، رقم المورد، اسم المورد، رمز المنتج، المنتج، وحدة القياس، الكمية)
   Quantities come in different units (litres, cylinders…), so every quantity figure is shown for one unit at a time:
   the unit filter defaults to the unit with the most lines, and picking a product switches to that product's unit.
   Storage mirrors sms055: Firestore collections `sales_report`, `transfers_report`, `receipts_report`, each with one
   `batch` document per imported file and `part` documents (<batchId>_p<n>) holding the rows as compact JSON.
   Readable by the department, the General Administration, the board and the administrator; written by the
   administrator or an editor of ca / gm (firestore.rules). The file readers and chart builders are shared from sms055.js. */
"use strict";
(function(){
const L=window.__smsLib;if(!L){console.warn('closures: sms055 library missing');return}
const {readGrids,hn,AD,num,serialIso,niceMax,colChart,bars,stack,pct,isoAdd,days,fdi,mName,me}=L;
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const UNIT='ca',PART_ROWS=4000,DEPT=['إدارة الشؤون التجارية والتحصيلات','Commercial Affairs & Collections'];
const qs=v=>{if(v==null||isNaN(v))return'—';const a=Math.abs(v);if(a>=1e6)return +(v/1e6).toFixed(a>=1e7?0:1)+tr(' مليون',' M');if(a>=1e4)return fn(Math.round(v/1e3))+tr(' ألف',' K');return fn(Math.round(v*10)/10)};
const q2=v=>fn(Math.round(v*100)/100);
const txt=v=>String(v==null?'':(v&&typeof v==='object'&&'serial' in v?serialIso(v.serial):v)).trim();

/* ---------- styles (on top of the .sm-* classes of sms055.js) ---------- */
(function(){const st=document.createElement('style');st.textContent=`
.cx-root.sm-root{--sm-c1:var(--cx-c1,#2a78d6);--sm-c2:var(--cx-c2,#eb6834);--sm-c3:var(--cx-c3,#1baf7a)}
.cx-sales .sm-hero{--hero-a:#3b1d8f;--hero-b:#6d3fd3;--hero-c:#4c2aa8}.cx-sales .cx-root{--cx-c1:#6d3fd3;--cx-c2:#eb6834;--cx-c3:#1baf7a}
.cx-transfers .sm-hero{--hero-a:#0b4f3a;--hero-b:#0f8a63;--hero-c:#0c6b4e}.cx-transfers .cx-root{--cx-c1:#0f8a63;--cx-c2:#2a78d6;--cx-c3:#eb6834}
.cx-receipts .sm-hero{--hero-a:#083a63;--hero-b:#0e73b6;--hero-c:#0a5a90}.cx-receipts .cx-root{--cx-c1:#0e73b6;--cx-c2:#eb6834;--cx-c3:#1baf7a}
.sm-tbl tr[data-cx-ent]{cursor:pointer}
.sm-tbl th[data-cx-sort]{cursor:pointer;white-space:nowrap}.sm-tbl th[data-cx-sort][aria-sort]::after{content:" ▾";font-size:10px}.sm-tbl th[data-cx-sort][aria-sort="ascending"]::after{content:" ▴"}
.cx-units{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12.5px;color:rgba(255,255,255,.85)}.cx-units b{font-variant-numeric:tabular-nums;color:#fff}
.cx-links{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
`;document.head.appendChild(st)})();

/* ---------- the three reports ---------- */
/* field: [key, aliases (normalised with hn), Arabic label, English label, type s=text n=number d=date m=month] */
const F_PROD=[['p',['رمز المنتج','رقم المنتج','كود المنتج','product_no','item_no','prod_cd','item_cd','product code'],'رمز المنتج','Product no.','s'],
  ['pn',['اسم المنتج','المنتج','الصنف','product','product_name','item','item_name'],'اسم المنتج','Product','s']];
const F_WH=[['w',['رقم المستودع','مستودع رقم','رمز المستودع','whse_no','warehouse_no','warehouse no','depot_no','depot no'],'رقم المستودع','Warehouse no.','s'],
  ['wn',['المستودع','اسم المستودع','warehouse','warehouse_name','depot','depot_name'],'المستودع','Warehouse','s']];
const F_UNIT=['u',['الوحدة','وحدة القياس','وحده','uom','unit','unit of measure'],'الوحدة','Unit','s'];
const F_PACK=['k',['pack_cd','pack','العبوه','كود العبوه'],'PACK_CD','PACK_CD','s'];
const F_QTY=['q',['الكمية','qty','quantity','الكميه الاجماليه','إجمالي الكمية','total qty','total_qty','total quantity'],'الكمية','Quantity','n'];
/* optional: the product temperature at reading (°C), used to correct litres / m³ to 15 °C (units.js) */
const F_TEMP=['tc',['درجة الحرارة','درجه الحراره','الحراره','الحرارة','temperature','temp','temp_c','obs temp','observed temperature'],'درجة الحرارة','Temperature','n'];
/* optional: the density measured at 15 °C for the line (kg/m³ or kg/l, ISO 3675 / ISO 12185); wins over the product setting */
const F_DENS=['dc',['الكثافة عند 15','الكثافه عند 15','الكثافة','الكثافه','density at 15','density15','density @15','density','dens'],'الكثافة عند 15°م','Density at 15 °C','n'];
const REPORTS={
 sales:{col:'sales_report',file:'Sales_Report',color:['#6d3fd3','#ddd6fe'],grain:'month',
  label:['تحليل المبيعات','Sales analysis'],head:['مخرجات تحليل المبيعات','Sales analysis outputs'],title:['لوحة مؤشرات تحليل المبيعات','Sales analysis dashboard'],
  sub:['المبيعات والزبائن والمستودعات','sales, customers and warehouses'],
  desc:['حوّل تقرير «تحليل المبيعات» (مركز تقارير الإقفال) إلى مؤشرات: الكميات الشهرية حسب المنتج والمستودع والمنطقة، وأعلى الزبائن، والزبائن الذين توقفوا عن الشراء.','Turn the “Sales analysis” report (closing-reports centre) into KPIs: monthly quantities by product, warehouse and region, top customers and customers who stopped buying.'],
  hint:['من «مركز تقارير الإقفال» في منظومة مبيعات البريقة حدّد تاريخ البداية والنهاية، ونزّل «تحليل المبيعات» بصيغة Excel، ثم اضغط «استيراد ملف تحليل المبيعات». يتعرّف الاستيراد على الأعمدة تلقائيًا.','In the Brega sales closing-reports centre pick the start and end dates, download “Sales analysis” as Excel, then press “Import sales analysis file”. Columns are recognised automatically.'],
  fields:[['rg',['المنطقة رقم','رقم المنطقة','منطقة رقم','region_no','region no','region code'],'المنطقة رقم','Region no.','s'],['rgn',['المنطقة','region','region_name','اسم المنطقة'],'المنطقة','Region','s'],
    ['mo',['الشهر','month','الفترة','period','yyyy-mm'],'الشهر','Month','m'],...F_WH,
    ['c',['رقم الحساب','رقم العميل','رقم الزبون','كود العميل','account_no','account no','acct_no','customer_no','cust_no','customer_id'],'رقم الحساب','Account no.','s'],
    ['cn',['اسم الزبون','اسم العميل','الزبون','العميل','customer','customer_name','cust_name'],'اسم الزبون','Customer','s'],
    ['th',['trnsfr_thru','transfer thru','trnsfr thru','transfer_thru'],'TRNSFR_THRU','TRNSFR_THRU','s'],...F_PROD,F_PACK,F_UNIT,F_QTY,F_TEMP,F_DENS],
  date:'mo',need:r=>!!(r.c||r.cn),
  ent:{k:'c',n:'cn',ar:'الزبون',en:'Customer',pl:['زبون','customers'],search:['بحث باسم الزبون أو رقم الحساب…','Search by customer name or account…']},
  filters:[{k:'rgn',ar:'المنطقة',en:'Region',all:['كل المناطق','All regions']},{k:'wn',ar:'المستودع',en:'Warehouse',all:['كل المستودعات','All warehouses']}],
  groups:[{k:'wn',kind:'bars',ar:'حسب المستودع',en:'By warehouse'},{k:'rgn',kind:'stack',ar:'حسب المنطقة',en:'By region'}],
  counters:[{k:'wn',ar:'المستودعات',en:'Warehouses'},{k:'rgn',ar:'المناطق',en:'Regions'}],
  verb:['اشتروا','bought'],stopped:['زبائن توقّفوا عن الشراء','Customers who stopped buying']},
 transfers:{col:'transfers_report',file:'Transfers_Report',color:['#0f8a63','#a7f3d0'],grain:'day',
  label:['حركة التحويلات','Transfers'],head:['مخرجات حركة التحويلات','Transfers outputs'],title:['لوحة مؤشرات حركة التحويلات','Transfers dashboard'],
  sub:['حركة المنتج بين المستودعات والمخازن','product movement across warehouses and stores'],
  desc:['حوّل تقرير «حركة التحويلات» (مركز تقارير الإقفال) إلى مؤشرات: الكميات اليومية حسب نوع الحركة والمستودع والمخزن والمنتج، وعدد المستندات.','Turn the “Transfers” report (closing-reports centre) into KPIs: daily quantities by movement type, warehouse, store and product, and the number of documents.'],
  hint:['من «مركز تقارير الإقفال» في منظومة مبيعات البريقة حدّد تاريخ البداية والنهاية، ونزّل «حركة التحويلات» بصيغة Excel، ثم اضغط «استيراد ملف حركة التحويلات». يتعرّف الاستيراد على الأعمدة تلقائيًا.','In the Brega sales closing-reports centre pick the start and end dates, download “Transfers” as Excel, then press “Import transfers file”. Columns are recognised automatically.'],
  fields:[...F_WH,['s',['رقم المخزن','المخزن رقم','رمز المخزن','رقم الخزان','store_no','store no','tank_no','tank no','location_no'],'رقم المخزن','Store no.','s'],
    ['sn',['وصف المخزن','اسم المخزن','المخزن','الخزان','store','store_name','tank','location'],'وصف المخزن','Store','s'],...F_PROD,
    ['pi',['prd_ind','product indicator'],'PRD_IND','PRD_IND','s'],F_PACK,F_UNIT,
    ['d',['التاريخ','تاريخ','تاريخ الحركة','date','trx_date','trans_date'],'التاريخ','Date','d'],
    ['mt',['نوع الحركة','رمز الحركة','trx_type','trans_type','movement type','movement_type'],'نوع الحركة','Movement type','s'],
    ['mn',['وصف الحركة','بيان الحركة','اسم الحركة','trx_desc','movement','movement desc','description'],'وصف الحركة','Movement','s'],
    ['doc',['رقم المستند','رقم الوثيقة','رقم السند','doc_no','document no','document_no','voucher no'],'رقم المستند','Document no.','s'],
    ['it',['item_no','رقم البند','رقم السطر','line_no'],'ITEM_NO','ITEM_NO','s'],
    ['pq',['الكمية السابقة','الكميه السابقه','prev_qty','previous qty','previous quantity','opening qty'],'الكمية السابقة','Previous qty','n'],
    F_QTY,F_TEMP,F_DENS,['pc',['process_counter','process counter'],'PROCESS_COUNTER','PROCESS_COUNTER','s']],
  date:'d',need:r=>!!(r.p||r.pn),
  ent:{k:'s',n:'sn',ar:'المخزن',en:'Store',pl:['مخزن','stores'],search:['بحث باسم المخزن أو رقمه…','Search by store name or number…'],parent:'wn'},
  filters:[{k:'mn',ar:'نوع الحركة',en:'Movement',all:['كل الحركات','All movements']},{k:'wn',ar:'المستودع',en:'Warehouse',all:['كل المستودعات','All warehouses']}],
  groups:[{k:'mn',kind:'stack',ar:'حسب نوع الحركة',en:'By movement type'},{k:'wn',kind:'bars',ar:'حسب المستودع',en:'By warehouse'}],
  counters:[{k:'wn',ar:'المستودعات',en:'Warehouses'},{k:'doc',ar:'المستندات',en:'Documents'},{k:'mn',ar:'أنواع الحركة',en:'Movement types'}],
  docs:'doc',verb:['تحرّكت','moved'],stopped:['مخازن بلا حركة في هذه الفترة','Stores with no movement this period']},
 receipts:{col:'receipts_report',file:'Receipts_Report',color:['#0e73b6','#bae6fd'],grain:'day',
  label:['أذونات الاستلام','Receipts'],head:['مخرجات أذونات الاستلام','Receipts outputs'],title:['لوحة مؤشرات أذونات الاستلام','Receipts dashboard'],
  sub:['الاستلام من المصادر الخارجية والمحلية','receipts from external and local sources'],
  desc:['حوّل تقرير «أذونات الاستلام» (مركز تقارير الإقفال) إلى مؤشرات: الكميات المستلمة يوميًا حسب المورد والمستودع والمنتج ونوع الاستلام، وعدد الأذونات، وزمن النقل بين المغادرة والوصول.','Turn the “Receipts” report (closing-reports centre) into KPIs: quantities received per day by supplier, warehouse, product and receipt type, the number of permits, and the transit time between departure and arrival.'],
  hint:['من «مركز تقارير الإقفال» في منظومة مبيعات البريقة حدّد تاريخ البداية والنهاية، ونزّل «أذونات الاستلام» بصيغة Excel، ثم اضغط «استيراد ملف أذونات الاستلام». يتعرّف الاستيراد على الأعمدة تلقائيًا.','In the Brega sales closing-reports centre pick the start and end dates, download “Receipts” as Excel, then press “Import receipts file”. Columns are recognised automatically.'],
  fields:[['doc',['رقم المستند','رقم الوثيقة','رقم السند','doc_no','document no','document_no','voucher no'],'رقم المستند','Document no.','s'],
    ['d',['التاريخ','تاريخ','تاريخ الاستلام','date','receipt date','receipt_date','trx_date'],'التاريخ','Date','d'],...F_WH,
    ['dd',['تاريخ المغادرة','تاريخ الخروج','تاريخ الشحن','departure date','departure_date','depart_date','ship date'],'تاريخ المغادرة','Departure date','d'],
    ['ad',['تاريخ الوصول','arrival date','arrival_date','arrive_date'],'تاريخ الوصول','Arrival date','d'],
    ['pm',['رقم الإذن','رقم إذن الاستلام','رقم اذن الاستلام','permit no','permit_no','receipt no','receipt_no','رقم الشحنة','رقم الرحلة'],'رقم الإذن','Permit no.','s'],
    ['rt',['نوع الاستلام','نوع الإذن','نوع المصدر','نوع','receipt type','receipt_type','source type','type'],'نوع الاستلام','Receipt type','s'],
    ['sp',['رقم المورد','رقم المصدر','رمز المورد','كود المورد','supplier_no','supplier no','vendor_no','vendor no','source_no'],'رقم المورد','Supplier no.','s'],
    ['spn',['اسم المورد','المورد','اسم المصدر','المصدر','الجهة الموردة','supplier','supplier_name','vendor','vendor_name','source'],'اسم المورد','Supplier','s'],
    ...F_PROD,F_UNIT,F_QTY,F_TEMP,F_DENS],
  date:'d',need:r=>!!(r.p||r.pn),
  ent:{k:'sp',n:'spn',ar:'المورد',en:'Supplier',pl:['مورد','suppliers'],search:['بحث باسم المورد أو رقمه…','Search by supplier name or number…']},
  filters:[{k:'rt',ar:'نوع الاستلام',en:'Receipt type',all:['كل الأنواع','All types']},{k:'wn',ar:'المستودع',en:'Warehouse',all:['كل المستودعات','All warehouses']}],
  groups:[{k:'wn',kind:'bars',ar:'حسب المستودع',en:'By warehouse'},{k:'rt',kind:'stack',ar:'حسب نوع الاستلام',en:'By receipt type'}],
  counters:[{k:'pm',ar:'الأذونات',en:'Permits'},{k:'wn',ar:'المستودعات',en:'Warehouses'}],
  docs:'pm',transit:['dd','ad'],verb:['ورّدوا','supplied'],stopped:['موردون لم يورّدوا في هذه الفترة','Suppliers with no receipts this period']}};
const IDS=Object.keys(REPORTS);
const ST={};IDS.forEach(id=>ST[id]={loaded:false,loading:null,err:'',batches:[],rows:[],local:null,pending:null,canImport:false});
const lbl=(cfg,k)=>{const f=cfg.fields.find(x=>x[0]===k);return f?tr(f[2],f[3]):k};
const nameKey=(cfg,k)=>k===cfg.ent.k?cfg.ent.n:k;

/* ---------- column recognition & parsing ---------- */
function mapHeader(cfg,row){const m={};const H=row.map(x=>hn(txt(x)));cfg.fields.forEach(f=>{const al=f[1].map(hn);const i=H.findIndex(h=>h&&al.includes(h));if(i>=0)m[f[0]]=i});return m}
function findHeader(cfg,grid){let best=null;for(let i=0;i<Math.min(grid.length,25);i++){const m=mapHeader(cfg,grid[i]||[]);const n=Object.keys(m).length;if(!best||n>best.n)best={i,m,n}}return best&&best.n>=4&&best.m.q!=null&&best.m[cfg.date]!=null?best:null}
function parseDate(v,dm){if(v&&typeof v==='object'&&'serial' in v)return serialIso(v.serial);if(typeof v==='number')return v>20000&&v<80000?serialIso(v):'';
  const s=AD(v).trim();let x=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(x)return x[1]+'-'+x[2].padStart(2,'0')+'-'+x[3].padStart(2,'0');
  x=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);if(x){let [a,b,y]=[+x[1],+x[2],+x[3]];if(y<100)y+=2000;const mo=dm?b:a,da=dm?a:b;if(mo<1||mo>12||da<1||da>31)return '';return y+'-'+String(mo).padStart(2,'0')+'-'+String(da).padStart(2,'0')}
  const n=parseFloat(s);if(!isNaN(n)&&/^\d+(\.\d+)?$/.test(s)&&n>20000&&n<80000)return serialIso(n);return ''}
/* month cells: "2026-08", "08/2026", "2026/8", "Aug-26", an Excel date or serial */
function parseMonth(v){if(v&&typeof v==='object'&&'serial' in v)return serialIso(v.serial).slice(0,7)+'-01';if(typeof v==='number')return v>20000&&v<80000?serialIso(v).slice(0,7)+'-01':(v>=190001&&v<=299912?String(v).slice(0,4)+'-'+String(v).slice(4)+'-01':'');
  const s=AD(v).trim();let x=s.match(/^(\d{4})[\/\-.](\d{1,2})(?:[\/\-.]\d{1,2})?/);if(x)return x[1]+'-'+x[2].padStart(2,'0')+'-01';
  x=s.match(/^(\d{1,2})[\/\-.](\d{4})$/);if(x)return x[2]+'-'+x[1].padStart(2,'0')+'-01';
  x=s.match(/^(\d{6})$/);if(x)return s.slice(0,4)+'-'+s.slice(4)+'-01';
  const d=parseDate(v,false);return d?d.slice(0,7)+'-01':''}
function parseRows(cfg,grid){const h=findHeader(cfg,grid);if(!h)return null;const m=h.m,body=grid.slice(h.i+1);const g=(r,k)=>m[k]==null?'':r[m[k]];
  let dm=false;body.forEach(r=>{const v=g(r,cfg.date);if(typeof v==='string'){const x=AD(v).match(/^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);if(x&&+x[1]>12)dm=true}});
  const rows=[];let skipped=0;
  body.forEach(r=>{if(!r||!r.some(x=>String(x&&typeof x==='object'?'1':x).trim()))return;const o={};
    cfg.fields.forEach(f=>{const k=f[0],v=g(r,k);if(f[4]==='n')o[k]=num(v);else if(f[4]==='d')o[k]=parseDate(v,dm);else if(f[4]==='m')o[k]=parseMonth(v);else o[k]=txt(v)});
    o.d=o[cfg.date]||'';if(!o.d||!o.q||!cfg.need(o)){skipped++;return}
    if(!o.p&&o.pn)o.p=o.pn;if(!o.pn&&o.p)o.pn=o.p;const e=cfg.ent;if(!o[e.k]&&o[e.n])o[e.k]=o[e.n];if(!o[e.n]&&o[e.k])o[e.n]=o[e.k];
    rows.push(o)});
  return {rows,skipped,map:m,header:h.i+1}}
async function parseFile(cfg,file){const sheets=await readGrids(file);let best=null;
  for(const s of sheets){const r=parseRows(cfg,s.grid);if(r&&(!best||r.rows.length>best.rows.length))best=Object.assign(r,{sheet:s.name})}
  if(!best)throw new Error(tr(`لم نجد أعمدة تقرير «${cfg.label[0]}» في الملف (${lbl(cfg,cfg.date)}، ${lbl(cfg,'q')}، ${lbl(cfg,cfg.ent.n)}، ${lbl(cfg,'pn')}…).`,`The “${cfg.label[1]}” columns (${lbl(cfg,cfg.date)}, ${lbl(cfg,'q')}, ${lbl(cfg,cfg.ent.n)}, ${lbl(cfg,'pn')}…) were not found in the file.`));
  if(!best.rows.length)throw new Error(tr('الأعمدة موجودة لكن لا توجد صفوف صالحة (تاريخ + كمية + منتج أو جهة).','Columns found, but no valid rows (date + quantity + product or party).'));
  const ds=best.rows.map(r=>r.d).sort();best.from=ds[0];best.to=cfg.grain==='month'?monthEnd(ds[ds.length-1]):ds[ds.length-1];best.file=file.name;return best}
function monthEnd(d){const y=+d.slice(0,4),m=+d.slice(5,7);return new Date(Date.UTC(y,m,0)).toISOString().slice(0,10)}

/* ---------- compact storage: dictionaries for text columns ---------- */
function encode(cfg,rows){const keys=cfg.fields.map(f=>f[0]),types=cfg.fields.map(f=>f[4]);const D={},idx={};keys.forEach(k=>{D[k]=[];idx[k]=new Map()});
  const R=rows.map(r=>keys.map((k,j)=>{const t=types[j];if(t==='n')return +r[k]||0;const v=String(r[k]||'');if(t!=='s')return v;let i=idx[k].get(v);if(i==null){i=D[k].length;D[k].push(v);idx[k].set(v,i)}return i}));
  const DD={};cfg.fields.forEach(f=>{if(f[4]==='s')DD[f[0]]=D[f[0]]});return JSON.stringify({F:keys,D:DD,R})}
function decode(cfg,s){const o=JSON.parse(s);const keys=o.F||[];const T={};cfg.fields.forEach(f=>T[f[0]]=f[4]);return (o.R||[]).map(a=>{const r={};keys.forEach((k,i)=>{const t=T[k],v=a[i];if(!t)return;r[k]=t==='n'?(+v||0):t==='s'?String((o.D[k]||[])[v]??''):String(v||'')});r.d=r[cfg.date]||'';return r})}
async function load(id,force){const S1=ST[id],cfg=REPORTS[id];if(S1.loading&&!force)return S1.loading;
  S1.loading=(async()=>{const FS=window.__fs&&window.__fs();if(!FS){S1.loaded=true;return}
    try{const q=await FS.collection(cfg.col).get();const docs=[];q.forEach(d=>docs.push(Object.assign({id:d.id},d.data())));
      const parts={};docs.filter(d=>d.kind==='part').forEach(p=>{(parts[p.batch]=parts[p.batch]||[]).push(p)});
      const batches=docs.filter(d=>d.kind==='batch').sort((a,b)=>(a.from||'').localeCompare(b.from||''));const rows=[];
      batches.forEach(b=>{const ps=(parts[b.id]||[]).sort((x,y)=>x.n-y.n);b.ok=ps.length===b.parts;if(!b.ok)return;ps.forEach(p=>{try{decode(cfg,p.d).forEach(r=>{r.b=b.id;rows.push(r)})}catch(e){b.ok=false}})});
      S1.batches=batches;S1.rows=rows;S1.err=''}
    catch(e){S1.err=e&&e.code==='permission-denied'?'denied':(e&&e.message)||'error';S1.batches=[];S1.rows=[]}
    S1.loaded=true;
    try{const u=await recMyUnit();const m=me();S1.canImport=m.role==='admin'||(!!S.canWrite&&m.role!=='viewer'&&(u===UNIT||u==='gm'))}catch(e){S1.canImport=false}
  })();return S1.loading}
function overlaps(id,P){return ST[id].batches.filter(b=>!(b.to<P.from||b.from>P.to))}
async function saveImport(id,P,replace){const cfg=REPORTS[id],FS=window.__fs(),m=me();const bid='b'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
  const chunks=[];for(let i=0;i<P.rows.length;i+=PART_ROWS)chunks.push(P.rows.slice(i,i+PART_ROWS));
  for(let i=0;i<chunks.length;i+=4){const b=FS.batch();chunks.slice(i,i+4).forEach((ch,j)=>b.set(FS.collection(cfg.col).doc(bid+'_p'+(i+j)),{kind:'part',batch:bid,n:i+j,d:encode(cfg,ch),createdBy:m.email}));await b.commit()}
  const tot=P.rows.reduce((a,r)=>a+r.q,0);
  await FS.collection(cfg.col).doc(bid).set({kind:'batch',file:String(P.file||'').slice(0,200),sheet:String(P.sheet||'').slice(0,100),from:P.from,to:P.to,rows:P.rows.length,parts:chunks.length,
    value:Math.round(tot*100)/100,returns:0,createdBy:m.email,by:String(m.name||m.email||'').slice(0,120),at:new Date().toISOString()});
  let removed=0;if(replace){for(const o of overlaps(id,P)){await delBatch(id,o,true);removed++}}
  try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'استيراد '+cfg.label[0],col:cfg.col,id:bid,summary:`${P.file} · ${P.from} – ${P.to} · ${P.rows.length} صف`+(removed?` · استُبدل ${removed} ملف`:'')})}catch(e){}
  return bid}
async function delBatch(id,b,quiet){const cfg=REPORTS[id],FS=window.__fs();const n=b.parts||0;const ops=[];for(let i=0;i<n;i++)ops.push(FS.collection(cfg.col).doc(b.id+'_p'+i));
  for(let i=0;i<ops.length;i+=400){const bt=FS.batch();ops.slice(i,i+400).forEach(r=>bt.delete(r));await bt.commit()}
  await FS.collection(cfg.col).doc(b.id).delete();
  if(!quiet){const m=me();try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'حذف ملف '+cfg.label[0],col:cfg.col,id:b.id,summary:`${b.file||''} · ${b.from} – ${b.to}`})}catch(e){}}}

/* ---------- analysis ---------- */
function UI(id){return S.ui['cx_'+id]||(S.ui['cx_'+id]={pre:'all',from:'',to:'',unit:'',prod:'',f:{},q:'',sort:'q',dir:'desc',lim:25})}
function raw(id){const s=ST[id];return s.local?s.local.rows:s.rows}
/* quantities shown in cubic metres (units.js): litres ÷ 1000, tonnes ÷ density; counts (cylinders) stay as they are */
function data(id){const r=raw(id);if(window.FuelUnits){FuelUnits.seen(r);return FuelUnits.normalize(r)}return r}
const qty=v=>fn(Math.round(v*100)/100);
function bounds(rows){if(!rows.length)return null;let a=rows[0].d,b=a;rows.forEach(r=>{if(r.d<a)a=r.d;if(r.d>b)b=r.d});return {min:a,max:b}}
function range(id){const cfg=REPORTS[id],u=UI(id),bd=bounds(data(id));if(!bd)return null;let from=bd.min,to=bd.max;
  if(cfg.grain==='month'){if(u.pre==='month')from=to;else if(u.pre==='3')from=addMonths(to,-2);else if(u.pre==='12')from=addMonths(to,-11);else if(u.pre==='custom'){from=(u.from||bd.min).slice(0,7)+'-01';to=(u.to||bd.max).slice(0,7)+'-01'}}
  else{if(u.pre==='day')from=to;else if(u.pre==='7')from=isoAdd(to,-6);else if(u.pre==='month')from=to.slice(0,8)+'01';else if(u.pre==='custom'){from=u.from||bd.min;to=u.to||bd.max}}
  if(from<bd.min&&u.pre!=='custom')from=bd.min;return {from,to,min:bd.min,max:bd.max}}
function addMonths(d,n){const y=+d.slice(0,4),m=+d.slice(5,7)-1+n;return new Date(Date.UTC(y,m,1)).toISOString().slice(0,10)}
/* the unit shown: the chosen product's unit, else the chosen unit, else the unit with most lines */
function unitOf(id){const u=UI(id),rows=data(id);if(u.prod){const r=rows.find(r=>r.p===u.prod);if(r&&r.u)return r.u}
  if(u.unit)return u.unit;const c=new Map();rows.forEach(r=>c.set(r.u||'',(c.get(r.u||'')||0)+1));let best='',n=-1;c.forEach((v,k)=>{if(v>n){n=v;best=k}});return best}
function pick(id,rows,from,to,unit){const cfg=REPORTS[id],u=UI(id);return rows.filter(r=>r.d>=from&&r.d<=to&&(r.u||'')===unit&&(!u.prod||r.p===u.prod)&&cfg.filters.every(f=>!u.f[f.k]||r[f.k]===u.f[f.k]))}
function agg(cfg,rows){const A={n:rows.length,q:0,ent:new Map(),prod:new Map(),day:new Map(),grp:{},docs:new Set(),tr:0,trn:0};
  cfg.groups.forEach(g=>A.grp[g.k]=new Map());cfg.counters.forEach(c=>A.grp[c.k]=A.grp[c.k]||new Map());const E=cfg.ent;
  rows.forEach(r=>{A.q+=r.q;
    const ek=r[E.k]||'—';let e=A.ent.get(ek);if(!e)A.ent.set(ek,e={id:ek,name:r[E.n]||ek,parent:E.parent?r[E.parent]||'':'',n:0,q:0,last:'',days:new Set(),prods:new Set(),docs:new Set(),tr:0,trn:0});
    e.n++;e.q+=r.q;if(r.d>e.last)e.last=r.d;e.days.add(r.d);e.prods.add(r.p);
    if(cfg.docs&&r[cfg.docs]){e.docs.add(r[cfg.docs]);A.docs.add(r[cfg.docs])}
    if(cfg.transit&&r[cfg.transit[0]]&&r[cfg.transit[1]]){const t=days(r[cfg.transit[0]],r[cfg.transit[1]])-1;if(t>=0&&t<365){e.tr+=t;e.trn++;A.tr+=t;A.trn++}}
    let p=A.prod.get(r.p);if(!p)A.prod.set(r.p,p={p:r.p,pn:r.pn,u:r.u,n:0,q:0,ent:new Set()});p.n++;p.q+=r.q;p.ent.add(ek);
    for(const k in A.grp){const gk=r[k]||'—';let g=A.grp[k].get(gk);if(!g)A.grp[k].set(gk,g={l:gk,n:0,q:0});g.n++;g.q+=r.q}
    let d=A.day.get(r.d);if(!d)A.day.set(r.d,d={d:r.d,q:0,n:0,ent:new Set()});d.q+=r.q;d.n++;d.ent.add(ek)});
  A.ents=[...A.ent.values()].sort((a,b)=>b.q-a.q);A.prods=[...A.prod.values()].sort((a,b)=>b.q-a.q);
  A.g={};for(const k in A.grp)A.g[k]=[...A.grp[k].values()].sort((a,b)=>b.q-a.q);return A}
function prodColors(id){const m=new Map();const tot=new Map();data(id).forEach(r=>tot.set(r.p,(tot.get(r.p)||0)+r.n));[...tot.entries()].sort((a,b)=>b[1]-a[1]).forEach(([p],i)=>m.set(p,i<3?'sm-k'+(i+1):'sm-k4'));return m}

/* ---------- page ---------- */
I['cx-sales']='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 20h16"/><rect x="6" y="11" width="3" height="6" rx="1"/><rect x="11" y="7" width="3" height="10" rx="1"/><rect x="16" y="4" width="3" height="13" rx="1"/></svg>';
I['cx-transfers']='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 8h13l-3-3M20 16H7l3 3"/></svg>';
I['cx-receipts']='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 8l8-4 8 4v9l-8 4-8-4z"/><path d="M4 8l8 4 8-4M12 12v9"/></svg>';
function heroHtml(cfg,sub,v,delta,unit,extra){return `<div class="sm-hero"><div><span class="sm-eyebrow">${esc(tr(DEPT[0],DEPT[1]))}</span>
  <h2>${esc(tr(cfg.title[0],cfg.title[1]))}</h2><p>${sub}</p>${extra||''}</div>${v==null?'':`<div class="sm-hero-v"><b>${v}</b><span>${esc(tr('إجمالي الكمية','Total quantity'))}${unit?' · '+esc(unit):''}</span>${delta||''}</div>`}</div>`}
function toolbar(id){const cfg=REPORTS[id],s=ST[id];return `<div class="toolbar sm-acts">
  ${s.canImport||!s.rows.length||s.err?`<button class="btn primary" data-cx="import" data-cx-id="${id}">${I.up}${esc(tr(`استيراد ملف ${cfg.label[0]}`,`Import ${cfg.label[1].toLowerCase()} file`))}</button>`:''}
  ${data(id).length?`<button class="btn" data-cx="export" data-cx-id="${id}">${I.down}${esc(tr('تصدير التحليل Excel','Export analysis (Excel)'))}</button><button class="btn" data-cx="print" data-cx-id="${id}">${esc(tr('طباعة / PDF','Print / PDF'))}</button>`:''}
  ${s.batches.length?`<button class="btn ghost" data-cx="files" data-cx-id="${id}">${esc(tr('الملفات المستوردة','Imported files'))} <span class="num">${s.batches.length}</span></button>`:''}
  ${s.local?`<button class="btn ghost" data-cx="drop-local" data-cx-id="${id}">${esc(tr('إغلاق العرض المؤقت','Close temporary view'))}</button>`:''}
  <input type="file" id="cx-file-${id}" accept="${FileImport.ACCEPT}" hidden></div>`}
function periodLabel(cfg,d){return cfg.grain==='month'?mName(+d.slice(5,7)-1)+' '+d.slice(0,4):fdi(d)}
function view(id){const cfg=REPORTS[id],s=ST[id];
  if(!s.loaded){load(id).then(()=>{if(S.view===id)render()});return `<div class="cx-${id}">`+heroHtml(cfg,esc(tr('جارٍ تحميل البيانات…','Loading data…')),null)+'<div class="panel muted">'+esc(tr('جارٍ التحميل…','Loading…'))+'</div></div>'}
  const rows=data(id);
  const head=`<div class="page-head" style="margin-bottom:8px"><div><h2 style="margin:0">${esc(tr(cfg.head[0],cfg.head[1]))}</h2></div>${toolbar(id)}</div>`;
  if(!rows.length){const denied=s.err==='denied';
    return `<div class="cx-${id}">`+head+heroHtml(cfg,esc(tr(cfg.desc[0],cfg.desc[1])),null)+
    `<div class="panel sm-empty">${I['cx-'+id].replace('<svg','<svg style="width:46px;height:46px;color:var(--accent)"')}<h3>${esc(denied?tr('هذه اللوحة لإدارة الشؤون التجارية والتحصيلات','This dashboard belongs to Commercial Affairs & Collections'):tr('لم يُستورد أي ملف بعد','No file imported yet'))}</h3>
     <p>${esc(denied?tr('لا تملك صلاحية الاطلاع على البيانات المحفوظة. يمكنك استيراد ملف وتحليله مؤقتًا على جهازك دون حفظه.','You cannot view the saved data. You can still import a file and analyse it on this device without saving it.'):tr(cfg.hint[0],cfg.hint[1]))}</p>
     <button class="btn primary" data-cx="import" data-cx-id="${id}">${I.up}${esc(tr(`استيراد ملف ${cfg.label[0]}`,`Import ${cfg.label[1].toLowerCase()} file`))}</button></div></div>`}
  const u=UI(id),R=range(id),unit=unitOf(id),cur=pick(id,rows,R.from,R.to,unit),A=agg(cfg,cur);
  const M=cfg.grain==='month';const len=M?(+R.to.slice(0,4)*12+ +R.to.slice(5,7))-(+R.from.slice(0,4)*12+ +R.from.slice(5,7))+1:days(R.from,R.to);
  const pFrom=M?addMonths(R.from,-len):isoAdd(R.from,-len),pTo=M?addMonths(R.from,-1):isoAdd(R.from,-1);
  const prevRows=pick(id,rows,pFrom,pTo,unit),B=prevRows.length?agg(cfg,prevRows):null;
  const dQ=B&&B.q?((A.q-B.q)*100/Math.abs(B.q)):null;
  const deltaHtml=dQ==null?'':`<span class="sm-delta">${dQ>=0?'▲':'▼'} ${Math.abs(Math.round(dQ*10)/10)}% ${esc(tr('عن الفترة السابقة','vs previous period'))}</span>`;
  const PC=prodColors(id);
  const units=uniq(rows.map(r=>r.u));const prods=[...new Map(rows.filter(r=>(r.u||'')===unit).map(r=>[r.p,r.pn])).entries()].sort((a,b)=>a[1].localeCompare(b[1],'ar'));
  const ents=A.ents,top10=ents.slice(0,10),top10q=top10.reduce((s,c)=>s+c.q,0),E=cfg.ent;
  const activeDays=A.day.size,avg=activeDays?A.q/activeDays:0,tp=A.prods[0];
  /* totals of every unit in the period (all products), for the hero line */
  const byUnit=new Map();rows.filter(r=>r.d>=R.from&&r.d<=R.to&&cfg.filters.every(f=>!u.f[f.k]||r[f.k]===u.f[f.k])).forEach(r=>byUnit.set(r.u||'',(byUnit.get(r.u||'')||0)+r.q));
  const unitsHtml=byUnit.size>1?`<div class="cx-units" style="margin-top:8px">${[...byUnit.entries()].sort((a,b)=>b[1]-a[1]).map(([k,v])=>`<span><b>${qs(v)}</b> ${esc(k||'—')}</span>`).join('')}</div>`:'';
  // trend
  let pts;if(M){pts=[];for(let d=R.from;d<=R.to;d=addMonths(d,1)){const x=A.day.get(d);pts.push({l:mName(+d.slice(5,7)-1).slice(0,3)+' '+d.slice(2,4),y:x?x.q:0,tip:`${periodLabel(cfg,d)} · ${qs(x?x.q:0)} ${unit} · ${x?x.ent.size:0} ${tr(E.pl[0],E.pl[1])}`})}}
  else if(len<=62){pts=[];for(let d=R.from;d<=R.to;d=isoAdd(d,1)){const x=A.day.get(d);pts.push({l:String(+d.slice(8)),y:x?x.q:0,tip:`${fdi(d)} · ${qs(x?x.q:0)} ${unit} · ${x?x.n:0} ${tr('حركة','lines')} · ${x?x.ent.size:0} ${tr(E.pl[0],E.pl[1])}`})}}
  else{const mm=new Map();cur.forEach(r=>{const k=r.d.slice(0,7);const o=mm.get(k)||{q:0,c:new Set()};o.q+=r.q;o.c.add(r[E.k]);mm.set(k,o)});
    pts=[...mm.keys()].sort().map(k=>{const o=mm.get(k);return {l:mName(+k.slice(5)-1).slice(0,3)+' '+k.slice(2,4),y:o.q,tip:`${mName(+k.slice(5)-1)} ${k.slice(0,4)} · ${qs(o.q)} ${unit} · ${o.c.size} ${tr(E.pl[0],E.pl[1])}`}})}
  const lost=B?B.ents.filter(c=>c.q>0&&!A.ent.has(c.id)).sort((a,b)=>b.q-a.q):[];const fresh=B?A.ents.filter(c=>!B.ent.has(c.id)):[];
  // written reading
  const ins=[];const prevLbl=`${periodLabel(cfg,pFrom)} – ${periodLabel(cfg,pTo)}`;
  if(dQ!=null)ins.push([dQ>=0?'o':'w',dQ>=0?'▲':'▼',tr(`إجمالي الكمية <b>${qs(A.q)}</b> ${esc(unit)}، ${dQ>=0?'بزيادة':'بانخفاض'} <b>${Math.abs(Math.round(dQ*10)/10)}%</b> عن الفترة السابقة المماثلة (${prevLbl}: ${qs(B.q)}).`,`Total quantity <b>${qs(A.q)}</b> ${esc(unit)}, ${dQ>=0?'up':'down'} <b>${Math.abs(Math.round(dQ*10)/10)}%</b> on the same-length previous period (${prevLbl}: ${qs(B.q)}).`)]);
  else ins.push(['','i',tr(`إجمالي الكمية <b>${qs(A.q)}</b> ${esc(unit)} خلال <b>${activeDays}</b> ${M?'شهر':'يوم'}، بمتوسط <b>${qs(avg)}</b> ${M?'شهريًا':'يوميًا'}. لا تتوفر بيانات للفترة السابقة للمقارنة؛ استورد ملف الفترة السابقة لتظهر المقارنة.`,`Total quantity <b>${qs(A.q)}</b> ${esc(unit)} over <b>${activeDays}</b> ${M?'months':'days'}, <b>${qs(avg)}</b> ${M?'a month':'a day'} on average. No previous-period data to compare; import it to see the comparison.`)]);
  if(tp&&A.q&&A.prods.length>1)ins.push(['','1',tr(`المنتج الأعلى: <b>${esc(tp.pn)}</b> بنسبة <b>${pct(tp.q,A.q)}%</b> من الكمية (${qs(tp.q)} ${esc(unit)}).`,`Top product: <b>${esc(tp.pn)}</b>, <b>${pct(tp.q,A.q)}%</b> of the quantity (${qs(tp.q)} ${esc(unit)}).`)]);
  if(ents.length>1){const sh=pct(top10q,A.q);ins.push([sh>=60&&ents.length>10?'w':'','10',tr(`أعلى 10 ${E.pl[1]==='customers'?'زبائن':E.pl[0]==='مورد'?'موردين':'مخازن'} يمثلون <b>${sh}%</b> من الكمية. الأكبر: <b>${esc(ents[0].name)}</b> (${qs(ents[0].q)} ${esc(unit)}).`,`The top 10 ${E.pl[1]} are <b>${sh}%</b> of the quantity. Largest: <b>${esc(ents[0].name)}</b> (${qs(ents[0].q)} ${esc(unit)}).`)])}
  cfg.groups.forEach(g=>{const list=A.g[g.k];if(list&&list.length>1)ins.push(['','·',tr(`${tr(g.ar,g.en)}: <b>${list.length}</b>، الأعلى <b>${esc(list[0].l)}</b> بنسبة <b>${pct(list[0].q,A.q)}%</b>.`,`${g.en}: <b>${list.length}</b>, led by <b>${esc(list[0].l)}</b> with <b>${pct(list[0].q,A.q)}%</b>.`)])});
  if(cfg.docs&&A.docs.size)ins.push(['','#',tr(`<b>${fn(A.docs.size)}</b> ${cfg.docs==='pm'?'إذن استلام':'مستند'} في الفترة، بمتوسط <b>${qs(A.q/A.docs.size)}</b> ${esc(unit)} لكل ${cfg.docs==='pm'?'إذن':'مستند'}.`,`<b>${fn(A.docs.size)}</b> ${cfg.docs==='pm'?'permits':'documents'} in the period, <b>${qs(A.q/A.docs.size)}</b> ${esc(unit)} each on average.`)]);
  if(cfg.transit&&A.trn){const slow=ents.filter(e=>e.trn&&e.tr/e.trn>A.tr/A.trn*1.5&&e.trn>=2);ins.push([slow.length?'w':'o','⏱',tr(`متوسط زمن النقل (من المغادرة إلى الوصول) <b>${Math.round(A.tr/A.trn*10)/10}</b> يوم${slow.length?`، وأبطأ من المتوسط بوضوح: <b>${slow.slice(0,3).map(e=>esc(e.name)).join('، ')}</b>`:''}.`,`Average transit time (departure to arrival) <b>${Math.round(A.tr/A.trn*10)/10}</b> days${slow.length?`; clearly slower than average: <b>${slow.slice(0,3).map(e=>esc(e.name)).join(', ')}</b>`:''}.`)])}
  if(lost.length)ins.push(['w','!',tr(`<b>${lost.length}</b> ${E.pl[0]} ${cfg.verb[0]} في الفترة السابقة ولا حركة لهم في هذه الفترة (${qs(lost.reduce((s,c)=>s+c.q,0))} ${esc(unit)} سابقًا). راجع القائمة أدناه للمتابعة.`,`<b>${lost.length}</b> ${E.pl[1]} ${cfg.verb[1]} in the previous period but have nothing in this one (${qs(lost.reduce((s,c)=>s+c.q,0))} ${esc(unit)} before). See the list below.`)]);
  if(fresh.length)ins.push(['o','+',tr(`<b>${fresh.length}</b> ${E.pl[0]} جديد مقارنة بالفترة السابقة.`,`<b>${fresh.length}</b> new ${E.pl[1]} compared with the previous period.`)]);
  const negQ=cur.filter(r=>r.q<0).length;if(negQ)ins.push(['w','−',tr(`<b>${fn(negQ)}</b> حركة بكمية سالبة (عكس قيد أو إرجاع) محسوبة ضمن الإجمالي.`,`<b>${fn(negQ)}</b> lines with a negative quantity (reversal or return) are included in the total.`)]);
  // table
  const q=norm(u.q||'');let tbl=A.ents.filter(c=>!q||norm(c.name+' '+c.id+' '+c.parent).includes(q));
  const key={q:c=>c.q,n:c=>c.n,name:c=>c.name,last:c=>c.last,days:c=>c.days.size,docs:c=>c.docs.size,tr:c=>c.trn?c.tr/c.trn:-1,prods:c=>c.prods.size}[u.sort]||(c=>c.q);
  tbl.sort((a,b)=>{const x=key(a),y=key(b);const r=typeof x==='string'?x.localeCompare(y,'ar'):x-y;return u.dir==='asc'?r:-r});const tblN=tbl.length;tbl=tbl.slice(0,u.lim);
  const th=(k,l,cls='')=>`<th class="${cls}" data-cx-sort="${k}" ${u.sort===k?`aria-sort="${u.dir==='asc'?'ascending':'descending'}"`:''}>${esc(l)}</th>`;
  const mxShare=Math.max(1,...ents.slice(0,1).map(c=>c.q));
  const sub=esc(tr(`الفترة ${periodLabel(cfg,R.from)} – ${periodLabel(cfg,R.to)} · ${fn(A.n)} حركة · ${A.ents.length} ${E.pl[0]} · البيانات المتاحة ${periodLabel(cfg,R.min)} – ${periodLabel(cfg,R.max)}`,`Period ${periodLabel(cfg,R.from)} – ${periodLabel(cfg,R.to)} · ${fn(A.n)} lines · ${A.ents.length} ${E.pl[1]} · data available ${periodLabel(cfg,R.min)} – ${periodLabel(cfg,R.max)}`));
  const presets=M?[['all',tr('كل البيانات','All data')],['12',tr('آخر 12 شهرًا','Last 12 months')],['3',tr('آخر 3 أشهر','Last 3 months')],['month',tr('آخر شهر','Latest month')],['custom',tr('مخصّص','Custom')]]
    :[['all',tr('كل البيانات','All data')],['month',tr('آخر شهر','Latest month')],['7',tr('آخر 7 أيام','Last 7 days')],['day',tr('آخر يوم','Latest day')],['custom',tr('مخصّص','Custom')]];
  const filtered=u.prod||u.unit||u.pre!=='all'||Object.values(u.f).some(Boolean);const unitNote=window.FuelUnits?FuelUnits.explain(cur):'';
  const barsOf=(list,go)=>bars(list.slice(0,10).map(c=>({l:c.name||c.l,v:c.q,go:go?c.id:'',tip:`${c.name||c.l}: ${q2(c.q)} ${unit} · ${pct(c.q,A.q)}%`})),{fmt:qs});
  const stackOf=list=>stack(list.slice(0,3).map((t,i)=>({l:t.name||t.l,v:t.q,k:'sm-k'+(i+1)})).concat(list.length>3?[{l:tr('أخرى','Other'),v:list.slice(3).reduce((s,t)=>s+t.q,0),k:'sm-k4'}]:[]));
  const stackHtml=h=>h.replace(/ د\.ل \(/g,' '+unit+' (').replace(/ LYD \(/g,' '+unit+' (');
  return `<div class="cx-${id}">`+head+(s.local?`<div class="banner info">${esc(tr(`عرض مؤقت للملف «${s.local.file}» على هذا الجهاز فقط، ولم يُحفظ في المنظومة.`,`Temporary view of “${s.local.file}” on this device only; not saved.`))}</div>`:'')+
  heroHtml(cfg,sub,qs(A.q),deltaHtml,unit,unitsHtml)+
  `<div class="sm-bar">
    <div class="an-f">${esc(tr('الفترة','Period'))}<div class="sm-presets">${presets.map(([k,l])=>`<button type="button" data-cx-pre="${k}" data-cx-id="${id}" aria-pressed="${u.pre===k}">${esc(l)}</button>`).join('')}</div></div>
    ${u.pre==='custom'?(M?`<label class="an-f">${esc(tr('من','From'))}<input type="month" id="cx-from" value="${esc(R.from.slice(0,7))}" min="${R.min.slice(0,7)}" max="${R.max.slice(0,7)}"></label><label class="an-f">${esc(tr('إلى','To'))}<input type="month" id="cx-to" value="${esc(R.to.slice(0,7))}" min="${R.min.slice(0,7)}" max="${R.max.slice(0,7)}"></label>`
      :`<label class="an-f">${esc(tr('من','From'))}<input type="date" id="cx-from" value="${esc(R.from)}" min="${R.min}" max="${R.max}"></label><label class="an-f">${esc(tr('إلى','To'))}<input type="date" id="cx-to" value="${esc(R.to)}" min="${R.min}" max="${R.max}"></label>`):''}
    ${units.length>1?`<label class="an-f">${esc(tr('الوحدة','Unit'))}<select id="cx-unit" ${u.prod?'disabled':''}>${units.map(x=>`<option value="${esc(x)}" ${x===unit?'selected':''}>${esc(x)}</option>`).join('')}</select></label>`:''}
    <label class="an-f">${esc(tr('المنتج','Product'))}<select id="cx-prod"><option value="">${esc(tr('كل منتجات الوحدة','All products of the unit'))}</option>${prods.map(([p,n])=>`<option value="${esc(p)}" ${u.prod===p?'selected':''}>${esc(n)}${n!==p?` (${esc(p)})`:''}</option>`).join('')}</select></label>
    ${cfg.filters.map(f=>{const vals=uniq(rows.map(r=>r[f.k]));return vals.length>1?`<label class="an-f">${esc(tr(f.ar,f.en))}<select data-cx-f="${f.k}"><option value="">${esc(tr(f.all[0],f.all[1]))}</option>${vals.map(v=>`<option ${u.f[f.k]===v?'selected':''}>${esc(v)}</option>`).join('')}</select></label>`:''}).join('')}
    ${filtered?`<button class="btn sm ghost" data-cx="clear" data-cx-id="${id}">${esc(tr('مسح التصفية','Clear filters'))}</button>`:''}
  </div>
  ${unitNote?`<p class="sm-note">${esc(unitNote)}</p>`:''}
  ${!A.n?`<div class="panel empty"><b>${esc(tr('لا توجد حركات','No lines'))}</b>${esc(tr('ضمن التصفية الحالية.','for the current filter.'))}</div>`:`
  <div class="kpis an-kpis">
   ${anKpi(tr('إجمالي الكمية','Total quantity'),qs(A.q),esc(`${unit} · ${fn(A.n)} ${tr('حركة','lines')}`))}
   ${anKpi(tr(E.pl[1]==='customers'?'الزبائن النشطون':E.pl[0]==='مورد'?'الموردون':'المخازن النشطة',E.pl[1]==='customers'?'Active customers':E.pl[0]==='مورد'?'Suppliers':'Active stores'),fn(A.ents.length),esc(B?tr(`${fresh.length} جديد · ${lost.length} توقّف`,`${fresh.length} new · ${lost.length} stopped`):tr(`متوسط ${qs(A.ents.length?A.q/A.ents.length:0)} ${unit}`,`avg ${qs(A.ents.length?A.q/A.ents.length:0)} ${unit}`)))}
   ${anKpi(tr(M?'متوسط الكمية الشهرية':'متوسط الكمية اليومية',M?'Average monthly quantity':'Average daily quantity'),qs(avg),esc(`${unit} · ${activeDays} ${tr(M?'شهر':'يوم','active '+(M?'months':'days'))}`))}
   ${cfg.counters.map(c=>{const n=c.k===cfg.docs?A.docs.size:(A.g[c.k]||[]).filter(x=>x.l!=='—').length;const top=(A.g[c.k]||[])[0];return anKpi(tr(c.ar,c.en),fn(n),esc(c.k===cfg.docs?tr(`متوسط ${qs(n?A.q/n:0)} ${unit} لكل مستند`,`avg ${qs(n?A.q/n:0)} ${unit} each`):top?tr(`الأعلى: ${top.l}`,`Top: ${top.l}`):''))}).join('')}
   ${cfg.transit?anKpi(tr('متوسط زمن النقل','Average transit'),A.trn?(Math.round(A.tr/A.trn*10)/10)+' '+tr('يوم','days'):'—',esc(tr(`من المغادرة إلى الوصول · ${fn(A.trn)} حركة`,`departure to arrival · ${fn(A.trn)} lines`))):''}
   ${anKpi(tr('المنتجات','Products'),fn(A.prods.length),esc(tp?tr(`الأعلى: ${tp.pn}`,`Top: ${tp.pn}`):''))}
  </div>
  <div class="an-grid sm-root cx-root">
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('قراءة تحليلية','Reading'))} <small>${esc(tr('تُكتب تلقائيًا من البيانات المصفّاة','written from the filtered data'))}</small></h3><ul class="sm-ins">${ins.map(x=>`<li><span class="ic ${x[0]}">${x[1]}</span><span>${x[2]}</span></li>`).join('')}</ul></section>
   <section class="panel an-wide"><h3 class="section-title">${esc(M||len>62?tr('الكمية الشهرية','Monthly quantity'):tr('الكمية اليومية','Daily quantity'))} <small>${esc(unit)} · ${esc(tr('مرّر على العمود للتفاصيل','hover a column for details'))}</small></h3>${colChart(pts,tr('الكمية حسب الفترة','Quantity by period'))}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('المنتجات','Products'))} <small>${esc(tr(`الكمية بوحدة ${unit}`,`quantity in ${unit}`))}</small></h3>
    ${stackHtml(stack(A.prods.slice(0,3).map(p=>({l:p.pn,v:p.q,k:PC.get(p.p)||'sm-k4'})).concat(A.prods.length>3?[{l:tr('منتجات أخرى','Other products'),v:A.prods.slice(3).reduce((s,p)=>s+p.q,0),k:'sm-k4'}]:[])))}
    <div class="sm-prod">${A.prods.map(p=>`<div class="sm-pc"><h4><i class="${PC.get(p.p)||'sm-k4'}"></i>${esc(p.pn)} ${p.p!==p.pn?`<small class="muted num">${esc(p.p)}</small>`:''}</h4>
      <span class="q">${qty(p.q)} <small style="font-size:13px;font-weight:400">${esc(p.u||'')}</small></span>
      <span class="s">${esc(tr('الحصة','Share'))} <b class="num">${pct(p.q,A.q)}%</b> · ${fn(p.n)} ${esc(tr('حركة','lines'))} · ${p.ent.size} ${esc(tr(E.pl[0],E.pl[1]))}</span></div>`).join('')}</div></section>
   <section class="panel"><h3 class="section-title">${esc(tr(`أعلى 10 ${E.pl[1]==='customers'?'زبائن':E.pl[0]==='مورد'?'موردين':'مخازن'}`,`Top 10 ${E.pl[1]}`))} <small>${esc(unit)} · ${esc(tr('اضغط للتفاصيل','click for details'))}</small></h3>${barsOf(top10,true)}</section>
   ${cfg.groups.map(g=>{const list=A.g[g.k]||[];return `<section class="panel"><h3 class="section-title">${esc(tr(g.ar,g.en))} <small>${esc(unit)}</small></h3>${list.length?(g.kind==='stack'?stackHtml(stackOf(list)):barsOf(list,false)):`<div class="empty" style="padding:20px">${esc(tr('لا توجد بيانات','No data'))}</div>`}</section>`}).join('')}
   ${B?`<section class="panel an-wide"><h3 class="section-title">${esc(tr(cfg.stopped[0],cfg.stopped[1]))} <small>${esc(tr(`${cfg.verb[0]} في ${prevLbl}`,`${cfg.verb[1]} in ${prevLbl}`))}</small></h3>${lost.length?bars(lost.slice(0,10).map(c=>({l:c.name,v:c.q,go:c.id,tip:`${c.name}: ${q2(c.q)} ${unit} ${tr('في الفترة السابقة','in the previous period')} · ${tr('آخر حركة','last')} ${periodLabel(cfg,c.last)}`})),{tone:'fl',fmt:qs}):`<div class="empty" style="padding:20px">${esc(tr('كل جهات الفترة السابقة لها حركة في هذه الفترة','Every previous-period party appears again'))}</div>`}</section>`:''}
   <section class="panel an-wide"><h3 class="section-title">${esc(tr(E.ar==='الزبون'?'الزبائن':E.ar==='المورد'?'الموردون':'المخازن',E.pl[1][0].toUpperCase()+E.pl[1].slice(1)))} <small>${tblN} ${esc(tr(`${E.pl[0]} · اضغط على الصف للتفاصيل`,`${E.pl[1]} · click a row for details`))}</small></h3>
    <div class="toolbar" style="margin:0 0 8px"><div class="grow od-search">${I.search}<input type="search" id="cx-q" value="${esc(u.q||'')}" placeholder="${esc(tr(E.search[0],E.search[1]))}" aria-label="${esc(tr('بحث','Search'))}"></div></div>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">#</th>${th('name',tr(E.ar,E.en))}${E.parent?`<th class="nosort">${esc(lbl(cfg,E.parent))}</th>`:''}${th('n',tr('الحركات','Lines'),'n')}${th('days',tr(M?'الأشهر':'الأيام',M?'Months':'Days'),'n')}${cfg.docs?th('docs',tr(cfg.docs==='pm'?'الأذونات':'المستندات',cfg.docs==='pm'?'Permits':'Documents'),'n'):''}${th('prods',tr('المنتجات','Products'),'n')}${th('q',tr('الكمية','Quantity'),'n')}<th class="nosort n">${esc(tr('الحصة','Share'))}</th>${cfg.transit?th('tr',tr('زمن النقل','Transit'),'n'):''}${th('last',tr('آخر حركة','Last'),'n')}</tr></thead>
    <tbody>${tbl.map((c,i)=>`<tr data-cx-ent="${esc(c.id)}" data-cx-id="${id}" tabindex="0"><td class="n">${i+1}</td><td class="t-strong">${esc(c.name)} ${c.id!==c.name?`<small class="muted num">${esc(c.id)}</small>`:''}</td>${E.parent?`<td>${esc(c.parent)}</td>`:''}<td class="n">${fn(c.n)}</td><td class="n">${c.days.size}</td>${cfg.docs?`<td class="n">${fn(c.docs.size)}</td>`:''}<td class="n">${c.prods.size}</td><td class="n"><b>${q2(c.q)}</b></td><td class="n">${pct(c.q,A.q)}%<span class="sm-share" style="width:${Math.max(0,c.q)*40/mxShare}px"></span></td>${cfg.transit?`<td class="n">${c.trn?Math.round(c.tr/c.trn*10)/10+' '+tr('يوم','d'):'—'}</td>`:''}<td class="n">${periodLabel(cfg,c.last)}</td></tr>`).join('')}</tbody></table></div>
    ${tblN>u.lim?`<div style="text-align:center;margin-top:8px"><button class="btn sm" data-cx="more" data-cx-id="${id}">${esc(tr(`عرض المزيد (${tblN-u.lim} متبقٍ)`,`Show more (${tblN-u.lim} left)`))}</button></div>`:''}</section>
  </div>`}</div>`}

function after(id){const u=UI(id);
  const bind=(sel,fn2)=>{const e=document.querySelector(sel);if(e)e.addEventListener('change',()=>{fn2(e);u.lim=25;render()})};
  bind('#cx-prod',e=>{u.prod=e.value});bind('#cx-unit',e=>{u.unit=e.value;u.prod=''});
  bind('#cx-from',e=>{u.from=e.value.length===7?e.value+'-01':e.value});bind('#cx-to',e=>{u.to=e.value.length===7?e.value+'-01':e.value});
  document.querySelectorAll('#view select[data-cx-f]').forEach(e=>e.addEventListener('change',()=>{u.f[e.dataset.cxF]=e.value;u.lim=25;render()}));
  const q=document.getElementById('cx-q');if(q)q.addEventListener('input',debounce(()=>{u.q=q.value;u.lim=25;const p=q.selectionStart;render();const n=document.getElementById('cx-q');if(n){n.focus();try{n.setSelectionRange(p,p)}catch(e){}}},250));
  const f=document.getElementById('cx-file-'+id);if(f)f.addEventListener('change',()=>{const file=f.files&&f.files[0];f.value='';if(file)importFile(id,file)})}

/* ---------- import ---------- */
async function importFile(id,file){const cfg=REPORTS[id],s=ST[id];toast(tr('جارٍ قراءة الملف…','Reading the file…'));let P;
  try{P=await parseFile(cfg,file)}catch(e){if(e&&e.cancelled)return;modal(tr('تعذّر الاستيراد','Import failed'),`<p style="line-height:1.8">${esc(e.message||String(e))}</p>`,`<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`);return}
  s.pending=P;const ov=s.local?[]:overlaps(id,P);const es=new Set(P.rows.map(r=>r[cfg.ent.k])).size;
  const byU=new Map();P.rows.forEach(r=>byU.set(r.u||'',(byU.get(r.u||'')||0)+r.q));
  modal(tr(`مراجعة ملف ${cfg.label[0]} قبل الحفظ`,`Review the ${cfg.label[1].toLowerCase()} file`),`
   <div class="sm-mini"><div><b class="num">${fn(P.rows.length)}</b><span>${esc(tr('حركة مقروءة','lines read'))}</span></div><div><b class="num">${periodLabel(cfg,P.from)} – ${periodLabel(cfg,P.to)}</b><span>${esc(tr('الفترة','Period'))}</span></div>
    ${[...byU.entries()].sort((a,b)=>b[1]-a[1]).slice(0,2).map(([k,v])=>`<div><b class="num">${qs(v)}</b><span>${esc(k||tr('الكمية','quantity'))}</span></div>`).join('')}<div><b class="num">${fn(es)}</b><span>${esc(tr(cfg.ent.pl[0],cfg.ent.pl[1]))}</span></div></div>
   <p class="muted" style="margin:4px 0;font-size:12.5px">${esc(tr(`الملف: ${P.file} · الورقة: ${P.sheet} · صف العناوين: ${P.header}`,`File: ${P.file} · sheet: ${P.sheet} · header row: ${P.header}`))}${P.skipped?` · <b style="color:var(--warn)">${esc(tr(`تُجوهل ${P.skipped} صف بلا تاريخ أو كمية أو جهة`,`${P.skipped} rows skipped (no date, quantity or party)`))}</b>`:''}</p>
   ${window.FuelUnits?`<p class="muted" style="margin:0 0 4px;font-size:12px">${esc(tr('الكميات أعلاه بوحدات الملف الأصلية، وتُعرض في اللوحة بالمتر المكعب.','Quantities above are in the file’s own units; the dashboard shows them in cubic metres.'))}</p>`:''}
   <div class="sm-cols-map">${cfg.fields.filter(f=>(f!==F_TEMP&&f!==F_DENS)||P.map[f[0]]!=null).map(f=>`<span class="${P.map[f[0]]!=null?'ok':'no'}">${P.map[f[0]]!=null?'✓':'✗'} ${esc(tr(f[2],f[3]))}</span>`).join('')}</div>
   ${ov.length?`<div class="banner warn" style="margin-top:8px">${esc(tr(`تتداخل هذه الفترة مع ${ov.length} ملف محفوظ سابقًا:`,`This period overlaps ${ov.length} saved file(s):`))} ${ov.map(b=>esc(`${b.file} (${fdi(b.from)} – ${fdi(b.to)})`)).join('، ')}
     <label style="display:flex;gap:6px;align-items:center;margin-top:6px"><input type="checkbox" id="cx-replace" checked> ${esc(tr('استبدال الملفات المتداخلة بهذا الملف (يمنع تكرار الأرقام)','Replace the overlapping files with this one (avoids double counting)'))}</label></div>`:''}
   ${!s.canImport?`<p class="muted" style="font-size:12.5px">${esc(tr('الحفظ في المنظومة متاح لمحرري إدارة الشؤون التجارية والتحصيلات والإدارة العامة ومدير النظام. يمكنك عرض التحليل مؤقتًا على جهازك.','Saving is for editors of Commercial Affairs & Collections, the General Administration and the administrator. You can view the analysis on this device.'))}</p>`:''}
   <p id="cx-msg" style="font-size:12.5px;margin:6px 0 0"></p>`,
   `${s.canImport?`<button type="button" class="btn primary" data-cx="save" data-cx-id="${id}">${esc(tr('حفظ في المنظومة','Save to the system'))}</button>`:''}<button type="button" class="btn" data-cx="view-local" data-cx-id="${id}">${esc(tr('عرض دون حفظ','View without saving'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}

/* ---------- entity detail ---------- */
function entDetail(id,ek){const cfg=REPORTS[id],R=range(id),unit=unitOf(id),rows=pick(id,data(id),R.from,R.to,unit).filter(r=>(r[cfg.ent.k]||'—')===ek);if(!rows.length)return;const A=agg(cfg,rows),x=A.ents[0];
  const byDay=[...A.day.values()].sort((a,b)=>b.d.localeCompare(a.d));
  modal(`${x.name}${x.id!==x.name?' ('+x.id+')':''}`,`<div class="sm-mini"><div><b class="num">${q2(x.q)}</b><span>${esc(unit)}</span></div><div><b class="num">${fn(x.n)}</b><span>${esc(tr('حركة','lines'))}</span></div><div><b class="num">${x.days.size}</b><span>${esc(tr(cfg.grain==='month'?'شهر':'يوم',cfg.grain==='month'?'months':'days'))}</span></div>${cfg.docs?`<div><b class="num">${fn(x.docs.size)}</b><span>${esc(tr(cfg.docs==='pm'?'إذن':'مستند',cfg.docs==='pm'?'permits':'documents'))}</span></div>`:''}${cfg.transit&&x.trn?`<div><b class="num">${Math.round(x.tr/x.trn*10)/10}</b><span>${esc(tr('يوم نقل في المتوسط','avg transit days'))}</span></div>`:''}<div><b class="num">${periodLabel(cfg,x.last)}</b><span>${esc(tr('آخر حركة','last'))}</span></div></div>
   <h4 style="margin:8px 0 4px">${esc(tr('حسب المنتج','By product'))}</h4><div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('الحركات','Lines'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th><th class="nosort n">${esc(tr('الحصة','Share'))}</th></tr></thead><tbody>${A.prods.map(p=>`<tr><td>${esc(p.pn)}</td><td class="n">${fn(p.n)}</td><td class="n">${q2(p.q)} ${esc(p.u)}</td><td class="n">${pct(p.q,A.q)}%</td></tr>`).join('')}</tbody></table></div>
   <h4 style="margin:10px 0 4px">${esc(tr(cfg.grain==='month'?'حسب الشهر':'حسب اليوم',cfg.grain==='month'?'By month':'By day'))}</h4><div class="tablewrap" style="max-height:260px;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('التاريخ','Date'))}</th><th class="nosort n">${esc(tr('الحركات','Lines'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th></tr></thead><tbody>${byDay.map(d=>`<tr><td class="n">${periodLabel(cfg,d.d)}</td><td class="n">${d.n}</td><td class="n">${q2(d.q)}</td></tr>`).join('')}</tbody></table></div>`,
   `<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- export, print, files ---------- */
function exportXlsx(id){const cfg=REPORTS[id],R=range(id),unit=unitOf(id),rows=pick(id,data(id),R.from,R.to,unit),A=agg(cfg,rows),E=cfg.ent;const K=tr('البند','Item'),V=tr('القيمة','Value');
  const sum=[{[K]:tr('الفترة','Period'),[V]:`${R.from} – ${R.to}`},{[K]:tr('الوحدة','Unit'),[V]:unit},{[K]:tr('إجمالي الكمية','Total quantity'),[V]:Math.round(A.q*100)/100},{[K]:tr('عدد الحركات','Lines'),[V]:A.n},
    {[K]:tr(E.ar,E.en),[V]:A.ents.length},{[K]:tr('المنتجات','Products'),[V]:A.prods.length}];
  if(cfg.docs)sum.push({[K]:tr(cfg.docs==='pm'?'الأذونات':'المستندات',cfg.docs==='pm'?'Permits':'Documents'),[V]:A.docs.size});if(cfg.transit)sum.push({[K]:tr('متوسط زمن النقل (يوم)','Average transit (days)'),[V]:A.trn?Math.round(A.tr/A.trn*10)/10:''});
  const prod=A.prods.map(p=>({'رمز المنتج':p.p,'اسم المنتج':p.pn,'الوحدة':p.u,'الحركات':p.n,['الكمية ('+unit+')']:Math.round(p.q*1000)/1000,'الحصة %':pct(p.q,A.q),['عدد '+E.pl[0]]:p.ent.size}));
  const ent=A.ents.map((c,i)=>{const o={'م':i+1,['رقم '+E.ar]:c.id,[E.ar]:c.name};if(E.parent)o[lbl(cfg,E.parent)]=c.parent;Object.assign(o,{'الحركات':c.n,'الأيام':c.days.size});if(cfg.docs)o[cfg.docs==='pm'?'الأذونات':'المستندات']=c.docs.size;Object.assign(o,{'المنتجات':c.prods.size,['الكمية ('+unit+')']:Math.round(c.q*1000)/1000,'الحصة %':pct(c.q,A.q)});if(cfg.transit)o['متوسط زمن النقل']=c.trn?Math.round(c.tr/c.trn*10)/10:'';o['آخر حركة']=c.last;return o});
  const grp=[];cfg.groups.forEach(g=>(A.g[g.k]||[]).forEach(x=>grp.push({'التصنيف':g.ar,'البند':x.l,'الحركات':x.n,'الكمية':Math.round(x.q*1000)/1000,'الحصة %':pct(x.q,A.q)})));
  const day=[...A.day.values()].sort((a,b)=>a.d.localeCompare(b.d)).map(d=>({[cfg.grain==='month'?'الشهر':'التاريخ']:cfg.grain==='month'?d.d.slice(0,7):d.d,'الحركات':d.n,[E.ar]:d.ent.size,'الكمية':Math.round(d.q*1000)/1000}));
  saveXlsx([[tr('ملخص','Summary'),sum],[tr('المنتجات','Products'),prod],[tr(E.ar,E.en),ent],[tr('التصنيفات','Groups'),grp],[tr(cfg.grain==='month'?'شهري':'يومي',cfg.grain==='month'?'Monthly':'Daily'),day]],`${cfg.file}-analysis-${R.from}_${R.to}.xlsx`)}
function filesDlg(id){const cfg=REPORTS[id],s=ST[id];modal(tr(`ملفات ${cfg.label[0]} المستوردة`,`Imported ${cfg.label[1].toLowerCase()} files`),`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الملف','File'))}</th><th class="nosort">${esc(tr('الفترة','Period'))}</th><th class="nosort n">${esc(tr('الصفوف','Rows'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th><th class="nosort">${esc(tr('استورده','Imported by'))}</th>${s.canImport?'<th class="nosort"></th>':''}</tr></thead>
  <tbody>${s.batches.slice().reverse().map(b=>`<tr><td>${esc(b.file||'—')}${b.ok===false?` <span class="pill p-crit">${esc(tr('غير مكتمل','incomplete'))}</span>`:''}</td><td class="n">${fdi(b.from)} – ${fdi(b.to)}</td><td class="n">${fn(b.rows)}</td><td class="n">${qs(b.value||0)}</td><td><small>${esc(b.by||'')}<br>${fdi(b.at)}</small></td>${s.canImport?`<td><button type="button" class="btn sm ghost danger" data-cx-del="${esc(b.id)}" data-cx-id="${id}">${esc(tr('حذف','Delete'))}</button></td>`:''}</tr>`).join('')}</tbody></table></div>`,
  `<button type="button" class="btn" data-cx="reload" data-cx-id="${id}">${esc(tr('تحديث','Refresh'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- events ---------- */
document.addEventListener('click',async e=>{
  const pre=e.target.closest('[data-cx-pre]');if(pre){const id=pre.dataset.cxId,u=UI(id);u.pre=pre.dataset.cxPre;if(u.pre==='custom'){const R=range(id);u.from=u.from||R.from;u.to=u.to||R.to}u.lim=25;render();return}
  const so=e.target.closest('[data-cx-sort]');if(so&&IDS.includes(S.view)){const u=UI(S.view),k=so.dataset.cxSort;if(u.sort===k)u.dir=u.dir==='asc'?'desc':'asc';else{u.sort=k;u.dir=k==='name'?'asc':'desc'}render();return}
  const en=e.target.closest('[data-cx-ent]');if(en&&document.getElementById('view').contains(en)){entDetail(en.dataset.cxId,en.dataset.cxEnt);return}
  const del=e.target.closest('[data-cx-del]');if(del){const id=del.dataset.cxId;if(!del.classList.contains('armed')){del.classList.add('armed');del.textContent=tr('تأكيد الحذف','Confirm delete');setTimeout(()=>{del.classList.remove('armed');del.textContent=tr('حذف','Delete')},4000);return}
    const b=ST[id].batches.find(x=>x.id===del.dataset.cxDel);if(!b)return;del.disabled=true;try{await delBatch(id,b);toast(tr('تم حذف الملف','File deleted'));await load(id,true);filesDlg(id);render()}catch(err){toast(tr('تعذّر الحذف: ','Could not delete: ')+(err.message||''));del.disabled=false}return}
  const a=e.target.closest('[data-cx]');if(!a)return;const act=a.dataset.cx,id=a.dataset.cxId;if(!REPORTS[id])return;const s=ST[id];
  if(act==='import'){const f=document.getElementById('cx-file-'+id);if(f)f.click();return}
  if(act==='export'){exportXlsx(id);return}
  if(act==='print'){try{window.print()}catch(e2){}return}
  if(act==='files'){filesDlg(id);return}
  if(act==='more'){UI(id).lim+=50;render();return}
  if(act==='clear'){Object.assign(UI(id),{pre:'all',from:'',to:'',unit:'',prod:'',f:{}});render();return}
  if(act==='drop-local'){s.local=null;render();return}
  if(act==='reload'){a.disabled=true;await load(id,true);dlg.close();render();toast(tr('تم تحديث البيانات','Data refreshed'));return}
  if(act==='view-local'){s.local=s.pending;s.pending=null;dlg.close();Object.assign(UI(id),{pre:'all',unit:'',prod:'',f:{},q:''});render();return}
  if(act==='save'){const P=s.pending;if(!P)return;const msg=document.getElementById('cx-msg');a.disabled=true;msg.style.color='';msg.textContent=tr('جارٍ الحفظ…','Saving…');
    try{const rep=document.getElementById('cx-replace');await saveImport(id,P,!!(rep&&rep.checked));s.pending=null;s.local=null;await load(id,true);dlg.close();Object.assign(UI(id),{pre:'all',unit:'',prod:'',f:{},q:''});render();toast(tr(`تم حفظ ${fn(P.rows.length)} حركة`,`${fn(P.rows.length)} lines saved`))}
    catch(err){a.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية الحفظ في هذه اللوحة','You are not allowed to save to this dashboard'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}}
});
document.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;const en=e.target.closest&&e.target.closest('[data-cx-ent]');if(en&&document.getElementById('view').contains(en)){e.preventDefault();entDetail(en.dataset.cxId,en.dataset.cxEnt)}});

/* ---------- navigation: tasks of Commercial Affairs & Collections, next to SMS055 ---------- */
if(window.I18N&&I18N.add){const d={};IDS.forEach(id=>d[REPORTS[id].label[0]]=REPORTS[id].label[1]);I18N.add(d)}
IDS.forEach(id=>VIEWS.push({id,label:REPORTS[id].label[0],icon:'cx-'+id,group:'مهام إدارة الشؤون التجارية والتحصيلات',nav:false}));
ORG_TASKS.ca=(ORG_TASKS.ca||[]).concat(IDS);
const __render=render;
render=function(){if(REPORTS[S.view]){recompute();document.getElementById('view').innerHTML=dbBanner()+view(S.view);after(S.view);return}__render()};
/* the department's page in the org dashboard: one card with the three dashboards */
const __rp=recPanel;
recPanel=function(unit){const h=__rp(unit);if(unit!==UNIT)return h;
  const cards=IDS.map(id=>{const cfg=REPORTS[id],s=ST[id];let body;
    if(!s.loaded){load(id).then(()=>{if(S.view==='org')render()});body=`<span class="muted">${esc(tr('جارٍ التحميل…','Loading…'))}</span>`}
    else if(!s.rows.length)body=`<span class="muted">${esc(s.err==='denied'?tr('البيانات متاحة للإدارة والإدارة العامة.','Data is available to the department and the General Administration.'):tr('لم يُستورد أي ملف بعد.','No file imported yet.'))}</span>`;
    else{const bd=bounds(s.rows),from=bd.max.slice(0,8)+'01',unit=unitOf(id),rs=s.rows.filter(r=>r.d>=from&&(r.u||'')===unit),A=agg(cfg,rs);body=`<b class="num">${qs(A.q)}</b> ${esc(unit)} · ${fn(A.ents.length)} ${esc(tr(cfg.ent.pl[0],cfg.ent.pl[1]))} <small class="muted">(${esc(tr('آخر شهر','latest month'))})</small>`}
    return `<div><h4 style="margin:0 0 2px;font-size:14px">${esc(tr(cfg.label[0],cfg.label[1]))}</h4><span style="font-size:12.5px">${body}</span></div>`}).join('');
  return `<section class="panel"><h3 class="section-title">${esc(tr('لوحات مركز تقارير الإقفال','Closing-reports dashboards'))} <small>${esc(tr('تحليل المبيعات، حركة التحويلات، أذونات الاستلام','sales analysis, transfers, receipts'))}</small></h3><div class="sm-mini" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))">${cards}</div>
   <div class="cx-links">${IDS.map(id=>`<button type="button" class="btn sm primary" data-org-go="${id}">${esc(tr(REPORTS[id].label[0],REPORTS[id].label[1]))}</button>`).join('')}</div></section>`+h};
/* duties guide: shortcuts next to the department's duties */
const __vod=viewOrgDuties;
viewOrgDuties=function(el){__vod(el);const sec=el.querySelector('#od-'+UNIT+' .od-acts');if(!sec)return;
  IDS.slice().reverse().forEach(id=>{if(!sec.querySelector(`[data-org-go="${id}"]`))sec.insertAdjacentHTML('afterbegin',`<button type="button" class="btn sm" data-org-go="${id}">${esc(tr(REPORTS[id].label[0],REPORTS[id].label[1]))}</button>`)})};
window.__cx={ST,REPORTS,data,parseFile:(id,f)=>parseFile(REPORTS[id],f),load:(id,force)=>load(id,force),save:(id,P,replace)=>saveImport(id,P,replace)};
})();
