/* ===== لوحة مؤشرات وتحليل منظومة SMS055 (إدارة الشؤون التجارية والتحصيلات) =====
   Loaded after duties.js. The department imports the SMS055 export (Excel from the sales closing-reports centre:
   التاريخ، رقم العميل، اسم الزبون، رقم المنتج، اسم المنتج، PACK_CD، وحدة القياس، النوع، الكمية، القيمة، الكمية المسترجعة، القيمة المسترجعة)
   and the page turns it into KPIs, charts and a written reading.
   Data lives in the Firestore collection `sms055`: one `batch` document per imported file (period, totals, who/when) and
   `part` documents (<batchId>_p<n>) that hold the rows as compact JSON. Readable by the department, the General
   Administration, the board and the administrator; written by the administrator or an editor of ca / gm (firestore.rules). */
"use strict";
(function(){
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const UNIT='ca',COL='sms055',PART_ROWS=5000;
const SM={loaded:false,loading:null,err:'',batches:[],rows:[],local:null,canImport:false};

/* ---------- styles ---------- */
(function(){const st=document.createElement('style');st.textContent=`
.sm-root{--sm-c1:#2a78d6;--sm-c2:#eb6834;--sm-c3:#1baf7a;--sm-c4:#8b9099}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .sm-root{--sm-c1:#3987e5;--sm-c2:#d95926;--sm-c3:#199e70;--sm-c4:#80858d}}
:root[data-theme="dark"] .sm-root{--sm-c1:#3987e5;--sm-c2:#d95926;--sm-c3:#199e70;--sm-c4:#80858d}
.sm-hero{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:18px;align-items:center;padding:18px 22px;border-radius:16px;margin-bottom:14px;color:#fff;
  background:radial-gradient(520px 260px at 0% 0%,var(--hero-glow,rgba(255,255,255,.12)),transparent 60%),linear-gradient(135deg,var(--hero-a,#0b3168),var(--hero-b,#0063ad) 60%,var(--hero-c,#0a4f8a))}
.sm-hero h2{margin:4px 0 4px;font-size:clamp(19px,2.4vw,24px);color:#fff}
.sm-hero p{margin:0;color:rgba(255,255,255,.82);font-size:13.5px;line-height:1.75;max-width:70ch}
.sm-eyebrow{display:inline-block;font-size:12px;padding:3px 10px;border-radius:999px;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.2)}
.sm-hero-v{text-align:center;min-width:170px}.sm-hero-v b{display:block;font-size:clamp(24px,3vw,32px);font-family:var(--f-display);font-variant-numeric:tabular-nums;line-height:1.2}
.sm-hero-v span{font-size:12.5px;color:rgba(255,255,255,.8)}
.sm-hero-v .sm-delta{display:inline-block;margin-top:4px;padding:2px 8px;border-radius:999px;background:rgba(255,255,255,.16);font-size:12px;color:#fff}
.sm-bar{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin:0 0 14px}
.sm-bar .an-f select,.sm-bar .an-f input[type=search]{height:36px;border:1px solid var(--line);background:var(--surface);border-radius:8px;padding-inline:8px;color:var(--ink);font:inherit;min-width:140px}
.sm-presets{display:flex;gap:4px;flex-wrap:wrap}
.sm-presets button{height:36px;padding:0 12px;border-radius:8px;border:1px solid var(--line);background:var(--surface);color:var(--ink);font:inherit;font-size:12.5px;cursor:pointer}
.sm-presets button[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--on-accent,#fff)}
.sm-presets button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.sm-note{font-size:12.5px;color:var(--muted);margin:-6px 0 12px}
.sm-plot{position:relative;height:210px;margin:6px 0 0;padding-inline-start:52px}
.sm-grid{position:absolute;inset-inline:52px 0;border-top:1px dashed var(--line);height:0}
.sm-grid span{position:absolute;inset-inline-end:calc(100% + 6px);top:-8px;font-size:11px;color:var(--muted);white-space:nowrap;font-variant-numeric:tabular-nums}
.sm-cols{position:absolute;top:0;bottom:22px;inset-inline-start:52px;inset-inline-end:0;display:flex;align-items:stretch;gap:2px;border-bottom:1px solid var(--line)}
.sm-col{flex:1;min-width:0;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;position:relative;cursor:default;border-radius:4px}
.sm-col:hover{background:var(--accent-soft)}
.sm-col i{display:block;width:min(26px,78%);background:var(--ocf);border-radius:4px 4px 0 0;min-height:0}
.sm-col i.neg{background:var(--flame)}
.sm-col small{position:absolute;bottom:-20px;font-size:10.5px;color:var(--muted);white-space:nowrap}
.sm-col .sm-cl{position:absolute;font-size:10.5px;color:var(--ink-2,var(--muted));white-space:nowrap;font-variant-numeric:tabular-nums;transform:translateY(-16px)}
.sm-stack{display:flex;height:16px;gap:2px;border-radius:5px;overflow:hidden;margin:4px 0 10px}
.sm-stack span{display:block;height:100%}
.sm-legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12.5px}
.sm-legend span{display:inline-flex;align-items:center;gap:6px}.sm-legend i{width:10px;height:10px;border-radius:3px;display:inline-block}
.sm-k1{background:var(--sm-c1)}.sm-k2{background:var(--sm-c2)}.sm-k3{background:var(--sm-c3)}.sm-k4{background:var(--sm-c4)}
.sm-ins{margin:0;padding:0;list-style:none;display:grid;gap:8px}
.sm-ins li{display:flex;gap:10px;align-items:flex-start;font-size:13.5px;line-height:1.75;padding:8px 10px;border-radius:10px;background:var(--surface-2,var(--accent-soft))}
.sm-ins li b{font-variant-numeric:tabular-nums}
.sm-ins .ic{flex:none;width:22px;height:22px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:700;color:#fff;background:var(--accent)}
.sm-ins .ic.w{background:var(--warn)}.sm-ins .ic.c{background:var(--crit)}.sm-ins .ic.o{background:var(--ok)}
.sm-prod{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,220px),1fr));gap:10px}
.sm-pc{border:1px solid var(--line);border-radius:12px;padding:12px 14px;background:var(--surface);display:grid;gap:4px}
.sm-pc h4{margin:0;font-size:14px;display:flex;gap:8px;align-items:center}.sm-pc h4 i{width:10px;height:10px;border-radius:3px;flex:none}
.sm-pc .q{font-family:var(--f-display);font-size:22px;font-weight:600;font-variant-numeric:tabular-nums}
.sm-pc .s{font-size:12.5px;color:var(--muted);line-height:1.7}
.sm-tbl td.n,.sm-tbl th.n{text-align:end;font-variant-numeric:tabular-nums}
.sm-tbl tr[data-sm-cust]{cursor:pointer}
.sm-tbl th[data-sm-sort]{cursor:pointer;white-space:nowrap}.sm-tbl th[data-sm-sort][aria-sort]::after{content:" ▾";font-size:10px}.sm-tbl th[data-sm-sort][aria-sort="ascending"]::after{content:" ▴"}
.sm-share{display:inline-block;height:6px;border-radius:3px;background:var(--ocf);vertical-align:middle;margin-inline-start:6px}
.sm-empty{padding:34px 20px;text-align:center;display:grid;gap:10px;justify-items:center}
.sm-empty h3{margin:0}.sm-empty p{margin:0;color:var(--muted);max-width:62ch;line-height:1.8}
.sm-cols-map{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:6px;font-size:12.5px;margin:8px 0}
.sm-cols-map span{padding:5px 8px;border-radius:8px;border:1px solid var(--line)}.sm-cols-map span.ok{border-color:var(--ok)}.sm-cols-map span.no{border-color:var(--crit);color:var(--crit)}
.sm-mini{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin:6px 0 10px}
.sm-mini div{border:1px solid var(--line);border-radius:10px;padding:8px 10px}.sm-mini b{display:block;font-size:18px;font-variant-numeric:tabular-nums}.sm-mini span{font-size:12px;color:var(--muted)}
@media (max-width:720px){.sm-hero{grid-template-columns:1fr}.sm-hero-v{text-align:start}.sm-plot{padding-inline-start:44px}.sm-cols,.sm-grid{inset-inline-start:44px}}
@media print{.sm-bar,.sm-acts{display:none}}
`;document.head.appendChild(st)})();

/* ---------- helpers ---------- */
const money=v=>anShort(v);
const money2=v=>fn(Math.round(v*100)/100);
const pct=(a,b)=>b?Math.round(a*1000/b)/10:0;
const DAY=86400000;
const isoAdd=(s,n)=>new Date(Date.parse(s+'T00:00:00Z')+n*DAY).toISOString().slice(0,10);
const days=(a,b)=>Math.round((Date.parse(b+'T00:00:00Z')-Date.parse(a+'T00:00:00Z'))/DAY)+1;
const MN_AR=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'],MN_EN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const mName=m=>tr(MN_AR[m],MN_EN[m]);
const fdi=s=>s?s.slice(0,10).replace(/-/g,'/'):'—';
const AD=s=>String(s==null?'':s).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
function num(v){if(typeof v==='number')return isFinite(v)?v:0;const s=AD(v).replace(/[,\s٬]/g,'').replace('٫','.');if(!s)return 0;const neg=/^\(.*\)$/.test(s);const n=parseFloat(s.replace(/[()]/g,''));return isNaN(n)?0:(neg?-n:n)}
function me(){return (window.__me&&window.__me())||{}}

/* ---------- column recognition ---------- */
const FIELDS={d:['التاريخ','تاريخ','تاريخ الفاتوره','date','trx_date','inv_date'],c:['رقم العميل','رقم الزبون','كود العميل','كود الزبون','customer_no','cust_no','customer_id','cust_cd'],
  cn:['اسم الزبون','اسم العميل','الزبون','العميل','customer','customer_name','cust_name'],p:['رقم المنتج','كود المنتج','رمز المنتج','product_no','item_no','prod_cd','item_cd'],
  pn:['اسم المنتج','المنتج','الصنف','product','product_name','item','item_name'],k:['pack_cd','pack','العبوه','كود العبوه'],u:['وحده القياس','الوحده','uom','unit'],
  t:['النوع','نوع','type','trx_type'],q:['الكميه','qty','quantity'],v:['القيمه','value','amount'],rq:['الكميه المسترجعه','الكميه المرتجعه','returned qty','ret_qty'],rv:['القيمه المسترجعه','القيمه المرتجعه','returned value','ret_value']};
const FIELD_AR={d:'التاريخ',c:'رقم العميل',cn:'اسم الزبون',p:'رقم المنتج',pn:'اسم المنتج',k:'PACK_CD',u:'وحدة القياس',t:'النوع',q:'الكمية',v:'القيمة',rq:'الكمية المسترجعة',rv:'القيمة المسترجعة'};
const FIELD_EN={d:'Date',c:'Customer no.',cn:'Customer name',p:'Product no.',pn:'Product name',k:'PACK_CD',u:'Unit',t:'Type',q:'Quantity',v:'Value',rq:'Returned qty',rv:'Returned value'};
const hn=s=>norm(String(s||'')).replace(/[_\-]+/g,'_').replace(/\s+/g,' ').trim();
function mapHeader(row){const m={};const H=row.map(hn);
  for(const f in FIELDS){const al=FIELDS[f].map(hn);const i=H.findIndex(h=>h&&al.includes(h));if(i>=0)m[f]=i}return m}
function findHeader(grid){let best=null;for(let i=0;i<Math.min(grid.length,25);i++){const m=mapHeader(grid[i]||[]);const n=Object.keys(m).length;if(!best||n>best.n)best={i,m,n}}return best&&best.n>=3?best:null}

/* ---------- reading files: every format through the shared reader (fileimport.js) ---------- */
const readGrids=file=>FileImport.readGrids(file);

/* dates: Excel serial, ISO, or m/d/yyyy (the SMS055 export shows 8/31/2026 0:00); d/m/yyyy when a first part is above 12 */
function serialIso(n){return new Date(Date.UTC(1899,11,30)+Math.round(n*86400000)).toISOString().slice(0,10)}
function parseRows(grid){const h=findHeader(grid);if(!h)return null;const m=h.m;const body=grid.slice(h.i+1);
  const g=(r,f)=>m[f]==null?'':r[m[f]];let dm=false;
  body.forEach(r=>{const v=g(r,'d');if(typeof v==='string'){const x=AD(v).match(/^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);if(x&&+x[1]>12)dm=true}});
  const pd=v=>{if(v&&typeof v==='object'&&'serial' in v)return serialIso(v.serial);if(typeof v==='number')return v>20000&&v<80000?serialIso(v):'';
    const s=AD(v).trim();let x=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(x)return x[1]+'-'+x[2].padStart(2,'0')+'-'+x[3].padStart(2,'0');
    x=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);if(x){let [a,b,y]=[+x[1],+x[2],+x[3]];if(y<100)y+=2000;const mo=dm?b:a,da=dm?a:b;if(mo<1||mo>12||da<1||da>31)return '';return y+'-'+String(mo).padStart(2,'0')+'-'+String(da).padStart(2,'0')}
    const n=parseFloat(s);if(!isNaN(n)&&/^\d+(\.\d+)?$/.test(s)&&n>20000&&n<80000)return serialIso(n);return ''};
  const rows=[];let skipped=0;
  body.forEach(r=>{if(!r||!r.some(x=>String(x&&typeof x==='object'?'1':x).trim()))return;const d=pd(g(r,'d'));
    const c=String(g(r,'c')).trim(),cn=String(g(r,'cn')).trim(),q=num(g(r,'q')),v=num(g(r,'v')),rq=num(g(r,'rq')),rv=num(g(r,'rv'));
    if(!d||(!c&&!cn)||(!q&&!v&&!rq&&!rv)){skipped++;return}
    rows.push({d,c:c||cn,cn:cn||c,p:String(g(r,'p')).trim()||String(g(r,'pn')).trim()||'—',pn:String(g(r,'pn')).trim()||String(g(r,'p')).trim()||'—',k:String(g(r,'k')).trim(),u:String(g(r,'u')).trim(),t:String(g(r,'t')).trim(),q,v,rq,rv})});
  return {rows,skipped,map:m,header:h.i+1}}
async function parseFile(file){const sheets=await readGrids(file);let best=null;
  for(const s of sheets){const r=parseRows(s.grid);if(r&&(!best||r.rows.length>best.rows.length))best=Object.assign(r,{sheet:s.name})}
  if(!best)throw new Error(tr('لم نجد أعمدة تقرير SMS055 في الملف (التاريخ، رقم العميل، اسم الزبون، الكمية، القيمة…).','The SMS055 columns (date, customer, quantity, value…) were not found in the file.'));
  if(!best.rows.length)throw new Error(tr('الأعمدة موجودة لكن لا توجد صفوف صالحة (تاريخ + زبون + كمية أو قيمة).','Columns found, but no valid rows (date + customer + quantity or value).'));
  const ds=best.rows.map(r=>r.d).sort();best.from=ds[0];best.to=ds[ds.length-1];best.file=file.name;return best}

/* ---------- compact storage ---------- */
function encode(rows){const C={},P={};const R=rows.map(r=>{if(!(r.c in C))C[r.c]=r.cn;if(!(r.p in P))P[r.p]=[r.pn,r.u];return [r.d,r.c,r.p,r.k,r.t,r.q,r.v,r.rq,r.rv]});return JSON.stringify({C,P,R})}
function decode(s){const o=JSON.parse(s);return (o.R||[]).map(a=>{const p=o.P[a[2]]||['',''];return {d:a[0],c:String(a[1]),cn:o.C[a[1]]||String(a[1]),p:String(a[2]),pn:p[0]||String(a[2]),u:p[1]||'',k:a[3]||'',t:a[4]||'',q:+a[5]||0,v:+a[6]||0,rq:+a[7]||0,rv:+a[8]||0}})}
async function load(force){if(SM.loading&&!force)return SM.loading;
  SM.loading=(async()=>{const FS=window.__fs&&window.__fs();if(!FS){SM.loaded=true;return}
    try{const qs=await FS.collection(COL).get();const docs=[];qs.forEach(d=>docs.push(Object.assign({id:d.id},d.data())));
      const parts={};docs.filter(d=>d.kind==='part').forEach(p=>{(parts[p.batch]=parts[p.batch]||[]).push(p)});
      const batches=docs.filter(d=>d.kind==='batch').sort((a,b)=>(a.from||'').localeCompare(b.from||''));const rows=[];
      batches.forEach(b=>{const ps=(parts[b.id]||[]).sort((x,y)=>x.n-y.n);b.ok=ps.length===b.parts;if(!b.ok)return;ps.forEach(p=>{try{decode(p.d).forEach(r=>{r.b=b.id;rows.push(r)})}catch(e){b.ok=false}})});
      SM.batches=batches;SM.rows=rows;SM.err=''}
    catch(e){SM.err=e&&e.code==='permission-denied'?'denied':(e&&e.message)||'error';SM.batches=[];SM.rows=[]}
    SM.loaded=true;
    try{const u=await recMyUnit();const m=me();SM.canImport=m.role==='admin'||(!!S.canWrite&&m.role!=='viewer'&&(u===UNIT||u==='gm'))}catch(e){SM.canImport=false}
  })();return SM.loading}
async function saveImport(P,replace){const FS=window.__fs(),m=me();const id='b'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
  const chunks=[];for(let i=0;i<P.rows.length;i+=PART_ROWS)chunks.push(P.rows.slice(i,i+PART_ROWS));
  for(let i=0;i<chunks.length;i+=4){const b=FS.batch();chunks.slice(i,i+4).forEach((ch,j)=>b.set(FS.collection(COL).doc(id+'_p'+(i+j)),{kind:'part',batch:id,n:i+j,d:encode(ch),createdBy:m.email}));await b.commit()}
  const tot=P.rows.reduce((a,r)=>{a.v+=r.v;a.rv+=r.rv;return a},{v:0,rv:0});
  await FS.collection(COL).doc(id).set({kind:'batch',file:String(P.file||'').slice(0,200),sheet:String(P.sheet||'').slice(0,100),from:P.from,to:P.to,rows:P.rows.length,parts:chunks.length,
    value:Math.round(tot.v*100)/100,returns:Math.round(tot.rv*100)/100,createdBy:m.email,by:String(m.name||m.email||'').slice(0,120),at:new Date().toISOString()});
  let removed=0;if(replace){for(const o of overlaps(P)){await delBatch(o,true);removed++}}
  try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'استيراد SMS055',col:COL,id,summary:`${P.file} · ${P.from} – ${P.to} · ${P.rows.length} صف`+(removed?` · استُبدل ${removed} ملف`:'')})}catch(e){}
  return id}
function overlaps(P){return SM.batches.filter(b=>!(b.to<P.from||b.from>P.to))}
async function delBatch(b,quiet){const FS=window.__fs();const n=b.parts||0;const ops=[];for(let i=0;i<n;i++)ops.push(FS.collection(COL).doc(b.id+'_p'+i));
  for(let i=0;i<ops.length;i+=400){const bt=FS.batch();ops.slice(i,i+400).forEach(r=>bt.delete(r));await bt.commit()}
  await FS.collection(COL).doc(b.id).delete();
  if(!quiet){const m=me();try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'حذف ملف SMS055',col:COL,id:b.id,summary:`${b.file||''} · ${b.from} – ${b.to}`})}catch(e){}}}

/* ---------- analysis ---------- */
function UI(){return S.ui.sm||(S.ui.sm={pre:'all',from:'',to:'',prod:'',type:'',q:'',sort:'net',dir:'desc',lim:25})}
function raw(){return SM.local?SM.local.rows:SM.rows}
/* quantities shown in cubic metres (units.js); the value columns (LYD) are untouched */
function data(){const r=raw();if(window.FuelUnits){FuelUnits.seen(r);return FuelUnits.normalize(r)}return r}
const qty=v=>fn(Math.round(v*100)/100);
function bounds(rows){if(!rows.length)return null;let a=rows[0].d,b=a;rows.forEach(r=>{if(r.d<a)a=r.d;if(r.d>b)b=r.d});return {min:a,max:b}}
function range(){const u=UI(),bd=bounds(data());if(!bd)return null;let from=bd.min,to=bd.max;
  if(u.pre==='day')from=to;else if(u.pre==='7')from=isoAdd(to,-6);else if(u.pre==='month')from=to.slice(0,8)+'01';else if(u.pre==='custom'){from=u.from||bd.min;to=u.to||bd.max}
  if(from<bd.min&&u.pre!=='custom')from=bd.min;return {from,to,min:bd.min,max:bd.max}}
function pick(rows,from,to){const u=UI();return rows.filter(r=>r.d>=from&&r.d<=to&&(!u.prod||r.p===u.prod)&&(!u.type||r.t===u.type))}
function agg(rows){const A={n:rows.length,v:0,rv:0,q:0,rq:0,cust:new Map(),prod:new Map(),type:new Map(),day:new Map()};
  rows.forEach(r=>{A.v+=r.v;A.rv+=r.rv;
    let c=A.cust.get(r.c);if(!c)A.cust.set(r.c,c={c:r.c,cn:r.cn,n:0,v:0,rv:0,last:'',days:new Set(),prods:new Set()});c.n++;c.v+=r.v;c.rv+=r.rv;if(r.d>c.last)c.last=r.d;c.days.add(r.d);c.prods.add(r.p);
    let p=A.prod.get(r.p);if(!p)A.prod.set(r.p,p={p:r.p,pn:r.pn,u:r.u,n:0,q:0,v:0,rq:0,rv:0,cust:new Set()});p.n++;p.q+=r.q;p.v+=r.v;p.rq+=r.rq;p.rv+=r.rv;p.cust.add(r.c);
    const tk=r.t||'—';let t=A.type.get(tk);if(!t)A.type.set(tk,t={t:tk,n:0,v:0});t.n++;t.v+=r.v-r.rv;
    let d=A.day.get(r.d);if(!d)A.day.set(r.d,d={d:r.d,v:0,rv:0,n:0,cust:new Set()});d.v+=r.v;d.rv+=r.rv;d.n++;d.cust.add(r.c)});
  A.net=A.v-A.rv;A.custs=[...A.cust.values()].map(c=>Object.assign(c,{net:c.v-c.rv}));A.prods=[...A.prod.values()].map(p=>Object.assign(p,{net:p.v-p.rv})).sort((a,b)=>b.net-a.net);
  A.types=[...A.type.values()].sort((a,b)=>b.v-a.v);return A}
/* product colors follow the product (overall rank in all data), not the filtered rank */
function prodColors(){const m=new Map();const tot=new Map();data().forEach(r=>tot.set(r.p,(tot.get(r.p)||0)+r.v-r.rv));[...tot.entries()].sort((a,b)=>b[1]-a[1]).forEach(([p],i)=>m.set(p,i<3?'sm-k'+(i+1):'sm-k4'));return m}

/* ---------- charts ---------- */
function niceMax(v){if(v<=0)return 1;const e=Math.pow(10,Math.floor(Math.log10(v)));const f=v/e;return (f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*e}
function colChart(pts,label){if(!pts.length)return `<div class="empty" style="padding:20px">${tr('لا توجد بيانات ضمن التصفية الحالية','No data for the current filter')}</div>`;
  const mx=niceMax(Math.max(...pts.map(p=>p.y),0)),step=Math.ceil(pts.length/12);let iMax=0;pts.forEach((p,i)=>{if(p.y>pts[iMax].y)iMax=i});
  const grid=[1,.5,0].map(f=>`<div class="sm-grid" style="top:${(1-f)*188}px"><span>${money(mx*f)}</span></div>`).join('');
  return `<div class="sm-plot" role="img" aria-label="${esc(label)}">${grid}<div class="sm-cols">${pts.map((p,i)=>{const h=Math.max(0,p.y)*100/mx;
    return `<div class="sm-col" data-tip="${esc(p.tip)}" tabindex="0">${i===iMax||i===pts.length-1&&pts.length>1?`<span class="sm-cl" style="bottom:${h}%">${money(p.y)}</span>`:''}<i class="${p.y<0?'neg':''}" style="height:${p.y>0?Math.max(1,h):0}%"></i>${i%step===0?`<small>${esc(p.l)}</small>`:''}</div>`}).join('')}</div></div>`}
function bars(rows,opt={}){if(!rows.length)return `<div class="empty" style="padding:20px">${tr('لا توجد بيانات ضمن التصفية الحالية','No data for the current filter')}</div>`;
  const max=Math.max(1,...rows.map(r=>r.v));
  return `<div class="an-bars">${rows.map(r=>`<div class="an-row${r.go?' go':''}" ${r.go?`data-sm-cust="${esc(r.go)}" role="button" tabindex="0"`:''} data-tip="${esc(r.tip||(r.l+': '+money2(r.v)))}">
   <span class="an-lbl" title="${esc(r.l)}">${esc(r.l)}</span><span class="an-track"><span class="an-fill${opt.tone?' '+opt.tone:''}" style="width:${r.v>0?Math.max(1.5,r.v*100/max):0}%"></span></span><span class="an-val num">${opt.fmt?opt.fmt(r.v):money(r.v)}</span></div>`).join('')}</div>`}
function stack(parts){const tot=parts.reduce((s,p)=>s+Math.max(0,p.v),0)||1;
  return `<div class="sm-stack">${parts.filter(p=>p.v>0).map(p=>`<span class="${p.k}" style="width:${p.v*100/tot}%" data-tip="${esc(p.l+': '+money2(p.v)+' '+tr('د.ل','LYD')+' ('+pct(p.v,tot)+'%)')}"></span>`).join('')}</div>
   <div class="sm-legend">${parts.map(p=>`<span><i class="${p.k}"></i>${esc(p.l)} <b class="num">${money(p.v)}</b> <small class="muted">${pct(Math.max(0,p.v),tot)}%</small></span>`).join('')}</div>`}

/* ---------- page ---------- */
I.sms='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 20h16"/><path d="M6 16l4-5 3 3 5-7"/><circle cx="18" cy="7" r="1.6"/></svg>';
I.up='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 16V4M7 9l5-5 5 5M5 20h14"/></svg>';
function heroHtml(sub,v,delta){return `<div class="sm-hero"><div><span class="sm-eyebrow">${esc(tr('إدارة الشؤون التجارية والتحصيلات','Commercial Affairs & Collections'))}</span>
  <h2>${esc(tr('لوحة مؤشرات منظومة SMS055','SMS055 sales dashboard'))}</h2><p>${sub}</p></div>${v==null?'':`<div class="sm-hero-v"><b>${v}</b><span>${esc(tr('صافي المبيعات · د.ل','Net sales · LYD'))}</span>${delta||''}</div>`}</div>`}
function toolbar(){return `<div class="toolbar sm-acts">
  ${SM.canImport||!SM.rows.length||SM.err?`<button class="btn primary" data-sm="import">${I.up}${esc(tr('استيراد ملف SMS055','Import SMS055 file'))}</button>`:''}
  ${data().length?`<button class="btn" data-sm="export">${I.down}${esc(tr('تصدير التحليل Excel','Export analysis (Excel)'))}</button><button class="btn" data-sm="print">${esc(tr('طباعة / PDF','Print / PDF'))}</button>`:''}
  ${SM.batches.length?`<button class="btn ghost" data-sm="files">${esc(tr('الملفات المستوردة','Imported files'))} <span class="num">${SM.batches.length}</span></button>`:''}
  ${SM.local?`<button class="btn ghost" data-sm="drop-local">${esc(tr('إغلاق العرض المؤقت','Close temporary view'))}</button>`:''}
  <input type="file" id="sm-file" accept="${FileImport.ACCEPT}" hidden></div>`}
function viewSms(){
  if(!SM.loaded){load().then(()=>{if(S.view==='sms055')render()});return heroHtml(esc(tr('جارٍ تحميل بيانات SMS055…','Loading SMS055 data…')),null)+'<div class="panel muted">'+esc(tr('جارٍ التحميل…','Loading…'))+'</div>'}
  const rows=data();
  const head=`<div class="page-head" style="margin-bottom:8px"><div><h2 style="margin:0">${esc(tr('مخرجات منظومة SMS055','SMS055 outputs'))}</h2></div>${toolbar()}</div>`;
  if(!rows.length){const denied=SM.err==='denied';
    return head+heroHtml(esc(tr('حوّل تقرير SMS055 (مركز تقارير الإقفال) إلى مؤشرات: صافي المبيعات، المرتجعات، الزبائن النشطون، المنتجات، الاتجاه اليومي، والزبائن الذين توقفوا عن الشراء.','Turn the SMS055 report (closing-reports centre) into KPIs: net sales, returns, active customers, products, the daily trend and customers who stopped buying.')),null)+
    `<div class="panel sm-empty">${I.sms.replace('<svg','<svg style="width:46px;height:46px;color:var(--accent)"')}<h3>${esc(denied?tr('هذه اللوحة لإدارة الشؤون التجارية والتحصيلات','This dashboard belongs to Commercial Affairs & Collections'):tr('لم يُستورد أي ملف بعد','No file imported yet'))}</h3>
     <p>${esc(denied?tr('لا تملك صلاحية الاطلاع على البيانات المحفوظة. يمكنك استيراد ملف وتحليله مؤقتًا على جهازك دون حفظه.','You cannot view the saved data. You can still import a file and analyse it on this device without saving it.'):
       tr('من «مركز تقارير الإقفال» في منظومة مبيعات البريقة حدّد تاريخ البداية والنهاية، ونزّل «نموذج SMS055» بصيغة Excel، ثم اضغط «استيراد ملف SMS055». يتعرّف الاستيراد على الأعمدة تلقائيًا.','In the Brega sales closing-reports centre pick the start and end dates, download “SMS055 form” as Excel, then press “Import SMS055 file”. Columns are recognised automatically.'))}</p>
     <button class="btn primary" data-sm="import">${I.up}${esc(tr('استيراد ملف SMS055','Import SMS055 file'))}</button></div>`}
  const u=UI(),R=range(),cur=pick(rows,R.from,R.to),A=agg(cur);
  const len=days(R.from,R.to),pFrom=isoAdd(R.from,-len),pTo=isoAdd(R.from,-1),prevRows=pick(rows,pFrom,pTo),B=prevRows.length?agg(prevRows):null;
  const dNet=B&&B.net?((A.net-B.net)*100/Math.abs(B.net)):null;
  const deltaHtml=dNet==null?'':`<span class="sm-delta">${dNet>=0?'▲':'▼'} ${Math.abs(Math.round(dNet*10)/10)}% ${esc(tr('عن الفترة السابقة','vs previous period'))}</span>`;
  const PC=prodColors();
  const prods=[...new Map(rows.map(r=>[r.p,r.pn])).entries()].sort((a,b)=>a[1].localeCompare(b[1],'ar'));
  const types=uniq(rows.map(r=>r.t));const unitNote=window.FuelUnits?FuelUnits.explain(cur):'';
  const custs=A.custs.slice().sort((a,b)=>b.net-a.net);const top10=custs.slice(0,10),top10v=top10.reduce((s,c)=>s+c.net,0);
  const activeDays=A.day.size,avgDay=activeDays?A.net/activeDays:0,retRate=pct(A.rv,A.v);
  // trend: daily up to 62 days, monthly beyond
  let pts;if(len<=62){pts=[];for(let d=R.from;d<=R.to;d=isoAdd(d,1)){const x=A.day.get(d);const net=x?x.v-x.rv:0;
      pts.push({l:String(+d.slice(8)),y:net,tip:`${fdi(d)} · ${tr('صافي','Net')} ${money2(net)} · ${tr('مبيعات','Sales')} ${money2(x?x.v:0)} · ${tr('مرتجع','Returns')} ${money2(x?x.rv:0)} · ${x?x.cust.size:0} ${tr('زبون','customers')}`})}}
  else{const mm=new Map();cur.forEach(r=>{const k=r.d.slice(0,7);const o=mm.get(k)||{v:0,rv:0,c:new Set()};o.v+=r.v;o.rv+=r.rv;o.c.add(r.c);mm.set(k,o)});
    pts=[...mm.keys()].sort().map(k=>{const o=mm.get(k),net=o.v-o.rv;return {l:mName(+k.slice(5)-1).slice(0,3)+' '+k.slice(2,4),y:net,tip:`${mName(+k.slice(5)-1)} ${k.slice(0,4)} · ${tr('صافي','Net')} ${money2(net)} · ${o.c.size} ${tr('زبون','customers')}`}})}
  // customers who stopped / new ones vs the previous period
  const lost=B?B.custs.filter(c=>c.net>0&&!A.cust.has(c.c)).sort((a,b)=>b.net-a.net):[];const fresh=B?A.custs.filter(c=>!B.cust.has(c.c)):[];
  const lostV=lost.reduce((s,c)=>s+c.net,0);
  const retC=A.custs.filter(c=>c.rv>0).sort((a,b)=>b.rv-a.rv);
  // written reading
  const ins=[];const tp=A.prods[0];
  if(dNet!=null)ins.push([dNet>=0?'o':'w',dNet>=0?'▲':'▼',tr(`صافي المبيعات <b>${money(A.net)}</b> د.ل، ${dNet>=0?'بزيادة':'بانخفاض'} <b>${Math.abs(Math.round(dNet*10)/10)}%</b> عن الفترة السابقة المماثلة (${fdi(pFrom)} – ${fdi(pTo)}: ${money(B.net)} د.ل).`,`Net sales <b>${money(A.net)}</b> LYD, ${dNet>=0?'up':'down'} <b>${Math.abs(Math.round(dNet*10)/10)}%</b> on the same-length previous period (${fdi(pFrom)} – ${fdi(pTo)}: ${money(B.net)} LYD).`)]);
  else ins.push(['','i',tr(`صافي المبيعات <b>${money(A.net)}</b> د.ل خلال <b>${activeDays}</b> يوم بيع، بمتوسط <b>${money(avgDay)}</b> د.ل يوميًا. لا تتوفر بيانات للفترة السابقة للمقارنة؛ استورد ملف الفترة السابقة لتظهر المقارنة.`,`Net sales <b>${money(A.net)}</b> LYD over <b>${activeDays}</b> sales days, <b>${money(avgDay)}</b> LYD a day on average. No previous-period data to compare; import it to see the comparison.`)]);
  if(tp&&A.net)ins.push(['','1',tr(`المنتج الأعلى: <b>${esc(tp.pn)}</b> بنسبة <b>${pct(tp.net,A.net)}%</b> من صافي المبيعات (${qty(tp.q)} ${esc(tp.u||'')}).`,`Top product: <b>${esc(tp.pn)}</b>, <b>${pct(tp.net,A.net)}%</b> of net sales (${qty(tp.q)} ${esc(tp.u||'')}).`)]);
  if(custs.length){const sh=pct(top10v,A.net);ins.push([sh>=60?'w':'','10',tr(`أعلى 10 زبائن يمثلون <b>${sh}%</b> من صافي المبيعات${sh>=60?'، تركّز مرتفع يستحق متابعة المخاطر الائتمانية':''}. أكبر زبون: <b>${esc(custs[0].cn)}</b> (${money(custs[0].net)} د.ل).`,`The top 10 customers are <b>${sh}%</b> of net sales${sh>=60?', a high concentration worth credit-risk follow-up':''}. Largest: <b>${esc(custs[0].cn)}</b> (${money(custs[0].net)} LYD).`)])}
  if(A.rv>0)ins.push([retRate>=2?'c':'w','↩',tr(`المرتجعات <b>${money(A.rv)}</b> د.ل (<b>${retRate}%</b> من المبيعات) لدى <b>${retC.length}</b> زبون، أعلاها لدى <b>${esc(retC[0].cn)}</b>.`,`Returns <b>${money(A.rv)}</b> LYD (<b>${retRate}%</b> of sales) from <b>${retC.length}</b> customers, most from <b>${esc(retC[0].cn)}</b>.`)]);
  else ins.push(['o','✓',tr('لا توجد مرتجعات في الفترة المحددة.','No returns in the selected period.')]);
  if(lost.length)ins.push(['w','!',tr(`<b>${lost.length}</b> زبون اشتروا في الفترة السابقة ولم يشتروا في هذه الفترة (صافي مشترياتهم السابقة ${money(lostV)} د.ل). راجع القائمة أدناه للمتابعة.`,`<b>${lost.length}</b> customers bought in the previous period but not in this one (previous net ${money(lostV)} LYD). See the list below.`)]);
  if(fresh.length)ins.push(['o','+',tr(`<b>${fresh.length}</b> زبون جديد مقارنة بالفترة السابقة.`,`<b>${fresh.length}</b> new customers compared with the previous period.`)]);
  const neg=custs.filter(c=>c.net<0);if(neg.length)ins.push(['c','−',tr(`<b>${neg.length}</b> زبون صافي حركته سالب (المرتجع أكبر من المبيعات).`,`<b>${neg.length}</b> customers have a negative net (returns above sales).`)]);
  // customer table
  const q=norm(u.q||'');let tbl=A.custs.filter(c=>!q||norm(c.cn+' '+c.c).includes(q));const key={net:c=>c.net,v:c=>c.v,rv:c=>c.rv,n:c=>c.n,cn:c=>c.cn,last:c=>c.last,days:c=>c.days.size}[u.sort]||(c=>c.net);
  tbl.sort((a,b)=>{const x=key(a),y=key(b);const r=typeof x==='string'?x.localeCompare(y,'ar'):x-y;return u.dir==='asc'?r:-r});const tblN=tbl.length;tbl=tbl.slice(0,u.lim);
  const th=(k,l,cls='')=>`<th class="${cls}" data-sm-sort="${k}" ${u.sort===k?`aria-sort="${u.dir==='asc'?'ascending':'descending'}"`:''}>${esc(l)}</th>`;
  const mxShare=Math.max(1,...custs.slice(0,1).map(c=>c.net));
  const sub=esc(tr(`الفترة ${fdi(R.from)} – ${fdi(R.to)} · ${fn(A.n)} حركة · ${A.custs.length} زبون · البيانات المتاحة ${fdi(R.min)} – ${fdi(R.max)}`,`Period ${fdi(R.from)} – ${fdi(R.to)} · ${fn(A.n)} lines · ${A.custs.length} customers · data available ${fdi(R.min)} – ${fdi(R.max)}`));
  return head+(SM.local?`<div class="banner info">${esc(tr(`عرض مؤقت للملف «${SM.local.file}» على هذا الجهاز فقط، ولم يُحفظ في المنظومة.`,`Temporary view of “${SM.local.file}” on this device only; not saved.`))}</div>`:'')+
  heroHtml(sub,money(A.net),deltaHtml)+
  `<div class="sm-bar">
    <div class="an-f">${esc(tr('الفترة','Period'))}<div class="sm-presets">${[['all',tr('كل البيانات','All data')],['month',tr('آخر شهر','Latest month')],['7',tr('آخر 7 أيام','Last 7 days')],['day',tr('آخر يوم','Latest day')],['custom',tr('مخصّص','Custom')]].map(([k,l])=>`<button type="button" data-sm-pre="${k}" aria-pressed="${u.pre===k}">${esc(l)}</button>`).join('')}</div></div>
    ${u.pre==='custom'?`<label class="an-f">${esc(tr('من','From'))}<input type="date" id="sm-from" value="${esc(R.from)}" min="${R.min}" max="${R.max}"></label><label class="an-f">${esc(tr('إلى','To'))}<input type="date" id="sm-to" value="${esc(R.to)}" min="${R.min}" max="${R.max}"></label>`:''}
    <label class="an-f">${esc(tr('المنتج','Product'))}<select id="sm-prod"><option value="">${esc(tr('كل المنتجات','All products'))}</option>${prods.map(([p,n])=>`<option value="${esc(p)}" ${u.prod===p?'selected':''}>${esc(n)} (${esc(p)})</option>`).join('')}</select></label>
    ${types.length>1?`<label class="an-f">${esc(tr('النوع','Type'))}<select id="sm-type"><option value="">${esc(tr('كل الأنواع','All types'))}</option>${types.map(t=>`<option ${u.type===t?'selected':''}>${esc(t)}</option>`).join('')}</select></label>`:''}
    ${u.prod||u.type||u.pre!=='all'?`<button class="btn sm ghost" data-sm="clear">${esc(tr('مسح التصفية','Clear filters'))}</button>`:''}
  </div>
  ${unitNote?`<p class="sm-note">${esc(unitNote)}</p>`:''}
  ${!A.n?`<div class="panel empty"><b>${esc(tr('لا توجد حركات','No lines'))}</b>${esc(tr('ضمن التصفية الحالية.','for the current filter.'))}</div>`:`
  <div class="kpis an-kpis">
   ${anKpi(tr('إجمالي المبيعات','Gross sales'),money(A.v),esc(tr(`د.ل · ${fn(A.n)} حركة`,`LYD · ${fn(A.n)} lines`)))}
   ${anKpi(tr('المرتجعات','Returns'),money(A.rv),esc(tr(`${retRate}% من المبيعات`,`${retRate}% of sales`)),retRate>=2?'crit':A.rv?'warn':'ok')}
   ${anKpi(tr('الزبائن النشطون','Active customers'),fn(A.custs.length),esc(B?tr(`${fresh.length} جديد · ${lost.length} توقّف`,`${fresh.length} new · ${lost.length} stopped`):tr(`متوسط ${money(A.custs.length?A.net/A.custs.length:0)} د.ل للزبون`,`avg ${money(A.custs.length?A.net/A.custs.length:0)} LYD each`)))}
   ${anKpi(tr('متوسط المبيعات اليومية','Average daily sales'),money(avgDay),esc(tr(`د.ل · ${activeDays} يوم بيع`,`LYD · ${activeDays} sales days`)))}
   ${anKpi(tr('تركّز أعلى 10 زبائن','Top-10 concentration'),pct(top10v,A.net)+'%',esc(tr('من صافي المبيعات','of net sales')),pct(top10v,A.net)>=60?'warn':'')}
   ${anKpi(tr('المنتجات المباعة','Products sold'),fn(A.prods.length),esc(tp?tr(`الأعلى: ${tp.pn}`,`Top: ${tp.pn}`):''))}
  </div>
  <div class="an-grid sm-root">
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('قراءة تحليلية','Reading'))} <small>${esc(tr('تُكتب تلقائيًا من البيانات المصفّاة','written from the filtered data'))}</small></h3><ul class="sm-ins">${ins.map(x=>`<li><span class="ic ${x[0]}">${x[1]}</span><span>${x[2]}</span></li>`).join('')}</ul></section>
   <section class="panel an-wide"><h3 class="section-title">${esc(len<=62?tr('صافي المبيعات اليومية','Daily net sales'):tr('صافي المبيعات الشهرية','Monthly net sales'))} <small>${esc(tr('د.ل · مرّر على العمود للتفاصيل','LYD · hover a column for details'))}</small></h3>${colChart(pts,tr('صافي المبيعات حسب اليوم','Net sales by day'))}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('المنتجات','Products'))} <small>${esc(tr('الكمية بوحدة كل منتج، والقيمة بالدينار','quantity in each product’s unit, value in LYD'))}</small></h3>
    ${stack(A.prods.slice(0,3).map(p=>({l:p.pn,v:p.net,k:PC.get(p.p)||'sm-k4'})).concat(A.prods.length>3?[{l:tr('منتجات أخرى','Other products'),v:A.prods.slice(3).reduce((s,p)=>s+p.net,0),k:'sm-k4'}]:[]))}
    <div class="sm-prod">${A.prods.map(p=>`<div class="sm-pc"><h4><i class="${PC.get(p.p)||'sm-k4'}"></i>${esc(p.pn)} <small class="muted num">${esc(p.p)}</small></h4>
      <span class="q">${qty(p.q)} <small style="font-size:13px;font-weight:400">${esc(p.u||'')}</small></span>
      <span class="s">${esc(tr('القيمة','Value'))} <b class="num">${money2(p.v)}</b> · ${esc(tr('صافي','Net'))} <b class="num">${money2(p.net)}</b> ${esc(tr('د.ل','LYD'))}<br>${esc(tr('متوسط سعر الوحدة','Avg unit price'))} <b class="num">${p.q?fn(Math.round(p.v/p.q*1000)/1000):'—'}</b> · ${p.cust.size} ${esc(tr('زبون','customers'))}${p.rq||p.rv?`<br>${esc(tr('مسترجع','Returned'))} <b class="num">${qty(p.rq)}</b> ${esc(p.u||'')} · <b class="num">${money2(p.rv)}</b> ${esc(tr('د.ل','LYD'))}`:''}</span></div>`).join('')}</div></section>
   <section class="panel"><h3 class="section-title">${esc(tr('أعلى 10 زبائن','Top 10 customers'))} <small>${esc(tr('صافي المبيعات · اضغط للتفاصيل','net sales · click for details'))}</small></h3>${bars(top10.map(c=>({l:c.cn,v:c.net,go:c.c,tip:`${c.cn} (${c.c}): ${money2(c.net)} ${tr('د.ل','LYD')} · ${pct(c.net,A.net)}%`})))}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('أعلى المرتجعات','Largest returns'))} <small>${esc(tr('القيمة المسترجعة · د.ل','returned value · LYD'))}</small></h3>${retC.length?bars(retC.slice(0,10).map(c=>({l:c.cn,v:c.rv,go:c.c,tip:`${c.cn}: ${money2(c.rv)} ${tr('د.ل','LYD')} · ${pct(c.rv,c.v)}% ${tr('من مشترياته','of its purchases')}`})),{tone:'fl'}):`<div class="empty" style="padding:20px">${esc(tr('لا توجد مرتجعات','No returns'))}</div>`}</section>
   ${A.types.length>1?`<section class="panel"><h3 class="section-title">${esc(tr('حسب نوع الحركة','By line type'))} <small>${esc(tr('عمود «النوع» في التقرير','the report’s “Type” column'))}</small></h3>${stack(A.types.slice(0,3).map((t,i)=>({l:t.t,v:t.v,k:'sm-k'+(i+1)})).concat(A.types.length>3?[{l:tr('أخرى','Other'),v:A.types.slice(3).reduce((s,t)=>s+t.v,0),k:'sm-k4'}]:[]))}</section>`:''}
   ${B?`<section class="panel${A.types.length>1?'':' an-wide'}"><h3 class="section-title">${esc(tr('زبائن توقّفوا عن الشراء','Customers who stopped buying'))} <small>${esc(tr(`اشتروا في ${fdi(pFrom)} – ${fdi(pTo)}`,`bought in ${fdi(pFrom)} – ${fdi(pTo)}`))}</small></h3>${lost.length?bars(lost.slice(0,10).map(c=>({l:c.cn,v:c.net,go:c.c,tip:`${c.cn}: ${money2(c.net)} ${tr('د.ل في الفترة السابقة','LYD in the previous period')} · ${tr('آخر شراء','last purchase')} ${fdi(c.last)}`})),{tone:'fl'}):`<div class="empty" style="padding:20px">${esc(tr('كل زبائن الفترة السابقة اشتروا في هذه الفترة','Every previous-period customer bought again'))}</div>`}</section>`:''}
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('الزبائن','Customers'))} <small>${tblN} ${esc(tr('زبون · اضغط على الصف للتفاصيل','customers · click a row for details'))}</small></h3>
    <div class="toolbar" style="margin:0 0 8px"><div class="grow od-search">${I.search}<input type="search" id="sm-q" value="${esc(u.q||'')}" placeholder="${esc(tr('بحث باسم الزبون أو رقمه…','Search by customer name or number…'))}" aria-label="${esc(tr('بحث في الزبائن','Search customers'))}"></div></div>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">#</th>${th('cn',tr('الزبون','Customer'))}${th('n',tr('الحركات','Lines'),'n')}${th('days',tr('أيام الشراء','Days'),'n')}${th('v',tr('المبيعات','Sales'),'n')}${th('rv',tr('المرتجع','Returns'),'n')}${th('net',tr('الصافي','Net'),'n')}<th class="nosort n">${esc(tr('الحصة','Share'))}</th>${th('last',tr('آخر شراء','Last'),'n')}</tr></thead>
    <tbody>${tbl.map((c,i)=>`<tr data-sm-cust="${esc(c.c)}" tabindex="0"><td class="n">${i+1}</td><td class="t-strong">${esc(c.cn)} <small class="muted num">${esc(c.c)}</small></td><td class="n">${fn(c.n)}</td><td class="n">${c.days.size}</td><td class="n">${money2(c.v)}</td><td class="n">${c.rv?money2(c.rv):'—'}</td><td class="n"><b>${money2(c.net)}</b></td><td class="n">${pct(c.net,A.net)}%<span class="sm-share" style="width:${Math.max(0,c.net)*40/mxShare}px"></span></td><td class="n">${fdi(c.last)}</td></tr>`).join('')}</tbody></table></div>
    ${tblN>u.lim?`<div style="text-align:center;margin-top:8px"><button class="btn sm" data-sm="more">${esc(tr(`عرض المزيد (${tblN-u.lim} متبقٍ)`,`Show more (${tblN-u.lim} left)`))}</button></div>`:''}</section>
  </div>`}`}

function after(){const u=UI();
  [['sm-prod','prod'],['sm-type','type'],['sm-from','from'],['sm-to','to']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('change',()=>{u[k]=e.value;u.lim=25;render()})});
  const q=document.getElementById('sm-q');if(q)q.addEventListener('input',debounce(()=>{u.q=q.value;u.lim=25;const p=q.selectionStart;render();const n=document.getElementById('sm-q');if(n){n.focus();try{n.setSelectionRange(p,p)}catch(e){}}},250));
  const f=document.getElementById('sm-file');if(f)f.addEventListener('change',()=>{const file=f.files&&f.files[0];f.value='';if(file)importFile(file)})}

/* ---------- import ---------- */
async function importFile(file){toast(tr('جارٍ قراءة الملف…','Reading the file…'));let P;
  try{P=await parseFile(file)}catch(e){if(e&&e.cancelled)return;modal(tr('تعذّر الاستيراد','Import failed'),`<p style="line-height:1.8">${esc(e.message||String(e))}</p>`,`<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`);return}
  SM.pending=P;const ov=SM.local?[]:overlaps(P);const tot=P.rows.reduce((a,r)=>{a.v+=r.v;a.rv+=r.rv;return a},{v:0,rv:0});const cs=new Set(P.rows.map(r=>r.c)).size;
  const FL=EN()?FIELD_EN:FIELD_AR;
  modal(tr('مراجعة ملف SMS055 قبل الحفظ','Review the SMS055 file'),`
   <div class="sm-mini"><div><b class="num">${fn(P.rows.length)}</b><span>${esc(tr('حركة مقروءة','lines read'))}</span></div><div><b class="num">${fdi(P.from)} – ${fdi(P.to)}</b><span>${esc(tr('الفترة','Period'))}</span></div>
    <div><b class="num">${money(tot.v-tot.rv)}</b><span>${esc(tr('صافي المبيعات د.ل','Net sales LYD'))}</span></div><div><b class="num">${fn(cs)}</b><span>${esc(tr('زبون','customers'))}</span></div></div>
   <p class="muted" style="margin:4px 0;font-size:12.5px">${esc(tr(`الملف: ${P.file} · الورقة: ${P.sheet} · صف العناوين: ${P.header}`,`File: ${P.file} · sheet: ${P.sheet} · header row: ${P.header}`))}${P.skipped?` · <b style="color:var(--warn)">${esc(tr(`تُجوهل ${P.skipped} صف بلا تاريخ أو زبون أو قيمة`,`${P.skipped} rows skipped (no date, customer or value)`))}</b>`:''}</p>
   <div class="sm-cols-map">${Object.keys(FIELDS).map(k=>`<span class="${P.map[k]!=null?'ok':'no'}">${P.map[k]!=null?'✓':'✗'} ${esc(FL[k])}</span>`).join('')}</div>
   ${ov.length?`<div class="banner warn" style="margin-top:8px">${esc(tr(`تتداخل هذه الفترة مع ${ov.length} ملف محفوظ سابقًا:`,`This period overlaps ${ov.length} saved file(s):`))} ${ov.map(b=>esc(`${b.file} (${fdi(b.from)} – ${fdi(b.to)})`)).join('، ')}
     <label style="display:flex;gap:6px;align-items:center;margin-top:6px"><input type="checkbox" id="sm-replace" checked> ${esc(tr('استبدال الملفات المتداخلة بهذا الملف (يمنع تكرار الأرقام)','Replace the overlapping files with this one (avoids double counting)'))}</label></div>`:''}
   ${!SM.canImport?`<p class="muted" style="font-size:12.5px">${esc(tr('الحفظ في المنظومة متاح لمحرري إدارة الشؤون التجارية والتحصيلات والإدارة العامة ومدير النظام. يمكنك عرض التحليل مؤقتًا على جهازك.','Saving is for editors of Commercial Affairs & Collections, the General Administration and the administrator. You can view the analysis on this device.'))}</p>`:''}
   <p id="sm-msg" style="font-size:12.5px;margin:6px 0 0"></p>`,
   `${SM.canImport?`<button type="button" class="btn primary" data-sm="save">${esc(tr('حفظ في المنظومة','Save to the system'))}</button>`:''}<button type="button" class="btn" data-sm="view-local">${esc(tr('عرض دون حفظ','View without saving'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}

/* ---------- customer detail ---------- */
function custDetail(c){const R=range(),rows=pick(data(),R.from,R.to).filter(r=>r.c===c);if(!rows.length)return;const A=agg(rows),x=A.custs[0];
  const byDay=[...A.day.values()].sort((a,b)=>b.d.localeCompare(a.d));
  modal(`${x.cn} (${x.c})`,`<div class="sm-mini"><div><b class="num">${money2(x.net)}</b><span>${esc(tr('صافي د.ل','Net LYD'))}</span></div><div><b class="num">${money2(x.rv)}</b><span>${esc(tr('مرتجع د.ل','Returns LYD'))}</span></div><div><b class="num">${x.days.size}</b><span>${esc(tr('يوم شراء','purchase days'))}</span></div><div><b class="num">${fdi(x.last)}</b><span>${esc(tr('آخر شراء','last purchase'))}</span></div></div>
   <h4 style="margin:8px 0 4px">${esc(tr('حسب المنتج','By product'))}</h4><div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th><th class="nosort n">${esc(tr('القيمة','Value'))}</th><th class="nosort n">${esc(tr('المسترجع','Returned'))}</th></tr></thead><tbody>${A.prods.map(p=>`<tr><td>${esc(p.pn)}</td><td class="n">${qty(p.q)} ${esc(p.u)}</td><td class="n">${money2(p.v)}</td><td class="n">${p.rv?money2(p.rv):'—'}</td></tr>`).join('')}</tbody></table></div>
   <h4 style="margin:10px 0 4px">${esc(tr('حسب اليوم','By day'))}</h4><div class="tablewrap" style="max-height:260px;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('التاريخ','Date'))}</th><th class="nosort n">${esc(tr('الحركات','Lines'))}</th><th class="nosort n">${esc(tr('المبيعات','Sales'))}</th><th class="nosort n">${esc(tr('المرتجع','Returns'))}</th></tr></thead><tbody>${byDay.map(d=>`<tr><td class="n">${fdi(d.d)}</td><td class="n">${d.n}</td><td class="n">${money2(d.v)}</td><td class="n">${d.rv?money2(d.rv):'—'}</td></tr>`).join('')}</tbody></table></div>`,
   `<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- export & print ---------- */
function exportXlsx(){const R=range(),rows=pick(data(),R.from,R.to),A=agg(rows);
  const sum=[{[tr('البند','Item')]:tr('الفترة','Period'),[tr('القيمة','Value')]:`${R.from} – ${R.to}`},{[tr('البند','Item')]:tr('إجمالي المبيعات (د.ل)','Gross sales (LYD)'),[tr('القيمة','Value')]:Math.round(A.v*100)/100},
    {[tr('البند','Item')]:tr('المرتجعات (د.ل)','Returns (LYD)'),[tr('القيمة','Value')]:Math.round(A.rv*100)/100},{[tr('البند','Item')]:tr('صافي المبيعات (د.ل)','Net sales (LYD)'),[tr('القيمة','Value')]:Math.round(A.net*100)/100},
    {[tr('البند','Item')]:tr('عدد الحركات','Lines'),[tr('القيمة','Value')]:A.n},{[tr('البند','Item')]:tr('الزبائن النشطون','Active customers'),[tr('القيمة','Value')]:A.custs.length},{[tr('البند','Item')]:tr('أيام البيع','Sales days'),[tr('القيمة','Value')]:A.day.size}];
  const prod=A.prods.map(p=>({'رقم المنتج':p.p,'اسم المنتج':p.pn,'وحدة القياس':p.u,'الكمية':Math.round(p.q*1000)/1000,'القيمة':Math.round(p.v*100)/100,'الكمية المسترجعة':Math.round(p.rq*1000)/1000,'القيمة المسترجعة':Math.round(p.rv*100)/100,'الصافي':Math.round(p.net*100)/100,'عدد الزبائن':p.cust.size}));
  const cust=A.custs.slice().sort((a,b)=>b.net-a.net).map((c,i)=>({'م':i+1,'رقم العميل':c.c,'اسم الزبون':c.cn,'الحركات':c.n,'أيام الشراء':c.days.size,'المبيعات':Math.round(c.v*100)/100,'المرتجع':Math.round(c.rv*100)/100,'الصافي':Math.round(c.net*100)/100,'الحصة %':pct(c.net,A.net),'آخر شراء':c.last}));
  const day=[...A.day.values()].sort((a,b)=>a.d.localeCompare(b.d)).map(d=>({'التاريخ':d.d,'الحركات':d.n,'الزبائن':d.cust.size,'المبيعات':Math.round(d.v*100)/100,'المرتجع':Math.round(d.rv*100)/100,'الصافي':Math.round((d.v-d.rv)*100)/100}));
  saveXlsx([[tr('ملخص','Summary'),sum],[tr('المنتجات','Products'),prod],[tr('الزبائن','Customers'),cust],[tr('يومي','Daily'),day]],`SMS055-analysis-${R.from}_${R.to}.xlsx`)}
function printView(){try{window.print()}catch(e){}}

/* ---------- files dialog ---------- */
function filesDlg(){modal(tr('ملفات SMS055 المستوردة','Imported SMS055 files'),`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl"><thead><tr><th class="nosort">${esc(tr('الملف','File'))}</th><th class="nosort">${esc(tr('الفترة','Period'))}</th><th class="nosort n">${esc(tr('الصفوف','Rows'))}</th><th class="nosort n">${esc(tr('الصافي د.ل','Net LYD'))}</th><th class="nosort">${esc(tr('استورده','Imported by'))}</th>${SM.canImport?'<th class="nosort"></th>':''}</tr></thead>
  <tbody>${SM.batches.slice().reverse().map(b=>`<tr><td>${esc(b.file||'—')}${b.ok===false?` <span class="pill p-crit">${esc(tr('غير مكتمل','incomplete'))}</span>`:''}</td><td class="n">${fdi(b.from)} – ${fdi(b.to)}</td><td class="n">${fn(b.rows)}</td><td class="n">${money2((b.value||0)-(b.returns||0))}</td><td><small>${esc(b.by||'')}<br>${fdi(b.at)}</small></td>${SM.canImport?`<td><button type="button" class="btn sm ghost danger" data-sm-del="${esc(b.id)}">${esc(tr('حذف','Delete'))}</button></td>`:''}</tr>`).join('')}</tbody></table></div>`,
  `<button type="button" class="btn" data-sm="reload">${esc(tr('تحديث','Refresh'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- events ---------- */
document.addEventListener('click',async e=>{
  const pre=e.target.closest('[data-sm-pre]');if(pre){const u=UI();u.pre=pre.dataset.smPre;if(u.pre==='custom'){const R=range();u.from=u.from||R.from;u.to=u.to||R.to}u.lim=25;render();return}
  const so=e.target.closest('[data-sm-sort]');if(so){const u=UI(),k=so.dataset.smSort;if(u.sort===k)u.dir=u.dir==='asc'?'desc':'asc';else{u.sort=k;u.dir=k==='cn'?'asc':'desc'}render();return}
  const cu=e.target.closest('[data-sm-cust]');if(cu&&document.getElementById('view').contains(cu)){custDetail(cu.dataset.smCust);return}
  const del=e.target.closest('[data-sm-del]');if(del){if(!del.classList.contains('armed')){del.classList.add('armed');del.textContent=tr('تأكيد الحذف','Confirm delete');setTimeout(()=>{del.classList.remove('armed');del.textContent=tr('حذف','Delete')},4000);return}
    const b=SM.batches.find(x=>x.id===del.dataset.smDel);if(!b)return;del.disabled=true;try{await delBatch(b);toast(tr('تم حذف الملف','File deleted'));await load(true);filesDlg();render()}catch(err){toast(tr('تعذّر الحذف: ','Could not delete: ')+(err.message||''));del.disabled=false}return}
  const a=e.target.closest('[data-sm]');if(!a)return;const act=a.dataset.sm;
  if(act==='import'){const f=document.getElementById('sm-file');if(f)f.click();return}
  if(act==='export'){exportXlsx();return}
  if(act==='print'){printView();return}
  if(act==='files'){filesDlg();return}
  if(act==='more'){UI().lim+=50;render();return}
  if(act==='clear'){Object.assign(UI(),{pre:'all',from:'',to:'',prod:'',type:''});render();return}
  if(act==='drop-local'){SM.local=null;render();return}
  if(act==='reload'){a.disabled=true;await load(true);dlg.close();render();toast(tr('تم تحديث البيانات','Data refreshed'));return}
  if(act==='view-local'){SM.local=SM.pending;SM.pending=null;dlg.close();Object.assign(UI(),{pre:'all',prod:'',type:'',q:''});render();return}
  if(act==='save'){const P=SM.pending;if(!P)return;const msg=document.getElementById('sm-msg');a.disabled=true;msg.style.color='';msg.textContent=tr('جارٍ الحفظ…','Saving…');
    try{const rep=document.getElementById('sm-replace');await saveImport(P,!!(rep&&rep.checked));SM.pending=null;SM.local=null;await load(true);dlg.close();Object.assign(UI(),{pre:'all',prod:'',type:'',q:''});render();toast(tr(`تم حفظ ${fn(P.rows.length)} حركة`,`${fn(P.rows.length)} lines saved`))}
    catch(err){a.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية الحفظ في هذه اللوحة','You are not allowed to save to this dashboard'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}}
});
document.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;const cu=e.target.closest&&e.target.closest('[data-sm-cust]');if(cu&&document.getElementById('view').contains(cu)){e.preventDefault();custDetail(cu.dataset.smCust)}});

/* ---------- navigation: a task of Commercial Affairs & Collections ---------- */
if(window.I18N&&I18N.add)I18N.add({'لوحة مؤشرات SMS055':'SMS055 dashboard','مهام إدارة الشؤون التجارية والتحصيلات':'Commercial Affairs & Collections tasks'});
VIEWS.push({id:'sms055',label:'لوحة مؤشرات SMS055',icon:'sms',group:'مهام إدارة الشؤون التجارية والتحصيلات',nav:false});
ORG_TASKS.ca=(ORG_TASKS.ca||[]).concat(['sms055']);
const __render=render;
render=function(){if(S.view==='sms055'){recompute();document.getElementById('view').innerHTML=dbBanner()+viewSms();after();return}__render()};
/* the department's page in the org dashboard gets a summary card with a link */
const __rp=recPanel;
recPanel=function(id){const h=__rp(id);if(id!==UNIT)return h;
  let body;if(!SM.loaded){load().then(()=>{if(S.view==='org')render()});body=`<p class="muted">${esc(tr('جارٍ التحميل…','Loading…'))}</p>`}
  else if(!SM.rows.length)body=`<p class="muted" style="font-size:13px;line-height:1.8">${esc(SM.err==='denied'?tr('البيانات متاحة للإدارة والإدارة العامة.','Data is available to the department and the General Administration.'):tr('لم يُستورد أي ملف بعد. استورد تقرير SMS055 لتظهر المؤشرات.','No file imported yet. Import the SMS055 report to see the KPIs.'))}</p>`;
  else{const bd=bounds(SM.rows),from=bd.max.slice(0,8)+'01',A=agg(SM.rows.filter(r=>r.d>=from));
    body=`<div class="sm-mini"><div><b class="num">${money(A.net)}</b><span>${esc(tr('صافي آخر شهر د.ل','Net, latest month LYD'))}</span></div><div><b class="num">${fn(A.custs.length)}</b><span>${esc(tr('زبون نشط','active customers'))}</span></div><div><b class="num">${pct(A.rv,A.v)}%</b><span>${esc(tr('نسبة المرتجع','return rate'))}</span></div></div><p class="muted" style="font-size:12px;margin:0">${esc(tr(`${fdi(from)} – ${fdi(bd.max)}`,`${fdi(from)} – ${fdi(bd.max)}`))}</p>`}
  return `<section class="panel"><h3 class="section-title">${esc(tr('مخرجات منظومة SMS055','SMS055 outputs'))} <small>${esc(tr('المبيعات والمرتجعات والزبائن','sales, returns and customers'))}</small></h3>${body}<button type="button" class="btn sm primary" data-org-go="sms055" style="margin-top:8px">${esc(tr('فتح لوحة مؤشرات SMS055','Open the SMS055 dashboard'))}</button></section>`+h};
/* duties guide: a shortcut next to the department's duties */
const __vod=viewOrgDuties;
viewOrgDuties=function(el){__vod(el);const sec=el.querySelector('#od-'+UNIT+' .od-acts');if(sec&&!sec.querySelector('[data-org-go="sms055"]'))sec.insertAdjacentHTML('afterbegin',`<button type="button" class="btn sm primary" data-org-go="sms055">${esc(tr('لوحة SMS055','SMS055 dashboard'))}</button>`)};
window.__sms=SM;window.__smsParse=parseFile;
/* shared with the other closing-reports dashboards (closures.js): file readers, number/date helpers and the small chart builders */
window.__smsLib={readGrids,hn,AD,num,serialIso,niceMax,colChart,bars,stack,money,money2,pct,isoAdd,days,fdi,mName,me};
})();
