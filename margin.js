/* ===== لوحة مؤشرات هامش الربح (إدارة الشؤون التجارية والتحصيلات) =====
   Loaded after closures.js. Quantities come from the «تحليل المبيعات» data (Sales_Report, collection sales_report,
   shared with the sales dashboard so a month is never counted twice). Prices and costs live in one Firestore document,
   margin/prices, with an effective-dated history per product, so a price change never rewrites past months:
     products: [{p, pn, u, h:[{from, cost, loc, fx, by, at, why}]}]   prices per base unit (m³ for fuels, else the report unit)
     rates:    [{from, rate}]   local currency per one unit of the foreign currency
     cur, target, min, foreign: [customer account numbers that pay in the foreign currency]
   Two price lists: «loc» in Libyan dinars for local consumers, «fx» in the foreign currency for foreign customers.
   Definitions (shown in the page): gross profit = revenue − cost of goods sold; margin % = profit ÷ revenue;
   markup % = profit ÷ cost. Ratios are always computed from summed totals, never averaged across lines.
   Read by the department, the General Administration, the board and the administrator; written by the administrator
   or an editor of ca / gm (firestore.rules). Every change is written to the activity log with old and new values. */
"use strict";
(function(){
const L=window.__smsLib;if(!L){console.warn('margin: sms055 library missing');return}
const {readGrids,hn,colChart,bars,stack,pct,fdi,mName,me}=L;
const CX=()=>window.__cx;
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const UNIT='ca',COL='margin',DOC='prices',VIEW='margin';
const PM={loaded:false,loading:null,err:'',cfg:null,canEdit:false,edit:false,pendingPrices:null,pendingSales:null,local:null};
const DEF=()=>({products:[],rates:[],cur:'USD',target:15,min:5,foreign:[]});
const CURS=[['USD','دولار أمريكي','US dollar'],['EUR','يورو','Euro'],['GBP','جنيه إسترليني','Pound sterling'],['TND','دينار تونسي','Tunisian dinar'],['EGP','جنيه مصري','Egyptian pound']];

/* ---------- styles (on top of the .sm-* classes of sms055.js) ---------- */
(function(){const st=document.createElement('style');st.textContent=`
.pm-page .sm-hero{--hero-glow:rgba(253,224,71,.16);--hero-a:#3d2a05;--hero-b:#a16207;--hero-c:#7c4a03}
.pm-root.sm-root{--sm-c1:#0f766e;--sm-c2:#c2410c;--sm-c3:#a16207;--sm-c4:#8b9099}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .pm-root.sm-root{--sm-c1:#14b8a6;--sm-c2:#f97316;--sm-c3:#eab308}}
:root[data-theme="dark"] .pm-root.sm-root{--sm-c1:#14b8a6;--sm-c2:#f97316;--sm-c3:#eab308}
.pm-units{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12.5px;color:rgba(255,255,255,.85);margin-top:6px}.pm-units b{font-variant-numeric:tabular-nums;color:#fff}
.pm-hero-m{display:flex;gap:18px;flex-wrap:wrap;justify-content:center}.pm-hero-m>div{text-align:center}
.pm-tbl td,.pm-tbl th{vertical-align:middle}
.pm-tbl td.n,.pm-tbl th.n{text-align:end;font-variant-numeric:tabular-nums;white-space:nowrap}
.pm-tbl input[type=number]{width:96px;height:32px;border:1px solid var(--line);border-radius:7px;background:var(--surface);color:var(--ink);font:inherit;font-variant-numeric:tabular-nums;padding-inline:6px;text-align:end}
.pm-tbl input[type=number]:focus-visible,.pm-set input:focus-visible,.pm-set select:focus-visible,.pm-tool input:focus-visible,.pm-tool select:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.pm-tbl .sub{display:block;font-size:11.5px;color:var(--muted);font-weight:400}
.pm-tbl tr.pm-chg td{background:var(--accent-soft)}
.pm-tbl tr[data-pm-cust]{cursor:pointer}.pm-tbl th[data-pm-sort]{cursor:pointer;white-space:nowrap}.pm-tbl th[data-pm-sort][aria-sort]::after{content:" ▾";font-size:10px}.pm-tbl th[data-pm-sort][aria-sort="ascending"]::after{content:" ▴"}
.pm-neg{color:var(--crit)}.pm-pos{color:var(--ok)}
.pm-set{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin:0 0 10px}
.pm-set label{display:grid;gap:4px;font-size:12.5px;color:var(--muted)}
.pm-set input,.pm-set select,.pm-tool input,.pm-tool select{height:34px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);font:inherit;padding-inline:8px;min-width:0}
.pm-set input[type=number]{width:110px}.pm-set input[type=text]{width:220px}
.pm-tools{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,320px),1fr));gap:14px}
.pm-tool{display:grid;gap:10px;align-content:start}
.pm-tool .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px}
.pm-tool label{display:grid;gap:4px;font-size:12.5px;color:var(--muted)}
.pm-tool input[type=range]{width:100%;accent-color:var(--accent)}
.pm-out{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px}
.pm-out div{border:1px solid var(--line);border-radius:10px;padding:8px 10px;background:var(--surface)}
.pm-out b{display:block;font-size:17px;font-variant-numeric:tabular-nums}.pm-out span{font-size:12px;color:var(--muted)}
.pm-out .good b{color:var(--ok)}.pm-out .bad b{color:var(--crit)}
.pm-def{margin:0;padding:0;list-style:none;display:grid;gap:6px;font-size:13px;line-height:1.75}
.pm-def b{color:var(--ink)}
.pm-seg{display:inline-block;font-size:11.5px;padding:1px 8px;border-radius:999px;border:1px solid var(--line);white-space:nowrap}
.pm-seg.fx{background:color-mix(in srgb,var(--sm-c2,#c2410c) 14%,transparent);border-color:color-mix(in srgb,var(--sm-c2,#c2410c) 45%,transparent)}
button.pm-seg{cursor:pointer;font:inherit;font-size:11.5px;background:transparent;color:var(--ink)}
button.pm-seg.fx{background:color-mix(in srgb,#c2410c 14%,transparent)}
.pm-gauge{height:8px;border-radius:4px;background:var(--line);position:relative;overflow:hidden;min-width:70px}
.pm-gauge i{position:absolute;inset-block:0;inset-inline-start:0;background:var(--ok);border-radius:4px}
.pm-gauge i.w{background:var(--warn)}.pm-gauge i.c{background:var(--crit)}
@media print{.pm-tools,.pm-edit-only{display:none}}
`;document.head.appendChild(st)})();

/* ---------- formatting ---------- */
const qs=v=>{if(v==null||isNaN(v))return'—';const a=Math.abs(v);if(a>=1e6)return +(v/1e6).toFixed(a>=1e7?0:1)+tr(' مليون',' M');if(a>=1e4)return fn(Math.round(v/1e3))+tr(' ألف',' K');return fn(Math.round(v*10)/10)};
const q2=v=>v==null||isNaN(v)?'—':fn(Math.round(v*100)/100);
const q3=v=>v==null||isNaN(v)?'—':fn(Math.round(v*1000)/1000);
const p4=v=>v==null||isNaN(v)?'—':fn(Math.round(v*10000)/10000);
const pc=v=>v==null||!isFinite(v)?'—':(Math.round(v*10)/10)+'%';
const bd=v=>`<bdi>${esc(v)}</bdi>`;
const LYD=()=>tr('د.ل','LYD');
const M3=()=>tr('م³','m³');
const today=()=>new Date().toISOString().slice(0,10);
const r4=v=>Math.round(v*10000)/10000;
const toNum=v=>{if(v==null)return null;const s=String(v).trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[,\s٬]/g,'').replace('٫','.');if(s==='')return null;const n=parseFloat(s);return isFinite(n)?n:null};
const curName=c=>{const x=CURS.find(k=>k[0]===c);return x?tr(x[1],x[2]):c};

/* ---------- units: fuels in m³ ---------- */
/* the shared converter (units.js, window.FuelUnits) when present, else litres ÷ 1000, tonnes × 1000 ÷ typical density15 */
const DENS=[[/سفن|بحري|marine|bunker|mgo/i,890],[/بنزين|gasoline|petrol/i,740],[/ديزل|سولار|diesel|gas ?oil/i,835],[/كيروسين|kerosene|jet/i,800],[/ثقيل|مازوت|زيت وقود|fuel oil|hfo/i,970],[/غاز|مسال|lpg/i,540]];
const SERVICE=/خدمات|خدمة|صيانة|نقل وتشغيل|توريد|service|maintenance|transport/i;
function m3Of(q,u,p,pn,tc){const F=window.FuelUnits;if(!F&&SERVICE.test(pn||'')&&!DENS.some(x=>x[0].test(pn||'')))return null;if(F&&F.toM3){try{const x=F.toM3(q,u,p,pn,tc);return x?x.q:null}catch(e){}}
  const k=String(u||'').toLowerCase();if(/م3|م³|مكعب|m3|m³|cbm/.test(k))return +q;if(/لتر|ليتر|litre|liter|ltr/.test(k))return q/1000;
  if(/طن|ton/.test(k)){const d=(DENS.find(x=>x[0].test(pn||''))||[])[1];return d?q*1000/d:null}return null}
function baseOf(r){const m=m3Of(r.q,r.u,r.p,r.pn,r.tc);return m==null?{q:r.q,u:r.u||'',m3:false}:{q:m,u:'م³',m3:true}}
const uLbl=u=>u==='م³'?M3():u;

/* ---------- storage ---------- */
async function load(force){if(PM.loading&&!force)return PM.loading;
  PM.loading=(async()=>{const FS=window.__fs&&window.__fs();const cx=CX();
    if(cx&&cx.load)try{await cx.load('sales',force)}catch(e){}
    if(!FS){PM.cfg=DEF();PM.loaded=true;return}
    try{const d=await FS.collection(COL).doc(DOC).get();PM.cfg=Object.assign(DEF(),d.exists?d.data():{});PM.err=''}
    catch(e){PM.err=e&&e.code==='permission-denied'?'denied':(e&&e.message)||'error';PM.cfg=DEF()}
    if(PM.err==='denied'&&window.__denyPrep)await window.__denyPrep();
    ['products','rates','foreign'].forEach(k=>{if(!Array.isArray(PM.cfg[k]))PM.cfg[k]=[]});
    PM.loaded=true;
    try{const u=await recMyUnit();const m=me();PM.canEdit=m.role==='admin'||(!!S.canWrite&&m.role!=='viewer'&&(u===UNIT||u==='gm'))}catch(e){PM.canEdit=false}
  })();return PM.loading}
async function saveCfg(cfg,action,summary){const FS=window.__fs(),m=me();
  const doc={products:cfg.products.slice(0,400),rates:cfg.rates.slice(-400),cur:String(cfg.cur||'USD').slice(0,8),target:+cfg.target||0,min:+cfg.min||0,
    foreign:cfg.foreign.map(String).slice(0,5000),updatedBy:m.email,by:String(m.name||m.email||'').slice(0,120),at:new Date().toISOString()};
  await FS.collection(COL).doc(DOC).set(doc);PM.cfg=Object.assign(DEF(),doc);
  try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action,col:COL,id:DOC,summary:String(summary).slice(0,1500)})}catch(e){}}

/* ---------- prices in effect ---------- */
function prodCfg(p){return PM.cfg.products.find(x=>x.p===p)}
function hist(p){const P=prodCfg(p);return P&&Array.isArray(P.h)?P.h.slice().sort((a,b)=>String(a.from).localeCompare(String(b.from))):[]}
/* the entry in effect on day d; before the first entry the first one applies (flagged in the reading) */
function priceAt(p,d){const H=hist(p);if(!H.length)return null;let e=null;H.forEach(h=>{if(h.from<=d)e=h});return e||H[0]}
function current(p){return priceAt(p,'9999-12-31')}
function rateAt(d){const R=PM.cfg.rates.slice().sort((a,b)=>String(a.from).localeCompare(String(b.from)));if(!R.length)return null;let e=null;R.forEach(x=>{if(x.from<=d)e=x});return +(e||R[0]).rate||null}
function rateNow(){return rateAt('9999-12-31')}
let FXSET=new Set();
const isFx=c=>FXSET.has(String(c));

/* ---------- data ---------- */
function UI(){return S.ui.pm||(S.ui.pm={tab:'full',pre:'all',from:'',to:'',prod:'',seg:'',rgn:'',wn:'',q:'',sort:'gp',dir:'desc',lim:25,sim:{p:'',seg:'loc',chg:5},calc:{cost:'',m:'',price:''}})}
function src(){const cx=CX();if(PM.local)return PM.local.rows;return cx&&cx.ST&&cx.ST.sales?cx.ST.sales.rows:[]}
function rowsAll(){return src().map(r=>{const b=baseOf(r);return {d:r.d,c:r.c||r.cn||'',cn:r.cn||r.c||'',p:r.p||r.pn||'',pn:r.pn||r.p||'',u0:r.u||'',q0:r.q,q:b.q,u:b.u,m3:b.m3,wn:r.wn||'',rgn:r.rgn||''}})}
function bounds(rows){if(!rows.length)return null;let a=rows[0].d,b=a;rows.forEach(r=>{if(r.d<a)a=r.d;if(r.d>b)b=r.d});return {min:a,max:b}}
function addMonths(d,n){const y=+d.slice(0,4),m=+d.slice(5,7)-1+n;return new Date(Date.UTC(y,m,1)).toISOString().slice(0,10)}
function monthsBetween(a,b){return (+b.slice(0,4)-+a.slice(0,4))*12+(+b.slice(5,7)-+a.slice(5,7))+1}
function range(rows){const u=UI(),bd=bounds(rows);if(!bd)return null;let from=bd.min.slice(0,7)+'-01',to=bd.max.slice(0,7)+'-01';
  if(u.pre==='month')from=to;else if(u.pre==='3')from=addMonths(to,-2);else if(u.pre==='12')from=addMonths(to,-11);
  else if(u.pre==='custom'){from=(u.from||bd.min).slice(0,7)+'-01';to=(u.to||bd.max).slice(0,7)+'-01';if(from>to)[from,to]=[to,from]}
  if(from<bd.min.slice(0,7)+'-01'&&u.pre!=='custom')from=bd.min.slice(0,7)+'-01';return {from,to,min:bd.min,max:bd.max}}
function pick(rows,from,to){const u=UI(),toE=to.slice(0,7)+'-31';return rows.filter(r=>r.d>=from&&r.d<=toE&&(!u.prod||r.p===u.prod)&&(!u.rgn||r.rgn===u.rgn)&&(!u.wn||r.wn===u.wn)&&(!u.seg||(u.seg==='fx')===isFx(r.c)))}

/* one sale line valued at the prices in effect for its month */
function line(r){const fx=isFx(r.c),e=priceAt(r.p,r.d);
  if(!e||e.cost==null)return {fx,ok:false,why:'cost'};
  let up,fxu=0;if(fx){const rt=rateAt(r.d);if(e.fx==null)return {fx,ok:false,why:'fx'};if(!rt)return {fx,ok:false,why:'rate'};fxu=+e.fx;up=fxu*rt}
  else{if(e.loc==null)return {fx,ok:false,why:'loc'};up=+e.loc}
  const rev=r.q*up,cogs=r.q*(+e.cost);return {fx,ok:true,rev,cogs,gp:rev-cogs,fxRev:fx?r.q*fxu:0,up,cost:+e.cost}}
function agg(rows){const A={n:rows.length,rev:0,cogs:0,gp:0,fxRev:0,loc:{rev:0,cogs:0,gp:0,q:0},fx:{rev:0,cogs:0,gp:0,q:0},qU:new Map(),prod:new Map(),cust:new Map(),mon:new Map(),wn:new Map(),rgn:new Map(),miss:new Map(),missQ:0,okQ:0};
  rows.forEach(r=>{const x=line(r);A.qU.set(r.u,(A.qU.get(r.u)||0)+r.q);
    let p=A.prod.get(r.p);if(!p)A.prod.set(r.p,p={p:r.p,pn:r.pn,u:r.u,q:0,rev:0,cogs:0,gp:0,qLoc:0,qFx:0,gpLoc:0,gpFx:0,revLoc:0,revFx:0,fxRev:0,miss:0,cust:new Set()});p.q+=r.q;p.cust.add(r.c);
    let c=A.cust.get(r.c);if(!c)A.cust.set(r.c,c={c:r.c,cn:r.cn,fx:isFx(r.c),q:new Map(),rev:0,cogs:0,gp:0,n:0,prods:new Set(),fxRev:0});c.n++;c.q.set(r.u,(c.q.get(r.u)||0)+r.q);c.prods.add(r.p);
    const mk=r.d.slice(0,7);let mo=A.mon.get(mk);if(!mo)A.mon.set(mk,mo={k:mk,rev:0,cogs:0,gp:0,fxGp:0,locGp:0});
    if(!x.ok){p.miss+=r.q;A.missQ+=r.q;let mm=A.miss.get(r.p);if(!mm)A.miss.set(r.p,mm={p:r.p,pn:r.pn,u:r.u,q:0,why:new Set()});mm.q+=r.q;mm.why.add(x.why);return}
    A.okQ+=r.q;A.rev+=x.rev;A.cogs+=x.cogs;A.gp+=x.gp;A.fxRev+=x.fxRev;const sg=x.fx?A.fx:A.loc;sg.rev+=x.rev;sg.cogs+=x.cogs;sg.gp+=x.gp;sg.q+=r.q;
    p.rev+=x.rev;p.cogs+=x.cogs;p.gp+=x.gp;p.fxRev+=x.fxRev;if(x.fx){p.qFx+=r.q;p.gpFx+=x.gp;p.revFx+=x.rev}else{p.qLoc+=r.q;p.gpLoc+=x.gp;p.revLoc+=x.rev}
    c.rev+=x.rev;c.cogs+=x.cogs;c.gp+=x.gp;c.fxRev+=x.fxRev;mo.rev+=x.rev;mo.cogs+=x.cogs;mo.gp+=x.gp;if(x.fx)mo.fxGp+=x.gp;else mo.locGp+=x.gp;
    [['wn',r.wn],['rgn',r.rgn]].forEach(([k,v])=>{if(!v)return;let g=A[k].get(v);if(!g)A[k].set(v,g={k:v,gp:0,rev:0});g.gp+=x.gp;g.rev+=x.rev})});
  A.m=A.rev?A.gp*100/A.rev:null;A.mk=A.cogs?A.gp*100/A.cogs:null;
  A.prods=[...A.prod.values()].sort((a,b)=>b.gp-a.gp);A.custs=[...A.cust.values()];return A}

/* ---------- page parts ---------- */
I.pm='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7.5 3.5c2.7 3.5 4.2 5.8 4.2 8a4.2 4.2 0 0 1-8.4 0c0-2.2 1.5-4.5 4.2-8z"/><rect x="14.5" y="4" width="6.5" height="16" rx="1.6"/><path d="M14.5 9.5h6.5M14.5 14h6.5"/><path d="M3 20.5h9"/></svg>';
function heroHtml(sub,A,delta,units){return `<div class="sm-hero"><div><span class="sm-eyebrow">${esc(tr('إدارة الشؤون التجارية والتحصيلات','Commercial Affairs & Collections'))}</span>
  <h2>${esc(tr('لوحة مؤشرات هامش الربح','Profit margin dashboard'))}</h2><p>${sub}</p>${units||''}</div>
  ${A?`<div class="pm-hero-m"><div class="sm-hero-v"><b>${qs(A.gp)}</b><span>${esc(tr('إجمالي الربح · د.ل','Gross profit · LYD'))}</span>${delta||''}</div><div class="sm-hero-v"><b>${pc(A.m)}</b><span>${esc(tr('هامش الربح','Gross margin'))}</span></div></div>`:''}</div>`}
function toolbar(){const has=src().length;const T=window.__pmx;if(T&&UI().tab==='full')return `<div class="toolbar sm-acts">
  <button class="btn primary" data-pm="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button>
  ${T.has()?`<button class="btn" data-px="export">${I.down}${esc(tr('تصدير التحليل Excel','Export analysis (Excel)'))}</button><button class="btn" data-pm="print">${esc(tr('طباعة / PDF','Print / PDF'))}</button>`:''}
  <button class="btn ghost" data-px="template">${esc(tr('قالب بيانات البريقة','Brega data template'))}</button>
  <input type="file" id="pm-file" accept="${FileImport.ACCEPT}" hidden></div>`;
  return `<div class="toolbar sm-acts">
  <button class="btn primary" data-pm="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button>
  ${has?`<button class="btn" data-pm="export">${I.down}${esc(tr('تصدير التحليل Excel','Export analysis (Excel)'))}</button><button class="btn" data-pm="print">${esc(tr('طباعة / PDF','Print / PDF'))}</button>`:''}
  <button class="btn ghost" data-pm="template">${esc(tr('نموذج قائمة الأسعار','Price-list template'))}</button>
  ${PM.local?`<button class="btn ghost" data-pm="drop-local">${esc(tr('إغلاق العرض المؤقت','Close temporary view'))}</button>`:''}
  <input type="file" id="pm-file" accept="${FileImport.ACCEPT}" hidden></div>`}
function statusOf(m){const c=PM.cfg;if(m==null)return ['p-neutral',tr('بلا سعر','No price'),''];if(m<0)return ['p-crit',tr('خسارة','Loss'),'c'];if(m<+c.min)return ['p-crit',tr('تحت الحد الأدنى','Below minimum'),'c'];if(m<+c.target)return ['p-warn',tr('تحت المستهدف','Below target'),'w'];return ['p-ok',tr('ضمن المستهدف','On target'),'']}
const mOf=(price,cost)=>price!=null&&cost!=null&&price>0?(price-cost)*100/price:null;
const kOf=(price,cost)=>price!=null&&cost!=null&&cost>0?(price-cost)*100/cost:null;

/* the price board: one row per product (from the data and from the saved list) */
function boardProducts(rows){const m=new Map();rows.forEach(r=>{if(!m.has(r.p))m.set(r.p,{p:r.p,pn:r.pn,u:r.u})});
  PM.cfg.products.forEach(P=>{if(!m.has(P.p))m.set(P.p,{p:P.p,pn:P.pn||P.p,u:P.u||''})});return [...m.values()].sort((a,b)=>String(a.pn).localeCompare(String(b.pn),'ar'))}
function board(rows){const c=PM.cfg,rt=rateNow(),E=PM.edit&&PM.canEdit;const list=boardProducts(rows);
  const inp=(p,k,v)=>`<input type="number" step="any" min="0" inputmode="decimal" data-pm-in="${esc(p)}" data-k="${k}" value="${v==null?'':esc(String(r4(+v)))}" aria-label="${esc(k)}">`;
  const body=list.map(x=>{const e=current(x.p)||{};const cost=e.cost==null?null:+e.cost,loc=e.loc==null?null:+e.loc,fxp=e.fx==null?null:+e.fx;const fxL=fxp!=null&&rt?fxp*rt:null;
    const mL=mOf(loc,cost),mF=mOf(fxL,cost);const st=statusOf(mL==null?mF:(mF==null?mL:Math.min(mL,mF)));const perL=x.u==='م³';
    const sub=v=>perL&&v!=null?`<span class="sub">${p4(v/1000)} ${esc(tr('للتر','/ litre'))}</span>`:'';
    return `<tr data-pm-row="${esc(x.p)}"><td class="t-strong">${esc(x.pn)} <small class="muted num">${esc(x.p)}</small><span class="sub">${esc(tr('السعر لكل','price per'))} ${esc(uLbl(x.u)||'—')}${e.from?' · '+esc(tr('ساري من','since'))+' '+fdi(e.from):''}</span></td>
     <td class="n">${E?inp(x.p,'cost',cost):`${p4(cost)}${sub(cost)}`}</td>
     <td class="n">${E?inp(x.p,'loc',loc):`${p4(loc)}${sub(loc)}`}</td>
     <td class="n ${mL!=null&&mL<0?'pm-neg':''}">${loc!=null&&cost!=null?`${p4(loc-cost)}<span class="sub">${pc(mL)} · ${esc(tr('إضافة','markup'))} ${pc(kOf(loc,cost))}</span>`:'—'}</td>
     <td class="n">${E?inp(x.p,'fx',fxp):`${p4(fxp)} <small class="muted">${esc(c.cur)}</small>${fxL!=null?`<span class="sub">= ${p4(fxL)} ${esc(LYD())}</span>`:''}`}</td>
     <td class="n ${mF!=null&&mF<0?'pm-neg':''}">${fxL!=null&&cost!=null?`${p4(fxL-cost)}<span class="sub">${pc(mF)} · ${esc(tr('إضافة','markup'))} ${pc(kOf(fxL,cost))}</span>`:'—'}</td>
     <td><span class="pill ${st[0]}">${esc(st[1])}</span></td>
     <td>${hist(x.p).length?`<button type="button" class="btn sm ghost" data-pm-hist="${esc(x.p)}">${esc(tr('السجل','History'))} <span class="num">${hist(x.p).length}</span></button>`:''}</td></tr>`}).join('');
  const set=E?`<div class="pm-set">
     <label>${esc(tr('ساري من تاريخ','Effective from'))}<input type="date" id="pm-eff" value="${today()}"></label>
     <label>${esc(tr('سعر الصرف (د.ل لكل وحدة)','Exchange rate (LYD per unit)'))}<input type="number" step="any" min="0" id="pm-rate" value="${rt==null?'':esc(String(rt))}"></label>
     <label>${esc(tr('العملة الأجنبية','Foreign currency'))}<select id="pm-cur">${CURS.map(k=>`<option value="${k[0]}" ${c.cur===k[0]?'selected':''}>${k[0]} · ${esc(tr(k[1],k[2]))}</option>`).join('')}</select></label>
     <label>${esc(tr('الهامش المستهدف %','Target margin %'))}<input type="number" step="any" min="0" max="99" id="pm-target" value="${esc(String(c.target))}"></label>
     <label>${esc(tr('الحد الأدنى للهامش %','Minimum margin %'))}<input type="number" step="any" min="0" max="99" id="pm-min" value="${esc(String(c.min))}"></label>
     <label>${esc(tr('سبب التغيير','Reason for the change'))}<input type="text" id="pm-why" maxlength="200" placeholder="${esc(tr('مثال: قرار وزارة الاقتصاد رقم …','e.g. Ministry decision no. …'))}"></label></div>`
   :`<p class="sm-note" style="margin:0 0 8px">${esc(tr(`سعر الصرف ${rt==null?'غير محدد':p4(rt)+' د.ل لكل '+c.cur} · المستهدف ${c.target}% · الحد الأدنى ${c.min}% · الأسعار لكل متر مكعب للوقود (و«للتر» تحتها للمقارنة).`,`Exchange rate ${rt==null?'not set':p4(rt)+' LYD per '+c.cur} · target ${c.target}% · minimum ${c.min}% · fuel prices per cubic metre (per litre shown below).`))}</p>`;
  return `<section class="panel an-wide" id="pm-board"><h3 class="section-title">${esc(tr('أسعار المنتجات والهوامش','Product prices & margins'))} <small>${esc(tr('السعر المحلي بالدينار للمستهلكين، وسعر الأجانب بالعملة الأجنبية','local price in LYD for consumers, foreign price in the foreign currency'))}</small>
    ${PM.canEdit?(E?`<span style="margin-inline-start:auto;display:inline-flex;gap:6px"><button type="button" class="btn sm primary" data-pm="save-prices">${esc(tr('حفظ الأسعار','Save prices'))}</button><button type="button" class="btn sm ghost" data-pm="cancel-edit">${esc(tr('إلغاء','Cancel'))}</button></span>`:`<button type="button" class="btn sm" data-pm="edit" style="margin-inline-start:auto">${esc(tr('تعديل الأسعار','Edit prices'))}</button>`):''}</h3>
    ${set}
    ${list.length?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('تكلفة الوحدة د.ل','Unit cost LYD'))}</th><th class="nosort n">${esc(tr('السعر المحلي د.ل','Local price LYD'))}</th><th class="nosort n">${esc(tr('الهامش المحلي للوحدة','Local margin / unit'))}</th><th class="nosort n">${esc(tr('سعر الأجانب','Foreign price'))} (${esc(c.cur)})</th><th class="nosort n">${esc(tr('هامش الأجانب للوحدة','Foreign margin / unit'))}</th><th class="nosort">${esc(tr('الحالة','Status'))}</th><th class="nosort"></th></tr></thead><tbody>${body}</tbody></table></div>`
     :`<div class="empty" style="padding:20px">${esc(tr('لا توجد منتجات بعد. استورد تقرير تحليل المبيعات أو قائمة أسعار.','No products yet. Import the sales analysis report or a price list.'))}</div>`}
    <p id="pm-bmsg" style="font-size:12.5px;margin:6px 0 0"></p></section>`}

function tools(rows,A){const u=UI(),list=boardProducts(rows);const s=u.sim;if(!s.p&&list.length)s.p=(A.prods[0]||list[0]).p;
  return `<section class="panel an-wide pm-tools-wrap"><h3 class="section-title">${esc(tr('أدوات التسعير','Pricing tools'))} <small>${esc(tr('تُحسب فورًا ولا تغيّر الأسعار المحفوظة','instant, does not change the saved prices'))}</small></h3>
   <div class="pm-tools">
    <div class="pm-tool"><h4 style="margin:0">${esc(tr('محاكاة تغيير السعر','Price change simulator'))}</h4>
     <div class="row"><label>${esc(tr('المنتج','Product'))}<select id="pm-sim-p">${list.map(x=>`<option value="${esc(x.p)}" ${s.p===x.p?'selected':''}>${esc(x.pn)}</option>`).join('')}</select></label>
      <label>${esc(tr('قائمة الأسعار','Price list'))}<select id="pm-sim-seg"><option value="loc" ${s.seg==='loc'?'selected':''}>${esc(tr('المحلي (د.ل)','Local (LYD)'))}</option><option value="fx" ${s.seg==='fx'?'selected':''}>${esc(tr('الأجانب','Foreign'))} (${esc(PM.cfg.cur)})</option></select></label>
      <label>${esc(tr('التغيير %','Change %'))}<input type="number" id="pm-sim-chg" step="0.5" min="-50" max="100" value="${esc(String(s.chg))}"></label></div>
     <input type="range" id="pm-sim-rng" min="-30" max="30" step="0.5" value="${esc(String(Math.max(-30,Math.min(30,s.chg))))}" aria-label="${esc(tr('نسبة تغيير السعر','Price change percent'))}">
     <div class="pm-out" id="pm-sim-out" aria-live="polite"></div></div>
    <div class="pm-tool"><h4 style="margin:0">${esc(tr('حاسبة التسعير','Pricing calculator'))}</h4>
     <div class="row"><label>${esc(tr('التكلفة للوحدة د.ل','Unit cost LYD'))}<input type="number" id="pm-c-cost" step="any" min="0" value="${esc(u.calc.cost)}"></label>
      <label>${esc(tr('الهامش المطلوب %','Target margin %'))}<input type="number" id="pm-c-m" step="any" min="0" max="99" value="${esc(u.calc.m===''?String(PM.cfg.target):u.calc.m)}"></label>
      <label>${esc(tr('أو سعر بيع للاختبار د.ل','Or a test price LYD'))}<input type="number" id="pm-c-price" step="any" min="0" value="${esc(u.calc.price)}"></label></div>
     <div class="pm-out" id="pm-calc-out" aria-live="polite"></div>
     <p class="sm-note" style="margin:0">${esc(tr('السعر لهامش مستهدف = التكلفة ÷ (1 − الهامش). هامش 20% يساوي إضافة 25% على التكلفة.','Price for a target margin = cost ÷ (1 − margin). A 20% margin equals a 25% markup.'))}</p></div>
   </div></section>`}
function simOut(){const u=UI(),s=u.sim,el=document.getElementById('pm-sim-out');if(!el)return;const rows=rowsAll(),R=range(rows);if(!R){el.innerHTML='';return}
  const cur=pick(rows,R.from,R.to).filter(r=>r.p===s.p&&(s.seg==='fx')===isFx(r.c));const e=current(s.p)||{};const rt=rateNow();
  const cost=e.cost==null?null:+e.cost;let price=s.seg==='fx'?(e.fx==null||!rt?null:+e.fx*rt):(e.loc==null?null:+e.loc);
  if(cost==null||price==null){el.innerHTML=`<div style="grid-column:1/-1"><span>${esc(tr('حدّد التكلفة والسعر لهذا المنتج في جدول الأسعار أولًا.','Set this product’s cost and price in the price table first.'))}</span></div>`;return}
  const q=cur.reduce((a,r)=>a+r.q,0),np=price*(1+(+s.chg||0)/100),gp0=(price-cost)*q,gp1=(np-cost)*q,m0=mOf(price,cost),m1=mOf(np,cost),be=np>cost?gp0/(np-cost):null;
  const u0=(cur[0]&&cur[0].u)||(boardProducts(rows).find(x=>x.p===s.p)||{}).u||'';
  el.innerHTML=[[tr('السعر الحالي','Current price'),p4(s.seg==='fx'?+e.fx:price)+' '+(s.seg==='fx'?PM.cfg.cur:LYD())],
    [tr('السعر الجديد','New price'),p4(s.seg==='fx'?+e.fx*(1+(+s.chg||0)/100):np)+' '+(s.seg==='fx'?PM.cfg.cur:LYD())],
    [tr('الهامش الحالي ← الجديد','Margin now → new'),pc(m0)+' ← '+pc(m1),m1>=m0?'good':'bad'],
    [tr('ربح الفترة الحالي د.ل','Period profit now LYD'),q2(gp0)],[tr('ربح الفترة بالسعر الجديد د.ل','Period profit at new price LYD'),q2(gp1),gp1>=gp0?'good':'bad'],
    [tr('الفرق د.ل','Difference LYD'),(gp1-gp0>=0?'+':'')+q2(gp1-gp0),gp1>=gp0?'good':'bad'],
    [tr(`الكمية في الفترة (${uLbl(u0)})`,`Period quantity (${uLbl(u0)})`),q3(q)],
    [tr('الكمية اللازمة للحفاظ على الربح','Volume that keeps the profit'),be==null?'—':q3(be)+(q?` (${pc((be-q)*100/q)})`:'')]].map(x=>`<div class="${x[2]||''}"><b>${esc(String(x[1]))}</b><span>${esc(x[0])}</span></div>`).join('')}
function calcOut(){const u=UI(),el=document.getElementById('pm-calc-out');if(!el)return;const cost=toNum(u.calc.cost),m=toNum(u.calc.m===''?PM.cfg.target:u.calc.m),pr=toNum(u.calc.price),rt=rateNow();
  if(cost==null){el.innerHTML=`<div style="grid-column:1/-1"><span>${esc(tr('أدخل التكلفة للوحدة.','Enter the unit cost.'))}</span></div>`;return}
  const out=[];if(m!=null&&m<100){const p=cost/(1-m/100);out.push([tr('السعر المحلي المطلوب د.ل','Required local price LYD'),p4(p)],[tr('ما يعادله بالإضافة على التكلفة','Same as a markup of'),pc(kOf(p,cost))]);if(rt)out.push([tr('سعر الأجانب المطلوب','Required foreign price')+' '+PM.cfg.cur,p4(p/rt)])}
  if(pr!=null&&pr>0){const mm=mOf(pr,cost);out.push([tr('هامش سعر الاختبار','Margin of the test price'),pc(mm),mm<0?'bad':'good'],[tr('الإضافة على التكلفة','Markup'),pc(kOf(pr,cost))],[tr('الربح للوحدة د.ل','Profit per unit LYD'),p4(pr-cost),pr<cost?'bad':'good'])}
  el.innerHTML=out.map(x=>`<div class="${x[2]||''}"><b>${esc(String(x[1]))}</b><span>${esc(x[0])}</span></div>`).join('')}

/* ---------- the page ---------- */
function viewMargin(){
  if(!PM.loaded){load().then(()=>{if(S.view===VIEW)render()});return `<div class="pm-page">`+heroHtml(esc(tr('جارٍ تحميل البيانات…','Loading data…')),null)+'<div class="panel muted">'+esc(tr('جارٍ التحميل…','Loading…'))+'</div></div>'}
  FXSET=new Set(PM.cfg.foreign.map(String));
  const all=rowsAll();
  const T=window.__pmx,tab=T?(UI().tab||'full'):'prices';
  const head=`<div class="page-head" style="margin-bottom:8px"><div><h2 style="margin:0">${esc(tr('هامش الربح','Profit margin'))}</h2></div>${toolbar()}</div>`+(T&&PM.err!=='denied'?`<div class="px-tabs" role="tablist" aria-label="${esc(tr('طريقة العرض','View'))}">${[['full',tr('الربحية الشاملة (بيانات البريقة)','Full profitability (Brega data)')],['prices',tr('الأسعار والتحليل السريع','Prices & quick analysis')]].map(([k,l])=>`<button type="button" role="tab" data-pm-tab="${k}" aria-selected="${tab===k}">${esc(l)}</button>`).join('')}</div>`:'');
  /* no permission to read the saved data: the full-profitability view still opens, so a file can be imported and analysed on this device without saving (as in the debts dashboard) */
  if(PM.err==='denied'&&T)return `<div class="pm-page">`+head+T.view()+'</div>';
  if(PM.err==='denied')return `<div class="pm-page">`+head+heroHtml(esc(tr('الربحية حسب المنتج والكمية والفترة، بسعرين: محلي للمستهلكين وبالعملة الأجنبية للأجانب.','Profitability by product, quantity and period, with two price lists.')),null)+`<div class="panel sm-empty">${I.pm.replace('<svg','<svg style="width:46px;height:46px;color:var(--accent)"')}<h3>${esc(tr('هذه اللوحة لإدارة الشؤون التجارية والتحصيلات','This dashboard belongs to Commercial Affairs & Collections'))}</h3><p>${esc(tr('الأسعار والتكاليف بيانات مالية يطّلع عليها الإدارة والإدارة العامة ومجلس الإدارة ومدير النظام فقط.','Prices and costs are financial data for the department, the General Administration, the board and the administrator only.'))}</p></div></div>`;
  if(T&&tab==='full')return `<div class="pm-page">`+head+T.view()+'</div>';
  if(!all.length)return `<div class="pm-page">`+head+heroHtml(esc(tr('حوّل كميات المبيعات إلى ربحية: الإيرادات والتكلفة وإجمالي الربح وهامشه لكل منتج وكمية وفترة، بسعرين — محلي بالدينار للمستهلكين وبالعملة الأجنبية للأجانب.','Turn sales quantities into profitability: revenue, cost, gross profit and margin by product, quantity and period, with two price lists — local in LYD and foreign currency for foreign customers.')),null)+
    `<div class="panel sm-empty">${I.pm.replace('<svg','<svg style="width:46px;height:46px;color:var(--accent)"')}<h3>${esc(tr('لا توجد كميات مبيعات بعد','No sales quantities yet'))}</h3>
     <p>${esc(tr('تُقرأ الكميات من تقرير «تحليل المبيعات» (Sales_Report) في مركز تقارير الإقفال. استورده هنا أو من لوحة تحليل المبيعات، ثم حدّد تكلفة كل منتج وسعريه. يمكنك أيضًا استيراد قائمة أسعار Excel (رمز المنتج، التكلفة، السعر المحلي، سعر الأجانب).','Quantities come from the “Sales analysis” report (Sales_Report). Import it here or in the sales analysis dashboard, then set each product’s cost and two prices. You can also import an Excel price list (product code, cost, local price, foreign price).'))}</p>
     <button class="btn primary" data-pm="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button></div>`+board(all)+'</div>';
  const u=UI(),R=range(all),cur=pick(all,R.from,R.to),A=agg(cur);
  const len=monthsBetween(R.from,R.to),pFrom=addMonths(R.from,-len),pTo=addMonths(R.from,-1),prevRows=pick(all,pFrom,pTo),B=prevRows.length?agg(prevRows):null;
  const dGp=B&&B.gp?((A.gp-B.gp)*100/Math.abs(B.gp)):null;
  const deltaHtml=dGp==null?'':`<span class="sm-delta">${dGp>=0?'▲':'▼'} ${Math.abs(Math.round(dGp*10)/10)}% ${esc(tr('عن الفترة السابقة','vs previous period'))}</span>`;
  const units=[...A.qU.entries()].sort((a,b)=>(b[0]==='م³')-(a[0]==='م³')||b[1]-a[1]);
  const unitsHtml=`<div class="pm-units">${units.map(([k,v])=>`<span>${esc(tr('الكمية','Quantity'))} <b>${k==='م³'?q3(v):qs(v)}</b> ${esc(uLbl(k)||'—')}</span>`).join('')}</div>`;
  const m3=A.qU.get('م³')||0;let m3gp=0;cur.forEach(r=>{if(r.u==='م³'){const x=line(r);if(x.ok)m3gp+=x.gp}});
  const prodsAll=[...new Map(all.map(r=>[r.p,r.pn])).entries()].sort((a,b)=>String(a[1]).localeCompare(String(b[1]),'ar'));
  const rgns=uniq(all.map(r=>r.rgn)),whs=uniq(all.map(r=>r.wn));
  const c=PM.cfg;const below=boardProducts(all).filter(x=>{const e=current(x.p)||{};const rt=rateNow();const mL=mOf(e.loc==null?null:+e.loc,e.cost==null?null:+e.cost),mF=mOf(e.fx==null||!rt?null:+e.fx*rt,e.cost==null?null:+e.cost);return (mL!=null&&mL<+c.min)||(mF!=null&&mF<+c.min)});
  const missing=[...A.miss.values()];
  /* reading */
  const ins=[];const tp=A.prods.find(p=>p.rev>0);
  if(dGp!=null)ins.push([dGp>=0?'o':'w',dGp>=0?'▲':'▼',tr(`إجمالي الربح <b>${qs(A.gp)}</b> د.ل بهامش <b>${pc(A.m)}</b>، ${dGp>=0?'بزيادة':'بانخفاض'} <b>${Math.abs(Math.round(dGp*10)/10)}%</b> عن الفترة السابقة المماثلة (${qs(B.gp)} د.ل بهامش ${pc(B.m)}).`,`Gross profit <b>${qs(A.gp)}</b> LYD at a <b>${pc(A.m)}</b> margin, ${dGp>=0?'up':'down'} <b>${Math.abs(Math.round(dGp*10)/10)}%</b> on the same-length previous period (${qs(B.gp)} LYD at ${pc(B.m)}).`)]);
  else ins.push(['','i',tr(`إجمالي الربح <b>${qs(A.gp)}</b> د.ل من إيرادات <b>${qs(A.rev)}</b> د.ل، أي هامش <b>${pc(A.m)}</b> وإضافة على التكلفة <b>${pc(A.mk)}</b>. استورد الفترة السابقة لتظهر المقارنة.`,`Gross profit <b>${qs(A.gp)}</b> LYD on revenue of <b>${qs(A.rev)}</b> LYD: a <b>${pc(A.m)}</b> margin and a <b>${pc(A.mk)}</b> markup. Import the previous period to see the comparison.`)]);
  if(m3)ins.push(['','م³',tr(`الربح لكل متر مكعب من الوقود <b>${q2(m3?m3gp/m3:0)}</b> د.ل (أي ${p4(m3?m3gp/m3/1000:0)} د.ل للتر) على <b>${q3(m3)}</b> م³.`,`Profit per cubic metre of fuel <b>${q2(m3?m3gp/m3:0)}</b> LYD (${p4(m3?m3gp/m3/1000:0)} LYD per litre) over <b>${q3(m3)}</b> m³.`)]);
  if(tp&&A.gp>0)ins.push(['','1',tr(`المنتج الأكثر ربحًا: <b>${bd(tp.pn)}</b> بـ <b>${pct(tp.gp,A.gp)}%</b> من إجمالي الربح، وهامشه ${pc(tp.rev?tp.gp*100/tp.rev:null)}.`,`Most profitable product: <b>${bd(tp.pn)}</b>, <b>${pct(tp.gp,A.gp)}%</b> of gross profit at a ${pc(tp.rev?tp.gp*100/tp.rev:null)} margin.`)]);
  const lowP=A.prods.filter(p=>p.rev>0).sort((a,b)=>a.gp/a.rev-b.gp/b.rev)[0];if(lowP&&lowP!==tp)ins.push([lowP.gp<0?'c':'w','↓',tr(`أضعف هامش: <b>${bd(lowP.pn)}</b> (${pc(lowP.gp*100/lowP.rev)}).`,`Weakest margin: <b>${bd(lowP.pn)}</b> (${pc(lowP.gp*100/lowP.rev)}).`)]);
  if(A.fx.rev>0){const sq=pct(A.fx.q,A.fx.q+A.loc.q),sg=pct(A.fx.gp,A.gp);const mf=A.fx.gp*100/A.fx.rev,ml=A.loc.rev?A.loc.gp*100/A.loc.rev:null;
    ins.push([sg>sq?'o':'','$',tr(`مبيعات الأجانب <b>${q2(A.fxRev)}</b> ${esc(c.cur)} (${qs(A.fx.rev)} د.ل): <b>${sq}%</b> من الكمية و<b>${sg}%</b> من الربح، بهامش ${pc(mf)} مقابل ${pc(ml)} للمحلي.`,`Foreign sales <b>${q2(A.fxRev)}</b> ${esc(c.cur)} (${qs(A.fx.rev)} LYD): <b>${sq}%</b> of quantity and <b>${sg}%</b> of profit, at a ${pc(mf)} margin vs ${pc(ml)} local.`)])}
  else if(!c.foreign.length)ins.push(['','$',tr('لم يُصنَّف أي زبون أجنبيًا بعد، فكل الكميات مسعّرة بالسعر المحلي. صنّف الزبائن الأجانب من جدول الزبائن أدناه.','No customer is marked as foreign yet, so every quantity is priced at the local price. Mark foreign customers in the customer table below.')]);
  if(A.m!=null&&A.m<+c.target&&A.rev>0){const need=A.cogs/(1-c.target/100);ins.push([A.m<+c.min?'c':'w','!',tr(`الهامش أقل من المستهدف (${c.target}%). للوصول إليه بالكميات نفسها يلزم رفع متوسط الأسعار <b>${pc((need/A.rev-1)*100)}</b> أو خفض التكلفة <b>${pc((1-A.rev*(1-c.target/100)/A.cogs)*100)}</b>.`,`Margin is below the ${c.target}% target. At the same volumes, prices must rise <b>${pc((need/A.rev-1)*100)}</b> on average, or cost fall <b>${pc((1-A.rev*(1-c.target/100)/A.cogs)*100)}</b>.`)])}
  else if(A.m!=null)ins.push(['o','✓',tr(`الهامش ضمن المستهدف (${c.target}%).`,`Margin is on the ${c.target}% target.`)]);
  if(below.length)ins.push(['c','−',tr(`<b>${below.length}</b> منتج بسعر حالي يعطي هامشًا تحت الحد الأدنى (${c.min}%): ${below.slice(0,4).map(x=>bd(x.pn)).join('، ')}.`,`<b>${below.length}</b> products have a current price below the ${c.min}% minimum margin: ${below.slice(0,4).map(x=>bd(x.pn)).join(', ')}.`)]);
  if(missing.length)ins.push(['w','?',tr(`<b>${missing.length}</b> منتج بلا تكلفة أو سعر (${missing.slice(0,3).map(x=>bd(x.pn)).join('، ')})؛ كمياته مستبعدة من الربح حتى تحدد أسعاره.`,`<b>${missing.length}</b> products have no cost or price (${missing.slice(0,3).map(x=>bd(x.pn)).join(', ')}); their quantities are left out until priced.`)]);
  const custs=A.custs.filter(x=>x.rev>0).sort((a,b)=>b.gp-a.gp);const t5=custs.slice(0,5).reduce((s,x)=>s+x.gp,0);if(custs.length>5&&A.gp>0){const sh=pct(t5,A.gp);ins.push([sh>=50?'w':'','5',tr(`أعلى 5 زبائن يحققون <b>${sh}%</b> من الربح، أكبرهم <b>${bd(custs[0].cn)}</b> (${qs(custs[0].gp)} د.ل).`,`The top 5 customers bring <b>${sh}%</b> of profit; the largest is <b>${bd(custs[0].cn)}</b> (${qs(custs[0].gp)} LYD).`)])}
  /* trend */
  const pts=[];for(let k=R.from;k<=R.to;k=addMonths(k,1)){const o=A.mon.get(k.slice(0,7))||{rev:0,gp:0,fxGp:0,locGp:0};pts.push({l:mName(+k.slice(5,7)-1).slice(0,3)+' '+k.slice(2,4),y:o.gp,tip:`${mName(+k.slice(5,7)-1)} ${k.slice(0,4)} · ${tr('ربح','Profit')} ${q2(o.gp)} ${LYD()} · ${tr('هامش','margin')} ${pc(o.rev?o.gp*100/o.rev:null)} · ${tr('محلي','local')} ${q2(o.locGp)} · ${tr('أجانب','foreign')} ${q2(o.fxGp)}`})}
  /* customer table */
  const q=norm(u.q||'');let tbl=A.custs.filter(x=>!q||norm(x.cn+' '+x.c).includes(q));const key={gp:x=>x.gp,rev:x=>x.rev,m:x=>x.rev?x.gp/x.rev:-1e9,cn:x=>x.cn,n:x=>x.n}[u.sort]||(x=>x.gp);
  tbl.sort((a,b)=>{const x=key(a),y=key(b);const r=typeof x==='string'?x.localeCompare(y,'ar'):x-y;return u.dir==='asc'?r:-r});const tblN=tbl.length;tbl=tbl.slice(0,u.lim);
  const th=(k,l,cls='')=>`<th class="${cls}" data-pm-sort="${k}" ${u.sort===k?`aria-sort="${u.dir==='asc'?'ascending':'descending'}"`:''}>${esc(l)}</th>`;
  const qtyTxt=mq=>[...mq.entries()].sort((a,b)=>(b[0]==='م³')-(a[0]==='م³')).map(([k,v])=>`<bdi>${k==='م³'?q3(v):qs(v)} ${esc(uLbl(k))}</bdi>`).join('<br>');
  const sub=esc(tr(`الفترة ${fdi(R.from).slice(0,7)} – ${fdi(R.to).slice(0,7)} · ${fn(A.n)} حركة · ${A.custs.length} زبون · البيانات المتاحة ${fdi(R.min).slice(0,7)} – ${fdi(R.max).slice(0,7)}`,`Period ${fdi(R.from).slice(0,7)} – ${fdi(R.to).slice(0,7)} · ${fn(A.n)} lines · ${A.custs.length} customers · data ${fdi(R.min).slice(0,7)} – ${fdi(R.max).slice(0,7)}`));
  const pr=A.prods.filter(p=>p.rev>0);
  return `<div class="pm-page">`+head+(PM.local?`<div class="banner info">${esc(tr(`عرض مؤقت للملف «${PM.local.file}» على هذا الجهاز فقط، ولم يُحفظ في المنظومة.`,`Temporary view of “${PM.local.file}” on this device only; not saved.`))}</div>`:'')+
  heroHtml(sub,A,deltaHtml,unitsHtml)+
  `<div class="sm-bar">
    <div class="an-f">${esc(tr('الفترة','Period'))}<div class="sm-presets">${[['all',tr('كل البيانات','All data')],['12',tr('آخر 12 شهرًا','Last 12 months')],['3',tr('آخر 3 أشهر','Last 3 months')],['month',tr('آخر شهر','Latest month')],['custom',tr('مخصّص','Custom')]].map(([k,l])=>`<button type="button" data-pm-pre="${k}" aria-pressed="${u.pre===k}">${esc(l)}</button>`).join('')}</div></div>
    ${u.pre==='custom'?`<label class="an-f">${esc(tr('من','From'))}<input type="month" id="pm-from" value="${esc(R.from.slice(0,7))}"></label><label class="an-f">${esc(tr('إلى','To'))}<input type="month" id="pm-to" value="${esc(R.to.slice(0,7))}"></label>`:''}
    <label class="an-f">${esc(tr('المنتج','Product'))}<select id="pm-prod"><option value="">${esc(tr('كل المنتجات','All products'))}</option>${prodsAll.map(([p,n])=>`<option value="${esc(p)}" ${u.prod===p?'selected':''}>${esc(n)}</option>`).join('')}</select></label>
    <label class="an-f">${esc(tr('فئة الزبون','Customer type'))}<select id="pm-seg"><option value="">${esc(tr('الكل','All'))}</option><option value="loc" ${u.seg==='loc'?'selected':''}>${esc(tr('محلي (د.ل)','Local (LYD)'))}</option><option value="fx" ${u.seg==='fx'?'selected':''}>${esc(tr('أجانب','Foreign'))} (${esc(c.cur)})</option></select></label>
    ${rgns.length>1?`<label class="an-f">${esc(tr('المنطقة','Region'))}<select id="pm-rgn"><option value="">${esc(tr('كل المناطق','All regions'))}</option>${rgns.map(x=>`<option ${u.rgn===x?'selected':''}>${esc(x)}</option>`).join('')}</select></label>`:''}
    ${whs.length>1?`<label class="an-f">${esc(tr('المستودع','Warehouse'))}<select id="pm-wn"><option value="">${esc(tr('كل المستودعات','All warehouses'))}</option>${whs.map(x=>`<option ${u.wn===x?'selected':''}>${esc(x)}</option>`).join('')}</select></label>`:''}
    ${u.prod||u.seg||u.rgn||u.wn||u.pre!=='all'?`<button class="btn sm ghost" data-pm="clear">${esc(tr('مسح التصفية','Clear filters'))}</button>`:''}
  </div>
  ${missing.length?`<div class="banner warn">${esc(tr(`${missing.length} منتج بلا تكلفة أو سعر، وكمياته مستبعدة من حساب الربح:`,`${missing.length} products have no cost or price and are left out of profit:`))} ${missing.slice(0,6).map(x=>bd(x.pn)).join(tr('، ',', '))}${PM.canEdit?` <button type="button" class="btn sm" data-pm="edit">${esc(tr('تحديد الأسعار','Set prices'))}</button>`:''}</div>`:''}
  ${!A.n?`<div class="panel empty"><b>${esc(tr('لا توجد حركات','No lines'))}</b>${esc(tr('ضمن التصفية الحالية.','for the current filter.'))}</div>`:`
  <div class="kpis an-kpis">
   ${anKpi(tr('الإيرادات','Revenue'),qs(A.rev),esc(tr(`د.ل مكافئ · ${fn(A.n)} حركة`,`LYD equivalent · ${fn(A.n)} lines`)))}
   ${anKpi(tr('تكلفة المبيعات','Cost of sales'),qs(A.cogs),esc(LYD()))}
   ${anKpi(tr('إجمالي الربح','Gross profit'),qs(A.gp),esc(B?tr(`السابقة ${qs(B.gp)} د.ل`,`previous ${qs(B.gp)} LYD`):LYD()),A.gp<0?'crit':'')}
   ${anKpi(tr('هامش الربح','Gross margin'),pc(A.m),esc(tr(`المستهدف ${c.target}% · الحد الأدنى ${c.min}%`,`target ${c.target}% · minimum ${c.min}%`)),A.m==null?'':A.m<+c.min?'crit':A.m<+c.target?'warn':'ok')}
   ${anKpi(tr('الإضافة على التكلفة','Markup on cost'),pc(A.mk),esc(tr('الربح ÷ التكلفة','profit ÷ cost')))}
   ${anKpi(tr('الربح لكل م³','Profit per m³'),m3?q2(m3gp/m3):'—',esc(m3?tr(`د.ل · ${p4(m3gp/m3/1000)} للتر`,`LYD · ${p4(m3gp/m3/1000)} per litre`):tr('لا توجد كميات وقود','no fuel quantities')))}
   ${anKpi(tr('مبيعات الأجانب','Foreign sales'),q2(A.fxRev),esc(`${c.cur} · ${pct(A.fx.gp,A.gp)}% ${tr('من الربح','of profit')}`))}
   ${anKpi(tr('منتجات تحت الحد الأدنى','Products below minimum'),fn(below.length),esc(missing.length?tr(`${missing.length} بلا سعر`,`${missing.length} without a price`):tr('كل المنتجات مسعّرة','every product priced')),below.length?'crit':missing.length?'warn':'ok')}
  </div>
  <div class="an-grid sm-root pm-root">
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('قراءة تحليلية','Reading'))} <small>${esc(tr('تُكتب تلقائيًا من البيانات المصفّاة','written from the filtered data'))}</small></h3><ul class="sm-ins">${ins.map(x=>`<li><span class="ic ${x[0]}">${esc(x[1])}</span><span>${x[2]}</span></li>`).join('')}</ul></section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('إجمالي الربح الشهري','Monthly gross profit'))} <small>${esc(tr('د.ل · مرّر على العمود للهامش والتفصيل','LYD · hover a column for margin and split'))}</small></h3>${colChart(pts,tr('إجمالي الربح حسب الشهر','Gross profit by month'))}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('الربح حسب المنتج','Profit by product'))} <small>${esc(LYD())}</small></h3>${bars(A.prods.filter(p=>p.rev).slice(0,10).map(p=>({l:p.pn,v:p.gp,tip:`${p.pn}: ${q2(p.gp)} ${LYD()} · ${tr('هامش','margin')} ${pc(p.rev?p.gp*100/p.rev:null)}`})),{fmt:qs})}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('هامش الربح % حسب المنتج','Margin % by product'))} <small>${esc(tr('الربح ÷ الإيرادات','profit ÷ revenue'))}</small></h3>${bars(pr.map(p=>({l:p.pn,v:Math.round(p.gp*1000/p.rev)/10,tip:`${p.pn}: ${pc(p.gp*100/p.rev)} · ${tr('إضافة','markup')} ${pc(p.cogs?p.gp*100/p.cogs:null)}`})).sort((a,b)=>b.v-a.v),{fmt:v=>pc(v)})}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('المحلي والأجانب','Local and foreign'))} <small>${esc(tr('حصة كل قائمة أسعار من الربح','each price list’s share of profit'))}</small></h3>
    ${stack([{l:tr('محلي (د.ل)','Local (LYD)'),v:A.loc.gp,k:'sm-k1'},{l:tr('أجانب','Foreign')+' ('+c.cur+')',v:A.fx.gp,k:'sm-k2'}])}
    <div class="sm-mini" style="margin-top:10px"><div><b class="num">${pc(A.loc.rev?A.loc.gp*100/A.loc.rev:null)}</b><span>${esc(tr('هامش المحلي','local margin'))}</span></div><div><b class="num">${pc(A.fx.rev?A.fx.gp*100/A.fx.rev:null)}</b><span>${esc(tr('هامش الأجانب','foreign margin'))}</span></div><div><b class="num">${fn(A.custs.filter(x=>x.fx).length)}</b><span>${esc(tr('زبون أجنبي','foreign customers'))}</span></div></div></section>
   <section class="panel"><h3 class="section-title">${esc(tr('الربح للوحدة حسب المنتج','Profit per unit by product'))} <small>${esc(tr('د.ل لكل م³ للوقود','LYD per m³ for fuels'))}</small></h3>${bars(pr.filter(p=>p.q).map(p=>({l:p.pn+' / '+uLbl(p.u),v:p.gp/(p.q-p.miss||1),tip:`${p.pn}: ${q2(p.gp/(p.q-p.miss||1))} ${LYD()} / ${uLbl(p.u)}`})).sort((a,b)=>b.v-a.v),{fmt:q2})}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('المنتجات','Products'))} <small>${esc(tr('الكمية والإيرادات والتكلفة والربح لكل منتج','quantity, revenue, cost and profit per product'))}</small></h3>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th><th class="nosort n">${esc(tr('منها أجانب','Foreign qty'))}</th><th class="nosort n">${esc(tr('الإيرادات د.ل','Revenue LYD'))}</th><th class="nosort n">${esc(tr('التكلفة د.ل','Cost LYD'))}</th><th class="nosort n">${esc(tr('الربح د.ل','Profit LYD'))}</th><th class="nosort n">${esc(tr('الهامش','Margin'))}</th><th class="nosort n">${esc(tr('الربح للوحدة','Profit / unit'))}</th><th class="nosort n">${esc(tr('الحصة','Share'))}</th></tr></thead>
    <tbody>${A.prods.map(p=>{const m=p.rev?p.gp*100/p.rev:null,st=statusOf(m),w=m==null?0:Math.max(0,Math.min(100,m*100/Math.max(1,+c.target*2)));return `<tr><td class="t-strong">${esc(p.pn)} <small class="muted num">${esc(p.p)}</small>${p.miss?`<span class="sub" style="color:var(--warn)">${esc(tr('بلا سعر لـ','unpriced'))} ${q3(p.miss)} ${esc(uLbl(p.u))}</span>`:''}</td><td class="n">${p.u==='م³'?q3(p.q):fn(Math.round(p.q))} <small class="muted">${esc(uLbl(p.u))}</small></td><td class="n">${p.qFx?(p.u==='م³'?q3(p.qFx):fn(Math.round(p.qFx))):'—'}</td><td class="n">${q2(p.rev)}</td><td class="n">${q2(p.cogs)}</td><td class="n ${p.gp<0?'pm-neg':''}"><b>${q2(p.gp)}</b></td><td class="n"><span style="display:inline-flex;gap:6px;align-items:center">${pc(m)}<span class="pm-gauge" title="${esc(st[1])}"><i class="${st[2]}" style="width:${w}%"></i></span></span></td><td class="n">${p.q-p.miss?q2(p.gp/(p.q-p.miss)):'—'}</td><td class="n">${A.gp>0?pct(Math.max(0,p.gp),A.gp)+'%':'—'}</td></tr>`}).join('')}</tbody>
    <tfoot><tr><td class="t-strong">${esc(tr('الإجمالي','Total'))}</td><td class="n">${qtyTxt(A.qU)}</td><td class="n">${A.fx.q?q3(A.fx.q):'—'}</td><td class="n"><b>${q2(A.rev)}</b></td><td class="n"><b>${q2(A.cogs)}</b></td><td class="n"><b>${q2(A.gp)}</b></td><td class="n"><b>${pc(A.m)}</b></td><td class="n"></td><td class="n"></td></tr></tfoot></table></div></section>
   ${board(all)}
   ${tools(all,A)}
   <section class="panel"><h3 class="section-title">${esc(tr('أعلى 10 زبائن ربحًا','Top 10 customers by profit'))} <small>${esc(tr('إجمالي الربح · د.ل','gross profit · LYD'))}</small></h3>${bars(custs.slice(0,10).map(x=>({l:x.cn,v:x.gp,tip:`${x.cn}: ${q2(x.gp)} ${LYD()} · ${tr('هامش','margin')} ${pc(x.rev?x.gp*100/x.rev:null)}${x.fx?' · '+tr('أجنبي','foreign'):''}`})),{fmt:qs})}</section>
   ${A.wn.size>1?`<section class="panel"><h3 class="section-title">${esc(tr('الربح حسب المستودع','Profit by warehouse'))} <small>${esc(LYD())}</small></h3>${bars([...A.wn.values()].sort((a,b)=>b.gp-a.gp).slice(0,10).map(g=>({l:g.k,v:g.gp,tip:`${g.k}: ${q2(g.gp)} ${LYD()} · ${pc(g.rev?g.gp*100/g.rev:null)}`})),{fmt:qs})}</section>`:A.rgn.size>1?`<section class="panel"><h3 class="section-title">${esc(tr('الربح حسب المنطقة','Profit by region'))}</h3>${bars([...A.rgn.values()].sort((a,b)=>b.gp-a.gp).map(g=>({l:g.k,v:g.gp})),{fmt:qs})}</section>`:''}
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('الزبائن','Customers'))} <small>${tblN} ${esc(tr('زبون',' customers'))}${PM.canEdit?' · '+esc(tr('اضغط «محلي/أجنبي» لتغيير قائمة الأسعار التي تُطبَّق على الزبون','press “Local/Foreign” to change the price list applied to the customer')):''}</small></h3>
    <div class="toolbar" style="margin:0 0 8px"><div class="grow od-search">${I.search}<input type="search" id="pm-q" value="${esc(u.q||'')}" placeholder="${esc(tr('بحث باسم الزبون أو رقم الحساب…','Search by customer name or account…'))}" aria-label="${esc(tr('بحث في الزبائن','Search customers'))}"></div></div>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">#</th>${th('cn',tr('الزبون','Customer'))}<th class="nosort">${esc(tr('الفئة','Type'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th>${th('rev',tr('الإيرادات د.ل','Revenue LYD'),'n')}${th('gp',tr('الربح د.ل','Profit LYD'),'n')}${th('m',tr('الهامش','Margin'),'n')}</tr></thead>
    <tbody>${tbl.map((x,i)=>`<tr data-pm-cust="${esc(x.c)}" tabindex="0"><td class="n">${i+1}</td><td class="t-strong">${esc(x.cn)} <small class="muted num">${esc(x.c)}</small></td><td>${PM.canEdit&&!PM.local?`<button type="button" class="pm-seg${x.fx?' fx':''}" data-pm-fx="${esc(x.c)}" aria-pressed="${x.fx}">${esc(x.fx?tr('أجنبي','Foreign'):tr('محلي','Local'))}</button>`:`<span class="pm-seg${x.fx?' fx':''}">${esc(x.fx?tr('أجنبي','Foreign'):tr('محلي','Local'))}</span>`}</td><td class="n">${qtyTxt(x.q)}</td><td class="n">${q2(x.rev)}</td><td class="n ${x.gp<0?'pm-neg':''}"><b>${q2(x.gp)}</b></td><td class="n">${pc(x.rev?x.gp*100/x.rev:null)}</td></tr>`).join('')}</tbody></table></div>
    ${tblN>u.lim?`<div style="text-align:center;margin-top:8px"><button class="btn sm" data-pm="more">${esc(tr(`عرض المزيد (${tblN-u.lim} متبقٍ)`,`Show more (${tblN-u.lim} left)`))}</button></div>`:''}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('التعريفات','Definitions'))} <small>${esc(tr('ليقصد الجميع الشيء نفسه','so everyone means the same thing'))}</small></h3><ul class="pm-def">
    <li>${tr('<b>إجمالي الربح</b> = الإيرادات − تكلفة المبيعات (الكمية × تكلفة الوحدة). هذا ربح إجمالي وليس صافيًا: لا يشمل المصروفات التشغيلية.','<b>Gross profit</b> = revenue − cost of sales (quantity × unit cost). It is gross, not net: operating expenses are not included.')}</li>
    <li>${tr('<b>هامش الربح %</b> = إجمالي الربح ÷ الإيرادات × 100، دائمًا أقل من 100%. <b>الإضافة على التكلفة %</b> = إجمالي الربح ÷ التكلفة × 100. هامش 20% = إضافة 25%.','<b>Gross margin %</b> = gross profit ÷ revenue × 100, always below 100%. <b>Markup %</b> = gross profit ÷ cost × 100. A 20% margin = a 25% markup.')}</li>
    <li>${tr('<b>السعران</b>: المحلي بالدينار للمستهلكين المحليين، وسعر الأجانب بالعملة الأجنبية ويُحوَّل إلى الدينار بسعر الصرف الساري في شهر البيع. الزبون أجنبي إذا صُنِّف كذلك في جدول الزبائن.','<b>Two prices</b>: local in LYD for local consumers; the foreign price is in the foreign currency, converted at the exchange rate in effect in the month of sale. A customer is foreign when marked so in the customer table.')}</li>
    <li>${tr('<b>الكميات بالمتر المكعب</b>: اللتر ÷ 1000، والطن ÷ الكثافة عند 15°م (كثافة نموذجية لكل منتج). الأسعار والتكاليف للوقود لكل م³، وما لا يُحوَّل (الأسطوانات) بوحدته.','<b>Quantities in cubic metres</b>: litres ÷ 1000, tonnes ÷ the density at 15 °C (a typical density per product). Fuel prices and costs are per m³; what cannot be converted (cylinders) stays in its unit.')}</li>
    <li>${tr('<b>تاريخ السريان</b>: كل تغيير سعر يُحفظ بتاريخ سريان، فتُحسب كل فترة بأسعارها ولا يتغير ربح الأشهر الماضية عند تعديل السعر.','<b>Effective date</b>: every price change is saved with an effective date, so each period uses its own prices and past months do not change.')}</li></ul></section>
  </div>`}</div>`}

function after(){const u=UI();if(window.__pmx&&u.tab==='full')window.__pmx.after();
  [['pm-prod','prod'],['pm-seg','seg'],['pm-rgn','rgn'],['pm-wn','wn']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('change',()=>{u[k]=e.value;u.lim=25;render()})});
  [['pm-from','from'],['pm-to','to']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('change',()=>{u[k]=e.value?e.value+'-01':'';render()})});
  const q=document.getElementById('pm-q');if(q)q.addEventListener('input',debounce(()=>{u.q=q.value;u.lim=25;const p=q.selectionStart;render();const n=document.getElementById('pm-q');if(n){n.focus();try{n.setSelectionRange(p,p)}catch(e){}}},250));
  const f=document.getElementById('pm-file');if(f)f.addEventListener('change',()=>{const file=f.files&&f.files[0];f.value='';if(file)importFile(file)});
  const sp=document.getElementById('pm-sim-p'),ss=document.getElementById('pm-sim-seg'),sc=document.getElementById('pm-sim-chg'),sr=document.getElementById('pm-sim-rng');
  if(sp){sp.addEventListener('change',()=>{u.sim.p=sp.value;simOut()});ss.addEventListener('change',()=>{u.sim.seg=ss.value;simOut()});
    sc.addEventListener('input',()=>{u.sim.chg=toNum(sc.value)||0;sr.value=Math.max(-30,Math.min(30,u.sim.chg));simOut()});sr.addEventListener('input',()=>{u.sim.chg=+sr.value;sc.value=sr.value;simOut()});simOut()}
  [['pm-c-cost','cost'],['pm-c-m','m'],['pm-c-price','price']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('input',()=>{u.calc[k]=e.value;calcOut()})});calcOut();
  document.querySelectorAll('[data-pm-in]').forEach(e=>e.addEventListener('input',()=>{const tr0=e.closest('tr');if(tr0)tr0.classList.add('pm-chg')}))}

/* ---------- saving prices from the board ---------- */
function applyPrices(cfg,changes,eff,why){const m=me(),at=new Date().toISOString();const log=[];
  changes.forEach(ch=>{let P=cfg.products.find(x=>x.p===ch.p);if(!P){P={p:ch.p,pn:ch.pn||ch.p,u:ch.u||'',h:[]};cfg.products.push(P)}if(ch.pn)P.pn=ch.pn;if(ch.u)P.u=ch.u;P.h=Array.isArray(P.h)?P.h:[];
    const old=priceAt(ch.p,eff)||{};const e={from:eff,cost:ch.cost===undefined?(old.cost??null):ch.cost,loc:ch.loc===undefined?(old.loc??null):ch.loc,fx:ch.fx===undefined?(old.fx??null):ch.fx,by:String(m.name||m.email||'').slice(0,120),at,why:String(why||'').slice(0,200)};
    const diff=['cost','loc','fx'].filter(k=>(old[k]??null)!==(e[k]??null));if(!diff.length)return;
    P.h=P.h.filter(h=>h.from!==eff);P.h.push(e);P.h.sort((a,b)=>String(a.from).localeCompare(String(b.from)));if(P.h.length>200)P.h=P.h.slice(-200);
    log.push(`${P.pn}: `+diff.map(k=>`${{cost:'التكلفة',loc:'السعر المحلي',fx:'سعر الأجانب'}[k]} ${old[k]??'—'} ← ${e[k]??'—'}`).join('، '))});
  return log}
async function savePrices(btn){const msg=document.getElementById('pm-bmsg');const cfg=JSON.parse(JSON.stringify(PM.cfg));const eff=(document.getElementById('pm-eff')||{}).value||today();const why=(document.getElementById('pm-why')||{}).value||'';
  const changes=new Map();const list=boardProducts(rowsAll());let bad='';
  document.querySelectorAll('[data-pm-in]').forEach(e=>{const p=e.dataset.pmIn,k=e.dataset.k;const v=e.value.trim()===''?null:toNum(e.value);if(e.value.trim()!==''&&(v==null||v<0)){bad=bad||p;return}
    const cur0=current(p)||{};const ov=cur0[k]==null?null:r4(+cur0[k]);if((v==null?null:r4(v))===ov)return;const x=list.find(y=>y.p===p)||{};const ch=changes.get(p)||{p,pn:x.pn,u:x.u};ch[k]=v==null?null:r4(v);changes.set(p,ch)});
  if(bad){msg.style.color='var(--crit)';msg.textContent=tr('قيمة غير صحيحة: أدخل أرقامًا موجبة.','Invalid value: enter positive numbers.');return}
  const log=applyPrices(cfg,[...changes.values()],eff,why);
  const rate=toNum((document.getElementById('pm-rate')||{}).value),curC=(document.getElementById('pm-cur')||{}).value||cfg.cur,tg=toNum((document.getElementById('pm-target')||{}).value),mn=toNum((document.getElementById('pm-min')||{}).value);
  const oldRate=rateAt(eff);if(rate!=null&&rate>0&&rate!==oldRate){cfg.rates=cfg.rates.filter(x=>x.from!==eff);cfg.rates.push({from:eff,rate:r4(rate)});cfg.rates.sort((a,b)=>String(a.from).localeCompare(String(b.from)));log.push(`سعر الصرف ${oldRate??'—'} ← ${r4(rate)} (${curC})`)}
  if(curC!==cfg.cur){log.push(`العملة ${cfg.cur} ← ${curC}`);cfg.cur=curC}
  if(tg!=null&&tg>=0&&tg<100&&tg!==+cfg.target){log.push(`الهامش المستهدف ${cfg.target}% ← ${tg}%`);cfg.target=tg}
  if(mn!=null&&mn>=0&&mn<100&&mn!==+cfg.min){log.push(`الحد الأدنى ${cfg.min}% ← ${mn}%`);cfg.min=mn}
  if(!log.length){PM.edit=false;render();toast(tr('لا تغييرات','No changes'));return}
  btn.disabled=true;msg.style.color='';msg.textContent=tr('جارٍ الحفظ…','Saving…');
  try{await saveCfg(cfg,'تعديل أسعار هامش الربح',`ساري من ${eff}${why?' · '+why:''} · `+log.join(' · '));PM.edit=false;render();toast(tr(`تم حفظ ${log.length} تغيير`,`${log.length} changes saved`))}
  catch(err){btn.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية تعديل الأسعار','You are not allowed to change prices'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}}
async function toggleFx(c){const cfg=JSON.parse(JSON.stringify(PM.cfg));const s=new Set(cfg.foreign.map(String));const was=s.has(c);if(was)s.delete(c);else s.add(c);cfg.foreign=[...s];
  const r=rowsAll().find(x=>x.c===c);
  try{await saveCfg(cfg,'تصنيف زبون في هامش الربح',`${r?r.cn:c} (${c}): ${was?'أجنبي ← محلي':'محلي ← أجنبي'}`);render();toast(was?tr('صار الزبون محليًا','Customer set to local'):tr('صار الزبون أجنبيًا','Customer set to foreign'))}
  catch(err){toast((err&&err.code==='permission-denied')?tr('لا تملك صلاحية التعديل','Not allowed'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||'')))}}

function histDlg(p){const H=hist(p).slice().reverse(),P=prodCfg(p)||{};const c=PM.cfg;
  modal(tr('سجل أسعار ','Price history: ')+(P.pn||p),`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">${esc(tr('ساري من','From'))}</th><th class="nosort n">${esc(tr('التكلفة','Cost'))}</th><th class="nosort n">${esc(tr('المحلي د.ل','Local LYD'))}</th><th class="nosort n">${esc(tr('الأجانب','Foreign'))} ${esc(c.cur)}</th><th class="nosort n">${esc(tr('الهامش المحلي','Local margin'))}</th><th class="nosort">${esc(tr('بواسطة','By'))}</th><th class="nosort">${esc(tr('السبب','Reason'))}</th></tr></thead>
   <tbody>${H.map(h=>`<tr><td class="n">${fdi(h.from)}</td><td class="n">${p4(h.cost)}</td><td class="n">${p4(h.loc)}</td><td class="n">${p4(h.fx)}</td><td class="n">${pc(mOf(h.loc==null?null:+h.loc,h.cost==null?null:+h.cost))}</td><td><small>${esc(h.by||'')}<br>${fdi(h.at)}</small></td><td><small>${esc(h.why||'')}</small></td></tr>`).join('')}</tbody></table></div>`,
   `<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}
function custDetail(c){const all=rowsAll(),R=range(all);if(!R)return;const rows=pick(all,R.from,R.to).filter(r=>r.c===c);if(!rows.length)return;const A=agg(rows),x=A.custs[0];
  modal(`${x.cn} (${x.c})`,`<div class="sm-mini"><div><b class="num">${q2(x.gp)}</b><span>${esc(tr('الربح د.ل','Profit LYD'))}</span></div><div><b class="num">${pc(x.rev?x.gp*100/x.rev:null)}</b><span>${esc(tr('الهامش','Margin'))}</span></div><div><b class="num">${q2(x.rev)}</b><span>${esc(tr('الإيرادات د.ل','Revenue LYD'))}</span></div><div><b>${esc(x.fx?tr('أجنبي','Foreign'):tr('محلي','Local'))}</b><span>${esc(tr('قائمة الأسعار','Price list'))}</span></div></div>
   <h4 style="margin:8px 0 4px">${esc(tr('حسب المنتج','By product'))}</h4><div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('الكمية','Quantity'))}</th><th class="nosort n">${esc(tr('الإيرادات','Revenue'))}</th><th class="nosort n">${esc(tr('الربح','Profit'))}</th><th class="nosort n">${esc(tr('الهامش','Margin'))}</th></tr></thead><tbody>${A.prods.map(p=>`<tr><td>${esc(p.pn)}</td><td class="n">${p.u==='م³'?q3(p.q):fn(Math.round(p.q))} ${esc(uLbl(p.u))}</td><td class="n">${q2(p.rev)}</td><td class="n">${q2(p.gp)}</td><td class="n">${pc(p.rev?p.gp*100/p.rev:null)}</td></tr>`).join('')}</tbody></table></div>`,
   `<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}

/* ---------- import: a price list, or the Sales_Report quantities ---------- */
const PF={p:['رمز المنتج','رقم المنتج','كود المنتج','product_no','product code','item_no','prod_cd','code'],pn:['اسم المنتج','المنتج','الصنف','product','product_name','item'],u:['الوحده','وحده القياس','uom','unit'],
  cost:['تكلفه الوحده','التكلفه','تكلفه','سعر التكلفه','unit cost','cost','unit_cost'],
  loc:['السعر المحلي','سعر المستهلك','السعر بالدينار','سعر البيع المحلي','local price','price lyd','price_lyd','local_price'],
  fx:['سعر الاجانب','السعر بالعمله الاجنبيه','السعر الاجنبي','سعر البيع للاجانب','foreign price','price usd','price_usd','foreign_price','fx price'],
  from:['تاريخ السريان','ساري من','effective from','effective_from','from']};
function priceList(grids){for(const s of grids){const g=s.grid;for(let i=0;i<Math.min(g.length,20);i++){const H=(g[i]||[]).map(x=>hn(String(x&&typeof x==='object'?'':x)));const m={};
    for(const k in PF){const al=PF[k].map(hn);const j=H.findIndex(h=>h&&al.some(a=>h===a||h.startsWith(a+' (')||h.startsWith(a+' ')));if(j>=0)m[k]=j}
    if((m.p!=null||m.pn!=null)&&(m.cost!=null||m.loc!=null||m.fx!=null)&&m.cost!==m.loc){const rows=[];g.slice(i+1).forEach(r=>{if(!r)return;const v=k=>m[k]==null?'':r[m[k]];const p=String(v('p')||v('pn')).trim();if(!p)return;
      const o={p,pn:String(v('pn')||p).trim(),u:String(v('u')||'').trim()};['cost','loc','fx'].forEach(k=>{if(m[k]!=null){const n=toNum(v(k));o[k]=n==null||n<0?undefined:r4(n)}});
      const fv=v('from');if(fv&&typeof fv==='object'&&fv.serial)o.from=L.serialIso(fv.serial);else if(/^\d{4}-\d{2}-\d{2}/.test(String(fv)))o.from=String(fv).slice(0,10);
      rows.push(o)});if(rows.length)return {rows,map:m,sheet:s.name}}}}return null}
async function importFile(file){toast(tr('جارٍ قراءة الملف…','Reading the file…'));let grids;
  try{grids=await readGrids(file)}catch(e){if(e&&e.cancelled)return;fail(e.message||String(e));return}
  if(window.__pmx&&await window.__pmx.tryImport(grids,file))return;
  const pl=priceList(grids);if(pl){pl.file=file.name;PM.pendingPrices=pl;priceReview(pl);return}
  const cx=CX();if(!cx){fail(tr('تعذّر قراءة الملف.','Could not read the file.'));return}
  let P;try{P=await cx.parseFile('sales',file)}catch(e){fail(tr('الملف ليس قائمة أسعار (رمز المنتج + التكلفة أو السعر) ولا تقرير «تحليل المبيعات».','The file is neither a price list (product code + cost or price) nor a “Sales analysis” report.')+' '+(e.message||''));return}
  PM.pendingSales=P;salesReview(P)}
function fail(m){modal(tr('تعذّر الاستيراد','Import failed'),`<p style="line-height:1.8">${esc(m)}</p>`,`<button type="button" class="btn" data-act="close-dlg">${esc(tr('إغلاق','Close'))}</button>`)}
function priceReview(pl){const c=PM.cfg;const rows=pl.rows.map(o=>{const e=current(o.p)||{};return `<tr><td>${esc(o.pn)} <small class="muted num">${esc(o.p)}</small></td>${['cost','loc','fx'].map(k=>`<td class="n">${o[k]===undefined?'—':`${p4(e[k])} ← <b>${p4(o[k])}</b>`}</td>`).join('')}<td class="n">${o.from?fdi(o.from):'—'}</td></tr>`}).join('');
  modal(tr('مراجعة قائمة الأسعار قبل الاعتماد','Review the price list'),`<p class="muted" style="margin:0 0 6px;font-size:12.5px">${esc(tr(`الملف: ${pl.file} · الورقة: ${pl.sheet} · ${pl.rows.length} منتج. أسعار الوقود لكل متر مكعب، وسعر الأجانب بعملة ${c.cur}.`,`File: ${pl.file} · sheet: ${pl.sheet} · ${pl.rows.length} products. Fuel prices per cubic metre; foreign price in ${c.cur}.`))}</p>
   <div class="tablewrap" style="max-height:340px;min-height:0"><table class="sm-tbl pm-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('التكلفة','Cost'))}</th><th class="nosort n">${esc(tr('المحلي','Local'))}</th><th class="nosort n">${esc(tr('الأجانب','Foreign'))}</th><th class="nosort n">${esc(tr('ساري من','From'))}</th></tr></thead><tbody>${rows}</tbody></table></div>
   ${PM.canEdit?`<div class="pm-set" style="margin-top:8px"><label>${esc(tr('ساري من (للصفوف بلا تاريخ)','Effective from (rows without a date)'))}<input type="date" id="pm-pl-eff" value="${today()}"></label><label>${esc(tr('سبب التغيير','Reason'))}<input type="text" id="pm-pl-why" maxlength="200"></label></div>`:`<p class="muted" style="font-size:12.5px">${esc(tr('اعتماد الأسعار متاح لمحرري إدارة الشؤون التجارية والتحصيلات والإدارة العامة ومدير النظام.','Approving prices is for editors of Commercial Affairs & Collections, the General Administration and the administrator.'))}</p>`}
   <p id="pm-msg" style="font-size:12.5px;margin:6px 0 0"></p>`,
   `${PM.canEdit?`<button type="button" class="btn primary" data-pm="save-pl">${esc(tr('اعتماد الأسعار','Approve prices'))}</button>`:''}<button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}
function salesReview(P){const cx=CX();const ov=PM.local?[]:(cx.ST.sales.batches||[]).filter(b=>!(b.to<P.from||b.from>P.to));const saved=PM.cfg;FXSET=new Set(saved.foreign.map(String));
  const rows=P.rows.map(r=>{const b=baseOf(r);return {d:r.d,c:r.c||r.cn,p:r.p||r.pn,q:b.q,u:b.u}});const priced=rows.filter(r=>line(r).ok).length;
  const m3=rows.filter(r=>r.u==='م³').reduce((a,r)=>a+r.q,0);const canSave=cx.ST.sales.canImport;
  modal(tr('مراجعة ملف تحليل المبيعات','Review the sales analysis file'),`
   <div class="sm-mini"><div><b class="num">${fn(P.rows.length)}</b><span>${esc(tr('حركة مقروءة','lines read'))}</span></div><div><b class="num">${fdi(P.from).slice(0,7)} – ${fdi(P.to).slice(0,7)}</b><span>${esc(tr('الفترة','Period'))}</span></div>
    <div><b class="num">${q3(m3)}</b><span>${esc(tr('م³ من الوقود','m³ of fuel'))}</span></div><div><b class="num">${pct(priced,rows.length)}%</b><span>${esc(tr('حركات مسعّرة','lines priced'))}</span></div></div>
   <p class="muted" style="margin:4px 0;font-size:12.5px">${esc(tr(`الملف: ${P.file} · الورقة: ${P.sheet}. تُحفظ الكميات مع بيانات لوحة «تحليل المبيعات» حتى لا يتكرر الشهر نفسه.`,`File: ${P.file} · sheet: ${P.sheet}. Quantities are saved with the “Sales analysis” data so a month is never counted twice.`))}${P.skipped?` · <b style="color:var(--warn)">${esc(tr(`تُجوهل ${P.skipped} صف`,`${P.skipped} rows skipped`))}</b>`:''}</p>
   ${ov.length?`<div class="banner warn" style="margin-top:8px">${esc(tr(`تتداخل هذه الفترة مع ${ov.length} ملف محفوظ سابقًا.`,`This period overlaps ${ov.length} saved file(s).`))}<label style="display:flex;gap:6px;align-items:center;margin-top:6px"><input type="checkbox" id="pm-replace" checked> ${esc(tr('استبدال الملفات المتداخلة بهذا الملف (يمنع تكرار الأرقام)','Replace the overlapping files (avoids double counting)'))}</label></div>`:''}
   <p id="pm-msg" style="font-size:12.5px;margin:6px 0 0"></p>`,
   `${canSave?`<button type="button" class="btn primary" data-pm="save-sales">${esc(tr('حفظ في المنظومة','Save to the system'))}</button>`:''}<button type="button" class="btn" data-pm="view-local">${esc(tr('عرض دون حفظ','View without saving'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}

/* ---------- export ---------- */
function exportXlsx(){const all=rowsAll(),R=range(all),A=agg(pick(all,R.from,R.to)),c=PM.cfg,rt=rateNow();
  const K=(a,e)=>tr(a,e);
  const sum=[[K('الفترة','Period'),`${R.from.slice(0,7)} – ${R.to.slice(0,7)}`],[K('الإيرادات (د.ل)','Revenue (LYD)'),Math.round(A.rev*100)/100],[K('تكلفة المبيعات (د.ل)','Cost of sales (LYD)'),Math.round(A.cogs*100)/100],
    [K('إجمالي الربح (د.ل)','Gross profit (LYD)'),Math.round(A.gp*100)/100],[K('هامش الربح %','Gross margin %'),A.m==null?'':Math.round(A.m*100)/100],[K('الإضافة على التكلفة %','Markup %'),A.mk==null?'':Math.round(A.mk*100)/100],
    [K('مبيعات الأجانب','Foreign sales')+' ('+c.cur+')',Math.round(A.fxRev*100)/100],[K('الكمية م³','Quantity m³'),Math.round((A.qU.get('م³')||0)*1000)/1000],[K('سعر الصرف الحالي','Current exchange rate'),rt||''],[K('الهامش المستهدف %','Target margin %'),+c.target],[K('الحد الأدنى %','Minimum %'),+c.min]]
    .map(([a,b])=>({[K('البند','Item')]:a,[K('القيمة','Value')]:b}));
  const prod=A.prods.map(p=>({'رمز المنتج':p.p,'اسم المنتج':p.pn,'الوحدة':uLbl(p.u),'الكمية':Math.round(p.q*1000)/1000,'كمية الأجانب':Math.round(p.qFx*1000)/1000,'الإيرادات':Math.round(p.rev*100)/100,'التكلفة':Math.round(p.cogs*100)/100,'إجمالي الربح':Math.round(p.gp*100)/100,'هامش الربح %':p.rev?Math.round(p.gp*10000/p.rev)/100:'','الإضافة %':p.cogs?Math.round(p.gp*10000/p.cogs)/100:'','الربح للوحدة':p.q-p.miss?Math.round(p.gp*100/(p.q-p.miss))/100:'','كمية بلا سعر':Math.round(p.miss*1000)/1000}));
  const cust=A.custs.slice().sort((a,b)=>b.gp-a.gp).map((x,i)=>({'م':i+1,'رقم الحساب':x.c,'اسم الزبون':x.cn,'الفئة':x.fx?'أجنبي':'محلي','الإيرادات':Math.round(x.rev*100)/100,'التكلفة':Math.round(x.cogs*100)/100,'إجمالي الربح':Math.round(x.gp*100)/100,'هامش الربح %':x.rev?Math.round(x.gp*10000/x.rev)/100:''}));
  const mon=[...A.mon.values()].sort((a,b)=>a.k.localeCompare(b.k)).map(o=>({'الشهر':o.k+'-01','الإيرادات':Math.round(o.rev*100)/100,'التكلفة':Math.round(o.cogs*100)/100,'إجمالي الربح':Math.round(o.gp*100)/100,'ربح المحلي':Math.round(o.locGp*100)/100,'ربح الأجانب':Math.round(o.fxGp*100)/100,'هامش الربح %':o.rev?Math.round(o.gp*10000/o.rev)/100:''}));
  const prices=[];c.products.forEach(P=>hist(P.p).forEach(h=>prices.push({'رمز المنتج':P.p,'اسم المنتج':P.pn,'الوحدة':uLbl(P.u),'تاريخ السريان':h.from,'تكلفة الوحدة':h.cost??'','السعر المحلي':h.loc??'',['سعر الأجانب ('+c.cur+')']:h.fx??'','بواسطة':h.by||'','السبب':h.why||''})));
  saveXlsx([[tr('ملخص','Summary'),sum],[tr('المنتجات','Products'),prod],[tr('الزبائن','Customers'),cust],[tr('شهري','Monthly'),mon],[tr('سجل الأسعار','Price history'),prices]],`Profit-margin-${R.from.slice(0,7)}_${R.to.slice(0,7)}.xlsx`)}
function template(){const c=PM.cfg;const rows=boardProducts(rowsAll()).map(x=>{const e=current(x.p)||{};return {'رمز المنتج':x.p,'اسم المنتج':x.pn,'الوحدة':uLbl(x.u),'تكلفة الوحدة':e.cost??'','السعر المحلي':e.loc??'','سعر الأجانب':e.fx??'','تاريخ السريان':today()}});
  if(!rows.length)rows.push({'رمز المنتج':'44347','اسم المنتج':'وقود الديزل','الوحدة':'م³','تكلفة الوحدة':'','السعر المحلي':'','سعر الأجانب':'','تاريخ السريان':today()});
  saveXlsx([[tr('قائمة الأسعار','Price list'),rows]],'Brega-price-list.xlsx')}

/* ---------- events ---------- */
document.addEventListener('click',async e=>{if(S.view!==VIEW&&!(document.getElementById('dlg')||{}).open)return;
  const pre=e.target.closest('[data-pm-pre]');if(pre){const u=UI();u.pre=pre.dataset.pmPre;if(u.pre==='custom'){const R=range(rowsAll());if(R){u.from=u.from||R.from;u.to=u.to||R.to}}u.lim=25;render();return}
  const tb=e.target.closest('[data-pm-tab]');if(tb){UI().tab=tb.dataset.pmTab;render();window.scrollTo(0,0);return}
  const so=e.target.closest('[data-pm-sort]');if(so){const u=UI(),k=so.dataset.pmSort;if(u.sort===k)u.dir=u.dir==='asc'?'desc':'asc';else{u.sort=k;u.dir=k==='cn'?'asc':'desc'}render();return}
  const fx=e.target.closest('[data-pm-fx]');if(fx){e.stopPropagation();fx.disabled=true;await toggleFx(fx.dataset.pmFx);return}
  const hs=e.target.closest('[data-pm-hist]');if(hs){histDlg(hs.dataset.pmHist);return}
  const cu=e.target.closest('[data-pm-cust]');if(cu&&S.view===VIEW&&document.getElementById('view').contains(cu)){custDetail(cu.dataset.pmCust);return}
  const a=e.target.closest('[data-pm]');if(!a)return;const act=a.dataset.pm;
  if(act==='import'){const f=document.getElementById('pm-file');if(f)f.click();return}
  if(act==='export'){exportXlsx();return}
  if(act==='template'){template();return}
  if(act==='print'){try{window.print()}catch(e2){}return}
  if(act==='more'){UI().lim+=50;render();return}
  if(act==='clear'){Object.assign(UI(),{pre:'all',from:'',to:'',prod:'',seg:'',rgn:'',wn:''});render();return}
  if(act==='drop-local'){PM.local=null;render();return}
  if(act==='edit'){PM.edit=true;render();const b=document.getElementById('pm-board');if(b)b.scrollIntoView({block:'start'});return}
  if(act==='cancel-edit'){PM.edit=false;render();return}
  if(act==='save-prices'){await savePrices(a);return}
  if(act==='view-local'){PM.local=PM.pendingSales;PM.pendingSales=null;dlg.close();Object.assign(UI(),{pre:'all',prod:'',q:''});render();return}
  if(act==='save-sales'){const P=PM.pendingSales;if(!P)return;const msg=document.getElementById('pm-msg');a.disabled=true;msg.textContent=tr('جارٍ الحفظ…','Saving…');
    try{const rep=document.getElementById('pm-replace');await CX().save('sales',P,!!(rep&&rep.checked));PM.pendingSales=null;PM.local=null;await CX().load('sales',true);dlg.close();Object.assign(UI(),{pre:'all',prod:''});render();toast(tr(`تم حفظ ${fn(P.rows.length)} حركة`,`${fn(P.rows.length)} lines saved`))}
    catch(err){a.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية الحفظ','You are not allowed to save'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}return}
  if(act==='save-pl'){const pl=PM.pendingPrices;if(!pl)return;const msg=document.getElementById('pm-msg');const eff=(document.getElementById('pm-pl-eff')||{}).value||today(),why=(document.getElementById('pm-pl-why')||{}).value||'';
    const cfg=JSON.parse(JSON.stringify(PM.cfg));const byEff=new Map();pl.rows.forEach(o=>{const f=o.from||eff;(byEff.get(f)||byEff.set(f,[]).get(f)).push(o)});
    let log=[];[...byEff.keys()].sort().forEach(f=>{const save0=PM.cfg;PM.cfg=cfg;log=log.concat(applyPrices(cfg,byEff.get(f),f,why));PM.cfg=save0});
    if(!log.length){dlg.close();toast(tr('لا تغييرات على الأسعار','No price changes'));return}
    a.disabled=true;msg.textContent=tr('جارٍ الحفظ…','Saving…');
    try{await saveCfg(cfg,'استيراد قائمة أسعار هامش الربح',`${pl.file} · ${log.join(' · ')}`);PM.pendingPrices=null;dlg.close();render();toast(tr(`تم اعتماد ${log.length} تغيير`,`${log.length} changes approved`))}
    catch(err){a.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية تعديل الأسعار','You are not allowed to change prices'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}return}
});
document.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;const cu=e.target.closest&&e.target.closest('[data-pm-cust]');if(cu&&S.view===VIEW&&!e.target.closest('[data-pm-fx]')&&document.getElementById('view').contains(cu)){e.preventDefault();custDetail(cu.dataset.pmCust)}});

/* ---------- navigation: a task of Commercial Affairs & Collections ---------- */
if(window.I18N&&I18N.add)I18N.add({'هامش الربح':'Profit margin'});
VIEWS.push({id:VIEW,label:'هامش الربح',icon:'pm',group:'مهام إدارة الشؤون التجارية والتحصيلات',nav:false});
ORG_TASKS.ca=(ORG_TASKS.ca||[]).concat([VIEW]);
const __render=render;
render=function(){if(S.view===VIEW){recompute();document.getElementById('view').innerHTML=dbBanner()+viewMargin();after();return}__render()};
/* the department's page in the org dashboard: a summary card with a link */
const __rp=recPanel;
recPanel=function(unit){const h=__rp(unit);if(unit!==UNIT)return h;let body;
  if(!PM.loaded){load().then(()=>{if(S.view==='org')render()});body=`<span class="muted">${esc(tr('جارٍ التحميل…','Loading…'))}</span>`}
  else{FXSET=new Set(PM.cfg.foreign.map(String));const all=rowsAll();if(!all.length||PM.err)body=`<span class="muted">${esc(PM.err==='denied'?tr('البيانات متاحة للإدارة والإدارة العامة.','Data is available to the department and the General Administration.'):tr('لا توجد كميات مبيعات بعد.','No sales quantities yet.'))}</span>`;
    else{const bd=bounds(all),k=bd.max.slice(0,7),A=agg(all.filter(r=>r.d.slice(0,7)===k));body=`<div class="sm-mini"><div><b class="num">${qs(A.gp)}</b><span>${esc(tr('ربح آخر شهر د.ل','Profit, latest month LYD'))}</span></div><div><b class="num">${pc(A.m)}</b><span>${esc(tr('هامش الربح','margin'))}</span></div></div>`}}
  return `<section class="panel"><h3 class="section-title">${esc(tr('هامش الربح','Profit margin'))} <small>${esc(tr('الربحية والأسعار','profitability & prices'))}</small></h3>${body}<button type="button" class="btn sm primary" data-org-go="${VIEW}" style="margin-top:8px">${esc(tr('فتح لوحة هامش الربح','Open the profit margin dashboard'))}</button></section>`+h};
const __vod=viewOrgDuties;
viewOrgDuties=function(el){__vod(el);const sec=el.querySelector('#od-'+UNIT+' .od-acts');if(sec&&!sec.querySelector(`[data-org-go="${VIEW}"]`))sec.insertAdjacentHTML('beforeend',`<button type="button" class="btn sm" data-org-go="${VIEW}">${esc(tr('هامش الربح','Profit margin'))}</button>`)};
window.__pm={PM,load,agg:rows=>agg(rows),rowsAll,line,priceAt};
})();
