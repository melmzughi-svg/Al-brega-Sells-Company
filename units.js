/* ===== وحدات الكميات: عرض موحّد بالمتر المكعب (Fuel quantity units) =====
   The closing-reports exports state quantities in whatever unit the source system used: litres for one product,
   tonnes for another, cylinders for LPG. The dashboards (SMS055, sales analysis, transfers, receipts, and any module
   that calls window.FuelUnits) show every volume/mass quantity in cubic metres (م³), the international reference:
     litres  → m³  : ÷ 1000 (exact)
     tonnes  → m³  : ÷ density (t/m³) at 15 °C, per product
     kg      → m³  : ÷ density (kg/m³)
     cylinders and other counts stay as they are (a count cannot become a volume without the fill mass).
   International standards applied (calcVersion 2):
     · volume correction to 15 °C: ASTM D1250-04 / API MPMS Ch. 11.1-2004 / ISO 91, Table 54B (refined products), computed
       from the standard's own implementation procedure (ρ15 → ρ60, ITS-90 temperatures, CTL rounded to 5 decimals);
     · tonnes ↔ volume: ASTM D1250 Table 56, metric tonnes in air (commercial weight, kg in air = V15 × (ρ15 − 1.1)),
       or in vacuum (true mass, V15 × ρ15) when the administrator chooses it;
     · density at 15 °C: the value measured on the delivery (ISO 3675 / ISO 12185) when the imported line carries it,
       else the administrator's setting, else a reference value inside the product specification range
       (EN 228, EN 590, ASTM D1655, ISO 8217, EN 589) — shown with its source in Settings ← «وحدات الكميات» (Firestore config/units).
   Loaded before the dashboards. Every row keeps its original quantity (q0) and unit (u0) for tooltips and exports. */
"use strict";
(function(){
const EN=()=>!!(window.I18N&&I18N.lang==='en');
const T=(a,e)=>EN()?e:a;
const M3='م³';
/* the product types sold by the company (the «النوع» list of the sales system): gasoline, diesel, LPG, heavy fuel oil
   and marine fuel. Marine fuel is checked first because «marine gas oil» would otherwise read as diesel; its default
   0.890 t/m³ is marine gas oil (MGO), heavier marine fuels reach 0.991 and are set per product in Settings. */
const FAM=[ // [family id, Arabic, English, reference density kg/m³ at 15 °C, name patterns (normalised, no spaces), product codes]
  ['marine','وقود السفن','Marine fuel',890,/سفن|بحري|marine|bunker|mgo|mdo|ifo/,[]],
  ['gasoline','البنزين','Gasoline',740,/بنزين|gasoline|petrol|mogas/,['33500']],
  ['diesel','وقود الديزل','Diesel',835,/ديزل|سولار|diesel|gasoil/,['44347']],
  ['lpg','غاز مسال','LPG',540,/مسال|غاز|lpg|propane|butane|بوتان|بروبان/,['16291']],
  ['kerosene','كيروسين','Kerosene',800,/كيروسين|كيروسين|كاز|kerosene|kero|jet|avtur/,[]],
  ['fueloil','الزيت الثقيل','Heavy fuel oil',970,/ثقيل|زيتوقود|fueloil|hfo|mazut|مازوت/,[]]];
/* the specification each product type is sold to: density limits at 15 °C (kg/m³) and the measurement table.
   [min, max, standard, table]. The reference value above lies inside the range; a measured density always wins. */
const SPEC={
  gasoline:[720,775,'EN 228','54B'],
  diesel:[820,845,'EN 590','54B'],
  kerosene:[775,840,'ASTM D1655 / DEF STAN 91-091 (Jet A-1)','54B'],
  fueloil:[900,1010,'ISO 8217 (RMG ≤ 991 · RMK ≤ 1010)','54B'],
  marine:[800,1010,'ISO 8217 (DMA ≤ 890 · DMB ≤ 900 · RM ≤ 1010)','54B'],
  lpg:[500,585,'EN 589 / ISO 9162 (C3 ≈ 508 · C4 ≈ 585)','54E']};
/* contract / line types that are not a product (services, maintenance, transport, supply): never converted to a volume */
const NOT_PRODUCT=/خدم|صيان|نقلوتشغيل|^نقل|تشغيل|توريد|service|maintenance|transport|supply/;
const U={over:{},names:{},temp:15,temps:{},basis:'air',loaded:false,ver:0};
const CALC_VERSION=2;
/* ---------- temperature: litres and m³ read at the tank temperature are corrected to standard volume at 15 °C ----------
   V15 = Vobs × CTL. CTL (the volume correction factor) per ASTM D1250-04 / API MPMS 11.1-2004, Table 54B (generalized
   refined products), following the standard's implementation procedure:
     1. commodity group and constants K0, K1, K2 from the density at 60 °F (ρ60):
          fuel oils 838.3127 ≤ ρ60 ≤ 1163.5      K0 103.8720   K1 0.2701   K2 0
          jet / kerosene 787.5195 ≤ ρ60 < 838.3127  K0 330.3010   K1 0        K2 0
          transition 770.352 ≤ ρ60 < 787.5195     K0 1489.0670  K1 0        K2 −0.0018684
          gasolines 611.16 ≤ ρ60 < 770.352        K0 192.4571   K1 0.2438   K2 0
     2. temperatures are ITS-90 and are shifted to IPTS-68 (the scale the coefficients were fitted on);
     3. ρ60 is adjusted by δ60 = 0.01374979547 °F to ρ*, α60 = (K0/ρ* + K1)/ρ* + K2;
     4. CTL(t) = exp(−α60·ΔT·(1 + 0.8·α60·(ΔT + δ60))), ΔT = t(IPTS-68, °F) − 60.0068749;
     5. 15 °C base: ρ15 is converted to ρ60 by iteration (ρ60 = ρ15 / CTL(59 °F)), and CTL15(t) = CTL(t) / CTL(59 °F).
   Inputs are rounded as the standard does (density to 0.1 kg/m³, temperature to 0.05 °C) and the factor to 5 decimals.
   Check values (Table 54B): 835.0 at 30 °C → 0.98716; 740.0 at 30 °C → 0.98151; 800.0 at 35 °C → 0.98131;
   970.0 at 40 °C → 0.98241; any density at 15 °C → 1.00000.
   Tonnes need no temperature: mass ÷ density at 15 °C is already V15. LPG is outside Table 54B (ρ < 611): it needs
   Table 54E (GPA TP-27) with the vapour pressure, so its volume stays as read unless the administrator enters the factor
   from the supplier's 54E certificate. The tables named in a supply contract remain the reference for custody transfer. */
const D60=0.01374979547;
const GRP=[[838.3127,103.8720,0.2701,0],[787.5195,330.3010,0,0],[770.352,1489.0670,0,-0.0018684],[611.16,192.4571,0.2438,0]];
const RHO_MIN=611.16,RHO_MAX=1163.5;
function grp(r){for(const g of GRP)if(r>=g[0])return g;return GRP[3]}
/* ITS-90 °F → IPTS-68 °F (API MPMS 11.1-2004, §11.1.5.3) */
function t68F(tF){const tc=(tF-32)/1.8,tau=tc/630,a=[-0.148759,-0.267408,1.08076,1.269056,-4.089591,-1.871251,7.438081,-3.536296];
  let s=a[7];for(let i=6;i>=0;i--)s=a[i]+s*tau;return (tc-s*tau)*1.8+32}
/* CTL from 60 °F to tF for a density at 60 °F */
function ctl60(r60,tF,g){g=g||grp(r60);const K0=g[1],K1=g[2],K2=g[3];
  const A=D60/2*((K0/r60+K1)/r60+K2),B=(2*K0+K1*r60)/(K0+(K1+K2*r60)*r60);
  const rs=r60*(1+(Math.exp(A*(1+0.8*A))-1)/(1+A*(1+1.6*A)*B));const a=(K0/rs+K1)/rs+K2;const dT=t68F(tF)-60.0068749;
  return Math.exp(-a*dT*(1+0.8*a*(dT+D60)))}
/* density at 15 °C → density at 60 °F (iteration of the standard; converges in a few steps) */
function rho60(r15){let r=r15;for(let i=0;i<50;i++){const n=r15/ctl60(r,59);if(Math.abs(n-r)<1e-9)return n;r=n}return r}
/* the unrounded Table 54B factor, or null outside the table (density 611.16–1163.5 kg/m³ at 60 °F, −18 to 90 °C here) */
function ctl54B(d15,t){if(!(d15>0)||!(t>=-18&&t<=90))return null;const r60=rho60(d15);if(r60<RHO_MIN||r60>RHO_MAX)return null;const g=grp(r60);
  return ctl60(r60,t*1.8+32,g)/ctl60(r60,59,g)}
function vcf(d15,t){d15=Math.round(+d15*10)/10;t=Math.round(+t*20)/20;if(t===15)return 1;const f=ctl54B(d15,t);return f==null?1:Math.round(f*100000)/100000}
/* the 1980 tables (ASTM D1250-80) closed form, kept only to compare old figures (agrees with 2004 to ±0.0001) */
function vcf1980(d,t){const a=d<770.5?346.4228/(d*d)+0.4388/d:d<787.5?-0.00336312+2680.3206/(d*d):d<839?594.5418/(d*d):186.9696/(d*d)+0.4862/d,dt=t-15;return Math.exp(-a*dt*(1+0.8*a*dt))}
/* density observed at tObs (hydrometer / density meter, kg/m³) → density at 15 °C (ISO 91 / Table 53B, by iteration) */
function density15From(dObs,tObs){let d=+dObs;for(let i=0;i<20;i++){const f=ctl54B(d,tObs);if(f==null)return null;const n=dObs/f;if(Math.abs(n-d)<0.001){d=n;break}d=n}return Math.round(d*10)/10}
/* density per m³ used to turn mass into standard volume: in air (Table 56, ρ15 − 1.1) or in vacuum (ρ15) */
function massDensity(d15,basis){return (basis||U.basis)==='vacuum'?d15:d15-1.1}
/* the observed temperature for a product: the row's own reading, else its type's setting, else the general setting */
function tempFor(code,name,rowT){const t=+rowT;if(rowT!==''&&rowT!=null&&isFinite(t)&&t!==0&&t>=-18&&t<=90)return {t,src:'row'};const f=family(code,name);
  if(f&&U.temps['@'+f[0]]!=null&&isFinite(+U.temps['@'+f[0]]))return {t:+U.temps['@'+f[0]],src:'type'};return {t:+U.temp||15,src:'general'}}
const nrm=s=>(window.norm?norm(String(s||'')):String(s||'').toLowerCase()).replace(/[\s\/\\\-_.()]/g,'');
/* unit kind from the unit text: L, T, KG, M3, or '' for counts (cylinders…) */
function kind(u){const s=nrm(u);if(!s)return'';
  if(/م3|م³|m3|m³|مترمكعب|cbm|cubic/.test(s))return'M3';
  if(/اسطوان|سلندر|cyl|عبوه|عبوة|قاروره|كرتون|حبه|قطعه|pcs|unit/.test(s))return'';
  if(/كغم|كجم|كيلو|kg|kilo/.test(s))return'KG';
  if(/طن|ton|^mt$|^t$|tonne/.test(s))return'T';
  if(/لتر|ليتر|ltr|lit|^l$|liter|litre|سالب/.test(s))return'L';return''}
function isService(name){const n=nrm(name);return !!n&&NOT_PRODUCT.test(n)&&!FAM.some(f=>f[4].test(n))}
function family(code,name){const c=String(code||'').trim(),n=nrm(name);for(const f of FAM){if(f[5].includes(c)||(n&&f[4].test(n)))return f}return null}
/* density in kg/m³ for a product: admin override by code, else the family default, else null */
function density(code,name,rowD){const c=String(code||'').trim();const f=family(c,name);let m=+rowD;if(m>0&&m<5)m*=1000;/* kg/l → kg/m³ */
  if(m>=400&&m<=1200)return {d:Math.round(m*10)/10,src:'measured',fam:f||null};
  if(U.over[c]>0)return {d:U.over[c],src:'override',fam:f||null};if(!f)return null;
  if(U.over['@'+f[0]]>0)return {d:U.over['@'+f[0]],src:'override',fam:f};return {d:f[3],src:'default',fam:f}}
/* convert one quantity; returns {q,u,d} with u=م³, or null when the unit is a count or no density is known */
function lpgVcf(code,name){const v=+U.lpgVcf;return v>0.8&&v<1.2?v:null}
function toM3(q,u,code,name,rowT,rowD){if(isService(name))return null;const k=kind(u);if(!k)return null;const n=+q||0;
  if(k==='M3'||k==='L'){const v=k==='L'?n/1000:n;const dd=density(code,name,rowD);const tt=tempFor(code,name,rowT);const lpg=!!(dd&&dd.fam&&dd.fam[0]==='lpg');
    if(lpg){const lf=lpgVcf(code,name);return lf?{q:v*lf,u:M3,k,d:dd.d,src:dd.src,t:tt.t,vcf:lf,qObs:v,table:'54E',lpg:true}:{q:v,u:M3,k,t:tt.t,vcf:1,lpg:true}}
    if(!dd||tt.t===15)return {q:v,u:M3,k,t:tt.t,vcf:1,d:dd?dd.d:0,src:dd?dd.src:''};
    const f=vcf(dd.d,tt.t);return {q:v*f,u:M3,k,d:dd.d,src:dd.src,t:tt.t,tsrc:tt.src,vcf:f,qObs:v,table:'54B'}}
  const dd=density(code,name,rowD);if(!dd)return null;const md=massDensity(dd.d);return {q:k==='T'?n*1000/md:n/md,u:M3,k,d:dd.d,src:dd.src,basis:U.basis}}
/* full calculation record for one quantity, to store with a movement so any figure can be recomputed later */
function convert({q,u,code,name,t,d}){const c=toM3(q,u,code,name,t,d);if(!c)return null;const md=c.d?massDensity(c.d):0;
  return {entered:+q,enteredUnit:u,tObs:c.t!=null?c.t:null,density15:c.d||null,densitySource:c.src||'',basis:U.basis,table:c.table||(c.k==='T'||c.k==='KG'?'56':''),vcf:c.vcf||1,
    v15_m3:c.q,v15_l:c.q*1000,tonnes:md?c.q*md/1000:null,calcVersion:CALC_VERSION}}
/* normalise rows of a dashboard in place-safe copies: q (and rq when present) in م³, u=م³; originals in q0/u0; memoised per array */
const memo=new WeakMap();
function normalize(rows){if(!rows||!rows.length)return rows||[];const m=memo.get(rows);if(m&&m.ver===U.ver)return m.out;
  const out=rows.map(r=>{if(!r.u&&!r.u0)return r;const u0=r.u0||r.u,q0=r.q0!=null?r.q0:r.q;const c=toM3(q0,u0,r.p,r.pn,r.tc,r.dc);if(!c){if(r.u0)return Object.assign({},r,{u:u0,q:q0,rq:r.rq0!=null?r.rq0:r.rq});return r}
    const o=Object.assign({},r,{q:c.q,u:M3,q0,u0,uk:c.k,ud:c.d||0,us:c.src||'',ut:c.t,uv:c.vcf||1});if(r.rq!=null){o.rq0=r.rq0!=null?r.rq0:r.rq;o.rq=toM3(o.rq0,u0,r.p,r.pn,r.tc,r.dc).q}return o});
  memo.set(rows,{ver:U.ver,out});return out}
/* text for the conversions actually used in a set of rows, e.g. «لتر ÷ 1000 · طن × 1000 ÷ (835.0 − 1.1) (وقود الديزل)» */
function explain(rows){const seen=new Map();(rows||[]).forEach(r=>{if(!r.u0||r.u!==M3)return;const key=r.uk+(r.uk==='T'||r.uk==='KG'?'|'+r.p:'');if(!seen.has(key))seen.set(key,r)});
  const cor=new Map();(rows||[]).forEach(r=>{if(r.u===M3&&r.uv&&r.uv!==1){const k=r.pn+'|'+r.ut;if(!cor.has(k))cor.set(k,r)}});
  const air=U.basis!=='vacuum';const div=d=>air?`(${(+d).toFixed(1)} − 1.1)`:(+d).toFixed(1);
  const parts=[];seen.forEach(r=>{if(r.uk==='L')parts.push(T('لتر ÷ 1000','litres ÷ 1000'));else if(r.uk==='M3')parts.push(T('م³ كما هي','m³ as is'));
    else parts.push(T(`${r.uk==='T'?'طن × 1000':'كغم'} ÷ ${div(r.ud)} (${r.pn}${r.us==='measured'?'، كثافة مقيسة':''})`,`${r.uk==='T'?'tonnes × 1000':'kg'} ÷ ${div(r.ud)} (${r.pn}${r.us==='measured'?', measured density':''})`))});
  const corTxt=[...cor.values()].slice(0,6).map(r=>T(`${r.pn}: ${r.ut}°م × ${r.uv.toFixed(5)}`,`${r.pn}: ${r.ut} °C × ${r.uv.toFixed(5)}`));
  if(corTxt.length)parts.push(T('مصحّحة إلى 15°م بجدول ASTM D1250 54B: ','corrected to 15 °C with ASTM D1250 Table 54B: ')+corTxt.join('، '));
  const mass=[...seen.values()].some(r=>r.uk==='T'||r.uk==='KG');
  return parts.length?T('الكميات بالمتر المكعب القياسي عند 15°م: ','Quantities in standard cubic metres at 15 °C: ')+parts.join(' · ')+(mass?(air?T(' · الطن متري في الهواء (جدول ASTM D1250 رقم 56)، والكثافة عند 15°م من الإعدادات',' · metric tonnes in air (ASTM D1250 Table 56); density at 15 °C from Settings'):T(' · الطن كتلة حقيقية في الفراغ، والكثافة عند 15°م من الإعدادات',' · tonnes in vacuum (true mass); density at 15 °C from Settings')):''):''}
/* products seen by the dashboards, so the settings panel can list them */
function seen(rows){(rows||[]).forEach(r=>{if(r.p&&r.pn&&!U.names[r.p])U.names[r.p]=r.pn})}
/* admin configuration: config/units {densities:{code:kg/m³}, temps:{'@type':°C}, temp, basis:'air'|'vacuum', lpgVcf, names:{code:name}} */
async function load(){if(U.loaded)return;U.loaded=true;try{const FS=window.__fs&&window.__fs();if(!FS)return;
  FS.collection('config').doc('units').onSnapshot(d=>{const x=(d.exists&&d.data())||{};U.over={};Object.entries(x.densities||{}).forEach(([k,v])=>{if(+v>0)U.over[k]=+v});U.temp=isFinite(+x.temp)&&x.temp!==''&&x.temp!=null?+x.temp:15;U.temps={};Object.entries(x.temps||{}).forEach(([k,v])=>{if(v!==''&&v!=null&&isFinite(+v))U.temps[k]=+v});
    U.basis=x.basis==='vacuum'?'vacuum':'air';U.lpgVcf=+x.lpgVcf>0?+x.lpgVcf:null;Object.assign(U.names,x.names||{});U.ver++;if(typeof render==='function'&&['settings','sms055','sales','transfers','receipts','org','margin'].includes(window.S&&S.view))render()},()=>{})}catch(e){}}
async function save(dens,temps,temp,basis,lpg){const FS=window.__fs(),m=(window.__me&&window.__me())||{};const names={};Object.keys(dens).forEach(k=>{if(U.names[k])names[k]=U.names[k]});
  temps=temps||U.temps;temp=temp==null?U.temp:temp;basis=basis==='vacuum'?'vacuum':basis==='air'?'air':U.basis;lpg=lpg===undefined?U.lpgVcf:lpg;
  await FS.collection('config').doc('units').set({densities:dens,temps,temp,basis,lpgVcf:lpg||null,names,calcVersion:CALC_VERSION,by:m.email||'',at:new Date().toISOString()});
  try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'تعديل كثافات المنتجات',col:'config',id:'units',summary:(Object.entries(dens).map(([k,v])=>`${names[k]||k}: ${v}`).join('، ')+` · درجة الحرارة العامة ${temp}°م · الطن ${basis==='vacuum'?'في الفراغ':'في الهواء'}`+(lpg?` · عامل الغاز المسال ${lpg}`:'')+(Object.keys(temps).length?' · '+Object.entries(temps).map(([k,v])=>`${k}: ${v}°م`).join('، '):'')).slice(0,1500)})}catch(e){}}

/* ---------- settings panel (administrator) ---------- */
(function(){if(typeof document==='undefined'||!document.createElement)return;const st=document.createElement('style');st.id='fu-css';st.textContent=`
#fu-panel{padding:0;overflow:hidden;grid-column:1/-1;min-width:0}
#fu-panel .fu-body>section{min-width:0}
#fu-panel .fu-hero{position:relative;padding:20px 22px 18px;background:linear-gradient(135deg,color-mix(in srgb,var(--accent) 14%,var(--surface)) 0%,var(--surface) 62%);border-bottom:1px solid var(--line)}
#fu-panel .fu-hero h3{margin:0;font-family:var(--f-display);font-size:18px;letter-spacing:-.2px}
#fu-panel .fu-hero p{margin:4px 0 12px;color:var(--muted);font-size:13px;font-weight:500;max-width:880px;line-height:1.75}
#fu-panel .fu-std{display:flex;flex-wrap:wrap;gap:6px}
#fu-panel .fu-std span{direction:ltr;font:600 11.5px/1 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;padding:6px 9px;border-radius:6px;background:var(--surface);border:1px solid var(--line);color:var(--ink)}
#fu-panel .fu-std span b{color:var(--accent);font-weight:700}
#fu-panel .fu-ver{position:absolute;inset-inline-end:22px;top:20px;font-size:11.5px;color:var(--muted);font-weight:600;display:flex;gap:6px;align-items:center}
#fu-panel .fu-ver i{width:8px;height:8px;border-radius:50%;background:var(--ok);box-shadow:0 0 0 3px var(--ok-soft)}
#fu-panel .fu-body{padding:18px 22px 20px;display:grid;gap:18px}
#fu-panel .fu-h{margin:0 0 10px;font-size:14px;display:flex;align-items:center;gap:8px}
#fu-panel .fu-h small{color:var(--muted);font-weight:500;font-size:12px}
#fu-panel .fu-n{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:6px;background:var(--accent);color:var(--on-accent);font-size:12px}
#fu-panel .fu-methods{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:12px}
#fu-panel .fu-m{border:1px solid var(--line);border-radius:10px;padding:14px;background:var(--surface2)}
#fu-panel .fu-m h4{margin:0 0 2px;font-size:13.5px}
#fu-panel .fu-m .fu-ref{font:600 11px ui-monospace,Menlo,Consolas,monospace;color:var(--accent);direction:ltr;display:inline-block}
#fu-panel .fu-f{direction:ltr;text-align:center;margin:10px 0 8px;padding:10px 8px;border-radius:8px;background:var(--surface);border:1px dashed var(--line2,var(--line));font:600 13.5px/1.5 "Cambria Math","STIX Two Math","Times New Roman",serif}
#fu-panel .fu-f sub{font-size:10px}
#fu-panel .fu-m p{margin:0;color:var(--muted);font-size:12.5px;font-weight:500;line-height:1.7}
#fu-panel .fu-m ol{margin:6px 0 0;padding-inline-start:18px;font-size:12.5px;font-weight:500;line-height:1.9}
#fu-panel .fu-conv{border:1px solid var(--line);border-radius:12px;overflow:hidden}
#fu-panel .fu-cin{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;padding:14px;background:var(--surface2);border-bottom:1px solid var(--line)}
#fu-panel .fu-cin label{display:grid;gap:4px;font-size:12px;color:var(--muted);font-weight:600}
#fu-panel .fu-cin input,#fu-panel .fu-cin select{width:100%;box-sizing:border-box}
#fu-panel .fu-out{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:1px;background:var(--line)}
@media (max-width:620px){#fu-panel .fu-out{grid-template-columns:repeat(2,minmax(0,1fr))}#fu-panel .fu-hero{padding:16px}#fu-panel .fu-body{padding:14px}#fu-panel .fu-ver{position:static;margin-bottom:8px}}
#fu-panel .fu-t{background:var(--surface);padding:12px 14px}
#fu-panel .fu-t small{display:block;color:var(--muted);font-size:11.5px;font-weight:600}
#fu-panel .fu-t b{display:block;font-size:20px;letter-spacing:-.3px;font-variant-numeric:tabular-nums;direction:ltr;text-align:start;unicode-bidi:plaintext}
#fu-panel .fu-t em{font-style:normal;color:var(--muted);font-size:11.5px;font-weight:500}
#fu-panel .fu-t.key{background:color-mix(in srgb,var(--accent) 9%,var(--surface))}
#fu-panel .fu-t.key b{color:var(--accent)}
#fu-panel .fu-trace{padding:10px 14px;font:500 12px/1.7 ui-monospace,Menlo,Consolas,monospace;color:var(--muted);direction:ltr;text-align:left;background:var(--surface);border-top:1px solid var(--line);overflow-x:auto;white-space:nowrap}
#fu-panel .fu-warn{padding:8px 14px;font-size:12.5px;color:var(--warn);background:var(--warn-soft);border-top:1px solid var(--line)}
#fu-panel .fu-grid2{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(260px,1fr);gap:18px;align-items:start}
@media (max-width:980px){#fu-panel .fu-grid2{grid-template-columns:1fr}}
#fu-panel table.fu-tb{width:100%;border-collapse:separate;border-spacing:0}
#fu-panel table.fu-tb th{font-size:11.5px;color:var(--muted);font-weight:700;text-align:start;padding:8px 10px;border-bottom:1px solid var(--line);background:var(--surface2);position:static}
#fu-panel table.fu-tb td{padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:middle;font-size:13px}
#fu-panel tr.fu-fam td{background:color-mix(in srgb,var(--accent) 4%,var(--surface))}
#fu-panel .fu-bar{position:relative;height:6px;border-radius:3px;background:var(--surface2);border:1px solid var(--line);margin-top:6px;min-width:120px;direction:ltr}
#fu-panel .fu-bar i{position:absolute;top:0;bottom:0;background:color-mix(in srgb,var(--ok) 35%,transparent);border-radius:3px}
#fu-panel .fu-bar u{position:absolute;top:-4px;width:3px;height:12px;border-radius:2px;background:var(--accent);transform:translateX(-1px)}
#fu-panel .fu-bar u.out{background:var(--crit)}
#fu-panel .fu-spec{font-size:11.5px;color:var(--muted);font-weight:600}
#fu-panel .fu-spec span{direction:ltr;unicode-bidi:isolate}
#fu-panel .fu-pill{display:inline-block;font-size:11px;font-weight:700;padding:3px 8px;border-radius:999px;background:var(--neutral-soft);color:var(--neutral)}
#fu-panel .fu-pill.ok{background:var(--ok-soft);color:var(--ok)}#fu-panel .fu-pill.edit{background:var(--info-soft);color:var(--info)}#fu-panel .fu-pill.bad{background:var(--crit-soft);color:var(--crit)}
#fu-panel .fu-tbl{font:500 11px ui-monospace,Menlo,Consolas,monospace;color:var(--muted);direction:ltr;display:inline-block}
#fu-panel .fu-vcf{border:1px solid var(--line);border-radius:12px;overflow:hidden}
#fu-panel .fu-vcf header{padding:10px 14px;background:var(--surface2);border-bottom:1px solid var(--line);font-size:13px}
#fu-panel .fu-vcf header small{display:block;color:var(--muted);font-weight:500;font-size:11.5px}
#fu-panel .fu-vcf table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums;direction:ltr}
#fu-panel .fu-vcf td,#fu-panel .fu-vcf th{padding:6px 12px;font:600 12.5px ui-monospace,Menlo,Consolas,monospace;border-bottom:1px solid var(--line);text-align:right}
#fu-panel .fu-vcf th{font-size:11px;color:var(--muted);text-align:right}
#fu-panel .fu-vcf td:first-child,#fu-panel .fu-vcf th:first-child{text-align:left}
#fu-panel .fu-vcf tr.on td{background:color-mix(in srgb,var(--accent) 10%,var(--surface));color:var(--accent)}
#fu-panel .fu-set{display:flex;flex-wrap:wrap;gap:12px 18px;align-items:center;padding:12px 14px;border:1px solid var(--line);border-radius:10px;background:var(--surface2)}
#fu-panel .fu-set label{display:inline-flex;flex-wrap:wrap;gap:8px;align-items:center;font-size:13px;max-width:100%}
#fu-panel .fu-note{font-size:12px;color:var(--muted);font-weight:500;line-height:1.8;border-inline-start:3px solid var(--line);padding-inline-start:10px;margin:0}
#fu-panel .fu-foot{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between;padding-top:4px}
`;document.head.appendChild(st)})();
function specTxt(f){const x=f&&SPEC[f[0]];return x?`${x[2]}: ${x[0]}–${x[1]}`:''}
const nf=(v,d)=>(+v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
/* range bar: spec band and the current value, on a common 480–1080 kg/m³ scale */
function bar(f,d){const x=f&&SPEC[f[0]];if(!x)return '';const L=480,H=1080,P=v=>Math.max(0,Math.min(100,(v-L)/(H-L)*100));const out=d!==''&&(d<x[0]||d>x[1]);
  return `<div class="fu-bar" title="${esc(x[0]+'–'+x[1]+' kg/m³')}"><i style="left:${P(x[0]).toFixed(2)}%;width:${(P(x[1])-P(x[0])).toFixed(2)}%"></i>${d!==''?`<u class="${out?'out':''}" style="left:${P(+d).toFixed(2)}%"></u>`:''}</div>`}
function panel(){const me=(window.__me&&window.__me())||{};if(me.role!=='admin')return '';load();
  const ORD=['gasoline','fueloil','lpg','diesel','marine','kerosene'];/* the order of the sales system's «النوع» list */
  const famRows=FAM.slice().sort((a,b)=>ORD.indexOf(a[0])-ORD.indexOf(b[0])).map(f=>{const o=U.over['@'+f[0]];return {c:'@'+f[0],shown:T('كل منتجات النوع','all products of this type'),n:T(f[1],f[2]),d:o>0?o:f[3],src:o>0?'override':'default',fam:f}});
  const codes=new Set([...Object.keys(U.names),...Object.keys(U.over).filter(k=>k[0]!=='@'),...FAM.flatMap(f=>f[5])].filter(c=>!isService(U.names[c])));
  const prodRows=[...codes].map(c=>{const dd=density(c,U.names[c]);const f=family(c,U.names[c]);return {c,n:U.names[c]||(f?T(f[1],f[2]):''),d:dd?dd.d:'',src:dd?dd.src:'',fam:f}}).sort((a,b)=>a.n.localeCompare(b.n,'ar'));
  const rows=famRows.concat(prodRows);const air=U.basis!=='vacuum';
  const out=r=>{const x=r.fam&&SPEC[r.fam[0]];return x&&r.d!==''&&(r.d<x[0]||r.d>x[1])};
  const famOpts=famRows.map(r=>`<option value="${esc(r.c)}" data-d="${esc(String(r.d))}"${r.c==='@diesel'?' selected':''}>${esc(r.n)}</option>`).join('');
  const src=r=>r.src==='override'?`<span class="fu-pill edit">${esc(T('معدّلة','edited'))}</span>`:r.src==='default'?`<span class="fu-pill ok">${esc(T('مرجعية ضمن المواصفة','reference, in spec'))}</span>`:`<span class="fu-pill bad">${esc(T('غير معروفة','unknown'))}</span>`;
  return `<div class="panel" id="fu-panel">
  <div class="fu-hero"><div class="fu-ver"><i></i>${esc(T('محرّك الحساب، الإصدار '+CALC_VERSION,'Calculation engine v'+CALC_VERSION))}</div>
   <h3 class="section-title" style="margin:0">${esc(T('وحدات الكميات والتحويلات القياسية','Quantity units & standard conversions'))}</h3>
   <p>${esc(T('كل لوحات المنظومة تعرض الكميات بالمتر المكعب القياسي عند 15°م، محسوبة بمعادلات جداول قياس البترول الدولية لا بقيم تقريبية: تصحيح الحجم بالحرارة، وتحويل الطن بالوزن في الهواء، وكثافة كل منتج ضمن حدود مواصفته الدولية.','Every dashboard shows quantities in standard cubic metres at 15 °C, computed with the equations of the international Petroleum Measurement Tables rather than rules of thumb: temperature volume correction, tonnes as weight in air, and each product’s density within its international specification.'))}</p>
   <div class="fu-std"><span><b>ISO 91</b>:2017</span><span><b>ASTM D1250</b>-04</span><span><b>API MPMS</b> 11.1-2004</span><span>Table <b>54B</b> · CTL</span><span>Table <b>53B</b> · ρ15</span><span>Table <b>56</b> · t in air</span><span><b>ISO 12185</b> / 3675 · ρ</span></div></div>
  <div class="fu-body">
   <section><h4 class="fu-h"><span class="fu-n">1</span>${esc(T('منهجية الحساب','Calculation method'))}</h4>
   <div class="fu-methods">
    <div class="fu-m"><h4>${esc(T('تصحيح الحجم إلى 15°م','Volume correction to 15 °C'))}</h4><span class="fu-ref">ASTM D1250-04 · Table 54B</span>
     <div class="fu-f">V<sub>15</sub> = V<sub>obs</sub> × CTL(ρ<sub>15</sub>, t)</div>
     <p>${esc(T('اللتر المقروء عند حرارة الخزان يتمدد بالحرارة، فيُعاد إلى حجمه عند 15°م بالمعادلات الرسمية للمعيار (تحويل ρ15 إلى ρ60 وحرارة ITS-90)، والعامل مقرّب إلى 5 خانات. مثال: ديزل 835.0 عند 30°م ← 0.98716.','Litres read at tank temperature are expanded by heat, so they are brought back to 15 °C with the standard’s own equations (ρ15 → ρ60, ITS-90 temperatures), factor rounded to 5 decimals. Example: diesel 835.0 at 30 °C → 0.98716.'))}</p></div>
    <div class="fu-m"><h4>${esc(T('الطن والحجم','Tonnes and volume'))}</h4><span class="fu-ref">ASTM D1250 · Table 56</span>
     <div class="fu-f">${air?'t = V<sub>15</sub> × (ρ<sub>15</sub> − 1.1) / 1000':'t = V<sub>15</sub> × ρ<sub>15</sub> / 1000'}</div>
     <p>${esc(air?T('الطن المتري في الهواء هو الوزن التجاري في الفواتير وبوالص الشحن؛ يُطرح 1.1 كغم/م³ لدفع الهواء. يمكن اختيار الكتلة الحقيقية في الفراغ من الإعدادات أدناه.','Metric tonnes in air are the commercial weight on invoices and bills of lading; 1.1 kg/m³ of air buoyancy is deducted. True mass in vacuum can be chosen in the settings below.'):T('الكتلة الحقيقية في الفراغ. الوزن التجاري في الفواتير عادةً في الهواء (ρ15 − 1.1)، ويمكن اختياره من الإعدادات أدناه.','True mass in vacuum. Invoices usually state weight in air (ρ15 − 1.1), selectable in the settings below.'))}</p></div>
    <div class="fu-m"><h4>${esc(T('مصدر الكثافة عند 15°م','Density at 15 °C: source order'))}</h4><span class="fu-ref">ISO 12185 · ISO 3675 · ISO 91</span>
     <ol><li>${esc(T('الكثافة المقيسة في سطر ملف الاستيراد (شهادة الجودة)','Measured density on the imported line (quality certificate)'))}</li><li>${esc(T('كثافة المنتج أو نوعه المعدّلة في الجدول أدناه','Product or type density edited in the table below'))}</li><li>${esc(T('القيمة المرجعية داخل حدود المواصفة الدولية للمنتج','Reference value inside the product’s international specification'))}</li></ol></div>
   </div></section>
   <section><h4 class="fu-h"><span class="fu-n">2</span>${esc(T('المحوّل القياسي','Standard converter'))} <small>${esc(T('جرّب أي كمية لترى الحساب كاملًا','try any quantity to see the full calculation'))}</small></h4>
   <div class="fu-grid2"><div class="fu-conv">
    <div class="fu-cin"><label>${esc(T('المنتج','Product'))}<select id="fu-c-p">${famOpts}</select></label>
     <label>${esc(T('الكمية','Quantity'))}<input type="number" id="fu-c-q" value="10000" step="any" min="0"></label>
     <label>${esc(T('الوحدة','Unit'))}<select id="fu-c-u"><option value="L">${esc(T('لتر مقروء','observed litres'))}</option><option value="M3">${esc(T('م³ مقروء','observed m³'))}</option><option value="T">${esc(air?T('طن متري (في الهواء)','metric tonnes (in air)'):T('طن متري (في الفراغ)','metric tonnes (vacuum)'))}</option></select></label>
     <label>${esc(T('حرارة المنتج °م','Product temperature °C'))}<input type="number" id="fu-c-t" value="30" step="0.05" min="-18" max="90"></label>
     <label>${esc(T('الكثافة عند 15°م كغم/م³','Density at 15 °C kg/m³'))}<input type="number" id="fu-c-d" value="835" step="0.1" min="400" max="1200"></label></div>
    <div class="fu-out" id="fu-c-out"></div><div class="fu-trace" id="fu-c-trace"></div><div class="fu-warn" id="fu-c-warn" hidden></div></div>
    <div class="fu-vcf"><header>${esc(T('جدول 54B: عامل التصحيح','Table 54B: correction factor'))}<small id="fu-v-sub"></small></header><table id="fu-v-tb"></table></div></div></section>
   <section><h4 class="fu-h"><span class="fu-n">3</span>${esc(T('الإعدادات العامة','General settings'))}</h4>
   <div class="fu-set"><label>${esc(T('الطن في الفواتير','Tonnes on invoices'))} <select id="fu-basis"><option value="air"${air?' selected':''}>${esc(T('في الهواء (وزن تجاري، جدول 56)','in air (commercial weight, Table 56)'))}</option><option value="vacuum"${air?'':' selected'}>${esc(T('في الفراغ (كتلة حقيقية)','in vacuum (true mass)'))}</option></select></label>
    <label>${esc(T('درجة حرارة المنتج عند القراءة (°م)','Product temperature at reading (°C)'))} <input type="number" id="fu-temp" min="-18" max="90" step="0.05" value="${esc(String(U.temp))}" style="width:90px"></label>
    <label>${esc(T('كثافة مقيسة','Measured density'))} <input type="number" id="fu-cd" min="400" max="1200" step="0.1" placeholder="${esc(T('كغم/م³','kg/m³'))}" style="width:100px"> ${esc(T('عند','at'))} <input type="number" id="fu-ct" min="-18" max="90" step="0.05" placeholder="°C" style="width:70px"> <button type="button" class="btn sm" data-fu="d15">${esc(T('← الكثافة عند 15°م (جدول 53B)','→ density at 15 °C (Table 53B)'))}</button> <b id="fu-cd-out" class="num" style="font-size:13px"></b></label></div>
   <p class="fu-note" style="margin-top:10px">${esc(T('الحرارة المعتمدة لكل سطر: عمود «درجة الحرارة» في ملف الاستيراد، وإلا حرارة النوع في الجدول، وإلا الحرارة العامة (15°م = بلا تصحيح). والكثافة المعتمدة: عمود «الكثافة عند 15°م» في ملف الاستيراد إن وُجد. الغاز المسال خارج جدول 54B ويحتاج جدول 54E (GPA TP-27) مع ضغط البخار، فيُدخَل عامله من شهادة المورّد. عند النزاع في الاستلام والتسليم تُعتمد الجداول المذكورة في عقد التوريد.','Temperature used per line: the «Temperature» column of the imported file, else the type temperature in the table, else the general temperature (15 °C = no correction). Density used: the «Density at 15 °C» column of the imported file when present. LPG is outside Table 54B and needs Table 54E (GPA TP-27) with the vapour pressure, so enter its factor from the supplier’s certificate. In custody-transfer disputes the tables named in the supply contract prevail.'))}</p></section>
   <section><h4 class="fu-h"><span class="fu-n">4</span>${esc(T('كثافات المنتجات عند 15°م','Product densities at 15 °C'))} <small>${esc(T('الشريط الأخضر حدود المواصفة الدولية، والعلامة قيمة المنتج','green band = international specification, marker = product value'))}</small></h4>
   <div class="tablewrap" style="max-height:none;min-height:0;border:1px solid var(--line);border-radius:12px"><table class="fu-tb"><thead><tr><th class="nosort">${esc(T('رمز المنتج','Product code'))}</th><th class="nosort">${esc(T('المنتج','Product'))}</th><th class="nosort">${esc(T('الكثافة كغم/م³ عند 15°م','Density kg/m³ at 15 °C'))}</th><th class="nosort">${esc(T('المواصفة (حدود الكثافة)','Specification (density limits)'))}</th><th class="nosort">${esc(T('الجدول','Table'))}</th><th class="nosort">${esc(T('درجة الحرارة °م','Temperature °C'))}</th><th class="nosort">${esc(T('المصدر','Source'))}</th></tr></thead>
   <tbody>${rows.map(r=>{const sp=r.fam&&SPEC[r.fam[0]];return `<tr${r.shown?' class="fu-fam"':''}><td class="num">${r.shown?`<b>${esc(T('نوع','Type'))}</b> <small class="muted">${esc(r.shown)}</small>`:esc(r.c)}</td><td><b>${esc(r.n||'—')}</b></td><td><input type="number" min="400" max="1200" step="0.1" data-fu-d="${esc(r.c)}" value="${esc(String(r.d))}" style="width:110px" aria-label="${esc(T('الكثافة','Density'))}">${out(r)?` <span class="fu-pill bad">${esc(T('خارج المواصفة','outside spec'))}</span>`:''}</td>
     <td><div class="fu-spec">${sp?`${esc(sp[2])} · <span>${sp[0]}–${sp[1]}</span>`:'—'}</div>${bar(r.fam,r.d)}</td><td><span class="fu-tbl">${esc(sp?(sp[3]==='54E'?'54E · GPA TP-27':'54B · 53B · 56'):'—')}</span></td>
     <td>${r.shown?(r.fam[0]==='lpg'?`<label style="display:inline-flex;gap:4px;align-items:center;font-size:12px">${esc(T('عامل 54E','54E factor'))} <input type="number" id="fu-lpg" min="0.8" max="1.2" step="0.0001" value="${esc(U.lpgVcf?String(U.lpgVcf):'')}" placeholder="${esc(T('بلا تصحيح','none'))}" style="width:90px"></label>`:`<input type="number" min="-18" max="90" step="0.05" data-fu-t="${esc(r.c)}" value="${esc(U.temps[r.c]!=null?String(U.temps[r.c]):'')}" placeholder="${esc(String(U.temp))}" style="width:80px" aria-label="${esc(T('درجة الحرارة','Temperature'))}">`):`<span class="muted">${esc(T('حسب النوع','by type'))}</span>`}</td><td>${src(r)}</td></tr>`}).join('')}
   <tr><td><input type="text" id="fu-new-c" placeholder="${esc(T('رمز جديد','new code'))}" style="width:110px"></td><td><input type="text" id="fu-new-n" placeholder="${esc(T('اسم المنتج','product name'))}"></td><td><input type="number" id="fu-new-d" min="400" max="1200" step="0.1" placeholder="835" style="width:110px"></td><td></td><td></td><td></td><td></td></tr></tbody></table></div></section>
   <div class="fu-foot"><span class="muted" style="font-size:12px;font-weight:500">${esc(T('كل تعديل يُسجَّل في سجل التدقيق بالقيم الجديدة واسم من عدّلها.','Every change is written to the audit log with the new values and who made it.'))}</span><span><span id="fu-msg" class="muted" style="font-size:12.5px;margin-inline-end:10px"></span><button type="button" class="btn primary" data-fu="save">${esc(T('حفظ الإعدادات','Save settings'))}</button></span></div>
  </div></div>`}
/* live converter + 54B reference column */
function conv(){const P=document.getElementById('fu-panel');if(!P||!document.getElementById('fu-c-out'))return;
  const fam=(document.getElementById('fu-c-p')||{}).value||'@diesel',lpg=fam==='@lpg';
  const q=parseFloat(document.getElementById('fu-c-q').value),u=document.getElementById('fu-c-u').value,t=parseFloat(document.getElementById('fu-c-t').value),d=parseFloat(document.getElementById('fu-c-d').value);
  const W=document.getElementById('fu-c-warn'),O=document.getElementById('fu-c-out'),TR=document.getElementById('fu-c-trace');
  const fail=m=>{O.innerHTML='';TR.textContent='';W.hidden=false;W.textContent=m};
  if(!(q>=0))return fail(T('أدخل الكمية','Enter a quantity'));if(!(t>=-18&&t<=90))return fail(T('درجة الحرارة بين −18 و90°م','Temperature must be between −18 and 90 °C'));if(!(d>=400&&d<=1200))return fail(T('الكثافة بين 400 و1200 كغم/م³','Density must be between 400 and 1200 kg/m³'));
  const ctl=lpg?(+U.lpgVcf>0?+U.lpgVcf:1):vcf(d,t);const md=massDensity(d),air=U.basis!=='vacuum';
  let v15,vObs;if(u==='T'){v15=q*1000/md;vObs=v15/ctl}else{vObs=u==='L'?q/1000:q;v15=vObs*ctl}
  const tAir=v15*(d-1.1)/1000,tVac=v15*d/1000;const f=FAM.find(x=>'@'+x[0]===fam),sp=f&&SPEC[f[0]];
  const tile=(k,lab,val,unit)=>`<div class="fu-t${k?' key':''}"><small>${esc(lab)}</small><b>${esc(val)}</b><em>${esc(unit)}</em></div>`;
  O.innerHTML=tile(0,T('الحجم المقروء','Observed volume'),nf(vObs*1000,0),T('لتر عند '+t+'°م','litres at '+t+' °C'))+tile(0,T('عامل التصحيح','Correction factor'),ctl.toFixed(5),lpg?T('جدول 54E (من المورّد)','Table 54E (supplier)'):'CTL · Table 54B')
   +tile(1,T('الحجم القياسي','Standard volume'),nf(v15,3),T('م³ عند 15°م','m³ at 15 °C'))+tile(0,T('اللترات القياسية','Standard litres'),nf(v15*1000,0),T('لتر عند 15°م','litres at 15 °C'))
   +tile(air,T('طن متري في الهواء','Metric tonnes in air'),nf(tAir,3),'Table 56')+tile(!air,T('طن في الفراغ','Tonnes in vacuum'),nf(tVac,3),T('كتلة حقيقية','true mass'));
  TR.textContent=u==='T'?`V15 = ${nf(q,3)} t × 1000 ÷ ${air?`(${d.toFixed(1)} − 1.1)`:d.toFixed(1)} = ${nf(v15,3)} m³   ·   Vobs = V15 ÷ CTL(${d.toFixed(1)}, ${t} °C) = ${nf(v15,3)} ÷ ${ctl.toFixed(5)} = ${nf(vObs,3)} m³`
   :`V15 = ${nf(vObs,3)} m³ × CTL(${d.toFixed(1)}, ${t} °C) ${ctl.toFixed(5)} = ${nf(v15,3)} m³   ·   t(air) = V15 × (${d.toFixed(1)} − 1.1) ÷ 1000 = ${nf(tAir,3)}`;
  const w=[];if(lpg&&!(+U.lpgVcf>0))w.push(T('الغاز المسال لا يُصحَّح بجدول 54B؛ أدخل عامل 54E من شهادة المورّد في جدول الكثافات.','LPG is not corrected with Table 54B; enter the 54E factor from the supplier’s certificate in the density table.'));
  if(!lpg&&rho60(d)<RHO_MIN)w.push(T('الكثافة أقل من مدى جدول 54B (611.16)، فلا يُطبَّق تصحيح.','Density below the Table 54B range (611.16): no correction applied.'));
  if(sp&&(d<sp[0]||d>sp[1]))w.push(T(`الكثافة خارج حدود ${sp[2]} (${sp[0]}–${sp[1]})`,`Density outside ${sp[2]} limits (${sp[0]}–${sp[1]})`));
  W.hidden=!w.length;W.textContent=w.join(' ');
  const VT=document.getElementById('fu-v-tb'),VS=document.getElementById('fu-v-sub');const temps=[0,5,10,15,20,25,30,35,40,45,50];if(Math.round(t*20)/20!==t||!temps.includes(t))temps.push(t);temps.sort((a,b)=>a-b);
  VS.textContent=lpg?T('الغاز المسال: جدول 54E خارج هذا الجدول','LPG: Table 54E, outside this table'):T(`ρ15 = ${d.toFixed(1)} كغم/م³ · ρ60 = ${rho60(d).toFixed(2)}`,`ρ15 = ${d.toFixed(1)} kg/m³ · ρ60 = ${rho60(d).toFixed(2)}`);
  VT.innerHTML=`<thead><tr><th>t °C</th><th>CTL</th><th>L15 / 1000 L</th></tr></thead><tbody>${lpg?'':temps.map(x=>{const c=vcf(d,x);return `<tr${x===t?' class="on"':''}><td>${x.toFixed(x%1?2:1)}</td><td>${c.toFixed(5)}</td><td>${nf(c*1000,1)}</td></tr>`}).join('')}</tbody>`}
document.addEventListener('input',e=>{if(!e.target.closest||!e.target.closest('#fu-panel .fu-cin'))return;if(e.target.id==='fu-c-p'){const o=e.target.selectedOptions[0];if(o&&o.dataset.d)document.getElementById('fu-c-d').value=o.dataset.d}conv()});
document.addEventListener('change',e=>{if(e.target&&e.target.id==='fu-c-p'){const o=e.target.selectedOptions[0];if(o&&o.dataset.d)document.getElementById('fu-c-d').value=o.dataset.d;conv()}});
if(typeof MutationObserver!=='undefined'&&typeof document!=='undefined'&&document.documentElement)new MutationObserver(()=>{const o=document.getElementById('fu-c-out');if(o&&!o.dataset.ready){o.dataset.ready='1';conv()}}).observe(document.documentElement,{childList:true,subtree:true});
document.addEventListener('click',e=>{const b=e.target.closest('[data-fu="d15"]');if(!b)return;const o=document.getElementById('fu-cd-out');const d=parseFloat((document.getElementById('fu-cd')||{}).value),t=parseFloat((document.getElementById('fu-ct')||{}).value);
  const r=d>0&&t>=-18&&t<=90?density15From(d,t):null;o.style.color=r?'var(--accent)':'var(--crit)';o.textContent=r?T(`= ${r.toFixed(1)} كغم/م³`,`= ${r.toFixed(1)} kg/m³`):T('أدخل كثافة بين 611 و1163 ودرجة حرارة بين −18 و90','enter a density 611–1163 and a temperature −18 to 90')});
document.addEventListener('click',async e=>{const b=e.target.closest('[data-fu="save"]');if(!b)return;const msg=document.getElementById('fu-msg');const dens={};let bad='';
  document.querySelectorAll('#fu-panel [data-fu-d]').forEach(i=>{const v=parseFloat(i.value);const c=i.dataset.fuD;const fam=c[0]==='@';const f=fam?FAM.find(x=>'@'+x[0]===c):family(c,U.names[c]);if(!i.value)return;if(!(v>=400&&v<=1200)){bad=fam?T(f[1],f[2]):c;return}
    const base=f?(fam?f[3]:(U.over['@'+f[0]]>0?U.over['@'+f[0]]:f[3])):null;if(base==null||Math.abs(base-v)>0.01)dens[c]=Math.round(v*10)/10});
  const nc=(document.getElementById('fu-new-c')||{}).value||'',nn=(document.getElementById('fu-new-n')||{}).value||'',nd=parseFloat((document.getElementById('fu-new-d')||{}).value);
  if(nc.trim()&&nd>=400&&nd<=1200){dens[nc.trim()]=Math.round(nd*10)/10;if(nn.trim())U.names[nc.trim()]=nn.trim().slice(0,80)}else if(nc.trim()){bad=nc}
  const temps={};let badT='';document.querySelectorAll('#fu-panel [data-fu-t]').forEach(i=>{if(i.value==='')return;const v=parseFloat(i.value);if(!(v>=-18&&v<=90)){badT=i.dataset.fuT;return}temps[i.dataset.fuT]=v});
  const gt=parseFloat((document.getElementById('fu-temp')||{}).value);if(!(gt>=-18&&gt<=90))badT=badT||'general';
  const basis=(document.getElementById('fu-basis')||{}).value||U.basis;const li=document.getElementById('fu-lpg');const lv=li&&li.value!==''?parseFloat(li.value):null;
  if(badT){msg.style.color='var(--crit)';msg.textContent=T('درجة الحرارة يجب أن تكون بين −18 و90°م','Temperature must be between −18 and 90 °C');return}
  if(bad){msg.style.color='var(--crit)';msg.textContent=T(`الكثافة يجب أن تكون بين 400 و1200 كغم/م³ (${bad})`,`Density must be between 400 and 1200 kg/m³ (${bad})`);return}
  if(lv!=null&&!(lv>0.8&&lv<1.2)){msg.style.color='var(--crit)';msg.textContent=T('عامل الغاز المسال يجب أن يكون بين 0.8 و1.2','The LPG factor must be between 0.8 and 1.2');return}
  b.disabled=true;msg.style.color='';msg.textContent=T('جارٍ الحفظ…','Saving…');
  try{await save(dens,temps,gt,basis,lv);msg.textContent=T('تم الحفظ، وتُعاد قراءة اللوحات بالإعدادات الجديدة','Saved; the dashboards now use the new settings')}catch(err){msg.style.color='var(--crit)';msg.textContent=T('تعذّر الحفظ: ','Could not save: ')+(err.message||'')}b.disabled=false});
if(window.I18N&&I18N.add)I18N.add({'وحدات الكميات':'Quantity units'});
window.FuelUnits={M3,kind,family,density,toM3,convert,vcf,vcf1980,ctl54B,rho60,density15From,massDensity,tempFor,isService,normalize,explain,seen,load,save,panel,state:U,families:FAM,spec:SPEC,calcVersion:CALC_VERSION};
})();
