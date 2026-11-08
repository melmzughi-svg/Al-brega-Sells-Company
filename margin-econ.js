/* ===== هامش الربح: النموذج الشامل على بيانات شركة البريقة =====
   Loaded after margin.js. Every figure is per product × depot (and its region) × month, from one Excel workbook with
   ten sheets (a template is downloadable from the page):
     reference (one document margin/econ_ref): المنتجات (code, name, category, unit, density at 15 °C),
       هيكل السعر (effective-dated: product cost, transport, Brega distribution margin, station margin, fees, final price),
       أسعار الفئات (effective-dated price per customer category: local LYD and foreign currency, discount %, commission)
     monthly (margin/econ_YYYY-MM): الاستلامات والتكلفة (per shipment: quantity, purchase value, sea freight, insurance,
       port fees, customs), تكاليف المستودعات (road transport, storage and handling), الكميات والمخزون (opening, received,
       sold, transfers in/out, closing, allowed loss %), الدعم الحكومي (subsidy per unit, due, collected, date),
       المصروفات التشغيلية (salaries, maintenance, depreciation, electricity, admin), الذمم (credit sales, overdue)
     sales by customer (margin/econs_YYYY-MM_N, split so no document passes the size limit): المبيعات حسب الزبون
   Quantities are stored in m³ at 15 °C: litres ÷ 1000 and m³ corrected with the volume correction factor at the row
   temperature (units.js, Table 54B; not for LPG, lubricants or asphalt), tonnes × 1000 ÷ density at 15 °C. Prices are
   stored per m³ (a price per litre × 1000, per tonne × density ÷ 1000).
   Formulas (shown on the page):
     net revenue = quantity × price − discounts and commissions + subsidy
     cost of goods sold = quantity × weighted average cost (landed cost of receipts rolled per depot and product)
     gross profit = net revenue − cost of goods sold − value of losses − depot costs (road transport, storage)
     margin % = gross profit ÷ net revenue × 100; margin per m³ / litre / tonne = gross profit ÷ quantity
     loss ratio = actual loss ÷ received, actual loss = opening + received + transfers in − transfers out − sold − closing
     official vs actual margin = Brega distribution margin in the price structure vs gross profit per m³
     net profit = gross profit − operating expenses allocated by quantity sold (without them the page says gross only)
   Ratios are always computed from summed totals. Read and written under the margin rules (firestore.rules). */
"use strict";
(function(){
const PMa=window.__pm,L=window.__smsLib;if(!PMa||!L){console.warn('margin-econ: margin.js missing');return}
const {readGrids,hn,colChart,bars,pct,fdi,mName,me,serialIso}=L;
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const COL='margin',REF='econ_ref',PART=2500,VIEW='margin';
const X={loaded:false,loading:null,err:'',ref:{prods:[],struct:[],segs:[]},months:new Map(),ver:0,pending:null,mc:null};

/* customer categories, shared with the debts dashboard (whichever file loads first defines them) */
if(!window.CUSTOMER_CATS){window.CUSTOMER_CATS=[{k:'station',ar:'محطة وقود',en:'Fuel station',sector:'private'},{k:'gov',ar:'جهة حكومية',en:'Government body',sector:'gov'},
  {k:'private',ar:'شركة خاصة',en:'Private company',sector:'private'},{k:'aviation',ar:'طيران',en:'Aviation',sector:'private'},{k:'industry',ar:'صناعة',en:'Industry',sector:'private'},{k:'power',ar:'كهرباء',en:'Electricity',sector:'gov'}];
  window.custCat=function(t){const s=(window.norm?norm(String(t||'')):String(t||'')).toLowerCase().trim();if(!s)return '';const K=window.CUSTOMER_CATS.find(c=>c.k===s||c.ar===s||c.en.toLowerCase()===s);if(K)return K.k;
    if(/محط|station/.test(s))return 'station';if(/كهرب|electric|power|gecol/.test(s))return 'power';if(/طير|طاير|طائر|جوي|مطار|aviat|airline|airport/.test(s))return 'aviation';
    if(/صناع|مصنع|industr|factory/.test(s))return 'industry';if(/حكوم|وزار|بلدي|هيئ|مصلح|جهاز|government|ministry|municip/.test(s))return 'gov';if(/شرك|company|ltd|corp/.test(s))return 'private';return ''}}
const catLbl=k=>{if(!k)return tr('غير مصنّف','Unclassified');const c=(window.CUSTOMER_CATS||[]).find(x=>x.k===k);return c?tr(c.ar,c.en):k};
const isFxText=v=>/^(نعم|yes|y|true|1|x|✓|اجنبي|أجنبي|foreign)$/i.test(String(v==null?'':v).trim());

/* ---------- styles ---------- */
(function(){const st=document.createElement('style');st.textContent=`
.px-tabs{display:flex;gap:4px;padding:4px;border:1px solid var(--line);border-radius:12px;background:var(--surface);width:max-content;max-width:100%;overflow-x:auto;margin:0 0 12px}
.px-tabs button{border:0;background:transparent;color:var(--muted);font:inherit;font-size:13.5px;font-weight:600;padding:7px 14px;border-radius:9px;cursor:pointer;white-space:nowrap}
.px-tabs button[aria-selected="true"]{background:var(--accent);color:#fff}
.px-tabs button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.px-badge{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;padding:3px 10px;border-radius:999px;background:rgba(255,255,255,.16);color:#fff;border:1px solid rgba(255,255,255,.28);margin-top:8px}
.px-cov{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,210px),1fr));gap:8px;margin:0 0 12px}
.px-cov div{display:grid;grid-template-columns:auto 1fr;gap:2px 8px;align-items:start;border:1px solid var(--line);border-radius:10px;padding:8px 10px;background:var(--surface);font-size:12.5px}
.px-cov i{font-style:normal;width:20px;height:20px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:700;grid-row:span 2}
.px-cov .ok i{background:color-mix(in srgb,var(--ok) 18%,transparent);color:var(--ok)}.px-cov .no i{background:var(--line);color:var(--muted)}
.px-cov b{font-size:13px}.px-cov span{color:var(--muted);line-height:1.5}
.px-wf{display:grid;gap:6px}
.px-wf .r{display:grid;grid-template-columns:minmax(120px,190px) 1fr minmax(84px,auto);gap:10px;align-items:center;font-size:13px}
.px-wf .t{position:relative;height:22px;background:color-mix(in srgb,var(--line) 45%,transparent);border-radius:5px}
.px-wf .t i{position:absolute;inset-block:3px;border-radius:4px;background:var(--sm-c1,#0f766e)}
.px-wf .t i.neg{background:var(--crit)}.px-wf .t i.pos{background:var(--ok)}.px-wf .t i.tot{background:var(--sm-c3,#a16207);inset-block:0}
.px-wf .v{text-align:end;font-variant-numeric:tabular-nums;font-weight:600;white-space:nowrap}
.px-wf .r.tot .l{font-weight:700}
.px-ps{display:grid;gap:10px}
.px-ps .r{display:grid;grid-template-columns:minmax(110px,170px) 1fr minmax(92px,auto);gap:10px;align-items:center;font-size:13px}
.px-ps .bar{position:relative;display:flex;height:20px;border-radius:5px;overflow:visible;background:color-mix(in srgb,var(--line) 40%,transparent)}
.px-ps .bar span{height:100%}.px-ps .bar span:first-child{border-start-start-radius:5px;border-end-start-radius:5px}
.px-ps .k1{background:#64748b}.px-ps .k2{background:#0ea5e9}.px-ps .k3{background:var(--sm-c3,#a16207)}.px-ps .k4{background:#14b8a6}.px-ps .k5{background:#a78bfa}.px-ps .k6{background:repeating-linear-gradient(135deg,var(--crit) 0 4px,transparent 4px 8px);opacity:.75}
.px-ps .mk{position:absolute;inset-block:-4px;width:0;border-inline-start:2px solid var(--ink)}
.px-ps .mk::after{content:attr(data-l);position:absolute;top:-16px;transform:translateX(50%);font-size:10.5px;color:var(--muted);white-space:nowrap}
.px-leg{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12px;color:var(--muted);margin-top:10px}.px-leg i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-inline-end:5px;vertical-align:-1px}
.px-line{width:100%;height:auto;display:block}
.px-line text{font-size:11px;fill:var(--muted)}.px-line .g{stroke:var(--line)}.px-line .a{stroke:var(--sm-c1,#0f766e);fill:none;stroke-width:2.4}.px-line .b{stroke:var(--crit);fill:none;stroke-width:2;stroke-dasharray:5 4}
.px-line circle.a{fill:var(--sm-c1,#0f766e);stroke:none}.px-line circle.b{fill:var(--crit);stroke:none}
.px-two{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,420px),1fr));gap:14px}
.px-flag{display:inline-block;font-size:11px;padding:0 7px;border-radius:999px;background:color-mix(in srgb,var(--crit) 14%,transparent);color:var(--crit);white-space:nowrap}
.px-flag.w{background:color-mix(in srgb,var(--warn) 18%,transparent);color:var(--warn)}.px-flag.o{background:color-mix(in srgb,var(--ok) 16%,transparent);color:var(--ok)}
.px-dual{display:grid;gap:3px;min-width:90px}.px-dual span{display:block;height:6px;border-radius:3px}.px-dual .a{background:var(--sm-c1,#0f766e)}.px-dual .b{background:var(--crit)}
.px-steps{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,220px),1fr));gap:12px;margin:14px 0;text-align:start}
.px-steps div{border:1px solid var(--line);border-radius:12px;padding:12px;background:var(--surface)}.px-steps b{display:block;margin-bottom:4px}
.px-steps em{font-style:normal;display:inline-grid;place-items:center;width:24px;height:24px;border-radius:50%;background:var(--accent);color:#fff;font-weight:700;margin-inline-end:6px}
.px-fx td:first-child{font-weight:600;white-space:nowrap}.px-fx td:last-child{font-family:inherit;color:var(--muted)}
.px-tbl td.n,.px-tbl th.n{text-align:end;font-variant-numeric:tabular-nums;white-space:nowrap}.px-tbl th[data-px-sort]{cursor:pointer;white-space:nowrap}.px-tbl th[data-px-sort][aria-sort]::after{content:" ▾";font-size:10px}.px-tbl th[data-px-sort][aria-sort="ascending"]::after{content:" ▴"}
.px-tbl .sub{display:block;font-size:11.5px;color:var(--muted);font-weight:400}
.px-seg{display:inline-flex;gap:4px;padding:3px;border:1px solid var(--line);border-radius:9px;margin-inline-start:auto}.px-seg button{border:0;background:transparent;font:inherit;font-size:12px;padding:3px 10px;border-radius:7px;cursor:pointer;color:var(--muted)}.px-seg button[aria-pressed="true"]{background:var(--accent-soft);color:var(--ink);font-weight:600}
@media print{.px-tabs,.px-seg{display:none}}
`;document.head.appendChild(st)})();

/* ---------- formatting ---------- */
const fS=v=>{if(v==null||!isFinite(v))return'—';const a=Math.abs(v);if(a>=1e6)return fn(+(v/1e6).toFixed(a>=1e8?0:2))+tr(' مليون',' M');if(a>=1e4)return fn(Math.round(v/1e3))+tr(' ألف',' K');return fn(Math.round(v))};
const f2=v=>v==null||!isFinite(v)?'—':fn(Math.round(v*100)/100);
const f3=v=>v==null||!isFinite(v)?'—':fn(Math.round(v*1000)/1000);
const f4=v=>v==null||!isFinite(v)?'—':fn(Math.round(v*10000)/10000);
const pc=v=>v==null||!isFinite(v)?'—':(Math.round(v*10)/10)+'%';
const bd=v=>`<bdi>${esc(v)}</bdi>`;
const LYD=()=>tr('د.ل','LYD'),M3=()=>tr('م³','m³');
const r2=v=>Math.round(v*100)/100,r4=v=>Math.round(v*10000)/10000,r6=v=>Math.round(v*1e6)/1e6;
const num=v=>{if(v==null||v==='')return null;if(typeof v==='number')return isFinite(v)?v:null;if(typeof v==='object')return null;
  let s=String(v).trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[٬,\s]/g,'').replace('٫','.').replace(/%$/,'');if(/^\(.*\)$/.test(s))s='-'+s.slice(1,-1);
  return /^-?\d*\.?\d+(e-?\d+)?$/i.test(s)?parseFloat(s):null};
const AR_M=['يناير','فبراير','مارس','ابريل','مايو','يونيو','يوليو','اغسطس','سبتمبر','اكتوبر','نوفمبر','ديسمبر'];
function monthOf(v){if(v==null||v==='')return '';if(typeof v==='object'&&v.serial)return serialIso(v.serial).slice(0,7);if(typeof v==='number')return v>20000&&v<80000?serialIso(v).slice(0,7):'';
  const s=String(v).trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d));let m=s.match(/^(\d{4})[-\/.](\d{1,2})/);if(m&&+m[2]>=1&&+m[2]<=12)return m[1]+'-'+m[2].padStart(2,'0');
  m=s.match(/^(\d{1,2})[-\/.](\d{4})$/);if(m&&+m[1]>=1&&+m[1]<=12)return m[2]+'-'+m[1].padStart(2,'0');
  m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);if(m){const a=+m[1],b=+m[2];const mo=a>12?b:b>12?a:b;if(mo>=1&&mo<=12)return m[3]+'-'+String(mo).padStart(2,'0')}
  const n=window.norm?norm(s):s;const i=AR_M.findIndex(x=>n.includes(window.norm?norm(x):x));const y=s.match(/\d{4}/);if(i>=0&&y)return y[0]+'-'+String(i+1).padStart(2,'0');return ''}
function dateOf(v){if(v==null||v==='')return '';if(typeof v==='object'&&v.serial)return serialIso(v.serial);if(typeof v==='number')return v>20000&&v<80000?serialIso(v):'';
  const s=String(v).trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d));let m=s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);if(m){let a=+m[1],b=+m[2];const d=a>12?a:b>12?b:a,mo=a>12?b:b>12?a:b;return `${m[3]}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`}
  const mm=monthOf(s);return mm?mm+'-01':''}
const clean=o=>{const r={};for(const k in o){const v=o[k];if(v==null||v===''||(typeof v==='number'&&!isFinite(v)))continue;r[k]=typeof v==='number'?r6(v):v}return r};
const S_=v=>String(v==null?'':(typeof v==='object'&&v.serial?serialIso(v.serial):v)).trim().slice(0,160);

/* ---------- the workbook: sheets and columns ---------- */
/* field: [key, header, aliases, type: s text · n number · m month · d date · c category · b yes/no, required] */
const SH=[
 {k:'prods',ref:1,ar:'المنتجات',en:'Products',al:['products','المنتجات والوحدات'],why:['المنتج ووحدته وكثافته عند 15°م، لتحويل الطن واللتر إلى متر مكعب','Product, unit and density at 15 °C, to convert tonnes and litres to m³'],
  f:[['p','رمز المنتج',['كود المنتج','رقم المنتج','product code','product_no','code'],'s'],['pn','اسم المنتج',['المنتج','product','product name'],'s',1],['cat','الفئة',['category','نوع المنتج'],'s'],
     ['u','وحدة القياس',['الوحدة','وحدة الأساس','unit','uom'],'s'],['d15','الكثافة عند 15°م',['الكثافة','الكثافة كغم/م3','density','density15','density at 15'],'n']]},
 {k:'struct',ref:1,ar:'هيكل السعر',en:'Price structure',al:['price structure','هيكل الاسعار','هيكل الأسعار'],why:['السعر الرسمي وتفصيله وهامش البريقة الرسمي','Official price, its build-up and the official Brega margin'],
  f:[['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],['from','تاريخ السريان',['ساري من','effective from','from'],'d'],['pu','وحدة السعر',['الوحدة','price unit'],'s'],
     ['cost','تكلفة المنتج',['تكلفة المصفاة او الاستيراد','تكلفة المصفاة','تكلفة الاستيراد','product cost'],'n'],['trans','النقل',['تكلفة النقل','transport'],'n'],['dist','هامش توزيع البريقة',['هامش التوزيع','هامش البريقة','distribution margin'],'n'],
     ['stn','هامش المحطة',['station margin'],'n'],['fees','الرسوم',['الضرائب والرسوم','رسوم للدولة','رسوم الدولة','fees'],'n'],['price','السعر النهائي',['السعر الرسمي','السعر الرسمي النهائي','السعر الرسمي المعتمد','official price','final price'],'n']]},
 {k:'segs',ref:1,ar:'أسعار الفئات',en:'Category prices',al:['segment prices','اسعار الفئات','أسعار البيع حسب الفئة','اسعار البيع حسب الفئه','أسعار البيع','أسعار البيع حسب فئة الزبون'],why:['سعر البيع لكل فئة زبائن بالدينار والعملة الأجنبية، والخصم والعمولة','Selling price per customer category, local and foreign, with discount and commission'],
  f:[['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],['seg','فئة الزبون',['الفئة','customer category','segment'],'c',1],['from','تاريخ السريان',['ساري من','effective from','from'],'d'],['pu','وحدة السعر',['price unit'],'s'],
     ['loc','السعر المحلي',['السعر بالدينار','سعر البيع للوحدة','سعر البيع','local price','price lyd'],'n'],['fx','سعر الأجانب',['السعر بالعملة الأجنبية','foreign price'],'n'],['cur','العملة',['currency'],'s'],['rate','سعر الصرف',['exchange rate'],'n'],
     ['disc','نسبة الخصم %',['الخصم %','نسبة الخصم','نسبة الخصم أو العمولة','discount %','discount'],'n'],['comm','العمولة للوحدة',['عمولة الموزع','commission per unit','commission'],'n']]},
 {k:'recv',ar:'الاستلامات والتكلفة',en:'Receipts & cost',al:['receipts cost','الشحنات','المشتريات','الاستلامات','الاستلامات وتكلفة الشراء'],why:['تكلفة الشراء والشحن والتأمين والموانئ والجمارك لكل شحنة: أساس متوسط التكلفة المرجح','Purchase, freight, insurance, port and customs per shipment: the basis of the weighted average cost'],
  f:[['m','الشهر',['month','الفترة','تاريخ الاستلام','التاريخ'],'m',1],['w','المستودع',['depot','warehouse'],'s'],['rg','المنطقة',['region'],'s'],['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],
     ['src','المصدر',['source','الجهة الموردة'],'s'],['ship','رقم الشحنة',['رقم الشحنة / الفاتورة','رقم الفاتورة','shipment','shipment no'],'s'],['q','الكمية',['الكمية المستلمة','quantity','qty'],'n',1],['u','الوحدة',['unit','uom'],'s'],['tc','درجة الحرارة',['temperature','temp'],'n'],
     ['val','قيمة الشراء',['القيمة','purchase value','value'],'n'],['fr','الشحن البحري',['freight','sea freight'],'n'],['ins','التأمين',['insurance'],'n'],['port','رسوم الموانئ',['رسوم الموانئ والجمارك','port fees'],'n'],['cus','الجمارك',['customs'],'n'],
     ['oth','تكاليف مباشرة أخرى',['other direct costs'],'n',0,1],['road','النقل البري',['road transport'],'n',0,1],['stor','التخزين والمناولة',['storage'],'n',0,1]]},
 {k:'dep',po:1,ar:'تكاليف المستودعات',en:'Depot costs',al:['depot costs','تكاليف النقل والتخزين'],why:['النقل البري والتخزين والمناولة لكل مستودع','Road transport, storage and handling per depot'],
  f:[['m','الشهر',['month'],'m',1],['w','المستودع',['depot','warehouse'],'s',1],['rg','المنطقة',['region'],'s'],['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],
     ['road','النقل البري',['نقل الشاحنات','road transport'],'n'],['stor','التخزين والمناولة',['التخزين','storage'],'n']]},
 {k:'stock',ar:'الكميات والمخزون',en:'Quantities & stock',al:['stock','الكميات','المخزون','حركة الكميات','حركة الكميات والفاقد'],why:['المخزون والمستلم والمباع والتحويلات: أساس الفاقد والكمية المباعة','Stock, received, sold and transfers: the basis of losses and sold quantity'],
  f:[['m','الشهر',['month'],'m',1],['w','المستودع',['depot','warehouse'],'s',1],['rg','المنطقة',['region'],'s'],['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],['u','الوحدة',['unit','uom'],'s'],['tc','درجة الحرارة',['temperature','temp'],'n'],
     ['op','مخزون أول المدة',['رصيد اول المدة','opening stock'],'n'],['rc','الكمية المستلمة',['المستلم','الوارد','received'],'n'],['sd','الكمية المباعة',['المباع','المبيعات','sold'],'n'],['ti','تحويلات واردة',['transfers in'],'n'],['to','تحويلات صادرة',['transfers out'],'n'],
     ['cl','مخزون آخر المدة',['رصيد اخر المدة','الرصيد الفعلي','الرصيد الفعلي (الجرد)','closing stock'],'n'],['al','الفاقد المسموح %',['نسبة الفاقد المسموح','نسبة الفاقد المسموح بها','allowed loss %'],'n']]},
 {k:'sales',ar:'المبيعات حسب الزبون',en:'Sales by customer',al:['sales','المبيعات','sales by customer'],why:['الكمية والقيمة والخصومات لكل زبون: أفضل وأضعف الزبائن والفئات','Quantity, value and discounts per customer: best and worst customers and categories'],
  f:[['m','الشهر',['month'],'m',1],['w','المستودع',['depot','warehouse'],'s'],['rg','المنطقة',['region'],'s'],['c','رمز الزبون',['رقم الحساب','رقم الزبون','customer code','account'],'s'],['cn','اسم الزبون',['الزبون','customer','customer name'],'s'],
     ['seg','فئة الزبون',['الفئة','customer category','segment'],'c'],['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],['q','الكمية',['الكمية المباعة','quantity','qty'],'n',1],['u','الوحدة',['unit','uom'],'s'],['tc','درجة الحرارة',['temperature','temp'],'n'],
     ['v','قيمة البيع',['القيمة','قيمة المبيعات قبل الخصم','قيمة المبيعات','sales value','value'],'n'],['disc','الخصومات',['الخصم','الخصومات والعمولات','discounts'],'n'],['comm','العمولات',['العمولة','commissions'],'n'],['fx','زبون أجنبي',['أجنبي','foreign'],'b']]},
 {k:'sub',ar:'الدعم الحكومي',en:'Government subsidy',al:['subsidy','الدعم','التعويض'],why:['الدعم لكل وحدة والمستحق والمحصّل؛ بدونه تظهر منتجات مدعومة بربح سالب','Subsidy per unit, due and collected; without it subsidised products show a loss'],
  f:[['m','الشهر',['month'],'m',1],['p','رمز المنتج',['كود المنتج','product code','code'],'s'],['pn','اسم المنتج',['المنتج','product'],'s'],['pu','وحدة الدعم',['وحدة السعر','unit'],'s'],['rate','الدعم للوحدة',['قيمة الدعم للوحدة','subsidy per unit'],'n'],
     ['due','المستحق',['المبلغ المستحق','الدعم المستحق','due'],'n'],['col','المحصل',['المبلغ المحصل','المحصّل','المحصل فعلا','collected'],'n'],['cd','تاريخ التحصيل',['collection date'],'d']]},
 {k:'opex',ar:'المصروفات التشغيلية',en:'Operating expenses',al:['opex','المصروفات','expenses'],why:['الرواتب والصيانة والإهلاك والكهرباء والإدارية: للانتقال من الربح الإجمالي إلى الصافي','Salaries, maintenance, depreciation, electricity, admin: from gross to net profit'],
  f:[['m','الشهر',['month'],'m',1],['w','المستودع',['المستودع / الإدارة','الإدارة','depot','warehouse'],'s'],['k','البند',['نوع المصروف','بند المصروف','item'],'s',1],['a','المبلغ',['amount'],'n',1]]},
 {k:'ar',ar:'الذمم',en:'Receivables',al:['receivables','المديونية','الديون'],why:['المبيعات الآجلة والمتأخرات لكل زبون: الربح على الورق مقابل المحصّل','Credit sales and overdue per customer: profit on paper vs collected'],
  f:[['m','الشهر',['month','حتى شهر'],'m',1],['c','رمز الزبون',['رقم الحساب','customer code','account'],'s'],['cn','اسم الزبون',['الزبون','customer'],'s'],['cr','المبيعات الآجلة',['الرصيد','الرصيد المستحق','credit sales','balance'],'n'],
     ['od','المتأخر',['الديون المتأخرة','overdue'],'n'],['days','أيام التأخر',['days overdue'],'n']]}];
const SHK=Object.fromEntries(SH.map(s=>[s.k,s]));
const MONTHLY=['recv','dep','stock','sales','sub','opex','ar'];

/* ---------- units: quantities to m³ at 15 °C, prices per m³ ---------- */
function refMap(prods){return new Map((prods||[]).map(x=>[x.p,x]))}
function dens(p,pn,RP){const x=RP.get(p);if(x&&x.d15>0)return +x.d15;const F=window.FuelUnits;if(F&&F.density){const d=F.density(p,pn||(x&&x.pn));if(d&&d.d)return d.d}return null}
function vcfOk(p,pn,RP){const x=RP.get(p)||{};const c=(window.norm?norm(String(x.cat||'')):String(x.cat||'')).toLowerCase();if(/زيوت|شحوم|اسفلت|أسفلت|غاز|lpg|lube|grease|asphalt|bitum/.test(c))return false;
  const F=window.FuelUnits;const f=F&&F.family?F.family(p,pn||x.pn):null;return !(f&&f[0]==='lpg')}
function kindOf(u){const s=String(u||'').trim();if(!s)return '';if(/برميل|bbl|barrel/i.test(s))return 'BBL';const F=window.FuelUnits;return F&&F.kind?F.kind(s):''}
/* returns {q (m³ at 15 °C), vf (factor), bad} */
function qM3(q,u,p,pn,tc,RP){const n=num(q);if(n==null)return null;const uu=String(u||'').trim()||(RP.get(p)||{}).u||'م³';const k=kindOf(uu);const d=dens(p,pn,RP);
  if(k==='M3'||k==='L'||k==='BBL'){let v=k==='L'?n/1000:k==='BBL'?n*0.158987:n;const t=num(tc);let vf=1;
    if(t!=null&&t!==15&&d&&vcfOk(p,pn,RP)&&window.FuelUnits&&FuelUnits.vcf)vf=FuelUnits.vcf(d,t);return {q:v*vf,vf}}
  /* tonnes → standard m³ per ASTM D1250 Table 56 (in air by default; Settings ← «وحدات الكميات») */
  if(k==='T'||k==='KG'){if(!d)return {q:n,bad:'dens'};const md=mD(d);return {q:k==='T'?n*1000/md:n/md,vf:1}}
  return {q:n,bad:'unit'}}
/* multiplier from a price per unit to a price per m³ (blank unit = per litre) */
function perM3(pu,p,pn,RP){const k=kindOf(String(pu||'').trim()||'لتر');if(k==='L')return 1000;if(k==='M3')return 1;if(k==='BBL')return 1/0.158987;const d=dens(p,pn,RP);if(!d)return null;if(k==='T')return mD(d)/1000;if(k==='KG')return mD(d);return null}
function mD(d){const F=window.FuelUnits;return F&&F.massDensity?F.massDensity(d):d}

/* ---------- import ---------- */
const hx=s=>hn(String(s==null?'':s));
/* the grey sample row of a data-request workbook («(مثال) بنزين», «مثال-001») is never data */
const isSample=x=>typeof x==='string'&&/^\(?\s*مثال/.test(hx(x));
/* a percentage column whose header has no «%» (Excel 5% is stored 0.05) is read as a fraction: 0.05 becomes 5 */
const PCT=['disc','al'];
function headerMap(H,def){const m={};const used=new Set();
  def.f.forEach(([k,h,al])=>{const cand=[h,...al].map(hx);let j=H.findIndex((x,i)=>!used.has(i)&&x&&cand.includes(x));if(j<0)j=H.findIndex((x,i)=>!used.has(i)&&x&&cand.some(a=>x.startsWith(a+' (')||x.startsWith(a+' ('.replace(' ',''))));if(j>=0){m[k]=j;used.add(j)}});return m}
function fitSheet(grid,def){let best=null;for(let i=0;i<Math.min(grid.length,15);i++){const H=(grid[i]||[]).map(x=>hx(typeof x==='object'&&x?'':x));const m=headerMap(H,def);
    const req=def.f.filter(f=>f[4]).every(f=>m[f[0]]!=null)&&(def.po||m.p!=null||m.pn!=null||!def.f.some(f=>f[0]==='p'));const n=Object.keys(m).length;if(req&&(!best||n>best.n))best={row:i,m,n}}return best}
function detect(grids){const out=new Map();
  grids.forEach(s=>{const nm=hx(s.name);let def=SH.find(d=>[d.ar,d.en,...d.al].map(hx).includes(nm));let fit=def&&fitSheet(s.grid,def);
    if(!fit){let b=null;SH.forEach(d=>{const f=fitSheet(s.grid,d);if(f&&f.n>=5&&(!b||f.n>b.f.n))b={d,f}});if(b){def=b.d;fit=b.f;fit.byHead=1}}
    if(def&&fit){const o=out.get(def.k)||{def,sheets:[],fits:[]};o.sheets.push(s);o.fits.push(fit);out.set(def.k,o)}});
  if(!out.size)return null;if(![...out.values()].some(o=>o.fits.some(f=>!f.byHead))&&out.size<2)return null;return out}
function parse(found,file){const P={file:file.name,sheets:{},months:new Set(),warn:new Map(),counts:{}};const w=(k,msg)=>{P.warn.set(k,(P.warn.get(k)||{msg,n:0}));P.warn.get(k).n++};
  /* reference first: densities and units of the workbook win over the saved ones */
  const prods0=found.get('prods');const RP=refMap(X.ref.prods);
  const rowsOf=k=>{const o=found.get(k);if(!o)return null;const out=[];let skip=0;
    o.sheets.forEach((s,si)=>{const fit=o.fits[si];const H=s.grid[fit.row]||[];const frac=PCT.filter(k2=>(k2==='disc'?k==='segs':k==='stock')&&fit.m[k2]!=null&&!/[%٪]/.test(String(H[fit.m[k2]]||'')));s.grid.slice(fit.row+1).forEach(r=>{if(!r||!r.some(x=>x!==''&&x!=null))return;if(r.some(isSample)){skip++;return}const v=k2=>fit.m[k2]==null?'':r[fit.m[k2]];const o2={};
      o.def.f.forEach(([k2,,, t])=>{const x=v(k2);if(t==='n')o2[k2]=num(x);else if(t==='m')o2[k2]=monthOf(x);else if(t==='d')o2[k2]=dateOf(x);else if(t==='b')o2[k2]=isFxText(x)?1:null;else o2[k2]=S_(x)});
      frac.forEach(k2=>{if(o2[k2]!=null&&o2[k2]>0&&o2[k2]<1)o2[k2]*=100});out.push(o2)})});return {rows:out,skip}};
  if(prods0){const R=rowsOf('prods');const rows=[];R.rows.forEach(o=>{const p=o.p||o.pn;if(!p)return;const d=o.d15>0&&o.d15<5?o.d15*1000:o.d15;const x=clean({p,pn:o.pn||p,cat:o.cat,u:o.u,d15:d>0&&d<2000?d:null});rows.push(x);RP.set(p,Object.assign({},RP.get(p)||{},x))});P.sheets.prods=rows}
  const pid=o=>{if(o.p)return o.p;if(!o.pn)return '';const n=window.norm?norm(o.pn):o.pn;for(const [k,x] of RP)if((window.norm?norm(x.pn||''):x.pn)===n)return k;return o.pn};
  const priceRows=(k,keys)=>{const R=rowsOf(k);if(!R)return null;const rows=[];R.rows.forEach(o=>{const p=pid(o);if(!p)return;const f=perM3(o.pu,p,o.pn,RP);if(f==null){w('pu-'+k,tr(`«${SHK[k].ar}»: وحدة سعر لا تُحوَّل (طن بلا كثافة؟)`,`“${SHK[k].en}”: a price unit that cannot be converted (tonnes without a density?)`));return}
    const x={p,pn:o.pn||(RP.get(p)||{}).pn||p,from:o.from||'2000-01-01',pu:o.pu||'لتر'};keys.forEach(kk=>{if(o[kk]!=null)x[kk]=['disc','rate'].includes(kk)?o[kk]:o[kk]*f});
    if(k==='segs'){x.seg=window.custCat(o.seg)||String(o.seg||'').slice(0,40);x.sl=String(o.seg||'').slice(0,60);x.cur=o.cur;x.disc=o.disc;x.rate=o.rate}
    if(!keys.some(kk=>o[kk]!=null))return;rows.push(clean(x))});return rows};
  const st=priceRows('struct',['cost','trans','dist','stn','fees','price']);if(st)P.sheets.struct=st;
  const sg=priceRows('segs',['loc','fx','comm']);if(sg)P.sheets.segs=sg;
  MONTHLY.forEach(k=>{const R=rowsOf(k);if(!R)return;const rows=[];let skip=0;
    R.rows.forEach(o=>{if(!o.m){skip++;return}const p=k==='opex'||k==='ar'?'':pid(o);const x={m:o.m};
      if(k==='recv'||k==='sales'||k==='stock'){if(!p){skip++;return}}
      if(k==='recv'||k==='sales'){const c=qM3(o.q,o.u,p,o.pn,o.tc,RP);if(!c){skip++;return}if(c.bad)w(c.bad+'-'+p,c.bad==='dens'?tr(`${o.pn||p}: لا توجد كثافة لتحويل الطن، أضفها في ورقة المنتجات`,`${o.pn||p}: no density to convert tonnes; add it on the Products sheet`):tr(`${o.pn||p}: الوحدة «${o.u}» لا تُحوَّل إلى م³`,`${o.pn||p}: unit “${o.u}” cannot be converted to m³`));
        Object.assign(x,{p,q:c.q,q0:num(o.q),u0:o.u||'',tc:o.tc,vf:c.vf!==1?c.vf:null})}
      if(k==='stock'){let vf=null;const cv=key=>{if(o[key]==null)return null;const c=qM3(o[key],o.u,p,o.pn,o.tc,RP);if(c&&c.bad)w(c.bad+'-'+p,tr(`${o.pn||p}: كميات المخزون بوحدة لا تُحوَّل`,`${o.pn||p}: stock quantities in a unit that cannot be converted`));if(c&&c.vf!==1)vf=c.vf;return c?c.q:null};
        Object.assign(x,{p,op:cv('op'),rc:cv('rc'),sd:cv('sd'),ti:cv('ti'),to:cv('to'),cl:cv('cl'),al:o.al,u0:o.u||'',tc:o.tc});x.vf=vf}
      if(k==='recv'){Object.assign(x,{w:o.w,rg:o.rg,src:o.src,ship:o.ship,val:o.val,fr:o.fr,ins:o.ins,port:o.port,cus:o.oth!=null?(o.cus||0)+o.oth:o.cus});
        /* road transport and storage written on the receipts sheet are depot costs of that month, depot and product */
        if(o.road!=null||o.stor!=null)(P.recvDep=P.recvDep||[]).push(clean({m:o.m,w:o.w,rg:o.rg,p,road:o.road,stor:o.stor}))}
      if(k==='sales'){const raw=String(o.seg||'');const fx=o.fx||/اجانب|أجانب|اجنبي|أجنبي|foreign/i.test(raw)?1:null;Object.assign(x,{w:o.w,rg:o.rg,c:o.c||o.cn,cn:o.cn||o.c,seg:window.custCat(raw)||(fx?'':raw.slice(0,40)),v:o.v,disc:o.disc,comm:o.comm,fx})}
      if(k==='stock'||k==='dep')Object.assign(x,{w:o.w,rg:o.rg});
      if(k==='dep'){Object.assign(x,{p:o.p||o.pn?p:'',road:o.road,stor:o.stor});if(x.road==null&&x.stor==null){skip++;return}}
      if(k==='sub'){if(!p){skip++;return}const f=o.rate!=null?perM3(o.pu,p,o.pn,RP):1;if(f==null){w('pu-sub',tr('«الدعم الحكومي»: وحدة دعم لا تُحوَّل','“Subsidy”: a unit that cannot be converted'));return}Object.assign(x,{p,rate:o.rate!=null?o.rate*f:null,pu:o.pu,due:o.due,col:o.col,cd:o.cd})}
      if(k==='opex'){if(o.a==null){skip++;return}Object.assign(x,{w:o.w,k:String(o.k||'').slice(0,60),a:o.a})}
      if(k==='ar'){if(!o.c&&!o.cn){skip++;return}Object.assign(x,{c:o.c||o.cn,cn:o.cn||o.c,cr:o.cr,od:o.od,days:o.days})}
      rows.push(clean(x));P.months.add(o.m)});
    P.sheets[k]=rows;P.counts[k]={n:rows.length,skip}});
  if(P.recvDep){if(!P.sheets.dep){P.sheets.dep=P.recvDep;P.counts.dep={n:P.recvDep.length,skip:0}}delete P.recvDep}
  ['prods','struct','segs'].forEach(k=>{if(P.sheets[k])P.counts[k]={n:P.sheets[k].length,skip:0}});
  return P}
async function tryImport(grids,file){const found=detect(grids);if(!found)return false;let P;try{P=parse(found,file)}catch(e){console.error(e);return false}
  if(!Object.keys(P.sheets).length)return false;X.pending=P;review(P);return true}
function review(P){const can=PMa.PM.canEdit;const ms=[...P.months].sort();const over=ms.filter(m=>X.months.has(m));
  const rows=SH.map(d=>{const c=P.counts[d.k];return `<tr><td class="t-strong">${esc(tr(d.ar,d.en))}</td><td class="n">${c?fn(c.n):'—'}</td><td class="n">${c&&c.skip?fn(c.skip):''}</td><td>${c?`<span class="px-flag o">${esc(tr('مقروءة','read'))}</span>`:`<span class="muted" style="font-size:12px">${esc(tr('غير موجودة في الملف','not in the file'))}</span>`}</td></tr>`}).join('');
  modal(tr('مراجعة ملف بيانات هامش الربح','Review the profit margin workbook'),`
   <div class="sm-mini"><div><b class="num">${fn(Object.keys(P.sheets).length)}</b><span>${esc(tr('ورقة مقروءة من 10','sheets read of 10'))}</span></div><div><b class="num">${ms.length?esc(ms[0]+(ms.length>1?' – '+ms[ms.length-1]:'')):'—'}</b><span>${esc(tr('الأشهر','Months'))}</span></div>
    <div><b class="num">${fn(Object.values(P.counts).reduce((a,c)=>a+c.n,0))}</b><span>${esc(tr('صف','rows'))}</span></div></div>
   <p class="muted" style="margin:6px 0;font-size:12.5px">${esc(tr(`الملف: ${P.file}. الكميات تُحوَّل إلى م³ عند 15°م والأسعار إلى سعر المتر المكعب. الأوراق الموجودة تستبدل بياناتها في الأشهر نفسها، وتبقى الأوراق الأخرى كما هي.`,`File: ${P.file}. Quantities are converted to m³ at 15 °C and prices to a price per m³. The sheets in the file replace their data for the same months; other sheets are kept.`))}</p>
   <div class="tablewrap" style="max-height:300px;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(tr('الورقة','Sheet'))}</th><th class="nosort n">${esc(tr('صفوف','Rows'))}</th><th class="nosort n">${esc(tr('متجاهلة','Skipped'))}</th><th class="nosort"></th></tr></thead><tbody>${rows}</tbody></table></div>
   ${over.length?`<div class="banner warn" style="margin-top:8px">${esc(tr(`أشهر محفوظة سابقًا ستُحدَّث: ${over.join('، ')}`,`Months already saved will be updated: ${over.join(', ')}`))}</div>`:''}
   ${P.warn.size?`<ul style="margin:8px 0 0;padding-inline-start:18px;font-size:12.5px;color:var(--warn);line-height:1.7">${[...P.warn.values()].slice(0,8).map(x=>`<li>${esc(x.msg)}${x.n>1?` (${fn(x.n)})`:''}</li>`).join('')}</ul>`:''}
   ${can?'':`<p class="muted" style="font-size:12.5px">${esc(tr('الحفظ متاح لمحرري إدارة الشؤون التجارية والتحصيلات والإدارة العامة ومدير النظام.','Saving is for editors of Commercial Affairs & Collections, the General Administration and the administrator.'))}</p>`}
   <p id="px-msg" style="font-size:12.5px;margin:6px 0 0"></p>`,
   `${can?`<button type="button" class="btn primary" data-px="save">${esc(tr('حفظ في المنظومة','Save to the system'))}</button>`:''}<button type="button" class="btn" data-px="view-local">${esc(tr('عرض دون حفظ','View without saving'))}</button><button type="button" class="btn ghost" data-act="close-dlg">${esc(tr('إلغاء','Cancel'))}</button>`)}

/* merge an import into the stored data (in memory); returns the documents to write */
function merge(P,ref,months){const keyed=(old,neu,key)=>{const ks=new Set(neu.map(key));return old.filter(x=>!ks.has(key(x))).concat(neu)};
  const R={prods:ref.prods||[],struct:ref.struct||[],segs:ref.segs||[]};
  if(P.sheets.prods)R.prods=keyed(R.prods,P.sheets.prods,x=>x.p);
  if(P.sheets.struct)R.struct=keyed(R.struct,P.sheets.struct,x=>x.p+'|'+x.from);
  if(P.sheets.segs)R.segs=keyed(R.segs,P.sheets.segs,x=>x.p+'|'+x.seg+'|'+x.from);
  const M=new Map(months);[...P.months].forEach(m=>{const o=Object.assign({m,recv:[],dep:[],stock:[],sub:[],opex:[],ar:[],sales:[]},M.get(m)||{});MONTHLY.forEach(k=>{if(P.sheets[k])o[k]=P.sheets[k].filter(r=>r.m===m).map(r=>{const y=Object.assign({},r);delete y.m;return y})});M.set(m,o)});
  return {R,M,refTouched:!!(P.sheets.prods||P.sheets.struct||P.sheets.segs)}}
async function save(P,btn){const msg=document.getElementById('px-msg');const FS=window.__fs(),m=me();const sig={updatedBy:m.email,by:String(m.name||m.email||'').slice(0,120),at:new Date().toISOString()};
  const {R,M,refTouched}=merge(P,X.ref,X.months);btn.disabled=true;msg.style.color='';msg.textContent=tr('جارٍ الحفظ…','Saving…');
  try{const W=[];if(refTouched)W.push([REF,Object.assign({prods:R.prods.slice(0,6000),struct:R.struct.slice(0,6000),segs:R.segs.slice(0,6000),src:String(P.file).slice(0,200)},sig)]);
    [...P.months].sort().forEach(mm=>{const o=M.get(mm);const sales=o.sales||[];const parts=[];for(let i=0;i<sales.length;i+=PART)parts.push(sales.slice(i,i+PART));
      const old=X.months.get(mm);const sp=P.sheets.sales?parts.length:(old&&old.sp)||0;
      W.push(['econ_'+mm,Object.assign({m:mm,recv:o.recv,dep:o.dep,stock:o.stock,sub:o.sub,opex:o.opex,ar:o.ar,sp,src:String(P.file).slice(0,200)},sig)]);
      if(P.sheets.sales)parts.forEach((s,k)=>W.push(['econs_'+mm+'_'+k,Object.assign({m:mm,k,sales:s},sig)]))});
    for(const [id,doc] of W)await FS.collection(COL).doc(id).set(doc);
    const cnt=Object.entries(P.counts).map(([k,c])=>`${SHK[k].ar} ${c.n}`).join('، ');
    try{await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:m.email,name:m.name||'',action:'استيراد بيانات هامش الربح',col:COL,id:'econ',summary:`${P.file} · ${[...P.months].sort().join('، ')} · ${cnt}`.slice(0,1500)})}catch(e){}
    X.ref=R;X.months=M;X.local=null;X.pending=null;X.ver++;dlg.close();UI().pre='all';render();toast(tr('تم حفظ بيانات هامش الربح','Profit margin data saved'))}
  catch(err){btn.disabled=false;msg.style.color='var(--crit)';msg.textContent=(err&&err.code==='permission-denied')?tr('لا تملك صلاحية الحفظ','You are not allowed to save'):(tr('تعذّر الحفظ: ','Could not save: ')+(err.message||''))}}

/* ---------- storage ---------- */
async function load(force){if(X.loading&&!force)return X.loading;
  X.loading=(async()=>{const FS=window.__fs&&window.__fs();if(!FS){X.loaded=true;return}
    try{const snap=await FS.collection(COL).get();const parts=new Map();const months=new Map();let ref={prods:[],struct:[],segs:[]};
      snap.forEach(d=>{const id=d.id,x=d.data()||{};if(id===REF)ref=Object.assign(ref,x);else if(/^econ_\d{4}-\d{2}$/.test(id))months.set(x.m,Object.assign({sales:[]},x));else if(/^econs_\d{4}-\d{2}_\d+$/.test(id)){(parts.get(x.m)||parts.set(x.m,[]).get(x.m)).push(x)}});
      parts.forEach((ps,m)=>{const M=months.get(m);if(!M)return;ps.filter(p=>p.k<(M.sp||0)).sort((a,b)=>a.k-b.k).forEach(p=>{M.sales=M.sales.concat(p.sales||[])})});
      X.ref=ref;X.months=months;X.err=''}
    catch(e){X.err=e&&e.code==='permission-denied'?'denied':(e&&e.message)||'error'}
    X.loaded=true;X.ver++})();return X.loading}

/* ---------- the model ---------- */
function bySorted(list,key){const m=new Map();(list||[]).forEach(x=>{const k=key(x);(m.get(k)||m.set(k,[]).get(k)).push(x)});m.forEach(v=>v.sort((a,b)=>String(a.from).localeCompare(String(b.from))));return m}
const at=(list,m)=>{let e=null;(list||[]).forEach(x=>{if(String(x.from||'').slice(0,7)<=m)e=x});return e};
const board=(p,d)=>{try{return PMa.PM.cfg&&PMa.priceAt?PMa.priceAt(p,d):null}catch(e){return null}};
function rateAt(m){const R=((PMa.PM.cfg||{}).rates||[]).slice().sort((a,b)=>String(a.from).localeCompare(String(b.from)));let e=null;R.forEach(x=>{if(String(x.from).slice(0,7)<=m)e=+x.rate});return e||(R[0]?+R[0].rate:null)}
function build(){const src=X.local||X;const key=X.local?'L'+X.local.ver:X.ver;if(X.mc&&X.mc.key===key)return X.mc.out;
  const ref=src.ref,RP=refMap(ref.prods),ST=bySorted(ref.struct,x=>x.p),SG=bySorted(ref.segs,x=>x.p);
  const months=[...src.months.keys()].sort();const names=new Map();(ref.prods||[]).forEach(x=>names.set(x.p,x.pn));
  const FX=new Set(((PMa.PM.cfg||{}).foreign||[]).map(String));const wRg=new Map();
  const cells=new Map();const cell=(m,w,p)=>{const k=m+'|'+(w||'')+'|'+p;let c=cells.get(k);if(!c)cells.set(k,c={m,w:w||'',p,rq:0,rv:0,sold:null,sq:0,op:null,cl:null,rcS:null,ti:0,to:0,al:null,dist:0,opex:0,lines:[]});return c};
  const pool=new Map(),subs=new Map(),opexRows=[],depRows=[],arRows=new Map(),subRows=[];
  months.forEach(m=>{const M=src.months.get(m);
    (M.recv||[]).forEach(r=>{if(r.w&&r.rg)wRg.set(r.w,r.rg);if(r.pn)names.set(r.p,names.get(r.p)||r.pn);const c=cell(m,r.w,r.p);const v=(+r.val||0)+(+r.fr||0)+(+r.ins||0)+(+r.port||0)+(+r.cus||0);c.rq+=+r.q||0;if(v>0){c.rv+=v;c.rqv=(c.rqv||0)+(+r.q||0);const pk=m+'|'+r.p,pl=pool.get(pk)||{q:0,v:0};pl.q+=+r.q||0;pl.v+=v;pool.set(pk,pl)}});
    (M.stock||[]).forEach(r=>{if(r.w&&r.rg)wRg.set(r.w,r.rg);const c=cell(m,r.w,r.p);if(r.op!=null)c.op=(c.op||0)+(+r.op);if(r.cl!=null)c.cl=(c.cl||0)+(+r.cl);if(r.rc!=null)c.rcS=(c.rcS||0)+(+r.rc);if(r.sd!=null)c.sold=(c.sold||0)+(+r.sd);c.ti+=+r.ti||0;c.to+=+r.to||0;if(r.al!=null)c.al=+r.al});
    (M.sales||[]).forEach(r=>{if(r.w&&r.rg)wRg.set(r.w,r.rg);const c=cell(m,r.w,r.p);c.lines.push(r);c.sq+=+r.q||0});
    (M.dep||[]).forEach(r=>{if(r.w&&r.rg)wRg.set(r.w,r.rg);depRows.push(Object.assign({m},r))});
    (M.opex||[]).forEach(r=>opexRows.push(Object.assign({m},r)));
    (M.sub||[]).forEach(r=>{subs.set(m+'|'+r.p,r);subRows.push(Object.assign({m},r))});
    (M.ar||[]).forEach(r=>arRows.set(m+'|'+r.c,Object.assign({m},r)))});
  /* weighted average cost per depot and product, rolled month by month */
  const state=new Map();const list=[...cells.values()].sort((a,b)=>a.m.localeCompare(b.m));
  list.forEach(c=>{const k=c.w+'|'+c.p;const s=state.get(k)||{avg:null,qty:null};const op=c.op!=null?c.op:(s.qty!=null?s.qty:0);let avg=s.avg;
    if(c.rqv>0)avg=(avg!=null&&op>0)?(op*avg+c.rv)/(op+c.rqv):c.rv/c.rqv;c.est='';
    if(avg==null){const pl=pool.get(c.m+'|'+c.p);const s0=at(ST.get(c.p),c.m);const b=board(c.p,c.m+'-15');
      if(pl&&pl.q>0){avg=pl.v/pl.q;c.est='pool'}else if(s0&&s0.cost!=null){avg=(+s0.cost)+(+s0.trans||0);c.est='struct'}else if(b&&b.cost!=null){avg=+b.cost;c.est='board'}}
    c.avg=avg;c.rc=c.rcS!=null?c.rcS:c.rq;if(c.sold==null)c.sold=c.sq;c.op0=op;
    c.loss=(c.cl!=null&&(c.op!=null||s.qty!=null))?op+c.rc+c.ti-c.to-c.sold-c.cl:null;c.allowed=c.al!=null?c.al/100*c.rc:null;
    const close=c.cl!=null?c.cl:op+c.rc+c.ti-c.to-c.sold;state.set(k,{avg:c.est?s.avg:avg,qty:close});c.rg=wRg.get(c.w)||''});
  /* depot costs and operating expenses, allocated by quantity sold */
  const alloc=(rows,pickCells,field,amt)=>rows.forEach(r=>{const cs=pickCells(r);if(!cs.length)return;const tot=cs.reduce((a,c)=>a+Math.max(0,c.sold),0);cs.forEach(c=>{c[field]+=amt(r)*(tot>0?Math.max(0,c.sold)/tot:1/cs.length)})});
  const inM=m=>list.filter(c=>c.m===m);
  alloc(depRows,r=>r.p?[cell(r.m,r.w,r.p)]:inM(r.m).filter(c=>c.w===r.w),'dist',r=>(+r.road||0)+(+r.stor||0));
  alloc(opexRows,r=>r.w?inM(r.m).filter(c=>c.w===r.w):inM(r.m),'opex',r=>+r.a||0);
  [...cells.values()].forEach(c=>{if(!c.rg)c.rg=wRg.get(c.w)||''});
  /* lines: every sold m³ is a line; costs follow the quantity */
  const priceFor=(p,seg,m,fx)=>{const Ls=SG.get(p)||[];const ok=x=>String(x.from||'').slice(0,7)<=m;
    const pickS=f=>{let e=null;Ls.forEach(x=>{if(ok(x)&&f(x))e=x});return e};
    if(fx){const e=pickS(x=>x.seg===seg&&x.fx!=null)||pickS(x=>x.fx!=null);if(e){const rt=e.rate||rateAt(m);if(rt)return {u:e.fx*rt,src:'seg',e}}}
    const e=pickS(x=>x.seg===seg&&x.loc!=null);if(e)return {u:e.loc,src:'seg',e};
    const s0=at(ST.get(p),m);if(s0&&s0.price!=null)return {u:+s0.price,src:'official'};
    const e2=pickS(x=>x.loc!=null);if(e2)return {u:e2.loc,src:'seg',e:e2};
    const b=board(p,m+'-15');if(b){if(fx&&b.fx!=null){const rt=rateAt(m);if(rt)return {u:+b.fx*rt,src:'board'}}if(b.loc!=null)return {u:+b.loc,src:'board'}}return null};
  const lines=[];let unpriced=0,unpricedQ=0,mismatch=0;
  [...cells.values()].forEach(c=>{const d=dens(c.p,names.get(c.p),RP);const s0=at(ST.get(c.p),c.m);const sb=subs.get(c.m+'|'+c.p);const subU=sb&&sb.rate!=null?+sb.rate:0;
    const lossV=c.loss!=null&&c.loss>0&&c.avg!=null?c.loss*c.avg:0;const base=Math.max(c.sold,c.sq);
    if(c.sq>c.sold+1e-6&&c.sold>0)mismatch++;
    const uC=base>0?{loss:lossV/base,dist:c.dist/base,opex:c.opex/base}:null;
    const mk=(r,q,resid)=>{const fx=!!r.fx||FX.has(String(r.c||''));const seg=r.seg||'';let gross=r.v!=null?+r.v:null,src='file',pe=null;
      if(gross==null){const pr=priceFor(c.p,seg,c.m,fx);if(pr){gross=q*pr.u;src=pr.src;pe=pr.e}else{gross=0;src='none';unpriced++;unpricedQ+=q}}
      if(!pe){const Ls=SG.get(c.p)||[];Ls.forEach(x=>{if(String(x.from||'').slice(0,7)<=c.m&&x.seg===seg)pe=x})}
      const disc=(r.disc!=null?+r.disc:(pe&&pe.disc!=null&&src!=='file'?gross*pe.disc/100:0))+(r.comm!=null?+r.comm:(pe&&pe.comm!=null?q*pe.comm:0));
      lines.push({m:c.m,w:c.w,rg:c.rg,p:c.p,c:resid?'':String(r.c||''),cn:resid?'':String(r.cn||r.c||''),seg:resid?'':seg,fx,q,t:d?q*d/1000:0,gross,disc,sub:q*subU,
        cogs:c.avg!=null?q*c.avg:0,loss:uC?q*uC.loss:0,dist:uC?q*uC.dist:0,opex:uC?q*uC.opex:0,off:s0&&s0.dist!=null?q*s0.dist:null,src,est:c.est,noCost:c.avg==null&&q>0})};
    c.lines.forEach(r=>mk(r,+r.q||0,false));
    if(c.sold>c.sq+1e-6)mk({},c.sold-c.sq,true);
    if(base<=0&&(lossV||c.dist||c.opex))lines.push({m:c.m,w:c.w,rg:c.rg,p:c.p,c:'',cn:'',seg:'',fx:false,q:0,t:0,gross:0,disc:0,sub:0,cogs:0,loss:lossV,dist:c.dist,opex:c.opex,off:null,src:'cost'})});
  const out={months,cells:[...cells.values()],lines,names,RP,ST,SG,subRows,opexRows,arRows,unpriced,unpricedQ,mismatch,
    has:{prods:(ref.prods||[]).length,struct:(ref.struct||[]).length,segs:(ref.segs||[]).length,recv:0,dep:depRows.length,stock:0,sales:0,sub:subRows.length,opex:opexRows.length,ar:arRows.size}};
  months.forEach(m=>{const M=src.months.get(m);out.has.recv+=(M.recv||[]).length;out.has.stock+=(M.stock||[]).length;out.has.sales+=(M.sales||[]).length});
  X.mc={key,out};return out}
function sum(ls){const A={q:0,t:0,gross:0,disc:0,sub:0,cogs:0,loss:0,dist:0,opex:0,off:0,offQ:0,n:0,noCost:0};
  ls.forEach(x=>{A.q+=x.q;A.t+=x.t;A.gross+=x.gross;A.disc+=x.disc;A.sub+=x.sub;A.cogs+=x.cogs;A.loss+=x.loss;A.dist+=x.dist;A.opex+=x.opex;if(x.off!=null){A.off+=x.off;A.offQ+=x.q}if(x.q>0)A.n++;if(x.noCost)A.noCost+=x.q});return fin(A)}
function fin(A){A.nr=A.gross-A.disc+A.sub;A.gp=A.nr-A.cogs-A.loss-A.dist;A.np=A.gp-A.opex;A.m=A.nr>0?A.gp*100/A.nr:null;A.pm3=A.q?A.gp/A.q:null;A.pt=A.t?A.gp/A.t:null;
  A.noSub=A.gp-A.sub;A.offM3=A.offQ?A.off/A.offQ:null;A.mk=A.cogs?A.gp*100/A.cogs:null;return A}
function group(ls,key){const m=new Map();ls.forEach(x=>{const k=key(x);(m.get(k)||m.set(k,[]).get(k)).push(x)});return [...m.entries()].map(([k,v])=>Object.assign(sum(v),{k,ls:v}))}
function lossOf(cs){let rc=0,loss=0,al=0,alN=0,n=0;cs.forEach(c=>{if(c.loss==null)return;n++;rc+=c.rc;loss+=c.loss;if(c.allowed!=null){al+=c.allowed;alN++}});return {rc,loss,al:alN?al:null,ratio:rc>0&&n?loss*100/rc:null,alR:rc>0&&alN?al*100/rc:null,n}}

/* ---------- filters ---------- */
function UI(){return S.ui.pmx||(S.ui.pmx={pre:'all',from:'',to:'',p:'',w:'',rg:'',seg:'',grp:'w',ps:'gp',pd:'desc'})}
const addM=(m,n)=>{const y=+m.slice(0,4),mo=+m.slice(5,7)-1+n;const d=new Date(Date.UTC(y,mo,1));return d.toISOString().slice(0,7)};
function range(B){const u=UI();if(!B.months.length)return null;const min=B.months[0],max=B.months[B.months.length-1];let from=min,to=max;
  if(u.pre==='month')from=to;else if(u.pre==='3')from=addM(to,-2);else if(u.pre==='12')from=addM(to,-11);else if(u.pre==='custom'){from=(u.from||min).slice(0,7);to=(u.to||max).slice(0,7);if(from>to)[from,to]=[to,from]}
  if(from<min&&u.pre!=='custom')from=min;return {from,to,min,max}}
function fl(B,from,to,noSeg){const u=UI();return B.lines.filter(x=>x.m>=from&&x.m<=to&&(!u.p||x.p===u.p)&&(!u.w||x.w===u.w)&&(!u.rg||x.rg===u.rg)&&(noSeg||!u.seg||(u.seg==='fx'?x.fx:(!x.fx&&x.seg===u.seg))))}
function fc(B,from,to){const u=UI();return B.cells.filter(c=>c.m>=from&&c.m<=to&&(!u.p||c.p===u.p)&&(!u.w||c.w===u.w)&&(!u.rg||c.rg===u.rg))}

/* ---------- charts ---------- */
function waterfall(A){const steps=[[tr('المبيعات بالأسعار','Sales at price'),A.gross,'tot'],[tr('الخصومات والعمولات','Discounts & commissions'),-A.disc],[tr('الدعم الحكومي','Government subsidy'),A.sub],[tr('الإيراد الصافي','Net revenue'),A.nr,'tot'],
  [tr('تكلفة البضاعة المباعة','Cost of goods sold'),-A.cogs],[tr('قيمة الفاقد','Value of losses'),-A.loss],[tr('النقل البري والتخزين','Road transport & storage'),-A.dist],[tr('الربح الإجمالي','Gross profit'),A.gp,'tot']];
  if(A.opex)steps.push([tr('المصروفات التشغيلية','Operating expenses'),-A.opex],[tr('الربح الصافي','Net profit'),A.np,'tot']);
  const mx=Math.max(1,A.gross,A.nr,A.gross+A.sub);let run=0;
  return `<div class="px-wf" role="img" aria-label="${esc(tr('من المبيعات إلى الربح','From sales to profit'))}">${steps.map(([l,v,t])=>{let a,b;if(t){a=Math.min(0,v);b=Math.max(0,v);run=v}else{a=Math.min(run,run+v);b=Math.max(run,run+v);run+=v}
    const s=Math.max(0,a)*100/mx,wd=Math.max(0.6,(Math.min(mx,b)-Math.max(0,a))*100/mx);
    return `<div class="r${t?' tot':''}" data-tip="${esc(l+': '+f2(v)+' '+LYD())}"><span class="l">${esc(l)}</span><span class="t"><i class="${t?'tot':v<0?'neg':'pos'}" style="inset-inline-start:${s}%;width:${v===0&&!t?0:wd}%"></i></span><span class="v ${v<0?'pm-neg':''}">${fS(v)}</span></div>`}).join('')}</div>`}
function lineChart(pts,la,lb){if(pts.length<2)return `<div class="empty" style="padding:20px">${esc(tr('يلزم شهران على الأقل لرسم الاتجاه','At least two months are needed for a trend'))}</div>`;
  const W=640,H=200,pl=38,pr=12,pt=14,pb=26;const vals=pts.flatMap(p=>[p.a,p.b]).filter(v=>v!=null&&isFinite(v));let lo=Math.min(0,...vals),hi=Math.max(1,...vals);const pad=(hi-lo)*.1;hi+=pad;lo=lo<0?lo-pad:0;
  const rtl=document.documentElement.dir!=='ltr';const x=i=>{const t=pl+(W-pl-pr)*i/(pts.length-1);return rtl?W-t+pl-pr:t};const y=v=>pt+(H-pt-pb)*(1-(v-lo)/(hi-lo));
  const path=k=>{let d='',on=false;pts.forEach((p,i)=>{const v=p[k];if(v==null||!isFinite(v)){on=false;return}d+=(on?'L':'M')+x(i).toFixed(1)+' '+y(v).toFixed(1);on=true});return d};
  const grid=[0,.5,1].map(f=>{const v=lo+(hi-lo)*f;return `<line class="g" x1="${pl}" x2="${W-pr}" y1="${y(v)}" y2="${y(v)}"/><text x="${rtl?W-4:4}" y="${y(v)+4}" text-anchor="${rtl?'end':'start'}">${Math.round(v*10)/10}%</text>`}).join('');
  const step=Math.ceil(pts.length/8);const lab=pts.map((p,i)=>i%step===0||i===pts.length-1?`<text x="${x(i)}" y="${H-6}" text-anchor="middle">${esc(p.l)}</text>`:'').join('');
  const dots=k=>pts.map((p,i)=>p[k]==null||!isFinite(p[k])?'':`<circle class="${k}" cx="${x(i)}" cy="${y(p[k])}" r="3"><title>${esc(p.l+': '+(k==='a'?la:lb)+' '+pc(p[k]))}</title></circle>`).join('');
  return `<svg class="px-line" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(la+' · '+lb)}">${grid}<path class="a" d="${path('a')}"/><path class="b" d="${path('b')}"/>${dots('a')}${dots('b')}${lab}</svg>
   <div class="px-leg"><span><i style="background:var(--sm-c1,#0f766e)"></i>${esc(la)}</span><span><i style="background:var(--crit)"></i>${esc(lb)}</span></div>`}
function structure(B,ps,R){const rows=ps.map(p=>{const s=at(B.ST.get(p.k),R.to);if(!s)return null;const parts=[['k1',s.cost],['k2',s.trans],['k3',s.dist],['k4',s.stn],['k5',s.fees]].map(([k,v])=>[k,v==null?0:+v]);
    const tot=parts.reduce((a,x)=>a+x[1],0);const price=s.price!=null?+s.price:tot;const gap=Math.max(0,tot-price);return {p,s,parts,tot,price,gap,cost:p.q?(p.cogs/p.q):null}}).filter(Boolean);
  if(!rows.length)return `<div class="empty" style="padding:20px">${esc(tr('استورد ورقة «هيكل السعر» لعرض تفصيل السعر الرسمي.','Import the “Price structure” sheet to show the official price build-up.'))}</div>`;
  const mx=Math.max(...rows.map(r=>Math.max(r.tot,r.price,r.cost||0)))||1;const L1=v=>f4(v/1000);
  const names=[['k1',tr('تكلفة المنتج','Product cost')],['k2',tr('النقل','Transport')],['k3',tr('هامش توزيع البريقة','Brega distribution margin')],['k4',tr('هامش المحطة','Station margin')],['k5',tr('الرسوم','Fees')],['k6',tr('الفرق المدعوم','Subsidised gap')]];
  return `<div class="px-ps">${rows.map(r=>`<div class="r"><span class="t-strong">${bd(r.p.pn)}<span class="sub muted" style="display:block;font-size:11.5px;font-weight:400">${esc(tr('ساري من','since'))} ${esc(String(r.s.from||'').slice(0,10))}</span></span>
    <span class="bar">${r.parts.map(([k,v])=>v>0?`<span class="${k}" style="width:${v*100/mx}%" data-tip="${esc((names.find(n=>n[0]===k)||[])[1]+': '+L1(v)+' '+LYD()+' '+tr('للتر','per litre'))}"></span>`:'').join('')}
     <i class="mk" style="inset-inline-start:${Math.min(100,r.price*100/mx)}%" data-l="${esc(tr('السعر','price'))}"></i>${r.cost!=null?`<i class="mk" style="inset-inline-start:${Math.min(100,r.cost*100/mx)}%;border-color:var(--crit);border-inline-start-style:dashed" data-l="${esc(tr('التكلفة الفعلية','actual cost'))}"></i>`:''}</span>
    <span class="n" style="text-align:end;font-variant-numeric:tabular-nums"><b>${L1(r.price)}</b> <small class="muted">${esc(tr('د.ل/لتر','LYD/L'))}</small>${r.gap>0?`<span class="px-flag" style="display:block;margin-top:2px">${esc(tr('مدعوم','subsidised'))} ${L1(r.gap)}</span>`:''}</span></div>`).join('')}</div>
   <div class="px-leg">${names.slice(0,5).map(([k,l])=>`<span><i class="${k}" style="${{k1:'background:#64748b',k2:'background:#0ea5e9',k3:'background:var(--sm-c3,#a16207)',k4:'background:#14b8a6',k5:'background:#a78bfa'}[k]}"></i>${esc(l)}</span>`).join('')}<span><i style="background:var(--ink);width:2px"></i>${esc(tr('السعر الرسمي','official price'))}</span><span><i style="background:var(--crit);width:2px"></i>${esc(tr('التكلفة الفعلية المحققة','actual realised cost'))}</span></div>`}

/* ---------- the page ---------- */
const kpi=(l,v,s,cls)=>anKpi(l,v,s||'',cls||'');
function covStrip(B){return `<div class="px-cov">${SH.map(d=>{const n=B.has[d.k]||0;return `<div class="${n?'ok':'no'}" title="${esc(tr(d.why[0],d.why[1]))}"><i>${n?'✓':'–'}</i><b>${esc(tr(d.ar,d.en))}</b><span>${n?esc(tr(`${fn(n)} صف`,`${fn(n)} rows`)):esc(tr('غير مستوردة · ','not imported · ')+tr(d.why[0],d.why[1]))}</span></div>`}).join('')}</div>`}
/* why this account cannot read the saved financial data: the role, department and sign-in method as the rules see them */
/* before the reason is shown, read what the rules use: whether administrators must sign in with Google (config/security)
   and how this session signed in (the ID token), so an administrator who used a password is told so instead of «sign in again» */
window.__denyPrep=async function(){try{const fs=window.__fs&&window.__fs();if(fs){const s=await fs.doc('config/security').get();window.__adminGoogle=!!(s.exists&&s.data().adminGoogle)}}catch(e){}
  try{const a=window.__auth&&window.__auth();const u=a&&a.currentUser;if(u&&u.getIdTokenResult){const t=await u.getIdTokenResult();window.__signInProvider=(t&&t.signInProvider)||''}}catch(e){}};
window.__denyWhy=function(){const m=me()||{};const a=window.__auth&&window.__auth();const u=a&&a.currentUser;
  const pv=u?(u.providerData||[]).map(x=>x&&x.providerId):[];const google=window.__signInProvider?window.__signInProvider==='google.com':(pv.includes('google.com')&&!pv.includes('password'));
  const unit=m.unit?(((typeof ORG_BY!=='undefined'&&ORG_BY.get(m.unit))||{}).title||m.unit):tr('بلا إدارة','no department');
  const role=(window.__roleAr||{})[m.role]||m.role||tr('بلا صلاحية','no role');
  if(u&&u.emailVerified===false)return tr(`بريد الحساب ${m.email||''} غير مؤكد، والمنظومة لا تعطي أي صلاحية قبل تأكيده. افتح رسالة التأكيد في بريدك ثم ادخل من جديد.`,`The account e-mail ${m.email||''} is not verified; no permission applies until it is. Open the verification e-mail, then sign in again.`);
  if(m.role==='admin'&&window.__adminGoogle&&!google)return tr('حسابك مدير النظام، لكنك دخلت بكلمة المرور، وصلاحيات المدير لا تعمل إلا بالدخول بحساب Google. اخرج ثم ادخل بزر «الدخول بحساب Google».','Your account is the administrator, but you signed in with a password and administrator rights only work after signing in with Google. Sign out, then use «Sign in with Google».');
  if(m.role==='admin'&&!google)return tr('حسابك مدير النظام لكنك دخلت بكلمة المرور. إن كان خيار «دخول المدير بحساب Google» مفعّلًا فلا تعمل صلاحيات المدير إلا بالدخول بحساب Google: اخرج ثم ادخل بزر «الدخول بحساب Google».','Your account is the administrator but you signed in with a password. If «administrators sign in with Google» is on, administrator rights only work after signing in with Google: sign out, then use «Sign in with Google».');
  if(m.role==='admin')return tr('حسابك مدير النظام ودخلت بحساب Google، ومع ذلك رفضت قواعد البيانات القراءة. اخرج ثم ادخل من جديد؛ وإن تكرر فتأكد في Firebase أن وثيقة roles باسم بريدك تحمل role = admin بالحروف الصغيرة نفسها لبريد الدخول.','Your account is the administrator and you signed in with Google, yet the database rules refused the read. Sign out and in again; if it persists, check in Firebase that the roles document named after your e-mail (same lower-case spelling) has role = admin.');
  return tr(`حسابك ${m.email||''}: الصلاحية «${role}»، والإدارة «${unit}». تُمنح من حساب مدير النظام: الإعدادات ← المستخدمون والصلاحيات ← الإدارة «إدارة الشؤون التجارية والتحصيلات» (أو الصلاحية «مدير النظام»).`,`Your account ${m.email||''}: role «${role}», department «${unit}». Granted from the administrator account: Settings → Users and permissions → department «Commercial Affairs & Collections» (or role «Administrator»).`)};
function deniedLocal(){return `<div class="panel sm-empty px-empty">${I.pm.replace('<svg','<svg style="width:52px;height:52px;color:var(--accent)"')}<h3>${esc(tr('هذه اللوحة لإدارة الشؤون التجارية والتحصيلات','This dashboard belongs to Commercial Affairs & Collections'))}</h3>
   <p style="max-width:720px;margin-inline:auto">${esc(tr('لا تملك صلاحية الاطلاع على بيانات الربحية المحفوظة. يمكنك استيراد ملف وتحليله مؤقتًا على جهازك دون حفظه.','You cannot view the saved profitability data. You can still import a file and analyse it on this device without saving it.'))}</p>
   <p class="px-why" style="max-width:720px;margin-inline:auto;font-size:13px"><b>${esc(tr('السبب: ','Why: '))}</b>${esc(window.__denyWhy())}</p>
   <div class="cx-links" style="justify-content:center"><button class="btn primary" data-pm="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button></div></div>`}
function onboarding(){return `<div class="panel sm-empty px-empty">${I.pm.replace('<svg','<svg style="width:52px;height:52px;color:var(--accent)"')}<h3>${esc(tr('الربحية الشاملة على بيانات البريقة','Full profitability on Brega’s data'))}</h3>
   <p style="max-width:720px;margin-inline:auto">${esc(tr('كل رقم لكل منتج ولكل مستودع أو منطقة ولكل شهر: الإيراد الصافي بعد الخصومات والدعم، وتكلفة البضاعة بمتوسط التكلفة المرجح، والفاقد، والهامش الرسمي مقابل الفعلي، والربح الإجمالي والصافي.','Every figure per product, depot or region and month: net revenue after discounts and subsidy, cost of goods at weighted average cost, losses, official vs actual margin, gross and net profit.'))}</p>
   <div class="px-steps"><div><b><em>1</em>${esc(tr('نزّل القالب','Download the template'))}</b><span class="muted">${esc(tr('عشر أوراق مرتبة حسب الغرض، مع ورقة «اقرأني» تشرح كل عمود.','Ten sheets by purpose, with a “Read me” sheet explaining every column.'))}</span></div>
    <div><b><em>2</em>${esc(tr('املأ ما يتوفر','Fill what you have'))}</b><span class="muted">${esc(tr('يكفي للبداية: المنتجات وهيكل السعر والكميات والاستلامات. كل ورقة إضافية تفتح مؤشرًا جديدًا.','Enough to start: products, price structure, quantities and receipts. Each extra sheet unlocks more indicators.'))}</span></div>
    <div><b><em>3</em>${esc(tr('استورد الملف','Import the file'))}</b><span class="muted">${esc(tr('تُراجع الأوراق قبل الحفظ، وتُحوَّل الكميات إلى م³ عند 15°م.','Sheets are reviewed before saving; quantities are converted to m³ at 15 °C.'))}</span></div></div>
   <div style="display:flex;gap:8px;justify-content:center;flex-wrap:wrap"><button class="btn primary" data-px="template">${I.down}${esc(tr('تنزيل قالب البيانات','Download the data template'))}</button><button class="btn" data-pm="import">${I.up}${esc(tr('استيراد ملف Excel','Import Excel file'))}</button><button class="btn ghost" data-pm-tab="prices">${esc(tr('الوضع السريع من تقرير المبيعات','Quick mode from the sales report'))}</button></div></div>
  <section class="panel"><h3 class="section-title">${esc(tr('أوراق القالب وما يفتحه كل منها','Template sheets and what each unlocks'))}</h3>${covStrip({has:{}})}</section>`}
/* the page while Brega's selling prices are missing: what can be measured now, and what unlocks the margin */
function waitBody(B,A,P,G,LS,u,lossCell){const cpm=x=>x.q?x.cogs/x.q:null;
  const need=[[tr('أسعار الفئات','Category prices'),tr('رمز المنتج، فئة الزبون، تاريخ السريان، وحدة السعر، السعر المحلي (د.ل)، سعر الأجانب والعملة وسعر الصرف، نسبة الخصم %، العمولة للوحدة','product code, customer category, effective date, price unit, local price (LYD), foreign price, currency and rate, discount %, commission per unit')],
    [tr('أو: قيمة البيع في ورقة المبيعات','Or: sales value on the sales sheet'),tr('«قيمة البيع» و«الخصومات» و«العمولات» لكل سطر (زبون × منتج × مستودع × شهر)','“Sales value”, “Discounts” and “Commissions” on each line (customer × product × depot × month)')],
    [tr('هيكل السعر','Price structure'),tr('رمز المنتج، تاريخ السريان، وحدة السعر، تكلفة المنتج، النقل، هامش توزيع البريقة، هامش المحطة، الرسوم، السعر النهائي: للهامش الرسمي مقابل الفعلي','product code, effective date, price unit, product cost, transport, Brega distribution margin, station margin, fees, final price: for official vs actual margin')],
    [tr('الدعم الحكومي','Government subsidy'),tr('الشهر، رمز المنتج، وحدة الدعم، الدعم للوحدة، المستحق، المحصّل: بدونه تظهر المنتجات المدعومة بخسارة','month, product code, subsidy unit, subsidy per unit, due, collected: without it subsidised products show a loss')]];
  return `<section class="panel an-wide" style="border-inline-start:4px solid var(--warn)"><h3 class="section-title">${esc(tr('بانتظار أسعار المنتجات من شركة البريقة','Awaiting product prices from Brega'))}</h3>
    <p style="margin:0 0 10px;line-height:1.8">${esc(tr('الكميات والتكلفة والفاقد مستوردة، لكن لا يوجد سعر بيع ولا قيمة بيع لأي سطر مبيعات في هذه الفترة، فلا يُحسب الإيراد ولا الربح ولا الهامش بعد. تظهر هنا المؤشرات الممكنة الآن، ويكتمل الهامش تلقائيًا عند استيراد إحدى الأوراق التالية (ورقة واحدة تكفي، ولا حاجة لإعادة الأوراق الأخرى):','Quantities, cost and losses are imported, but no sales line in this period has a selling price or a sales value, so revenue, profit and margin cannot be computed yet. The measures available now are shown here; the margin completes itself once one of these sheets is imported (one sheet is enough, the others need not be re-imported):'))}</p>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-fx"><thead><tr><th class="nosort">${esc(tr('الورقة','Sheet'))}</th><th class="nosort">${esc(tr('الأعمدة المطلوبة','Columns needed'))}</th></tr></thead><tbody>${need.map(([a,b])=>`<tr><td>${esc(a)}</td><td>${esc(b)}</td></tr>`).join('')}</tbody></table></div>
    <p class="sm-note" style="margin:8px 0 0">${esc(tr('القالب الجاهز: زر «قالب بيانات البريقة». الأسعار بالدينار لكل لتر افتراضيًا، أو لكل م³ أو طن حسب «وحدة السعر».','Ready template: the “Brega data template” button. Prices are LYD per litre by default, or per m³ or tonne as set in “Price unit”.'))}</p></section>
  <div class="kpis an-kpis">
   ${kpi(tr('الكمية المباعة','Quantity sold'),f3(A.q),esc(tr(`م³ عند 15°م · ${f3(A.q*1000)} لتر`,`m³ at 15 °C · ${f3(A.q*1000)} litres`)))}
   ${kpi(tr('تكلفة البضاعة المباعة','Cost of goods sold'),fS(A.cogs),esc(tr('الكمية × متوسط التكلفة المرجح','quantity × weighted average cost')))}
   ${kpi(tr('تكلفة البضاعة لكل م³','Cost per m³'),f2(cpm(A)),esc(tr(`د.ل · ${f4(cpm(A)==null?null:cpm(A)/1000)} للتر`,`LYD · ${f4(cpm(A)==null?null:cpm(A)/1000)} per litre`)))}
   ${kpi(tr('نسبة الفاقد','Loss ratio'),pc(LS.ratio),esc(LS.ratio==null?tr('يلزم مخزون أول وآخر المدة','needs opening and closing stock'):tr(`${f3(LS.loss)} م³ من ${f3(LS.rc)} مستلمة${LS.alR!=null?` · المسموح ${pc(LS.alR)}`:''}`,`${f3(LS.loss)} of ${f3(LS.rc)} m³ received${LS.alR!=null?` · allowed ${pc(LS.alR)}`:''}`)),LS.ratio!=null&&LS.alR!=null?(LS.ratio>LS.alR?'crit':'ok'):'')}
   ${kpi(tr('قيمة الفاقد','Value of losses'),fS(A.loss),esc(tr('الفاقد × متوسط التكلفة','losses × average cost')))}
   ${kpi(tr('النقل البري والتخزين','Road transport and storage'),fS(A.dist),esc(tr('تكاليف المستودعات','depot costs')))}
   ${kpi(tr('هامش الربح','Profit margin'),'—',esc(tr('بانتظار أسعار المنتجات','awaiting product prices')))}
  </div>
  <div class="an-grid sm-root pm-root">
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('التكلفة والفاقد حسب المنتج','Cost and losses by product'))}</h3>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('المباع م³','Sold m³'))}</th><th class="nosort n">${esc(tr('تكلفة البضاعة','COGS'))}</th><th class="nosort n">${esc(tr('التكلفة لكل م³','Cost per m³'))}</th><th class="nosort n">${esc(tr('فاقد ونقل وتخزين','Losses, transport, storage'))}</th><th class="nosort n">${esc(tr('الفاقد','Losses'))}</th></tr></thead>
    <tbody>${P.map(p=>`<tr><td class="t-strong">${bd(p.pn)} <small class="muted num">${esc(p.k)}</small></td><td class="n">${f3(p.q)}</td><td class="n">${f2(p.cogs)}</td><td class="n">${f2(cpm(p))}<span class="sub">${f4(cpm(p)==null?null:cpm(p)/1000)} ${esc(tr('للتر','per litre'))}</span></td><td class="n">${f2(p.loss+p.dist)}</td><td class="n">${lossCell(p.L)}</td></tr>`).join('')}</tbody></table></div></section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('الفاقد حسب المستودع','Losses by depot'))}</h3>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(tr('المستودع','Depot'))}</th><th class="nosort n">${esc(tr('المباع م³','Sold m³'))}</th><th class="nosort n">${esc(tr('تكلفة البضاعة','COGS'))}</th><th class="nosort n">${esc(tr('التكلفة لكل م³','Cost per m³'))}</th><th class="nosort n">${esc(tr('الفاقد','Losses'))}</th></tr></thead>
    <tbody>${G.map(g=>`<tr><td class="t-strong">${bd(g.k)}</td><td class="n">${f3(g.q)}</td><td class="n">${f2(g.cogs)}</td><td class="n">${f2(cpm(g))}</td><td class="n">${lossCell(g.L)}</td></tr>`).join('')}</tbody></table></div></section>
  </div>`}
function view(){if(!X.loaded){load().then(()=>{if(S.view===VIEW)render()});return `<div class="panel muted">${esc(tr('جارٍ تحميل بيانات الربحية…','Loading profitability data…'))}</div>`}
  if((X.err==='denied'||PMa.PM.err==='denied')&&!X.local)return deniedLocal();
  const B=build();if(!B.months.length)return onboarding();
  const u=UI(),R=range(B),c=PMa.PM.cfg||{target:15,min:5};
  const cur=fl(B,R.from,R.to),A=sum(cur),cs=fc(B,R.from,R.to),LS=lossOf(cs);
  const len=(+R.to.slice(0,4)-+R.from.slice(0,4))*12+(+R.to.slice(5,7)-+R.from.slice(5,7))+1;const pF=addM(R.from,-len),pT=addM(R.from,-1);const prev=fl(B,pF,pT);const PB=prev.length?sum(prev):null;
  /* no selling price and no sales value for anything in the period: profit cannot be computed yet, so the page shows
     quantities, cost and losses and says which sheets unlock the margin (instead of a «loss» equal to the cost) */
  const wait=A.q>0&&!(A.gross>0)&&!(A.sub>0);
  const gross=!B.has.opex;const head=gross?A.gp:A.np;const dH=PB?(gross?PB.gp:PB.np):null;const dPct=dH?((head-dH)*100/Math.abs(dH)):null;
  /* subsidy due vs collected */
  const subR=B.subRows.filter(r=>r.m>=R.from&&r.m<=R.to&&(!u.p||r.p===u.p));const qP=new Map();fl(B,R.from,R.to,true).forEach(x=>qP.set(x.m+'|'+x.p,(qP.get(x.m+'|'+x.p)||0)+x.q));
  const subP=new Map();subR.forEach(r=>{const o=subP.get(r.p)||{p:r.p,due:0,col:0,q:0,rateQ:0,last:''};const q=qP.get(r.m+'|'+r.p)||0;const due=r.due!=null?+r.due:(r.rate!=null?q*r.rate:0);o.due+=due;o.col+=+r.col||0;o.q+=q;if(r.rate!=null)o.rateQ+=q*r.rate;if(r.cd&&r.cd>o.last)o.last=r.cd;subP.set(r.p,o)});
  const SB=[...subP.values()].reduce((a,o)=>({due:a.due+o.due,col:a.col+o.col}),{due:0,col:0});const subLast=[...subP.values()].reduce((a,o)=>o.last>a?o.last:a,'');
  /* receivables at the end of the period */
  const arM=[...new Set([...B.arRows.values()].map(r=>r.m))].filter(m=>m<=R.to).sort().pop();const AR=arM?[...B.arRows.values()].filter(r=>r.m===arM):[];const arOd=AR.reduce((a,r)=>a+(+r.od||0),0),arCr=AR.reduce((a,r)=>a+(+r.cr||0),0);const arBy=new Map(AR.map(r=>[String(r.c),r]));
  /* groups */
  const nm=p=>B.names.get(p)||p;
  const P=group(cur,x=>x.p).map(g=>Object.assign(g,{pn:nm(g.k),L:lossOf(cs.filter(c=>c.p===g.k))}));
  const sortK={gp:x=>x.gp,m:x=>x.m==null?-1e9:x.m,q:x=>x.q,nr:x=>x.nr,pm3:x=>x.pm3==null?-1e9:x.pm3,pn:x=>x.pn,loss:x=>x.L.ratio==null?-1e9:x.L.ratio}[u.ps]||(x=>x.gp);
  P.sort((a,b)=>{const x=sortK(a),y=sortK(b);const r=typeof x==='string'?x.localeCompare(y,'ar'):x-y;return u.pd==='asc'?r:-r});
  const G=group(cur.filter(x=>u.grp==='rg'?x.rg:x.w),x=>u.grp==='rg'?x.rg:x.w).map(g=>Object.assign(g,{L:lossOf(cs.filter(c=>(u.grp==='rg'?c.rg:c.w)===g.k))})).sort((a,b)=>b.gp-a.gp);
  const SG=group(cur.filter(x=>x.c||x.seg||x.fx),x=>x.fx?'@fx':x.seg).sort((a,b)=>b.gp-a.gp);
  const C=group(cur.filter(x=>x.c),x=>x.c).map(g=>Object.assign(g,{cn:g.ls[0].cn,seg:g.ls[0].fx?'@fx':g.ls[0].seg,ar:arBy.get(String(g.k))}));
  const top=C.slice().sort((a,b)=>b.gp-a.gp).slice(0,10),bot=C.slice().sort((a,b)=>a.gp-b.gp).slice(0,10);
  const months=[];for(let m=R.from;m<=R.to;m=addM(m,1))months.push(m);
  const Mo=new Map(group(cur,x=>x.m).map(g=>[g.k,g]));const pts=months.map(m=>{const g=Mo.get(m);return {l:mName(+m.slice(5,7)-1).slice(0,3)+' '+m.slice(2,4),y:g?(gross?g.gp:g.np):0,tip:`${mName(+m.slice(5,7)-1)} ${m.slice(0,4)} · ${tr('ربح إجمالي','gross profit')} ${f2(g?g.gp:0)} ${LYD()} · ${tr('هامش','margin')} ${pc(g?g.m:null)}`}});
  const lpts=months.map(m=>{const g=Mo.get(m);const Lm=lossOf(cs.filter(c=>c.m===m));return {l:mName(+m.slice(5,7)-1).slice(0,3)+' '+m.slice(2,4),a:g?g.m:null,b:Lm.ratio}});
  const opexK=group(B.opexRows.filter(r=>r.m>=R.from&&r.m<=R.to&&(!u.w||r.w===u.w)).map(r=>({k:r.k,q:0,t:0,gross:0,disc:0,sub:0,cogs:0,loss:0,dist:0,opex:+r.a||0})),x=>x.k||tr('أخرى','Other')).sort((a,b)=>b.opex-a.opex);
  const negNoSub=P.filter(p=>p.noSub<0&&p.gp>=0);const est=cur.filter(x=>x.est&&x.q>0).reduce((a,x)=>a+x.q,0);
  const status=m=>m==null?['p-neutral',tr('بلا إيراد','No revenue')]:m<0?['p-crit',tr('خسارة','Loss')]:m<+c.min?['p-crit',tr('تحت الحد الأدنى','Below minimum')]:m<+c.target?['p-warn',tr('تحت المستهدف','Below target')]:['p-ok',tr('ضمن المستهدف','On target')];
  /* reading */
  const ins=[];const H=gross?tr('الربح الإجمالي','Gross profit'):tr('الربح الصافي','Net profit');
  ins.push([dPct==null?'':dPct>=0?'o':'w',dPct==null?'i':dPct>=0?'▲':'▼',tr(`${H} <b>${fS(head)}</b> د.ل من إيراد صافٍ <b>${fS(A.nr)}</b> د.ل، أي هامش إجمالي <b>${pc(A.m)}</b> و<b>${f2(A.pm3)}</b> د.ل لكل م³ (${f4(A.pm3==null?null:A.pm3/1000)} للتر)${dPct!=null?`، ${dPct>=0?'بزيادة':'بانخفاض'} <b>${Math.abs(Math.round(dPct*10)/10)}%</b> عن الفترة السابقة المماثلة`:''}.`,
    `${H} <b>${fS(head)}</b> LYD on net revenue of <b>${fS(A.nr)}</b> LYD: a <b>${pc(A.m)}</b> gross margin and <b>${f2(A.pm3)}</b> LYD per m³ (${f4(A.pm3==null?null:A.pm3/1000)} per litre)${dPct!=null?`, ${dPct>=0?'up':'down'} <b>${Math.abs(Math.round(dPct*10)/10)}%</b> on the previous period`:''}.`)]);
  if(A.sub>0)ins.push([A.noSub<0?'c':'w','د',tr(`الدعم الحكومي <b>${fS(A.sub)}</b> د.ل يمثل <b>${pct(A.sub,A.nr)}%</b> من الإيراد الصافي. بدونه يصبح الربح الإجمالي <b>${fS(A.noSub)}</b> د.ل${negNoSub.length?`، وتنقلب ${negNoSub.length} منتجات إلى خسارة: ${negNoSub.slice(0,3).map(p=>bd(p.pn)).join('، ')}`:''}.`,
    `Government subsidy of <b>${fS(A.sub)}</b> LYD is <b>${pct(A.sub,A.nr)}%</b> of net revenue. Without it gross profit would be <b>${fS(A.noSub)}</b> LYD${negNoSub.length?`, and ${negNoSub.length} products turn to a loss: ${negNoSub.slice(0,3).map(p=>bd(p.pn)).join(', ')}`:''}.`)]);
  else if(!B.has.sub)ins.push(['w','د',tr('لم تُستورد ورقة «الدعم الحكومي»؛ المنتجات المدعومة تظهر بربح أقل من الحقيقي أو بخسارة لأن سعر بيعها أقل من تكلفتها.','The “Government subsidy” sheet is not imported; subsidised products show a lower profit or a loss because their selling price is below cost.')]);
  const rk=P.filter(p=>p.nr>0&&p.m!=null);if(rk.length>1){const b=rk.slice().sort((a,b)=>b.m-a.m);ins.push(['','1',tr(`أعلى هامش: <b>${bd(b[0].pn)}</b> (${pc(b[0].m)}، ${f2(b[0].pm3)} د.ل/م³). أضعف هامش: <b>${bd(b[b.length-1].pn)}</b> (${pc(b[b.length-1].m)}).`,`Highest margin: <b>${bd(b[0].pn)}</b> (${pc(b[0].m)}, ${f2(b[0].pm3)} LYD/m³). Lowest: <b>${bd(b[b.length-1].pn)}</b> (${pc(b[b.length-1].m)}).`)])}
  if(LS.ratio!=null){const over=G.filter(g=>g.L.ratio!=null&&g.L.alR!=null&&g.L.ratio>g.L.alR);ins.push([over.length?'c':LS.alR!=null&&LS.ratio<=LS.alR?'o':'','٪',tr(`نسبة الفاقد <b>${pc(LS.ratio)}</b> من الكمية المستلمة${LS.alR!=null?` مقابل مسموح ${pc(LS.alR)}`:''}، وقيمته <b>${fS(A.loss)}</b> د.ل${over.length?`. ${over.length} ${u.grp==='rg'?'منطقة تتجاوز':'مستودع يتجاوز'} المسموح: ${over.slice(0,3).map(g=>bd(g.k)).join('، ')}`:''}.`,
    `Loss ratio <b>${pc(LS.ratio)}</b> of received${LS.alR!=null?` vs ${pc(LS.alR)} allowed`:''}, worth <b>${fS(A.loss)}</b> LYD${over.length?`. ${over.length} ${u.grp==='rg'?'regions':'depots'} exceed the allowance: ${over.slice(0,3).map(g=>bd(g.k)).join(', ')}`:''}.`)])}
  if(A.offM3!=null&&A.pm3!=null){const gap=A.pm3-A.offM3;ins.push([gap<0?'w':'o','≈',tr(`الهامش الفعلي <b>${f2(A.pm3)}</b> د.ل/م³ مقابل هامش التوزيع الرسمي <b>${f2(A.offM3)}</b> د.ل/م³ (${gap>=0?'+':''}${f2(gap)}). ${gap<0?'الفرق يذهب إلى الفاقد وتكاليف النقل والتخزين والخصومات.':'المحقق أعلى من الرسمي.'}`,
    `Actual margin <b>${f2(A.pm3)}</b> LYD/m³ vs the official distribution margin <b>${f2(A.offM3)}</b> LYD/m³ (${gap>=0?'+':''}${f2(gap)}). ${gap<0?'The gap goes to losses, road transport, storage and discounts.':'Realised is above official.'}`)])}
  if(SB.due>0)ins.push([SB.col<SB.due*.8?'w':'o','$',tr(`الدعم المستحق <b>${fS(SB.due)}</b> د.ل، المحصّل منه <b>${fS(SB.col)}</b> (${pct(SB.col,SB.due)}%)${subLast?`، آخر تحصيل ${fdi(subLast)}`:''}. المتبقي <b>${fS(SB.due-SB.col)}</b> د.ل.`,`Subsidy due <b>${fS(SB.due)}</b> LYD, collected <b>${fS(SB.col)}</b> (${pct(SB.col,SB.due)}%)${subLast?`, last on ${fdi(subLast)}`:''}. Outstanding <b>${fS(SB.due-SB.col)}</b> LYD.`)]);
  if(top.length>3&&A.gp>0){const sh=pct(top.reduce((a,x)=>a+Math.max(0,x.gp),0),A.gp);const odT=top.filter(x=>x.ar&&+x.ar.od>0);ins.push([sh>=60?'w':'','10',tr(`أعلى 10 زبائن يحققون <b>${sh}%</b> من الربح${odT.length?`؛ ${odT.length} منهم عليهم متأخرات بـ ${fS(odT.reduce((a,x)=>a+ +x.ar.od,0))} د.ل، فربحهم على الورق لم يُحصَّل بعد`:''}.`,`The top 10 customers bring <b>${sh}%</b> of profit${odT.length?`; ${odT.length} of them owe ${fS(odT.reduce((a,x)=>a+ +x.ar.od,0))} LYD overdue, so their profit is on paper only`:''}.`)])}
  if(gross)ins.push(['i','Σ',tr('لم تُستورد «المصروفات التشغيلية»، لذلك تعرض اللوحة <b>الربح الإجمالي فقط</b> وليس الربح الصافي.','“Operating expenses” are not imported, so the dashboard shows <b>gross profit only</b>, not net profit.')]);
  if(est>0)ins.push(['w','~',tr(`<b>${f3(est)}</b> م³ مبيعة بتكلفة تقديرية (من شحنات مستودعات أخرى أو من هيكل السعر) لعدم وجود استلام بقيمة في مستودعها.`,`<b>${f3(est)}</b> m³ sold at an estimated cost (other depots’ shipments or the price structure) as their depot has no costed receipt.`)]);
  if(B.unpriced)ins.push(['c','?',tr(`${fn(B.unpriced)} سطر مبيعات بلا قيمة ولا سعر (${f3(B.unpricedQ)} م³)، فإيراده صفر. أكمل «قيمة البيع» أو «أسعار الفئات».`,`${fn(B.unpriced)} sales lines have no value and no price (${f3(B.unpricedQ)} m³), so their revenue is zero. Fill “Sales value” or “Category prices”.`)]);
  /* parts */
  const wsel=[...new Set(B.cells.map(x=>x.w).filter(Boolean))].sort(),rsel=[...new Set(B.cells.map(x=>x.rg).filter(Boolean))].sort(),psel=[...new Set(B.cells.map(x=>x.p))].map(p=>[p,nm(p)]).sort((a,b)=>String(a[1]).localeCompare(String(b[1]),'ar'));
  const segOpts=[...new Set(B.lines.filter(x=>x.seg&&!x.fx).map(x=>x.seg))];
  const th=(k,l,cls='n')=>`<th class="${cls}" data-px-sort="${k}" ${u.ps===k?`aria-sort="${u.pd==='asc'?'ascending':'descending'}"`:''}>${esc(l)}</th>`;
  const gauge=m=>{const w=m==null?0:Math.max(0,Math.min(100,m*100/Math.max(1,+c.target*2)));const st=status(m);return `<span style="display:inline-flex;gap:6px;align-items:center">${pc(m)}<span class="pm-gauge" title="${esc(st[1])}"><i class="${st[0]==='p-crit'?'c':st[0]==='p-warn'?'w':''}" style="width:${w}%"></i></span></span>`};
  const lossCell=L0=>L0.ratio==null?'—':`${pc(L0.ratio)}${L0.alR!=null?`<span class="sub">${esc(tr('المسموح','allowed'))} ${pc(L0.alR)} ${L0.ratio>L0.alR?`<span class="px-flag">${esc(tr('تجاوز','over'))}</span>`:''}</span>`:''}`;
  const custRow=(x,i)=>`<tr><td class="n">${i+1}</td><td class="t-strong">${bd(x.cn)} <small class="muted num">${esc(x.k)}</small><span class="sub">${esc(x.seg==='@fx'?tr('أجنبي','Foreign'):catLbl(x.seg))}</span></td><td class="n">${f3(x.q)}</td><td class="n ${x.gp<0?'pm-neg':''}"><b>${f2(x.gp)}</b></td><td class="n">${pc(x.m)}</td><td>${x.ar&&+x.ar.od>0?`<span class="px-flag w" title="${esc(tr('متأخرات','overdue'))}">${esc(tr('متأخر','overdue'))} ${fS(+x.ar.od)}</span>`:''}</td></tr>`;
  const custTbl=list=>list.length?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort n">#</th><th class="nosort">${esc(tr('الزبون','Customer'))}</th><th class="nosort n">${esc(M3())}</th><th class="nosort n">${esc(tr('الربح د.ل','Profit LYD'))}</th><th class="nosort n">${esc(tr('الهامش','Margin'))}</th><th class="nosort"></th></tr></thead><tbody>${list.map(custRow).join('')}</tbody></table></div>`:`<div class="empty" style="padding:20px">${esc(tr('استورد ورقة «المبيعات حسب الزبون».','Import the “Sales by customer” sheet.'))}</div>`;
  const sub=esc(tr(`الفترة ${R.from} – ${R.to} · ${P.length} منتج · ${wsel.length} مستودع · البيانات المتاحة ${R.min} – ${R.max}`,`Period ${R.from} – ${R.to} · ${P.length} products · ${wsel.length} depots · data ${R.min} – ${R.max}`));
  return `${X.local?`<div class="banner info">${esc(tr(`عرض مؤقت للملف «${X.local.file}» على هذا الجهاز فقط، ولم يُحفظ.`,`Temporary view of “${X.local.file}” on this device only; not saved.`))} <button class="btn sm ghost" data-px="drop-local">${esc(tr('إغلاق العرض المؤقت','Close temporary view'))}</button></div>`:''}
  <div class="sm-hero"><div><span class="sm-eyebrow">${esc(tr('إدارة الشؤون التجارية والتحصيلات','Commercial Affairs & Collections'))}</span><h2>${esc(tr('لوحة مؤشرات هامش الربح','Profit margin dashboard'))}</h2><p>${sub}</p>
    <span class="px-badge">${esc(wait?tr('بانتظار أسعار المنتجات: الكميات والتكلفة والفاقد فقط','Awaiting product prices: quantities, cost and losses only'):gross?tr('ربح إجمالي فقط: المصروفات التشغيلية غير مستوردة','Gross profit only: operating expenses not imported'):tr('ربح صافٍ بعد المصروفات التشغيلية','Net profit after operating expenses'))}</span></div>
   <div class="pm-hero-m">${wait?`<div class="sm-hero-v"><b>${esc(tr('بانتظار الأسعار','Awaiting prices'))}</b><span>${esc(tr('هامش الربح','Profit margin'))}</span></div><div class="sm-hero-v"><b>${fS(A.cogs)}</b><span>${esc(tr('تكلفة البضاعة المباعة · د.ل','Cost of goods sold · LYD'))}</span></div><div class="sm-hero-v"><b>${f2(A.q?A.cogs/A.q:null)}</b><span>${esc(tr('تكلفة البضاعة لكل م³','Cost per m³'))}</span></div>`:`<div class="sm-hero-v"><b>${fS(head)}</b><span>${esc(H+' · '+LYD())}</span>${dPct!=null?`<span class="sm-delta">${dPct>=0?'▲':'▼'} ${Math.abs(Math.round(dPct*10)/10)}% ${esc(tr('عن الفترة السابقة','vs previous period'))}</span>`:''}</div><div class="sm-hero-v"><b>${pc(A.m)}</b><span>${esc(tr('هامش الربح الإجمالي','Gross margin'))}</span></div><div class="sm-hero-v"><b>${f2(A.pm3)}</b><span>${esc(tr('د.ل لكل م³','LYD per m³'))}</span></div>`}</div></div>
  ${covStrip(B)}
  <div class="sm-bar">
    <div class="an-f">${esc(tr('الفترة','Period'))}<div class="sm-presets">${[['all',tr('كل البيانات','All data')],['12',tr('آخر 12 شهرًا','Last 12 months')],['3',tr('آخر 3 أشهر','Last 3 months')],['month',tr('آخر شهر','Latest month')],['custom',tr('مخصّص','Custom')]].map(([k,l])=>`<button type="button" data-px-pre="${k}" aria-pressed="${u.pre===k}">${esc(l)}</button>`).join('')}</div></div>
    ${u.pre==='custom'?`<label class="an-f">${esc(tr('من','From'))}<input type="month" id="px-from" value="${esc(R.from)}"></label><label class="an-f">${esc(tr('إلى','To'))}<input type="month" id="px-to" value="${esc(R.to)}"></label>`:''}
    <label class="an-f">${esc(tr('المنتج','Product'))}<select id="px-p"><option value="">${esc(tr('كل المنتجات','All products'))}</option>${psel.map(([p,n])=>`<option value="${esc(p)}" ${u.p===p?'selected':''}>${esc(n)}</option>`).join('')}</select></label>
    ${rsel.length?`<label class="an-f">${esc(tr('المنطقة','Region'))}<select id="px-rg"><option value="">${esc(tr('كل المناطق','All regions'))}</option>${rsel.map(x=>`<option ${u.rg===x?'selected':''}>${esc(x)}</option>`).join('')}</select></label>`:''}
    ${wsel.length?`<label class="an-f">${esc(tr('المستودع','Depot'))}<select id="px-w"><option value="">${esc(tr('كل المستودعات','All depots'))}</option>${wsel.map(x=>`<option ${u.w===x?'selected':''}>${esc(x)}</option>`).join('')}</select></label>`:''}
    <label class="an-f">${esc(tr('فئة الزبون','Customer category'))}<select id="px-seg"><option value="">${esc(tr('كل الفئات','All categories'))}</option>${segOpts.map(k=>`<option value="${esc(k)}" ${u.seg===k?'selected':''}>${esc(catLbl(k))}</option>`).join('')}<option value="fx" ${u.seg==='fx'?'selected':''}>${esc(tr('الأجانب (العملة الأجنبية)','Foreign (foreign currency)'))}</option></select></label>
    ${u.p||u.w||u.rg||u.seg||u.pre!=='all'?`<button class="btn sm ghost" data-px="clear">${esc(tr('مسح التصفية','Clear filters'))}</button>`:''}
  </div>
  ${!cur.length?`<div class="panel empty"><b>${esc(tr('لا توجد بيانات','No data'))}</b>${esc(tr('ضمن التصفية الحالية.','for the current filter.'))}</div>`:wait?waitBody(B,A,P,G,LS,u,lossCell):`
  <div class="kpis an-kpis">
   ${kpi(tr('الإيراد الصافي','Net revenue'),fS(A.nr),esc(tr(`بعد خصومات ${fS(A.disc)} ودعم ${fS(A.sub)} د.ل`,`after ${fS(A.disc)} discounts and ${fS(A.sub)} subsidy`)))}
   ${kpi(tr('تكلفة البضاعة المباعة','Cost of goods sold'),fS(A.cogs),esc(tr('الكمية × متوسط التكلفة المرجح','quantity × weighted average cost')))}
   ${kpi(tr('الربح الإجمالي','Gross profit'),fS(A.gp),esc(PB?tr(`السابقة ${fS(PB.gp)} د.ل`,`previous ${fS(PB.gp)} LYD`):tr('بعد الفاقد والنقل والتخزين','after losses, transport and storage')),A.gp<0?'crit':'')}
   ${kpi(tr('هامش الربح','Gross margin'),pc(A.m),esc(tr(`المستهدف ${c.target}% · الحد الأدنى ${c.min}%`,`target ${c.target}% · minimum ${c.min}%`)),A.m==null?'':A.m<+c.min?'crit':A.m<+c.target?'warn':'ok')}
   ${kpi(tr('الهامش لكل م³','Margin per m³'),f2(A.pm3),esc(tr(`د.ل · ${f4(A.pm3==null?null:A.pm3/1000)} للتر · ${f2(A.pt)} للطن`,`LYD · ${f4(A.pm3==null?null:A.pm3/1000)} per litre · ${f2(A.pt)} per tonne`)))}
   ${kpi(tr('نسبة الفاقد','Loss ratio'),pc(LS.ratio),esc(LS.ratio==null?tr('يلزم مخزون أول وآخر المدة','needs opening and closing stock'):tr(`${f3(LS.loss)} م³ من ${f3(LS.rc)} مستلمة${LS.alR!=null?` · المسموح ${pc(LS.alR)}`:''}`,`${f3(LS.loss)} of ${f3(LS.rc)} m³ received${LS.alR!=null?` · allowed ${pc(LS.alR)}`:''}`)),LS.ratio!=null&&LS.alR!=null?(LS.ratio>LS.alR?'crit':'ok'):'')}
   ${kpi(tr('الهامش الرسمي مقابل الفعلي','Official vs actual margin'),A.offM3==null?'—':`${f2(A.offM3)} / ${f2(A.pm3)}`,esc(A.offM3==null?tr('يلزم «هيكل السعر»','needs the price structure'):tr('د.ل لكل م³: الرسمي / المحقق','LYD per m³: official / realised')),A.offM3!=null&&A.pm3!=null?(A.pm3<A.offM3?'warn':'ok'):'')}
   ${kpi(tr('الدعم: المحصّل من المستحق','Subsidy: collected of due'),SB.due?pct(SB.col,SB.due)+'%':'—',esc(SB.due?tr(`${fS(SB.col)} من ${fS(SB.due)} د.ل`,`${fS(SB.col)} of ${fS(SB.due)} LYD`):tr('يلزم «الدعم الحكومي»','needs the government subsidy sheet')),SB.due?(SB.col<SB.due*.8?'warn':'ok'):'')}
   ${kpi(tr('الربح الصافي','Net profit'),gross?'—':fS(A.np),esc(gross?tr('يلزم «المصروفات التشغيلية»','needs operating expenses'):tr(`بعد مصروفات ${fS(A.opex)} د.ل`,`after ${fS(A.opex)} LYD expenses`)),gross?'':A.np<0?'crit':'')}
   ${kpi(tr('الذمم المتأخرة','Overdue receivables'),AR.length?fS(arOd):'—',esc(AR.length?tr(`من آجل ${fS(arCr)} د.ل · ${arM}`,`of ${fS(arCr)} LYD credit · ${arM}`):tr('يلزم «الذمم»','needs receivables')),AR.length&&arOd>0?'warn':'')}
  </div>
  <div class="an-grid sm-root pm-root">
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('قراءة تحليلية','Reading'))} <small>${esc(tr('تُكتب تلقائيًا من البيانات المصفّاة','written from the filtered data'))}</small></h3><ul class="sm-ins">${ins.map(x=>`<li><span class="ic ${x[0]}">${esc(x[1])}</span><span>${x[2]}</span></li>`).join('')}</ul></section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('من المبيعات إلى الربح','From sales to profit'))} <small>${esc(tr('د.ل · كل خطوة بمعادلتها','LYD · each step by its formula'))}</small></h3>${waterfall(A)}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('ترتيب المنتجات حسب الربحية','Products ranked by profitability'))} <small>${esc(tr('اضغط عنوان العمود للترتيب','click a column header to sort'))}</small></h3>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort n">#</th>${th('pn',tr('المنتج','Product'),'')}${th('q',tr('المباع م³','Sold m³'))}${th('nr',tr('الإيراد الصافي','Net revenue'))}<th class="nosort n">${esc(tr('تكلفة البضاعة','COGS'))}</th><th class="nosort n">${esc(tr('فاقد ونقل وتخزين','Losses, transport, storage'))}</th>${th('gp',tr('الربح الإجمالي','Gross profit'))}${th('m',tr('الهامش','Margin'))}${th('pm3',tr('لكل م³','Per m³'))}<th class="nosort n">${esc(tr('الرسمي / الفعلي','Official / actual'))}</th>${th('loss',tr('الفاقد','Losses'))}<th class="nosort n">${esc(tr('بدون الدعم','Without subsidy'))}</th></tr></thead>
    <tbody>${P.map((p,i)=>`<tr><td class="n">${i+1}</td><td class="t-strong">${bd(p.pn)} <small class="muted num">${esc(p.k)}</small><span class="sub"><span class="pill ${status(p.m)[0]}">${esc(status(p.m)[1])}</span></span></td><td class="n">${f3(p.q)}<span class="sub">${p.t?f3(p.t)+' '+esc(tr('طن','t')):''}</span></td><td class="n">${f2(p.nr)}${p.sub?`<span class="sub">${esc(tr('منه دعم','incl. subsidy'))} ${fS(p.sub)}</span>`:''}</td><td class="n">${f2(p.cogs)}</td><td class="n">${f2(p.loss+p.dist)}</td><td class="n ${p.gp<0?'pm-neg':''}"><b>${f2(p.gp)}</b></td><td class="n">${gauge(p.m)}</td><td class="n">${f2(p.pm3)}<span class="sub">${f4(p.pm3==null?null:p.pm3/1000)} ${esc(tr('للتر','/ L'))}</span></td><td class="n">${p.offM3==null?'—':`${f2(p.offM3)} / ${f2(p.pm3)}${p.pm3!=null&&p.pm3<p.offM3?`<span class="sub"><span class="px-flag w">${esc(tr('أقل من الرسمي','below official'))}</span></span>`:''}`}</td><td class="n">${lossCell(p.L)}</td><td class="n ${p.noSub<0?'pm-neg':''}">${p.sub?f2(p.noSub):'—'}</td></tr>`).join('')}</tbody>
    <tfoot><tr><td></td><td class="t-strong">${esc(tr('الإجمالي','Total'))}</td><td class="n"><b>${f3(A.q)}</b></td><td class="n"><b>${f2(A.nr)}</b></td><td class="n"><b>${f2(A.cogs)}</b></td><td class="n"><b>${f2(A.loss+A.dist)}</b></td><td class="n"><b>${f2(A.gp)}</b></td><td class="n"><b>${pc(A.m)}</b></td><td class="n"><b>${f2(A.pm3)}</b></td><td class="n">${A.offM3==null?'':f2(A.offM3)+' / '+f2(A.pm3)}</td><td class="n">${pc(LS.ratio)}</td><td class="n">${A.sub?f2(A.noSub):''}</td></tr></tfoot></table></div></section>
   <section class="panel"><h3 class="section-title">${esc(tr('هامش الربح % حسب المنتج','Margin % by product'))} <small>${esc(tr('الربح الإجمالي ÷ الإيراد الصافي','gross profit ÷ net revenue'))}</small></h3>${bars(P.filter(p=>p.m!=null).map(p=>({l:p.pn,v:Math.round(p.m*10)/10,tip:`${p.pn}: ${pc(p.m)} · ${f2(p.pm3)} ${LYD()}/${M3()}`})).sort((a,b)=>b.v-a.v),{fmt:v=>pc(v)})}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('الربح الإجمالي حسب المنتج','Gross profit by product'))} <small>${esc(LYD())}</small></h3>${bars(P.slice().sort((a,b)=>b.gp-a.gp).map(p=>({l:p.pn,v:p.gp,tip:`${p.pn}: ${f2(p.gp)} ${LYD()}`})),{fmt:fS})}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('هيكل السعر الرسمي','Official price structure'))} <small>${esc(tr('د.ل للتر · مع التكلفة الفعلية المحققة','LYD per litre · with the actual realised cost'))}</small></h3>${structure(B,P,R)}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(u.grp==='rg'?tr('مقارنة المناطق','Regions compared'):tr('مقارنة المستودعات','Depots compared'))} <small>${esc(tr('الفاقد بجانب الهامش','losses next to margin'))}</small>
     <span class="px-seg" role="group"><button type="button" data-px-grp="w" aria-pressed="${u.grp!=='rg'}">${esc(tr('المستودعات','Depots'))}</button><button type="button" data-px-grp="rg" aria-pressed="${u.grp==='rg'}">${esc(tr('المناطق','Regions'))}</button></span></h3>
    ${G.length?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(u.grp==='rg'?tr('المنطقة','Region'):tr('المستودع','Depot'))}</th><th class="nosort n">${esc(tr('المباع م³','Sold m³'))}</th><th class="nosort n">${esc(tr('الإيراد الصافي','Net revenue'))}</th><th class="nosort n">${esc(tr('الربح الإجمالي','Gross profit'))}</th><th class="nosort n">${esc(tr('الهامش','Margin'))}</th><th class="nosort n">${esc(tr('لكل م³','Per m³'))}</th><th class="nosort n">${esc(tr('الفاقد','Losses'))}</th><th class="nosort">${esc(tr('الهامش مقابل الفاقد','Margin vs losses'))}</th></tr></thead>
     <tbody>${G.map(g=>{const mxM=Math.max(1,...G.map(x=>Math.abs(x.m||0))),mxL=Math.max(.5,...G.map(x=>x.L.ratio||0));return `<tr><td class="t-strong">${bd(g.k)}${u.grp!=='rg'&&g.ls[0].rg?`<span class="sub">${bd(g.ls[0].rg)}</span>`:''}</td><td class="n">${f3(g.q)}</td><td class="n">${f2(g.nr)}</td><td class="n ${g.gp<0?'pm-neg':''}"><b>${f2(g.gp)}</b></td><td class="n">${gauge(g.m)}</td><td class="n">${f2(g.pm3)}</td><td class="n">${lossCell(g.L)}</td><td><span class="px-dual" title="${esc(tr('الهامش','margin')+' '+pc(g.m)+' · '+tr('الفاقد','losses')+' '+pc(g.L.ratio))}"><span class="a" style="width:${Math.max(0,(g.m||0))*100/mxM}%"></span><span class="b" style="width:${Math.max(0,(g.L.ratio||0))*100/mxL}%"></span></span></td></tr>`}).join('')}</tbody></table></div>
     <div class="px-leg"><span><i style="background:var(--sm-c1,#0f766e)"></i>${esc(tr('هامش الربح %','Margin %'))}</span><span><i style="background:var(--crit)"></i>${esc(tr('نسبة الفاقد %','Loss ratio %'))}</span></div>`:`<div class="empty" style="padding:20px">${esc(tr('لا توجد مستودعات في البيانات.','No depots in the data.'))}</div>`}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(gross?tr('الربح الإجمالي الشهري','Monthly gross profit'):tr('الربح الصافي الشهري','Monthly net profit'))} <small>${esc(tr('د.ل · مرّر على العمود للتفصيل','LYD · hover a column for detail'))}</small></h3>${colChart(pts,tr('الربح حسب الشهر','Profit by month'))}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('اتجاه الهامش والفاقد','Margin and loss trend'))} <small>${esc(tr('% شهريًا','% by month'))}</small></h3>${lineChart(lpts,tr('هامش الربح','Gross margin'),tr('نسبة الفاقد','Loss ratio'))}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('فئات الزبائن','Customer categories'))} <small>${esc(tr('محطات، جهات حكومية، شركات، طيران، صناعة، كهرباء، أجانب','stations, government, companies, aviation, industry, electricity, foreign'))}</small></h3>
    ${SG.length?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(tr('الفئة','Category'))}</th><th class="nosort n">${esc(tr('المباع م³','Sold m³'))}</th><th class="nosort n">${esc(tr('متوسط السعر الصافي د.ل/لتر','Avg net price LYD/L'))}</th><th class="nosort n">${esc(tr('الخصومات والعمولات','Discounts & commissions'))}</th><th class="nosort n">${esc(tr('الربح الإجمالي','Gross profit'))}</th><th class="nosort n">${esc(tr('الهامش','Margin'))}</th><th class="nosort n">${esc(tr('حصة الربح','Profit share'))}</th></tr></thead>
     <tbody>${SG.map(g=>`<tr><td class="t-strong">${esc(g.k==='@fx'?tr('أجانب (العملة الأجنبية)','Foreign (foreign currency)'):catLbl(g.k))}</td><td class="n">${f3(g.q)}</td><td class="n">${g.q?f4((g.gross-g.disc)/g.q/1000):'—'}</td><td class="n">${f2(g.disc)}</td><td class="n ${g.gp<0?'pm-neg':''}"><b>${f2(g.gp)}</b></td><td class="n">${gauge(g.m)}</td><td class="n">${A.gp>0?pct(Math.max(0,g.gp),A.gp)+'%':'—'}</td></tr>`).join('')}</tbody></table></div>`:`<div class="empty" style="padding:20px">${esc(tr('استورد ورقة «المبيعات حسب الزبون» مع عمود «فئة الزبون».','Import “Sales by customer” with the “Customer category” column.'))}</div>`}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('أكثر 10 زبائن ربحًا وأقلهم','Top and bottom 10 customers by profit'))} <small>${esc(tr('مع المتأخرات من ورقة الذمم: الربح على الورق لا يعني أن المال حُصِّل','with overdue from receivables: profit on paper is not cash collected'))}</small></h3>
    <div class="px-two"><div><h4 style="margin:0 0 6px">${esc(tr('الأعلى ربحًا','Most profitable'))}</h4>${custTbl(top)}</div><div><h4 style="margin:0 0 6px">${esc(tr('الأقل ربحًا','Least profitable'))}</h4>${custTbl(bot)}</div></div></section>
   <section class="panel"><h3 class="section-title">${esc(tr('الدعم الحكومي: المستحق والمحصّل','Government subsidy: due and collected'))} <small>${esc(LYD())}</small></h3>
    ${subP.size?`<div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-tbl"><thead><tr><th class="nosort">${esc(tr('المنتج','Product'))}</th><th class="nosort n">${esc(tr('الدعم د.ل/لتر','Subsidy LYD/L'))}</th><th class="nosort n">${esc(tr('المستحق','Due'))}</th><th class="nosort n">${esc(tr('المحصّل','Collected'))}</th><th class="nosort n">${esc(tr('المتبقي','Outstanding'))}</th><th class="nosort n">${esc(tr('آخر تحصيل','Last collected'))}</th></tr></thead>
     <tbody>${[...subP.values()].sort((a,b)=>b.due-a.due).map(o=>`<tr><td class="t-strong">${bd(nm(o.p))}</td><td class="n">${o.q?f4(o.rateQ/o.q/1000):'—'}</td><td class="n">${f2(o.due)}</td><td class="n">${f2(o.col)}<span class="sub">${o.due?pct(o.col,o.due)+'%':''}</span></td><td class="n ${o.due-o.col>0?'pm-neg':''}">${f2(o.due-o.col)}</td><td class="n">${o.last?fdi(o.last):'—'}</td></tr>`).join('')}</tbody>
     <tfoot><tr><td class="t-strong">${esc(tr('الإجمالي','Total'))}</td><td></td><td class="n"><b>${f2(SB.due)}</b></td><td class="n"><b>${f2(SB.col)}</b></td><td class="n"><b>${f2(SB.due-SB.col)}</b></td><td></td></tr></tfoot></table></div>`:`<div class="empty" style="padding:20px">${esc(tr('استورد ورقة «الدعم الحكومي».','Import the “Government subsidy” sheet.'))}</div>`}</section>
   <section class="panel"><h3 class="section-title">${esc(tr('المصروفات التشغيلية','Operating expenses'))} <small>${esc(tr('تُوزَّع على المنتجات حسب الكمية المباعة','allocated to products by quantity sold'))}</small></h3>
    ${opexK.length?bars(opexK.map(g=>({l:g.k,v:g.opex,tip:`${g.k}: ${f2(g.opex)} ${LYD()}`})),{fmt:fS})+`<div class="sm-mini" style="margin-top:10px"><div><b class="num">${fS(A.opex)}</b><span>${esc(tr('مصروفات الفترة د.ل','Period expenses LYD'))}</span></div><div><b class="num">${A.q?f2(A.opex/A.q):'—'}</b><span>${esc(tr('د.ل لكل م³ مباع','LYD per m³ sold'))}</span></div><div><b class="num ${A.np<0?'pm-neg':''}">${fS(A.np)}</b><span>${esc(tr('الربح الصافي د.ل','Net profit LYD'))}</span></div></div>`
     :`<div class="empty" style="padding:20px"><b>${esc(tr('الربح الإجمالي فقط','Gross profit only'))}</b>${esc(tr('استورد ورقة «المصروفات التشغيلية» (الرواتب، الصيانة، الإهلاك، الكهرباء، الإدارية) لحساب الربح الصافي.','Import “Operating expenses” (salaries, maintenance, depreciation, electricity, admin) to compute net profit.'))}</div>`}</section>
   <section class="panel an-wide"><h3 class="section-title">${esc(tr('المعادلات','Formulas'))} <small>${esc(tr('ليقصد الجميع الشيء نفسه','so everyone means the same thing'))}</small></h3>
    <div class="tablewrap" style="max-height:none;min-height:0"><table class="sm-tbl px-fx"><thead><tr><th class="nosort">${esc(tr('المؤشر','Indicator'))}</th><th class="nosort">${esc(tr('المعادلة','Formula'))}</th></tr></thead><tbody>
     ${[[tr('الإيراد الصافي','Net revenue'),tr('(الكمية × السعر) − الخصومات والعمولات + الدعم','(quantity × price) − discounts and commissions + subsidy')],[tr('تكلفة البضاعة المباعة','Cost of goods sold'),tr('الكمية × متوسط التكلفة المرجح (قيمة الشراء + الشحن + التأمين + الموانئ + الجمارك، مرجّحة مع مخزون أول المدة لكل مستودع)','quantity × weighted average cost (purchase + freight + insurance + ports + customs, weighted with opening stock per depot)')],
       [tr('الربح الإجمالي','Gross profit'),tr('الإيراد الصافي − تكلفة البضاعة − قيمة الفاقد − النقل البري والتخزين','net revenue − cost of goods − value of losses − road transport and storage')],[tr('هامش الربح %','Gross margin %'),tr('الربح الإجمالي ÷ الإيراد الصافي × 100','gross profit ÷ net revenue × 100')],
       [tr('الهامش لكل م³ / لتر / طن','Margin per m³ / litre / tonne'),tr('الربح الإجمالي ÷ الكمية المباعة (م³ عند 15°م؛ اللتر = م³ ÷ 1000؛ الطن = م³ × الكثافة ÷ 1000)','gross profit ÷ quantity sold (m³ at 15 °C; litre = m³ ÷ 1000; tonne = m³ × density ÷ 1000)')],
       [tr('نسبة الفاقد','Loss ratio'),tr('الفاقد الفعلي ÷ الكمية المستلمة؛ الفاقد = أول المدة + المستلم + الوارد − الصادر − المباع − آخر المدة','actual loss ÷ received; loss = opening + received + in − out − sold − closing')],
       [tr('الهامش الرسمي مقابل الفعلي','Official vs actual margin'),tr('هامش توزيع البريقة في هيكل السعر مقابل الربح الإجمالي المحقق لكل م³ بعد التكاليف والفاقد','Brega distribution margin in the price structure vs realised gross profit per m³ after costs and losses')],
       [tr('الربح الصافي','Net profit'),tr('الربح الإجمالي − المصروفات التشغيلية الموزعة حسب الكمية المباعة','gross profit − operating expenses allocated by quantity sold')]].map(([a,b])=>`<tr><td>${esc(a)}</td><td>${esc(b)}</td></tr>`).join('')}</tbody></table></div>
    <p class="sm-note" style="margin:8px 0 0">${esc(tr('الكميات بالمتر المكعب عند 15°م: اللتر ÷ 1000 مصحّحًا بعامل تصحيح الحجم عند درجة حرارة السطر، والطن × 1000 ÷ الكثافة عند 15°م. النسب تُحسب من المجاميع دائمًا ولا تُجمع المتوسطات. السعر الأجنبي يُحوَّل بسعر الصرف في ورقة «أسعار الفئات» أو بسعر الصرف المحفوظ في تبويب الأسعار.','Quantities in m³ at 15 °C: litres ÷ 1000 corrected with the volume correction factor at the row temperature, tonnes × 1000 ÷ density at 15 °C. Ratios always come from totals, never averaged. Foreign prices use the exchange rate in “Category prices” or the saved rate in the prices tab.'))}${B.mismatch?' '+esc(tr(`تنبيه: في ${B.mismatch} حالة مجموع مبيعات الزبائن أكبر من الكمية المباعة في ورقة المخزون؛ اعتُمد مجموع الزبائن.`,`Note: in ${B.mismatch} cases customer sales exceed the sold quantity on the stock sheet; the customer total is used.`)):''}</p></section>
  </div>`}`}
function after(){const u=UI();[['px-p','p'],['px-w','w'],['px-rg','rg'],['px-seg','seg']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('change',()=>{u[k]=e.value;render()})});
  [['px-from','from'],['px-to','to']].forEach(([id,k])=>{const e=document.getElementById(id);if(e)e.addEventListener('change',()=>{u[k]=e.value||'';render()})})}

/* ---------- template and export ---------- */
function template(){const RP=refMap(X.ref.prods);const base=[['33500','بنزين','وقود','لتر',740],['44347','ديزل','وقود','لتر',835],['','كيروسين','وقود','لتر',800],['16291','غاز الطهي (LPG)','غاز','طن',540],['','وقود الطائرات','وقود','لتر',800],
    ['','زيت الوقود الثقيل','وقود','طن',970],['','وقود السفن','وقود','طن',890],['','الزيوت والشحوم','زيوت وشحوم','لتر',890],['','الأسفلت','أسفلت','طن',1020]];
  const prods=(X.ref.prods&&X.ref.prods.length?X.ref.prods.map(x=>[x.p===x.pn?'':x.p,x.pn,x.cat||'',x.u||'',x.d15||'']):base).map(r=>({'رمز المنتج':r[0],'اسم المنتج':r[1],'الفئة':r[2],'وحدة القياس':r[3],'الكثافة عند 15°م':r[4]}));
  const first=new Date().toISOString().slice(0,7)+'-01';const plist=prods.map(x=>[x['رمز المنتج'],x['اسم المنتج']]);
  const blank=def=>{const o={};def.f.filter(f=>!f[5]).forEach(f=>o[f[1]]='');return [o]};
  const struct=plist.map(([p,n])=>({'رمز المنتج':p,'اسم المنتج':n,'تاريخ السريان':first,'وحدة السعر':'لتر','تكلفة المنتج':'','النقل':'','هامش توزيع البريقة':'','هامش المحطة':'','الرسوم':'','السعر النهائي':''}));
  const segs=[];plist.forEach(([p,n])=>(window.CUSTOMER_CATS||[]).forEach(cg=>segs.push({'رمز المنتج':p,'اسم المنتج':n,'فئة الزبون':cg.ar,'تاريخ السريان':first,'وحدة السعر':'لتر','السعر المحلي':'','سعر الأجانب':'','العملة':(PMa.PM.cfg||{}).cur||'USD','سعر الصرف':'','نسبة الخصم %':'','العمولة للوحدة':''})));
  const readme=[...SH.map(d=>({'الورقة':d.ar,'الغرض':d.why[0],'الأعمدة':d.f.filter(f=>!f[5]).map(f=>f[1]+(f[4]?' (إلزامي)':'')).join('، ')})),
    {'الورقة':'قواعد عامة','الغرض':'كل رقم لكل منتج ولكل مستودع أو منطقة ولكل شهر. الشهر بصيغة 2026-09 أو تاريخ داخل الشهر.','الأعمدة':''},
    {'الورقة':'الوحدات','الغرض':'الوحدة: لتر أو م³ أو طن أو كغم أو برميل. تُحوَّل كل الكميات إلى م³ عند 15°م؛ اللتر والمتر المكعب يُصحَّحان بدرجة الحرارة إن وُجدت، والطن بالكثافة عند 15°م من ورقة المنتجات. إن تُركت الوحدة فارغة تُؤخذ وحدة المنتج.','الأعمدة':''},
    {'الورقة':'وحدة السعر','الغرض':'الأسعار والتكاليف والدعم بالدينار لكل وحدة السعر (لتر افتراضيًا، أو م³ أو طن). سعر الأجانب بالعملة الأجنبية ويُحوَّل بسعر الصرف.','الأعمدة':''},
    {'الورقة':'فئات الزبائن','الغرض':(window.CUSTOMER_CATS||[]).map(x=>x.ar).join('، ')+'. الزبون الأجنبي: «نعم» في عمود «زبون أجنبي».','الأعمدة':''},
    {'الورقة':'الفاقد','الغرض':'الفاقد الفعلي = أول المدة + المستلم + التحويلات الواردة − الصادرة − المباع − آخر المدة. الفاقد المسموح نسبة مئوية من المستلم.','الأعمدة':''},
    {'الورقة':'التحديث','الغرض':'يمكن استيراد شهر واحد أو عدة أشهر أو ورقة واحدة؛ الأوراق الموجودة تستبدل بياناتها في الأشهر نفسها فقط.','الأعمدة':''}];
  saveXlsx([['اقرأني',readme],['المنتجات',prods],['هيكل السعر',struct],['أسعار الفئات',segs],...['recv','dep','stock','sales','sub','opex','ar'].map(k=>[SHK[k].ar,blank(SHK[k])])],'Brega-profit-margin-template.xlsx')}
function exportXlsx(){const B=build(),R=range(B);if(!R)return;const cur=fl(B,R.from,R.to),A=sum(cur),cs=fc(B,R.from,R.to),LS=lossOf(cs);const nm=p=>B.names.get(p)||p;
  /* a group with no price and no sales value has no revenue yet: its profit columns stay empty instead of a «loss» equal to its cost */
  const row=g=>!(g.gross>0)&&!(g.sub>0)?Object.assign(row0(g),{'الإيراد الصافي':'','الربح الإجمالي':'','هامش الربح %':'','الربح لكل م³':'','الربح الصافي':'','الحالة':'بانتظار الأسعار'}):row0(g);
  const row0=g=>({'المباع م³':r6(g.q),'المبيعات':r2(g.gross),'الخصومات والعمولات':r2(g.disc),'الدعم':r2(g.sub),'الإيراد الصافي':r2(g.nr),'تكلفة البضاعة':r2(g.cogs),'قيمة الفاقد':r2(g.loss),'النقل والتخزين':r2(g.dist),'الربح الإجمالي':r2(g.gp),'هامش الربح %':g.m==null?'':r2(g.m),'الربح لكل م³':g.pm3==null?'':r2(g.pm3),'الهامش الرسمي لكل م³':g.offM3==null?'':r2(g.offM3),'المصروفات التشغيلية':r2(g.opex),'الربح الصافي':r2(g.np)});
  const sumS=[['الفترة',`${R.from} – ${R.to}`],['الإيراد الصافي',r2(A.nr)],['تكلفة البضاعة المباعة',r2(A.cogs)],['الربح الإجمالي',r2(A.gp)],['هامش الربح %',A.m==null?'':r2(A.m)],['الربح لكل م³',A.pm3==null?'':r2(A.pm3)],['نسبة الفاقد %',LS.ratio==null?'':r2(LS.ratio)],['المصروفات التشغيلية',r2(A.opex)],['الربح الصافي',B.has.opex?r2(A.np):'غير متاح: المصروفات غير مستوردة']].map(([a,b])=>({'البند':a,'القيمة':b}));
  saveXlsx([['ملخص',sumS],['المنتجات',group(cur,x=>x.p).map(g=>Object.assign({'رمز المنتج':g.k,'اسم المنتج':nm(g.k)},row(g)))],['المستودعات',group(cur,x=>x.w).map(g=>Object.assign({'المستودع':g.k,'المنطقة':g.ls[0].rg||''},row(g),{'نسبة الفاقد %':(l=>l.ratio==null?'':r2(l.ratio))(lossOf(cs.filter(c=>c.w===g.k)))}))],
    ['الأشهر',group(cur,x=>x.m).sort((a,b)=>a.k.localeCompare(b.k)).map(g=>Object.assign({'الشهر':g.k},row(g)))],['الزبائن',group(cur.filter(x=>x.c),x=>x.c).sort((a,b)=>b.gp-a.gp).map(g=>Object.assign({'رمز الزبون':g.k,'اسم الزبون':g.ls[0].cn,'الفئة':g.ls[0].fx?'أجنبي':catLbl(g.ls[0].seg)},row(g)))]],`Profit-margin-full-${R.from}_${R.to}.xlsx`)}

/* ---------- events ---------- */
document.addEventListener('click',async e=>{if(S.view!==VIEW&&!(document.getElementById('dlg')||{}).open)return;
  const pre=e.target.closest('[data-px-pre]');if(pre){const u=UI();u.pre=pre.dataset.pxPre;if(u.pre==='custom'){const B=build(),R=range(B);if(R){u.from=u.from||R.from;u.to=u.to||R.to}}render();return}
  const so=e.target.closest('[data-px-sort]');if(so){const u=UI(),k=so.dataset.pxSort;if(u.ps===k)u.pd=u.pd==='asc'?'desc':'asc';else{u.ps=k;u.pd=k==='pn'?'asc':'desc'}render();return}
  const gp=e.target.closest('[data-px-grp]');if(gp){UI().grp=gp.dataset.pxGrp;render();return}
  const a=e.target.closest('[data-px]');if(!a)return;const act=a.dataset.px;
  if(act==='template'){template();return}
  if(act==='export'){exportXlsx();return}
  if(act==='clear'){Object.assign(UI(),{pre:'all',from:'',to:'',p:'',w:'',rg:'',seg:''});render();return}
  if(act==='drop-local'){X.local=null;render();return}
  if(act==='view-local'){const P=X.pending;if(!P)return;const {R,M}=merge(P,X.ref,X.months);X.local={ref:R,months:M,file:P.file,ver:Date.now()};X.pending=null;dlg.close();UI().pre='all';render();return}
  if(act==='save'){const P=X.pending;if(P)await save(P,a);return}});

window.__pmx={view,after,load,tryImport,template,exportXlsx,has:()=>{const B=X.loaded?build():null;return !!(B&&B.months.length)},state:X,build,sum,fl,range,qM3,perM3};
})();
