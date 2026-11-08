/* ===== Power BI export / import next to every Excel button =====
   - every «Excel» export button gets a joined Power BI button (button group). It runs the same export, but saveXlsx
     is routed here: a dialog lists the tables, then a .zip downloads with one UTF-8 CSV per table (ISO dates, plain numbers),
     a ready Power Query script with the column types, and a short read-me.
   - every Excel import gets a Power BI counterpart: a dialog explains «Export data» in Power BI, then opens the same file
     picker. The settings import now also reads .csv (Power BI exports CSV with a BOM).
   Loaded after settings.js (saveXlsx, zipStore) and fileimport.js (readWorkbook) and before odoo.js. */
(function(){
"use strict";
const en=()=>!!(window.I18N&&I18N.lang==='en');
const T=(a,e)=>en()?e:a;
const ICON='<svg class="bi-ic" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><rect x="3" y="12" width="4.6" height="9" rx="1.4" fill="#F2C811" opacity=".55"/><rect x="9.7" y="7.5" width="4.6" height="13.5" rx="1.4" fill="#E8B100" opacity=".8"/><rect x="16.4" y="3" width="4.6" height="18" rx="1.4" fill="#C99700"/></svg>';
window.__biIcon=ICON;

/* ---------- data shaping ---------- */
const ISO=/^\d{4}-\d{2}-\d{2}$/;
const NUMS=/^-?\d+(\.\d+)?$/;
function columns(rows){const k=[];rows.forEach(r=>Object.keys(r).forEach(x=>{if(!k.includes(x))k.push(x)}));return k}
function colType(rows,k){let n=0,int=true,num=true,date=true,bool=true;
  for(const r of rows){const v=r[k];if(v==null||v==='')continue;n++;
    if(typeof v==='boolean'){num=false;date=false;continue}bool=false;
    if(typeof v==='number'&&isFinite(v)){date=false;if(!Number.isInteger(v))int=false;continue}
    num=false;if(!(typeof v==='string'&&ISO.test(v.trim())))date=false;
    if(!num&&!date)break}
  if(!n)return'text';if(bool)return'logical';if(num)return int?'int':'number';if(date)return'date';return'text'}
/* a text cell that Excel would read as a formula is prefixed with an apostrophe (CSV injection) */
function cellOut(v){if(v==null)return'';if(typeof v==='boolean')return v?'TRUE':'FALSE';if(typeof v==='number')return isFinite(v)?String(v):'';
  let s=String(v).replace(/\r\n?/g,'\n');if(/^[=+@\t\r]/.test(s)||(/^-/.test(s)&&!NUMS.test(s)))s="'"+s;
  return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}
function toCsv(rows,keys){return '\uFEFF'+[keys.map(cellOut).join(','),...rows.map(r=>keys.map(k=>cellOut(r[k])).join(','))].join('\r\n')+'\r\n'}
function fileSafe(s){return String(s).replace(/[\/\\:*?"<>|\x00-\x1f]/g,' ').replace(/\s+/g,' ').trim().slice(0,60)||'table'}
const mq=s=>'"'+String(s).replace(/"/g,'""')+'"';
const MT={int:'Int64.Type',number:'type number',date:'type date',logical:'type logical',text:'type text'};
function mScript(tables,title){
  const L=[];
  L.push('// ===== Power BI — '+title+' — '+TODAY_ISO+' =====');
  L.push('// الأسرع: Power BI Desktop ← Get data ← Text/CSV ← اختر ملف الجدول ← Load.');
  L.push('// لأنواع أعمدة صحيحة (تواريخ وأرقام): Get data ← Blank query ← Advanced editor ← الصق استعلام الجدول، وعدّل FolderPath إلى مجلد فك الضغط.');
  L.push('// Fastest: Get data → Text/CSV → pick the table file. For typed columns: Blank query → Advanced editor → paste a query below, set FolderPath.');
  tables.forEach(t=>{
    L.push('','// ---- '+t.name+'  ('+t.rows.length+' rows) — query name: '+t.name+' ----');
    L.push('let');
    L.push('    FolderPath = "C:\\Brega\\PowerBI\\",');
    L.push('    Source = Csv.Document(File.Contents(FolderPath & '+mq(t.file)+'), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),');
    L.push('    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),');
    const ty=t.keys.map(k=>'{'+mq(k)+', '+MT[t.types[k]]+'}');
    L.push('    Typed = Table.TransformColumnTypes(Headers, {'+ty.join(', ')+'}, "en-US")');
    L.push('in','    Typed')});
  return L.join('\r\n')+'\r\n'}
function readme(tables,title){
  return '\uFEFF'+[
   'منظومة الإدارة العامة للمبيعات — حزمة Power BI',
   title+' — '+TODAY_ISO,'',
   'المحتوى:',...tables.map(t=>'  • '+t.file+' — '+t.rows.length+' صف، '+t.keys.length+' عمود'),
   '  • BregaPowerBI.pq — استعلامات Power Query بأنواع الأعمدة','',
   'طريقة الفتح:',
   '  1) فك ضغط الملف في مجلد، مثلًا C:\\Brega\\PowerBI',
   '  2) في Power BI Desktop: Get data ← Text/CSV ← اختر ملف الجدول ← Load.',
   '  3) أو للأنواع الدقيقة: Get data ← Blank query ← Advanced editor، والصق الاستعلام من BregaPowerBI.pq.','',
   'الملفات بترميز UTF-8، والتواريخ بصيغة YYYY-MM-DD، والأرقام دون فواصل آلاف.',
   'للإعادة إلى المنظومة: من Power BI اختر «…» ← Export data ← .xlsx أو .csv، واحتفظ بأسماء الأعمدة كما هي، ثم استخدم زر استيراد Power BI.',
   'هذه بيانات داخلية لشركة البريقة: لا تنشر ملف .pbix أو التقرير للعامة.','',
   'Brega sales system — Power BI pack. Unzip, then Power BI Desktop → Get data → Text/CSV. Typed queries are in BregaPowerBI.pq. Internal data: do not publish.'
  ].join('\r\n')+'\r\n'}
function pack(sheets,filename){
  const used=new Set();
  const tables=sheets.filter(([,r])=>r&&r.length).map(([name,rows])=>{const keys=columns(rows).filter(k=>k!=='');let f=fileSafe(name);let i=2;while(used.has(f))f=fileSafe(name)+' '+(i++);used.add(f);
    const types={};keys.forEach(k=>types[k]=colType(rows,k));return {name:String(name),file:f+'.csv',rows,keys,types}});
  const title=String(filename||'').replace(/\.xlsx?$/i,'');
  return {tables,title,zipName:title+'-PowerBI.zip'}}
function zipBlob(p){const files=p.tables.map(t=>[t.file,toCsv(t.rows,t.keys)]);files.push(['BregaPowerBI.pq',mScript(p.tables,p.title)],['اقرأني-README.txt',readme(p.tables,p.title)]);
  const b=zipStore(files);return new Blob([b],{type:'application/zip'})}
window.__biExport=(tables,title)=>exportDialog(tables.map(t=>[t.name,t.rows]),title+'.xlsx');
window.__biPack=(sheets,filename)=>{const p=pack(sheets,filename);return {tables:p.tables.map(t=>({name:t.name,file:t.file,rows:t.rows.length,keys:t.keys,types:t.types})),zipName:p.zipName,csv:p.tables.map(t=>toCsv(t.rows,t.keys)),pq:mScript(p.tables,p.title),blob:p.tables.length?zipBlob(p):null}};

/* ---------- dialog ---------- */
let dl=null;
function dialog(){if(dl&&document.body.contains(dl))return dl;dl=document.createElement('dialog');dl.id='bi-dlg';dl.className='bi-dlg';dl.setAttribute('aria-labelledby','bi-title');document.body.appendChild(dl);
  dl.addEventListener('click',e=>{if(e.target===dl)dl.close()});return dl}
function show(title,sub,body,foot){const d=dialog();
  d.innerHTML=`<div class="m-head"><span class="bi-badge">${ICON}</span><div><h3 id="bi-title">${esc(title)}</h3>${sub?`<p class="muted bi-sub">${esc(sub)}</p>`:''}</div><button type="button" class="iconbtn" data-bi="close" style="margin-inline-start:auto" aria-label="${esc(T('إغلاق','Close'))}">${I.close}</button></div><div class="m-body">${body}</div><div class="m-foot">${foot}</div>`;
  if(!d.open)d.showModal();const f=d.querySelector('.m-foot .primary');f&&f.focus()}
const TYL={int:['عدد صحيح','Whole number'],number:['رقم','Decimal'],date:['تاريخ','Date'],logical:['صح/خطأ','True/False'],text:['نص','Text']};
let PEND=null;
function exportDialog(sheets,filename){
  const p=pack(sheets,filename);
  if(!p.tables.length){toast(T('لا توجد بيانات للتصدير','Nothing to export'));return}
  PEND=p;
  const fmt=n=>{try{return Number(n).toLocaleString(en()?'en-US':'ar-LY')}catch(e){return String(n)}};
  const list=p.tables.map(t=>{const c={};t.keys.forEach(k=>{c[t.types[k]]=(c[t.types[k]]||0)+1});
    const chips=Object.entries(c).map(([ty,n])=>`<span class="bi-ty bi-ty-${ty}">${esc(T(TYL[ty][0],TYL[ty][1]))} ${fmt(n)}</span>`).join('');
    return `<li><div class="bi-t"><b>${esc(t.name)}</b><span class="muted">${esc(t.file)}</span></div><div class="bi-n"><span>${fmt(t.rows.length)} ${esc(T('صف','rows'))}</span><span>${fmt(t.keys.length)} ${esc(T('عمود','columns'))}</span></div><div class="bi-tys">${chips}</div></li>`}).join('');
  const body=`<ul class="bi-tables">${list}</ul>
   <div class="bi-pack"><div class="bi-pk-h">${esc(T('محتوى الحزمة','What is in the pack'))}</div><ul>
    <li>${esc(T('ملف CSV لكل جدول بترميز UTF-8، تواريخ YYYY-MM-DD وأرقام بلا فواصل.','One UTF-8 CSV per table, YYYY-MM-DD dates, plain numbers.'))}</li>
    <li>${esc(T('استعلام Power Query جاهز (BregaPowerBI.pq) يضبط نوع كل عمود.','A ready Power Query script (BregaPowerBI.pq) that sets every column type.'))}</li>
    <li>${esc(T('ملف «اقرأني» بخطوات الفتح.','A read-me with the steps.'))}</li></ul></div>
   <ol class="bi-steps"><li>${esc(T('فك ضغط الحزمة في مجلد.','Unzip the pack into a folder.'))}</li><li>${esc(T('Power BI Desktop ← Get data ← Text/CSV ← اختر الملف ← Load.','Power BI Desktop → Get data → Text/CSV → pick the file → Load.'))}</li><li>${esc(T('للأنواع الدقيقة الصق الاستعلام من BregaPowerBI.pq في Blank query.','For exact types, paste the query from BregaPowerBI.pq into a Blank query.'))}</li></ol>
   <p class="muted bi-note">${esc(T('بيانات داخلية: لا تنشر التقرير أو ملف .pbix للعامة.','Internal data: do not publish the report or the .pbix file.'))}</p>`;
  const foot=`${window.__biDash?`<button type="button" class="btn primary" data-bi="dash">${ICON}${esc(T('تحليل البيانات','Analyse the data'))}</button>`:''}<button type="button" class="btn${window.__biDash?'':' primary'}" data-bi="zip">${I.down}${esc(T('تنزيل حزمة Power BI','Download Power BI pack'))}</button>${p.tables.length===1?`<button type="button" class="btn" data-bi="csv">${esc(T('ملف CSV فقط','CSV only'))}</button>`:''}<button type="button" class="btn" data-bi="close" style="margin-inline-start:auto">${esc(T('إلغاء','Cancel'))}</button>`;
  show(T('تصدير إلى Power BI','Export to Power BI'),p.zipName,body,foot)}
async function save(blob,filename){await S.downloads.save({filename,data:blob});toast(T('تم تنزيل الملف','File downloaded'))}

/* ---------- import counterpart ---------- */
let IMPT=null;
function importDialog(target){IMPT=target;
  const body=`<ol class="bi-steps">
    <li>${esc(T('في تقرير Power BI افتح قائمة «…» على الجدول أو المخطط ← Export data.','In the Power BI report open «…» on the table or visual → Export data.'))}</li>
    <li>${esc(T('اختر «البيانات الملخّصة» بصيغة .xlsx أو .csv.','Choose summarized data as .xlsx or .csv.'))}</li>
    <li>${esc(T('احتفظ بأسماء الأعمدة كما صدّرتها المنظومة (مثل «رقم العقد»)، ثم اختر الملف هنا.','Keep the column names as the system exported them (e.g. «رقم العقد»), then pick the file here.'))}</li></ol>
   <p class="muted bi-note">${esc(T('تُعرض نتيجة القراءة قبل الحفظ كما في استيراد Excel، ولا يُحفظ شيء دون تأكيدك.','The file is checked and shown before anything is saved, as with an Excel import.'))}</p>`;
  const foot=`<button type="button" class="btn primary" data-bi="pick">${I.up}${esc(T('اختيار ملف Power BI','Choose Power BI file'))}</button><button type="button" class="btn" data-bi="close" style="margin-inline-start:auto">${esc(T('إلغاء','Cancel'))}</button>`;
  show(T('استيراد من Power BI','Import from Power BI'),T('ملف .xlsx أو .csv مُصدَّر من Power BI','An .xlsx or .csv exported from Power BI'),body,foot)}

/* a real click on the Excel button itself always means Excel */
document.addEventListener('click',e=>{if(e.isTrusted&&e.target.closest&&e.target.closest('[data-bi-done]'))ARM=0},true);
document.addEventListener('click',async e=>{const b=e.target.closest&&e.target.closest('[data-bi]');if(!b)return;const k=b.dataset.bi;
  if(k==='close'){dialog().close();return}
  if(k==='zip'&&PEND){const p=PEND;try{await save(zipBlob(p),p.zipName);dialog().close()}catch(err){console.error(err);toast(T('تعذّر تجهيز الملف','Could not prepare the file'))}return}
  if(k==='dash'&&PEND&&window.__biDash){const p=PEND;dialog().close();window.__biDash(p.tables,p.title);return}
  if(k==='csv'&&PEND){const p=PEND,t=p.tables[0];try{await save(new Blob([toCsv(t.rows,t.keys)],{type:'text/csv;charset=utf-8'}),p.title+'-PowerBI.csv');dialog().close()}catch(err){console.error(err);toast(T('تعذّر تجهيز الملف','Could not prepare the file'))}return}
  if(k==='pick'&&IMPT){const t=IMPT;dialog().close();if(t.isConnected===false){toast(T('أعد فتح الصفحة ثم حاول مرة أخرى','Reopen the page and try again'));return}t.click();return}
  if(k==='exp'){const src=document.getElementById(b.dataset.biFor);if(!src)return;ARM=Date.now();src.click();return}
  if(k==='imp'){const src=document.getElementById(b.dataset.biFor);if(!src)return;const inp=src.matches('input[type=file]')?src:src.querySelector('input[type=file]');importDialog(inp||src);return}
});

/* ---------- route saveXlsx while a Power BI button is armed ---------- */
let ARM=0;
const __save=saveXlsx;
saveXlsx=async function(sheets,filename){if(ARM&&Date.now()-ARM<4000){ARM=0;exportDialog(sheets,filename);return}return __save.apply(this,arguments)};

/* ---------- CSV parser (kept for callers of window.__biParseCsv) ---------- */
function parseCsv(text){text=text.replace(/^\uFEFF/,'');const first=(text.split(/\r?\n/,1)[0]||'');
  const d=[',',';','\t'].map(c=>[c,first.split(c).length]).sort((a,b)=>b[1]-a[1])[0][0];
  const out=[];let row=[],f='',q=false;
  for(let i=0;i<text.length;i++){const c=text[i];
    if(q){if(c==='"'){if(text[i+1]==='"'){f+='"';i++}else q=false}else f+=c;continue}
    if(c==='"'&&f===''){q=true;continue}
    if(c===d){row.push(f);f='';continue}
    if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(f);out.push(row);row=[];f='';continue}
    f+=c}
  if(f!==''||row.length){row.push(f);out.push(row)}
  return out}
window.__biParseCsv=parseCsv;
/* reading every import format (CSV from Power BI included) is done by fileimport.js */

/* ---------- decorate the Excel buttons ---------- */
let seq=0;
const isExcel=b=>/excel/i.test(b.textContent||'')||b.dataset.dt==='xlsx-tasks'||b.dataset.act==='export-all';
function decorate(){
  const roots=[document.getElementById('view'),document.getElementById('drawer-root'),document.getElementById('dlg')].filter(Boolean);
  roots.forEach(root=>root.querySelectorAll('.btn:not([data-bi-done]):not([data-bi])').forEach(b=>{
    const txt=(b.textContent||'').trim();
    const exp=isExcel(b)&&(b.tagName==='BUTTON')&&(/تصدير|export/i.test(txt)||!!b.querySelector('svg path[d^="M12 4v11"]'));
    const imp=(b.tagName==='LABEL'&&isExcel(b)&&b.querySelector('input[type=file]'))||(b.dataset.sm==='import')||(b.dataset.cx==='import')||(b.dataset.dbt==='import')||(b.dataset.pm==='import');
    if(!exp&&!imp)return;
    b.dataset.biDone='1';if(!b.id)b.id='bi-src-'+(++seq);
    const sm=b.classList.contains('sm');const prim=b.classList.contains('primary');
    const g=document.createElement('span');g.className='bi-grp'+(sm?' sm':'');g.setAttribute('role','group');
    g.setAttribute('aria-label',exp?T('تصدير','Export'):T('استيراد','Import'));
    b.before(g);g.appendChild(b);
    const p=document.createElement('button');p.type='button';p.className='btn bi-btn'+(sm?' sm':'')+(prim?' bi-on-primary':'');
    p.dataset.bi=exp?'exp':'imp';p.dataset.biFor=b.id;
    const lab=exp?T('تصدير إلى Power BI','Export to Power BI'):T('استيراد من Power BI','Import from Power BI');
    p.title=lab;p.setAttribute('aria-label',lab);
    p.innerHTML=ICON+`<span class="bi-lbl">Power BI</span>`;
    g.appendChild(p)}))}
window.__biDecorate=decorate;
const __render=render;render=function(){__render();try{decorate()}catch(e){console.warn('bi',e)}};
let q=0;new MutationObserver(()=>{if(q)return;q=requestAnimationFrame(()=>{q=0;try{decorate()}catch(e){}})}).observe(document.body,{childList:true,subtree:true});
if(window.I18N&&I18N.add)I18N.add({'تصدير إلى Power BI':'Export to Power BI','استيراد من Power BI':'Import from Power BI'});
decorate();
})();
