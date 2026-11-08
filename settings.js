"use strict";
/* ---- ZIP helpers ---- */
const CRC_T=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC_T[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
function zipStore(files){const enc=new TextEncoder();const parts=[],cd=[];let off=0;
  files.forEach(([name,str])=>{const nb=enc.encode(name),db=enc.encode(str),crc=crc32(db);
    const h=new DataView(new ArrayBuffer(30));h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(8,0,true);h.setUint32(14,crc,true);h.setUint32(18,db.length,true);h.setUint32(22,db.length,true);h.setUint16(26,nb.length,true);
    parts.push(new Uint8Array(h.buffer),nb,db);
    const c=new DataView(new ArrayBuffer(46));c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);c.setUint32(16,crc,true);c.setUint32(20,db.length,true);c.setUint32(24,db.length,true);c.setUint16(28,nb.length,true);c.setUint32(42,off,true);
    cd.push(new Uint8Array(c.buffer),nb);off+=30+nb.length+db.length});
  const cdSize=cd.reduce((s,a)=>s+a.length,0);const e=new DataView(new ArrayBuffer(22));e.setUint32(0,0x06054b50,true);e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,cdSize,true);e.setUint32(16,off,true);
  return new Blob([...parts,...cd,new Uint8Array(e.buffer)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
async function unzip(buf){const u8=new Uint8Array(buf),dv=new DataView(buf);let e=u8.length-22;while(e>=0&&dv.getUint32(e,true)!==0x06054b50)e--;if(e<0)throw new Error('zip');
  const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);const out={};const dec=new TextDecoder();
  for(let i=0;i<n;i++){const method=dv.getUint16(p+10,true),csize=dv.getUint32(p+20,true),nl=dv.getUint16(p+28,true),xl=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true);
    const name=dec.decode(u8.subarray(p+46,p+46+nl));const lnl=dv.getUint16(lo+26,true),lxl=dv.getUint16(lo+28,true);const data=u8.subarray(lo+30+lnl+lxl,lo+30+lnl+lxl+csize);
    out[name]={method,data};p+=46+nl+xl+cl}
  return {names:Object.keys(out),async text(name){const f=out[name];if(!f)return null;if(f.method===0)return dec.decode(f.data);
    const ds=new DecompressionStream('deflate-raw');const s=new Blob([f.data]).stream().pipeThrough(ds);return await new Response(s).text()},
    async bytes(name){const f=out[name];if(!f)return null;if(f.method===0)return f.data;
    const ds=new DecompressionStream('deflate-raw');const s=new Blob([f.data]).stream().pipeThrough(ds);return new Uint8Array(await new Response(s).arrayBuffer())}}}
/* ---- XLSX write ---- */
function xe(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'')}
function colL(i){let s='';i++;while(i){const r=(i-1)%26;s=String.fromCharCode(65+r)+s;i=Math.floor((i-1)/26)}return s}
function buildXlsx(sheets){const n=sheets.length;const names=[];
  sheets.forEach(([nm],i)=>{let x=String(nm).replace(/[\[\]\*\?\/\\:]/g,' ').slice(0,31)||('Sheet'+(i+1));while(names.includes(x))x=x.slice(0,28)+i;names.push(x)});
  const files=[['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>'],
   ['_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
   ['xl/workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>'+names.map((nm,i)=>`<sheet name="${xe(nm)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')+'</sheets></workbook>'],
   ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
   ['xl/styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0063AD"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>']];
  sheets.forEach(([,rows],si)=>{const keys=[];rows.forEach(r=>Object.keys(r).forEach(k=>{if(!keys.includes(k))keys.push(k)}));
    const cell=(ref,v,st)=>{if(v==null||v==='')return'';if(typeof v==='number'&&isFinite(v))return `<c r="${ref}" s="${st||2}"><v>${v}</v></c>`;return `<c r="${ref}" t="inlineStr"${st?` s="${st}"`:''}><is><t xml:space="preserve">${xe(v)}</t></is></c>`};
    let x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>';
    keys.forEach((k,i)=>{const w=Math.max(k.length,...rows.slice(0,300).map(r=>String(r[k]??'').length));x+=`<col min="${i+1}" max="${i+1}" width="${Math.min(50,Math.max(10,w+2))}" customWidth="1"/>`});
    x+='</cols><sheetData><row r="1">'+keys.map((k,i)=>cell(colL(i)+'1',k,1)).join('')+'</row>';
    rows.forEach((r,ri)=>{x+=`<row r="${ri+2}">`+keys.map((k,i)=>cell(colL(i)+(ri+2),r[k])).join('')+'</row>'});
    x+='</sheetData>'+(keys.length&&rows.length?`<autoFilter ref="A1:${colL(keys.length-1)}${rows.length+1}"/>`:'')+'</worksheet>';
    files.push([`xl/worksheets/sheet${si+1}.xml`,x])});
  return zipStore(files)}
saveXlsx=async function(sheets,filename){try{const blob=buildXlsx(sheets.map(([n,r])=>[n,r.length?r:[{'':'لا توجد بيانات'}]]));await S.downloads.save({filename,data:blob});toast('تم تنزيل الملف')}catch(e){console.error(e);toast('تعذّر تجهيز الملف')}};
/* ---- reading imports: readWorkbook() and every file format live in fileimport.js ---- */

/* ===== Cloud extras: banners, initial data load, roles (users), activity log ===== */
"use strict";
try{Object.assign(ARTICLES,JSON.parse(document.getElementById('articles-data').textContent))}catch(e){}
loadArticles=async function(){};
dbBanner=function(){
  const me=window.__me();
  if(S.dbState==='wait')return `<div class="banner info">جارٍ تحميل البيانات…</div>`;
  if(me&&me.role==='admin'&&S.loadedOnce&&!S.contracts.length&&!S.needs.length)return `<div class="banner info">قاعدة البيانات فارغة. <button class="btn primary sm" data-act="seed-now" style="margin-inline-start:8px">استيراد ملف البيانات الأولية (JSON)</button></div>`;
  if(!S.canWrite)return `<div class="banner info">حسابك بصلاحية «مشاهد»: يمكنك البحث والاطلاع والتصدير. لطلب صلاحية التعديل تواصل مع مدير النظام.</div>`;
  return '';
};
const __vs=viewSettings;
viewSettings=function(){const me=window.__me()||{};const admin=me.role==='admin';
  return __vs()+`<div class="grid2"><div class="panel"><h3 class="section-title">حسابي <small dir="ltr">${esc(me.email||'')}</small> <small>· ${esc(window.__roleAr[me.role]||'')}</small></h3>
   <p class="muted" style="margin:0 0 8px;font-size:13px">غيّر كلمة المرور من هنا مباشرة؛ ستُطلب منك إعادة تأكيد هويتك أولًا (بنافذة Google أو بكلمة المرور الحالية). إن كنت تدخل بحساب Google فقط فكلمة المرور اختيارية، وتعيينها يتيح لك الدخول بها أيضًا.</p>
   <form id="pw-form" novalidate style="display:grid;gap:8px;max-width:360px"><div class="field"><label for="pw-new">كلمة المرور الجديدة (8 أحرف على الأقل)</label><input id="pw-new" type="password" autocomplete="new-password" dir="ltr" minlength="8"></div><div class="field"><label for="pw-conf">تأكيد كلمة المرور</label><input id="pw-conf" type="password" autocomplete="new-password" dir="ltr"></div><div><button class="btn primary" type="submit" id="pw-save">تغيير كلمة المرور</button> <span id="pw-msg" role="status" style="font-size:12.5px"></span></div></form>
   <p class="muted" style="margin:8px 0 0;font-size:12px">أو <a href="#" data-act="pw-reset">أرسل رابط التغيير إلى البريد</a> (قد لا تصل الرسالة إلى بعض صناديق البريد).</p></div>
  ${admin&&window.FuelUnits?FuelUnits.panel():''}
  ${admin?`<div class="panel"><h3 class="section-title">النسخ الاحتياطي</h3><p class="muted" style="margin:0 0 8px;font-size:13px">ملف Excel من قسم التصدير أعلاه للقراءة والمشاركة. أما النسخة الاحتياطية الكاملة لكل البيانات فمن لوحة «النسخة الاحتياطية الكاملة» أسفل هذه الصفحة.</p></div>`:''}</div>
  ${admin?`<div class="panel"><h3 class="section-title">المستخدمون والصلاحيات</h3>
   <p class="muted" style="margin:0 0 10px;font-size:13px">يُنشئ كل زميل حسابه بنفسه من شاشة الدخول (إنشاء حساب)، ثم تمنحه الصلاحية هنا ببريده الإلكتروني. <b>مدير النظام</b>: كل الصلاحيات · <b>محرر</b>: إضافة وتعديل وحذف · <b>مشاهد</b>: اطلاع وبحث وتصدير فقط.</p>
   <div id="users-box" class="muted">جارٍ التحميل…</div>
   <fieldset style="margin-top:14px"><legend>منح صلاحية أو تعديلها</legend>
    <div class="field"><label for="u-email">البريد الإلكتروني</label><input id="u-email" type="email" dir="ltr"></div>
    <div class="field"><label for="u-name">الاسم</label><input id="u-name"></div>
    <div class="field"><label for="u-role">الصلاحية</label><select id="u-role"><option value="editor">محرر</option><option value="viewer">مشاهد</option><option value="admin">مدير النظام</option></select></div>
    <div class="field"><label for="u-unit">الإدارة</label><select id="u-unit"><option value="">— بلا إدارة (لا يرى سجلات ولا تكليفات) —</option>${ROLE_UNITS.map(u=>`<option value="${u}">${esc((ORG_BY.get(u)||{}).title||u)}</option>`).join('')}</select></div>
    <div style="grid-column:1/-1;display:flex;gap:8px;align-items:center"><button class="btn primary" data-act="save-role">حفظ</button><span id="u-msg" style="font-size:12.5px"></span></div></fieldset></div>
  <div class="panel"><h3 class="section-title">سجل النشاط <small>آخر 200 عملية</small></h3><div id="log-box" class="muted">جارٍ التحميل…</div></div>`:''}`};
const ROLE_UNITS=['board','gm','gm-rep','gm-fo','ca','w','e'].filter(u=>ORG_BY.get(u));
const __as=afterRender.settings;
afterRender.settings=()=>{__as();const me=window.__me()||{};if(me.role==='admin'){loadRoles();loadLog()}
  const pf=document.getElementById('pw-form');if(pf)pf.addEventListener('submit',changePassword);};
async function changePassword(e){e.preventDefault();const m=document.getElementById('pw-msg'),b=document.getElementById('pw-save'),n=document.getElementById('pw-new'),c=document.getElementById('pw-conf');const pw=n.value,pw2=c.value;
  const say=(t,ok)=>{m.style.color=ok?'var(--ok)':'var(--crit)';m.textContent=t};
  if(pw.length<8){say('كلمة المرور يجب ألا تقل عن 8 أحرف');n.focus();return}
  if(pw!==pw2){say('كلمتا المرور غير متطابقتين');c.focus();return}
  b.disabled=true;say('');
  try{const ok=await window.__reauth('لتغيير كلمة المرور أكّد هويتك أولًا.',true);if(!ok){say('لم يتم التأكيد، لم تتغير كلمة المرور');return}
    const u=window.__auth().currentUser;await u.updatePassword(pw);n.value='';c.value='';say('تم تغيير كلمة المرور. استخدمها في الدخول القادم مع بريدك.',true);
    try{const me=window.__me();await window.__fs().collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:me.email,name:me.name||'',action:'تغيير كلمة المرور',col:'settings',id:me.email,summary:'غيّر المستخدم كلمة مروره من «حسابي»'})}catch(err){}
  }catch(err){say((window.__authErr&&window.__authErr(err))||err.message||'تعذّر تغيير كلمة المرور')}finally{b.disabled=false}}
async function loadRoles(){const box=document.getElementById('users-box');if(!box)return;try{const qs=await window.__fs().collection('roles').get();const me=window.__me();const rows=[];qs.forEach(d=>rows.push(Object.assign({email:d.id},d.data())));
  box.innerHTML=`<div class="tablewrap" style="max-height:none;min-height:0"><table><thead><tr><th class="nosort">الاسم</th><th class="nosort">البريد</th><th class="nosort">الصلاحية</th><th class="nosort">الإدارة</th><th class="nosort"></th></tr></thead><tbody>${rows.map(u=>`<tr style="cursor:default"><td class="t-strong">${esc(u.name||'')}</td><td dir="ltr" style="text-align:right">${esc(u.email)}</td><td><span class="pill ${u.role==='admin'?'p-fl':u.role==='editor'?'p-ok':'p-neutral'}">${esc(window.__roleAr[u.role]||u.role)}</span></td><td>${u.unit?esc((ORG_BY.get(u.unit)||{}).title||u.unit):`<span class="pill p-crit">بلا إدارة</span>`}</td><td><button class="btn sm ghost" data-act="edit-role" data-e="${esc(u.email)}" data-n="${esc(u.name||'')}" data-r="${esc(u.role)}" data-u="${esc(u.unit||'')}">${I.edit}تعديل</button>${u.email!==me.email?`<button class="btn sm danger" data-act="del-role" data-e="${esc(u.email)}">إزالة</button>`:''}</td></tr>`).join('')}</tbody></table></div>`}catch(e){box.textContent='تعذّر التحميل: '+(e.message||'')}}
async function loadLog(){const box=document.getElementById('log-box');if(!box)return;try{const qs=await window.__fs().collection('log').orderBy('ts','desc').limit(200).get();const colAr={contracts:'عقد',needs:'منشأة',batches:'دفعة',extensions:'تمديد',letters:'مراسلة',config:'إعدادات',roles:'مستخدم',tickets:'طلب زبون',notices:'إشعار بيع',companies:'زبون'};const rows=[];qs.forEach(d=>rows.push(d.data()));
  box.innerHTML=rows.length?`<div class="tablewrap" style="max-height:420px;min-height:0"><table><thead><tr><th class="nosort">الوقت</th><th class="nosort">المستخدم</th><th class="nosort">العملية</th><th class="nosort">النوع</th><th class="nosort">السجل</th></tr></thead><tbody>${rows.map(r=>`<tr style="cursor:default"><td class="n">${esc((r.ts||'').slice(0,16).replace('T',' '))}</td><td>${esc(r.name||r.user||'')}</td><td>${esc(r.action)}</td><td>${esc(colAr[r.col]||r.col||'')}</td><td>${esc(r.summary||'')}</td></tr>`).join('')}</tbody></table></div>`:'لا توجد عمليات بعد.'}catch(e){box.textContent='تعذّر التحميل: '+(e.message||'')}}
document.addEventListener('click',async e=>{const a=e.target.closest('[data-act]');if(!a)return;const act=a.dataset.act;const FS=window.__fs();
  if(act==='seed-now'){const inp=document.createElement('input');inp.type='file';inp.accept='.json,application/json';inp.id='seed-file';inp.hidden=true;document.body.appendChild(inp);inp.onchange=async()=>{const f=inp.files&&inp.files[0];inp.remove();if(!f)return;const t=a.textContent;a.disabled=true;a.textContent='جارٍ التحميل…';try{const n=await window.__seedIfEmpty(f);toast('تم تحميل '+n+' سجل')}catch(err){toast('تعذّر التحميل: '+(err.message||''));a.disabled=false;a.textContent=t}};inp.click()}
  if(act==='pw-reset'){e.preventDefault();const m=document.getElementById('pw-msg');try{await window.__auth().sendPasswordResetEmail(window.__me().email);m.style.color='var(--ok)';m.textContent='أُرسلت الرسالة إلى بريدك'}catch(err){m.style.color='var(--crit)';m.textContent=err.message}}
  if(act==='edit-role'){document.getElementById('u-email').value=a.dataset.e;document.getElementById('u-name').value=a.dataset.n;document.getElementById('u-role').value=a.dataset.r;const us=document.getElementById('u-unit');if(us)us.value=a.dataset.u||'';document.getElementById('u-name').focus()}
  if(act==='save-role'){const m=document.getElementById('u-msg');const em=val('u-email').toLowerCase();if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)){m.style.color='var(--crit)';m.textContent='أدخل بريدًا إلكترونيًا صحيحًا';return}
    const newRole=document.getElementById('u-role').value;if(!['admin','editor','viewer'].includes(newRole)){m.style.color='var(--crit)';m.textContent='صلاحية غير معروفة';return}
    if(em===window.__me().email&&newRole!=='admin'){m.style.color='var(--crit)';m.textContent='لا يمكنك تخفيض صلاحيتك بنفسك حتى لا يُقفل النظام؛ اطلب ذلك من مدير آخر.';return}
    if(!(await window.__reauth('منح الصلاحيات إجراء حساس، أدخل كلمة المرور للمتابعة.'))){m.style.color='var(--crit)';m.textContent='أُلغي الإجراء';return}
    const unit=(document.getElementById('u-unit')||{}).value||'';if(unit&&!ROLE_UNITS.includes(unit)){m.style.color='var(--crit)';m.textContent='إدارة غير معروفة';return}
    try{await FS.collection('roles').doc(em).set(Object.assign({},unit?{unit}:{},{name:val('u-name').slice(0,120),role:newRole}));await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:window.__me().email,name:window.__me().name,action:'منح صلاحية',col:'roles',id:em,summary:val('u-name')+' ('+window.__roleAr[document.getElementById('u-role').value]+(unit?' · '+((ORG_BY.get(unit)||{}).title||unit):'')+')'});m.style.color='var(--ok)';m.textContent=unit?'تم الحفظ. يستطيع المستخدم الدخول الآن ببريده.':'تم الحفظ بلا إدارة: لن يرى سجلات ولا تكليفات حتى تحدد إدارته.';['u-email','u-name'].forEach(i=>document.getElementById(i).value='');loadRoles();loadLog()}catch(err){m.style.color='var(--crit)';m.textContent=err.message}}
  if(act==='del-role'){if(!a.classList.contains('armed')){a.classList.add('armed');a.textContent='تأكيد الإزالة';setTimeout(()=>{a.classList.remove('armed');a.textContent='إزالة'},4000);return}if(a.dataset.e===window.__me().email){toast('لا يمكنك إزالة حسابك بنفسك');return}if(!(await window.__reauth('إزالة صلاحية إجراء حساس، أدخل كلمة المرور للمتابعة.')))return;try{await FS.collection('roles').doc(a.dataset.e).delete();await FS.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:new Date().toISOString(),user:window.__me().email,name:window.__me().name,action:'إزالة صلاحية',col:'roles',id:a.dataset.e,summary:a.dataset.e});loadRoles();loadLog()}catch(err){toast(err.message)}}
});
