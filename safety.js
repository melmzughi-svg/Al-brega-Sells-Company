/* ===== حماية حساب المدير والنسخ الاحتياطي الكامل (admin safety & full backup) =====
   1) Two-step sign-in for administrators without a paid plan: when config/security.adminGoogle is on, the rules grant
      admin rights only to a session signed in with Google (which carries Google's own 2-step verification). It can be
      switched on only from a Google session, so the administrator cannot lock themselves out.
   2) Full backup: one JSON file with every collection the administrator can read (attachments optional), recorded in
      config/backup; the administrator is reminded when the last backup is older than 7 days. A backup file can be
      checked (counts per collection) before it is needed. */
"use strict";
(function(){
const me=()=>(window.__me&&window.__me())||null;
const FS=()=>window.__fs&&window.__fs();
const say=m=>{try{toast(m)}catch(e){}};
const roleAdmin=()=>{const m=me();return !!m&&m.role==='admin'};
const DAYS=7;
const COLS=['contracts','needs','batches','extensions','letters','companies','tickets','notices','config','records','assignments','meetings','duties','roles','sms055','sales_report','transfers_report','receipts_report','debts_report','margin','log'];
const SF={sec:null,bk:null,provider:'',loaded:false,busy:false};

async function provider(){const a=window.__auth&&window.__auth();const u=a&&a.currentUser;if(!u)return '';
  try{if(u.getIdTokenResult){const t=await u.getIdTokenResult();if(t&&t.signInProvider)return t.signInProvider}}catch(e){}
  const p=(u.providerData||[]).map(x=>x&&x.providerId);return p.includes('google.com')&&!p.includes('password')?'google.com':(p[0]||'')}
async function load(){const fs=FS();if(!fs||!me())return;
  try{const [s,b]=await Promise.all([fs.doc('config/security').get(),fs.doc('config/backup').get()]);SF.sec=s.exists?s.data():{};SF.bk=b.exists?b.data():null}catch(e){SF.sec=SF.sec||{}}
  SF.provider=await provider();window.__adminGoogle=!!(SF.sec&&SF.sec.adminGoogle);SF.loaded=true}
const enforced=()=>!!(SF.sec&&SF.sec.adminGoogle);
const weak=()=>roleAdmin()&&enforced()&&SF.provider!=='google.com';
const ageDays=()=>SF.bk&&SF.bk.at?Math.floor((Date.now()-Date.parse(SF.bk.at))/86400000):null;
const overdue=()=>{const a=ageDays();return a==null||a>=DAYS};

/* ---------- panels ---------- */
function secPanel(){
  const st=weak()?`<p class="sf-st over">دخلتَ بكلمة المرور، وصلاحيات المدير مقصورة على الدخول بحساب Google. اخرج ثم ادخل بزر «الدخول بحساب Google».</p>`
    :enforced()?`<p class="sf-st ok">مفعّل: صلاحيات المدير لا تعمل إلا لمن دخل بحساب Google.${SF.sec.at?` · ${esc(String(SF.sec.at).slice(0,10))}`:''}</p>`
    :`<p class="sf-st over">غير مفعّل: من يعرف كلمة مرور المدير يحصل على كل الصلاحيات.</p>`;
  const canOn=SF.provider==='google.com';
  return `<div class="panel" id="sf-sec"><h3 class="section-title">حماية حساب المدير <small>التحقق بخطوتين عبر Google</small></h3>
   <p class="muted" style="margin:0 0 8px;font-size:13px">عند التفعيل لا تعمل صلاحيات مدير النظام إلا لمن دخل بحساب Google، فتسري عليه حماية Google بخطوتين. بقية المستخدمين يدخلون كما هم. فعّل التحقق بخطوتين في حساب Google نفسه (myaccount.google.com/security).</p>
   ${st}
   ${enforced()?`<button class="btn ghost" type="button" data-sf="sec-off" ${canOn?'':'disabled'}>إيقاف</button>`
     :`<button class="btn" type="button" data-sf="sec-on" ${canOn?'':'disabled title="ادخل بحساب Google أولًا"'}>تفعيل</button>${canOn?'':'<p class="muted" style="font-size:12.5px;margin:6px 0 0">يُفعَّل من جلسة دخلتَ فيها بحساب Google، حتى لا تُقفل نفسك خارج النظام.</p>'}`}</div>`}
function bkPanel(){
  const a=ageDays();
  const st=SF.bk?`<p class="sf-st ${overdue()?'over':'ok'}">آخر نسخة: ${esc(String(SF.bk.at).slice(0,10))} (${a===0?'اليوم':'قبل '+a+' يوم'}) · ${esc(String(SF.bk.docs||0))} وثيقة${SF.bk.by?` · <span dir="ltr">${esc(SF.bk.by)}</span>`:''}</p>`
    :`<p class="sf-st over">لم تُؤخذ نسخة احتياطية كاملة بعد.</p>`;
  return `<div class="panel" id="sf-bk"><h3 class="section-title">النسخة الاحتياطية الكاملة <small>كل ${DAYS} أيام</small></h3>
   <p class="muted" style="margin:0 0 8px;font-size:13px">الخطة المجانية في Firebase لا تأخذ نسخًا احتياطية تلقائية. هذا الزر يحفظ كل البيانات في ملف JSON واحد على جهازك (العقود والسجلات والتكليفات والاجتماعات والصلاحيات وسجل النشاط والإعدادات وأسماء المسؤولين). احفظ الملف في مكان آمن خارج هذا الجهاز، فهو يحوي بيانات سرية.</p>
   ${st}
   <label style="display:flex;gap:6px;align-items:center;font-size:13px;margin:0 0 8px"><input type="checkbox" id="sf-files"> تضمين المرفقات (الملف أكبر بكثير)</label>
   <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn" type="button" data-sf="backup" ${SF.busy?'disabled':''}>${SF.busy?'جارٍ التحضير…':'أخذ نسخة الآن'}</button>
   <label class="btn ghost">فحص ملف نسخة<input type="file" id="sf-check" accept=".json,application/json" hidden></label></div>
   <div id="sf-res" class="muted" style="margin-top:8px;font-size:13px"></div></div>`}

/* ---------- actions ---------- */
function plain(v){if(v==null)return v;if(typeof v.toDate==='function'){try{return v.toDate().toISOString()}catch(e){}}
  if(Array.isArray(v))return v.map(plain);if(typeof v==='object'){const o={};for(const k in v)o[k]=plain(v[k]);return o}return v}
async function backup(){
  const fs=FS(),m=me();if(!fs||!m||SF.busy)return;SF.busy=true;rp();
  const res=()=>document.getElementById('sf-res');const withFiles=!!(document.getElementById('sf-files')||{}).checked;
  const out={app:'brega-sales',version:window.APP_VERSION||'',at:new Date().toISOString(),by:m.email,collections:{}},counts={},skipped=[];let docs=0;
  try{
    for(const c of COLS.concat(withFiles?['files']:[])){
      if(res())res().textContent='جارٍ قراءة '+c+'…';
      try{const qs=await fs.collection(c).get();const o={};qs.forEach(d=>{o[d.id]=plain(d.data())});out.collections[c]=o;counts[c]=qs.size;docs+=qs.size}
      catch(e){skipped.push(c)}}
    if(!withFiles){try{const qs=await fs.collection('files').get();const o={};qs.forEach(d=>{if(!/_\d+$/.test(d.id))o[d.id]=plain(d.data())});out.collections.filesMeta=o;counts.filesMeta=Object.keys(o).length;docs+=counts.filesMeta}catch(e){skipped.push('files')}}
    const json=JSON.stringify(out);const name='brega-backup-'+out.at.slice(0,10)+'.json';
    const blob=new Blob([json],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000);
    const rec={at:out.at,by:m.email,docs,bytes:json.length,files:withFiles,counts};
    try{await fs.doc('config/backup').set(rec);SF.bk=rec}catch(e){SF.bk=rec}
    try{await fs.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:out.at,user:m.email,name:m.name||'',action:'نسخة احتياطية كاملة',col:'config',id:'backup',summary:docs+' وثيقة'})}catch(e){}
    say('تم حفظ النسخة: '+docs+' وثيقة');
    SF.busy=false;rp();const r=res();if(r&&skipped.length)r.textContent='تعذّرت قراءة: '+skipped.join('، ');
  }catch(e){SF.busy=false;rp();const r=res();if(r)r.textContent=(e&&(e.code||e.message))||'تعذّر أخذ النسخة'}
}
async function check(f){const r=document.getElementById('sf-res');let j;
  try{j=JSON.parse(await f.text())}catch(e){r.textContent='الملف ليس JSON صالحًا';return}
  if(!j||j.app!=='brega-sales'||!j.collections){r.textContent='هذا ليس ملف نسخة احتياطية من المنظومة';return}
  const rows=Object.entries(j.collections).map(([c,o])=>`<tr><td>${esc(c)}</td><td class="n">${Object.keys(o||{}).length}</td></tr>`).join('');
  r.innerHTML=`<p style="margin:0 0 6px">نسخة سليمة: ${esc(String(j.at||'').slice(0,16).replace('T',' '))} · الإصدار ${esc(j.version||'—')} · <span dir="ltr">${esc(j.by||'')}</span></p><div class="tablewrap"><table><thead><tr><th>المجموعة</th><th class="n">الوثائق</th></tr></thead><tbody>${rows}</tbody></table></div>`}
async function setSec(on){const fs=FS(),m=me();if(!fs||!m)return;
  /* تأكيد الهوية بحساب Google نفسه (لا بكلمة المرور، حتى تبقى الجلسة جلسة Google) */
  const u=(window.__auth&&window.__auth().currentUser);
  if(u&&u.reauthenticateWithPopup){try{const p=new firebase.auth.GoogleAuthProvider();p.setCustomParameters({prompt:'select_account'});await u.reauthenticateWithPopup(p)}catch(e){if(e&&e.code!=='auth/popup-closed-by-user')say(window.__authErr?window.__authErr(e):(e.code||e.message));return}}
  SF.provider=await provider();
  const doc={adminGoogle:!!on,by:m.email,at:new Date().toISOString()};
  try{await fs.doc('config/security').set(doc);SF.sec=doc;window.__adminGoogle=!!on;say(on?'فُعّلت حماية حساب المدير':'أوقفت حماية حساب المدير');
    try{await fs.collection('log').add({at:firebase.firestore.FieldValue.serverTimestamp(),ts:doc.at,user:m.email,name:m.name||'',action:on?'تفعيل دخول المديرين بحساب Google':'إيقاف دخول المديرين بحساب Google',col:'config',id:'security',summary:''})}catch(e){}
  }catch(e){say(typeof writeErr==='function'?writeErr(e):(e.code||e.message))}rp()}

function rp(){const a=document.getElementById('sf-sec'),b=document.getElementById('sf-bk');
  if(a)a.outerHTML=secPanel();if(b){b.outerHTML=bkPanel();bindBk()}}
function bindBk(){const c=document.getElementById('sf-check');if(c)c.addEventListener('change',()=>{if(c.files[0])check(c.files[0]);c.value=''})}

let loading=false;
const __render=render;render=function(){__render();if(!roleAdmin())return;const v=document.getElementById('view');if(!v)return;
  if(!SF.loaded){if(!loading){loading=true;load().catch(()=>{}).then(()=>{loading=false;if(SF.loaded)render()})}return}
  if(S.view==='settings'&&!document.getElementById('sf-sec')){v.insertAdjacentHTML('beforeend',secPanel()+bkPanel());bindBk()}
  if(S.view!=='settings'&&!document.getElementById('sf-note')){
    const msg=weak()?'دخلتَ بكلمة المرور: صلاحيات المدير لا تعمل إلا بالدخول بحساب Google.':overdue()?(SF.bk?`مضى ${ageDays()} يوم على آخر نسخة احتياطية كاملة.`:'لم تُؤخذ نسخة احتياطية كاملة بعد.'):'';
    if(msg&&S.view==='home'||weak())v.insertAdjacentHTML('afterbegin',`<div class="sf-st over" id="sf-note" role="status">${esc(msg)} <button type="button" class="btn sm ghost" data-sf="go">الإعدادات</button></div>`)}};
document.addEventListener('click',e=>{const b=e.target.closest('[data-sf]');if(!b||!roleAdmin())return;const k=b.dataset.sf;
  if(k==='backup')backup();else if(k==='sec-on')setSec(true);else if(k==='sec-off')setSec(false);else if(k==='go')go('settings')});

const css=document.createElement('style');css.textContent=`
.sf-st{border-radius:10px;padding:9px 12px;font-size:13.5px;margin:0 0 10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}.sf-st.over{background:var(--warn-soft);color:var(--warn)}.sf-st.ok{background:var(--ok-soft);color:var(--ok)}`;document.head.appendChild(css);
})();
