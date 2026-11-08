/* ===== المرفقات والمسح الضوئي للمستندات (Attachments & document scanner) =====
   Loaded after the main app script. Files live in the Firestore collection `files`: files/<id> holds
   the metadata and files/<id>_<i> hold base64 chunks (each under the 1 MB document limit), so no
   Firebase Storage bucket is needed. Records keep a light `attachments` list [{id,name,type,size,n,pages,scanned,at,by}].
   The scanner runs entirely in the browser: live document detection on the camera feed, perspective
   correction, enhancement filters, multi-page review with drag-to-reorder, and a built-in PDF writer. */
"use strict";
(function(){
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()?en:ar;
const H=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const say=m=>{try{toast(m)}catch(e){console.log(m)}};
const FSX=()=>window.__fs?window.__fs():null;
const ME=()=>(window.__me&&window.__me())||{};
const CHUNK=700000,MAX_MB=15,WORK=2200;
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));
const mk=(w,h)=>{const c=document.createElement('canvas');c.width=Math.max(1,w);c.height=Math.max(1,h);return c};
const cx=c=>c.getContext('2d',{willReadFrequently:true});
const jpeg=(c,q)=>new Promise(r=>c.toBlob(r,'image/jpeg',q));
const fsize=b=>b<1024?b+' B':b<1048576?Math.round(b/1024)+' KB':(b/1048576).toFixed(b<10485760?1:0)+' MB';
const vib=n=>{try{navigator.vibrate&&navigator.vibrate(n)}catch(e){}};
const clamp=(v,a,b)=>v<a?a:v>b?b:v;

const sv=(p,w=2)=>`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`;
const IC={
 clip:sv('<path d="m21.4 11.1-9.2 9.2a6 6 0 0 1-8.5-8.5l9.2-9.2a4 4 0 0 1 5.7 5.7l-9.2 9.2a2 2 0 0 1-2.8-2.8l8.5-8.5"/>'),
 scan:sv('<path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2"/><path d="M7 12h10"/>'),
 up:sv('<path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>'),
 x:sv('<path d="M18 6 6 18M6 6l12 12"/>'),
 check:sv('<path d="m5 12.5 4.5 4.5L19 7.5"/>',2.4),
 rot:sv('<path d="M20 12a8 8 0 1 1-2.4-5.7L20 8.5"/><path d="M20 3.5v5h-5"/>'),
 crop:sv('<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>'),
 trash:sv('<path d="M3 6h18M8 6V4h8v2M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6M10 11v6M14 11v6"/>'),
 dl:sv('<path d="M12 4v12m0 0-4.5-4.5M12 16l4.5-4.5M4 20h16"/>'),
 ext:sv('<path d="M14 4h6v6M20 4l-9 9M18 14v4a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4"/>'),
 bolt:sv('<path d="M13 2 4 14h7l-1 8 9-12h-7z"/>'),
 sw:sv('<path d="M4 7h11a5 5 0 0 1 5 5M20 17H9a5 5 0 0 1-5-5"/><path d="m17 4 3 3-3 3M7 14l-3 3 3 3"/>'),
 img:sv('<rect x="3" y="3" width="18" height="18" rx="3"/><circle cx="9" cy="9" r="2"/><path d="m21 15-5-5L5 21"/>'),
 plus:sv('<path d="M12 5v14M5 12h14"/>'),
 back:sv('<path d="m15 18-6-6 6-6"/>'),
 next:sv('<path d="m9 18 6-6-6-6"/>'),
 wand:sv('<path d="m14.5 5.5 4 4L8 20l-4-4z"/><path d="M12 8l4 4M5 2.5v3M3.5 4h3M19 15.5v3M17.5 17h3M10.5 2.5v2M9.5 3.5h2"/>'),
 auto:sv('<path d="M12 2.5v3M12 18.5v3M2.5 12h3M18.5 12h3"/><circle cx="12" cy="12" r="5"/><path d="m10 12 1.5 1.5L14.5 10.5"/>'),
 cam:sv('<path d="M4 8h3l2-3h6l2 3h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13.5" r="3.5"/>'),
 full:sv('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'),
 file:sv('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>'),
 grip:sv('<circle cx="9" cy="6" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="18" r="1"/>',2.4),
 shield:sv('<path d="M12 3 4 6v6c0 4.5 3.4 8.2 8 9 4.6-.8 8-4.5 8-9V6z"/><path d="m9 12 2 2 4-4"/>')
};

/* ---------------- storage (Firestore, chunked) ---------------- */
const cache=new Map();
const toB64=blob=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result).split(',')[1]||'');r.onerror=()=>rej(r.error);r.readAsDataURL(blob)});
/* scope: who may open the file, inherited from the record it belongs to (see firestore.rules → files).
   {k:'all'} | {k:'units',u:[...]} | {k:'meeting',id}. Missing scope means a record everyone with a role reads. */
const normScope=sc=>{sc=typeof sc==='function'?sc():sc;if(!sc||!sc.k)return {k:'all'};
  if(sc.k==='units')return {k:'units',u:[...new Set((sc.u||[]).filter(Boolean).map(String))].slice(0,60)};
  if(sc.k==='meeting'&&sc.id)return {k:'meeting',id:String(sc.id)};return {k:'all'}};
async function upload(blob,name,extra,onProg,scope){
  const fs=FSX();if(!fs)throw new Error(tr('قاعدة البيانات غير متصلة','Database not connected'));
  if(blob.size>MAX_MB*1048576)throw new Error(tr(`الحجم أكبر من ${MAX_MB} ميغابايت`,`larger than ${MAX_MB} MB`));
  const data=await toB64(blob);const n=Math.max(1,Math.ceil(data.length/CHUNK));
  const id=fs.collection('files').doc().id;
  const m=ME();const meta=Object.assign({name,type:blob.type||'application/octet-stream',size:blob.size,n,at:new Date().toISOString(),by:m.name||m.email||''},extra||{});
  // the metadata (owner + scope) goes first: the rules check it before accepting each chunk
  await fs.collection('files').doc(id).set(Object.assign({},meta,{byEmail:m.email||'',scope:normScope(scope)}));
  try{for(let i=0;i<n;i++){await fs.collection('files').doc(id+'_'+i).set({d:data.slice(i*CHUNK,(i+1)*CHUNK)});onProg&&onProg((i+1)/(n+.25))}}
  catch(e){removeFile({id,n});throw e}
  cache.set(id,blob);onProg&&onProg(1);return Object.assign({id},meta);
}
async function getBlob(a){
  if(cache.has(a.id))return cache.get(a.id);
  const fs=FSX();if(!fs)throw new Error(tr('قاعدة البيانات غير متصلة','Database not connected'));
  let n=a.n;if(!n){const d=await fs.collection('files').doc(a.id).get();if(!d.exists)throw new Error(tr('الملف غير موجود','File not found'));n=d.data().n||1}
  const parts=await Promise.all(Array.from({length:n},(_,i)=>fs.collection('files').doc(a.id+'_'+i).get()));
  if(parts.some(p=>!p.exists))throw new Error(tr('الملف غير مكتمل','File is incomplete'));
  const bin=atob(parts.map(p=>p.data().d).join(''));const u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
  const b=new Blob([u],{type:a.type||'application/octet-stream'});cache.set(a.id,b);return b;
}
async function removeFile(a){const fs=FSX();if(!fs||!a)return;try{await Promise.all(Array.from({length:a.n||1},(_,i)=>fs.collection('files').doc(a.id+'_'+i).delete()));await fs.collection('files').doc(a.id).delete()}catch(e){}cache.delete(a.id)}
function saveBlob(blob,name){const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name||'file';document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(u);a.remove()},4000)}

async function decode(blob){
  if(window.createImageBitmap){try{return await createImageBitmap(blob)}catch(e){}}
  return new Promise((res,rej)=>{const u=URL.createObjectURL(blob);const i=new Image();i.onload=()=>{res(i);setTimeout(()=>URL.revokeObjectURL(u),1000)};i.onerror=()=>{URL.revokeObjectURL(u);rej(new Error('decode'))};i.src=u});
}
function scaleTo(src,max){const w=src.width,h=src.height;const s=Math.min(1,max/Math.max(w,h));const c=mk(Math.round(w*s),Math.round(h*s));const x=c.getContext('2d');x.imageSmoothingQuality='high';x.drawImage(src,0,0,c.width,c.height);return c}
async function compressImage(blob,name){
  if(!/^image\/(jpeg|png|webp|bmp)$/i.test(blob.type||''))return {blob,name};
  try{const b=await decode(blob);const big=Math.max(b.width,b.height)>2400;if(!big&&blob.size<900*1024)return {blob,name};
    const c=scaleTo(b,2400);const x=c.getContext('2d');x.globalCompositeOperation='destination-over';x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);
    const out=await jpeg(c,.84);if(out&&out.size<blob.size)return {blob:out,name:name.replace(/\.[^.]+$/,'')+'.jpg'}}catch(e){}
  return {blob,name};
}

/* ---------------- attachment field ---------------- */
const ST=new Map();let seq=0;
function kind(a){const t=(a.type||'').toLowerCase(),n=(a.name||'').toLowerCase();if(t.includes('pdf')||n.endsWith('.pdf'))return 'pdf';if(t.startsWith('image/'))return 'img';if(/\.(docx?|rtf|odt)$/.test(n)||t.includes('word'))return 'doc';if(/\.(xlsx?|csv|ods)$/.test(n)||t.includes('sheet')||t.includes('excel'))return 'xls';return 'file'}
const KLAB={pdf:'PDF',img:'IMG',doc:'DOC',xls:'XLS',file:'FILE'};
const pagesTxt=n=>EN()?`${n} page${n>1?'s':''}`:(n===1?'صفحة واحدة':n===2?'صفحتان':n<=10?`${n} صفحات`:`${n} صفحة`);
function itemHTML(a,key,ro){
  const k=kind(a);const sub=[fsize(a.size||0),a.pages?pagesTxt(a.pages):'',a.at?new Date(a.at).toLocaleDateString('en-GB'):''].filter(Boolean).join(' · ');
  return `<div class="att-item${a.scanned?' scanned':''}"><button type="button" class="att-open" data-att-act="open" data-key="${H(key)}" data-id="${H(a.id)}" title="${H(tr('عرض','View'))}">
   <span class="att-ic k-${k}">${a.scanned?IC.scan:''}<b>${KLAB[k]}</b></span><span class="att-meta"><b dir="auto">${H(a.name)}</b><small>${a.scanned?`<em>${H(tr('مسح ضوئي','Scanned'))}</em> · `:''}${H(sub)}</small></span></button>
   <button type="button" class="att-x" data-att-act="dl" data-key="${H(key)}" data-id="${H(a.id)}" title="${H(tr('تنزيل','Download'))}" aria-label="${H(tr('تنزيل','Download'))}">${IC.dl}</button>
   ${ro?'':`<button type="button" class="att-x rm" data-att-act="rm" data-key="${H(key)}" data-id="${H(a.id)}" title="${H(tr('إزالة','Remove'))}" aria-label="${H(tr('إزالة','Remove'))}">${IC.trash}</button>`}</div>`;
}
function pendHTML(pid,p){return `<div class="att-item pend" data-pend="${pid}"><span class="att-ic k-${p.k||'file'}"><b>${KLAB[p.k||'file']}</b></span><span class="att-meta"><b dir="auto">${H(p.name)}</b><small class="att-ptxt">${H(p.txt)}</small><span class="att-bar"><i style="width:${Math.round(p.p*100)}%"></i></span></span></div>`}
function listHTML(key){const s=ST.get(key);if(!s)return '';const h=s.list.map(a=>itemHTML(a,key,s.ro)).join('')+[...s.pending.entries()].map(([pid,p])=>pendHTML(pid,p)).join('');return h||(s.ro?`<div class="att-empty">${H(tr('لا توجد مرفقات','No attachments'))}</div>`:'')}
function paint(key){document.querySelectorAll('[data-att]').forEach(el=>{if(el.dataset.att===key)el.querySelector('.att-list').innerHTML=listHTML(key)})}
function prog(key,pid,p,txt){const s=ST.get(key);if(!s)return;const e=s.pending.get(pid);if(!e)return;e.p=p;if(txt)e.txt=txt;else e.txt=tr('جارٍ الرفع… ','Uploading… ')+Math.round(p*100)+'%';
  const el=document.querySelector(`[data-pend="${pid}"]`);if(el){el.querySelector('i').style.width=Math.round(p*100)+'%';el.querySelector('.att-ptxt').textContent=e.txt}else paint(key)}
function field(key,list,opts={}){
  let s=ST.get(key);
  if(!s||!s.pending.size||opts.reset)s={list:(list||[]).slice(),pending:new Map(),fresh:new Set(),waiters:[]};
  s.onChange=opts.onChange||null;s.scope=opts.scope||null;s.ro=!!opts.readonly;s.name=opts.name||'';ST.set(key,s);
  const zone=s.ro?'':`<div class="att-zone" data-att-drop="${H(key)}"><span class="att-zi">${IC.clip}</span>
    <span class="att-zt"><b>${H(tr('أرفق المستندات','Attach documents'))}</b><small>${H(tr('اسحب الملفات وأفلتها هنا، أو امسح المستند الورقي بكاميرا الجهاز','Drop files here, or scan a paper document with your camera'))}</small></span>
    <span class="att-zb"><button type="button" class="att-btn" data-att-act="pick" data-key="${H(key)}">${IC.up}${H(tr('رفع ملف','Upload file'))}</button><button type="button" class="att-btn att-scan" data-att-act="scan" data-key="${H(key)}">${IC.scan}${H(tr('مسح ضوئي','Scan'))}</button></span></div>`;
  return `<div class="att${s.ro?' ro':''}" data-att="${H(key)}">${zone}<div class="att-list">${listHTML(key)}</div></div>`;
}
function view(list,opts={}){if(!list||!list.length)return opts.empty?`<div class="att-empty">${H(tr('لا توجد مرفقات','No attachments'))}</div>`:'';const key='v'+(++seq);ST.set(key,{list:list.slice(),pending:new Map(),fresh:new Set(),ro:true,waiters:[]});
  if(opts.compact)return `<span class="att-chips">${list.map(a=>`<button type="button" class="att-chip" data-att-act="open" data-key="${key}" data-id="${H(a.id)}" title="${H(a.name)}">${IC.clip}<span dir="auto">${H(a.name)}</span></button>`).join('')}</span>`;
  return `<div class="att ro" data-att="${key}"><div class="att-list">${listHTML(key)}</div></div>`}
const clean=a=>{const o={};['id','name','type','size','n','pages','scanned','at','by'].forEach(k=>{if(a[k]!=null&&a[k]!=='')o[k]=a[k]});return o};
async function val(key,msgId){const s=ST.get(key);if(!s)return [];
  if(s.pending.size){const m=msgId&&document.getElementById(msgId);if(m)m.textContent=tr('جارٍ إكمال رفع المرفقات…','Finishing attachment upload…');while(s.pending.size)await sleep(200);if(m)m.textContent=''}
  return s.list.map(clean)}
async function addBlobs(key,items){
  const s=ST.get(key);if(!s||!items.length)return;let ok=0;
  await Promise.all(items.map(async it=>{
    const pid='p'+(++seq);s.pending.set(pid,{name:it.name,p:0,k:kind({type:it.blob.type,name:it.name}),txt:tr('جارٍ التجهيز…','Preparing…')});paint(key);
    try{let blob=it.blob,name=it.name;if(!(it.extra&&it.extra.scanned)){const c=await compressImage(blob,name);blob=c.blob;name=c.name}
      const a=await upload(blob,name,it.extra,p=>prog(key,pid,p),s.scope);s.list.push(a);s.fresh.add(a.id);ok++}
    catch(e){say(tr('تعذّر رفع «','Could not upload “')+it.name+tr('»: ','”: ')+(e.code==='permission-denied'?tr('لا تملك صلاحية الرفع (حدّث قواعد Firestore)','permission denied (update Firestore rules)'):(e.message||e.code||'')))}
    s.pending.delete(pid);paint(key);
  }));
  if(ok&&s.onChange){try{await s.onChange(s.list.map(clean))}catch(e){say(tr('تعذّر حفظ المرفقات','Could not save attachments'))}}
  if(ok)say(ok>1?tr(`تم إرفاق ${ok} ملفات`,`${ok} files attached`):tr('تم إرفاق الملف','File attached'));
}
function pickFiles(key){const i=document.createElement('input');i.type='file';i.multiple=true;i.accept='.pdf,image/*,.doc,.docx,.xls,.xlsx,.csv,.txt';
  i.onchange=()=>{const f=[...i.files];if(f.length)addBlobs(key,f.map(x=>({blob:x,name:x.name})))};i.click()}
document.addEventListener('click',async e=>{
  const b=e.target.closest('[data-att-act]');if(!b)return;e.preventDefault();const key=b.dataset.key,s=ST.get(key);if(!s)return;const act=b.dataset.attAct;
  if(act==='pick')return pickFiles(key);
  if(act==='scan'){const r=await scanOpen({name:s.name});if(r)addBlobs(key,[{blob:r.blob,name:r.name,extra:{pages:r.pages,scanned:true}}]);return}
  const i=s.list.findIndex(x=>x.id===b.dataset.id);const a=s.list[i];if(!a)return;
  if(act==='open')return viewer(s.list,i);
  if(act==='dl'){try{b.classList.add('spin');saveBlob(await getBlob(a),a.name)}catch(err){say(err.message||'')}finally{b.classList.remove('spin')}return}
  if(act==='rm'){if(!b.classList.contains('armed')){b.classList.add('armed');b.title=tr('اضغط مرة أخرى للتأكيد','Click again to confirm');setTimeout(()=>b.isConnected&&b.classList.remove('armed'),3000);return}
    s.list.splice(i,1);paint(key);
    if(s.onChange){try{await s.onChange(s.list.map(clean));removeFile(a);say(tr('أُزيل المرفق','Attachment removed'))}catch(err){s.list.splice(i,0,a);paint(key);say(tr('تعذّر الحذف','Could not remove'))}}
    else if(s.fresh.has(a.id))removeFile(a);}
});
document.addEventListener('dragover',e=>{const z=e.target.closest&&e.target.closest('[data-att-drop]');if(!z)return;e.preventDefault();z.classList.add('over')});
document.addEventListener('dragleave',e=>{const z=e.target.closest&&e.target.closest('[data-att-drop]');if(z&&!z.contains(e.relatedTarget))z.classList.remove('over')});
document.addEventListener('drop',e=>{const z=e.target.closest&&e.target.closest('[data-att-drop]');if(!z)return;e.preventDefault();z.classList.remove('over');const f=[...(e.dataTransfer.files||[])];if(f.length)addBlobs(z.dataset.attDrop,f.map(x=>({blob:x,name:x.name})))});

/* ---------------- viewer ---------------- */
let VD=null;const V={list:[],i:0,url:''};
function viewer(list,i){
  if(!VD){VD=document.createElement('dialog');VD.className='attv';document.body.appendChild(VD);
    VD.addEventListener('keydown',e=>{e.stopPropagation();if(e.key==='ArrowLeft'||e.key==='ArrowRight'){const rtl=document.documentElement.dir==='rtl';vgo((e.key==='ArrowLeft')===rtl?1:-1)}});
    VD.addEventListener('close',()=>{if(V.url)URL.revokeObjectURL(V.url);V.url='';VD.innerHTML=''});
    VD.addEventListener('click',async e=>{const b=e.target.closest('[data-v]');if(e.target===VD){VD.close();return}if(!b)return;const a=V.list[V.i];
      if(b.dataset.v==='x')VD.close();else if(b.dataset.v==='prev')vgo(-1);else if(b.dataset.v==='next')vgo(1);
      else if(b.dataset.v==='dl'){try{saveBlob(await getBlob(a),a.name)}catch(err){say(err.message)}}
      else if(b.dataset.v==='tab'&&V.url)window.open(V.url,'_blank','noopener');
      else if(b.dataset.v==='zoom')b.closest('.attv-body').classList.toggle('z')})}
  V.list=list;V.i=i;if(!VD.open)VD.showModal();vshow();
}
function vgo(d){if(V.list.length<2)return;V.i=(V.i+d+V.list.length)%V.list.length;vshow()}
async function vshow(){const a=V.list[V.i],k=kind(a),n=V.list.length;if(V.url){URL.revokeObjectURL(V.url);V.url=''}
  VD.innerHTML=`<div class="attv-head"><span class="att-ic k-${k}">${a.scanned?IC.scan:''}<b>${KLAB[k]}</b></span><div class="attv-t"><b dir="auto">${H(a.name)}</b><small>${H([fsize(a.size||0),a.pages?pagesTxt(a.pages):'',a.by||'',a.at?new Date(a.at).toLocaleString('en-GB'):''].filter(Boolean).join(' · '))}${n>1?` · ${V.i+1}/${n}`:''}</small></div>
   <button class="attv-b" data-v="dl" title="${H(tr('تنزيل','Download'))}">${IC.dl}<span>${H(tr('تنزيل','Download'))}</span></button><button class="attv-b" data-v="tab" title="${H(tr('فتح في نافذة جديدة','Open in new tab'))}">${IC.ext}</button><button class="attv-b" data-v="x" aria-label="${H(tr('إغلاق','Close'))}">${IC.x}</button></div>
   <div class="attv-body"><div class="att-spin"></div>${n>1?`<button class="attv-nav p flip" data-v="prev" aria-label="${H(tr('السابق','Previous'))}">${IC.back}</button><button class="attv-nav n flip" data-v="next" aria-label="${H(tr('التالي','Next'))}">${IC.next}</button>`:''}</div>`;
  const body=VD.querySelector('.attv-body');const want=V.i;
  try{const blob=await getBlob(a);if(want!==V.i||!VD.open)return;V.url=URL.createObjectURL(blob);
    const coarse=matchMedia('(pointer:coarse)').matches;let el;
    if(k==='img'){el=document.createElement('img');el.src=V.url;el.alt=a.name;el.dataset.v='zoom'}
    else if(k==='pdf'){el=await pdfPages(blob,body,want);if(!el&&!coarse){el=document.createElement('iframe');el.src=V.url+'#view=FitH';el.title=a.name}}
    if(want!==V.i||!VD.open)return;
    if(!el){el=document.createElement('div');el.className='attv-card';el.innerHTML=`<span class="att-ic big k-${k}"><b>${KLAB[k]}</b></span><b dir="auto">${H(a.name)}</b><span>${H(fsize(a.size||0))}</span><div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:center"><button class="att-btn att-scan" data-v="tab">${IC.ext}${H(tr('فتح الملف','Open file'))}</button><button class="att-btn" data-v="dl">${IC.dl}${H(tr('تنزيل','Download'))}</button></div>`}
    const sp=body.querySelector('.att-spin');if(sp)sp.remove();body.prepend(el)}
  catch(e){body.innerHTML=`<div class="attv-card"><b>${H(tr('تعذّر فتح الملف','Could not open the file'))}</b><span>${H(e.message||'')}</span></div>`}
}

/* PDFs are drawn page by page with pdf.js so they display the same on phones and desktops. */
let pdfjsP=null;
function pdfjs(){if(window.FileImport&&FileImport.pdfjs)return FileImport.pdfjs();if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);if(!pdfjsP)pdfjsP=new Promise((res,rej)=>{const s=document.createElement('script');s.src='vendor/pdfjs/pdf.min.js';s.integrity='sha384-/1qUCSGwTur9vjf/z9lmu/eCUYbpOTgSjmpbMQZ1/CtX2v/WcAIKqRv+U1DUCG6e';s.crossOrigin='anonymous';
  s.onload=()=>{if(!window.pdfjsLib)return rej(new Error('pdfjs'));pdfjsLib.GlobalWorkerOptions.workerSrc='vendor/pdfjs/pdf.worker.min.js';res(pdfjsLib)};s.onerror=()=>{pdfjsP=null;rej(new Error('pdfjs'))};document.head.appendChild(s)});return pdfjsP}
async function pdfPages(blob,body,want){
  try{const lib=await pdfjs();const doc=await lib.getDocument({data:new Uint8Array(await blob.arrayBuffer())}).promise;if(want!==V.i)return null;
    const wrap=document.createElement('div');wrap.className='attv-pages';const W=Math.min(body.clientWidth-32,980),dpr=Math.min(2,window.devicePixelRatio||1);
    (async()=>{for(let i=1;i<=doc.numPages;i++){if(!wrap.isConnected&&i>1)return;const pg=await doc.getPage(i);const v0=pg.getViewport({scale:1}),sc=W/v0.width,vp=pg.getViewport({scale:sc*dpr});
      const c=mk(vp.width,vp.height);c.style.width=Math.round(v0.width*sc)+'px';c.style.height=Math.round(v0.height*sc)+'px';const f=document.createElement('figure');f.appendChild(c);
      const cap=document.createElement('figcaption');cap.textContent=i+' / '+doc.numPages;f.appendChild(cap);wrap.appendChild(f);await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise}})();
    return wrap}catch(e){return null}
}

/* ---------------- image processing ---------------- */
const FULL=[[0,0],[1,0],[1,1],[0,1]];
const isFull=q=>q.every((p,i)=>Math.abs(p[0]-FULL[i][0])<1e-3&&Math.abs(p[1]-FULL[i][1])<1e-3);
function grayOf(d,n){const g=new Uint8ClampedArray(n);for(let i=0,j=0;i<n;i++,j+=4)g[i]=(d[j]*77+d[j+1]*151+d[j+2]*28)>>8;return g}
function blurU8(g,w,h,r){const t=new Float32Array(w*h),o=new Uint8ClampedArray(w*h),k=2*r+1;
  for(let y=0;y<h;y++){const row=y*w;let acc=0;for(let x=-r;x<=r;x++)acc+=g[row+clamp(x,0,w-1)];for(let x=0;x<w;x++){t[row+x]=acc/k;acc+=g[row+Math.min(w-1,x+r+1)]-g[row+Math.max(0,x-r)]}}
  for(let x=0;x<w;x++){let acc=0;for(let y=-r;y<=r;y++)acc+=t[clamp(y,0,h-1)*w+x];for(let y=0;y<h;y++){o[y*w+x]=acc/k;acc+=t[Math.min(h-1,y+r+1)*w+x]-t[Math.max(0,y-r)*w+x]}}
  return o}
const area=q=>Math.abs(q.reduce((s,p,i)=>{const n=q[(i+1)%4];return s+p[0]*n[1]-n[0]*p[1]},0))/2;
function convex(q){let sg=0;for(let i=0;i<4;i++){const a=q[i],b=q[(i+1)%4],c=q[(i+2)%4];const z=(b[0]-a[0])*(c[1]-b[1])-(b[1]-a[1])*(c[0]-b[0]);if(Math.abs(z)<1e-6)return false;const s=Math.sign(z);if(sg&&s!==sg)return false;sg=s}return true}
/* Finds the sheet of paper: Otsu threshold on a blurred thumbnail, largest bright region, then its four extreme corners. */
function detectQuad(g,w,h){
  const b=blurU8(g,w,h,2),N=w*h;const hist=new Uint32Array(256);for(let i=0;i<N;i++)hist[b[i]]++;
  let sum=0;for(let i=0;i<256;i++)sum+=i*hist[i];let sB=0,wB=0,best=0,t=127,m0=0,m1=0;
  for(let i=0;i<256;i++){wB+=hist[i];if(!wB)continue;const wF=N-wB;if(!wF)break;sB+=i*hist[i];const mB=sB/wB,mF=(sum-sB)/wF,v=wB*wF*(mB-mF)*(mB-mF);if(v>best){best=v;t=i;m0=mB;m1=mF}}
  if(m1-m0<26)return null;
  const lab=new Int32Array(N),st=new Int32Array(N);let cur=0,bestA=0,bestL=0;
  for(let i=0;i<N;i++){if(b[i]<=t||lab[i])continue;cur++;let sp=0,a=0;st[sp++]=i;lab[i]=cur;
    while(sp){const p=st[--sp];a++;const x=p%w;
      if(x>0&&!lab[p-1]&&b[p-1]>t){lab[p-1]=cur;st[sp++]=p-1}
      if(x<w-1&&!lab[p+1]&&b[p+1]>t){lab[p+1]=cur;st[sp++]=p+1}
      if(p>=w&&!lab[p-w]&&b[p-w]>t){lab[p-w]=cur;st[sp++]=p-w}
      if(p<N-w&&!lab[p+w]&&b[p+w]>t){lab[p+w]=cur;st[sp++]=p+w}}
    if(a>bestA){bestA=a;bestL=cur}}
  if(bestA<N*.1||bestA>N*.96)return null;
  let tl,tR,br,bl,s1=1e9,s2=-1e9,d1=-1e9,d2=1e9;
  for(let i=0;i<N;i++){if(lab[i]!==bestL)continue;const x=i%w,y=(i-x)/w,s=x+y,d=x-y;if(s<s1){s1=s;tl=[x,y]}if(s>s2){s2=s;br=[x,y]}if(d>d1){d1=d;tR=[x,y]}if(d<d2){d2=d;bl=[x,y]}}
  const q=[tl,tR,br,bl];const A=area(q);if(A<N*.1)return null;const fill=bestA/A;if(fill<.82||fill>1.18||!convex(q))return null;
  const c=[(tl[0]+tR[0]+br[0]+bl[0])/4,(tl[1]+tR[1]+br[1]+bl[1])/4];
  return q.map(([x,y])=>[clamp((x+.5-(x-c[0])*.008)/w,0,1),clamp((y+.5-(y-c[1])*.008)/h,0,1)]);
}
function detectOn(src){const s=Math.min(1,256/Math.max(src.width,src.height));const c=mk(Math.round(src.width*s),Math.round(src.height*s));const x=cx(c);x.drawImage(src,0,0,c.width,c.height);return detectQuad(grayOf(x.getImageData(0,0,c.width,c.height).data,c.width*c.height),c.width,c.height)}
function sq2quad(P){const [x0,y0]=P[0],[x1,y1]=P[1],[x2,y2]=P[2],[x3,y3]=P[3];const dx1=x1-x2,dx2=x3-x2,sx=x0-x1+x2-x3,dy1=y1-y2,dy2=y3-y2,sy=y0-y1+y2-y3;
  let g=0,h=0;if(Math.abs(sx)>1e-9||Math.abs(sy)>1e-9){const den=dx1*dy2-dx2*dy1;g=(sx*dy2-dx2*sy)/den;h=(dx1*sy-sx*dy1)/den}
  return {a:x1-x0+g*x1,b:x3-x0+h*x3,c:x0,d:y1-y0+g*y1,e:y3-y0+h*y3,f:y0,g,h}}
/* Perspective-corrects the quad (normalized corners tl,tr,br,bl) into a flat page, longest side capped at maxSide. */
function warp(src,q,maxSide){
  let W0=src.width,H0=src.height;const P0=q.map(([x,y])=>[x*W0,y*H0]);const dist=(a,b)=>Math.hypot(a[0]-b[0],a[1]-b[1]);
  let w=Math.max(dist(P0[0],P0[1]),dist(P0[3],P0[2])),h=Math.max(dist(P0[0],P0[3]),dist(P0[1],P0[2]));const s=Math.min(1,maxSide/Math.max(w,h));w=Math.max(1,Math.round(w*s));h=Math.max(1,Math.round(h*s));
  if(isFull(q))return scaleTo(src,maxSide);
  if(Math.max(W0,H0)>maxSide*2.2){src=scaleTo(src,Math.round(maxSide*2));W0=src.width;H0=src.height}
  const P=q.map(([x,y])=>[x*W0,y*H0]);const sd=cx(src instanceof HTMLCanvasElement?src:scaleTo(src,Math.max(W0,H0))).getImageData(0,0,W0,H0).data;
  const out=mk(w,h),ox=out.getContext('2d'),id=ox.createImageData(w,h),od=id.data,M=sq2quad(P);
  for(let y=0;y<h;y++){const v=(y+.5)/h;for(let x=0;x<w;x++){const u=(x+.5)/w,den=M.g*u+M.h*v+1;let sx=(M.a*u+M.b*v+M.c)/den-.5,sy=(M.d*u+M.e*v+M.f)/den-.5;
    sx=clamp(sx,0,W0-1.001);sy=clamp(sy,0,H0-1.001);const ix=sx|0,iy=sy|0,fx=sx-ix,fy=sy-iy,i0=(iy*W0+ix)*4,i1=i0+W0*4,o=(y*w+x)*4;
    for(let c=0;c<3;c++){const tp=sd[i0+c]+(sd[i0+4+c]-sd[i0+c])*fx,bt=sd[i1+c]+(sd[i1+4+c]-sd[i1+c])*fx;od[o+c]=tp+(bt-tp)*fy}od[o+3]=255}}
  ox.putImageData(id,0,0);return out;
}
function rotateC(c,r){r=((r%4)+4)%4;if(!r)return c;const o=r%2?mk(c.height,c.width):mk(c.width,c.height);const x=o.getContext('2d');x.translate(o.width/2,o.height/2);x.rotate(r*Math.PI/2);x.drawImage(c,-c.width/2,-c.height/2);return o}
/* Paper whitening: estimate the lighting from the brightest pixels in coarse blocks, divide it out, then stretch. */
function applyFilter(c,mode){
  if(!mode||mode==='orig')return c;
  const W=c.width,Hh=c.height,N=W*Hh,x=cx(c),id=x.getImageData(0,0,W,Hh),d=id.data,L=grayOf(d,N);
  const bs=Math.max(8,Math.round(Math.max(W,Hh)/44)),gw=Math.ceil(W/bs),gh=Math.ceil(Hh/bs);let G=new Float32Array(gw*gh);
  for(let gy=0;gy<gh;gy++)for(let gx=0;gx<gw;gx++){const h0=new Uint16Array(32);let cnt=0;const y1=Math.min(Hh,(gy+1)*bs),x1=Math.min(W,(gx+1)*bs);
    for(let y=gy*bs;y<y1;y+=2){const o=y*W;for(let xx=gx*bs;xx<x1;xx+=2){h0[L[o+xx]>>3]++;cnt++}}
    let k=cnt*.06,bin=31;while(bin>0&&k>0){k-=h0[bin];if(k>0)bin--}G[gy*gw+gx]=bin*8+4}
  const pass=(A,f)=>{const B=new Float32Array(A.length);for(let y=0;y<gh;y++)for(let xx=0;xx<gw;xx++){let m=f==='max'?0:0,n=0;for(let j=-1;j<=1;j++)for(let i=-1;i<=1;i++){const yy=y+j,xi=xx+i;if(yy<0||yy>=gh||xi<0||xi>=gw)continue;const v=A[yy*gw+xi];if(f==='max'){if(v>m)m=v}else{m+=v;n++}}B[y*gw+xx]=f==='max'?m:m/n}return B};
  G=pass(G,'max');G=pass(G,'avg');G=pass(G,'avg');
  const lo=.1,hi=.9,span=hi-lo;
  for(let y=0;y<Hh;y++){let fy=(y+.5)/bs-.5;fy=clamp(fy,0,gh-1);const y0=fy|0,y1=Math.min(gh-1,y0+1),ty=fy-y0;
    for(let xx=0;xx<W;xx++){let fx=(xx+.5)/bs-.5;fx=clamp(fx,0,gw-1);const x0=fx|0,x1=Math.min(gw-1,x0+1),tx=fx-x0;
      const a=G[y0*gw+x0]+(G[y0*gw+x1]-G[y0*gw+x0])*tx,b=G[y1*gw+x0]+(G[y1*gw+x1]-G[y1*gw+x0])*tx;const bg=Math.max(48,a+(b-a)*ty);const j=(y*W+xx)*4;
      if(mode==='magic'){let r=clamp((d[j]/bg-lo)/span,0,1),g=clamp((d[j+1]/bg-lo)/span,0,1),bb=clamp((d[j+2]/bg-lo)/span,0,1);r=Math.pow(r,1.25);g=Math.pow(g,1.25);bb=Math.pow(bb,1.25);
        const m=(r+g+bb)/3;d[j]=clamp(m+(r-m)*1.3,0,1)*255;d[j+1]=clamp(m+(g-m)*1.3,0,1)*255;d[j+2]=clamp(m+(bb-m)*1.3,0,1)*255}
      else{const n=L[y*W+xx]/bg;let v;if(mode==='gray')v=Math.pow(clamp((n-lo)/span,0,1),1.3)*255;else v=clamp((n-.64)/.16,0,1)*255;d[j]=d[j+1]=d[j+2]=v}}}
  x.putImageData(id,0,0);return c;
}

/* ---------------- minimal PDF writer (JPEG pages, no library) ---------------- */
function makePDF(pages,title){
  const enc=new TextEncoder(),parts=[],offs=[];let len=0;const push=x=>{const b=typeof x==='string'?enc.encode(x):x;parts.push(b);len+=b.length};
  const obj=(n,body)=>{offs[n]=len;push(n+' 0 obj\n'+body+'\nendobj\n')};const N=pages.length,info=3+N*3;
  push('%PDF-1.4\n');push(new Uint8Array([37,226,227,207,211,10]));
  obj(1,'<< /Type /Catalog /Pages 2 0 R >>');
  obj(2,`<< /Type /Pages /Kids [${pages.map((_,i)=>(3+i*3)+' 0 R').join(' ')}] /Count ${N} >>`);
  pages.forEach((p,i)=>{const pn=3+i*3,cn=pn+1,im=pn+2;const A4=595.28;let W,Hh;if(p.w<=p.h){W=A4;Hh=A4*p.h/p.w}else{Hh=A4;W=A4*p.w/p.h}W=+W.toFixed(2);Hh=+Hh.toFixed(2);
    obj(pn,`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W} ${Hh}] /Resources << /XObject << /Im0 ${im} 0 R >> >> /Contents ${cn} 0 R >>`);
    const cs=`q ${W} 0 0 ${Hh} 0 0 cm /Im0 Do Q`;obj(cn,`<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`);
    offs[im]=len;push(`${im} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.u8.length} >>\nstream\n`);push(p.u8);push('\nendstream\nendobj\n')});
  const u16=s=>'<FEFF'+[...s].map(ch=>{let o='';for(let i=0;i<ch.length;i++)o+=ch.charCodeAt(i).toString(16).padStart(4,'0');return o}).join('').toUpperCase()+'>';
  const dt=new Date(),pad=n=>String(n).padStart(2,'0');
  obj(info,`<< /Title ${u16(title||'Scan')} /Producer ${u16('Al-Brega Sales System · Document Scanner')} /CreationDate (D:${dt.getFullYear()}${pad(dt.getMonth()+1)}${pad(dt.getDate())}${pad(dt.getHours())}${pad(dt.getMinutes())}${pad(dt.getSeconds())}) >>`);
  const xref=len,total=info+1;push(`xref\n0 ${total}\n0000000000 65535 f \n`);for(let i=1;i<total;i++)push(String(offs[i]).padStart(10,'0')+' 00000 n \n');
  push(`trailer\n<< /Size ${total} /Root 1 0 R /Info ${info} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  return new Blob(parts,{type:'application/pdf'});
}

/* ---------------- scanner ---------------- */
const FILTERS=()=>[['magic',tr('تحسين ذكي','Magic'),IC.wand],['orig',tr('الأصل','Original'),IC.img],['gray',tr('رمادي','Grayscale'),''],['bw',tr('أبيض وأسود','B & W'),'']];
const PRESETS={hi:{max:2200,q:.86,l:['عالية','High']},md:{max:1650,q:.74,l:['متوازنة','Balanced']},lo:{max:1240,q:.6,l:['مضغوطة','Compact']}};
const store=(k,v)=>{try{if(v===undefined)return localStorage.getItem('scn_'+k);localStorage.setItem('scn_'+k,v)}catch(e){return null}};
let scanBusy=false;
function scanOpen(opts={}){
 if(scanBusy)return Promise.resolve(null);scanBusy=true;
 return new Promise(resolve=>{
  const d=document.createElement('dialog');d.className='scn';d.setAttribute('aria-label',tr('الماسح الضوئي','Document scanner'));document.body.appendChild(d);
  const now=new Date(),pad=n=>String(n).padStart(2,'0');
  const defName=`${opts.name?opts.name+' - ':''}${tr('مستند ممسوح','Scan')} ${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())} ${pad(now.getHours())}${pad(now.getMinutes())}`;
  let pages=[],stream=null,track=null,raf=0,devs=[],devIdx=-1,torch=false,auto=store('auto')!=='0',gFilter=store('filter')||'magic',preset=store('preset')||'md';
  let detQ=null,dispQ=null,miss=99,hist=[],lastDet=0,stableAt=0,cool=0,armed=true,lastGray=null,capGray=null,shooting=false,camOn=false,vw=0,vh=0,edit=null,closed=false;
  const sm=mk(200,150),smx=cx(sm);
  d.innerHTML=`<div class="scn-root" data-screen="cam">
  <section class="scn-cam">
   <header class="scn-top"><button class="scn-ib" data-sc="close" aria-label="${H(tr('إغلاق','Close'))}">${IC.x}</button>
    <div class="scn-title"><b>${H(tr('الماسح الضوئي','Document scanner'))}</b><small class="scn-sub"></small></div>
    <button class="scn-pill" data-sc="auto" aria-pressed="${auto}" title="${H(tr('التقاط تلقائي عند ثبات المستند','Auto-capture when the document is steady'))}">${IC.auto}<span>${H(tr('تلقائي','Auto'))}</span></button>
    <button class="scn-ib" data-sc="torch" hidden aria-pressed="false" aria-label="${H(tr('الفلاش','Flash'))}">${IC.bolt}</button>
    <button class="scn-ib" data-sc="switch" hidden aria-label="${H(tr('تبديل الكاميرا','Switch camera'))}">${IC.sw}</button></header>
   <div class="scn-stage"><video playsinline muted autoplay></video>
    <svg class="scn-ov" preserveAspectRatio="xMidYMid meet"><path class="scn-shade"/><polygon class="scn-quad"/><g class="scn-cns"><circle class="scn-cn"/><circle class="scn-cn"/><circle class="scn-cn"/><circle class="scn-cn"/></g></svg>
    <div class="scn-hint"><span class="dot"></span><span class="scn-hint-t">${H(tr('جارٍ تشغيل الكاميرا…','Starting camera…'))}</span></div>
    <div class="scn-flash"></div>
    <div class="scn-fail" hidden><div class="scn-fcard"><span class="scn-fic">${IC.cam}</span><b class="scn-fh"></b><p class="scn-fp"></p>
      <div class="scn-fbtns"><button class="scn-go" data-sc="native">${IC.cam}${H(tr('التقاط بكاميرا الجهاز','Use device camera'))}</button><button class="scn-ghost" data-sc="gallery">${IC.img}${H(tr('اختيار صور من الجهاز','Choose images'))}</button><button class="scn-ghost" data-sc="retry" hidden>${IC.rot}${H(tr('إعادة المحاولة','Try again'))}</button></div></div></div>
   </div>
   <footer class="scn-bot">
    <button class="scn-side" data-sc="gallery"><span class="scn-sq">${IC.img}</span><small>${H(tr('من الصور','Import'))}</small></button>
    <button class="scn-shutter" data-sc="shoot" aria-label="${H(tr('التقاط','Capture'))}"><svg viewBox="0 0 80 80"><circle cx="40" cy="40" r="37" class="r0"/><circle cx="40" cy="40" r="37" class="r1"/></svg></button>
    <button class="scn-side scn-stack" data-sc="review" disabled><span class="scn-thumbs"></span><em class="scn-count">0</em><small>${H(tr('مراجعة','Review'))}</small></button>
   </footer>
  </section>
  <section class="scn-rev">
   <header class="scn-top solid"><button class="scn-ib flip" data-sc="tocam" aria-label="${H(tr('رجوع للكاميرا','Back to camera'))}">${IC.back}</button>
    <div class="scn-title"><b>${H(tr('مراجعة المستند','Review document'))}</b><small class="scn-rsub"></small></div>
    <button class="scn-ib" data-sc="close" aria-label="${H(tr('إغلاق','Close'))}">${IC.x}</button></header>
   <div class="scn-filters" role="radiogroup" aria-label="${H(tr('تحسين الصورة','Enhancement'))}"></div>
   <div class="scn-grid"></div>
   <footer class="scn-done">
    <label class="scn-name"><span>${H(tr('اسم الملف','File name'))}</span><span class="scn-in"><input type="text" dir="auto" value="${H(defName)}"><em dir="ltr">.pdf</em></span></label>
    <div class="scn-q" role="radiogroup" aria-label="${H(tr('الجودة','Quality'))}">${Object.entries(PRESETS).map(([k,p])=>`<button data-sc="preset" data-p="${k}" aria-pressed="${k===preset}">${H(tr(p.l[0],p.l[1]))}</button>`).join('')}</div>
    <span class="scn-est"></span>
    <button class="scn-go big" data-sc="finish">${IC.check}<span>${H(tr('إنشاء PDF وإرفاقه','Create PDF & attach'))}</span></button>
   </footer>
  </section>
  <section class="scn-edit">
   <header class="scn-top solid"><button class="scn-ib flip" data-sc="ecancel" aria-label="${H(tr('رجوع','Back'))}">${IC.back}</button>
    <div class="scn-title"><b>${H(tr('قص وتعديل الصفحة','Crop & adjust page'))}</b><small>${H(tr('اسحب الزوايا الأربع إلى حواف الورقة','Drag the four corners to the paper edges'))}</small></div>
    <button class="scn-go" data-sc="eapply">${IC.check}<span>${H(tr('تطبيق','Apply'))}</span></button></header>
   <div class="scn-ed-stage"><div class="scn-ed-box"><canvas></canvas><svg preserveAspectRatio="none"><path class="scn-shade on"/><polygon class="scn-equad"/><g class="scn-hs"></g></svg></div></div>
   <div class="scn-ed-tools">
    <div class="scn-tl"><button class="scn-tool" data-sc="edetect">${IC.wand}<span>${H(tr('اكتشاف تلقائي','Auto detect'))}</span></button><button class="scn-tool" data-sc="efull">${IC.full}<span>${H(tr('الصفحة كاملة','Full image'))}</span></button><button class="scn-tool" data-sc="erot">${IC.rot}<span>${H(tr('تدوير','Rotate'))}</span></button></div>
    <div class="scn-tl scn-efl"></div></div>
   <canvas class="scn-loupe" width="240" height="240" hidden></canvas>
  </section>
  <div class="scn-sheet" hidden><div class="scn-sheet-c"><b>${H(tr('تجاهل المسح؟','Discard this scan?'))}</b><p class="scn-sheet-p"></p><div class="scn-fbtns"><button class="scn-ghost" data-sc="keep">${H(tr('متابعة المسح','Keep scanning'))}</button><button class="scn-danger" data-sc="discard">${IC.trash}${H(tr('تجاهل','Discard'))}</button></div></div></div>
  <div class="scn-busy" hidden><div class="scn-busy-c"><div class="scn-ring"><svg viewBox="0 0 80 80"><circle cx="40" cy="40" r="34" class="r0"/><circle cx="40" cy="40" r="34" class="r1"/></svg><span class="scn-ok">${IC.check}</span></div><b class="scn-busy-t"></b><small class="scn-busy-s"></small></div></div>
  <div class="scn-toast" role="status"></div>
  </div>`;
  const $=s=>d.querySelector(s),root=$('.scn-root'),video=$('video'),ov=$('.scn-ov'),grid=$('.scn-grid');
  const gal=document.createElement('input');gal.type='file';gal.accept='image/*';gal.multiple=true;
  const nat=document.createElement('input');nat.type='file';nat.accept='image/*';nat.setAttribute('capture','environment');
  let toastT=0;const stoast=m=>{const t=$('.scn-toast');t.textContent=m;t.classList.add('on');clearTimeout(toastT);toastT=setTimeout(()=>t.classList.remove('on'),2600)};
  const hint=(txt,state)=>{$('.scn-hint-t').textContent=txt;root.classList.toggle('det',state==='det'||state==='hold');root.classList.toggle('hold',state==='hold')};
  const sub=()=>{const n=pages.length;$('.scn-sub').textContent=n?tr(`التقطت ${pagesTxt(n)} · الصفحة التالية ${n+1}`,`${pagesTxt(n)} captured · next is page ${n+1}`):tr('الصفحة 1','Page 1');$('.scn-rsub').textContent=pagesTxt(n)+' · '+tr('اسحب البطاقة لإعادة الترتيب','drag a card to reorder')};

  /* ---- camera ---- */
  async function camStart(){
    if(camOn||closed)return;$('.scn-fail').hidden=true;hint(tr('جارٍ تشغيل الكاميرا…','Starting camera…'));
    if(!window.isSecureContext||!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia)return camFail('unsupported');
    const base={width:{ideal:2560},height:{ideal:1440}};
    try{stream=await navigator.mediaDevices.getUserMedia({audio:false,video:devIdx>=0&&devs[devIdx]?Object.assign({deviceId:{exact:devs[devIdx].deviceId}},base):Object.assign({facingMode:{ideal:'environment'}},base)})}
    catch(e){return camFail(e.name)}
    if(closed||root.dataset.screen!=='cam'){stream.getTracks().forEach(t=>t.stop());stream=null;return}
    camOn=true;video.srcObject=stream;try{await video.play()}catch(e){}
    track=stream.getVideoTracks()[0];try{await track.applyConstraints({advanced:[{focusMode:'continuous'}]})}catch(e){}
    const caps=track.getCapabilities?track.getCapabilities():{};$('[data-sc="torch"]').hidden=!caps.torch;torch=false;$('[data-sc="torch"]').setAttribute('aria-pressed','false');
    try{devs=(await navigator.mediaDevices.enumerateDevices()).filter(x=>x.kind==='videoinput');$('[data-sc="switch"]').hidden=devs.length<2;if(devIdx<0){const s=track.getSettings?track.getSettings():{};devIdx=Math.max(0,devs.findIndex(x=>x.deviceId===s.deviceId))}}catch(e){}
    const ready=()=>{vw=video.videoWidth;vh=video.videoHeight;if(!vw)return setTimeout(ready,60);ov.setAttribute('viewBox',`0 0 ${vw} ${vh}`);const s=200/Math.max(vw,vh);sm.width=Math.round(vw*s);sm.height=Math.round(vh*s);dispQ=null;hint(tr('وجّه الكاميرا نحو المستند','Point the camera at the document'))};ready();
    cancelAnimationFrame(raf);raf=requestAnimationFrame(loop);
  }
  function camStop(){camOn=false;cancelAnimationFrame(raf);if(stream)stream.getTracks().forEach(t=>t.stop());stream=null;track=null;video.srcObject=null}
  function camFail(name){camOn=false;const f=$('.scn-fail');f.hidden=false;hint('');
    const M={NotAllowedError:[tr('لم يُسمح باستخدام الكاميرا','Camera access was blocked'),tr('اسمح للموقع باستخدام الكاميرا من شريط العنوان أو إعدادات المتصفح ثم أعد المحاولة، أو التقط الصورة بتطبيق الكاميرا مباشرة.','Allow camera access from the address bar or browser settings and try again, or take the photo with the camera app.')],
      NotFoundError:[tr('لا توجد كاميرا متصلة','No camera found'),tr('لم نجد كاميرا في هذا الجهاز. يمكنك اختيار صور للمستند من الجهاز وسنحسّنها ونحولها إلى PDF.','This device has no camera. Choose photos of the document and we will enhance them and turn them into a PDF.')],
      NotReadableError:[tr('الكاميرا مشغولة','Camera is busy'),tr('قد يكون تطبيق آخر يستخدم الكاميرا. أغلقه ثم أعد المحاولة.','Another app may be using the camera. Close it and try again.')],
      unsupported:[tr('المتصفح لا يدعم الكاميرا المباشرة','Live camera not supported'),tr('استخدم كاميرا الجهاز أو اختر صورًا، وسنكمل التحسين والقص وإنشاء PDF هنا.','Use the device camera or pick images; enhancement, cropping and PDF creation still happen here.')]};
    const m=M[name]||[tr('تعذّر تشغيل الكاميرا','Could not start the camera'),tr('أعد المحاولة، أو استخدم كاميرا الجهاز أو الصور.','Try again, or use the device camera or images.')];
    $('.scn-fh').textContent=m[0];$('.scn-fp').textContent=m[1];$('[data-sc="retry"]').hidden=name==='NotFoundError'||name==='unsupported'}
  function guideQ(){const a=vw/vh;let w=.78,h=w*a*1.414;if(h>.84){h=.84;w=h/(a*1.414)}return [[.5-w/2,.5-h/2],[.5+w/2,.5-h/2],[.5+w/2,.5+h/2],[.5-w/2,.5+h/2]]}
  function loop(t){
    if(!camOn)return;raf=requestAnimationFrame(loop);if(!vw||video.readyState<2)return;
    if(t-lastDet>110){lastDet=t;detectFrame(t)}
    const target=detQ||guideQ();if(!dispQ)dispQ=target.map(p=>p.slice());const k=detQ?.35:.18;
    for(let i=0;i<4;i++){dispQ[i][0]+=(target[i][0]-dispQ[i][0])*k;dispQ[i][1]+=(target[i][1]-dispQ[i][1])*k}
    const P=dispQ.map(([x,y])=>[x*vw,y*vh]),pts=P.map(p=>p[0].toFixed(1)+','+p[1].toFixed(1)).join(' ');
    $('.scn-quad').setAttribute('points',pts);$('.scn-shade').setAttribute('d',`M0 0H${vw}V${vh}H0Z M${P.map(p=>p[0].toFixed(1)+' '+p[1].toFixed(1)).join(' L')}Z`);
    const r=Math.max(vw,vh)/90;d.querySelectorAll('.scn-cn').forEach((c,i)=>{c.setAttribute('cx',P[i][0]);c.setAttribute('cy',P[i][1]);c.setAttribute('r',r)});
  }
  function detectFrame(t){
    smx.drawImage(video,0,0,sm.width,sm.height);const g=grayOf(smx.getImageData(0,0,sm.width,sm.height).data,sm.width*sm.height);lastGray=g;
    const q=detectQuad(g,sm.width,sm.height);
    if(q){miss=0;detQ=q}else if(++miss>3){detQ=null;armed=true}
    hist.push(q);if(hist.length>8)hist.shift();
    const steady=detQ&&hist.length>=6&&hist.slice(-6).every(h=>h&&h.every((p,i)=>Math.hypot(p[0]-detQ[i][0],p[1]-detQ[i][1])<.018));
    if(!armed&&capGray&&diff(g,capGray)>18)armed=true;
    let ring=0;
    if(!detQ){stableAt=0;hint(pages.length?tr('ضع الصفحة التالية أمام الكاميرا','Place the next page in view'):tr('وجّه الكاميرا نحو المستند','Point the camera at the document'))}
    else if(auto&&steady&&armed&&t>cool){if(!stableAt)stableAt=t;ring=Math.min(1,(t-stableAt)/900);hint(tr('ثبّت الجهاز… جارٍ الالتقاط','Hold steady… capturing'),'hold');if(ring>=1){stableAt=0;shoot()}}
    else{stableAt=0;hint(auto&&!armed?tr('تم الالتقاط · اقلب الصفحة أو أبعد المستند','Captured · turn the page or move the document'):auto?tr('تم اكتشاف المستند · ثبّت الجهاز','Document found · hold steady'):tr('تم اكتشاف المستند · اضغط زر الالتقاط','Document found · tap the shutter'),'det')}
    $('.scn-shutter .r1').style.strokeDashoffset=String(232.5*(1-ring));
  }
  const diff=(a,b)=>{if(!a||!b||a.length!==b.length)return 255;let s=0;for(let i=0;i<a.length;i+=3)s+=Math.abs(a[i]-b[i]);return s/(a.length/3)};
  async function shoot(){
    if(shooting||!camOn||!vw)return;shooting=true;
    const s=Math.min(1,2600/Math.max(vw,vh)),c=mk(Math.round(vw*s),Math.round(vh*s));c.getContext('2d').drawImage(video,0,0,c.width,c.height);
    const q=detQ?detQ.map(p=>p.slice()):null;capGray=lastGray;armed=false;cool=performance.now()+1300;
    const fl=$('.scn-flash');fl.classList.remove('go');void fl.offsetWidth;fl.classList.add('go');vib(35);
    const pg=await addPage(c,q);fly(pg,q);shooting=false;
  }
  function fly(pg,q){
    const m=ov.getScreenCTM();const stack=$('.scn-thumbs').getBoundingClientRect();if(!m)return;
    const Q=(q||FULL).map(([x,y])=>{const p=ov.createSVGPoint();p.x=x*vw;p.y=y*vh;const r=p.matrixTransform(m);return [r.x,r.y]});
    const x0=Math.min(...Q.map(p=>p[0])),y0=Math.min(...Q.map(p=>p[1])),x1=Math.max(...Q.map(p=>p[0])),y1=Math.max(...Q.map(p=>p[1]));
    const im=document.createElement('img');im.className='scn-fly';im.src=pg.raw;Object.assign(im.style,{left:x0+'px',top:y0+'px',width:(x1-x0)+'px',height:(y1-y0)+'px'});d.appendChild(im);
    requestAnimationFrame(()=>requestAnimationFrame(()=>{Object.assign(im.style,{left:stack.left+'px',top:stack.top+'px',width:stack.width+'px',height:stack.height+'px',opacity:'.2',transform:'rotate(-8deg)'})}));
    setTimeout(()=>{im.remove();const c=$('.scn-count');c.classList.remove('pop');void c.offsetWidth;c.classList.add('pop')},650);
  }

  /* ---- pages ---- */
  const keyOf=p=>JSON.stringify([p.quad,p.rot,p.filter||gFilter]);
  async function addPage(canvas,q){
    const raw=scaleTo(canvas,260).toDataURL('image/jpeg',.7);
    const pg={id:'pg'+(++seq),src:await jpeg(canvas,.92),quad:q||FULL.map(p=>p.slice()),rot:0,filter:null,raw,thumb:'',key:'',out:null,exp:{}};
    pages.push(pg);stackUI();sub();pump();return pg;
  }
  let srcC={id:null,c:null};
  async function srcCanvas(pg){if(srcC.id===pg.id)return srcC.c;const b=await decode(pg.src);const c=mk(b.width,b.height);c.getContext('2d').drawImage(b,0,0);srcC={id:pg.id,c};return c}
  let pumping=false;
  async function pump(){if(pumping)return;pumping=true;try{for(;;){const pg=pages.find(p=>p.key!==keyOf(p));if(!pg||closed)break;const key=keyOf(pg);cardBusy(pg,true);await nextFrame();
      try{const src=await srcCanvas(pg);let c=warp(src,pg.quad,WORK);c=rotateC(c,pg.rot);applyFilter(c,pg.filter||gFilter);pg.out=await jpeg(c,.9);pg.thumb=scaleTo(c,420).toDataURL('image/jpeg',.8)}catch(e){console.warn(e)}
      pg.key=key;pg.exp={};cardBusy(pg,false);cardImg(pg);stackUI()}}finally{pumping=false}estimate()}
  function stackUI(){const n=pages.length,b=$('.scn-stack');b.disabled=!n;$('.scn-count').textContent=n;$('.scn-thumbs').innerHTML=pages.slice(-3).map(p=>`<img src="${p.thumb||p.raw}" alt="">`).join('')}
  async function addFiles(files,fromNative){
    if(!files.length)return;busy(true,tr('جارٍ تجهيز الصور…','Preparing images…'),'');let n=0;
    for(const f of files){try{const b=await decode(f);const c=scaleTo(b,2600);await addPage(c,detectOn(c));n++}catch(e){stoast(tr('تعذّرت قراءة صورة','Could not read an image'))}}
    busy(false);if(n){setScreen('rev')}
  }

  /* ---- review ---- */
  function cardHTML(p,i){return `<div class="scn-card" data-pid="${p.id}"><div class="scn-pg"><img src="${p.thumb||p.raw}" alt="${H(tr('صفحة','Page'))} ${i+1}" draggable="false"><span class="scn-no">${i+1}</span><span class="scn-grip">${IC.grip}</span></div>
    <div class="scn-acts"><button data-sc="edit" title="${H(tr('قص','Crop'))}" aria-label="${H(tr('قص','Crop'))}">${IC.crop}</button><button data-sc="rot" title="${H(tr('تدوير','Rotate'))}" aria-label="${H(tr('تدوير','Rotate'))}">${IC.rot}</button><button data-sc="del" title="${H(tr('حذف','Delete'))}" aria-label="${H(tr('حذف','Delete'))}">${IC.trash}</button></div></div>`}
  function renderGrid(){grid.innerHTML=pages.map(cardHTML).join('')+`<button class="scn-addc" data-sc="tocam">${IC.cam}<b>${H(tr('التقاط صفحة','Capture page'))}</b></button><button class="scn-addc" data-sc="gallery">${IC.img}<b>${H(tr('إضافة صور','Add images'))}</b></button>`;pages.forEach(p=>cardBusy(p,p.key!==keyOf(p)));sub();grid.classList.remove('ready');setTimeout(()=>grid.classList.add('ready'),700)}
  const card=pg=>grid.querySelector(`[data-pid="${pg.id}"]`);
  function cardBusy(pg,on){const c=card(pg);if(c)c.classList.toggle('busy',on)}
  function cardImg(pg){const c=card(pg);if(c)c.querySelector('img').src=pg.thumb||pg.raw}
  function renumber(){grid.querySelectorAll('.scn-card').forEach((c,i)=>{c.querySelector('.scn-no').textContent=i+1});sub()}
  async function filterBar(){
    const el=$('.scn-filters');const f0=FILTERS();el.innerHTML=f0.map(([k,l,ic])=>`<button class="scn-f" data-sc="filter" data-f="${k}" role="radio" aria-checked="${k===gFilter}"><span class="pv"><span class="ph"></span></span><span class="fl">${ic}${H(l)}</span></button>`).join('');
    const p=pages[0];if(!p)return;try{const src=await srcCanvas(p);const base=rotateC(warp(src,p.quad,150),p.rot);
      f0.forEach(([k])=>{const c=mk(base.width,base.height);c.getContext('2d').drawImage(base,0,0);applyFilter(c,k);const b=el.querySelector(`[data-f="${k}"] .pv`);if(b)b.innerHTML=`<img src="${c.toDataURL('image/jpeg',.8)}" alt="">`})}catch(e){}
  }
  let estT=0;function estimate(){clearTimeout(estT);estT=setTimeout(async()=>{if(root.dataset.screen!=='rev'||!pages.length)return;const el=$('.scn-est');el.classList.add('calc');
    try{let t=0;for(const p of pages){if(p.key!==keyOf(p))return;t+=(await expJpeg(p,preset)).u8.length}el.innerHTML=`<span>${H(tr('الحجم التقريبي','Approx. size'))}</span><b>${fsize(t+1500*pages.length)}</b>`}catch(e){}el.classList.remove('calc')},350)}
  async function expJpeg(p,pr){if(p.exp[pr])return p.exp[pr];const P=PRESETS[pr];const b=await decode(p.out);const c=scaleTo(b,P.max);const j=await jpeg(c,P.q);const r={u8:new Uint8Array(await j.arrayBuffer()),w:c.width,h:c.height};p.exp[pr]=r;return r}

  /* drag to reorder (mouse: drag; touch: long-press then drag) */
  let drag=null,lp=0;
  grid.addEventListener('pointerdown',e=>{const pg=e.target.closest('.scn-pg');if(!pg||e.button>0)return;const c=pg.closest('.scn-card');const st={c,x:e.clientX,y:e.clientY,id:e.pointerId,touch:e.pointerType!=='mouse',on:false};drag=st;
    if(st.touch){clearTimeout(lp);lp=setTimeout(()=>{if(drag===st)begin(st,e.clientX,e.clientY)},260)}});
  function begin(st,x,y){st.on=true;const r=st.c.getBoundingClientRect();st.ox=x-r.left;st.oy=y-r.top;st.c.classList.add('drag');grid.classList.add('dragging');vib(15);move(x,y)}
  function move(x,y){const st=drag;const g=grid.getBoundingClientRect();
    const tx=x-st.ox-(g.left+st.c.offsetLeft-grid.scrollLeft),ty=y-st.oy-(g.top+st.c.offsetTop-grid.scrollTop);st.c.style.transform=`translate(${tx}px,${ty}px) scale(1.05)`;
    st.c.style.pointerEvents='none';const over=document.elementFromPoint(x,y);st.c.style.pointerEvents='';const t=over&&over.closest('.scn-card');
    if(t&&t!==st.c&&grid.contains(t)&&!(t===st.last&&performance.now()-st.lt<360)){st.last=t;st.lt=performance.now();const all=[...grid.querySelectorAll('.scn-card')];const a=all.indexOf(st.c),b=all.indexOf(t);
      flipMove(()=>{if(a<b)t.after(st.c);else t.before(st.c)});const tx2=x-st.ox-(g.left+st.c.offsetLeft-grid.scrollLeft),ty2=y-st.oy-(g.top+st.c.offsetTop-grid.scrollTop);st.c.style.transform=`translate(${tx2}px,${ty2}px) scale(1.05)`}
    if(y<g.top+50)grid.scrollTop-=12;else if(y>g.bottom-50)grid.scrollTop+=12}
  function flipMove(mut){const cs=[...grid.querySelectorAll('.scn-card')].filter(c=>c!==drag.c);const r0=new Map(cs.map(c=>[c,c.getBoundingClientRect()]));mut();
    cs.forEach(c=>{const a=r0.get(c),b=c.getBoundingClientRect(),dx=a.left-b.left,dy=a.top-b.top;if(dx||dy){c.style.transition='none';c.style.transform=`translate(${dx}px,${dy}px)`;void c.offsetWidth;c.style.transition='transform .28s cubic-bezier(.2,.8,.2,1)';c.style.transform=''}})}
  d.addEventListener('pointermove',e=>{const st=drag;if(!st||e.pointerId!==st.id)return;if(!st.on){const m=Math.hypot(e.clientX-st.x,e.clientY-st.y);if(st.touch){if(m>8){clearTimeout(lp);drag=null}return}if(m>6)begin(st,e.clientX,e.clientY);return}e.preventDefault();move(e.clientX,e.clientY)});
  const endDrag=()=>{clearTimeout(lp);const st=drag;drag=null;if(!st||!st.on)return;st.c.classList.remove('drag');grid.classList.remove('dragging');st.c.style.transition='transform .25s cubic-bezier(.2,.8,.2,1)';st.c.style.transform='';setTimeout(()=>{st.c.style.transition=''},260);
    const order=[...grid.querySelectorAll('.scn-card')].map(c=>c.dataset.pid);pages.sort((a,b)=>order.indexOf(a.id)-order.indexOf(b.id));renumber();filterBar();st.c.dataset.justDragged='1';setTimeout(()=>delete st.c.dataset.justDragged,50)};
  d.addEventListener('pointerup',endDrag);d.addEventListener('pointercancel',endDrag);
  grid.addEventListener('touchmove',e=>{if(drag&&drag.on)e.preventDefault()},{passive:false});
  grid.addEventListener('contextmenu',e=>{if(e.target.closest('.scn-pg'))e.preventDefault()});

  /* ---- crop editor ---- */
  async function openEdit(pg){
    const src=await srcCanvas(pg);const c=$('.scn-ed-box canvas');const v=scaleTo(src,1800);c.width=v.width;c.height=v.height;c.getContext('2d').drawImage(v,0,0);
    edit={pg,q:pg.quad.map(p=>p.slice()),rot:pg.rot,filter:pg.filter,W:v.width,H:v.height,img:v};
    $('.scn-ed-box svg').setAttribute('viewBox',`0 0 ${v.width} ${v.height}`);
    $('.scn-hs').innerHTML=[0,1,2,3].map(i=>`<g class="scn-h" data-h="${i}"><circle class="hit"/><circle class="ring"/><circle class="core"/></g>`).join('');
    $('.scn-efl').innerHTML=[['',tr('مثل باقي الصفحات','Same as all'),'']].concat(FILTERS()).map(([k,l])=>`<button class="scn-chip" data-sc="efilter" data-f="${k}" aria-pressed="${(edit.filter||'')===k}">${H(l)}</button>`).join('');
    setScreen('edit');requestAnimationFrame(()=>{fitEdit();drawEdit()});
  }
  function fitEdit(){if(!edit)return;const st=$('.scn-ed-stage'),b=$('.scn-ed-box');const aw=st.clientWidth-40,ah=st.clientHeight-40;const s=Math.min(aw/edit.W,ah/edit.H);b.style.width=Math.round(edit.W*s)+'px';b.style.height=Math.round(edit.H*s)+'px';edit.s=s;drawEdit()}
  function drawEdit(){if(!edit)return;const P=edit.q.map(([x,y])=>[x*edit.W,y*edit.H]);const pts=P.map(p=>p.join(',')).join(' ');$('.scn-equad').setAttribute('points',pts);
    $('.scn-ed-box .scn-shade').setAttribute('d',`M0 0H${edit.W}V${edit.H}H0Z M${P.map(p=>p.join(' ')).join(' L')}Z`);const u=1/(edit.s||1);
    d.querySelectorAll('.scn-h').forEach((g,i)=>{g.querySelectorAll('circle').forEach(c=>{c.setAttribute('cx',P[i][0]);c.setAttribute('cy',P[i][1])});g.querySelector('.hit').setAttribute('r',26*u);g.querySelector('.ring').setAttribute('r',13*u);g.querySelector('.core').setAttribute('r',5*u)})}
  let eh=null;const loupe=$('.scn-loupe'),lx=loupe.getContext('2d');
  $('.scn-ed-box').addEventListener('pointerdown',e=>{const h=e.target.closest('.scn-h');if(!h||!edit)return;e.preventDefault();eh={i:+h.dataset.h,id:e.pointerId};h.classList.add('on');try{e.target.setPointerCapture(e.pointerId)}catch(x){}edMove(e)});
  function edMove(e){const b=$('.scn-ed-box').getBoundingClientRect();const x=clamp((e.clientX-b.left)/b.width,0,1),y=clamp((e.clientY-b.top)/b.height,0,1);edit.q[eh.i]=[x,y];drawEdit();
    const L=240,z=3/edit.s*.5;const sx=x*edit.W,sy=y*edit.H,r=L/2*(1/(edit.s*2.6));lx.fillStyle='#000';lx.fillRect(0,0,L,L);lx.imageSmoothingQuality='high';lx.drawImage(edit.img,sx-r,sy-r,2*r,2*r,0,0,L,L);
    lx.strokeStyle='#4fd1a5';lx.lineWidth=3;lx.beginPath();lx.moveTo(L/2,L/2-22);lx.lineTo(L/2,L/2+22);lx.moveTo(L/2-22,L/2);lx.lineTo(L/2+22,L/2);lx.stroke();
    loupe.hidden=false;const top=e.clientY-150<10?e.clientY+40:e.clientY-150;loupe.style.left=(e.clientX-60)+'px';loupe.style.top=top+'px';void z}
  d.addEventListener('pointermove',e=>{if(eh&&e.pointerId===eh.id){e.preventDefault();edMove(e)}});
  const edEnd=()=>{if(!eh)return;eh=null;loupe.hidden=true;d.querySelectorAll('.scn-h.on').forEach(h=>h.classList.remove('on'))};d.addEventListener('pointerup',edEnd);d.addEventListener('pointercancel',edEnd);
  addEventListener('resize',()=>{if(root.dataset.screen==='edit')fitEdit()});

  /* ---- screens, busy, finish ---- */
  function setScreen(s){root.dataset.screen=s;if(s==='cam'){camStart();sub()}else camStop();if(s==='rev'){renderGrid();filterBar();estimate()}}
  function busy(on,t,s){const b=$('.scn-busy');b.hidden=!on;b.classList.remove('done');if(on){$('.scn-busy-t').textContent=t||'';$('.scn-busy-s').textContent=s||'';$('.scn-busy .r1').style.strokeDashoffset='213.6'}}
  const bprog=p=>{$('.scn-busy .r1').style.strokeDashoffset=String(213.6*(1-p))};
  async function finish(){
    if(!pages.length)return;let name=($('.scn-name input').value||defName).replace(/[\\/:*?"<>|]+/g,'-').trim()||defName;name=name.replace(/\.pdf$/i,'')+'.pdf';
    busy(true,tr('جارٍ إنشاء ملف PDF…','Creating PDF…'),'');
    try{await pump();while(pumping)await sleep(80);const out=[];for(let i=0;i<pages.length;i++){$('.scn-busy-s').textContent=tr(`الصفحة ${i+1} من ${pages.length}`,`Page ${i+1} of ${pages.length}`);out.push(await expJpeg(pages[i],preset));bprog((i+1)/(pages.length+.4))}
      const blob=makePDF(out,name.replace(/\.pdf$/,''));bprog(1);$('.scn-busy').classList.add('done');$('.scn-busy-t').textContent=tr('تم إنشاء المستند','Document ready');$('.scn-busy-s').textContent=pagesTxt(pages.length)+' · '+fsize(blob.size);vib([20,40,20]);
      await sleep(900);done({blob,name,pages:pages.length})}
    catch(e){busy(false);stoast(tr('تعذّر إنشاء الملف','Could not create the file'))}
  }
  function done(r){if(closed)return;closed=true;camStop();d.classList.add('out');setTimeout(()=>{try{d.close()}catch(e){}d.remove()},260);scanBusy=false;resolve(r||null)}
  function askClose(){if(!pages.length)return done(null);$('.scn-sheet-p').textContent=tr(`ستُحذف ${pagesTxt(pages.length)} ملتقطة.`,`${pagesTxt(pages.length)} will be discarded.`);$('.scn-sheet').hidden=false}
  d.addEventListener('cancel',e=>{e.preventDefault();if(!$('.scn-sheet').hidden){$('.scn-sheet').hidden=true;return}const s=root.dataset.screen;if(s==='edit'){edit=null;setScreen('rev')}else askClose()});
  d.addEventListener('keydown',e=>{e.stopPropagation();if(root.dataset.screen==='cam'&&(e.key===' '||e.key==='Enter')&&!e.target.closest('input')){e.preventDefault();shoot()}});
  gal.onchange=()=>{addFiles([...gal.files]);gal.value=''};nat.onchange=()=>{addFiles([...nat.files],true);nat.value=''};
  d.addEventListener('click',async e=>{const b=e.target.closest('[data-sc]');if(!b)return;const a=b.dataset.sc,c=b.closest('.scn-card'),pg=c&&pages.find(p=>p.id===c.dataset.pid);
    if(a==='close')return askClose();
    if(a==='keep'){$('.scn-sheet').hidden=true;return}
    if(a==='discard')return done(null);
    if(a==='shoot')return shoot();
    if(a==='auto'){auto=!auto;store('auto',auto?'1':'0');b.setAttribute('aria-pressed',auto);stoast(auto?tr('الالتقاط التلقائي مفعّل','Auto-capture on'):tr('الالتقاط التلقائي متوقف','Auto-capture off'));return}
    if(a==='torch'&&track){torch=!torch;try{await track.applyConstraints({advanced:[{torch}]});b.setAttribute('aria-pressed',torch)}catch(err){torch=!torch}return}
    if(a==='switch'&&devs.length>1){devIdx=(devIdx+1)%devs.length;camStop();camStart();return}
    if(a==='gallery'){gal.click();return}
    if(a==='native'){nat.click();return}
    if(a==='retry'){camStart();return}
    if(a==='review'){setScreen('rev');return}
    if(a==='tocam'){setScreen('cam');return}
    if(a==='filter'){gFilter=b.dataset.f;store('filter',gFilter);$('.scn-filters').querySelectorAll('[data-f]').forEach(x=>x.setAttribute('aria-checked',x===b));pages.forEach(p=>cardBusy(p,p.key!==keyOf(p)));pump();return}
    if(a==='preset'){preset=b.dataset.p;store('preset',preset);d.querySelectorAll('[data-sc="preset"]').forEach(x=>x.setAttribute('aria-pressed',x===b));estimate();return}
    if(a==='finish')return finish();
    if(a==='rot'&&pg){pg.rot=(pg.rot+1)%4;const im=c.querySelector('img');im.style.transition='transform .3s';im.style.transform='rotate(90deg)';setTimeout(()=>{im.style.transition='none';im.style.transform=''},300);pump();if(pages[0]===pg)filterBar();return}
    if(a==='del'&&pg){c.classList.add('bye');setTimeout(()=>{pages=pages.filter(p=>p!==pg);c.remove();renumber();stackUI();estimate();if(!pages.length)setScreen('cam');else filterBar()},260);return}
    if(a==='edit'&&pg)return openEdit(pg);
    if(a==='ecancel'){edit=null;setScreen('rev');return}
    if(a==='edetect'&&edit){const q=detectOn(edit.img);if(q){edit.q=q;drawEdit();stoast(tr('تم اكتشاف حواف الورقة','Paper edges found'))}else stoast(tr('لم نتمكن من اكتشاف الحواف، اضبطها يدويًا','Could not find the edges, adjust them by hand'));return}
    if(a==='efull'&&edit){edit.q=FULL.map(p=>p.slice());drawEdit();return}
    if(a==='erot'&&edit){edit.rot=(edit.rot+1)%4;stoast(tr('ستُدوَّر الصفحة ','Page will be rotated ')+edit.rot*90+'°');return}
    if(a==='efilter'&&edit){edit.filter=b.dataset.f||null;d.querySelectorAll('[data-sc="efilter"]').forEach(x=>x.setAttribute('aria-pressed',x===b));return}
    if(a==='eapply'&&edit){const p=edit.pg;if(area(edit.q)<.02){stoast(tr('منطقة القص صغيرة جدًا','Crop area is too small'));return}p.quad=edit.q;p.rot=edit.rot;p.filter=edit.filter;edit=null;setScreen('rev');pump();if(pages[0]===p)filterBar();return}
  });
  d.addEventListener('click',e=>{const c=e.target.closest('.scn-pg');if(!c||e.target.closest('[data-sc]'))return;const card_=c.closest('.scn-card');if(card_.dataset.justDragged)return;const pg=pages.find(p=>p.id===card_.dataset.pid);if(pg)openEdit(pg)});
  d.showModal();sub();setScreen('cam');
 });
}

/* ---------------- styles ---------------- */
const css=document.createElement('style');css.textContent=`
.att{display:flex;flex-direction:column;gap:10px;min-width:0}
.att-zone{position:relative;display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding:14px 16px;border:1.5px dashed color-mix(in srgb,var(--accent) 45%,var(--line));border-radius:14px;background:linear-gradient(135deg,var(--accent-soft),var(--surface) 70%);transition:border-color .2s,box-shadow .2s,transform .2s}
.att-zone.over{border-style:solid;border-color:var(--accent);box-shadow:0 0 0 5px color-mix(in srgb,var(--accent) 16%,transparent);transform:scale(1.01)}
.att-zi{flex:none;width:42px;height:42px;border-radius:12px;display:grid;place-items:center;background:var(--surface);color:var(--accent);box-shadow:var(--shadow)}.att-zi svg{width:21px;height:21px}
.att-zt{flex:1;min-width:190px;display:flex;flex-direction:column;gap:2px}.att-zt b{font-size:14px}.att-zt small{color:var(--muted);font-size:12.5px;line-height:1.5}
.att-zb{display:flex;gap:8px;flex-wrap:wrap}
.att-btn{all:unset;box-sizing:border-box;cursor:pointer;display:inline-flex;align-items:center;gap:7px;height:38px;padding:0 15px;border-radius:10px;border:1px solid var(--line);background:var(--surface);font-size:13.5px;font-weight:600;color:var(--ink);transition:border-color .15s,color .15s,transform .15s,box-shadow .15s;white-space:nowrap}
.att-btn svg{width:17px;height:17px;flex:none}.att-btn:hover{border-color:var(--accent);color:var(--accent);transform:translateY(-1px)}.att-btn:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.att-btn.att-scan{position:relative;overflow:hidden;color:#fff;border-color:transparent;background:linear-gradient(135deg,#0b5c8a,#1784c4 60%,#2aa3a0);box-shadow:0 6px 18px rgba(11,92,138,.32)}
.att-btn.att-scan:hover{color:#fff;box-shadow:0 10px 24px rgba(11,92,138,.42)}
.att-btn.att-scan::after{content:'';position:absolute;inset:0;background:linear-gradient(105deg,transparent 30%,rgba(255,255,255,.4) 50%,transparent 70%);transform:translateX(-130%);animation:attShine 3.6s ease-in-out infinite}
@keyframes attShine{0%,62%{transform:translateX(-130%)}100%{transform:translateX(130%)}}
.att-list{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,250px),1fr));gap:8px}
.att-list:empty{display:none}
.att-item{display:flex;align-items:center;gap:2px;border:1px solid var(--line);border-radius:12px;background:var(--surface);padding:6px;min-width:0;animation:attIn .4s cubic-bezier(.2,.8,.2,1);transition:border-color .15s,box-shadow .15s}
.att-item:hover{border-color:color-mix(in srgb,var(--accent) 40%,var(--line));box-shadow:var(--shadow)}
@keyframes attIn{from{opacity:0;transform:translateY(6px) scale(.98)}}
.att-open{all:unset;cursor:pointer;flex:1;min-width:0;display:flex;align-items:center;gap:10px;padding:2px 4px;border-radius:8px}.att-open:focus-visible{outline:2px solid var(--accent)}
.att-ic{flex:none;position:relative;width:36px;height:44px;border-radius:6px 12px 6px 6px;display:grid;place-items:end center;padding-bottom:6px;color:#fff;font-size:9.5px;font-weight:800;letter-spacing:.4px;box-shadow:inset 0 -14px 20px rgba(0,0,0,.12)}
.att-ic::before{content:'';position:absolute;top:0;inset-inline-end:0;width:11px;height:11px;background:linear-gradient(225deg,var(--surface) 50%,rgba(255,255,255,.55) 50%);border-radius:0 12px 0 4px}
[dir=rtl] .att-ic::before{background:linear-gradient(135deg,var(--surface) 50%,rgba(255,255,255,.55) 50%);border-radius:12px 0 4px 0}
.att-ic svg{position:absolute;top:8px;left:50%;width:15px;height:15px;transform:translateX(-50%);opacity:.95}
.att-ic.k-pdf{background:#d93a2b}.att-ic.k-img{background:#1d8a5a}.att-ic.k-doc{background:#2b5fd9}.att-ic.k-xls{background:#157347}.att-ic.k-file{background:#667585}
.att-ic.big{width:64px;height:78px;font-size:13px;padding-bottom:10px}
.att-meta{min-width:0;flex:1;display:flex;flex-direction:column;gap:2px}.att-meta>b{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.att-meta small{font-size:11.5px;color:var(--muted)}.att-meta em{font-style:normal;color:var(--accent);font-weight:600}
.att-x{all:unset;cursor:pointer;flex:none;width:34px;height:34px;display:grid;place-items:center;border-radius:9px;color:var(--muted);transition:background .15s,color .15s}.att-x svg{width:17px;height:17px}.att-x:hover{background:var(--surface2);color:var(--ink)}
.att-x.rm:hover{color:var(--crit)}.att-x.armed{background:var(--crit);color:#fff;animation:attWig .3s}.att-x.spin svg{animation:attSpin 1s linear infinite}
@keyframes attWig{25%{transform:rotate(-8deg)}75%{transform:rotate(8deg)}}@keyframes attSpin{to{transform:rotate(360deg)}}
.att-item.pend{border-style:dashed}.att-item.pend .att-ic{opacity:.75}
.att-bar{display:block;height:4px;border-radius:4px;background:var(--line2);overflow:hidden;margin-top:3px}.att-bar i{display:block;height:100%;border-radius:4px;background:linear-gradient(90deg,var(--accent),#2aa3a0,var(--accent));background-size:200% 100%;animation:attBar 1.2s linear infinite;transition:width .35s}
@keyframes attBar{to{background-position:-200% 0}}
.att-empty{color:var(--muted);font-size:13px}
.att-chips{display:inline-flex;flex-wrap:wrap;gap:6px;margin-top:4px}.att-chip{all:unset;cursor:pointer;display:inline-flex;align-items:center;gap:5px;max-width:220px;font-size:12px;padding:3px 9px;border-radius:999px;background:var(--accent-soft);color:var(--accent)}.att-chip span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.att-chip svg{width:13px;height:13px;flex:none}.att-chip:hover{filter:brightness(.96)}
dialog.attv{width:min(1100px,calc(100% - 20px));height:calc(100dvh - 20px);max-height:none;padding:0;overflow:hidden;background:var(--surface)}
dialog.attv[open]{display:flex;flex-direction:column;animation:attPop .3s cubic-bezier(.2,.8,.2,1)}dialog.attv::backdrop{background:rgba(5,12,20,.72);backdrop-filter:blur(4px)}
@keyframes attPop{from{opacity:0;transform:scale(.96)}}
.attv-head{display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid var(--line)}.attv-t{flex:1;min-width:0;display:flex;flex-direction:column}.attv-t b{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:14.5px}.attv-t small{color:var(--muted);font-size:12px}
.attv-b{all:unset;cursor:pointer;display:inline-flex;align-items:center;gap:6px;height:36px;min-width:36px;justify-content:center;padding:0 8px;border-radius:9px;font-size:13px;font-weight:600}.attv-b:hover{background:var(--surface2)}.attv-b svg{width:18px;height:18px}
@media (max-width:600px){.attv-b span{display:none}}
.attv-body{position:relative;flex:1;min-height:0;display:grid;place-items:center;background:var(--surface2);overflow:auto}
.attv-body iframe{position:absolute;inset:0;width:100%;height:100%;border:0;background:#fff}
.attv-body>img{max-width:100%;max-height:100%;object-fit:contain;cursor:zoom-in;box-shadow:0 10px 40px rgba(0,0,0,.18)}.attv-body.z{place-items:start}.attv-body.z>img{max-width:none;max-height:none;cursor:zoom-out}
.attv-pages{position:absolute;inset:0;overflow:auto;display:flex;flex-direction:column;align-items:center;gap:18px;padding:18px 16px 30px}
.attv-pages figure{margin:0;display:flex;flex-direction:column;align-items:center;gap:6px;animation:attIn .4s cubic-bezier(.2,.8,.2,1)}.attv-pages canvas{background:#fff;border-radius:4px;box-shadow:0 8px 30px rgba(15,34,54,.18)}.attv-pages figcaption{font-size:11.5px;color:var(--muted)}
.attv-card{display:flex;flex-direction:column;align-items:center;gap:10px;text-align:center;padding:30px}.attv-card>span{color:var(--muted);font-size:13px}
.attv-nav{all:unset;cursor:pointer;position:absolute;top:50%;transform:translateY(-50%);width:44px;height:44px;border-radius:50%;display:grid;place-items:center;background:rgba(15,34,54,.6);color:#fff;z-index:2}.attv-nav svg{width:22px;height:22px}.attv-nav.p{inset-inline-start:12px}.attv-nav.n{inset-inline-end:12px}
[dir=rtl] .flip svg{transform:scaleX(-1)}
.att-spin{width:38px;height:38px;border-radius:50%;border:3px solid var(--line);border-top-color:var(--accent);animation:attSpin .8s linear infinite}

dialog.scn{position:fixed;inset:0;width:100vw;height:100vh;height:100dvh;max-width:none;max-height:none;margin:0;padding:0;border:0;border-radius:0;background:#03070c;color:#eef4fa;overflow:hidden;box-shadow:none;--ac:#4fd1a5;--ac2:#63b3e3}
dialog.scn[open]{animation:scnIn .4s cubic-bezier(.2,.8,.2,1)}dialog.scn.out{animation:scnOut .26s ease-in forwards}dialog.scn::backdrop{background:#000}
@keyframes scnIn{from{opacity:0;transform:scale(1.04)}}@keyframes scnOut{to{opacity:0;transform:scale(.97)}}
.scn-root{position:absolute;inset:0;font-family:var(--f-body);-webkit-tap-highlight-color:transparent}
.scn-root button{font:inherit;color:inherit}
.scn-root>section{position:absolute;inset:0;display:flex;flex-direction:column;opacity:0;visibility:hidden;transform:translateY(14px);transition:opacity .32s,transform .32s cubic-bezier(.2,.8,.2,1),visibility 0s .32s}
.scn-root[data-screen=cam] .scn-cam,.scn-root[data-screen=rev] .scn-rev,.scn-root[data-screen=edit] .scn-edit{opacity:1;visibility:visible;transform:none;transition:opacity .32s,transform .32s cubic-bezier(.2,.8,.2,1)}
.scn-top{display:flex;align-items:center;gap:10px;padding:max(12px,env(safe-area-inset-top)) 14px 12px;z-index:3}
.scn-cam .scn-top{position:absolute;top:0;inset-inline:0;background:linear-gradient(rgba(0,0,0,.65),transparent)}
.scn-top.solid{background:#0a121b;border-bottom:1px solid rgba(255,255,255,.07)}
.scn-ib{all:unset;cursor:pointer;flex:none;width:42px;height:42px;border-radius:50%;display:grid;place-items:center;background:rgba(255,255,255,.1);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);transition:background .2s,transform .15s}.scn-ib:hover{background:rgba(255,255,255,.18)}.scn-ib:active{transform:scale(.92)}.scn-ib svg{width:20px;height:20px}
.scn-ib[aria-pressed=true]{background:#ffd166;color:#1a1300}
.scn-title{flex:1;min-width:0;display:flex;flex-direction:column;gap:1px}.scn-title b{font-size:15.5px;font-weight:700}.scn-title small{font-size:12px;opacity:.7;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.scn-pill{all:unset;cursor:pointer;display:inline-flex;align-items:center;gap:6px;height:36px;padding:0 13px;border-radius:999px;background:rgba(255,255,255,.1);backdrop-filter:blur(10px);font-size:13px;font-weight:700;transition:background .2s,color .2s}.scn-pill svg{width:17px;height:17px}
.scn-pill[aria-pressed=true]{background:var(--ac);color:#03251a}
.scn-stage{position:relative;flex:1;min-height:0;overflow:hidden;background:radial-gradient(circle at 50% 40%,#101c28,#000 70%)}
.scn-stage video{position:absolute;inset:0;width:100%;height:100%;object-fit:contain}
.scn-ov{position:absolute;inset:0;width:100%;height:100%;pointer-events:none;overflow:visible}
.scn-shade{fill:rgba(2,6,12,.5);fill-rule:evenodd;opacity:0;transition:opacity .35s}.det .scn-shade,.scn-shade.on{opacity:1}
.scn-quad{fill:transparent;stroke:rgba(255,255,255,.85);stroke-width:2;stroke-dasharray:14 10;vector-effect:non-scaling-stroke;stroke-linejoin:round;transition:stroke .3s,fill .3s;animation:scnDash 1.2s linear infinite}
@keyframes scnDash{to{stroke-dashoffset:-24}}
.det .scn-quad{stroke:var(--ac);stroke-width:3;stroke-dasharray:none;animation:none;fill:rgba(79,209,165,.07);filter:drop-shadow(0 0 10px rgba(79,209,165,.85))}
.hold .scn-quad{fill:rgba(79,209,165,.16)}
.scn-cn{fill:#fff;stroke:var(--ac);stroke-width:3;vector-effect:non-scaling-stroke;opacity:0;transition:opacity .3s}.det .scn-cn{opacity:1}
.scn-hint{position:absolute;top:calc(max(12px,env(safe-area-inset-top)) + 66px);left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:9px;width:max-content;max-width:calc(100% - 32px);padding:9px 16px;border-radius:20px;background:rgba(8,16,26,.72);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);font-size:13.5px;font-weight:600;line-height:1.45;border:1px solid rgba(255,255,255,.08);transition:background .3s,border-color .3s;z-index:2}
.scn-hint-t{text-align:center}.scn-hint:has(.scn-hint-t:empty){opacity:0}
.scn-hint .dot{flex:none;width:9px;height:9px;border-radius:50%;background:#ffd166;animation:scnPulse 1.6s infinite}.det .scn-hint .dot{background:var(--ac)}.det .scn-hint{border-color:rgba(79,209,165,.45)}
@keyframes scnPulse{0%{box-shadow:0 0 0 0 rgba(255,209,102,.55)}70%{box-shadow:0 0 0 10px rgba(255,209,102,0)}100%{box-shadow:0 0 0 0 rgba(255,209,102,0)}}
.scn-flash{position:absolute;inset:0;background:#fff;opacity:0;pointer-events:none;z-index:4}.scn-flash.go{animation:scnFlash .5s ease-out}@keyframes scnFlash{0%{opacity:.85}100%{opacity:0}}
.scn-bot{position:absolute;bottom:0;inset-inline:0;display:flex;align-items:center;justify-content:space-between;padding:26px 28px max(24px,env(safe-area-inset-bottom));background:linear-gradient(transparent,rgba(0,0,0,.8));z-index:3}
.scn-shutter{all:unset;cursor:pointer;position:relative;width:80px;height:80px;border-radius:50%;display:grid;place-items:center}
.scn-shutter::before{content:'';width:62px;height:62px;border-radius:50%;background:#fff;box-shadow:0 4px 20px rgba(0,0,0,.4);transition:transform .15s,background .2s}.scn-shutter:hover::before{transform:scale(.96)}.scn-shutter:active::before{transform:scale(.86)}
.det .scn-shutter::before{background:#eafff7}
.scn-shutter svg{position:absolute;inset:0;width:100%;height:100%;transform:rotate(-90deg)}.scn-shutter .r0{fill:none;stroke:rgba(255,255,255,.55);stroke-width:4}.scn-shutter .r1{fill:none;stroke:var(--ac);stroke-width:5;stroke-linecap:round;stroke-dasharray:232.5;stroke-dashoffset:232.5;transition:stroke-dashoffset .12s linear}
[dir=rtl] .scn-shutter svg{transform:rotate(-90deg) scaleY(-1)}
.scn-side{all:unset;cursor:pointer;position:relative;width:68px;display:flex;flex-direction:column;align-items:center;gap:7px;font-size:11.5px;font-weight:600;transition:opacity .2s,transform .15s}.scn-side:active{transform:scale(.94)}
.scn-sq{width:52px;height:52px;border-radius:14px;background:rgba(255,255,255,.13);backdrop-filter:blur(10px);display:grid;place-items:center}.scn-sq svg{width:23px;height:23px}
.scn-thumbs{position:relative;width:52px;height:52px;border-radius:12px;background:rgba(255,255,255,.08)}
.scn-thumbs img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;border-radius:9px;border:2px solid #fff;box-shadow:0 4px 14px rgba(0,0,0,.5)}
.scn-thumbs img:nth-last-child(2){transform:rotate(-7deg) translateX(-3px)}.scn-thumbs img:nth-last-child(3){transform:rotate(7deg) translateX(3px)}
.scn-count{position:absolute;top:-8px;inset-inline-end:2px;min-width:24px;height:24px;padding:0 6px;border-radius:12px;background:#f08350;color:#fff;font-style:normal;font-size:12.5px;font-weight:800;display:grid;place-items:center;box-shadow:0 2px 8px rgba(0,0,0,.4)}
.scn-count.pop{animation:scnPop .45s cubic-bezier(.3,1.6,.5,1)}@keyframes scnPop{0%{transform:scale(.4)}100%{transform:scale(1)}}
.scn-stack[disabled]{opacity:.35;pointer-events:none}.scn-stack[disabled] .scn-count{display:none}
.scn-fly{position:fixed;z-index:9;border-radius:6px;object-fit:cover;box-shadow:0 18px 40px rgba(0,0,0,.6);pointer-events:none;transition:left .6s cubic-bezier(.55,0,.2,1),top .6s cubic-bezier(.55,0,.2,1),width .6s cubic-bezier(.55,0,.2,1),height .6s cubic-bezier(.55,0,.2,1),opacity .6s,transform .6s}
.scn-fail{position:absolute;inset:0;display:grid;place-items:center;padding:90px 18px 150px;z-index:2}
.scn-fcard{width:min(420px,100%);display:flex;flex-direction:column;align-items:center;gap:10px;text-align:center;padding:28px 22px;border-radius:22px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);backdrop-filter:blur(16px)}
.scn-fic{width:64px;height:64px;border-radius:50%;display:grid;place-items:center;background:linear-gradient(135deg,#0b5c8a,#2aa3a0);margin-bottom:4px}.scn-fic svg{width:30px;height:30px}.scn-fcard b{font-size:17px}.scn-fcard p{margin:0 0 6px;font-size:13.5px;line-height:1.7;opacity:.8}
.scn-fbtns{display:flex;flex-direction:column;gap:9px;width:100%}
.scn-go,.scn-ghost,.scn-danger{all:unset;box-sizing:border-box;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;height:44px;padding:0 18px;border-radius:12px;font-size:14px;font-weight:700;transition:transform .15s,box-shadow .2s,background .2s;white-space:nowrap}
.scn-go svg,.scn-ghost svg,.scn-danger svg{width:19px;height:19px;flex:none}.scn-go:active,.scn-ghost:active,.scn-danger:active{transform:scale(.97)}
.scn-go{background:linear-gradient(135deg,#2aa3a0,var(--ac));color:#022019;box-shadow:0 8px 24px rgba(79,209,165,.35)}.scn-go:hover{box-shadow:0 10px 30px rgba(79,209,165,.5)}
.scn-top .scn-go{height:40px;padding:0 16px}
.scn-ghost{background:rgba(255,255,255,.1)}.scn-ghost:hover{background:rgba(255,255,255,.16)}.scn-danger{background:#e5484d;color:#fff}
.scn-rev,.scn-edit{background:#060c13}
.scn-filters{display:flex;gap:12px;overflow-x:auto;padding:14px 16px 6px;scrollbar-width:none;justify-content:center}.scn-filters::-webkit-scrollbar{display:none}
.scn-f{all:unset;cursor:pointer;flex:none;display:flex;flex-direction:column;align-items:center;gap:7px;font-size:12px;font-weight:600;opacity:.7;transition:opacity .2s,transform .2s}.scn-f:hover{opacity:.95}
.scn-f .pv{width:62px;height:80px;border-radius:11px;overflow:hidden;border:2px solid rgba(255,255,255,.12);background:#152230;display:grid;place-items:center;transition:border-color .2s,box-shadow .2s,transform .2s}.scn-f .pv img{width:100%;height:100%;object-fit:cover}
.scn-f .ph{width:20px;height:20px;border-radius:50%;border:2px solid rgba(255,255,255,.2);border-top-color:var(--ac2);animation:attSpin .8s linear infinite}
.scn-f .fl{display:inline-flex;align-items:center;gap:4px}.scn-f .fl svg{width:13px;height:13px}
.scn-f[aria-checked=true]{opacity:1}.scn-f[aria-checked=true] .pv{border-color:var(--ac);box-shadow:0 0 0 3px rgba(79,209,165,.25),0 8px 22px rgba(0,0,0,.4);transform:translateY(-2px)}.scn-f[aria-checked=true] .fl{color:var(--ac)}
.scn-grid{position:relative;flex:1;min-height:0;overflow:auto;padding:14px 18px 28px;display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:22px 18px;align-content:start}
@media (max-width:520px){.scn-grid{grid-template-columns:repeat(2,1fr);gap:18px 14px;padding:12px 14px 24px}}
.scn-card{position:relative;animation:scnCard .45s cubic-bezier(.2,.8,.2,1) backwards;touch-action:pan-y;user-select:none;-webkit-user-select:none;-webkit-touch-callout:none}
.scn-card:nth-child(2){animation-delay:.04s}.scn-card:nth-child(3){animation-delay:.08s}.scn-card:nth-child(4){animation-delay:.12s}.scn-card:nth-child(n+5){animation-delay:.16s}
.scn-grid.ready .scn-card{animation:none}
@keyframes scnCard{from{opacity:0;transform:translateY(16px) scale(.96)}}
.scn-pg{position:relative;aspect-ratio:1/1.32;border-radius:12px;overflow:hidden;background:#111c27;display:grid;place-items:center;box-shadow:0 10px 28px rgba(0,0,0,.45);cursor:grab;transition:box-shadow .2s,transform .2s}
.scn-pg:hover{box-shadow:0 14px 34px rgba(0,0,0,.55),0 0 0 2px rgba(99,179,227,.5)}
.scn-pg img{max-width:100%;max-height:100%;object-fit:contain;pointer-events:none;background:#fff}
.scn-no{position:absolute;top:8px;inset-inline-start:8px;min-width:26px;height:26px;padding:0 7px;border-radius:13px;background:rgba(3,7,12,.78);font-size:12.5px;font-weight:800;display:grid;place-items:center;backdrop-filter:blur(6px)}
.scn-grip{position:absolute;top:8px;inset-inline-end:8px;width:26px;height:26px;border-radius:8px;background:rgba(3,7,12,.6);display:grid;place-items:center;opacity:0;transition:opacity .2s}.scn-grip svg{width:15px;height:15px}.scn-card:hover .scn-grip{opacity:1}
@media (pointer:coarse){.scn-grip{opacity:.85}}
.scn-card.busy .scn-pg::after{content:'';position:absolute;inset-inline:0;top:-40%;height:40%;background:linear-gradient(transparent,rgba(79,209,165,.5),transparent);animation:scnSweep 1.1s linear infinite}
@keyframes scnSweep{to{top:100%}}
.scn-card.drag{z-index:10;transition:none!important}.scn-card.drag .scn-pg{cursor:grabbing;box-shadow:0 30px 60px rgba(0,0,0,.7),0 0 0 3px var(--ac2)}.scn-grid.dragging .scn-card:not(.drag) .scn-acts{opacity:.3}
.scn-card.bye{transition:transform .26s,opacity .26s!important;transform:scale(.7)!important;opacity:0}
.scn-acts{display:flex;justify-content:center;gap:6px;margin-top:9px;transition:opacity .2s}
.scn-acts button{all:unset;cursor:pointer;width:36px;height:34px;border-radius:10px;display:grid;place-items:center;background:rgba(255,255,255,.07);transition:background .15s,color .15s}.scn-acts button svg{width:17px;height:17px}.scn-acts button:hover{background:rgba(255,255,255,.15)}.scn-acts [data-sc=del]:hover{background:#e5484d;color:#fff}
.scn-addc{all:unset;box-sizing:border-box;cursor:pointer;aspect-ratio:1/1.32;border-radius:12px;border:2px dashed rgba(255,255,255,.18);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;font-size:13px;color:rgba(255,255,255,.75);transition:border-color .2s,background .2s,color .2s}.scn-addc svg{width:28px;height:28px}.scn-addc:hover{border-color:var(--ac);color:var(--ac);background:rgba(79,209,165,.06)}
.scn-done{display:flex;flex-wrap:wrap;align-items:flex-end;gap:12px 16px;padding:14px 18px max(14px,env(safe-area-inset-bottom));background:#0a121b;border-top:1px solid rgba(255,255,255,.07)}
.scn-name{flex:1 1 260px;display:flex;flex-direction:column;gap:5px;font-size:12px;opacity:.95}.scn-name>span:first-child{opacity:.65}
.scn-in{display:flex;align-items:center;height:42px;border-radius:11px;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);padding-inline:12px;transition:border-color .2s}.scn-in:focus-within{border-color:var(--ac2)}
.scn-in input{flex:1;min-width:0;background:none;border:0;outline:0;color:inherit;font:inherit;font-size:14px}.scn-in em{font-style:normal;opacity:.5;font-size:13px}
.scn-q{display:flex;padding:3px;border-radius:12px;background:rgba(255,255,255,.07);height:42px;box-sizing:border-box}
.scn-q button{all:unset;cursor:pointer;padding:0 13px;border-radius:9px;font-size:13px;font-weight:600;opacity:.7;transition:background .2s,opacity .2s;display:grid;place-items:center}.scn-q button[aria-pressed=true]{background:rgba(255,255,255,.16);opacity:1}
.scn-est{display:flex;flex-direction:column;justify-content:center;height:42px;font-size:11.5px;min-width:86px}.scn-est span{opacity:.6}.scn-est b{font-size:14px;color:var(--ac)}.scn-est.calc{opacity:.5}
.scn-go.big{height:48px;padding:0 22px;font-size:15px;margin-inline-start:auto}
@media (max-width:640px){.scn-done{gap:10px}.scn-go.big{flex:1 1 100%;margin:0}.scn-q{flex:1}.scn-q button{flex:1}}
.scn-ed-stage{position:relative;flex:1;min-height:0;display:grid;place-items:center;overflow:hidden;touch-action:none;background:radial-gradient(circle at 50% 40%,#0f1b27,#03070c 75%)}
.scn-ed-box{position:relative;box-shadow:0 20px 60px rgba(0,0,0,.6)}.scn-ed-box canvas{display:block;width:100%;height:100%}.scn-ed-box svg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
.scn-equad{fill:rgba(79,209,165,.1);stroke:var(--ac);stroke-width:2.5;vector-effect:non-scaling-stroke;stroke-linejoin:round}
.scn-h{cursor:grab}.scn-h .hit{fill:transparent}.scn-h .ring{fill:rgba(79,209,165,.25);stroke:#fff;stroke-width:2.5;vector-effect:non-scaling-stroke;transition:fill .2s}.scn-h .core{fill:#fff}.scn-h.on .ring{fill:rgba(79,209,165,.55)}
.scn-loupe{position:fixed;width:120px;height:120px;border-radius:50%;border:3px solid #fff;box-shadow:0 12px 30px rgba(0,0,0,.6),0 0 0 2px var(--ac);pointer-events:none;z-index:20}
.scn-ed-tools{display:flex;flex-direction:column;gap:10px;padding:12px 16px max(14px,env(safe-area-inset-bottom));background:#0a121b;border-top:1px solid rgba(255,255,255,.07)}
.scn-tl{display:flex;gap:8px;overflow-x:auto;scrollbar-width:none;justify-content:center}
.scn-tool{all:unset;cursor:pointer;flex:none;display:inline-flex;align-items:center;gap:7px;height:40px;padding:0 14px;border-radius:11px;background:rgba(255,255,255,.08);font-size:13px;font-weight:600}.scn-tool svg{width:18px;height:18px}.scn-tool:hover{background:rgba(255,255,255,.15)}
.scn-chip{all:unset;cursor:pointer;flex:none;height:32px;padding:0 13px;border-radius:999px;border:1px solid rgba(255,255,255,.15);font-size:12.5px;font-weight:600;display:grid;place-items:center;opacity:.8}.scn-chip[aria-pressed=true]{background:var(--ac);border-color:var(--ac);color:#022019;opacity:1}
.scn-sheet{position:absolute;inset:0;z-index:30;display:grid;align-items:end;justify-items:center;background:rgba(0,0,0,.55);backdrop-filter:blur(3px);animation:scnFade .2s}
.scn-sheet-c{width:min(440px,100%);margin:0 0 max(16px,env(safe-area-inset-bottom));padding:22px;border-radius:22px;background:#111b26;border:1px solid rgba(255,255,255,.1);display:flex;flex-direction:column;gap:8px;animation:scnUp .3s cubic-bezier(.2,.8,.2,1)}
.scn-sheet-c b{font-size:17px}.scn-sheet-c p{margin:0 0 8px;opacity:.75;font-size:14px}.scn-sheet-c .scn-fbtns{flex-direction:row}.scn-sheet-c .scn-fbtns>*{flex:1}
@media (min-width:700px){.scn-sheet{align-items:center}}
@keyframes scnFade{from{opacity:0}}@keyframes scnUp{from{transform:translateY(40px);opacity:0}}
.scn-busy{position:absolute;inset:0;z-index:40;display:grid;place-items:center;background:rgba(3,7,12,.82);backdrop-filter:blur(8px);animation:scnFade .2s}
.scn-busy-c{display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center}.scn-busy-c b{font-size:16px;margin-top:8px}.scn-busy-c small{opacity:.7;font-size:13px}
.scn-ring{position:relative;width:96px;height:96px}.scn-ring svg{width:100%;height:100%;transform:rotate(-90deg)}.scn-ring .r0{fill:none;stroke:rgba(255,255,255,.12);stroke-width:6}.scn-ring .r1{fill:none;stroke:var(--ac);stroke-width:6;stroke-linecap:round;stroke-dasharray:213.6;stroke-dashoffset:213.6;transition:stroke-dashoffset .3s}
.scn-ok{position:absolute;inset:0;display:grid;place-items:center;opacity:0;transform:scale(.4);transition:opacity .3s,transform .45s cubic-bezier(.3,1.6,.5,1)}.scn-ok svg{width:44px;height:44px;color:var(--ac)}
.scn-busy.done .scn-ok{opacity:1;transform:none}.scn-busy.done .scn-ring .r1{stroke-dashoffset:0!important}
.scn-toast{position:absolute;left:50%;bottom:max(130px,calc(env(safe-area-inset-bottom) + 120px));transform:translate(-50%,20px);z-index:50;padding:10px 16px;border-radius:12px;background:rgba(20,32,44,.95);border:1px solid rgba(255,255,255,.1);font-size:13.5px;font-weight:600;opacity:0;pointer-events:none;transition:opacity .25s,transform .25s;max-width:calc(100% - 32px);text-align:center}
.scn-toast.on{opacity:1;transform:translate(-50%,0)}
@media (prefers-reduced-motion:reduce){.scn-root *,.att *{animation:none!important;transition:none!important}}
`;document.head.appendChild(css);

if(window.I18N&&I18N.add)I18N.add({'المرفقات':'Attachments','ملف العقد والمرفقات':'Contract file & attachments','مرفقات المحضر':'Minutes attachments'});

/* One-time migration (admin): stamp a scope on files uploaded before scopes existed, from the record that
   references each one. Files no record references get an empty unit list (admin and uploader only). */
const SCOPE_COLS=['contracts','needs','batches','extensions','letters','companies','tickets','notices','records','meetings','assignments'];
function attIds(v,out){if(Array.isArray(v))v.forEach(x=>attIds(x,out));else if(v&&typeof v==='object'){for(const[k,x]of Object.entries(v)){if(k==='attachments'&&Array.isArray(x))x.forEach(a=>{if(a&&a.id)out.push(String(a.id))});else if(x&&typeof x==='object')attIds(x,out)}}return out}
function scopeFor(col,id,d){if(col==='records')return {k:'units',u:[d.unit,'gm','board'].filter(Boolean)};
  if(col==='meetings')return {k:'meeting',id};
  if(col==='assignments')return {k:'units',u:[d.fromUnit,d.toUnit,'gm','board'].filter(Boolean)};return {k:'all'}}
async function migrateScopes(force){
  const fs=FSX();if(!fs||ME().role!=='admin')return null;
  const flag=fs.collection('config').doc('filesScope');if(!force){const f=await flag.get();if(f.exists&&(f.data().v||0)>=1)return 0}
  const want=new Map();
  for(const col of SCOPE_COLS){let snap;try{snap=await fs.collection(col).get()}catch(e){continue}
    snap.forEach(doc=>{const d=doc.data()||{};const sc=scopeFor(col,doc.id,d);attIds(d,[]).forEach(fid=>{const cur=want.get(fid);if(!cur||sc.k==='all')want.set(fid,sc)})})}
  const metas=await fs.collection('files').where('n','>=',1).get();
  let batch=fs.batch(),k=0,n=0;
  for(const doc of metas.docs){if(doc.data().scope)continue;batch.update(doc.ref,{scope:want.get(doc.id)||{k:'units',u:[]}});n++;
    if(++k>=400){await batch.commit();batch=fs.batch();k=0}}
  if(k)await batch.commit();
  await flag.set({v:1,at:new Date().toISOString(),by:ME().email||'',count:n});
  return n;
}
(function autoMigrate(tries){setTimeout(()=>{if(ME().role==='admin'&&FSX()){migrateScopes().then(n=>{if(n)console.info('files: scoped '+n+' legacy attachments')}).catch(e=>console.warn('files scope migration',e&&e.message))}else if(tries<60)autoMigrate(tries+1)},tries?5000:15000)})(0);

window.ATT={field,view,val,scan:scanOpen,migrateScopes,
  formField:(key,list,opts)=>`<div class="field full">${field(key,list,Object.assign({reset:true},opts||{}))}</div>`,
  _test:{detectQuad,warp,applyFilter,makePDF,grayOf}};
})();
