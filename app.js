/* Extensions and letters, settings, forms, actions, database wiring, global search, Excel.
   Classic script: shares the global scope with the other app files; load order is set in index.html. */
"use strict";
/* ============ Extensions & letters ============ */
function viewExt(){
  const tab=S.ui.et||'ext';
  const ext=[...S.extensions].sort((a,b)=>(b.decisionDate||'').localeCompare(a.decisionDate||''));
  const let_=[...S.letters].sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  const cByNo=no=>S.contracts.find(c=>normNo(c.no)===normNo(no));
  return `<div class="page-head"><div><h2>التمديدات والمراسلات</h2><p>يُسجَّل كل تمديد كقرار مستقل دون حذف السابق، ويُحسب الانتهاء الفعلي للعقد من آخر تمديد.</p></div>
  <div class="toolbar">${S.canWrite?`<button class="btn primary" data-act="${tab==='ext'?'add-ext':'add-letter'}">${I.plus}${tab==='ext'?'تمديد جديد':'مراسلة جديدة'}</button>`:''}</div></div>
  <div class="tabs" role="tablist"><button role="tab" data-et="ext" aria-selected="${tab==='ext'}">التمديدات (${ext.length})</button><button role="tab" data-et="let" aria-selected="${tab==='let'}">المراسلات (${let_.length})</button></div>
  ${tab==='ext'?(ext.length?`<div class="tablewrap"><table><thead><tr><th class="nosort">رقم العقد</th><th class="nosort">الشركة</th><th class="nosort">تاريخ قرار التمديد</th><th class="nosort">تاريخ الانتهاء الجديد</th><th class="nosort">ملاحظة</th></tr></thead><tbody>${ext.map(e=>{const c=cByNo(e.contractNo);return `<tr data-act="edit-ext" data-id="${e.id}"><td class="t-strong">${esc(e.contractNo)}</td><td>${c?esc(c.company):'<span class="tag dup">عقد غير موجود</span>'}</td><td class="n">${fd(e.decisionDate)}</td><td class="n">${fd(e.newEnd)}</td><td>${esc(e.note||'')}</td></tr>`}).join('')}</tbody></table></div>`:`<div class="panel empty"><b>لا توجد تمديدات مسجلة</b>عند صدور قرار تمديد، سجّله هنا أو من بطاقة العقد.</div>`)
   :(let_.length?`<div class="tablewrap"><table><thead><tr><th class="nosort">رقم العقد</th><th class="nosort">الشركة</th><th class="nosort">رقم المراسلة</th><th class="nosort">التاريخ</th><th class="nosort">الموضوع</th></tr></thead><tbody>${let_.map(e=>{const c=cByNo(e.contractNo);return `<tr data-act="edit-letter" data-id="${e.id}"><td class="t-strong">${esc(e.contractNo)}</td><td>${c?esc(c.company):'<span class="tag dup">عقد غير موجود</span>'}</td><td>${esc(e.ref||'')}</td><td class="n">${fd(e.date)}</td><td>${esc(e.note||'')}</td></tr>`}).join('')}</tbody></table></div>`:`<div class="panel empty"><b>لا توجد مراسلات مسجلة</b>سجّل الإخطارات والمخاطبات المتعلقة بكل عقد.</div>`)}`;
}
afterRender.ext=()=>{};
document.addEventListener('click',e=>{const t=e.target.closest('[data-et]');if(t){S.ui.et=t.dataset.et;render()}});

/* ============ Settings ============ */
/* ---- colour palettes (per device, localStorage cm_palette; default = Brega royal blue) ---- */
const PALETTES=[
 {id:'royal',name:'أزرق البريقة',desc:'ألوان شعار الشركة: أزرق مع لمسة حمراء',sw:['#0b3168','#0063ad','#c8161e','#f3f5f9'],meta:'#0b3168'},
 {id:'graphite',name:'الجرافيت المؤسسي',desc:'رمادي داكن محايد مع أزرق',sw:['#1b2029','#1f5fae','#5b9be6','#f4f5f7'],meta:'#1b2029'},
 {id:'gold',name:'الكحلي والذهبي',desc:'كحلي مع لمسات ذهبية وخلفية دافئة',sw:['#0e1b33','#0d4f8b','#d6a83a','#f5f4f0'],meta:'#0e1b33'}];
function applyPalette(id){const p=PALETTES.find(x=>x.id===id)||PALETTES[0];const r=document.documentElement;
  if(p.id==='royal')delete r.dataset.palette;else r.dataset.palette=p.id;
  const m=document.querySelector('meta[name="theme-color"]');if(m)m.setAttribute('content',p.meta);return p.id}
function currentPalette(){return document.documentElement.dataset.palette||'royal'}
function paletteChooser(){const cur=currentPalette();
  return `<div class="pal-opts" role="radiogroup" aria-label="نظام الألوان">${PALETTES.map(p=>`<button type="button" class="pal-opt" role="radio" aria-checked="${p.id===cur}" data-palette-pick="${p.id}"><span class="pal-sw" aria-hidden="true">${p.sw.map(c=>`<i style="background:${c}"></i>`).join('')}</span><b>${esc(p.name)}</b><small>${esc(p.desc)}</small></button>`).join('')}</div>`}
document.addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-palette-pick]');if(!b)return;
  const id=applyPalette(b.dataset.palettePick);try{localStorage.setItem('cm_palette',id)}catch(err){}
  document.querySelectorAll('[data-palette-pick]').forEach(x=>x.setAttribute('aria-checked',String(x.dataset.palettePick===id)))});
if(window.I18N&&I18N.add)I18N.add({'المظهر':'Appearance','نظام ألوان المنظومة على هذا الجهاز. يعمل مع المظهرين الفاتح والداكن.':'Colour scheme for this device. Works with both light and dark mode.','نظام الألوان':'Colour scheme',
  'أزرق البريقة':'Brega Royal Blue','ألوان شعار الشركة: أزرق مع لمسة حمراء':'The company logo colours: blue with a red accent','الجرافيت المؤسسي':'Corporate Graphite','رمادي داكن محايد مع أزرق':'Neutral charcoal with blue',
  'الكحلي والذهبي':'Navy & Gold','كحلي مع لمسات ذهبية وخلفية دافئة':'Navy with gold accents on a warm background'});
function viewSettings(){
  const s=S.settings;
  return `<div class="page-head"><div><h2>الإعدادات والاستيراد</h2><p>فترات التنبيه والقوائم المستخدمة في النماذج، واستيراد البيانات من Excel وتصدير نسخة كاملة.</p></div></div>
  <div class="grid2">
   <div class="panel"><h3 class="section-title">فترات التنبيه (أيام)</h3>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px">
     <div class="field"><label for="s-t1">تنبيه أول</label><input id="s-t1" type="number" min="1" value="${s.t1}"></div>
     <div class="field"><label for="s-t2">تنبيه ثانٍ</label><input id="s-t2" type="number" min="1" value="${s.t2}"></div>
     <div class="field"><label for="s-t3">تنبيه عاجل</label><input id="s-t3" type="number" min="1" value="${s.t3}"></div>
    </div>
    <h3 class="section-title" style="margin-top:16px">القوائم</h3>
    <div class="field"><label for="s-types">أنواع العقود (سطر لكل نوع)</label><textarea id="s-types" rows="5">${esc(s.types.join('\n'))}</textarea></div>
    <div class="field" style="margin-top:10px"><label for="s-depts">الإدارات</label><textarea id="s-depts" rows="4">${esc(s.depts.join('\n'))}</textarea></div>
    <div class="field" style="margin-top:10px"><label for="s-specials">الحالات الخاصة</label><textarea id="s-specials" rows="2">${esc(s.specials.join('\n'))}</textarea></div>
    ${S.canWrite?`<div style="margin-top:12px"><button class="btn primary" data-act="save-settings">حفظ الإعدادات</button></div>`:''}
   </div>
   <div class="panel" style="display:flex;flex-direction:column;gap:14px">
    <div><h3 class="section-title">المظهر</h3><p class="muted" style="margin:0 0 8px">نظام ألوان المنظومة على هذا الجهاز. يعمل مع المظهرين الفاتح والداكن.</p>${paletteChooser()}</div>
    <div><h3 class="section-title">تصدير</h3><p class="muted" style="margin:0 0 8px">ملف Excel واحد بكل الأوراق: العقود (مع الحالة المحسوبة)، الكشوفات، الدفعات، التمديدات، المراسلات.</p><button class="btn" data-act="export-all">${I.down}تصدير نسخة كاملة</button></div>
    ${S.canWrite?`<div><h3 class="section-title">استيراد من Excel</h3><p class="muted" style="margin:0 0 8px">يقبل ملف سجل العقود بنفس أعمدة ورقة «العقود» القديمة، أو ملفًا مُصدَّرًا من هذه المنظومة. السجلات الموجودة بنفس رقم العقد والشركة تُحدَّث ولا تتكرر.</p>
     <label class="btn">${I.up}اختيار ملف Excel<input type="file" id="imp-file" accept="${FileImport.ACCEPT}" hidden></label><div id="imp-res" style="margin-top:10px"></div></div>`:''}
    <div><h3 class="section-title">مصادر البيانات الأولية</h3><ul class="muted" style="margin:0;padding-inline-start:18px;font-size:13px">
     <li>سجل العقود (contracts-register.xlsx) — ${S.contracts.length} عقد.</li>
     <li>كشوفات الدفعات من وزارة الصناعة والمعادن لشهر أكتوبر 2026 — ${S.needs.length} منشأة في ${S.batches.length} دفعات.</li>
     <li>نسخ العقود الموقعة (PDF) — بيانات الطرف الثاني والبنود.</li></ul></div>
   </div>
  </div>`;
}
afterRender.settings=()=>{const f=document.getElementById('imp-file');if(f)f.addEventListener('change',importFile)};

/* ============ Forms ============ */
const dlg=document.getElementById('dlg');
function F(id,label,val,opts={}){const t=opts.type||'text';const req=opts.req?'<span class="req">*</span>':'';const full=opts.full?' full':'';
  if(t==='select')return `<div class="field${full}"><label for="${id}">${label}${req}</label><select id="${id}"><option value=""></option>${(opts.options||[]).map(o=>`<option ${o===val?'selected':''}>${esc(o)}</option>`).join('')}</select></div>`;
  if(t==='textarea')return `<div class="field${full}"><label for="${id}">${label}${req}</label><textarea id="${id}">${esc(val||'')}</textarea></div>`;
  return `<div class="field${full}"><label for="${id}">${label}${req}</label><input id="${id}" type="${t}" value="${esc(val??'')}" ${opts.list?`list="${opts.list}"`:''} ${t==='number'?'step="any"':''}>${opts.hint?`<span class="hint" id="${id}-hint">${opts.hint}</span>`:''}</div>`}
function dl(id,vals){return `<datalist id="${id}">${uniq(vals).map(v=>`<option value="${esc(v)}">`).join('')}</datalist>`}
function modal(title,body,foot){dlg.innerHTML=`<form method="dialog" id="mform" novalidate><div class="m-head"><h3>${esc(title)}</h3><button type="button" class="iconbtn" data-act="close-dlg" style="margin-inline-start:auto" aria-label="إغلاق">${I.close}</button></div><div class="m-body">${body}</div><div class="m-foot">${foot}</div></form>`;dlg.showModal();const first=dlg.querySelector('.m-body input,.m-body select');first&&first.focus()}
function val(id){const e=document.getElementById(id);return e?e.value.trim():''}
function contractForm(c){
  const isNew=!c;c=c||{currency:'دينار ليبي',dept:'إدارة المبيعات',party:{}};const p=c.party||{};const s=S.settings;
  modal(isNew?'عقد جديد':'تعديل العقد '+(c.no||''),`
   <fieldset><legend>بيانات أساسية</legend>
    ${F('f-no','رقم العقد',c.no||(isNew?'ع.ت.د/':''),{req:1,hint:'مثال: ع.ت.د/609'})}
    ${F('f-company','الشركة المتعاقدة',c.company,{req:1,list:'dl-comp'})}
    ${F('f-type','نوع العقد',c.type,{type:'select',options:uniq([...s.types,c.type])})}
    ${F('f-location','الموقع',c.location,{list:'dl-loc'})}
    ${F('f-subject','موضوع العقد',c.subject,{full:1})}
   </fieldset>
   <fieldset><legend>التواريخ</legend>
    ${F('f-sign','تاريخ التوقيع',c.signDate,{type:'date'})}${F('f-start','بداية السريان',c.startDate,{type:'date'})}${F('f-end','الانتهاء الأصلي',c.endDate,{type:'date'})}
    <div class="calc" id="f-calc"></div>
   </fieldset>
   <fieldset><legend>القيمة والكميات والمتابعة</legend>
    ${F('f-value','القيمة',c.value,{type:'number'})}${F('f-currency','العملة',c.currency,{type:'select',options:uniq([...s.currencies,c.currency])})}${F('f-qty','الكميات',c.qty,{hint:'مثال: 96468 لتر شهريًا'})}
    ${F('f-dept','الإدارة المختصة',c.dept,{type:'select',options:uniq([...s.depts,c.dept])})}${F('f-officer','مسؤول المتابعة',c.officer)}${F('f-special','حالة خاصة',c.special,{type:'select',options:s.specials})}
    ${F('f-pdf','مسار ملف PDF (اختياري)',c.pdfPath)}
    ${F('f-notes','ملاحظات وإجراءات',c.notes,{type:'textarea',full:1})}
   </fieldset>
   <fieldset><legend>ملف العقد والمرفقات</legend>${ATT.formField('dlg',c.attachments,{name:c.no?'عقد '+c.no:''})}</fieldset>
   <fieldset><legend>بيانات الطرف الثاني (من نسخة العقد)</legend>
    ${F('p-rep','الممثل القانوني / المفوض',p.rep)}${F('p-title','الصفة',p.repTitle)}${F('p-cr','رقم السجل التجاري',p.cr)}${F('p-phone','الهاتف',p.phone)}
    ${F('p-address','العنوان',p.address,{full:1})}${F('p-auth','سند التفويض / الهوية',p.auth,{full:1})}${F('p-purpose','الغرض من الاستعمال',p.purpose,{full:1})}${F('p-qtytext','الكمية كما وردت بالعقد',p.qtyText,{full:1})}${F('p-basis','أساس تحديد الكمية (تقرير اللجنة)',p.basis,{full:1})}
   </fieldset>
   ${dl('dl-comp',[...S.contracts.map(x=>x.company),...S.needs.map(x=>x.company)])}${dl('dl-loc',[...S.contracts.map(x=>x.location),...S.needs.map(x=>x.region)])}`,
  `<button type="submit" class="btn primary" data-act="save-contract" data-id="${c.id||''}">حفظ العقد</button><button type="button" class="btn" data-act="close-dlg">إلغاء</button>${!isNew?`<button type="button" class="btn danger" data-act="del" data-col="contracts" data-id="${c.id}" style="margin-inline-start:auto">حذف العقد</button>`:''}<span class="hint bad" id="f-err" style="margin-inline-start:${isNew?'auto':'8px'};color:var(--crit);font-size:12.5px"></span>`);
  const upd=()=>{const tmp={no:val('f-no'),startDate:val('f-start'),endDate:val('f-end'),special:val('f-special')};const d=derive(tmp);
    document.getElementById('f-calc').innerHTML=`<span>المدة: <b>${d.dur!=null?d.dur+' شهر':'—'}</b></span><span>الحالة المتوقعة: ${statusPill(d.status)}</span><span>المتبقي: <b>${d.remain!=null?d.remain+' يوم':'—'}</b></span>${d.exts?`<span>تمديدات مسجلة: <b>${d.exts}</b> (الانتهاء الفعلي ${fd(d.effEnd)})</span>`:''}`;
    const dup=S.contracts.find(x=>x.id!==c.id&&normNo(x.no)&&normNo(x.no)===normNo(tmp.no));const h=document.getElementById('f-no-hint');
    if(dup){h.textContent='تنبيه: هذا الرقم مستخدم لعقد '+dup.company;h.classList.add('bad')}else{h.textContent='مثال: ع.ت.د/609';h.classList.remove('bad')}};
  ['f-no','f-start','f-end','f-special'].forEach(i=>document.getElementById(i).addEventListener('input',upd));upd();
  document.getElementById('f-company').addEventListener('change',e=>{const n=S.needs.find(x=>norm(x.company)===norm(e.target.value));if(n){if(!val('f-location'))document.getElementById('f-location').value=n.region||'';if(!val('p-rep'))document.getElementById('p-rep').value=n.licensee||''}});
}
async function saveContract(id){
  const o={no:val('f-no'),company:val('f-company'),type:val('f-type'),location:val('f-location'),subject:val('f-subject'),signDate:val('f-sign'),startDate:val('f-start'),endDate:val('f-end'),value:num(val('f-value')),currency:val('f-currency'),qty:val('f-qty'),dept:val('f-dept'),officer:val('f-officer'),special:val('f-special'),pdfPath:val('f-pdf'),notes:val('f-notes'),
    party:{rep:val('p-rep'),repTitle:val('p-title'),cr:val('p-cr'),phone:val('p-phone'),address:val('p-address'),auth:val('p-auth'),purpose:val('p-purpose'),qtyText:val('p-qtytext'),basis:val('p-basis')}};
  const err=document.getElementById('f-err');
  if(!o.no||normNo(o.no)===normNo('ع.ت.د/')){err.textContent='أدخل رقم العقد.';document.getElementById('f-no').focus();return}
  if(!o.company){err.textContent='أدخل اسم الشركة المتعاقدة.';document.getElementById('f-company').focus();return}
  if(o.startDate&&o.endDate&&o.endDate<o.startDate){err.textContent='تاريخ الانتهاء يسبق تاريخ البداية.';return}
  o.attachments=await ATT.val('dlg','f-err');
  const old=id?S.contracts.find(x=>x.id===id):null;if(old){o.party=Object.assign({},old.party||{},o.party);['pdfAssetId','pdfName','pdfUrl','createdAt','source','docInfo'].forEach(k=>{if(old[k]!=null)o[k]=old[k]})}
  await write('contracts',id||uid('c'),o,!id);dlg.close();toast(id?'تم حفظ التعديلات':'تمت إضافة العقد');if(id)openContract(id);
}
function needForm(n){
  const isNew=!n;n=n||{batch:(S.ui.n&&S.ui.n.batch)||(batchesSorted().slice(-1)[0]||{}).id};
  modal(isNew?'منشأة جديدة في كشف الاحتياجات':'تعديل بيانات المنشأة',`
   <fieldset><legend>المنشأة</legend>
    <div class="field"><label for="n-batchsel">الدفعة<span class="req">*</span></label><select id="n-batchsel">${batchesSorted().map(b=>`<option value="${b.id}" ${b.id===n.batch?'selected':''}>${esc(b.name)}</option>`).join('')}</select></div>
    ${F('n-seq','التسلسل في الكشف',n.seq,{type:'number'})}
    ${F('n-company','اسم الشركة / المصنع',n.company,{req:1,full:1,list:'dl-comp'})}
    ${F('n-trade','الاسم التجاري — الشكل القانوني',n.tradeName)}${F('n-section','القسم في الكشف',n.section,{list:'dl-sec'})}${F('n-project','اسم المشروع (للمشروعات)',n.project,{full:1})}
    ${F('n-activity','نوع النشاط',n.activity,{list:'dl-act'})}${F('n-licensee','صاحب الترخيص أو المفوض',n.licensee)}${F('n-region','المنطقة',n.region,{list:'dl-reg'})}
    ${F('n-capacity','الطاقة الإنتاجية',n.capacity,{hint:'مثال: 10400م3/شهر'})}${F('n-power','الطاقة الكهربائية',n.power)}
   </fieldset>
   <fieldset><legend>الكميات المطلوبة للتشغيل (شهريًا)</legend>
    ${F('n-dg','ديزل مولدات (لتر)',n.dieselGen,{type:'number'})}${F('n-do','ديزل تشغيلية (لتر)',n.dieselOp,{type:'number'})}${F('n-gas','غاز (طن)',n.gas,{type:'number'})}${F('n-ho','زيت ثقيل (لتر)',n.heavyOil,{type:'number'})}${F('n-ke','كيروسين (لتر)',n.kerosene,{type:'number'})}${F('n-pq','كمية المشروع الشهرية (لتر)',n.projQty,{type:'number'})}${F('n-bq','الكمية حسب شركة البريقة',n.bregaQty,{type:'number'})}
    ${F('n-notes','ملاحظات',n.notes,{type:'textarea',full:1})}
   </fieldset>${dl('dl-comp',[...S.needs.map(x=>x.company),...S.contracts.map(x=>x.company)])}${dl('dl-act',S.needs.map(x=>x.activity))}${dl('dl-sec',S.needs.map(x=>x.section))}${dl('dl-reg',S.needs.map(x=>x.region))}`,
  `<button type="submit" class="btn primary" data-act="save-need" data-id="${n.id||''}">حفظ</button><button type="button" class="btn" data-act="close-dlg">إلغاء</button>${!isNew?`<button type="button" class="btn danger" data-act="del" data-col="needs" data-id="${n.id}" style="margin-inline-start:auto">حذف</button>`:''}<span id="f-err" style="color:var(--crit);font-size:12.5px;margin-inline-start:8px"></span>`);
}
async function saveNeed(id){
  const o={batch:val('n-batchsel'),seq:num(val('n-seq')),company:val('n-company'),activity:val('n-activity'),licensee:val('n-licensee'),region:val('n-region'),capacity:val('n-capacity'),power:val('n-power'),dieselGen:num(val('n-dg')),dieselOp:num(val('n-do')),gas:num(val('n-gas')),heavyOil:num(val('n-ho')),kerosene:num(val('n-ke')),projQty:num(val('n-pq')),bregaQty:num(val('n-bq')),tradeName:val('n-trade'),section:val('n-section'),project:val('n-project'),notes:val('n-notes')};
  if(!o.company){document.getElementById('f-err').textContent='أدخل اسم المنشأة.';return}
  const old=id?S.needs.find(x=>x.id===id):null;if(old){['page','extra','uncertain','createdAt','visitDate'].forEach(k=>{if(old[k]!=null)o[k]=old[k]})}
  await write('needs',id||uid('n'),o,!id);dlg.close();toast(id?'تم حفظ التعديلات':'تمت إضافة المنشأة');if(id)openNeed(id);
}
function batchForm(b){const isNew=!b;b=b||{order:(Math.max(0,...S.batches.map(x=>x.order||0))+1)};
  modal(isNew?'دفعة جديدة':'تعديل الدفعة',`<fieldset><legend>بيانات الدفعة</legend>${F('b-name','اسم الدفعة',b.name,{req:1,hint:'مثال: الدفعة السادسة'})}${F('b-order','الترتيب',b.order,{type:'number'})}${F('b-ref','الرقم الإشاري',b.ref)}${F('b-date','تاريخ الكتاب',b.date,{type:'date'})}${F('b-count','عدد المنشآت المعلن',b.count,{type:'number'})}${F('b-summary','ملخص الكتاب',b.summary,{type:'textarea',full:1})}</fieldset>`,
  `<button type="submit" class="btn primary" data-act="save-batch" data-id="${b.id||''}">حفظ</button><button type="button" class="btn" data-act="close-dlg">إلغاء</button>${!isNew?`<button type="button" class="btn danger" data-act="del" data-col="batches" data-id="${b.id}" style="margin-inline-start:auto">حذف الدفعة</button>`:''}<span id="f-err" style="color:var(--crit);font-size:12.5px;margin-inline-start:8px"></span>`)}
async function saveBatch(id){const o={name:val('b-name'),order:num(val('b-order')),ref:val('b-ref'),date:val('b-date'),count:num(val('b-count')),summary:val('b-summary')};if(!o.name){document.getElementById('f-err').textContent='أدخل اسم الدفعة.';return}
  const old=id?S.batches.find(x=>x.id===id):null;if(old&&old.pages)o.pages=old.pages;await write('batches',id||uid('b'),o,!id);dlg.close();toast('تم حفظ الدفعة')}
function extForm(e,no){const isNew=!e;e=e||{contractNo:no||''};
  modal(isNew?'تسجيل تمديد':'تعديل التمديد',`<fieldset><legend>قرار التمديد</legend>${F('e-no','رقم العقد',e.contractNo,{req:1,list:'dl-no'})}${F('e-date','تاريخ قرار التمديد',e.decisionDate,{type:'date'})}${F('e-end','تاريخ الانتهاء الجديد',e.newEnd,{type:'date',req:1})}${F('e-note','ملاحظة',e.note,{full:1})}</fieldset><fieldset><legend>المرفقات</legend>${ATT.formField('dlg',e.attachments,{name:'تمديد '+(e.contractNo||'')})}</fieldset>${dl('dl-no',S.contracts.map(c=>c.no))}`,
  `<button type="submit" class="btn primary" data-act="save-ext" data-id="${e.id||''}">حفظ</button><button type="button" class="btn" data-act="close-dlg">إلغاء</button>${!isNew?`<button type="button" class="btn danger" data-act="del" data-col="extensions" data-id="${e.id}" style="margin-inline-start:auto">حذف</button>`:''}<span id="f-err" style="color:var(--crit);font-size:12.5px;margin-inline-start:8px"></span>`)}
async function saveExt(id){const o={contractNo:val('e-no'),decisionDate:val('e-date'),newEnd:val('e-end'),note:val('e-note')};if(!o.contractNo||!o.newEnd){document.getElementById('f-err').textContent='رقم العقد وتاريخ الانتهاء الجديد مطلوبان.';return}
  o.attachments=await ATT.val('dlg','f-err');
  await write('extensions',id||uid('e'),o,!id);dlg.close();toast('تم حفظ التمديد');const c=S.contracts.find(x=>normNo(x.no)===normNo(o.contractNo));if(c&&document.querySelector('.drawer'))openContract(c.id)}
function letterForm(e,no){const isNew=!e;e=e||{contractNo:no||'',date:TODAY_ISO};
  modal(isNew?'تسجيل مراسلة':'تعديل المراسلة',`<fieldset><legend>المراسلة</legend>${F('l-no','رقم العقد',e.contractNo,{req:1,list:'dl-no'})}${F('l-ref','رقم المراسلة',e.ref)}${F('l-date','تاريخ المراسلة',e.date,{type:'date'})}${F('l-note','الموضوع / ملاحظة',e.note,{type:'textarea',full:1})}</fieldset><fieldset><legend>المرفقات</legend>${ATT.formField('dlg',e.attachments,{name:'مراسلة '+(e.ref||e.contractNo||'')})}</fieldset>${dl('dl-no',S.contracts.map(c=>c.no))}`,
  `<button type="submit" class="btn primary" data-act="save-letter" data-id="${e.id||''}">حفظ</button><button type="button" class="btn" data-act="close-dlg">إلغاء</button>${!isNew?`<button type="button" class="btn danger" data-act="del" data-col="letters" data-id="${e.id}" style="margin-inline-start:auto">حذف</button>`:''}<span id="f-err" style="color:var(--crit);font-size:12.5px;margin-inline-start:8px"></span>`)}
async function saveLetter(id){const o={contractNo:val('l-no'),ref:val('l-ref'),date:val('l-date'),note:val('l-note')};if(!o.contractNo){document.getElementById('f-err').textContent='أدخل رقم العقد.';return}
  o.attachments=await ATT.val('dlg','f-err');
  await write('letters',id||uid('l'),o,!id);dlg.close();toast('تم حفظ المراسلة');const c=S.contracts.find(x=>normNo(x.no)===normNo(o.contractNo));if(c&&document.querySelector('.drawer'))openContract(c.id)}
function quickAdd(){modal('إضافة سجل جديد',`<div class="batch-cards">
  <button type="button" class="bcard" data-act="add-contract" style="text-align:start;font:inherit;color:inherit">${I.contract.replace('<svg','<svg width="22" height="22"')}<h4>عقد بريقة</h4><span class="muted">عقد بيع وقود أو غاز مع شركة أو جهة</span></button>
  <button type="button" class="bcard" data-act="add-need" style="text-align:start;font:inherit;color:inherit">${I.factory.replace('<svg','<svg width="22" height="22"')}<h4>منشأة في كشف</h4><span class="muted">احتياجات مصنع من الوقود ضمن دفعة</span></button>
  <button type="button" class="bcard" data-act="add-ext" style="text-align:start;font:inherit;color:inherit">${I.ext.replace('<svg','<svg width="22" height="22"')}<h4>تمديد عقد</h4><span class="muted">قرار تمديد بتاريخ انتهاء جديد</span></button>
  <button type="button" class="bcard" data-act="add-letter" style="text-align:start;font:inherit;color:inherit">${I.file.replace('<svg','<svg width="22" height="22"')}<h4>مراسلة</h4><span class="muted">إخطار أو مخاطبة متعلقة بعقد</span></button>
  <button type="button" class="bcard" data-act="add-batch" style="text-align:start;font:inherit;color:inherit">${I.batch.replace('<svg','<svg width="22" height="22"')}<h4>دفعة وزارة</h4><span class="muted">كتاب إحالة دفعة جديدة من الكشوفات</span></button>
 </div>`,`<button type="button" class="btn" data-act="close-dlg">إغلاق</button>`)}

/* ============ Actions ============ */
document.addEventListener('click',async e=>{
  const a=e.target.closest('[data-act]');if(!a)return;const act=a.dataset.act,id=a.dataset.id;
  if(a.tagName==='INPUT')return;
  if(['save-contract','save-need','save-batch','save-ext','save-letter'].includes(act)){e.preventDefault();if(!S.canWrite)return;a.disabled=true;try{await({'save-contract':saveContract,'save-need':saveNeed,'save-batch':saveBatch,'save-ext':saveExt,'save-letter':saveLetter}[act])(id||null)}catch(err){const fe=document.getElementById('f-err');if(fe)fe.textContent=writeErr(err)}finally{a.disabled=false}return}
  switch(act){
   case 'close-drawer':closeDrawer();break;
   case 'close-dlg':dlg.close();break;
   case 'add-contract':if(dlg.open)dlg.close();contractForm();break;
   case 'edit-contract':contractForm(S.contracts.find(x=>x.id===id));break;
   case 'add-need':if(dlg.open)dlg.close();needForm();break;
   case 'edit-need':needForm(S.needs.find(x=>x.id===id));break;
   case 'add-batch':if(dlg.open)dlg.close();batchForm();break;
   case 'edit-batch':e.stopPropagation();batchForm(S.batches.find(x=>x.id===id));break;
   case 'batch-needs':e.stopPropagation();S.ui.n={q:'',batch:id,region:'',fuel:'',link:'',act:'',sort:'seq',dir:'asc'};go('needs');break;
   case 'add-ext':if(dlg.open)dlg.close();extForm(null,a.dataset.no);break;
   case 'edit-ext':if(S.canWrite)extForm(S.extensions.find(x=>x.id===id));break;
   case 'add-letter':if(dlg.open)dlg.close();letterForm(null,a.dataset.no);break;
   case 'edit-letter':if(S.canWrite)letterForm(S.letters.find(x=>x.id===id));break;
   case 'del':if(!a.classList.contains('armed')){a.classList.add('armed');a.textContent='اضغط مرة أخرى لتأكيد الحذف';setTimeout(()=>{a.classList.remove('armed');a.textContent='حذف'},4000);break}
     try{await S.db.collection(a.dataset.col).doc(id).delete();dlg.close();closeDrawer();toast('تم الحذف')}catch(err){toast(writeErr(err))}break;
   case 'save-settings':saveSettings();break;
   case 'export-contracts':exportContracts(filteredContracts());break;
   case 'export-needs':exportNeeds(filteredNeeds());break;
   case 'export-dir':exportDir();break;
   case 'export-org':exportOrg();break;
   case 'export-org-pdf':orgPdf();break;
   case 'export-all':exportAll();break;
   case 'do-import':doImport();break;
  }
});
document.addEventListener('click',e=>{const b=e.target.closest('[data-batch]');if(b&&!e.target.closest('button'))batchForm&&openBatchNeeds(b.dataset.batch)});
function openBatchNeeds(id){S.ui.n={q:'',batch:id,region:'',fuel:'',link:'',act:'',sort:'seq',dir:'asc'};go('needs')}
document.addEventListener('change',async e=>{const t=e.target;if(t.dataset&&t.dataset.act==='upload-pdf'&&t.files[0]){const id=t.dataset.id;const f=t.files[0];try{toast('جارٍ رفع الملف…');const r=await S.assets.upload(f);const c=S.contracts.find(x=>x.id===id);await write('contracts',id,Object.assign({},stripId(c),{pdfAssetId:r.id,pdfUrl:r.url,pdfName:f.name}),false);toast('تم إرفاق ملف العقد');openContract(id)}catch(err){toast('تعذّر رفع الملف: '+(err.message||err.code||''))}}});
document.getElementById('quick-add').addEventListener('click',()=>{if(!S.canWrite){toast('لا تملك صلاحية الإضافة');return}quickAdd()});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&document.querySelector('.drawer')&&!dlg.open)closeDrawer();if(e.key==='/'&&!/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)){e.preventDefault();document.getElementById('gq').focus()}});

/* ============ DB ============ */
function stripId(o){const r=Object.assign({},o);delete r.id;return r}
function writeErr(err){const c=err&&err.code;if(c==='invalid_argument')return 'لا تملك صلاحية الحفظ أو أن البيانات غير صالحة.';if(c==='quota_exceeded')return 'امتلأت سعة قاعدة البيانات.';if(c==='resource_exhausted')return 'عمليات كثيرة في وقت قصير، حاول بعد لحظات.';return 'تعذّر الحفظ، حاول مرة أخرى.'}
async function write(col,id,obj,isNew){
  if(!S.db)throw {code:'unavailable'};
  const now=new Date().toISOString();const data=JSON.parse(JSON.stringify(Object.assign({},obj,{updatedAt:now},isNew?{createdAt:now}:{})));
  Object.keys(data).forEach(k=>{if(data[k]===null||data[k]==='')delete data[k]});
  if(data.party){Object.keys(data.party).forEach(k=>{if(!data.party[k])delete data.party[k]})}
  await S.db.collection(col).doc(id).set(data);
}
async function saveSettings(){
  const s={t1:num(val('s-t1'))||90,t2:num(val('s-t2'))||60,t3:num(val('s-t3'))||30,types:val('s-types').split('\n').map(x=>x.trim()).filter(Boolean),depts:val('s-depts').split('\n').map(x=>x.trim()).filter(Boolean),specials:val('s-specials').split('\n').map(x=>x.trim()).filter(Boolean),currencies:S.settings.currencies};
  try{await S.db.doc('config/main').set(s);toast('تم حفظ الإعدادات')}catch(err){toast(writeErr(err))}
}
const rerender=debounce(()=>{renderNav();const ae=document.activeElement;const keep=ae&&ae.id&&/^(c-q|n-q|d-q|gq|tk-q|cu-q|nt-q|rc-q)$/.test(ae.id);
  if(keep&&S.view==='cs'){recompute();renderTk();renderCu();renderNt()}else if(keep&&S.view==='contracts'){recompute();renderContractsTable()}else if(keep&&S.view==='needs'){recompute();renderNeedsTable()}else if(keep&&S.view==='dir'){recompute();renderDir()}else if(keep&&S.view==='org'){const p=document.getElementById('rec-panel');if(p)renderRecTable(p.dataset.unit)}else if(!dlg.open)render()},120);
(async function init(){
  try{const t=localStorage.getItem('cm_theme');if(t)document.documentElement.dataset.theme=t;applyPalette(localStorage.getItem('cm_palette')||'royal');const v=localStorage.getItem('cm_view');if(v&&VIEWS.some(x=>x.id===v))S.view=v}catch(e){}
  try{history.replaceState({v:S.view,d:(history.state&&history.state.d)||0},'')}catch(e){}
  document.getElementById('today-line').textContent='اليوم: '+fd(TODAY_ISO);
  renderNav();render();
  if(!window.appRuntime||!window.appRuntime.use){setConn('none');return}
  const [db,user,assets,downloads]=await Promise.all([appRuntime.use('db'),appRuntime.use('user'),appRuntime.use('assets').catch(()=>null),appRuntime.use('downloads').catch(()=>null)]);
  S.user=user;S.assets=assets;S.downloads=downloads;
  if(!db){setConn('none');render();return}
  S.db=db;
  if(user){try{const w=await user.can('data.write');S.canWrite=w!==false}catch(e){}}
  let pending=COLS.length+1;const done=()=>{if(--pending===0){setConn('ready')}};
  COLS.forEach(col=>{let first=true;db.collection(col).onSnapshot(snap=>{S[col]=snap.docs.map(d=>Object.assign({id:d.id},d.data()));if(first){first=false;done()}rerender()},err=>{console.warn(col,err);if(first){first=false;done()}})});
  db.doc('config/notices').onSnapshot(snap=>{S.nm=snap.exists?snap.data():{data:{}};rerender()},()=>{});
  let f2=true;db.doc('config/main').onSnapshot(snap=>{if(snap.exists)S.settings=Object.assign({},S.settings,snap.data());if(f2){f2=false;done()}rerender()},()=>{if(f2){f2=false;done()}});
  setTimeout(()=>{if(S.dbState==='wait')setConn('ready')},8000);
})();
function setConn(s,ok){if(ok===true)s='ready';S.dbState=s;if(s==='ready')S.loadedOnce=true;const d=document.getElementById('conn-dot'),t=document.getElementById('conn-txt');
  if(s==='ready'){d.className='dot on';t.textContent='متصل · الحفظ تلقائي ومشترك';window.__syncReady=true;if(window.__sync)window.__sync.paint()}else if(s==='none'){d.className='dot off';t.textContent='غير متصل بقاعدة البيانات'}render()}
document.getElementById('theme-btn').addEventListener('click',()=>{const r=document.documentElement;const cur=r.dataset.theme||(matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light');const nx=cur==='dark'?'light':'dark';r.dataset.theme=nx;try{localStorage.setItem('cm_theme',nx)}catch(e){}});
async function loadArticles(){try{const r=await fetch('articles.json');if(r.ok){const j=await r.json();Object.assign(ARTICLES,j)}}catch(e){}}

/* ============ Global search ============ */
const gq=document.getElementById('gq'),gres=document.getElementById('gres');let gIdx=-1;
function gsearch(){const q=norm(gq.value),qn=normNo(gq.value);if(q.length<2){gres.hidden=true;return}
  const cs=S.contracts.filter(c=>norm([c.no,c.company,c.location,c.type,c.notes,c.party&&c.party.rep].join(' ')).includes(q)||normNo(c.no).includes(qn)).slice(0,8);
  const ns=S.needs.filter(n=>norm([n.company,n.licensee,n.region,n.activity].join(' ')).includes(q)).slice(0,8);
  const bs=S.batches.filter(b=>norm([b.name,b.ref,b.summary].join(' ')).includes(q)).slice(0,4);
  let h='';if(cs.length)h+='<h6>عقود البريقة</h6>'+cs.map(c=>`<button data-open-c="${c.id}"><b>${esc(c.company)}</b><span>${esc(c.no)} · ${esc(D.get(c.id)?.status||'')}</span></button>`).join('');
  if(ns.length)h+='<h6>كشوفات الوزارة</h6>'+ns.map(n=>`<button data-open-n="${n.id}"><b>${esc(n.company)}</b><span>${esc(batchName(n.batch))} · ${esc(n.region||'')}</span></button>`).join('');
  if(bs.length)h+='<h6>الدفعات</h6>'+bs.map(b=>`<button data-gb="${b.id}"><b>${esc(b.name)}</b><span>${esc(b.ref||'')}</span></button>`).join('');
  gres.innerHTML=h||'<div class="empty" style="padding:18px">لا نتائج لـ «'+esc(gq.value)+'»</div>';gres.hidden=false;gIdx=-1}
gq.addEventListener('input',debounce(gsearch,100));gq.addEventListener('focus',gsearch);
gq.addEventListener('keydown',e=>{const items=[...gres.querySelectorAll('button')];if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();gIdx=(gIdx+(e.key==='ArrowDown'?1:-1)+items.length)%items.length;items.forEach((b,i)=>b.classList.toggle('act',i===gIdx))}else if(e.key==='Enter'&&items.length){(items[gIdx]||items[0]).click()}else if(e.key==='Escape'){gres.hidden=true;gq.blur()}});
document.addEventListener('click',e=>{if(!e.target.closest('#gsearch'))gres.hidden=true;else if(e.target.closest('#gres button')){const b=e.target.closest('#gres button');if(b.dataset.gb)openBatchNeeds(b.dataset.gb);gres.hidden=true}});

/* ============ Excel export / import ============ */
async function saveXlsx(sheets,filename){
  if(!window.XLSX){toast('مكتبة Excel لم تُحمّل بعد');return}
  const wb=XLSX.utils.book_new();wb.Workbook={Views:[{RTL:true}]};
  sheets.forEach(([name,rows])=>{const ws=XLSX.utils.json_to_sheet(rows.length?rows:[{'':'لا توجد بيانات'}]);const keys=Object.keys(rows[0]||{});ws['!cols']=keys.map(k=>({wch:Math.min(45,Math.max(10,k.length+2,...rows.slice(0,200).map(r=>String(r[k]??'').length)))}));XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31))});
  const data=XLSX.write(wb,{type:'array',bookType:'xlsx'});
  if(!S.downloads){toast('التنزيل غير متاح في هذا العرض');return}
  try{await S.downloads.save({filename,data:new Uint8Array(data)});toast('تم تجهيز الملف')}catch(err){if(err.code!=='declined')toast('تعذّر تنزيل الملف')}
}
function contractRows(rows){return rows.map(({c,d},i)=>({'م':i+1,'رقم العقد':c.no,'الموقع':c.location||'','الشركة المتعاقدة':c.company,'نوع العقد':c.type||'','موضوع العقد':c.subject||'','تاريخ التوقيع':c.signDate||'','بداية السريان':c.startDate||'','الانتهاء الأصلي':c.endDate||'','الانتهاء الفعلي (بعد التمديد)':d.effEnd||'','المدة الأصلية (شهر)':d.dur??'','القيمة':c.value??'','العملة':c.currency||'','الكميات':c.qty||'','الإدارة المختصة':c.dept||'','مسؤول المتابعة':c.officer||'','حالة خاصة':c.special||'','الحالة':d.status,'المتبقي (يوم)':d.remain??'','التنبيه':d.alert,'عدد التمديدات':d.exts,'مسار ملف PDF':c.pdfPath||'','ملاحظات وإجراءات':c.notes||'','الممثل القانوني':(c.party||{}).rep||'','السجل التجاري':(c.party||{}).cr||'','الهاتف':(c.party||{}).phone||'','العنوان':(c.party||{}).address||''}))}
function needRows(rows){return rows.map(n=>({'الدفعة':batchName(n.batch),'ت':n.seq??'','اسم الشركة / المصنع':n.company,'نوع النشاط':n.activity||'','صاحب الترخيص او المفوض':n.licensee||'','المنطقة':n.region||'','الطاقة الانتاجية':n.capacity||'','الطاقة الكهربائية':n.power||'','ديزل/لتر مولدات':n.dieselGen??'','ديزل/لتر تشغيلية':n.dieselOp??'','غاز/طن':n.gas??'','زيت ثقيل/لتر':n.heavyOil??'','كيروسين/لتر':n.kerosene??'','كمية المشروع/لتر':n.projQty??'','الكمية حسب شركة البريقة':n.bregaQty??'','القسم':n.section||'','الاسم التجاري':n.tradeName||'','اسم المشروع':n.project||'','تاريخ الزيارة':n.visitDate||'','ملاحظات':n.notes||'','عقد البريقة المطابق':(MATCH.byNeed.get(n.id)||{}).no||''}))}
function exportContracts(rows){saveXlsx([['العقود',contractRows(rows)]],'عقود-البريقة-'+TODAY_ISO+'.xlsx')}
function exportNeeds(rows){saveXlsx([['الكشوفات',needRows(rows)]],'كشوفات-الاحتياجات-'+TODAY_ISO+'.xlsx')}
function exportDir(){const r=(S._dirRows||buildDir()).map(e=>({'الجهة':e.name,'المنطقة':e.region||'','النشاط':e.activity||'','المفوض / الممثل':e.person||'','الهاتف':e.phone||'','العقود':e.contracts.map(c=>c.no).join('، '),'الدفعات':e.needs.map(n=>batchName(n.batch)).join('، '),'ديزل معتمد':sumBy(e.needs,n=>(n.dieselGen||0)+(n.dieselOp||0))}));saveXlsx([['دليل الجهات',r]],'دليل-الجهات-'+TODAY_ISO+'.xlsx')}
function exportAll(){recompute();const cs=S.contracts.map(c=>({c,d:D.get(c.id)})).sort((a,b)=>(a.c.no||'').localeCompare(b.c.no||''));
  saveXlsx([['العقود',contractRows(cs)],['الكشوفات',needRows(filteredNeedsAll())],['الدفعات',batchesSorted().map(b=>({'الدفعة':b.name,'الترتيب':b.order,'الرقم الإشاري':b.ref||'','التاريخ':b.date||'','العدد المعلن':b.count??'','المسجل':S.needs.filter(n=>n.batch===b.id).length,'ملخص':b.summary||''}))],['التمديدات',S.extensions.map(e=>({'رقم العقد':e.contractNo,'تاريخ قرار التمديد':e.decisionDate||'','تاريخ الانتهاء الجديد':e.newEnd||'','ملاحظة':e.note||''}))],['المراسلات',S.letters.map(e=>({'رقم العقد':e.contractNo,'رقم المراسلة':e.ref||'','تاريخ المراسلة':e.date||'','ملاحظة':e.note||''}))]],'نسخة-المنظومة-'+TODAY_ISO+'.xlsx')}
function filteredNeedsAll(){const bo=id=>{const b=S.batches.find(x=>x.id===id);return b?b.order||0:999};return [...S.needs].sort((a,b)=>bo(a.batch)-bo(b.batch)||(a.seq||0)-(b.seq||0))}
let IMP=null;
const CMAP={'رقم العقد':'no','الموقع':'location','الشركة المتعاقدة':'company','نوع العقد':'type','موضوع العقد':'subject','تاريخ التوقيع':'signDate','بداية السريان':'startDate','الانتهاء الأصلي':'endDate','القيمة':'value','العملة':'currency','الكميات':'qty','الإدارة المختصة':'dept','مسؤول المتابعة':'officer','حالة خاصة':'special','مسار ملف PDF':'pdfPath','ملاحظات وإجراءات':'notes'};
const NMAP={'ت':'seq','اسم الشركة / المصنع':'company','نوع النشاط':'activity','صاحب الترخيص او المفوض':'licensee','المنطقة':'region','الطاقة الانتاجية':'capacity','الطاقة الكهربائية':'power','ديزل/لتر مولدات':'dieselGen','ديزل/لتر تشغيلية':'dieselOp','غاز/طن':'gas','زيت ثقيل/لتر':'heavyOil','كيروسين/لتر':'kerosene','كمية المشروع/لتر':'projQty','الكمية حسب شركة البريقة':'bregaQty','القسم':'section','الاسم التجاري':'tradeName','اسم المشروع':'project','تاريخ الزيارة':'visitDate','ملاحظات':'notes'};
function xdate(v){if(v==null||v==='')return'';if(v instanceof Date)return iso(v);if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400000));return iso(new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate()))}const m=String(v).match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);if(m)return m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');const m2=String(v).match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);if(m2)return m2[3]+'-'+m2[2].padStart(2,'0')+'-'+m2[1].padStart(2,'0');return''}
async function importFile(e){const f=e.target.files[0];if(!f)return;const res=document.getElementById('imp-res');
  try{const wb=await readWorkbook(f);
    const cSheet=wb.SheetNames.find(n=>n.includes('العقود'))||null;const nSheet=wb.SheetNames.find(n=>n.includes('الكشوفات')||n.includes('الاحتياجات'));
    const toRows=n=>wb.rows(n);
    let contracts=[],needs=[];
    const first=wb.SheetNames[0];const fr=toRows(cSheet||first);
    if(fr.length&&('رقم العقد' in fr[0]))contracts=fr.map(r=>{const o={};for(const[k,v]of Object.entries(CMAP))if(r[k]!==''&&r[k]!=null)o[v]=r[k];['signDate','startDate','endDate'].forEach(k=>{if(o[k])o[k]=xdate(o[k])});if(o.value!=null)o.value=num(o.value);Object.keys(o).forEach(k=>{if(typeof o[k]==='string')o[k]=o[k].trim()});return o}).filter(o=>o.company);
    if(nSheet){needs=toRows(nSheet).map(r=>{const o={};for(const[k,v]of Object.entries(NMAP))if(r[k]!==''&&r[k]!=null)o[v]=r[k];['seq','dieselGen','dieselOp','gas','heavyOil','kerosene','projQty','bregaQty'].forEach(k=>{if(o[k]!=null)o[k]=num(o[k])});o._batch=String(r['الدفعة']||'').trim();return o}).filter(o=>o.company)}
    let upd=0,add=0;contracts.forEach(o=>{const ex=S.contracts.find(c=>normNo(c.no)===normNo(o.no)&&norm(c.company)===norm(o.company));o._id=ex?ex.id:null;ex?upd++:add++});
    let nadd=0,nupd=0;needs.forEach(o=>{const b=S.batches.find(b=>norm(b.name)===norm(o._batch));o.batch=b?b.id:null;o._newBatch=b?null:o._batch;const ex=S.needs.find(n=>n.batch===o.batch&&norm(n.company)===norm(o.company));o._id=ex?ex.id:null;ex?nupd++:nadd++});
    IMP={contracts,needs};
    res.innerHTML=(contracts.length||needs.length)?`<div class="banner info">تمت قراءة الملف: ${contracts.length?`${contracts.length} عقد (${add} جديد، ${upd} تحديث)`:''}${contracts.length&&needs.length?' و':''}${needs.length?`${needs.length} منشأة (${nadd} جديدة، ${nupd} تحديث)`:''}.<div style="margin-top:8px"><button class="btn primary sm" data-act="do-import">تنفيذ الاستيراد</button></div></div>`:`<div class="banner">لم أجد ورقة بعمود «رقم العقد» أو ورقة «الكشوفات». تأكد من أن الصف الأول يحوي عناوين الأعمدة.</div>`;
  }catch(err){res.innerHTML='<div class="banner">'+esc(err.message||'تعذّرت قراءة الملف.')+'</div>'}
  e.target.value='';
}
async function doImport(){if(!IMP)return;const res=document.getElementById('imp-res');let n=0;
  try{for(const o of IMP.contracts){const id=o._id;delete o._id;const old=id?S.contracts.find(c=>c.id===id):null;await write('contracts',id||uid('c'),Object.assign({},old?stripId(old):{},o),!id);n++}
    const newB={};for(const o of IMP.needs){if(!o.batch&&o._newBatch){if(!newB[o._newBatch]){const bid=uid('b');await write('batches',bid,{name:o._newBatch,order:Math.max(0,...S.batches.map(x=>x.order||0))+1+Object.keys(newB).length},true);newB[o._newBatch]=bid}o.batch=newB[o._newBatch]}
      const id=o._id;['_id','_batch','_newBatch'].forEach(k=>delete o[k]);const old=id?S.needs.find(x=>x.id===id):null;await write('needs',id||uid('n'),Object.assign({},old?stripId(old):{},o),!id);n++}
    res.innerHTML=`<div class="banner info">تم استيراد ${n} سجل بنجاح.</div>`;IMP=null;toast('تم استيراد '+n+' سجل بنجاح')}
  catch(err){res.innerHTML=`<div class="banner">توقف الاستيراد بعد ${n} سجل: ${writeErr(err)}</div>`}}
