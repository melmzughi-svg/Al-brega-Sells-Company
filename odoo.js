/* ===== Module layout (business-suite style), applied on top of every screen after it renders =====
   - top navigation bar in the brand colour: apps button, app name, the app's menu (its sibling screens); the sidebar is retired
   - control panel: breadcrumb, action buttons, search/filter bar and status chips grouped in one sticky panel
   - list ⇄ cards (kanban) switch for the main record tables, built from the rendered table (cloned DOM, no new HTML from data)
   - record drawers open as a full form: breadcrumb, stage bar, content on a centred sheet
   Everything here re-arranges or decorates DOM the screens already rendered; click handling stays with the screens (delegated). */
(function(){
"use strict";
const en=()=>!!(window.I18N&&I18N.lang==='en');
const T=(a,e)=>en()?e:a;
const tl=s=>(window.I18N&&I18N.t&&en())?I18N.t(s):s;

/* apps = groups of screens; the first screen opens when the app's name is clicked */
const APPS=[
 {id:'rep',ar:'التقارير والمؤشرات',en:'Reports & KPIs',v:['dash','ana','sms055','sales','transfers','receipts','debts','margin']},
 {id:'sales',ar:'العقود والاحتياجات',en:'Contracts & needs',v:['contracts','needs','batches','ext','dir']},
 {id:'work',ar:'المهام والاجتماعات',en:'Tasks & meetings',v:['duties','meetings','calendar']},
 {id:'cs',ar:'خدمة الزبائن',en:'Customer service',v:['cs']},
 {id:'org',ar:'الهيكل التنظيمي',en:'Organization',v:['org']},
 {id:'settings',ar:'الإعدادات',en:'Settings',v:['settings']}];
function appOf(view){const a=APPS.find(a=>a.v.includes(view));if(a)return a;const v=VIEWS.find(x=>x.id===view);return {id:view,ar:v?v.label:view,en:v?tl(v.label):view,v:[view]}}
window.__appOf=appOf;
const icon=(id,s)=>window.__appIcon?window.__appIcon(id,s):'';

document.body.classList.add('o-ui');

/* ---------- navigation bar ---------- */
function navbar(){
  const top=document.querySelector('header.top');if(!top)return;
  const home=S.view==='home';top.classList.toggle('o-nav',!home);
  let m=document.getElementById('o-menu');
  if(!m){const c=document.getElementById('app-crumb');if(!c)return;c.insertAdjacentHTML('afterend','<nav class="o-menu" id="o-menu" aria-label="قائمة التطبيق"></nav>');m=document.getElementById('o-menu')}
  const crumb=document.getElementById('app-crumb');
  if(home){m.hidden=true;return}
  requestAnimationFrame(()=>{try{document.documentElement.style.setProperty('--o-navh',top.offsetHeight+'px')}catch(e){}});
  const a=appOf(S.view);const avail=a.v.filter(id=>VIEWS.some(x=>x.id===id));
  if(crumb){crumb.hidden=false;crumb.innerHTML=`<button type="button" class="o-appname" data-hm="${esc(avail[0]||S.view)}">${icon(S.view,24)}<b>${esc(T(a.ar,a.en))}</b></button>`}
  m.hidden=avail.length<2;
  m.innerHTML=avail.length<2?'':avail.map(id=>{const v=VIEWS.find(x=>x.id===id);let n='';try{n=v.count&&['cs','meetings'].includes(id)?v.count():''}catch(e){}
    return `<button type="button" data-hm="${esc(id)}" aria-current="${S.view===id||(id==='org'&&false)}">${esc(v.label)}${n?`<span class="o-mn">${esc(String(n))}</span>`:''}</button>`}).join('');
  const cur=m.querySelector('[aria-current="true"]');if(cur&&cur.scrollIntoView&&m.scrollWidth>m.clientWidth)try{cur.scrollIntoView({block:'nearest',inline:'nearest'})}catch(e){}
}

/* ---------- control panel ---------- */
const SUBKEY='oSub';
function controlPanel(){
  const v=document.getElementById('view');if(!v||S.view==='home')return;
  const ph=[...v.children].find(x=>x.classList&&x.classList.contains('page-head'));
  if(!ph||ph.closest('.o-cp'))return;
  const a=appOf(S.view);const vw=VIEWS.find(x=>x.id===S.view);
  const cp=document.createElement('div');cp.className='o-cp';ph.before(cp);cp.appendChild(ph);
  const h=ph.querySelector('h2');const p=ph.querySelector(':scope > div > p');
  if(h&&!h.previousElementSibling?.classList?.contains('o-bc')){
    const bc=document.createElement('nav');bc.className='o-bc';bc.setAttribute('aria-label',T('مسار التنقل','Breadcrumb'));
    const first=a.v.find(id=>VIEWS.some(x=>x.id===id))||S.view;
    bc.innerHTML=`<button type="button" data-hm="home">${esc(T('التطبيقات','Apps'))}</button><i aria-hidden="true">/</i>${a.v.length>1&&first!==S.view?`<button type="button" data-hm="${esc(first)}">${esc(T(a.ar,a.en))}</button><i aria-hidden="true">/</i>`:''}`;
    h.before(bc);
    if(p&&p.textContent.trim()){p.classList.add('o-sub');h.title=p.textContent.trim()}
  }
  /* search/filter row and status chips that follow the heading join the panel */
  let n=cp.nextElementSibling;let moved=0;
  while(n&&moved<3&&(n.classList.contains('toolbar')||n.classList.contains('chips')||n.classList.contains('an-chips')||n.classList.contains('tabs')||n.matches('.seg'))){
    const nx=n.nextElementSibling;n.classList.add('o-cp-row');cp.appendChild(n);n=nx;moved++}
  cp.classList.toggle('o-cp-has-rows',moved>0);
}

/* ---------- list ⇄ cards ---------- */
const KB_VIEWS=new Set(['contracts','needs','dir','batches','ext']);
const PILL_COL={'p-ok':'ok','p-warn':'warn','p-crit':'crit','p-info':'info','p-neutral':'neutral'};
function mode(){try{return localStorage.getItem('o_mode_'+S.view)||'list'}catch(e){return'list'}}
function setMode(m){try{localStorage.setItem('o_mode_'+S.view,m)}catch(e){}}
function mainTable(){const v=document.getElementById('view');if(!v)return null;
  return [...v.querySelectorAll('.tablewrap table')].find(t=>t.tBodies[0]&&t.tBodies[0].rows.length&&[...t.tBodies[0].rows].some(r=>[...r.attributes].some(at=>/^data-/.test(at.name))))||null}
function switcher(){
  const cp=document.querySelector('#view .o-cp');if(!cp||!KB_VIEWS.has(S.view))return;
  let sw=cp.querySelector('.o-vs');
  if(!sw){const ph=cp.querySelector('.page-head');if(!ph)return;
    sw=document.createElement('div');sw.className='o-vs';sw.setAttribute('role','group');sw.setAttribute('aria-label',T('طريقة العرض','View'));
    sw.innerHTML=`<button type="button" data-o-mode="list" title="${esc(T('عرض قائمة','List view'))}" aria-label="${esc(T('عرض قائمة','List view'))}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="3.5" cy="6" r="1.2" fill="currentColor"/><circle cx="3.5" cy="12" r="1.2" fill="currentColor"/><circle cx="3.5" cy="18" r="1.2" fill="currentColor"/></svg></button><button type="button" data-o-mode="kanban" title="${esc(T('عرض بطاقات','Card view'))}" aria-label="${esc(T('عرض بطاقات','Card view'))}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7.5" height="7.5" rx="1.8"/><rect x="13.5" y="3" width="7.5" height="7.5" rx="1.8"/><rect x="3" y="13.5" width="7.5" height="7.5" rx="1.8"/><rect x="13.5" y="13.5" width="7.5" height="7.5" rx="1.8"/></svg></button>`;
    (ph.querySelector(':scope > .toolbar')||ph).appendChild(sw)}
  const m=mode();sw.querySelectorAll('[data-o-mode]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.oMode===m)));
}
function txt(n){return (n&&n.textContent||'').replace(/\s+/g,' ').trim()}
function kanban(){
  const v=document.getElementById('view');if(!v)return;
  v.querySelectorAll('.o-kanban').forEach(k=>{if(!k.previousElementSibling||!k.previousElementSibling.classList.contains('tablewrap')||!KB_VIEWS.has(S.view)||mode()!=='kanban')k.remove()});
  v.querySelectorAll('.tablewrap.o-kb-hide').forEach(w=>{if(mode()!=='kanban'||!KB_VIEWS.has(S.view))w.classList.remove('o-kb-hide')});
  if(!KB_VIEWS.has(S.view)||mode()!=='kanban')return;
  const t=mainTable();if(!t)return;const wrap=t.closest('.tablewrap');
  const sig=t.tBodies[0].rows.length+'|'+txt(t.tBodies[0]).length;
  let k=wrap.nextElementSibling&&wrap.nextElementSibling.classList.contains('o-kanban')?wrap.nextElementSibling:null;
  if(k&&k.__t===t&&k.__sig===sig){wrap.classList.add('o-kb-hide');return}
  if(k)k.remove();
  const heads=[...t.tHead.rows[0].cells].map(c=>txt(c));
  k=document.createElement('div');k.className='o-kanban';k.setAttribute('role','list');k.__t=t;k.__sig=sig;
  [...t.tBodies[0].rows].forEach(tr=>{const cells=[...tr.cells];if(!cells.length)return;
    const card=document.createElement('div');card.className='o-card';card.tabIndex=0;card.setAttribute('role','listitem');
    const pillCell=cells.find(c=>c.querySelector('.pill'));const pill=pillCell&&pillCell.querySelector('.pill');
    const col=pill&&[...pill.classList].map(c=>PILL_COL[c]).find(Boolean);if(col)card.dataset.col=col;
    const head=document.createElement('div');head.className='o-card-h';
    const ref=document.createElement('span');ref.className='o-card-ref';cells[0].childNodes.forEach(n=>ref.appendChild(n.cloneNode(true)));head.appendChild(ref);
    if(pill)head.appendChild(pill.cloneNode(true));card.appendChild(head);
    if(cells[1]){const ttl=document.createElement('div');ttl.className='o-card-t';cells[1].childNodes.forEach(n=>ttl.appendChild(n.cloneNode(true)));card.appendChild(ttl)}
    const dl=document.createElement('dl');dl.className='o-card-f';let shown=0;
    cells.slice(2).forEach((c,i)=>{if(shown>=4||c===pillCell)return;const val=txt(c);if(!val||val==='—')return;
      const row=document.createElement('div');const dt=document.createElement('dt');dt.textContent=heads[i+2]||'';const dd=document.createElement('dd');c.childNodes.forEach(n=>dd.appendChild(n.cloneNode(true)));
      row.append(dt,dd);dl.appendChild(row);shown++});
    if(shown)card.appendChild(dl);
    card.addEventListener('click',e=>{if(e.target.closest('a,button'))return;tr.click()});
    card.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tr.click()}});
    k.appendChild(card)});
  wrap.after(k);wrap.classList.add('o-kb-hide');
}
document.addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-o-mode]');if(!b)return;setMode(b.dataset.oMode);switcher();kanban()});

/* ---------- record form (drawer) ---------- */
const STAGES=[
 ['لم يبدأ بعد','ساري','قريب الانتهاء','منتهي'],
 ['جديد','قيد المعالجة','بانتظار الزبون','مغلق']];
function form(){
  const d=document.querySelector('#drawer-root .drawer');if(!d||d.dataset.oForm)return;d.dataset.oForm='1';d.classList.add('o-form');
  const head=d.querySelector('.drawer-head'),body=d.querySelector('.drawer-body');if(!head||!body)return;
  const vw=VIEWS.find(x=>x.id===S.view);
  const bc=document.createElement('nav');bc.className='o-bc o-form-bc';bc.setAttribute('aria-label',T('مسار التنقل','Breadcrumb'));
  bc.innerHTML=`<button type="button" data-act="close-drawer">${esc(vw?vw.label:T('رجوع','Back'))}</button><i aria-hidden="true">/</i>`;
  const h3=head.querySelector('h3');if(h3)h3.before(bc);else head.prepend(bc);
  /* stage bar from the status pill in the header */
  const pills=[...head.querySelectorAll('.pill')].map(p=>txt(p));
  let st=null,stages=null;for(const s of STAGES){const hit=pills.find(p=>s.includes(p));if(hit){st=hit;stages=s;break}}
  if(stages){const i=stages.indexOf(st);const bar=document.createElement('div');bar.className='o-stages';bar.setAttribute('role','list');bar.setAttribute('aria-label',T('المرحلة','Stage'));
    bar.innerHTML=stages.map((s,j)=>`<span role="listitem" class="${j===i?'on':j<i?'done':''}"${j===i?' aria-current="step"':''}>${esc(s)}</span>`).join('');
    head.after(bar)}
  const sheet=document.createElement('div');sheet.className='o-sheet';while(body.firstChild)sheet.appendChild(body.firstChild);body.appendChild(sheet);
}
const dr=document.getElementById('drawer-root');
if(dr)new MutationObserver(()=>form()).observe(dr,{childList:true});

/* ---------- wiring ---------- */
function decorate(){try{navbar();controlPanel();switcher();kanban()}catch(e){console.warn('layout',e)}}
window.__oDecorate=decorate;
const __render=render;render=function(){__render();decorate()};
const __renderNav=renderNav;renderNav=function(){__renderNav();try{navbar()}catch(e){}};
/* partial redraws (filters, sorting, live data) replace the table only */
let q=0;const view=document.getElementById('view');
if(view)new MutationObserver(ms=>{if(ms.every(m=>m.target.closest&&m.target.closest('.o-kanban')))return;if(q)return;q=requestAnimationFrame(()=>{q=0;try{controlPanel();switcher();kanban()}catch(e){}})}).observe(view,{childList:true,subtree:true});
if(window.I18N&&I18N.add)I18N.add({'التقارير والمؤشرات':'Reports & KPIs','العقود والاحتياجات':'Contracts & needs','المهام والاجتماعات':'Tasks & meetings','قائمة التطبيق':'App menu'});
decorate();
document.addEventListener('DOMContentLoaded',decorate);
})();
