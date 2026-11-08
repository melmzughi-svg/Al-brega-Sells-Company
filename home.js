/* ===== Home screen: app launcher (one tile per module), apps button + current-app crumb in the top bar =====
   Original icons drawn here (two-tone glyphs on rounded tiles); colours follow each module, the page follows the active palette/theme. */
(function(){
"use strict";
const en=()=>!!(window.I18N&&I18N.lang==='en');
const T=(a,e)=>en()?e:a;
const L=v=>(window.I18N&&I18N.t&&en())?I18N.t(v.label):v.label;

/* glyphs: 48×48, fill A = main colour, B = light accent, W = white */
const G={
 home:'<rect x="6" y="6" width="15" height="15" rx="4" fill="var(--a)"/><rect x="27" y="6" width="15" height="15" rx="4" fill="var(--b)"/><rect x="6" y="27" width="15" height="15" rx="4" fill="var(--b)"/><rect x="27" y="27" width="15" height="15" rx="4" fill="var(--a)"/>',
 dash:'<rect x="6" y="6" width="16" height="21" rx="4" fill="var(--a)"/><rect x="26" y="6" width="16" height="12" rx="4" fill="var(--b)"/><rect x="26" y="22" width="16" height="20" rx="4" fill="var(--a)"/><rect x="6" y="31" width="16" height="11" rx="4" fill="var(--b)"/>',
 ana:'<rect x="6" y="28" width="8" height="14" rx="3" fill="var(--b)"/><rect x="18" y="19" width="8" height="23" rx="3" fill="var(--a)"/><rect x="30" y="10" width="8" height="32" rx="3" fill="var(--b)"/><path d="M7 21 19 12l9 5 13-11" fill="none" stroke="var(--a)" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>',
 cs:'<path d="M9 27v-4a15 15 0 0 1 30 0v4" fill="none" stroke="var(--a)" stroke-width="5" stroke-linecap="round"/><rect x="5" y="24" width="10" height="15" rx="4" fill="var(--b)"/><rect x="33" y="24" width="10" height="15" rx="4" fill="var(--b)"/><path d="M38 39c0 4-5 5-11 5" fill="none" stroke="var(--a)" stroke-width="3.5" stroke-linecap="round"/><circle cx="25" cy="44" r="3" fill="var(--a)"/>',
 calendar:'<rect x="5" y="9" width="38" height="34" rx="7" fill="var(--b)"/><path d="M5 16a7 7 0 0 1 7-7h24a7 7 0 0 1 7 7v4H5z" fill="var(--a)"/><rect x="13" y="4" width="5" height="10" rx="2.5" fill="var(--a)"/><rect x="30" y="4" width="5" height="10" rx="2.5" fill="var(--a)"/><rect x="12" y="26" width="7" height="6" rx="2" fill="var(--a)"/><rect x="21" y="26" width="7" height="6" rx="2" fill="var(--a)" opacity=".55"/><rect x="30" y="26" width="7" height="6" rx="2" fill="var(--a)" opacity=".55"/><rect x="12" y="34" width="7" height="5" rx="2" fill="var(--a)" opacity=".55"/>',
 contracts:'<path d="M12 4h17l10 10v26a4 4 0 0 1-4 4H12a4 4 0 0 1-4-4V8a4 4 0 0 1 4-4z" fill="var(--b)"/><path d="M29 4v10h10z" fill="var(--a)"/><rect x="14" y="18" width="16" height="3.5" rx="1.75" fill="var(--a)"/><rect x="14" y="25" width="12" height="3.5" rx="1.75" fill="var(--a)" opacity=".6"/><circle cx="34" cy="36" r="9" fill="var(--a)"/><path d="m30 36 3 3 5-6" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>',
 ext:'<rect x="4" y="12" width="34" height="26" rx="5" fill="var(--b)"/><path d="m5 15 16 12 16-12" fill="none" stroke="var(--a)" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/><circle cx="37" cy="12" r="9" fill="var(--a)"/><path d="M33 12h8m-3-3 3 3-3 3" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>',
 needs:'<path d="M4 42V22l10 6v-6l10 6v-6l10 6V8h9v34z" fill="var(--a)"/><rect x="34" y="4" width="9" height="5" rx="2" fill="var(--b)"/><rect x="9" y="33" width="6" height="5" rx="1.5" fill="var(--b)"/><rect x="19" y="33" width="6" height="5" rx="1.5" fill="var(--b)"/><rect x="29" y="33" width="6" height="5" rx="1.5" fill="var(--b)"/>',
 batches:'<path d="m24 6 19 9-19 9L5 15z" fill="var(--a)"/><path d="m5 23 19 9 19-9" fill="none" stroke="var(--b)" stroke-width="4.5" stroke-linejoin="round" stroke-linecap="round"/><path d="m5 32 19 9 19-9" fill="none" stroke="var(--a)" stroke-width="4.5" stroke-linejoin="round" stroke-linecap="round" opacity=".7"/>',
 dir:'<rect x="6" y="5" width="32" height="38" rx="6" fill="var(--b)"/><rect x="38" y="11" width="5" height="8" rx="2" fill="var(--a)"/><rect x="38" y="23" width="5" height="8" rx="2" fill="var(--a)" opacity=".6"/><circle cx="22" cy="19" r="6.5" fill="var(--a)"/><path d="M11 36c1-6 6-9 11-9s10 3 11 9z" fill="var(--a)"/>',
 org:'<rect x="16" y="4" width="16" height="12" rx="4" fill="var(--a)"/><path d="M24 16v7M12 23h24v7M12 23v7" fill="none" stroke="var(--a)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/><rect x="4" y="30" width="16" height="12" rx="4" fill="var(--b)"/><rect x="28" y="30" width="16" height="12" rx="4" fill="var(--b)"/>',
 duties:'<rect x="8" y="8" width="32" height="36" rx="6" fill="var(--b)"/><rect x="16" y="4" width="16" height="9" rx="4" fill="var(--a)"/><path d="m15 25 4 4 7-8" fill="none" stroke="var(--a)" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round"/><rect x="28" y="23" width="7" height="3.5" rx="1.75" fill="var(--a)"/><path d="m15 36 4 4 7-8" fill="none" stroke="var(--a)" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round" opacity=".55"/><rect x="28" y="34" width="7" height="3.5" rx="1.75" fill="var(--a)" opacity=".55"/>',
 meetings:'<rect x="3" y="12" width="30" height="24" rx="6" fill="var(--a)"/><path d="m35 21 10-6v18l-10-6z" fill="var(--b)"/><circle cx="13" cy="22" r="4" fill="var(--b)"/><circle cx="23" cy="22" r="4" fill="var(--b)" opacity=".7"/><path d="M7 32c1-3 3-5 6-5s5 2 6 5m-2 0c1-3 3-5 6-5s5 2 6 5" fill="none" stroke="var(--b)" stroke-width="2.4"/>',
 settings:'<circle cx="24" cy="24" r="15" fill="none" stroke="var(--b)" stroke-width="8" stroke-dasharray="5.9 5.9"/><circle cx="24" cy="24" r="12" fill="var(--a)"/><circle cx="24" cy="24" r="5" fill="var(--b)"/>',
 sales:'<rect x="6" y="26" width="8" height="16" rx="2.5" fill="var(--b)"/><rect x="20" y="16" width="8" height="26" rx="2.5" fill="var(--a)"/><rect x="34" y="8" width="8" height="34" rx="2.5" fill="var(--b)"/><path d="m8 20 12-8 10 5 12-11" fill="none" stroke="var(--a)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>',
 transfers:'<rect x="4" y="8" width="16" height="14" rx="4" fill="var(--b)"/><rect x="28" y="26" width="16" height="14" rx="4" fill="var(--b)"/><path d="M22 15h14m-4-4 4 4-4 4M26 33H12m4 4-4-4 4-4" fill="none" stroke="var(--a)" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round"/>',
 receipts:'<path d="M6 16 24 7l18 9v18l-18 9-18-9z" fill="var(--b)"/><path d="M6 16l18 9 18-9M24 25v18" fill="none" stroke="var(--a)" stroke-width="3.4" stroke-linejoin="round"/><path d="m15 11 18 9" fill="none" stroke="var(--a)" stroke-width="3.4" stroke-linecap="round"/><circle cx="37" cy="37" r="8" fill="var(--a)"/><path d="m33.5 37 2.5 2.5 5-5" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>',
 debts:'<path d="M7 5h18l8 8v28a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2z" fill="var(--b)"/><path d="M25 5v8h8" fill="var(--a)" opacity=".35"/><path d="M12 34v-5M18 34v-9M24 34v-13" fill="none" stroke="var(--a)" stroke-width="3.6" stroke-linecap="round"/><circle cx="35" cy="35" r="10" fill="var(--a)"/><path d="M35 30v5.5l3.6 2.4" fill="none" stroke="#fff" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round"/>',
 margin:'<path d="M14 5c5.4 7 8.2 11.2 8.2 15.4a8.2 8.2 0 0 1-16.4 0C5.8 16.2 8.6 12 14 5z" fill="var(--b)"/><rect x="27" y="23" width="15" height="19" rx="2.5" fill="var(--b)" opacity=".75"/><rect x="27" y="15" width="15" height="7" rx="2" fill="var(--b)" opacity=".5"/><rect x="27" y="6" width="15" height="8" rx="2.5" fill="var(--a)"/><circle cx="14" cy="34" r="8.5" fill="var(--a)"/><path d="m10.6 37.4 6.8-6.8" stroke="#fff" stroke-width="2.3" stroke-linecap="round"/><circle cx="11.2" cy="31.2" r="1.7" fill="#fff"/><circle cx="16.8" cy="36.8" r="1.7" fill="#fff"/><path d="M4 45h40" stroke="var(--b)" stroke-width="2.4" stroke-linecap="round" opacity=".6"/>',
 sms055:'<path d="M6 34a18 18 0 0 1 36 0" fill="none" stroke="var(--b)" stroke-width="7" stroke-linecap="round"/><path d="M6 34a18 18 0 0 1 9-15.6" fill="none" stroke="var(--a)" stroke-width="7" stroke-linecap="round"/><path d="m24 34 9-11" stroke="var(--a)" stroke-width="4" stroke-linecap="round"/><circle cx="24" cy="34" r="5" fill="var(--a)"/><rect x="12" y="40" width="24" height="4" rx="2" fill="var(--b)" opacity=".8"/>'
};
/* tile colours: [main, accent] */
const C={home:['#2b7bd3','#9cc8f5'],dash:['#2563eb','#93c5fd'],ana:['#7c3aed','#c4b5fd'],cs:['#0d9488','#99f6e4'],calendar:['#e5484d','#fecaca'],
 contracts:['#0369a1','#7dd3fc'],ext:['#d97706','#fde68a'],needs:['#16a34a','#bbf7d0'],batches:['#db2777','#fbcfe8'],dir:['#0891b2','#a5f3fc'],
 org:['#4f46e5','#c7d2fe'],duties:['#059669','#a7f3d0'],meetings:['#e11d48','#fecdd3'],settings:['#475569','#cbd5e1'],sms055:['#ea580c','#fed7aa'],sales:['#6d3fd3','#ddd6fe'],transfers:['#0f8a63','#a7f3d0'],receipts:['#0e73b6','#bae6fd'],debts:['#9b1c31','#fda4af'],margin:['#a16207','#fde68a']};
const ORDER=['dash','ana','contracts','needs','batches','ext','dir','sms055','debts','margin','sales','transfers','receipts','cs','calendar','meetings','duties','org','settings'];
const BADGE=new Set(['cs','meetings']);
const SUB={dash:['مؤشرات الإدارة العامة','Management KPIs'],ana:['تقارير ورسوم بيانية','Reports & charts'],contracts:['سجل عقود البيع','Sales contracts'],
 needs:['احتياجات المنشآت من الوقود','Facility fuel needs'],batches:['دفعات الكشوفات','Needs batches'],ext:['تمديدات ومراسلات العقود','Extensions & letters'],
 dir:['الشركات والجهات','Companies & entities'],sms055:['المبيعات والتحصيلات','Sales & collections'],debts:['أرصدة الزبائن والتحصيلات','Customer balances & collections'],sales:['كميات المبيعات الشهرية','Monthly sales quantities'],transfers:['حركة المستودعات والمخازن','Warehouse & store movements'],receipts:['الاستلام من الموردين','Receipts from suppliers'],margin:['الربحية والأسعار','Profitability & prices'],cs:['الطلبات والشكاوى','Requests & complaints'],
 calendar:['المواعيد والاستحقاقات','Dates & deadlines'],meetings:['محاضر وقرارات','Minutes & decisions'],duties:['التكليفات والمتابعة','Assignments & follow-up'],
 org:['الإدارات والمناصب','Departments & positions'],settings:['المستخدمون والاستيراد','Users & import']};

function icon(id,size){const c=C[id]||C.home;const g=G[id]||G.home;
  return `<span class="hm-ic" style="--a:${c[0]};--b:${c[1]}${size?';--s:'+size+'px':''}" aria-hidden="true"><svg viewBox="0 0 48 48">${g}</svg></span>`}
window.__appIcon=icon;

function apps(){const ids=VIEWS.map(v=>v.id).filter(id=>id!=='home');
  const ordered=[...ORDER.filter(id=>ids.includes(id)),...ids.filter(id=>!ORDER.includes(id))];
  return ordered.map(id=>VIEWS.find(v=>v.id===id))}
function badge(v){if(!BADGE.has(v.id)||!v.count)return'';let n='';try{n=v.count()}catch(e){}return n?`<span class="hm-badge">${esc(String(n))}</span>`:''}
function greet(){const h=new Date().getHours();return h<12?T('صباح الخير','Good morning'):T('مساء الخير','Good evening')}
function myName(){try{const m=window.__me&&window.__me();return (m&&(m.name||m.email))||''}catch(e){return''}}
const HS={q:''};
function logo(){const i=document.querySelector('.side .brand-mark img');return i&&i.getAttribute('src')?`<span class="hm-logo" aria-hidden="true"><img src="${i.src}" alt=""></span>`:''}

function viewHome(){
  const list=apps();const q=norm(HS.q||'');
  const d=new Date();const ds=d.toLocaleDateString(en()?'en-GB':'ar-LY-u-nu-latn',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
  return `<section class="hm" aria-labelledby="hm-h">
    <div class="hm-head">
      <div class="hm-who">${logo()}<div><h2 id="hm-h">${esc(greet())}${myName()?(en()?', ':'، ')+esc(myName()):''}</h2><p>${esc(T('منظومة الإدارة العامة للمبيعات والشؤون التجارية','General Management of Sales & Commercial Affairs'))} · ${esc(ds)}</p></div></div>
      <label class="hm-find"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
        <input id="hm-q" type="search" autocomplete="off" placeholder="${esc(T('ابحث عن تطبيق…','Find an app…'))}" aria-label="${esc(T('ابحث عن تطبيق','Find an app'))}" value="${esc(HS.q||'')}"></label>
    </div>
    <div class="hm-grid" role="list">${list.map(v=>{const s=SUB[v.id];const hay=norm(v.label+' '+(s?s.join(' '):'')+' '+L(v));
      return `<button type="button" role="listitem" class="hm-app" data-hm="${esc(v.id)}"${q&&!hay.includes(q)?' hidden':''} title="${esc(L(v))}">${icon(v.id)}${badge(v)}<b>${esc(v.label)}</b>${s?`<small>${esc(T(s[0],s[1]))}</small>`:''}</button>`}).join('')}</div>
    <p class="hm-none" id="hm-none" hidden>${esc(T('لا يوجد تطبيق بهذا الاسم','No app matches'))}</p>
  </section>`}
function bindHome(){const i=document.getElementById('hm-q');if(!i)return;
  const f=()=>{HS.q=i.value;const q=norm(i.value);let n=0;document.querySelectorAll('.hm-app').forEach(b=>{const v=VIEWS.find(x=>x.id===b.dataset.hm);const s=SUB[b.dataset.hm];
    const ok=!q||norm(v.label+' '+(s?s.join(' '):'')+' '+L(v)).includes(q);b.hidden=!ok;if(ok)n++});document.getElementById('hm-none').hidden=n>0};
  i.addEventListener('input',f);i.addEventListener('keydown',e=>{if(e.key==='Enter'){const b=document.querySelector('.hm-app:not([hidden])');if(b){e.preventDefault();b.click()}}});f()}

/* ---- view registration (kept out of the sidebar list so Alt+1… stay on the existing screens) ---- */
VIEWS.unshift({id:'home',label:'التطبيقات',icon:'dash',group:'عام',nav:false});
I.apps=I.apps||'<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="5" r="2.2"/><circle cx="12" cy="5" r="2.2"/><circle cx="19" cy="5" r="2.2"/><circle cx="5" cy="12" r="2.2"/><circle cx="12" cy="12" r="2.2"/><circle cx="19" cy="12" r="2.2"/><circle cx="5" cy="19" r="2.2"/><circle cx="12" cy="19" r="2.2"/><circle cx="19" cy="19" r="2.2"/></svg>';
if(window.I18N&&I18N.add)I18N.add({'التطبيقات':'Apps','الرئيسية':'Home'});

/* top bar: apps button + crumb with the current app */
function mountTop(){const top=document.querySelector('header.top');if(!top||document.getElementById('apps-btn'))return;
  top.insertAdjacentHTML('afterbegin',`<button type="button" class="apps-btn" id="apps-btn" title="${esc(T('التطبيقات (الشاشة الرئيسية)','Apps (home screen)'))}" aria-label="${esc(T('التطبيقات','Apps'))}">${I.apps}</button><span class="app-crumb" id="app-crumb"></span>`);
  const brand=document.querySelector('.side .brand');if(brand&&!brand.dataset.hm){brand.dataset.hm='1';brand.setAttribute('role','link');brand.tabIndex=0;brand.style.cursor='pointer';brand.title=T('الشاشة الرئيسية','Home screen')}}
function updTop(){mountTop();const home=S.view==='home';document.body.classList.toggle('on-home',home);
  const a=document.getElementById('apps-btn');if(a)a.setAttribute('aria-current',String(home));
  const c=document.getElementById('app-crumb');if(c){const v=VIEWS.find(x=>x.id===S.view);c.hidden=home||!v;if(v&&!home)c.innerHTML=icon(v.id,22)+`<b>${esc(v.label)}</b>`}}
document.addEventListener('click',e=>{const t=e.target.closest&&e.target.closest('#apps-btn,.side .brand[data-hm],[data-hm]');if(!t)return;
  if(t.id==='apps-btn'||t.classList.contains('brand')){go('home');return}
  const id=t.dataset.hm;if(id&&VIEWS.some(v=>v.id===id))go(id)});
document.addEventListener('keydown',e=>{if((e.key==='Enter'||e.key===' ')&&e.target.matches&&e.target.matches('.side .brand[data-hm]')){e.preventDefault();go('home')}});

/* back button: hidden on the home screen; with no earlier page it returns home */
updBack=function(){const b=document.getElementById('back-btn');if(b)b.hidden=S.view==='home'};
goBack=function(){if(history.state&&history.state.d>0)history.back();else{go('home',null,true);try{history.replaceState({v:'home',d:0},'')}catch(e){}}};

const __render=render;
render=function(){
  if(S.view==='home'){recompute();document.getElementById('view').innerHTML=dbBanner()+viewHome();bindHome()}
  else{HS.q='';__render()}
  updTop();
};
/* first screen after opening the system is always the launcher (the app loaded the remembered screen before this file) */
S.view='home';try{localStorage.setItem('cm_view','home');history.replaceState({v:'home',d:0},'')}catch(e){}
mountTop();renderNav();render();
/* modules loaded after this file (calendar…) add their screens: draw the launcher again once every script has run */
document.addEventListener('DOMContentLoaded',()=>{if(S.view==='home')render()});
})();
