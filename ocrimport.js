/* ===== قراءة الملفات الممسوحة ضوئيًا والصور (OCRImport) =====
   Used by the shared reader (fileimport.js) for a PDF that has no text layer (a scan) and for a picture of a report
   (.jpg .png .webp .bmp, a phone photo or a scanner image). Everything runs in the browser: no file leaves the device.
     1. every page is drawn on a canvas at about 300 dpi (pdf.js for a PDF), turned to grey and straightened (deskew);
     2. it is made black on white with a local threshold, and the long horizontal and vertical lines of the table are
        found and erased (they confuse the text recogniser); the vertical ones are kept as the column borders;
     3. Tesseract (vendor/tesseract, Arabic + English, loaded only when needed) reads the words with a confidence for each;
     4. the words are put back in rows and columns, numbers written with Arabic-Indic digits or thousands separators
        become numbers, and every cell keeps the lowest confidence of its words;
     5. a review window that cannot be skipped shows the table next to the picture: low-confidence cells are coloured,
        each one has to be checked (or corrected) before the import continues, and the user confirms that the figures
        match the original. Only then the rows go on to the dashboard's own import (and its own preview, where it has one).
   OCRImport.isImage(u8) · OCRImport.readImage(file) · OCRImport.readScannedPdf(doc, name) → [{name, grid, text:true, ocr}]
   A cancelled review throws an error with .cancelled = true. */
"use strict";
(function(){
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const H=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fail=(ar,en,extra)=>{const e=new Error(tr(ar,en));e.importError=true;if(extra)Object.assign(e,extra);return e};
const MAX_PAGES=20;
/* a number cell under this confidence (0–100) has to be checked by hand; text cells have a lower bar */
const NUM_OK=90,TEXT_OK=80;
const AR=/[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;

/* ---------- loading Tesseract on demand ---------- */
let libP=null;const workers={};
function lib(){if(window.Tesseract)return Promise.resolve(window.Tesseract);if(!libP)libP=new Promise((res,rej)=>{const s=document.createElement('script');s.src='vendor/tesseract/tesseract.min.js';s.integrity='sha384-r1ru3tcf6FhnCFR4B7pIFG+BhFF9LlFtz/P1y4pblWn3AGs9y3lBx5SKLNf4+rED';s.crossOrigin='anonymous';
  s.onload=()=>window.Tesseract?res(window.Tesseract):rej(new Error('tesseract'));s.onerror=()=>{libP=null;s.remove();rej(new Error('tesseract'))};document.head.appendChild(s)});return libP}
/* two readers: Arabic + English for every cell, and a digits-only one that reads every figure a second time */
const KINDS={text:['ara+eng',{tessedit_pageseg_mode:'7',preserve_interword_spaces:'1',user_defined_dpi:'300'}],
  block:['ara+eng',{tessedit_pageseg_mode:'6',preserve_interword_spaces:'1',user_defined_dpi:'300'}],
  raw:['ara+eng',{tessedit_pageseg_mode:'13',preserve_interword_spaces:'1',user_defined_dpi:'300'}],
  page:['ara+eng',{tessedit_pageseg_mode:'6',preserve_interword_spaces:'1',user_defined_dpi:'300'}],
  digits:['eng',{tessedit_pageseg_mode:'7',tessedit_char_whitelist:'0123456789.,-()/%',user_defined_dpi:'300'}]};
function worker(kind){const k=kind==='page'||kind==='raw'||kind==='block'?'text':kind;if(!workers[k])workers[k]=(async()=>{const T=await lib();
  const w=await T.createWorker(KINDS[k][0],1,{workerPath:'vendor/tesseract/worker.min.js',corePath:'vendor/tesseract/core',langPath:'vendor/tesseract/lang',gzip:true,workerBlobURL:false,errorHandler:e=>console.warn('OCR',e)});
  w._mode=null;return w})().catch(e=>{delete workers[k];throw e});
  return workers[k].then(async w=>{if(w._mode!==kind){await w.setParameters(KINDS[kind][1]);w._mode=kind}return w})}
/* free the memory the readers take once the user is done with imports for a while */
let idle=null;const release=()=>{clearTimeout(idle);idle=setTimeout(()=>{Object.keys(workers).forEach(async k=>{const p=workers[k];delete workers[k];try{(await p).terminate()}catch(e){}})},120000)};

/* ---------- picture → clean black-on-white page ---------- */
const IMG_MAGIC=[[0x89,0x50,0x4E,0x47],[0xFF,0xD8,0xFF],[0x42,0x4D],[0x47,0x49,0x46,0x38],[0x49,0x49,0x2A,0x00],[0x4D,0x4D,0x00,0x2A]];
const isWebp=u8=>u8[0]===0x52&&u8[1]===0x49&&u8[2]===0x46&&u8[3]===0x46&&u8[8]===0x57&&u8[9]===0x45&&u8[10]===0x42&&u8[11]===0x50;
const isTiff=u8=>IMG_MAGIC.slice(4).some(m=>m.every((x,i)=>u8[i]===x));
const isImage=u8=>isWebp(u8)||IMG_MAGIC.some(m=>m.every((x,i)=>u8[i]===x));
async function bitmapOf(file){
  if(window.createImageBitmap){try{return await createImageBitmap(file,{imageOrientation:'from-image'})}catch(e){}}
  const url=URL.createObjectURL(file);try{const im=new Image();im.src=url;await im.decode();return im}catch(e){return null}finally{setTimeout(()=>URL.revokeObjectURL(url),1000)}}
/* draw at a size the recogniser likes: about 2400 px across for a page, small photos are enlarged */
function canvasOf(src,w,h){let s=1;const big=Math.max(w,h);if(big<1600)s=Math.min(3,2400/big);else if(big>4200)s=4200/big;
  const c=document.createElement('canvas');c.width=Math.round(w*s);c.height=Math.round(h*s);const x=c.getContext('2d',{willReadFrequently:true});
  x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.imageSmoothingQuality='high';x.drawImage(src,0,0,c.width,c.height);return c}
const greyOf=c=>{const d=c.getContext('2d',{willReadFrequently:true}).getImageData(0,0,c.width,c.height).data;const g=new Uint8ClampedArray(c.width*c.height);
  for(let i=0,j=0;j<g.length;i+=4,j++)g[j]=(d[i]*299+d[i+1]*587+d[i+2]*114)/1000;return g};
/* local (Bradley) threshold on a lightly smoothed copy: copes with shadows and uneven light in phone photos */
function binarize(g,w,h){const sm=new Uint8ClampedArray(g.length);
  for(let y=0;y<h;y++)for(let x=0;x<w;x++){let s=0,n=0;for(let dy=-1;dy<=1;dy++){const yy=y+dy;if(yy<0||yy>=h)continue;for(let dx=-1;dx<=1;dx++){const xx=x+dx;if(xx<0||xx>=w)continue;s+=g[yy*w+xx];n++}}sm[y*w+x]=s/n}
  const I=new Float64Array((w+1)*(h+1));for(let y=0;y<h;y++){let r=0;for(let x=0;x<w;x++){r+=sm[y*w+x];I[(y+1)*(w+1)+x+1]=I[y*(w+1)+x+1]+r}}
  const R=Math.max(8,Math.round(Math.max(w,h)/60)),b=new Uint8Array(g.length);
  for(let y=0;y<h;y++){const y0=Math.max(0,y-R),y1=Math.min(h,y+R+1);for(let x=0;x<w;x++){const x0=Math.max(0,x-R),x1=Math.min(w,x+R+1);
    const sum=I[y1*(w+1)+x1]-I[y0*(w+1)+x1]-I[y1*(w+1)+x0]+I[y0*(w+1)+x0];const mean=sum/((x1-x0)*(y1-y0));b[y*w+x]=sm[y*w+x]<mean*0.85&&sm[y*w+x]<200?1:0}}
  return b}
/* skew: the angle (−4°…4°) at which the rows of ink line up best (largest spread of the row sums) */
function skewOf(b,w,h){const st=Math.max(1,Math.round(w/900));const pts=[];for(let y=0;y<h;y+=st)for(let x=0;x<w;x+=st)if(b[y*w+x])pts.push(x,y);
  if(pts.length<200)return 0;let best=0,bv=-1;
  const score=a=>{const t=Math.tan(a*Math.PI/180);const rows=new Float64Array(Math.ceil(h/st)+Math.ceil(w*0.08/st)+8);const off=Math.ceil(w*0.08/st)+4;
    for(let i=0;i<pts.length;i+=2){const r=Math.round((pts[i+1]-pts[i]*t)/st)+off;if(r>=0&&r<rows.length)rows[r]++}let s=0;for(const v of rows)s+=v*v;return s};
  for(let a=-4;a<=4.001;a+=0.25){const v=score(a);if(v>bv){bv=v;best=a}}
  const lo=best-0.25,hi=best+0.25;for(let a=lo;a<=hi+1e-9;a+=0.05){const v=score(a);if(v>bv){bv=v;best=a}}
  return Math.abs(best)<0.08?0:best}
const rotated=(c,deg)=>{if(!deg)return c;const r=deg*Math.PI/180;const o=document.createElement('canvas');o.width=c.width;o.height=c.height;const x=o.getContext('2d',{willReadFrequently:true});
  x.fillStyle='#fff';x.fillRect(0,0,o.width,o.height);x.translate(o.width/2,o.height/2);x.rotate(-r);x.drawImage(c,-c.width/2,-c.height/2);return o};
/* long straight runs of ink are table rules: erase them, and keep the vertical ones as column borders */
/* a horizontal rule is at least a sixth of the page wide (joined Arabic writing has long flat strokes, but not that long) */
function rules(b,w,h){const LH=Math.max(60,Math.round(w/6)),LV=Math.max(40,Math.round(h/24)),L=LV;const hz=[],vt=[];const kill=new Uint8Array(b.length);
  for(let y=1;y<h-1;y++){let a=-1;for(let x=0;x<=w;x++){const on=x<w&&(b[y*w+x]||b[(y-1)*w+x]||b[(y+1)*w+x]);if(on&&a<0)a=x;else if(!on&&a>=0){if(x-a>=LH){hz.push({y,x0:a,x1:x});for(let k=a;k<x;k++)kill[y*w+k]=1}a=-1}}}
  for(let x=1;x<w-1;x++){let a=-1;for(let y=0;y<=h;y++){const on=y<h&&(b[y*w+x]||b[y*w+x-1]||b[y*w+x+1]);if(on&&a<0)a=y;else if(!on&&a>=0){if(y-a>=LV){vt.push({x,y0:a,y1:y});for(let k=a;k<y;k++)kill[k*w+x]=1}a=-1}}}
  /* a 1-pixel border around what was erased goes too (anti-aliased edges of the lines) */
  const k2=new Uint8Array(b.length);for(let i=0;i<b.length;i++)if(kill[i]){const x=i%w;for(let dy=-2;dy<=2;dy++)for(let dx=-2;dx<=2;dx++){const xx=x+dx,j=i+dy*w+dx;if(xx>=0&&xx<w&&j>=0&&j<b.length)k2[j]=1}}
  const clean=b.slice();for(let i=0;i<b.length;i++)if(k2[i])clean[i]=0;
  const group=(arr,k,s0,s1)=>{arr.sort((p,q)=>p[k]-q[k]);const out=[];arr.forEach(r=>{const g=out[out.length-1];if(g&&r[k]-g.last<=3&&!(r[s0]>g[s1]+L||r[s1]<g[s0]-L)){g.sum+=r[k];g.n++;g.last=r[k];g[s0]=Math.min(g[s0],r[s0]);g[s1]=Math.max(g[s1],r[s1])}else out.push({sum:r[k],n:1,last:r[k],[s0]:r[s0],[s1]:r[s1]})});
    return out.map(g=>({at:g.sum/g.n,[s0]:g[s0],[s1]:g[s1]}))};
  return {clean,kill:k2,h:group(hz,'y','x0','x1'),v:group(vt,'x','y0','y1')}}
/* specks left by grain: lone pixels with no inked neighbour */
function despeckle(b,w,h){const o=b.slice();for(let y=1;y<h-1;y++)for(let x=1;x<w-1;x++){const i=y*w+x;if(!b[i])continue;const n=b[i-1]+b[i+1]+b[i-w]+b[i+w]+b[i-w-1]+b[i-w+1]+b[i+w-1]+b[i+w+1];if(n<=1)o[i]=0}return o}
/* light writing on a dark fill (a report's coloured header row): made dark on light, block by block */
function invertDark(g,w,h){const B=16,bw=Math.ceil(w/B),bh=Math.ceil(h/B);const dark=new Uint8Array(bw*bh);
  for(let by=0;by<bh;by++)for(let bx=0;bx<bw;bx++){let s=0,n=0;for(let y=by*B;y<Math.min(h,by*B+B);y++)for(let x=bx*B;x<Math.min(w,bx*B+B);x++){s+=g[y*w+x];n++}dark[by*bw+bx]=s/n<128?1:0}
  /* a dark area is a fill only when it is wide (not a thick letter or a logo): a run of 6 blocks or more */
  const fill=new Uint8Array(bw*bh);for(let by=0;by<bh;by++){let a=-1;for(let bx=0;bx<=bw;bx++){const on=bx<bw&&dark[by*bw+bx];if(on&&a<0)a=bx;else if(!on&&a>=0){if(bx-a>=6)for(let k=a;k<bx;k++)fill[by*bw+k]=1;a=-1}}}
  /* the fill's exact top and bottom: the pixel rows (near a fill block) that are mostly dark across the fill's width */
  const done=new Uint8Array(h*w);let any=false;
  for(let by=0;by<bh;by++){let a=-1;for(let bx=0;bx<=bw;bx++){const on=bx<bw&&fill[by*bw+bx];if(on&&a<0)a=bx;else if(!on&&a>=0){const x0=a*B,x1=Math.min(w,bx*B);
      /* the band of pixel rows that are mostly dark across the fill: from its first such row to its last, so the
         rows crossed by light writing (more light than dark) are inverted too */
      const ys=[];for(let y=Math.max(0,by*B-B);y<Math.min(h,by*B+2*B);y++){let d=0;for(let x=x0;x<x1;x++)if(g[y*w+x]<128)d++;if(d>0.5*(x1-x0))ys.push(y)}
      if(ys.length)for(let y=ys[0];y<=ys[ys.length-1];y++){let l=Math.max(0,x0-B),r=Math.min(w,x1+B)-1;while(l<r&&g[y*w+l]>=128&&!(y>ys[0]&&y<ys[ys.length-1]&&l>=x0))l++;while(r>l&&g[y*w+r]>=128&&!(y>ys[0]&&y<ys[ys.length-1]&&r<x1))r--;
        for(let x=l;x<=r;x++){const i=y*w+x;if(!done[i]){g[i]=255-g[i];done[i]=1}}any=true}a=-1}}}
  return any}
/* a page scanned sideways or upside down: turned by k quarter turns (clockwise) before anything else */
function turned(src,w,h,k){if(!k)return {src,w,h};const c=document.createElement('canvas');const side=k%2===1;c.width=side?h:w;c.height=side?w:h;const x=c.getContext('2d');
  x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.translate(c.width/2,c.height/2);x.rotate(k*Math.PI/2);x.drawImage(src,-w/2,-h/2,w,h);return {src:c,w:c.width,h:c.height}}
/* lines of writing make the ink jump from row to row; on a page lying on its side the jumps run across the columns */
function sideways(src,w,h){const k=Math.min(1,700/Math.max(w,h));const c=document.createElement('canvas');c.width=Math.max(1,Math.round(w*k));c.height=Math.max(1,Math.round(h*k));
  const x=c.getContext('2d',{willReadFrequently:true});x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(src,0,0,c.width,c.height);
  const g=greyOf(c),W=c.width,Hh=c.height;let m=0;for(const v of g)m+=v;m/=g.length;const rows=new Float64Array(Hh),cols=new Float64Array(W);
  for(let y=0;y<Hh;y++)for(let x2=0;x2<W;x2++)if(g[y*W+x2]<m*0.75){rows[y]++;cols[x2]++}
  const jump=a=>{let s=0,t=0;for(let i=1;i<a.length;i++){s+=(a[i]-a[i-1])**2;t+=a[i]}return t?s/(t*t)*a.length:0};
  return jump(cols)>1.4*jump(rows)}
function prepare(src,w,h){let c=canvasOf(src,w,h);let g=greyOf(c);invertDark(g,c.width,c.height);let b=binarize(g,c.width,c.height);
  const sk=skewOf(b,c.width,c.height);if(sk){c=rotated(c,sk);g=greyOf(c);invertDark(g,c.width,c.height);b=binarize(g,c.width,c.height)}
  const W=c.width,Ht=c.height;const r=rules(b,W,Ht);const clean=despeckle(r.clean,W,Ht);
  /* the page for the recogniser: grey, with the table rules painted white */
  const o=document.createElement('canvas');o.width=W;o.height=Ht;const x=o.getContext('2d',{willReadFrequently:true});const id=x.createImageData(W,Ht);
  for(let i=0,j=0;j<g.length;i+=4,j++){const v=r.kill[j]?255:g[j];id.data[i]=id.data[i+1]=id.data[i+2]=v;id.data[i+3]=255}x.putImageData(id,0,0);
  return {view:c,ocr:o,ink:clean,skew:sk,hRules:r.h,vRules:r.v,width:W,height:Ht}}

/* ---------- the table's cells ---------- */
/* a ruled table: rows between the horizontal rules, columns between the vertical ones (where both cross) */
function ruledCells(pg){const W=pg.width;const minLen=0.25*W;
  const hs=pg.hRules.filter(r=>r.x1-r.x0>=minLen).sort((a,b)=>a.at-b.at);if(hs.length<3)return null;
  const x0=Math.min(...hs.map(r=>r.x0)),x1=Math.max(...hs.map(r=>r.x1)),top=hs[0].at,bot=hs[hs.length-1].at;
  const vs=pg.vRules.filter(r=>r.at>=x0-8&&r.at<=x1+8&&r.y1-r.y0>=0.3*(bot-top)).map(r=>r.at).sort((a,b)=>a-b);if(vs.length<2)return null;
  const merge=(arr,gap)=>{const o=[];arr.forEach(v=>{if(o.length&&v-o[o.length-1].last<=gap){const g=o[o.length-1];g.s+=v;g.n++;g.last=v}else o.push({s:v,n:1,last:v})});return o.map(g=>g.s/g.n)};
  const ys=merge(hs.map(r=>r.at),8),xs=merge(vs,8);if(ys.length<3||xs.length<2)return null;
  const rows=[];for(let i=0;i+1<ys.length;i++){if(ys[i+1]-ys[i]<12)continue;const row=[];
    for(let j=xs.length-2;j>=0;j--){if(xs[j+1]-xs[j]<12)continue;row.push({box:[xs[j]+3,ys[i]+3,xs[j+1]-3,ys[i+1]-3]})}rows.push(EN()?row.reverse():row)}
  return rows.length?rows:null}
/* a table without rules: the text lines come from the recogniser's words, the columns from the white "rivers" —
   vertical strips with no ink on (almost) any line of the table, wider than a space between two words */
function riverCells(pg,words){if(!words.length)return [];const W=pg.width;
  const hs=words.map(w=>w.y1-w.y0).sort((a,b)=>a-b);const med=hs[hs.length>>1]||20;
  const ws=words.slice().sort((a,b)=>(a.y0+a.y1)-(b.y0+b.y1));const lines=[];
  ws.forEach(w=>{const m=(w.y0+w.y1)/2;const l=lines[lines.length-1];if(l&&Math.abs(m-l.m)<=0.55*med){l.w.push(w);l.m=(l.m*(l.w.length-1)+m)/l.w.length;l.y0=Math.min(l.y0,w.y0);l.y1=Math.max(l.y1,w.y1)}else lines.push({m,w:[w],y0:w.y0,y1:w.y1})});
  /* ink of each line, column by column */
  const inkX=l=>{const a=new Uint8Array(W);for(let y=Math.max(0,l.y0|0);y<Math.min(pg.height,Math.ceil(l.y1));y++)for(let x=0;x<W;x++)if(pg.ink[y*W+x])a[x]=1;return a};
  lines.forEach(l=>{l.ink=inkX(l);let a=-1,b=-1;for(let x=0;x<W;x++)if(l.ink[x]){if(a<0)a=x;b=x}l.x0=a;l.x1=b});
  /* the table's lines: from the first line with several separated pieces of text */
  const pieces=l=>{let n=0,gap=1e9;for(let x=l.x0;x<=l.x1;x++){if(l.ink[x]){if(gap>=0.9*med)n++;gap=0}else gap++}return n};
  const first=lines.findIndex(l=>pieces(l)>=3);const T=first<0?lines:lines.slice(first);
  const cov=new Uint16Array(W);T.forEach(l=>{for(let x=0;x<W;x++)cov[x]+=l.ink[x]});
  const allow=Math.max(2,Math.floor(0.12*T.length));const minGap=Math.max(8,0.75*med);
  const tx0=Math.min(...T.map(l=>l.x0)),tx1=Math.max(...T.map(l=>l.x1));
  const cols=[];let a=tx0,g0=-1;
  for(let x=tx0;x<=tx1+1;x++){const empty=x>tx1||cov[x]<=allow;if(empty&&g0<0)g0=x;else if(!empty&&g0>=0){if(x-g0>=minGap){if(g0>a)cols.push([a,g0]);a=x}g0=-1}}
  if(tx1+1>a)cols.push([a,tx1+1]);
  /* a line whose text runs into a gap (a long name on a few rows): that piece joins the column it touches */
  const out=lines.map(l=>{const bx=cols.map(([x0,x1])=>[x0,x1]);
    for(let k=0;k+1<cols.length;k++){const g0=cols[k][1],g1=cols[k+1][0];const runs=[];
      for(let x=g0;x<g1;){if(!l.ink[x]){x++;continue}let e=x;while(e<g1&&l.ink[e])e++;runs.push([x,e]);x=e}
      if(!runs.length)continue;
      let eL=-1;for(let q=g0-1;q>=cols[k][0];q--)if(l.ink[q]){eL=q+1;break}let sR=-1;for(let q=g1;q<cols[k+1][1];q++)if(l.ink[q]){sR=q;break}
      let i0=0,i1=runs.length-1;const near=0.75*med;
      while(i0<=i1&&eL>=0&&runs[i0][0]-eL<near){eL=runs[i0][1];bx[k][1]=Math.max(bx[k][1],eL);i0++}
      while(i1>=i0&&sR>=0&&sR-runs[i1][1]<near){sR=runs[i1][0];bx[k+1][0]=Math.min(bx[k+1][0],sR);i1--}}
    return bx});
  if(!EN())cols.reverse();
  return lines.map((l,i)=>{const bx=out[i];if(!EN())bx.reverse();return bx.map(([x0,x1])=>({box:[x0,l.y0-2,x1,l.y1+2]}))})}
/* ink inside a box (after the rules were erased): an empty cell is not read at all */
const inked=(pg,[x0,y0,x1,y1])=>{let n=0;const W=pg.width;for(let y=Math.max(0,y0|0);y<Math.min(pg.height,y1);y++)for(let x=Math.max(0,x0|0);x<Math.min(W,x1);x++)n+=pg.ink[y*W+x];return n};
/* the ink's own bounds inside a cell, so the crop is tight around the text */
function tight(pg,box){const W=pg.width;let [x0,y0,x1,y1]=box.map(Math.round);
  const full=y=>{let n=0;for(let x=Math.max(0,x0);x<Math.min(W,x1);x++)n+=pg.ink[y*W+x];return n>0.7*(x1-x0)};
  while(y0<y1&&full(y0))y0++;while(y1>y0&&full(y1-1))y1--;
  let a=1e9,b=1e9,c=-1,d=-1;
  for(let y=Math.max(0,y0|0);y<Math.min(pg.height,y1);y++)for(let x=Math.max(0,x0|0);x<Math.min(W,x1);x++)if(pg.ink[y*W+x]){if(x<a)a=x;if(x>c)c=x;if(y<b)b=y;if(y>d)d=y}
  return c<0?null:[a,b,c+1,d+1]}
/* crop of a cell, enlarged so the text is about 48 px high, with a white margin */
function crop(pg,box,lineH){const [x0,y0,x1,y1]=box;const w=x1-x0,h=y1-y0;const k=Math.max(1,Math.min(4,48/Math.max(1,lineH||h)));const pad=Math.round(12);
  const c=document.createElement('canvas');c.width=Math.round(w*k)+pad*2;c.height=Math.round(h*k)+pad*2;const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);
  x.imageSmoothingQuality='high';x.drawImage(pg.ocr,x0,y0,w,h,pad,pad,c.width-pad*2,c.height-pad*2);return c}
/* ---------- words → table ---------- */
const DIG=s=>String(s).replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/٫/g,'.').replace(/٬/g,',');
const JUNK=/^[|¦!\[\]{}()_~`'"«»‘’“”،؛:;.,\-—–=+*#^\\\/]+$/;
function wordsOf(data){const out=[];(data.blocks||[]).forEach(bl=>(bl.paragraphs||[]).forEach(p=>(p.lines||[]).forEach(l=>(l.words||[]).forEach(wd=>{
  let t=String(wd.text||'').replace(/[‎‏‪-‮]/g,'').replace(/^[|¦\[\]{}]+|[|¦\[\]{}]+$/g,'').trim();if(!t||JUNK.test(t)&&t.length<3)return;
  const b=wd.bbox;out.push({t,c:Math.max(0,Math.min(100,wd.confidence||0)),x0:b.x0,x1:b.x1,y0:b.y0,y1:b.y1,ar:AR.test(t)})}))));return out}
const isNum=t=>/^[-+(]?[\d٠-٩۰-۹][\d٠-٩۰-۹,٬.٫\s]*\)?-?%?$/.test(t);
/* the text of a cell from its words: Arabic reads right to left, numbers and Latin left to right */
const joinCell=ws=>{const ar=ws.filter(w=>w.ar).length>=ws.length/2;return ws.slice().sort((a,b)=>ar?b.x0-a.x0:a.x0-b.x0).map(w=>w.t).join(' ')};

/* a cell's value: numbers as numbers (Arabic-Indic digits, thousands separators, (negative)), codes with a leading
   zero and dates stay text. A figure that does not read as a clean number in a column of numbers is flagged. */
function valueOf(t){const s=DIG(t).replace(/\s+/g,'');if(!s)return '';
  const m=s.match(/^(\()?([-+])?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?(\))?(-)?$/);
  if(!m||(!!m[1]!==!!m[5]))return DIG(t).trim();if(/^0\d/.test(m[3])&&!m[4])return s;
  const n=Number(m[3].replace(/,/g,'')+(m[4]||''));if(!isFinite(n))return s;return (m[1]||m[2]==='-'||m[6])?-n:n}
const looksNum=t=>/^[-+(]?[\d.,٬٫٠-٩۰-۹]+\)?-?$/.test(DIG(t).replace(/\s/g,''));
const SUSPECT=t=>{const s=DIG(t).replace(/\s/g,'');return /[OoIlSsBZ|]/.test(s)&&/\d/.test(s)||/,,|\.\.|,\.|\.,|\.\d*\./.test(s)||/,\d{1,2}(?!\d)|,\d{4,}/.test(s.replace(/\.\d+$/,''))};
/* page tables → one sheet: the header row of the first page is kept once, the same header on later pages is dropped */
function toSheet(pages,name){const rows=[];let head=null;
  pages.forEach((p,pi)=>p.table.rows.forEach((r,ri)=>{const key=r.map(c=>c?c.t:'').join('|');
    if(ri<4&&head&&key===head)return;if(pi===0&&!head&&r.filter(Boolean).length>=Math.max(2,0.6*p.table.cols)&&!r.some(c=>c&&isNum(c.t)))head=key;
    rows.push(r.map(c=>c?Object.assign({page:pi},c):null))}));
  const ncol=Math.max(0,...rows.map(r=>r.length));
  /* columns of numbers: three quarters of their filled cells read as numbers */
  const numCol=Array.from({length:ncol},(_,j)=>{const v=rows.map(r=>r[j]).filter(c=>c&&c.t);return v.length>=2&&v.filter(c=>looksNum(c.t)).length>=0.75*v.length});
  const cells=rows.map(r=>Array.from({length:ncol},(_,j)=>{const c=r[j];if(!c)return {t:'',c:100,page:null,box:null,empty:true};
    let conf=c.c;const num=looksNum(c.t);const z=DIG(c.t).replace(/\s/g,'');
    /* 15.276 or 1.234.567: a dot where a thousands comma is usual — a thousand times off if it is wrong */
    if(/^\d{1,3}(?:\.\d{3})+$/.test(z)&&!c.alt){conf=Math.min(conf,40);c.alt=z.replace(/\./g,',')}
    /* Arabic words with stray Latin letters: the reader lost part of a word */
    if(AR.test(c.t)&&/(^|\s)[A-Za-z]{1,3}(?=\s|$)/.test(c.t))conf=Math.min(conf,40);if(numCol[j]&&!num&&!/^[\s\-–—]*$/.test(c.t)&&r!==rows[0])conf=Math.min(conf,35);if(num&&SUSPECT(c.t))conf=Math.min(conf,40);
    return {t:DIG(c.t).trim(),c:Math.round(conf),page:c.page,box:c.box,num:num||numCol[j],alt:c.alt?DIG(c.alt).trim():null,blank:c.blank&&!c.t}}));
  return {name,cells,numCol}}
const needsCheck=c=>!c.empty&&(c.t!==''||c.blank)&&c.c<(c.num?NUM_OK:TEXT_OK);

/* ---------- a table, or running text? ---------- */
/* in an Arabic page, short Latin "words" read with low confidence are smudges, stamps and frame lines, not text */
function cleanWords(ws){const ar=ws.filter(w=>w.ar).length;if(ar<0.6*ws.length)return ws;
  return ws.filter(w=>w.ar||/[\d٠-٩]/.test(w.t)||w.c>=50)}
const NUMW=/^[-+(]?[\d٠-٩][\d٠-٩,٬.٫\/\-]*\)?%?$/;
function pageKind(pg,words){if(words.length<6)return {kind:'empty'};
  const hs=words.map(w=>w.y1-w.y0).sort((a,b)=>a-b);const med=hs[hs.length>>1]||20;
  const nums=words.filter(w=>NUMW.test(w.t)).length;
  /* a ruled grid counts when most of the page's words are inside it and its cells hold single lines, not paragraphs */
  const ruled=ruledCells(pg);
  if(ruled){const bx=ruled.flat().map(c=>c.box);const X0=Math.min(...bx.map(b=>b[0])),Y0=Math.min(...bx.map(b=>b[1])),X1=Math.max(...bx.map(b=>b[2])),Y1=Math.max(...bx.map(b=>b[3]));
    const inside=words.filter(w=>(w.x0+w.x1)/2>=X0&&(w.x0+w.x1)/2<=X1&&(w.y0+w.y1)/2>=Y0&&(w.y0+w.y1)/2<=Y1);
    const ch=ruled.map(r=>r[0].box[3]-r[0].box[1]).sort((a,b)=>a-b);const cols=Math.max(...ruled.map(r=>r.length));
    /* words in cells taller than eight lines: a paragraph in a frame (a contract page), not a table cell */
    const cellOf=w=>{const m=(w.x0+w.x1)/2,y=(w.y0+w.y1)/2;for(const r of ruled)for(const c of r){const b=c.box;if(m>=b[0]-3&&m<=b[2]+3&&y>=b[1]-3&&y<=b[3]+3)return c}return null};
    const tall=inside.filter(w=>{const c=cellOf(w);return c&&c.box[3]-c.box[1]>8*med}).length;
    if(inside.length>=0.4*words.length&&ch[ch.length>>1]<=5*med&&tall<0.6*inside.length&&cols>=2&&ruled.length>=3&&inside.filter(w=>NUMW.test(w.t)).length>=3)return {kind:'table',ruled:true,rows:ruled,med}}
  /* no rules: at least three columns, filled on a good share of the lines, and figures among the words */
  const rows=riverCells(pg,words);const cols=rows.length?rows[0].length:0;
  const filled=rows.map(r=>r.filter(c=>words.some(w=>{const m=(w.x0+w.x1)/2,y=(w.y0+w.y1)/2;return m>=c.box[0]&&m<=c.box[2]&&y>=c.box[1]&&y<=c.box[3]})).length);
  const wide=filled.filter(n=>n>=3).length;
  if(cols>=3&&wide>=3&&wide>=0.4*rows.length&&nums>=Math.max(3,0.05*words.length))return {kind:'table',ruled:false,rows,med};
  return {kind:'text'}}
const notTable=()=>fail('هذا الملف نص مكتوب (مثل عقد أو كتاب رسمي) وليس جدول أرقام، فلا تُستخرج منه بيانات للوحة. لوحات المؤشرات تحتاج تقريرًا على شكل جدول: تحليل المبيعات، أو حركة التحويلات، أو أذونات الاستلام، أو SMS055، أو كشف الديون، أو ملف هامش الربح. وإن أردت حفظ العقد نفسه فافتحه من صفحة «العقود» ← تعديل العقد ← «ملف العقد والمرفقات»، واستخدم «رفع ملف» أو «مسح ضوئي».',
  'This file is running text (such as a contract or a letter), not a table of figures, so there is nothing to import into a dashboard. Dashboards need a report laid out as a table: sales analysis, transfers, receipts, SMS055, the debts statement or the profit-margin file. To keep the contract itself, open it in «Contracts» → edit → «Contract file and attachments» and use «Upload file» or «Scan».',{notTable:true});
/* ---------- reading ---------- */
const readText=async(kind,cv)=>{const w=await worker(kind);const {data}=await w.recognize(cv,{},{text:true,blocks:true});
  const ws=wordsOf(data);const t=ws.length?joinCell(ws):String(data.text||'').replace(/[‎‏]/g,'').trim();
  return {t:t.replace(/\s+/g,' ').trim(),c:ws.length?Math.min(...ws.map(x=>x.c)):Math.round(data.confidence||0)}};
const normNum=t=>DIG(t).replace(/[\s]/g,'');
const DATE=/^(?:\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}(?:[\/\-.]\d{1,2})?)$/;
async function recognisePages(srcs,onProgress){const pages=[],prose=[];
  try{for(let i=0;i<srcs.length;i++){const s=srcs[i];onProgress&&onProgress(i,0,tr('تجهيز الصفحة','Preparing the page'));
      /* the right way up: a page on its side is tried both ways, an upright page that reads badly also upside down */
      const tryTurn=async k=>{const t=turned(s.src,s.w,s.h,k);const p=prepare(t.src,t.w,t.h);let ink=0;for(let q=0;q<p.ink.length;q+=7)ink+=p.ink[q];
        if(ink*7<0.0015*p.ink.length)return {pg:p,words:[],score:-1,blank:true};
        const w=await worker('page');const {data}=await w.recognize(p.ocr,{},{blocks:true});const words=cleanWords(wordsOf(data));
        return {k,pg:p,words,score:words.reduce((a,x)=>a+(x.c>=60?x.c:0),0),mean:words.length?words.reduce((a,x)=>a+x.c,0)/words.length:0}};
      onProgress&&onProgress(i,0.02,tr('تحديد اتجاه الصفحة','Finding the page orientation'));
      let best;if(sideways(s.src,s.w,s.h)){/* the ink test is only a hint (frames, logos and tables mislead it): upright is tried too */
        const c0=await tryTurn(0),a=await tryTurn(1),b=await tryTurn(3);best=[c0,a,b].reduce((p,q)=>q.score>p.score*(p===c0?1.15:1)?q:p)}
      else{best=await tryTurn(0);if(!best.blank&&best.mean<55&&best.words.length>=6){const b=await tryTurn(2);if(b.score>best.score*1.3)best=b}}
      const pg=best.pg;
      /* a blank page (the back of a sheet, an empty end page) is skipped */
      if(best.blank)continue;
      /* the words of the whole page tell a table from running text (a contract, a letter) */
      onProgress&&onProgress(i,0.08,tr('تحديد أعمدة الجدول','Finding the table columns'));
      const words=best.words;
      const kind=pageKind(pg,words);pg.kind=kind.kind;
      if(kind.kind!=='table'){if(kind.kind==='text')prose.push(i);if(!pages.length&&prose.length>=Math.min(2,srcs.length))break;continue}
      let rows=kind.rows;pg.ruled=kind.ruled;
      /* an Arabic page: a cell read as a few Latin letters with low confidence is a smudge — left empty to be filled by hand */
      const arPage=words.filter(x=>x.ar).length>=0.6*words.length;
      const todo=[];rows.forEach(r=>r.forEach(c=>{if(!c.box)return;const tb=tight(pg,c.box);if(tb&&inked(pg,tb)>12)todo.push([c,tb]);else c.box=c.first?c.box:null}));
      let done=0;const step=()=>onProgress&&onProgress(i,0.1+0.9*(done/Math.max(1,todo.length*1.6)),tr('قراءة الخلايا','Reading the cells'));
      const lineH=kind.med||20;
      for(const [c,tb] of todo){const cv=crop(pg,tb,tb[3]-tb[1]>1.8*lineH?lineH:0);let r=await readText(tb[3]-tb[1]>1.8*lineH?'block':'text',cv);
        /* a doubtful or empty read: once more as a raw line (some short Arabic cells are only found that way), the surer read wins */
        if(r.c<60||!r.t){const r2=await readText('raw',cv);if(r2.t&&(!r.t||r2.c>r.c))r=r2}c.t=r.t;c.c=r.c;c.box=tb;if(!c.t&&c.first){c.t=c.first.t;c.c=Math.min(c.first.c,50)}if(arPage&&c.t&&!AR.test(c.t)&&!/[\d٠-٩]/.test(c.t)&&c.c<45)c.t='';if(!c.t){c.c=0;c.blank=true}done++;step()}
      /* every figure is read a second time by a digits-only reader: two different reads → the cell has to be checked */
      const ncol=Math.max(0,...rows.map(r=>r.length));
      const numCol=Array.from({length:ncol},(_,j)=>{const v=rows.map(r=>r[j]).filter(c=>c&&c.t);return v.length>=2&&v.filter(c=>looksNum(c.t)||DATE.test(normNum(c.t))).length>=0.6*v.length});
      const nums=[];rows.forEach(r=>r.forEach((c,j)=>{if(c&&c.t&&c.box&&(looksNum(c.t)||DATE.test(normNum(c.t))||(numCol[j]&&!AR.test(c.t))||(/[\d٠-٩]/.test(c.t)&&!AR.test(c.t.replace(/[٠-٩٫٬]/g,'')))))nums.push(c)}));
      for(const c of nums){const r=await readText('digits',crop(pg,c.box));const a=normNum(c.t),b=normNum(r.t);
        if(a!==b){c.alt=r.t;if(!looksNum(c.t)&&!DATE.test(a)&&looksNum(r.t)){c.t=r.t;c.alt=null}c.c=Math.min(c.c,a.replace(/[,.]/g,'')===b.replace(/[,.]/g,'')?55:30)}
        done+=0.6;step()}
      pg.table={rows:rows.map(r=>r.map(c=>c&&(c.t||c.blank)?{t:c.t||'',c:c.c,box:c.box,alt:c.alt||null,blank:!!c.blank}:null)),cols:ncol};pages.push(pg)}}
  finally{release()}
  if(!pages.length&&prose.length)throw notTable();
  return pages}
const loaderFail=()=>fail('تعذّر تحميل قارئ الملفات الممسوحة ضوئيًا. تحقق من الاتصال ثم أعد المحاولة.','Could not load the scanned-file reader. Check the connection and try again.');
async function run(srcs,name,meta){const ui=Review.progress(name,srcs.length);let pages;
  try{try{await worker('text');await worker('digits')}catch(e){throw loaderFail()}
    if(ui.cancelled())throw fail('أُلغي الاستيراد.','Import cancelled.',{cancelled:true});
    pages=await recognisePages(srcs,(i,p,st)=>ui.update(i,p,st));}
  catch(e){ui.close();throw e}
  const sheet=toSheet(pages,name);
  if(!sheet.cells.length||sheet.cells.every(r=>r.every(c=>c.empty))){ui.close();throw fail('لم أتعرف على أي نص في الملف. تأكد أن الصورة واضحة ومستقيمة وأن الإضاءة جيدة، أو صدّر التقرير بصيغة Excel.','No text could be recognised. Make sure the picture is sharp, straight and well lit, or export the report as Excel.')}
  /* a picture too poor to read (a photo of a screen, blurred, at an angle): most cells are doubtful — say so plainly */
  const cs=sheet.cells.flat().filter(c=>!c.empty&&(c.t||c.blank)).map(c=>c.c).sort((a,b)=>a-b);
  meta=Object.assign({},meta,{poor:!cs.length||cs[cs.length>>1]<55||cs.filter(c=>c<60).length>0.5*cs.length});
  const grid=await ui.review(sheet,pages,meta);
  return [{name,grid,text:true,ocr:{pages:pages.length,cells:sheet.cells.reduce((n,r)=>n+r.filter(c=>!c.empty).length,0),checked:sheet.checked||0}}]}
async function readImage(file){const name=String(file.name||'scan').replace(/\.[^.]+$/,'')||'scan';
  const bm=await bitmapOf(file);if(!bm){const u8=new Uint8Array(await file.slice(0,12).arrayBuffer());
    if(isTiff(u8))throw fail('صور TIFF لا تُعرض في هذا المتصفح. احفظ الصورة بصيغة JPG أو PNG أو PDF ثم أعد المحاولة.','TIFF pictures cannot be opened in this browser. Save the picture as JPG, PNG or PDF and try again.');
    throw fail('تعذّر فتح الصورة: يبدو أنها تالفة أو بصيغة غير مدعومة. احفظها بصيغة JPG أو PNG ثم أعد المحاولة.','Could not open the picture: it looks damaged or unsupported. Save it as JPG or PNG and try again.')}
  return run([{src:bm,w:bm.width,h:bm.height}],name,{kind:'image'})}
async function readScannedPdf(doc,name){const n=Math.min(doc.numPages,MAX_PAGES);const srcs=[];
  for(let i=1;i<=n;i++){const pg=await doc.getPage(i);const v1=pg.getViewport({scale:1});const s=Math.min(4.2,2400/Math.max(v1.width,v1.height)*1.4);const vp=pg.getViewport({scale:s});
    const c=document.createElement('canvas');c.width=Math.ceil(vp.width);c.height=Math.ceil(vp.height);const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);
    await pg.render({canvasContext:x,viewport:vp}).promise;pg.cleanup();srcs.push({src:c,w:c.width,h:c.height})}
  return run(srcs,name,{kind:'pdf',total:doc.numPages,read:n})}

/* ---------- the review window ---------- */
const Review={
  progress(name,n){const d=document.createElement('dialog');d.className='ocr-dlg';d.setAttribute('dir',EN()?'ltr':'rtl');let cancelled=false;
    d.innerHTML=`<div class="m-head"><h3>${H(tr('قراءة ملف ممسوح ضوئيًا','Reading a scanned file'))}</h3></div>
     <div class="m-body ocr-wait"><div class="ocr-spin" aria-hidden="true"></div><b dir="auto">${H(name)}</b><span class="ocr-st" role="status">${H(tr('تحميل قارئ النصوص العربية…','Loading the Arabic text reader…'))}</span>
     <div class="ocr-bar"><i style="width:2%"></i></div><small>${H(tr('تجري القراءة على جهازك ولا يُرسل الملف إلى أي مكان. قد تستغرق كل صفحة بضع ثوانٍ.','Reading happens on this device; the file is not sent anywhere. Each page can take a few seconds.'))}</small></div>
     <div class="m-foot"><button type="button" class="btn ghost" data-ocr="cancel">${H(tr('إلغاء','Cancel'))}</button></div>`;
    document.body.appendChild(d);d.addEventListener('cancel',e=>{e.preventDefault()});d.showModal();
    d.querySelector('[data-ocr="cancel"]').onclick=()=>{cancelled=true;d.querySelector('.ocr-st').textContent=tr('جارٍ الإلغاء…','Cancelling…')};
    const close=()=>{try{d.close()}catch(e){}d.remove()};
    return {cancelled:()=>cancelled,close,
      update(i,p,st){if(cancelled)throw fail('أُلغي الاستيراد.','Import cancelled.',{cancelled:true});const pct=Math.round(((i+(p||0))/n)*100);
        d.querySelector('.ocr-bar i').style.width=Math.max(2,pct)+'%';d.querySelector('.ocr-st').textContent=`${st} · ${tr('صفحة','page')} ${i+1} ${tr('من','of')} ${n} · ${pct}%`},
      review:(sheet,pages,meta)=>{close();return Review.open(sheet,pages,meta)}}},
  open(sheet,pages,meta){return new Promise((resolve,reject)=>{
    const S={cells:sheet.cells,pages,cur:null,page:0,ok:new Set()};
    const d=document.createElement('dialog');d.className='ocr-dlg ocr-review';d.setAttribute('dir',EN()?'ltr':'rtl');
    const poorNote=meta&&meta.poor?`<p class="ocr-note warn" role="alert">${H(tr('الصورة غير واضحة بما يكفي: أغلب الخلايا قُرئت بثقة منخفضة، وتصحيحها يدويًا يطول وقد يفوتك خطأ. الأفضل تصوير الورقة نفسها لا الشاشة، من الأعلى مباشرة وبإضاءة جيدة، أو تصدير التقرير من المنظومة بصيغة Excel.','The picture is not clear enough: most cells were read with low confidence, and correcting them by hand takes long and can miss errors. Photograph the paper itself (not a screen), straight from above in good light, or export the report as Excel.'))}</p>`:'';
    const pdfNote=meta&&meta.kind==='pdf'&&meta.total>meta.read?`<p class="ocr-note warn">${H(tr(`قُرئت أول ${meta.read} صفحة فقط من ${meta.total}. قسّم الملف واستورد الباقي في ملف آخر.`,`Only the first ${meta.read} of ${meta.total} pages were read. Split the file and import the rest separately.`))}</p>`:'';
    d.innerHTML=`<div class="m-head"><h3>${H(tr('مراجعة النص المستخرج من الملف الممسوح','Review the text read from the scan'))}</h3><span class="ocr-name" dir="auto">${H(sheet.name)}</span></div>
     <div class="ocr-main">
      <section class="ocr-src" aria-label="${H(tr('الأصل','Original'))}"><div class="ocr-src-h"><b>${H(tr('الأصل','Original'))}</b>${pages.length>1?`<span class="ocr-pages">${pages.map((_,i)=>`<button type="button" class="chip" data-pg="${i}">${i+1}</button>`).join('')}</span>`:''}</div>
       <div class="ocr-crop"><canvas aria-hidden="true"></canvas><small>${H(tr('اختر خلية لترى مكانها في الأصل مكبّرًا.','Pick a cell to see it enlarged from the original.'))}</small></div>
       <div class="ocr-page"><canvas aria-hidden="true"></canvas><i class="ocr-hl"></i></div></section>
      <section class="ocr-tbl-w">
       <p class="ocr-note">${H(tr('قراءة الصور آليًا قد تخطئ في الأرقام. طابق كل خلية ملوّنة مع الأصل: صحّحها بالكتابة فيها، أو اضغط ✓ (أو Enter) إن كانت صحيحة. الأحمر ثقة منخفضة، والأصفر متوسطة.','Machine reading of pictures can get figures wrong. Check every coloured cell against the original: type the correct value, or press ✓ (or Enter) if it is right. Red is low confidence, amber is medium.'))}</p>${poorNote}${pdfNote}
       <div class="ocr-legend"><span class="lv hi"></span>${H(tr('ثقة عالية','High confidence'))}<span class="lv mid"></span>${H(tr('متوسطة: تحتاج تأكيد','Medium: confirm'))}<span class="lv lo"></span>${H(tr('منخفضة: تحتاج تأكيد','Low: confirm'))}<span class="lv done"></span>${H(tr('روجعت','Checked'))}</div>
       <div class="ocr-tbl" role="region" tabindex="0" aria-label="${H(tr('الجدول المستخرج','Extracted table'))}"><table><tbody></tbody></table></div>
      </section></div>
     <div class="m-foot ocr-foot"><span class="ocr-left" role="status"></span><button type="button" class="btn ghost" data-ocr="next">${H(tr('الخلية التالية للمراجعة','Next cell to check'))}</button>
      <label class="ocr-agree"><input type="checkbox" data-ocr="agree"> ${H(tr('طابقت الأرقام مع المستند الأصلي','I checked the figures against the original'))}</label>
      <span style="margin-inline-start:auto"></span><button type="button" class="btn ghost" data-ocr="cancel">${H(tr('إلغاء','Cancel'))}</button><button type="button" class="btn primary" data-ocr="ok" disabled>${H(tr('اعتماد ومتابعة الاستيراد','Approve and continue'))}</button></div>`;
    document.body.appendChild(d);d.addEventListener('cancel',e=>{e.preventDefault()});
    const tb=d.querySelector('tbody'),okB=d.querySelector('[data-ocr="ok"]'),agree=d.querySelector('[data-ocr="agree"]'),left=d.querySelector('.ocr-left');
    const key=(r,c)=>r+':'+c;
    const lvl=c=>c.empty||(c.t===''&&!c.blank)?'':S.ok.has(c)?'done':!needsCheck(c)?'hi':c.c<(c.num?60:45)?'lo':'mid';
    const pending=()=>{let n=0;S.cells.forEach(r=>r.forEach(c=>{if(needsCheck(c)&&!S.ok.has(c))n++}));return n};
    const render=()=>{tb.innerHTML=S.cells.map((r,i)=>`<tr><th scope="row"><button type="button" class="iconbtn" data-del-row="${i}" title="${H(tr('حذف الصف','Delete row'))}" aria-label="${H(tr('حذف الصف','Delete row'))} ${i+1}">×</button>${i+1}</th>${r.map((c,j)=>
      `<td class="ocr-c ${lvl(c)}${c.num?' n':''}" data-r="${i}" data-c="${j}"><input value="${H(c.t)}" size="${Math.max(5,Math.min(40,[...String(c.t)].length+2))}" dir="auto" aria-label="${H(tr('صف','Row'))} ${i+1}، ${H(tr('عمود','column'))} ${j+1}${c.empty?'':' · '+H(tr('الثقة','confidence'))+' '+c.c+'%'}" title="${c.empty?'':H(tr('الثقة','Confidence'))+' '+c.c+'%'}">${needsCheck(c)&&!S.ok.has(c)?`<button type="button" class="ocr-ok" data-ok="${i}:${j}" title="${H(tr('صحيحة','Correct'))}" aria-label="${H(tr('تأكيد الخلية','Confirm cell'))}">✓</button>${c.alt?`<button type="button" class="ocr-alt" data-alt="${i}:${j}" dir="ltr" title="${H(tr('قراءة ثانية مختلفة: اضغط لاعتمادها','A second, different reading: click to use it'))}">${H(c.alt)}</button>`:''}`:''}</td>`).join('')}</tr>`).join('')
      +`<tr class="ocr-colrow"><th></th>${(S.cells[0]||[]).map((_,j)=>`<td><button type="button" class="btn ghost sm" data-del-col="${j}">${H(tr('حذف العمود','Delete column'))}</button></td>`).join('')}</tr>`;status()};
    const status=()=>{const n=pending();left.textContent=n?tr(`بقيت ${n} خلية تحتاج مراجعة`,`${n} cells still to check`):tr('روجعت كل الخلايا المشكوك فيها','Every doubtful cell is checked');left.classList.toggle('warn',!!n);okB.disabled=!!n||!agree.checked};
    const pageCv=d.querySelector('.ocr-page canvas'),hl=d.querySelector('.ocr-hl'),crop=d.querySelector('.ocr-crop canvas');
    const showPage=i=>{S.page=i;const p=pages[i];pageCv.width=p.view.width;pageCv.height=p.view.height;pageCv.getContext('2d').drawImage(p.view,0,0);d.querySelectorAll('[data-pg]').forEach(b=>b.classList.toggle('on',+b.dataset.pg===i));hl.hidden=true};
    const focusCell=(i,j)=>{const c=S.cells[i]&&S.cells[i][j];if(!c)return;S.cur=c;
      if(c.page!=null&&c.box){if(c.page!==S.page)showPage(c.page);const p=pages[c.page];const [x0,y0,x1,y1]=c.box;const pad=Math.max(12,(y1-y0)*0.6);
        const sx=Math.max(0,x0-pad*2),sy=Math.max(0,y0-pad),sw=Math.min(p.view.width-sx,x1-x0+pad*4),sh=Math.min(p.view.height-sy,y1-y0+pad*2);
        const k=Math.min(3,440/sw,110/sh);crop.width=Math.round(sw*k);crop.height=Math.round(sh*k);const cx=crop.getContext('2d');cx.imageSmoothingQuality='high';cx.drawImage(p.view,sx,sy,sw,sh,0,0,crop.width,crop.height);crop.parentNode.classList.add('on');
        const sc=pageCv.clientWidth/p.view.width;Object.assign(hl.style,{left:(x0-4)*sc+'px',top:(y0-4)*sc+'px',width:(x1-x0+8)*sc+'px',height:(y1-y0+8)*sc+'px'});hl.hidden=false;
        const box=d.querySelector('.ocr-page');const t=(y0*sc)-box.clientHeight/2;if(Math.abs(box.scrollTop-t)>box.clientHeight/3)box.scrollTop=t}
      else{crop.parentNode.classList.remove('on');hl.hidden=true}};
    const confirmCell=(i,j,moveOn)=>{const c=S.cells[i][j];S.ok.add(c);const td=tb.querySelector(`td[data-r="${i}"][data-c="${j}"]`);if(td){td.className=`ocr-c ${lvl(c)}${c.num?' n':''}`;td.querySelectorAll('.ocr-ok,.ocr-alt').forEach(b=>b.remove())}status();if(moveOn)next(i,j)};
    const next=(fi,fj)=>{const all=[];S.cells.forEach((r,i)=>r.forEach((c,j)=>{if(needsCheck(c)&&!S.ok.has(c))all.push([i,j])}));if(!all.length){agree.focus();return}
      const after=all.find(([i,j])=>fi==null||i>fi||(i===fi&&j>fj))||all[0];const inp=tb.querySelector(`td[data-r="${after[0]}"][data-c="${after[1]}"] input`);inp&&(inp.focus(),inp.select())};
    tb.addEventListener('focusin',e=>{const td=e.target.closest('td.ocr-c');if(td)focusCell(+td.dataset.r,+td.dataset.c)});
    tb.addEventListener('input',e=>{const td=e.target.closest('td.ocr-c');if(!td)return;const c=S.cells[+td.dataset.r][+td.dataset.c];c.t=e.target.value;c.edited=true;if(c.empty&&c.t)c.empty=false;if(!S.ok.has(c)){S.ok.add(c);td.className=`ocr-c ${lvl(c)}${c.num?' n':''}`;td.querySelectorAll('.ocr-ok,.ocr-alt').forEach(b=>b.remove());status()}});
    tb.addEventListener('keydown',e=>{const td=e.target.closest('td.ocr-c');if(!td||e.target.tagName!=='INPUT')return;const i=+td.dataset.r,j=+td.dataset.c;
      if(e.key==='Enter'){e.preventDefault();const c=S.cells[i][j];if(needsCheck(c)&&!S.ok.has(c))confirmCell(i,j,true);else next(i,j)}
      else if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();const n=tb.querySelector(`td[data-r="${i+(e.key==='ArrowDown'?1:-1)}"][data-c="${j}"] input`);n&&n.focus()}});
    tb.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;
      if(b.dataset.ok){const [i,j]=b.dataset.ok.split(':').map(Number);confirmCell(i,j,true)}
      else if(b.dataset.alt){const [i,j]=b.dataset.alt.split(':').map(Number);const c=S.cells[i][j];c.t=c.alt;c.edited=true;const inp=tb.querySelector(`td[data-r="${i}"][data-c="${j}"] input`);if(inp)inp.value=c.t;b.remove();confirmCell(i,j,true)}
      else if(b.dataset.delRow!=null){S.cells.splice(+b.dataset.delRow,1);render()}
      else if(b.dataset.delCol!=null){const j=+b.dataset.delCol;S.cells.forEach(r=>r.splice(j,1));render()}});
    d.addEventListener('click',e=>{const b=e.target.closest('[data-pg],[data-ocr]');if(!b)return;
      if(b.dataset.pg!=null)showPage(+b.dataset.pg);
      else if(b.dataset.ocr==='next')next(null,null);
      else if(b.dataset.ocr==='cancel'){d.close();d.remove();reject(fail('أُلغي الاستيراد.','Import cancelled.',{cancelled:true}))}
      else if(b.dataset.ocr==='ok'&&!okB.disabled){sheet.checked=S.ok.size;const grid=S.cells.map(r=>r.map(c=>c.empty&&!c.t?'':valueOf(c.t)));
        while(grid.length&&grid[grid.length-1].every(v=>v===''))grid.pop();d.close();d.remove();resolve(grid)}});
    agree.addEventListener('change',status);
    render();showPage(0);d.showModal();setTimeout(()=>next(null,null),50)})}};

window.OCRImport={isImage,readImage,readScannedPdf,_:{tight,crop,recognisePages,ruledCells,prepare,riverCells,toSheet,wordsOf,valueOf,needsCheck,skewOf,lib,worker,NUM_OK,TEXT_OK}};
})();
