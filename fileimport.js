/* ===== قراءة ملفات الاستيراد بكل الصيغ الشائعة (FileImport) =====
   One reader for every import button (SMS055, sales analysis, transfers, receipts, debts, profit margin and the
   contracts import in Settings). Loaded right after settings.js (which provides unzip()).
   The format is recognised from the file's content, not only its extension:
     ZIP  → .xlsx .xlsm .xltx .xltm (and WPS .et/.ett saved in the new format) · .xlsb · .ods .ots
     OLE2 → .xls .xlt and WPS .et .ett (BIFF8, and Excel 5/95 BIFF5)
     text → .csv .tsv .tab .txt (comma, semicolon, tab or pipe; UTF-8, UTF-16 or Windows-1256 Arabic)
            .htm .html (tables, including the «.xls» files that are really HTML), Excel 2003 XML (.xml),
            flat OpenDocument (.fods) and .json (array of rows, array of arrays, or {sheet:[rows]})
     PDF  → .pdf printed reports, rebuilt from the text layer (pdf.js, loaded on demand); a scan (no text layer) goes to
            the text recogniser with its compulsory review window (ocrimport.js)
     image → .jpg .png .webp .bmp .gif photos or scanner pictures of a report: text recogniser and review (ocrimport.js)
   FileImport.readGrids(file) → [{name, grid}] where grid is rows of cells: number, string, boolean, or {serial:n}
   for a date-formatted cell (Excel serial day, 1900 system). FileImport.readWorkbook(file) → {SheetNames, rows(name)}
   with the first row as the header (dates as yyyy-mm-dd). Errors are thrown with a clear message in the page language. */
"use strict";
(function(){
const EN=()=>{if(window.I18N&&I18N.lang)return I18N.lang==='en';try{return localStorage.getItem('cm_lang')==='en'}catch(e){return false}};
const tr=(ar,en)=>EN()&&en?en:ar;
const ACCEPT='.xlsx,.xlsm,.xltx,.xltm,.xlsb,.xls,.xlt,.et,.ett,.ods,.ots,.fods,.csv,.tsv,.tab,.txt,.htm,.html,.xml,.json,.pdf,.jpg,.jpeg,.png,.webp,.bmp,.gif,'+
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.oasis.opendocument.spreadsheet,text/csv,text/tab-separated-values,application/pdf,image/jpeg,image/png,image/webp';
const LIST='Excel (xlsx · xlsm · xlsb · xls) · WPS (et · ett) · OpenDocument (ods) · PDF (also scanned) · JPG · PNG · CSV · TSV · TXT · HTML · XML · JSON';
const MAX=200*1024*1024,MAX_CELLS=5e6;
const fail=(ar,en)=>{const e=new Error(tr(ar,en));e.importError=true;return e};
const unsupported=()=>fail('صيغة الملف غير مدعومة. الصيغ المقبولة: '+LIST+'. افتح الملف في Excel أو WPS واحفظه بإحدى هذه الصيغ ثم أعد المحاولة.','This file type is not supported. Accepted: '+LIST+'. Open the file in Excel or WPS, save it in one of these formats and try again.');
const damaged=()=>fail('تعذّرت قراءة الملف: يبدو أنه تالف أو غير مكتمل. افتحه في Excel أو WPS واحفظه من جديد ثم أعد المحاولة.','Could not read the file: it looks damaged or incomplete. Open it in Excel or WPS, save it again and retry.');
const locked=()=>fail('الملف محمي بكلمة مرور. أزل الحماية (ملف ← معلومات ← حماية المصنف) واحفظه ثم أعد الاستيراد.','The file is password-protected. Remove the protection, save it and import again.');
const isDateCode=code=>{const c=String(code||'').toLowerCase().replace(/"[^"]*"|\\.|\[[^\]]*\]/g,'');return /[dy]/.test(c)&&!/[#0?]/.test(c)};
/* built-in number formats that are dates: 14–22, 45–47, and the East-Asian locale dates 27–36, 50–58 (WPS writes these) */
const isDateId=id=>(id>=14&&id<=22)||(id>=27&&id<=36)||(id>=45&&id<=47)||(id>=50&&id<=58);
const isDateFmt=(id,custom)=>custom[id]!=null?isDateCode(custom[id]):isDateId(id);
const serialOf=(y,m,d,hh,mi,ss)=>(Date.UTC(y,m-1,d,hh||0,mi||0,ss||0)-Date.UTC(1899,11,30))/86400000;
const isoSerial=s=>{const x=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);return x?serialOf(+x[1],+x[2],+x[3],+(x[4]||0),+(x[5]||0),+(x[6]||0)):null};
const trimGrid=grid=>{while(grid.length&&!grid[grid.length-1].some(c=>c!==''&&c!=null))grid.pop();return grid.map(r=>{let n=r.length;while(n&&(r[n-1]===''||r[n-1]==null))n--;return Array.from({length:n},(_,i)=>r[i]==null?'':r[i])})};
const colIdx=ref=>{const L=(String(ref).match(/^[A-Z]+/)||[''])[0];let c=0;for(const ch of L)c=c*26+(ch.charCodeAt(0)-64);return L?c-1:-1};
const XP=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw damaged();return d};

/* ---------- Office Open XML (.xlsx .xlsm .xltx .xltm, WPS et/ett in the new format) ---------- */
async function readXlsx(z){const M='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const wbx=await z.text('xl/workbook.xml');const wb=XP(wbx);const pr=wb.getElementsByTagNameNS(M,'workbookPr')[0];const d1904=!!pr&&/^(1|true)$/.test(pr.getAttribute('date1904')||'');
  const rels={};const rx=await z.text('xl/_rels/workbook.xml.rels');if(rx)[...XP(rx).getElementsByTagName('Relationship')].forEach(r=>rels[r.getAttribute('Id')]=r.getAttribute('Target'));
  const path=t=>{t=String(t||'');return t.startsWith('/')?t.slice(1):'xl/'+t.replace(/^\.\//,'')};
  const ssName=Object.values(rels).find(t=>/sharedStrings/i.test(t));
  const shared=[];const ss=await z.text(ssName?path(ssName):'xl/sharedStrings.xml');
  if(ss){const d=XP(ss);[...d.getElementsByTagNameNS(M,'si')].forEach(si=>{const ts=[...si.getElementsByTagNameNS(M,'t')].filter(t=>!(t.parentNode&&t.parentNode.localName==='rPh'));shared.push(ts.map(t=>t.textContent).join(''))})}
  const dateSt=new Set();const sty=await z.text('xl/styles.xml');
  if(sty){const d=XP(sty);const custom={};[...d.getElementsByTagNameNS(M,'numFmt')].forEach(f=>custom[+f.getAttribute('numFmtId')]=f.getAttribute('formatCode')||'');
    const xfs=d.getElementsByTagNameNS(M,'cellXfs')[0];if(xfs)[...xfs.getElementsByTagNameNS(M,'xf')].forEach((xf,i)=>{if(isDateFmt(+(xf.getAttribute('numFmtId')||0),custom))dateSt.add(i)})}
  const out=[];let cells=0;
  for(const sh of [...wb.getElementsByTagNameNS(M,'sheet')]){const xml=await z.text(path(rels[sh.getAttributeNS(R,'id')]));if(!xml)continue;
    const d=XP(xml);const grid=[];
    [...d.getElementsByTagNameNS(M,'row')].forEach((row,ri)=>{const rn=+(row.getAttribute('r')||0);const at=rn>0?rn-1:grid.length;while(grid.length<at)grid.push([]);
      const vals=[];[...row.getElementsByTagNameNS(M,'c')].forEach((c,j)=>{let ci=colIdx(c.getAttribute('r')||'');if(ci<0)ci=vals.length;
        const ty=c.getAttribute('t'),s=+(c.getAttribute('s')||0),v=c.getElementsByTagNameNS(M,'v')[0];let val='';
        if(ty==='s'&&v)val=shared[+v.textContent]??'';
        else if(ty==='inlineStr')val=[...c.getElementsByTagNameNS(M,'t')].map(x=>x.textContent).join('');
        else if(ty==='str'||ty==='e')val=v?v.textContent:'';
        else if(ty==='b')val=v?v.textContent==='1':'';
        else if(ty==='d'){const n=v?isoSerial(v.textContent):null;val=n==null?(v?v.textContent:''):{serial:n}}
        else if(v){const n=parseFloat(v.textContent);if(isNaN(n))val=v.textContent;else if(dateSt.has(s))val={serial:d1904?n+1462:n};else val=n}
        vals[ci]=val;cells++});
      if(cells>MAX_CELLS)throw fail('الملف أكبر من أن يُقرأ في المتصفح. قسّمه إلى ملفات أصغر.','The file is too large to read in the browser. Split it into smaller files.');
      grid[at]=Array.from(vals,x=>x===undefined?'':x)});
    out.push({name:sh.getAttribute('name')||('Sheet'+(out.length+1)),grid:trimGrid(grid)})}
  return out}

/* ---------- Excel binary workbook (.xlsb) ---------- */
async function readXlsb(z){
  const recs=u8=>{const out=[];let p=0;const vi=()=>{let v=0,s=0,b;do{b=u8[p++];v|=(b&127)<<s;s+=7}while(b&128&&s<28);return v};
    while(p<u8.length){const t=vi(),n=vi();out.push([t,u8.subarray(p,p+n)]);p+=n}return out};
  const dv=b=>new DataView(b.buffer,b.byteOffset,b.byteLength);
  const wstr=(b,o)=>{const n=dv(b).getUint32(o,true);let s='';for(let i=0;i<n;i++)s+=String.fromCharCode(dv(b).getUint16(o+4+i*2,true));return [s,o+4+n*2]};
  const rk=x=>{let v;if(x&2)v=x>>2;else{const b=new DataView(new ArrayBuffer(8));b.setUint32(4,x&0xFFFFFFFC,true);v=b.getFloat64(0,true)}return x&1?v/100:v};
  const rels={};const rx=await z.text('xl/_rels/workbook.bin.rels');if(rx)[...XP(rx).getElementsByTagName('Relationship')].forEach(r=>rels[r.getAttribute('Id')]=r.getAttribute('Target'));
  const sheets=[];recs(await z.bytes('xl/workbook.bin')).forEach(([t,b])=>{if(t===156){const [rid,o]=wstr(b,8);const [nm]=wstr(b,o);sheets.push([nm,rels[rid]])}});
  const shared=[];const sb=await z.bytes('xl/sharedStrings.bin');if(sb)recs(sb).forEach(([t,b])=>{if(t===19)shared.push(wstr(b,1)[0])});
  const custom={},xfs=[];const st=await z.bytes('xl/styles.bin');let inXf=false;
  if(st)recs(st).forEach(([t,b])=>{if(t===44)custom[dv(b).getUint16(0,true)]=wstr(b,2)[0];else if(t===617)inXf=true;else if(t===618)inXf=false;else if(t===47&&inXf)xfs.push(dv(b).getUint16(2,true))});
  const out=[];
  for(const [nm,target] of sheets){const b=await z.bytes('xl/'+String(target||'').replace(/^\//,'').replace(/^xl\//,''));if(!b)continue;const grid=[];let row=0;
    recs(b).forEach(([t,c])=>{if(t===0){row=dv(c).getUint32(0,true);return}if(t<1||t>11)return;const d=dv(c);const col=d.getUint32(0,true),xf=d.getUint32(4,true)&0xFFFFFF;let v='';
      if(t===2||t===5||t===9){const n=t===2?rk(d.getUint32(8,true)):d.getFloat64(8,true);v=isDateFmt(xfs[xf]||0,custom)?{serial:n}:n}
      else if(t===7)v=shared[d.getUint32(8,true)]??'';else if(t===6||t===8)v=wstr(c,8)[0];else if(t===4||t===10)v=!!c[8];else return;
      (grid[row]||(grid[row]=[]))[col]=v});
    out.push({name:nm,grid:trimGrid(Array.from(grid,r=>r?Array.from(r,x=>x===undefined?'':x):[]))})}
  return out}

/* ---------- OpenDocument spreadsheet (.ods .ots, and flat .fods) ---------- */
function readOdsXml(xml){const T='urn:oasis:names:tc:opendocument:xmlns:table:1.0',O='urn:oasis:names:tc:opendocument:xmlns:office:1.0',X='urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  const d=XP(xml);const out=[];
  for(const tb of [...d.getElementsByTagNameNS(T,'table')]){const grid=[];let blankRows=0;
    const rows=[...tb.getElementsByTagNameNS(T,'table-row')];
    for(const r of rows){const rr=Math.min(+(r.getAttributeNS(T,'number-rows-repeated')||1),100000);const vals=[];let blank=0;
      for(const c of [...r.children].filter(n=>n.namespaceURI===T&&/^(covered-)?table-cell$/.test(n.localName))){const rep=Math.min(+(c.getAttributeNS(T,'number-columns-repeated')||1),16384);
        const ty=c.getAttributeNS(O,'value-type')||'';let v='';
        if(ty==='float'||ty==='percentage'||ty==='currency')v=parseFloat(c.getAttributeNS(O,'value'));
        else if(ty==='date'){const n=isoSerial(c.getAttributeNS(O,'date-value'));v=n==null?'':{serial:n}}
        else if(ty==='boolean')v=c.getAttributeNS(O,'boolean-value')==='true';
        else{const ps=[...c.getElementsByTagNameNS(X,'p')];v=ps.length?ps.map(p=>p.textContent).join('\n'):'';if(ty==='time'&&!v)v=c.getAttributeNS(O,'time-value')||''}
        if(typeof v==='number'&&isNaN(v))v='';
        if(v===''){blank+=rep;continue}
        while(blank){vals.push('');blank--}for(let i=0;i<rep;i++)vals.push(v)}
      if(!vals.length){blankRows+=rr;continue}
      while(blankRows){grid.push([]);blankRows--}for(let i=0;i<rr;i++)grid.push(vals.slice())
      if(grid.length*Math.max(1,vals.length)>MAX_CELLS)throw fail('الملف أكبر من أن يُقرأ في المتصفح. قسّمه إلى ملفات أصغر.','The file is too large to read in the browser. Split it into smaller files.')}
    out.push({name:tb.getAttributeNS(T,'name')||('Sheet'+(out.length+1)),grid:trimGrid(grid)})}
  return out}

/* ---------- OLE2 compound file → BIFF8 / BIFF5 (.xls .xlt, WPS .et .ett) ---------- */
function cfbStreams(buf){const u8=new Uint8Array(buf),d=new DataView(buf);if(u8.length<512)throw damaged();
  const ss=1<<d.getUint16(30,true),ms=1<<d.getUint16(32,true),nFat=d.getUint32(44,true),dir0=d.getUint32(48,true),cutoff=d.getUint32(56,true),mf0=d.getUint32(60,true),nMf=d.getUint32(64,true);let dif=d.getUint32(68,true),nDif=d.getUint32(72,true);
  if(ss!==512&&ss!==4096)throw damaged();
  const off=s=>(s+1)*ss;const END=0xFFFFFFFE;
  const fatSecs=[];for(let i=0;i<109&&fatSecs.length<nFat;i++)fatSecs.push(d.getUint32(76+i*4,true));
  while(dif<END&&nDif-->0&&fatSecs.length<nFat){for(let i=0;i<ss/4-1&&fatSecs.length<nFat;i++)fatSecs.push(d.getUint32(off(dif)+i*4,true));dif=d.getUint32(off(dif)+ss-4,true)}
  const fat=new Uint32Array(fatSecs.length*ss/4);fatSecs.forEach((s,i)=>{for(let j=0;j<ss/4;j++){const o=off(s)+j*4;fat[i*ss/4+j]=o+4<=u8.length?d.getUint32(o,true):END}});
  const chain=(s,t)=>{const out=[];let g=0;while(s<END&&s<t.length&&g++<t.length){out.push(s);s=t[s]}return out};
  const read=(s,size)=>{const c=chain(s,fat);const b=new Uint8Array(c.length*ss);c.forEach((x,i)=>b.set(u8.subarray(off(x),off(x)+ss),i*ss));return size==null?b:b.subarray(0,size)};
  const dir=read(dir0);const dd=new DataView(dir.buffer);const ents=[];
  for(let o=0;o+128<=dir.length;o+=128){const nl=dd.getUint16(o+64,true);if(!nl)continue;let nm='';for(let i=0;i<nl/2-1;i++)nm+=String.fromCharCode(dd.getUint16(o+i*2,true));
    ents.push({nm,type:dir[o+66],start:dd.getUint32(o+116,true),size:dd.getUint32(o+120,true)})}
  const root=ents.find(e=>e.type===5);let mini=null,mfat=null;
  const get=name=>{const e=ents.find(x=>x.type===2&&x.nm.toLowerCase()===name.toLowerCase());if(!e)return null;
    if(e.size>=cutoff||!root)return read(e.start,e.size);
    if(!mini){mini=read(root.start,root.size);const mb=read(mf0);mfat=new Uint32Array(mb.buffer,0,Math.min(nMf*ss/4,mb.length/4|0))}
    const c=chain(e.start,mfat);const b=new Uint8Array(c.length*ms);c.forEach((x,i)=>b.set(mini.subarray(x*ms,x*ms+ms),i*ms));return b.subarray(0,e.size)};
  return {names:ents.map(e=>e.nm),get}}
function readBiff(b){const d=new DataView(b.buffer,b.byteOffset,b.byteLength);const n=b.length;
  let biff=8,cp='windows-1252',d1904=false;const sheets=[],sst=[],custom={},xfFmt=[];
  const dec=(bytes)=>{try{return new TextDecoder(cp).decode(bytes)}catch(e){return new TextDecoder('windows-1252').decode(bytes)}};
  const u16s=(o,len)=>{let s='';for(let i=0;i<len;i++)s+=String.fromCharCode(d.getUint16(o+i*2,true));return s};
  /* string with a 16-bit (or 8-bit) count followed by an options byte (BIFF8), or plain code-page bytes (BIFF5) */
  const xstr=(o,wide)=>{const cch=wide?d.getUint16(o,true):b[o];o+=wide?2:1;if(biff<8)return dec(b.subarray(o,o+cch));const f=b[o++];
    if(f&8)o+=2;if(f&4)o+=4;return f&1?u16s(o,cch):dec8(o,cch)};
  const dec8=(o,len)=>{let s='';for(let i=0;i<len;i++)s+=String.fromCharCode(b[o+i]);return s};
  const rk=x=>{let v;if(x&2)v=x>>2;else{const t=new DataView(new ArrayBuffer(8));t.setUint32(4,x&0xFFFFFFFC,true);v=t.getFloat64(0,true)}return x&1?v/100:v};
  const num=(v,xf)=>isDateFmt(xfFmt[xf]??0,custom)?{serial:d1904?v+1462:v}:v;
  /* the shared string table can run over several CONTINUE records; a string cut at a boundary restarts with a new options byte */
  const readSst=segs=>{let si=0,p=0;
    const byte=()=>{while(si<segs.length&&p>=segs[si].length){si++;p=0}if(si>=segs.length)throw damaged();return segs[si][p++]};
    const rd=k=>{let v=0;for(let i=0;i<k;i++)v+=byte()*2**(8*i);return v};
    rd(4);const uniq=rd(4);
    for(let i=0;i<uniq;i++){while(si<segs.length&&p>=segs[si].length){si++;p=0}if(si>=segs.length)break;
      const cch=rd(2);let f=byte();const runs=f&8?rd(2):0,ext=f&4?rd(4):0;let s='';
      for(let k=0;k<cch;k++){if(p>=segs[si].length){si++;p=0;if(si>=segs.length)break;f=segs[si][p++]}s+=String.fromCharCode(f&1?byte()|(byte()<<8):byte())}
      for(let k=runs*4+ext;k>0;k--){if(si>=segs.length)break;if(p<segs[si].length)p++;else{si++;p=0;k++}}
      sst.push(s)}};
  let p=0,ws=null,pendingFormula=null,sstSegs=null,first=true;const stack=[];
  while(p+4<=n){const t=d.getUint16(p,true),len=d.getUint16(p+2,true),o=p+4;p=o+len;if(p>n)break;
    if(sstSegs&&t!==0x3C){readSst(sstSegs);sstSegs=null}
    switch(t){
      case 0x0809:case 0x0409:case 0x0209:case 0x0009:{const ver=d.getUint16(o,true),dt=d.getUint16(o+2,true);if(first){biff=ver===0x0600?8:ver===0x0500?5:t===0x0809?8:5;first=false}
        stack.push(ws);
        if(dt===0x10){const i=sheets.findIndex(s=>s.pos===o-4);ws=i>=0?sheets[i]:{name:'Sheet'+(sheets.length+1),grid:[]};if(i<0)sheets.push(ws);ws.grid=ws.grid||[]}else ws=null;break}
      case 0x000A:ws=stack.length?stack.pop():null;break;
      case 0x002F:throw locked();
      case 0x0042:{const c=d.getUint16(o,true);cp=c===1200?'utf-16le':c===1256?'windows-1256':c===1252||c===32769?'windows-1252':c>=874&&c<=1258?'windows-'+c:c===10000?'macintosh':'windows-1252';break}
      case 0x0022:d1904=d.getUint16(o,true)===1;break;
      case 0x0085:{const pos=d.getUint32(o,true),kind=b[o+5];const name=biff<8?dec(b.subarray(o+7,o+7+b[o+6])):xstr(o+6,false);if(kind===0)sheets.push({pos,name,grid:null});break}
      case 0x041E:case 0x001E:{if(biff>=8){custom[d.getUint16(o,true)]=xstr(o+2,true)}else{custom[d.getUint16(o,true)]=dec(b.subarray(o+3,o+3+b[o+2]))}break}
      case 0x00E0:xfFmt.push(d.getUint16(o+2,true));break;
      case 0x00FC:sstSegs=[b.subarray(o,o+len)];break;
      case 0x003C:if(sstSegs)sstSegs.push(b.subarray(o,o+len));break;
      default:if(ws){const g=ws.grid,r=d.getUint16(o,true),c=d.getUint16(o+2,true),xf=d.getUint16(o+4,true);const put=(rr,cc,v)=>{(g[rr]||(g[rr]=[]))[cc]=v};
        if(t===0x0203)put(r,c,num(d.getFloat64(o+6,true),xf));
        else if(t===0x027E)put(r,c,num(rk(d.getUint32(o+6,true)),xf));
        else if(t===0x00BD){const last=d.getUint16(o+len-2,true);for(let k=0,cc=c;cc<=last;cc++,k++){const q=o+4+k*6;put(r,cc,num(rk(d.getUint32(q+2,true)),d.getUint16(q,true)))}}
        else if(t===0x00FD)put(r,c,sst[d.getUint32(o+6,true)]??'');
        else if(t===0x0204||t===0x00D6)put(r,c,xstr(o+6,true));
        else if(t===0x0205){const v=b[o+6],err=b[o+7];put(r,c,err?'':!!v)}
        else if(t===0x0006||t===0x0406||t===0x0206){const tag=d.getUint16(o+12,true);if(tag===0xFFFF){const k=b[o+6];if(k===0)pendingFormula=[r,c];else if(k===1)put(r,c,!!b[o+8]);else put(r,c,'')}else put(r,c,num(d.getFloat64(o+6,true),xf))}
        else if(t===0x0207&&pendingFormula){put(pendingFormula[0],pendingFormula[1],biff<8?dec(b.subarray(o+2,o+2+d.getUint16(o,true))):xstr(o,true));pendingFormula=null}
        }}}
  if(sstSegs)readSst(sstSegs);
  return sheets.filter(s=>s.grid).map(s=>({name:s.name,grid:trimGrid(Array.from(s.grid,r=>r?Array.from(r,x=>x===undefined?'':x):[]))}))}

/* ---------- text: delimited, HTML tables, Excel 2003 XML, flat ODS, JSON ---------- */
function wideText(u8){if(u8[0]===0xFF&&u8[1]===0xFE||u8[0]===0xFE&&u8[1]===0xFF)return true;let z=0;const k=Math.min(u8.length,400);for(let i=1;i<k;i+=2)if(u8[i]===0)z++;return k>8&&z>k/4}
function decodeText(u8){
  if(u8[0]===0xFF&&u8[1]===0xFE)return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if(u8[0]===0xFE&&u8[1]===0xFF)return new TextDecoder('utf-16be').decode(u8.subarray(2));
  if(u8[0]===0xEF&&u8[1]===0xBB&&u8[2]===0xBF)return new TextDecoder('utf-8').decode(u8.subarray(3));
  /* UTF-16 without a mark (WPS «Unicode text»): every other byte is zero */
  if(wideText(u8))return new TextDecoder('utf-16le').decode(u8);
  try{return new TextDecoder('utf-8',{fatal:true}).decode(u8)}catch(e){}
  /* Excel and WPS save «CSV» on Arabic Windows in the ANSI code page */
  return new TextDecoder('windows-1256').decode(u8)}
function delimGrid(text,name){const lines=text.split(/\r\n|\n|\r/).filter(l=>l.trim()).slice(0,30);
  const score=ch=>{const c=lines.map(l=>{let n=0,q=false;for(const x of l){if(x==='"')q=!q;else if(x===ch&&!q)n++}return n});const nz=c.filter(x=>x>0);if(!nz.length)return -1;const m=Math.max(...nz),same=c.filter(x=>x===m).length;return same*100+m};
  const sep=['\t',',',';','|'].map(c=>[c,score(c)]).sort((a,b)=>b[1]-a[1])[0];const d=sep[1]>0?sep[0]:'\t';
  const rows=[];let row=[],f='',q=false,quoted=false;
  for(let i=0;i<text.length;i++){const c=text[i];
    if(q){if(c==='"'){if(text[i+1]==='"'){f+='"';i++}else q=false}else f+=c;continue}
    if(c==='"'&&f.trim()===''&&!quoted){q=true;quoted=true;f='';continue}
    if(c===d){row.push(quoted?f:f.trim());f='';quoted=false;continue}
    if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(quoted?f:f.trim());rows.push(row);row=[];f='';quoted=false;continue}
    f+=c}
  if(f!==''||row.length){row.push(quoted?f:f.trim());rows.push(row)}
  return [{name,grid:trimGrid(rows),text:true}]}
function htmlGrids(text){const d=new DOMParser().parseFromString(text,'text/html');const out=[];
  [...d.querySelectorAll('table')].forEach((t,i)=>{if(t.querySelector('table'))return;const grid=[];
    [...t.rows].forEach((r,ri)=>{const row=grid[ri]||(grid[ri]=[]);let ci=0;
      [...r.cells].forEach(c=>{while(row[ci]!==undefined)ci++;const v=(c.innerText!=null&&c.innerText!==''?c.innerText:c.textContent).replace(/ /g,' ').trim();
        const cs=Math.min(+c.colSpan||1,200),rs=Math.min(+c.rowSpan||1,500);row[ci]=v;for(let k=1;k<cs;k++)row[ci+k]='';
        for(let k=1;k<rs;k++){const rr=grid[ri+k]||(grid[ri+k]=[]);for(let j=0;j<cs;j++)rr[ci+j]=''}ci+=cs})});
    out.push({name:(t.caption&&t.caption.textContent.trim())||('table'+(out.length+1)),grid:trimGrid(grid.map(r=>Array.from(r,x=>x===undefined?'':x))),text:true})});
  return out}
function spreadsheetMl(text){const S='urn:schemas-microsoft-com:office:spreadsheet';const d=XP(text);const out=[];
  for(const w of [...d.getElementsByTagNameNS(S,'Worksheet')]){const grid=[];let r=0;
    for(const row of [...w.getElementsByTagNameNS(S,'Row')]){const ix=+(row.getAttributeNS(S,'Index')||0);if(ix)r=ix-1;const vals=[];let c=0;
      for(const cell of [...row.getElementsByTagNameNS(S,'Cell')]){const cx=+(cell.getAttributeNS(S,'Index')||0);if(cx)c=cx-1;const dt=cell.getElementsByTagNameNS(S,'Data')[0];let v='';
        if(dt){const ty=dt.getAttributeNS(S,'Type'),s=dt.textContent;v=ty==='Number'?parseFloat(s):ty==='DateTime'?(isoSerial(s)!=null?{serial:isoSerial(s)}:s):ty==='Boolean'?s==='1':s;if(typeof v==='number'&&isNaN(v))v=''}
        vals[c]=v;c+=1+(+(cell.getAttributeNS(S,'MergeAcross')||0))}
      grid[r++]=Array.from(vals,x=>x===undefined?'':x)}
    out.push({name:w.getAttributeNS(S,'Name')||('Sheet'+(out.length+1)),grid:trimGrid(Array.from(grid,x=>x||[]))})}
  return out}
function jsonGrids(text,name){let j;try{j=JSON.parse(text)}catch(e){throw fail('ملف JSON غير صالح: تأكد أنه مصفوفة صفوف مثل [{"العمود":"القيمة"}].','Invalid JSON file: it should be an array of rows such as [{"column":"value"}].')}
  const toGrid=a=>{if(!Array.isArray(a))return null;if(a.every(Array.isArray))return a.map(r=>r.map(cellJ));if(!a.every(r=>r&&typeof r==='object'))return null;
    const keys=[];a.forEach(r=>Object.keys(r).forEach(k=>{if(!keys.includes(k))keys.push(k)}));return [keys,...a.map(r=>keys.map(k=>cellJ(r[k])))]};
  const cellJ=v=>v==null?'':typeof v==='number'||typeof v==='boolean'?v:typeof v==='object'?JSON.stringify(v):String(v);
  const g=toGrid(j);if(g)return [{name,grid:trimGrid(g)}];
  if(j&&typeof j==='object'){const out=Object.entries(j).map(([k,v])=>{const gg=toGrid(v);return gg?{name:k,grid:trimGrid(gg)}:null}).filter(Boolean);if(out.length)return out}
  throw fail('ملف JSON لا يحتوي جدولًا: المطلوب مصفوفة صفوف مثل [{"العمود":"القيمة"}].','The JSON file has no table: expected an array of rows such as [{"column":"value"}].')}

/* ---------- PDF (.pdf): tables rebuilt from the text layer ----------
   pdf.js (vendor/pdfjs, loaded only when a PDF is chosen) gives every glyph with its position. Glyphs are placed on
   the page, grouped into lines by baseline, split into cells at wide gaps, and the cells are lined up into columns
   from the gaps shared by all the lines. Arabic is rebuilt from the visual order on the page: each glyph keeps its own
   text (so ligatures such as «لا» stay right), runs of Arabic are read right to left and numbers, codes and dates left
   to right. An Arabic page gives its columns from right to left, as the report reads. Pages that start with a
   different header become a separate sheet; repeated headers, page numbers and running titles are dropped.
   A PDF with no text layer (a scan or photo) gets a clear message instead. */
let pdfjsP=null;
function pdfjs(){if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);if(!pdfjsP)pdfjsP=new Promise((res,rej)=>{const s=document.createElement('script');s.src='vendor/pdfjs/pdf.min.js';s.integrity='sha384-/1qUCSGwTur9vjf/z9lmu/eCUYbpOTgSjmpbMQZ1/CtX2v/WcAIKqRv+U1DUCG6e';s.crossOrigin='anonymous';
  s.onload=()=>{if(!window.pdfjsLib)return rej(new Error('pdfjs'));pdfjsLib.GlobalWorkerOptions.workerSrc='vendor/pdfjs/pdf.worker.min.js';res(pdfjsLib)};s.onerror=()=>{pdfjsP=null;s.remove();rej(new Error('pdfjs'))};document.head.appendChild(s)});return pdfjsP}
const PDF_MAX_PAGES=400;
const AR=/[֐-ࣿיִ-﷿ﹰ-﻿]/,LTRC=/[A-Za-z0-9٠-٩۰-۹À-ɏ]/,MARK=/^[ؐ-ًؚ-ٰٟۖ-ۭ̀-ͯ]+$/;
const MIRROR={'(':')',')':'(','[':']',']':'[','{':'}','}':'{','<':'>','>':'<','«':'»','»':'«'};
const mxMul=(m,n)=>[m[0]*n[0]+m[2]*n[1],m[1]*n[0]+m[3]*n[1],m[0]*n[2]+m[2]*n[3],m[1]*n[2]+m[3]*n[3],m[0]*n[4]+m[2]*n[5]+m[4],m[1]*n[4]+m[3]*n[5]+m[5]];
/* every visible glyph of a page: {x (left), y (baseline), w, s (size), u (text)} in page units, y downwards */
async function pdfGlyphs(lib,page){const O=lib.OPS;const ol=await page.getOperatorList();const vt=page.getViewport({scale:1}).transform;
  const out=[];let ctm=[1,0,0,1,0,0],runs=0;const stack=[];const forms=[];
  let st={font:null,size:0,cs:0,ws:0,hs:1,lead:0,rise:0,tm:[1,0,0,1,0,0],x:0,y:0,lx:0,ly:0};
  const fontOf=name=>{try{return page.commonObjs.has(name)?page.commonObjs.get(name):null}catch(e){return null}};
  const show=glyphs=>{if(!st.font||!Array.isArray(glyphs))return;const fm=(st.font.fontMatrix||[0.001])[0]||0.001,size=st.size,vert=!!st.font.vertical;if(vert)return;
    const M=mxMul(vt,mxMul(ctm,st.tm));const sx=Math.hypot(M[0],M[1]),sy=Math.hypot(M[2],M[3]);const flat=Math.abs(M[1])<0.2*sx&&Math.abs(M[2])<0.2*sy;
    let x=0,prev=null;const run=++runs;
    for(const g of glyphs){if(typeof g==='number'){x-=g*size/1000;continue}if(!g)continue;
      const adv=(g.width||0)*fm*size+st.cs+(g.isSpace?st.ws:0);const u=g.unicode==null?'':String(g.unicode);
      if(flat&&u.trim()){const tx=(st.x+x*st.hs),ty=st.y+st.rise;const px=M[0]*tx+M[2]*ty+M[4],py=M[1]*tx+M[3]*ty+M[5];
        const w=Math.abs((g.width||0)*fm*size*st.hs*sx),s=Math.abs(size*sy);
        if(MARK.test(u)&&prev)prev.u+=u;else{const x0=M[0]<0?px-w:px;prev={x:x0,y:py,w,s,u,r:run};out.push(prev)}}
      else if(!u.trim())prev=null;
      x+=adv}
    st.x+=x*st.hs};
  /* vertical table rules (cell borders), drawn as thin lines or thin filled rectangles */
  const rules=[];const PAINT=new Set([O.stroke,O.closeStroke,O.fill,O.eoFill,O.fillStroke,O.eoFillStroke,O.closeFillStroke,O.closeEOFillStroke]);
  const vline=(x0,y0,x1,y1)=>{const M=mxMul(vt,ctm);const p=(x,y)=>[M[0]*x+M[2]*y+M[4],M[1]*x+M[3]*y+M[5]];const A=p(x0,y0),B=p(x1,y1);
    if(Math.abs(A[0]-B[0])<=2&&Math.abs(A[1]-B[1])>4)rules.push({x:(A[0]+B[0])/2,y0:Math.min(A[1],B[1]),y1:Math.max(A[1],B[1])})};
  const path=(ops,c)=>{let k=0,cx=0,cy=0;for(const op of ops){
    if(op===O.rectangle){const [x,y,w,h]=c.slice(k,k+4);k+=4;if(Math.abs(w)<=2)vline(x+w/2,y,x+w/2,y+h);else if(Math.abs(h)>4&&Math.abs(w)>4){vline(x,y,x,y+h);vline(x+w,y,x+w,y+h)}}
    else if(op===O.moveTo){cx=c[k++];cy=c[k++]}
    else if(op===O.lineTo){const x=c[k++],y=c[k++];vline(cx,cy,x,y);cx=x;cy=y}
    else if(op===O.curveTo){k+=6}else if(op===O.curveTo2||op===O.curveTo3){k+=4}}};
  const fa=ol.fnArray,aa=ol.argsArray;
  for(let i=0;i<fa.length;i++){const f=fa[i],a=aa[i]||[];
    switch(f){
      case O.constructPath:if(PAINT.has(fa[i+1])&&Array.isArray(a[0]))path(a[0],a[1]||[]);break;
      case O.save:stack.push({ctm,st:Object.assign({},st)});break;
      case O.restore:{const p=stack.pop();if(p){ctm=p.ctm;st=p.st}break}
      case O.transform:ctm=mxMul(ctm,a);break;
      case O.paintFormXObjectBegin:forms.push(ctm);if(Array.isArray(a[0])&&a[0].length===6)ctm=mxMul(ctm,a[0]);break;
      case O.paintFormXObjectEnd:if(forms.length)ctm=forms.pop();break;
      case O.beginText:st.tm=[1,0,0,1,0,0];st.x=st.y=st.lx=st.ly=0;break;
      case O.setFont:st.font=fontOf(a[0]);st.size=a[1];break;
      case O.setGState:(a[0]||[]).forEach(e=>{if(e&&e[0]==='Font'&&e[1]){st.font=fontOf(e[1][0]);st.size=e[1][1]}});break;
      case O.setCharSpacing:st.cs=a[0];break;
      case O.setWordSpacing:st.ws=a[0];break;
      case O.setHScale:st.hs=a[0]/100;break;
      case O.setLeading:st.lead=-a[0];break;
      case O.setTextRise:st.rise=a[0];break;
      case O.setTextMatrix:st.tm=a.slice(0,6);st.x=st.y=st.lx=st.ly=0;break;
      case O.moveText:st.x=st.lx+=a[0];st.y=st.ly+=a[1];break;
      case O.setLeadingMoveText:st.lead=a[1];st.x=st.lx+=a[0];st.y=st.ly+=a[1];break;
      case O.nextLine:st.x=st.lx;st.y=st.ly+=st.lead;break;
      case O.nextLineShowText:st.x=st.lx;st.y=st.ly+=st.lead;show(a[0]);break;
      case O.nextLineSetSpacingShowText:st.ws=a[0];st.cs=a[1];st.x=st.lx;st.y=st.ly+=st.lead;show(a[2]);break;
      case O.showText:case O.showSpacedText:show(a[0]);break}}
  const v1=page.getViewport({scale:1});return {glyphs:out,rules,width:Math.abs(v1.width),height:Math.abs(v1.height)}}
/* the text of one cell from its glyphs in visual (left to right) order */
function cellText(gs){const units=[];let last=null;
  for(const g of gs){if(last&&g.x-(last.x+last.w)>0.2*Math.max(g.s,last.s))units.push(' ');units.push(g.u);last=g}
  let s;
  if(!units.some(u=>AR.test(u)))s=units.join('');
  else{const ty=units.map(u=>AR.test(u)?'R':LTRC.test(u)?'L':'N');
    for(let i=0;i<ty.length;i++)if(ty[i]==='N'){let a=i-1;while(a>=0&&ty[a]==='N')a--;let b=i+1;while(b<ty.length&&ty[b]==='N')b++;ty[i]=a>=0&&b<ty.length&&ty[a]==='L'&&ty[b]==='L'?'L':'R'}
    const runs=[];for(let i=0;i<units.length;i++){const t=ty[i];if(runs.length&&runs[runs.length-1].t===t)runs[runs.length-1].u.push(units[i]);else runs.push({t,u:[units[i]]})}
    s=runs.reverse().map(r=>r.t==='L'?r.u.join(''):r.u.reverse().map(u=>u.length===1&&MIRROR[u]?MIRROR[u]:u).join('')).join('')}
  if(/[ﭐ-﷿ﹰ-﻿ﬀ-ﬆ]/.test(s))s=s.normalize('NFKC');
  return s.replace(/\s+/g,' ').trim()}
/* lines → pieces of text: a new piece where the text run changes (writers draw one run per cell) or at a wide gap */
const seg_x1=seg=>seg?seg.x1:0;
function pdfLines(glyphs,rules){const gs=glyphs.slice().sort((a,b)=>a.y-b.y||a.x-b.x);const lines=[];let cur=null;
  for(const g of gs){if(cur&&Math.abs(g.y-cur.y)<=0.4*Math.max(g.s,cur.s))cur.g.push(g);else{cur={y:g.y,s:g.s,g:[g]};lines.push(cur)}}
  return lines.map(l=>{l.g.sort((a,b)=>a.x-b.x);const mid=l.y-0.3*l.s;
    /* the cell borders that cross this line: text between two borders is one cell, whatever its gaps */
    const bx=[];(rules||[]).forEach(r=>{if(r.y0-1<=mid&&r.y1+1>=mid&&!bx.some(x=>Math.abs(x-r.x)<1.5))bx.push(r.x)});bx.sort((a,b)=>a-b);
    const cellOf=x=>{let k=0;while(k<bx.length&&bx[k]<x)k++;return k};const ruled=bx.length>=2;
    const segs=[];let seg=null;
    for(const g of l.g){const c=ruled?cellOf(g.x+g.w/2):0;const gap=g.x-seg_x1(seg),sz=seg?Math.max(g.s,seg.s):g.s;
      const same=seg&&(ruled&&c>0&&c<bx.length?seg.c===c:gap>=-0.5*sz&&(g.r===seg.r?gap<=0.9*sz:gap<=0.25*sz)&&(!ruled||seg.c===c));
      if(same){seg.g.push(g);seg.x1=Math.max(seg.x1,g.x+g.w);seg.s=Math.max(seg.s,g.s);seg.r=g.r}
      else{seg={x0:g.x,x1:g.x+g.w,s:g.s,r:g.r,c,g:[g]};segs.push(seg)}}
    segs.forEach(c=>c.t=cellText(c.g));return {y:l.y,s:l.s,bx,segs:segs.filter(c=>c.t)}}).filter(l=>l.segs.length)}
const isNumText=t=>/^[-+(]?[\d٠-٩۰-۹][\d٠-٩۰-۹,٬.٫\s]*\)?-?%?$/.test(t);
const headerLike=l=>l.segs.length>=3&&l.segs.filter(c=>!isNumText(c.t)).length>=0.6*l.segs.length;
const lineKey=l=>l.segs.map(c=>c.t).join(' ').replace(/\s+/g,' ');
const FOOT=/^(?:(?:page|صفحة|الصفحة)\s*[\d٠-٩]+(?:\s*(?:of|من|\/)\s*[\d٠-٩]+)?|[\d٠-٩]+\s*(?:\/|of|من)\s*[\d٠-٩]+|[\d٠-٩]{1,4})$/i;
/* numbers printed as text: thousands separators, Arabic-Indic digits, (negative) or trailing minus; codes with a leading zero stay text */
function pdfCell(t){const s=t.replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/٫/g,'.').replace(/٬/g,',');
  const z=s.replace(/\s/g,'');if(/^-?\d+(?:\.\d+)?E[+-]?\d+$/i.test(z))return Number(z); // Excel prints long numbers as 1.01083E+12
  const m=z.match(/^(\()?([-+])?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?(\))?(-)?$/);
  if(!m||(!!m[1]!==!!m[5]))return t;if(/^0\d/.test(m[3])&&!m[4])return t;
  const n=Number(m[3].replace(/,/g,'')+(m[4]||''));if(!isFinite(n))return t;return (m[1]||m[2]==='-'||m[6])?-n:n}
/* a group of lines that share one header → one sheet. Columns are the stretches of the page covered by the pieces of
   text of the table's lines (from its header down); the gaps no line crosses separate them. */
function pdfSheet(lines,pageW,head,name){
  const start=Math.max(0,head?lines.findIndex(l=>lineKey(l)===head):0);
  const multi=lines.slice(start).filter(l=>l.segs.length>=2);if(!multi.length)return {name,grid:lines.map(l=>[pdfCell(l.segs.map(c=>c.t).join(' '))]),text:true};
  const W=Math.ceil(Math.max(pageW,...multi.map(l=>l.segs[l.segs.length-1].x1)))+2;const cov=new Uint16Array(W);
  multi.forEach(l=>l.segs.forEach(c=>{if(c.x1-c.x0>0.45*pageW)return;for(let x=Math.max(0,Math.floor(c.x0));x<Math.min(W,Math.ceil(c.x1));x++)cov[x]++}));
  let cols=[];let a=-1;
  for(let x=0;x<=W;x++){const on=x<W&&cov[x]>0;if(on&&a<0)a=x;else if(!on&&a>=0){cols.push([a,x]);a=-1}}
  /* a ruled table: its columns are the spaces between the borders that most of its lines share */
  const ruledL=multi.filter(l=>l.bx.length>=2);
  if(ruledL.length>=0.5*multi.length){const bs=[];ruledL.forEach(l=>l.bx.forEach(x=>{const b=bs.find(b=>Math.abs(b.x-x)<1.5);if(b)b.n++;else bs.push({x,n:1})}));
    const xs=bs.filter(b=>b.n>=Math.max(2,0.3*ruledL.length)).map(b=>b.x).sort((a,b)=>a-b);
    const rc=[];for(let i=0;i+1<xs.length;i++)if(xs[i+1]-xs[i]>2&&multi.some(l=>l.segs.some(c=>{const m=(c.x0+c.x1)/2;return m>xs[i]&&m<xs[i+1]})))rc.push([xs[i],xs[i+1]]);
    if(rc.length>=2)cols=rc}
  if(!cols.length)cols.push([0,W]);
  let ar=0,lt=0;lines.forEach(l=>l.segs.forEach(c=>{if(AR.test(c.t))ar++;else if(/[A-Za-z]/.test(c.t))lt++}));const rtl=ar>=lt;
  const order=rtl?cols.map((_,i)=>cols.length-1-i):cols.map((_,i)=>i);const pos=new Map(order.map((ci,k)=>[ci,k]));
  const colOf=c=>{let best=-1,bo=0;cols.forEach(([x0,x1],i)=>{const o=Math.min(x1,c.x1)-Math.max(x0,c.x0);if(o>bo){bo=o;best=i}});
    if(best<0){const mid=(c.x0+c.x1)/2;let d=1e9;cols.forEach(([x0,x1],i)=>{const dd=Math.abs((x0+x1)/2-mid);if(dd<d){d=dd;best=i}})}return best};
  const grid=lines.map(l=>{const row=new Array(cols.length).fill(null);
    l.segs.forEach(c=>{const k=pos.get(colOf(c));row[k]=row[k]?row[k].concat(c.g):c.g.slice()});
    return row.map(gs=>gs?pdfCell(cellText(gs.sort((a,b)=>a.x-b.x))):'')});
  return {name,grid:trimGrid(grid),text:true}}
/* the header of a page: among its first lines, the one with the most cells that has no numbers in it */
const pageHeader=(lines,rep)=>{let h=null,hs=0;lines.slice(0,6).forEach(l=>{const sc=l.segs.length-(rep(l)?100:0);if(headerLike(l)&&!l.segs.some(c=>isNumText(c.t))&&(!h||sc>hs)){h=l;hs=sc}});return h};
async function readPdf(buf,base){let lib;try{lib=await pdfjs()}catch(e){throw fail('تعذّر تحميل قارئ ملفات PDF. تحقق من الاتصال ثم أعد المحاولة.','Could not load the PDF reader. Check the connection and try again.')}
  let doc;try{doc=await lib.getDocument({data:new Uint8Array(buf),isEvalSupported:false,disableFontFace:true,useSystemFonts:false,verbosity:0}).promise}
  catch(e){if(e&&e.name==='PasswordException')throw fail('ملف PDF محمي بكلمة مرور. افتحه واحفظه بدون كلمة مرور (أو صدّر التقرير بصيغة Excel) ثم أعد الاستيراد.','The PDF is password-protected. Save it without a password (or export the report as Excel) and import again.');throw damaged()}
  try{const n=Math.min(doc.numPages,PDF_MAX_PAGES);const pages=[];let chars=0;
    for(let i=1;i<=n;i++){const pg=await doc.getPage(i);const {glyphs,rules,width,height}=await pdfGlyphs(lib,pg);chars+=glyphs.length;pages.push({lines:pdfLines(glyphs,rules),width,height});pg.cleanup()}
    /* a scan or a phone photo saved as PDF: no text layer, so the pages are read by the text recogniser and reviewed */
    if(chars<8){if(!window.OCRImport)throw fail('ملف PDF هذا صورة ممسوحة ضوئيًا ولا يحتوي نصًا. صدّر التقرير من النظام بصيغة Excel ثم استورده.','This PDF is a scanned image with no text layer. Export the report as Excel and import that file.');
      return await OCRImport.readScannedPdf(doc,base)}
    /* running titles and footers: lines in the top or bottom margin that come back on several pages (page numbers aside) */
    const edge=(l,p)=>l.y<0.09*p.height||l.y>0.91*p.height;const shape=l=>l.segs.length<=4&&l.segs.reduce((n,c)=>n+c.t.length,0)<=40?lineKey(l).replace(/[\d٠-٩]+/g,'#'):lineKey(l);const seen={};
    pages.forEach(p=>new Set(p.lines.filter(l=>edge(l,p)).map(shape)).forEach(k=>seen[k]=(seen[k]||0)+1));
    const groups=[];let g=null;
    const PAGENO=/^(?:page|صفحة|الصفحة)\s*[\d٠-٩]+\s*(?:of|من|\/)\s*[\d٠-٩]+$/i;
    pages.forEach(p=>{const l0=p.lines.filter(l=>!(l.segs.length===1&&(PAGENO.test(l.segs[0].t)||(edge(l,p)&&FOOT.test(l.segs[0].t)))));
      const rep=l=>edge(l,p)&&seen[shape(l)]>1;const h=pageHeader(l0,rep),hk=h&&lineKey(h);const lines=l0.filter(l=>l===h||!rep(l));
      if(!g||(hk&&hk!==g.head)){g={head:hk,hasHead:false,lines:[],width:p.width};groups.push(g)}
      lines.forEach(l=>{if(g.head&&lineKey(l)===g.head){if(g.hasHead)return;g.hasHead=true}g.lines.push(l)});g.width=Math.max(g.width,p.width)});
    const out=groups.filter(x=>x.lines.length).map((x,i,all)=>pdfSheet(x.lines,x.width,x.head,all.length>1?base+' '+(i+1):base));
    if(doc.numPages>PDF_MAX_PAGES)console.warn('PDF import: only the first '+PDF_MAX_PAGES+' pages were read');
    return out}
  finally{try{doc.destroy()}catch(e){}}}

/* ---------- entry points ---------- */
const EXT=f=>(String(f&&f.name||'').match(/\.([a-z0-9]+)$/i)||['',''])[1].toLowerCase();
/* a scan is read (and reviewed) once: a dashboard that reads the same file twice gets the reviewed table again */
const SCANNED=new WeakMap();
async function readGrids(file){
  if(file&&typeof file==='object'&&SCANNED.has(file))return SCANNED.get(file).map(s=>Object.assign({},s,{grid:s.grid.map(r=>r.slice())}));
  const out=await readAny(file);if(out.some(s=>s.ocr)&&file&&typeof file==='object')SCANNED.set(file,out);return out}
async function readAny(file){
  if(!file||!file.size)throw fail('الملف فارغ.','The file is empty.');
  if(file.size>MAX)throw fail('الملف أكبر من 200 ميغابايت. قسّمه إلى ملفات أصغر.','The file is over 200 MB. Split it into smaller files.');
  const ext=EXT(file);const buf=await file.arrayBuffer();const u8=new Uint8Array(buf);const base=String(file.name||'sheet').replace(/\.[^.]+$/,'')||'sheet';
  let out;
  if(u8[0]===0x50&&u8[1]===0x4B){
    if(typeof DecompressionStream==='undefined')throw fail('المتصفح قديم ولا يدعم قراءة ملفات Excel. استخدم Chrome أو Edge حديثًا.','This browser cannot read Excel files. Use a recent Chrome or Edge.');
    let z;try{z=await unzip(buf)}catch(e){throw damaged()}
    const has=n=>z.names.includes(n);
    try{
      if(has('xl/workbook.xml'))out=await readXlsx(z);
      else if(has('xl/workbook.bin'))out=await readXlsb(z);
      else if(has('content.xml')&&/spreadsheet/.test((await z.text('mimetype'))||'spreadsheet'))out=readOdsXml(await z.text('content.xml'));
      else if(z.names.some(n=>/^Index\/.*\.iwa$/.test(n)))throw fail('ملفات Apple Numbers لا تُقرأ مباشرة. من Numbers اختر ملف ← تصدير إلى ← Excel، ثم استورد الملف الناتج.','Apple Numbers files cannot be read directly. In Numbers choose File → Export To → Excel, then import that file.');
      else if(has('word/document.xml')||has('ppt/presentation.xml')||(has('content.xml')))throw fail('هذا مستند نصي أو عرض تقديمي وليس جدول بيانات. اختر ملف الجدول (Excel أو WPS أو CSV).','This is a document or presentation, not a spreadsheet. Choose the spreadsheet file (Excel, WPS or CSV).');
      else throw unsupported();
    }catch(e){if(e.importError)throw e;console.error(e);throw damaged()}}
  else if(u8[0]===0xD0&&u8[1]===0xCF&&u8[2]===0x11&&u8[3]===0xE0){let c;try{c=cfbStreams(buf)}catch(e){throw damaged()}
    const lower=c.names.map(x=>x.toLowerCase());
    if(lower.includes('encryptedpackage')||lower.includes('encryptioninfo'))throw locked();
    const wb=c.get('Workbook')||c.get('Book');
    if(!wb){if(lower.includes('worddocument'))throw fail('هذا مستند Word وليس جدول بيانات.','This is a Word document, not a spreadsheet.');throw unsupported()}
    try{out=readBiff(wb)}catch(e){if(e.importError)throw e;console.error(e);throw damaged()}}
  else if(u8[0]===0x25&&u8[1]===0x50&&u8[2]===0x44&&u8[3]===0x46)out=await readPdf(buf,base);
  else{
    /* binary that is none of the above (images, archives…) */
    if(window.OCRImport&&OCRImport.isImage(u8))return await OCRImport.readImage(file);
    const bin=[[0x89,0x50,0x4E,0x47],[0xFF,0xD8,0xFF],[0x52,0x61,0x72,0x21],[0x37,0x7A,0xBC,0xAF],[0x1F,0x8B]];
    if(bin.some(m=>m.every((x,i)=>u8[i]===x)))throw unsupported();
    if(!wideText(u8)){let ctrl=0;const k=Math.min(u8.length,4000);for(let i=0;i<k;i++){const x=u8[i];if(x<9||(x>13&&x<32&&x!==27))ctrl++}if(ctrl>k*0.02)throw unsupported()}
    const text=decodeText(u8);
    const head=text.replace(/^﻿/,'').trimStart();
    if(ext==='json'||/^[\[{]/.test(head)&&ext!=='csv'&&ext!=='txt'&&ext!=='tsv')out=jsonGrids(head,base);
    else if(/^</.test(head)){
      if(/urn:schemas-microsoft-com:office:spreadsheet/.test(head.slice(0,4000)))out=spreadsheetMl(head);
      else if(/<office:document[\s>]/.test(head.slice(0,4000)))out=readOdsXml(head);
      else if(/<table[\s>]/i.test(head))out=htmlGrids(head);
      else throw unsupported()}
    else out=delimGrid(head,base)}
  out=(out||[]).filter(s=>s&&s.grid);
  if(!out.length||out.every(s=>!s.grid.length))throw fail('لم أجد بيانات في الملف: كل الأوراق فارغة.','No data found: every sheet in the file is empty.');
  return out}
/* text-sourced cells: numbers written as text become numbers (except codes with a leading zero); an apostrophe guard is removed */
const NUMS=/^-?\d+(\.\d+)?$/;
const cellIn=v=>{if(typeof v!=='string')return v;const s=v.trim().replace(/^'(?=[=+@\-])/,'');return NUMS.test(s)&&!/^-?0\d/.test(s)?Number(s):s};
const serialIso=n=>new Date(Date.UTC(1899,11,30)+Math.round(n*86400000)).toISOString().slice(0,10);
async function readWorkbook(file){const sheets=await readGrids(file);const names=[],rows={};
  sheets.forEach(s=>{let nm=s.name;while(rows[nm])nm+='_';const g=s.grid;const out=[];
    if(g.length){const h=g[0].map(x=>String(x&&typeof x==='object'?'':x).trim());g.slice(1).forEach(r=>{if(!r.some(x=>String(x&&typeof x==='object'?'1':x).trim()))return;const o={};
      h.forEach((k,i)=>{if(!k)return;let v=r[i]??'';if(v&&typeof v==='object'&&'serial' in v)v=serialIso(v.serial);else if(s.text)v=cellIn(v);o[k]=v});out.push(o)})}
    names.push(nm);rows[nm]=out});
  return {SheetNames:names,rows:n=>rows[n]||[]}}
window.FileImport={ACCEPT,LIST,readGrids,readWorkbook,decodeText,pdfjs,_:{readPdf,pdfCell,cellText,pdfGlyphs,pdfLines,readBiff,cfbStreams,readOdsXml,delimGrid,htmlGrids,spreadsheetMl,jsonGrids,isDateCode}};
/* the contracts and needs import in Settings (app.js) calls readWorkbook() */
window.readWorkbook=readWorkbook;
})();
