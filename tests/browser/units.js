/* Quantity conversions per international standards (units.js), no browser needed:
   ASTM D1250-04 / API MPMS 11.1-2004 Table 54B volume correction to 15 °C, Table 53B density to 15 °C,
   Table 56 tonnes in air, measured density per line, LPG outside 54B. Ends with «UNITS OK». */
'use strict';
const path=require('path');
global.window={};global.document={addEventListener(){}};
require(path.join(__dirname,'..','..','units.js'));
const F=window.FuelUnits;let bad=0;
const eq=(name,got,want,tol=0)=>{const ok=Math.abs(got-want)<=tol;if(!ok){bad++;console.log('FAIL',name,got,'≠',want)}};
/* Table 54B check values (5 decimals) */
[[835,30,0.98716],[740,30,0.98151],[800,35,0.98131],[970,40,0.98241],[890,25,0.99216],[835,0,1.01275],[1000,60,0.96945],[700,45,0.95953],[835,15,1]]
  .forEach(([d,t,w])=>eq(`vcf ${d}@${t}`,F.vcf(d,t),w));
/* 2004 equations stay within ±0.0004 of the 1980 closed form across the table range */
let m=0;for(let d=612;d<=1160;d+=3.1)for(let t=-18;t<=90;t+=2.5)m=Math.max(m,Math.abs(F.ctl54B(d,t)-F.vcf1980(d,t)));eq('2004 vs 1980',m<0.0005?1:0,1);
/* LPG is outside Table 54B: no correction unless a 54E factor is set */
eq('lpg 54B',F.vcf(540,30),1);eq('lpg volume',F.toM3(1000,'لتر','16291','غاز مسال',30).q,1);
F.state.lpgVcf=0.9655;eq('lpg 54E factor',F.toM3(1000,'لتر','16291','غاز مسال',30).q,0.9655,1e-9);F.state.lpgVcf=null;
/* litres at 30 °C → m³ at 15 °C */
eq('diesel litres',F.toM3(1000,'لتر','44347','ديزل',30).q,0.98716,1e-9);
/* Table 56: tonnes in air, and in vacuum */
eq('diesel t air',F.toM3(10,'طن','44347','ديزل').q,10000/(835-1.1),1e-9);
F.state.basis='vacuum';eq('diesel t vacuum',F.toM3(10,'طن','44347','ديزل').q,10000/835,1e-9);F.state.basis='air';
/* a measured density on the line wins (kg/l or kg/m³) */
eq('measured kg/l',F.toM3(10,'طن','44347','ديزل',null,0.84).d,840);eq('measured kg/m³',F.toM3(10,'طن','44347','ديزل',null,842.3).d,842.3);
/* Table 53B: observed density → density at 15 °C, round trip */
eq('d15 from observed',F.density15From(835*F.ctl54B(835,30),30),835,0.05);
/* round trip tonnes → litres → tonnes */
const c=F.convert({q:1000,u:'لتر',code:'44347',name:'ديزل',t:30});eq('convert tonnes',c.tonnes,0.98716*(835-1.1)/1000,1e-12);eq('calcVersion',c.calcVersion,2);
const back=F.toM3(c.tonnes,'طن','44347','ديزل').q;eq('round trip',back,c.v15_m3,1e-9);
/* every reference density lies inside its specification */
F.families.forEach(f=>{const s=F.spec[f[0]];eq('spec '+f[0],s&&f[3]>=s[0]&&f[3]<=s[1]?1:0,1)});
if(bad){console.log('UNITS FAILED',bad);process.exit(1)}console.log('UNITS OK');
