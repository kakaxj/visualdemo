const EXCEL_FILE = window.APP_CONFIG?.FISCAL_EXCEL_URL || "fiscal data.xlsx";
const EXCEL_LABEL = "fiscal data.xlsx";
const REQUIRED_SHEETS = [
  '01_预算口径脉冲','02_收支口径脉冲','03_财政收支缺口与新增债务融资',
  '04_一般公共预算收支缺口年内路径','05_一般公共预算收支增速差',
  '06_政府性基金与土地出让收入','07_财政性存款月度变动',
  '08_地方政府债券发行与净融资','09_地方政府债务余额','10_新增专项债发行进度'
];

const C = {ink:'#14181C', ink2:'#626C76', ink3:'#8D959D', rule:'#D3D9DF', zhu:'#A32B2B',
           zhuL:'#C98B84', indigo:'#1F4E79', indigoL:'#7FA0BE', brass:'#8A6E28'};
const SANS = getComputedStyle(document.documentElement).getPropertyValue('--sans');
const MONO = '"SF Mono","JetBrains Mono",Consolas,monospace';
let D = null;
const charts = {};

function cacheBust(url){
  const sep=url.includes('?')?'&':'?';
  return encodeURI(url)+sep+'v='+Date.now();
}

const num = v => {
  if(v===null || v===undefined || v==='') return null;
  const n=Number(v); return Number.isFinite(n)?n:null;
};
const fmt = (v,d=0)=> v==null ? '—' : Number(v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
const sgn = (v,d=2)=> v==null ? '—' : ((Number(v)>0?'+':'') + Number(v).toFixed(d));
const last = a => { for(let i=a.length-1;i>=0;i--) if(a[i]!=null) return a[i]; return null; };
const lastIdx = a => { for(let i=a.length-1;i>=0;i--) if(a[i]!=null) return i; return -1; };
const firstDate = (dates,a) => { for(let i=0;i<a.length;i++) if(a[i]!=null) return dates[i]; return null; };

function excelDateParts(v){
  if(v==null || v==='') return null;
  if(v instanceof Date && !isNaN(v)) return {y:v.getFullYear(),m:v.getMonth()+1,d:v.getDate()};
  if(typeof v==='number'){
    const p=XLSX.SSF.parse_date_code(v);
    if(p && p.y && p.m) return {y:p.y,m:p.m,d:p.d||1};
  }
  const s=String(v).trim();
  let m=s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.](\d{1,2}))?/);
  if(m) return {y:+m[1],m:+m[2],d:+(m[3]||1)};
  const dt=new Date(s);
  if(!isNaN(dt)) return {y:dt.getFullYear(),m:dt.getMonth()+1,d:dt.getDate()};
  return null;
}
function ym(v){const p=excelDateParts(v); return p?`${p.y}-${String(p.m).padStart(2,'0')}`:null;}
function ymd(v){const p=excelDateParts(v); return p?`${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`:null;}
function prevYearYM(s){if(!s) return null; return `${+s.slice(0,4)-1}-${s.slice(5,7)}`;}

function getRows(workbook,name){
  const sheet=workbook.Sheets[name];
  if(!sheet) throw new Error(`缺少 Sheet：${name}`);
  return XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null,blankrows:false});
}
function findHeader(rows,required){
  for(let i=0;i<Math.min(rows.length,20);i++){
    const row=(rows[i]||[]).map(x=>String(x??'').trim());
    if(required.every(x=>row.includes(x))) return i;
  }
  throw new Error(`未找到表头：${required.join('、')}`);
}
function tableObjects(workbook,name,required){
  const rows=getRows(workbook,name); const h=findHeader(rows,required);
  const headers=(rows[h]||[]).map(x=>String(x??'').trim());
  const out=[];
  for(let i=h+1;i<rows.length;i++){
    const r=rows[i]||[]; if(r.every(x=>x==null || x==='')) continue;
    const o={}; headers.forEach((k,j)=>{if(k) o[k]=r[j]??null;}); out.push(o);
  }
  return {rows,h,headers,data:out};
}
function align(union,dates,values){
  const m=new Map(); dates.forEach((d,i)=>{if(d) m.set(d,values[i]??null);});
  return union.map(d=>m.has(d)?m.get(d):null);
}
function yearDiff(dates,arr){
  const m=new Map(dates.map((d,i)=>[d,arr[i]]));
  return dates.map((d,i)=>{
    const a=arr[i], b=m.get(prevYearYM(d));
    return (a==null || b==null)?null:+(a-b).toFixed(6);
  });
}
function latestDate(dates,...arrs){
  let best=null;
  for(let i=0;i<dates.length;i++) if(arrs.some(a=>a && a[i]!=null)) best=dates[i];
  return best;
}

function normalizeWorkbook(workbook){
  const missing=REQUIRED_SHEETS.filter(s=>!workbook.Sheets[s]);
  if(missing.length) throw new Error('Excel 缺少结果 Sheet：'+missing.join('、'));

  // 01 预算口径
  const t01=tableObjects(workbook,'01_预算口径脉冲',['年份','广义安排占GDP (%)','预算口径脉冲 (pp)']);
  const budget=t01.data.map(r=>({
    year:num(r['年份']), ratio:num(r['广义安排占GDP (%)']), impulse:num(r['预算口径脉冲 (pp)']),
    deficit:num(r['赤字合计 (亿元)']), special:num(r['新增专项债 (亿元)']),
    treasury:num(r['特别国债合计 (亿元)']), total:num(r['广义财政安排 (亿元)'])
  })).filter(x=>x.year!=null).sort((a,b)=>a.year-b.year);

  // 02 收支口径脉冲
  const t02=tableObjects(workbook,'02_收支口径脉冲',['日期','净脉冲 (pp)','分子效应 (pp)','分母效应 (pp)','收支缺口率 (%)']);
  const d02=t02.data.map(r=>ym(r['日期']));
  const fi_gap0=t02.data.map(r=>num(r['净脉冲 (pp)']));
  const fi_num0=t02.data.map(r=>num(r['分子效应 (pp)']));
  const fi_den0=t02.data.map(r=>num(r['分母效应 (pp)']));
  const gap_ratio0=t02.data.map(r=>num(r['收支缺口率 (%)']));

  // 03 融资匹配
  const t03=tableObjects(workbook,'03_财政收支缺口与新增债务融资',['日期','收支缺口率 (%)','新增债务口径 (%)']);
  const d03=t03.data.map(r=>ym(r['日期']));
  const wedge_gap0=t03.data.map(r=>num(r['收支缺口率 (%)']));
  const new_ratio0=t03.data.map(r=>num(r['新增债务口径 (%)']));
  const govb_ratio0=t03.data.map(r=>{
    const n=num(r['政府债券净融资12M (亿元)']), g=num(r['12M滚动GDP (亿元)']);
    return (n==null || !g)?null:n/g*100;
  });

  // 05-09 月度序列
  const t05=tableObjects(workbook,'05_一般公共预算收支增速差',['日期','收支增速差 (pp)','收入累计同比 (%)','支出累计同比 (%)']);
  const d05=t05.data.map(r=>ym(r['日期']));
  const growth0=t05.data.map(r=>num(r['收支增速差 (pp)']));
  const rev0=t05.data.map(r=>num(r['收入累计同比 (%)']));
  const exp0=t05.data.map(r=>num(r['支出累计同比 (%)']));

  const t06=tableObjects(workbook,'06_政府性基金与土地出让收入',['日期','政府性基金收入累计同比 (%)','土地出让收入累计同比 (%)']);
  const d06=t06.data.map(r=>ym(r['日期']));
  const frev0=t06.data.map(r=>num(r['政府性基金收入累计同比 (%)']));
  const land0=t06.data.map(r=>num(r['土地出让收入累计同比 (%)']));

  const t07=tableObjects(workbook,'07_财政性存款月度变动',['日期','月度变动 (亿元)']);
  const d07=t07.data.map(r=>ym(r['日期']));
  const dep0=t07.data.map(r=>num(r['月度变动 (亿元)']));

  const t08=tableObjects(workbook,'08_地方政府债券发行与净融资',['日期','新增专项债 (亿元)','新增一般债 (亿元)','再融资债 (亿元)','净融资 (亿元)']);
  const d08=t08.data.map(r=>ym(r['日期']));
  const spe0=t08.data.map(r=>num(r['新增专项债 (亿元)']));
  const gen0=t08.data.map(r=>num(r['新增一般债 (亿元)']));
  const refi0=t08.data.map(r=>num(r['再融资债 (亿元)']));
  const net0=t08.data.map(r=>num(r['净融资 (亿元)']));

  const t09=tableObjects(workbook,'09_地方政府债务余额',['日期','债务余额 (亿元)','同比 (%)','债务余额/GDP (%)']);
  const d09=t09.data.map(r=>ym(r['日期']));
  const bal0=t09.data.map(r=>num(r['债务余额 (亿元)']));
  const baly0=t09.data.map(r=>num(r['同比 (%)']));
  const balratio0=t09.data.map(r=>num(r['债务余额/GDP (%)']));

  // 04 年内路径矩阵
  const rows04=getRows(workbook,'04_一般公共预算收支缺口年内路径');
  const h04=findHeader(rows04,['月份']);
  const hdr04=rows04[h04]||[];
  const pubYears=hdr04.slice(1).map(num).filter(y=>y && y>=2022);
  const pub_gap_by_year=Object.fromEntries(pubYears.map(y=>[String(y),Array(12).fill(null)]));
  for(let i=h04+1;i<rows04.length;i++){
    const m=num(rows04[i]?.[0]); if(!m || m<1 || m>12) continue;
    pubYears.forEach((y,j)=>{pub_gap_by_year[String(y)][m-1]=num(rows04[i]?.[j+1]);});
  }

  // 10 专项债发行进度 + 右侧年度额度摘要
  const rows10=getRows(workbook,'10_新增专项债发行进度');
  const h10=findHeader(rows10,['月份']);
  const hdr10=rows10[h10]||[];
  const spYears=hdr10.slice(1,6).map(num).filter(y=>y && y>=2022);
  const spCum=Object.fromEntries(spYears.map(y=>[String(y),Array(12).fill(null)]));
  for(let i=h10+1;i<rows10.length;i++){
    const m=num(rows10[i]?.[0]); if(!m || m<1 || m>12) continue;
    spYears.forEach((y,j)=>{spCum[String(y)][m-1]=num(rows10[i]?.[j+1]);});
  }
  const spSummary={};
  const yCol=hdr10.findIndex(x=>String(x??'').trim()==='年份');
  const qCol=hdr10.findIndex(x=>String(x??'').trim()==='年度额度 (亿元)');
  const iCol=hdr10.findIndex(x=>String(x??'').trim()==='最新累计发行 (亿元)');
  const pCol=hdr10.findIndex(x=>String(x??'').trim()==='完成率 (%)');
  if(yCol>=0 && qCol>=0){
    for(let i=h10+1;i<rows10.length;i++){
      const y=num(rows10[i]?.[yCol]); if(!y) continue;
      spSummary[String(y)]={quota:num(rows10[i]?.[qCol]),issued:num(rows10[i]?.[iCol]),progress:num(rows10[i]?.[pCol])};
    }
  }

  // Excel 中显式“数据截止”优先，否则取月度结果表最新日期。
  let asof=null;
  outer: for(let i=0;i<Math.min(rows10.length,8);i++) for(let j=0;j<(rows10[i]||[]).length;j++){
    if(String(rows10[i][j]??'').trim()==='数据截止'){
      asof=ymd(rows10[i]?.[j+1]); break outer;
    }
  }

  const union=[...new Set([...d02,...d03,...d05,...d06,...d07,...d08,...d09].filter(Boolean))].sort();
  const D={
    meta:{asof:asof || (latestDate(d09,bal0)||latestDate(d08,net0)||latestDate(d05,growth0))+'-01'},
    dates:union,
    budget,
    fi_gap:align(union,d02,fi_gap0), fi_num:align(union,d02,fi_num0), fi_den:align(union,d02,fi_den0), gap_ratio:align(union,d02,gap_ratio0),
    wedge_gap_ratio:align(union,d03,wedge_gap0), new_ratio:align(union,d03,new_ratio0), govb_ratio:align(union,d03,govb_ratio0),
    pub_growth_gap:align(union,d05,growth0), rev_yoy:align(union,d05,rev0), exp_yoy:align(union,d05,exp0),
    frev_yoy:align(union,d06,frev0), land_yoy:align(union,d06,land0), fdep_chg:align(union,d07,dep0),
    lgb_spe:align(union,d08,spe0), lgb_gen:align(union,d08,gen0), lgb_refi:align(union,d08,refi0), lgb_net:align(union,d08,net0),
    bal:align(union,d09,bal0), bal_yoy:align(union,d09,baly0), bal_ratio:align(union,d09,balratio0),
    pubYears,pub_gap_by_year,spYears,spCum,spSummary
  };
  D.fi_new=yearDiff(D.dates,D.new_ratio);
  D.fi_fin=yearDiff(D.dates,D.govb_ratio);
  return D;
}

function getChart(id){
  if(!charts[id]) charts[id]=echarts.init(document.getElementById(id),null,{renderer:'canvas'});
  return charts[id];
}
function base(o={}){return Object.assign({
  textStyle:{fontFamily:SANS,color:C.ink2},animationDuration:650,
  grid:{left:8,right:8,top:34,bottom:4,containLabel:true},
  tooltip:{trigger:'axis',backgroundColor:'#fff',borderColor:C.rule,borderWidth:1,textStyle:{color:C.ink,fontSize:12,fontFamily:SANS},extraCssText:'box-shadow:0 2px 10px rgba(0,0,0,.09)'},
  legend:{top:0,left:0,itemWidth:14,itemHeight:2,itemGap:16,textStyle:{fontSize:11.5,color:C.ink2,fontFamily:SANS}}
},o);}
function valAxis(o={}){return Object.assign({type:'value',axisLine:{show:false},axisTick:{show:false},splitLine:{lineStyle:{color:C.rule,type:[3,4]}},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO}},o);}
function zero(){return {symbol:'none',silent:true,lineStyle:{color:C.ink,width:1},data:[{yAxis:0,label:{show:false}}]};}
const pct=v=>v+'%';
function shadowTip(extra){return Object.assign({trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'#fff',borderColor:C.rule,borderWidth:1,textStyle:{color:C.ink,fontSize:12,fontFamily:SANS}},extra||{});}
function lgd(data){return {top:0,left:0,itemWidth:14,itemHeight:2,itemGap:16,data,textStyle:{fontSize:11.5,color:C.ink2,fontFamily:SANS}};}

const DISPLAY_START='2022-01';
function win(...arrs){
  let s=D.dates.length,e=-1;
  for(const a of arrs){for(let i=0;i<a.length;i++) if(a[i]!=null){s=Math.min(s,i);break;} for(let i=a.length-1;i>=0;i--) if(a[i]!=null){e=Math.max(e,i);break;}}
  const ds=Math.max(0,D.dates.indexOf(DISPLAY_START)); s=Math.max(s===D.dates.length?ds:s,ds); if(e<s)e=D.dates.length-1;
  const dates=D.dates.slice(s,e+1); const span=(+dates.at(-1).slice(0,4))-(+dates[0].slice(0,4));
  return {dates,cut:a=>a.slice(s,e+1),step:span>10?2:1,from:D.dates[s]};
}
function winAll(...arrs){
  let s=0,e=D.dates.length-1;
  for(const a of arrs){for(let i=0;i<a.length;i++) if(a[i]!=null){s=Math.max(s,i);break;} for(let i=a.length-1;i>=0;i--) if(a[i]!=null){e=Math.min(e,i);break;}}
  const ds=Math.max(0,D.dates.indexOf(DISPLAY_START)); s=Math.max(s,ds); if(e<s)e=D.dates.length-1;
  const dates=D.dates.slice(s,e+1); const span=(+dates.at(-1).slice(0,4))-(+dates[0].slice(0,4));
  return {dates,cut:a=>a.slice(s,e+1),step:span>10?2:1,from:D.dates[s]};
}
function timeAxis(W,o={}){return Object.assign({type:'category',data:W.dates,boundaryGap:false,axisLine:{lineStyle:{color:C.rule}},axisTick:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,interval:(i,v)=>v.slice(5)==='01' && (+v.slice(0,4))%W.step===0,formatter:v=>v.slice(0,4)}},o);}
function bridge(a){const o=a.map(()=>null);for(let i=1;i<a.length-1;i++) if(a[i]==null&&a[i-1]!=null&&a[i+1]!=null){o[i-1]=a[i-1];o[i]=(a[i-1]+a[i+1])/2;o[i+1]=a[i+1];}return o;}
function bridgeSeries(name,arr,color){return {name,type:'line',data:bridge(arr),symbol:'none',connectNulls:false,silent:true,legendHoverLink:false,emphasis:{disabled:true},lineStyle:{width:1.2,color,type:'dashed',opacity:.85},itemStyle:{color},z:2,tooltip:{show:false,trigger:'none'}};}

function renderHeaderAndCards(){
  const asofYM=D.meta.asof.slice(0,7);
  document.getElementById('asof').textContent=asofYM.replace('-', '年')+'月';
  const B=D.budget, bN=B.at(-1), bP=B.at(-2), gi=lastIdx(D.gap_ratio), ni=lastIdx(D.new_ratio);
  const rd=[
    {k:`预算口径脉冲 · ${bN?.year||'—'}年安排`,v:bN?sgn(bN.impulse):'—',u:'pp',neg:bN?.impulse<0,d:bN&&bP?`广义财政安排 ${(bN.total/10000).toFixed(2)}万亿 ← ${(bP.total/10000).toFixed(2)}万亿`:'—'},
    {k:'收支口径脉冲 · 12个月滚动',v:sgn(last(D.fi_gap)),u:'pp',neg:last(D.fi_gap)<0,d:`分子 ${sgn(last(D.fi_num))} · 分母 ${sgn(last(D.fi_den))}`},
    {k:'新增债务口径脉冲 · 12个月滚动',v:sgn(last(D.fi_new)),u:'pp',neg:last(D.fi_new)<0,d:ni>=0?`新增债务口径 ${fmt(D.new_ratio[ni],2)}% · 上年同期 ${fmt(D.new_ratio[D.dates.indexOf(prevYearYM(D.dates[ni]))],2)}%`:'—'},
    {k:'收支口径缺口率 · 12个月滚动',v:gi>=0?fmt(D.gap_ratio[gi],2):'—',u:'%',neg:false,d:gi>=0?`上年同期 ${fmt(D.gap_ratio[D.dates.indexOf(prevYearYM(D.dates[gi]))],2)}%`:'—'}
  ];
  document.getElementById('readout').innerHTML=rd.map(r=>`<div class="rd"><div class="k">${r.k}</div><div class="v ${r.neg?'neg':''}">${r.v}<span class="u">${r.u}</span></div><div class="d">${r.d}</div></div>`).join('');

  const y=+asofYM.slice(0,4), m=+asofYM.slice(5,7);
  const gapNow=D.pub_gap_by_year[String(y)]?.[m-1]??null, gapPrev=D.pub_gap_by_year[String(y-1)]?.[m-1]??null;
  const idx=D.dates.indexOf(asofYM), growth=idx>=0?D.pub_growth_gap[idx]:null, rev=idx>=0?D.rev_yoy[idx]:null, exp=idx>=0?D.exp_yoy[idx]:null, dep=idx>=0?D.fdep_chg[idx]:null;
  const sp=D.spSummary[String(y)]||{};
  const execRd=[
    {k:'一般公共预算收支缺口 · 年初至今',v:gapNow==null?'—':(gapNow/10000).toFixed(2),u:'万亿',neg:false,d:`上年同期 ${gapPrev==null?'—':(gapPrev/10000).toFixed(2)+'万亿'}`},
    {k:'一般公共预算收支增速差',v:growth==null?'—':sgn(growth,1),u:'pp',neg:growth!=null&&growth<0,d:`支出 ${exp==null?'—':sgn(exp,1)+'%'} · 收入 ${rev==null?'—':sgn(rev,1)+'%'}`},
    {k:'财政性存款 · 当月变动',v:dep==null?'—':sgn(dep/10000,2),u:'万亿',neg:dep!=null&&dep<0,d:dep==null?'—':(dep<0?'负值对应财政资金投放':'正值对应财政资金回笼')},
    {k:'新增专项债发行进度',v:sp.progress==null?'—':fmt(sp.progress,1),u:'%',neg:false,d:`${sp.issued==null?'—':(sp.issued/10000).toFixed(2)}万亿 / 额度 ${sp.quota==null?'—':(sp.quota/10000).toFixed(2)}万亿`}
  ];
  document.getElementById('exec_readout').innerHTML=execRd.map(r=>`<div class="rd"><div class="k">${r.k}</div><div class="v ${r.neg?'neg':''}">${r.v}<span class="u">${r.u}</span></div><div class="d">${r.d}</div></div>`).join('');
}

function renderSummary(){
  const el=document.getElementById('summaryText');
  if(!el || !D) return;

  const asofYM=D.meta.asof.slice(0,7), y=+asofYM.slice(0,4), m=+asofYM.slice(5,7);
  const zhDate=`${y}年${m}月`;
  const B=D.budget||[], bN=B.at(-1);
  const fi=last(D.fi_gap), gap=last(D.gap_ratio);
  const idx=D.dates.indexOf(asofYM);
  const growth=idx>=0?D.pub_growth_gap[idx]:null;
  const rev=idx>=0?D.rev_yoy[idx]:null, exp=idx>=0?D.exp_yoy[idx]:null;
  const dep=idx>=0?D.fdep_chg[idx]:null;
  const gapNow=D.pub_gap_by_year[String(y)]?.[m-1]??null;
  const sp=D.spSummary[String(y)]||{};

  const budgetTone=bN?.impulse==null?'预算安排保持稳定':
    bN.impulse>0.2?'预算安排较上年进一步扩张':
    bN.impulse<-0.2?'预算安排力度较上年有所回落，但广义财政安排仍维持较高规模':'预算安排总体保持稳定';
  const execTone=fi==null?'执行端保持平稳':
    fi>0.2?'收支口径脉冲为正，执行端支持力度较上年增强':
    fi<-0.2?'收支口径脉冲较上年同期有所回落':'收支口径脉冲总体平稳';
  const growthTone=growth==null?'':(growth>0.3?'支出增速快于收入':growth<-0.3?'收入增速快于支出':'收支增速大体接近');
  const depTone=dep==null?'':(dep<0?'当月财政性存款下降，体现资金投放':'当月财政性存款上升，体现阶段性资金回笼');

  let spTone='';
  if(sp.progress!=null){
    const py=String(y-1), prevCum=D.spCum?.[py]?.[m-1], prevQuota=D.spSummary?.[py]?.quota;
    const prevProg=(prevCum!=null&&prevQuota)?prevCum/prevQuota*100:null;
    if(prevProg!=null){
      const diff=sp.progress-prevProg;
      spTone=diff>3?'专项债发行进度快于上年同期':diff<-3?'专项债发行进度慢于上年同期':'专项债发行进度与上年同期大体接近';
    }else spTone='专项债发行继续推进';
  }

  const s1=`截至 <strong>${zhDate}</strong>，财政政策整体呈现“${budgetTone}、${execTone}”的组合特征。`;
  const metrics=[];
  if(bN?.impulse!=null) metrics.push(`${y}年预算口径脉冲为 <span class="key">${sgn(bN.impulse,2)}pp</span>${bN.total!=null?`，广义财政安排 <strong>${(bN.total/10000).toFixed(2)}万亿元</strong>`:''}`);
  if(fi!=null) metrics.push(`12个月滚动收支口径脉冲为 <span class="key">${sgn(fi,2)}pp</span>${gap!=null?`，收支缺口率 <strong>${fmt(gap,2)}%</strong>`:''}`);
  const s2=metrics.length?metrics.join('；')+'。':'';

  const exec=[];
  if(gapNow!=null) exec.push(`一般公共预算累计收支缺口 <strong>${(gapNow/10000).toFixed(2)}万亿元</strong>`);
  if(growthTone) exec.push(`${growthTone}${growth!=null?`（收支增速差 <span class="key">${sgn(growth,1)}pp</span>）`:''}`);
  if(sp.progress!=null) exec.push(`新增专项债累计发行 ${sp.issued!=null?`<strong>${(sp.issued/10000).toFixed(2)}万亿元</strong>，`:''}完成年度额度 <strong>${fmt(sp.progress,1)}%</strong>${spTone?'，'+spTone:''}`);
  if(depTone) exec.push(depTone);
  const s3=exec.length?'执行层面，'+exec.join('；')+'。':'';

  el.innerHTML=[s1,s2,s3].filter(Boolean).join(' ');
}

function renderCharts(){
  const Bchart=D.budget.filter(b=>b.year>=2022);
  getChart('c_bud').setOption(base({
    legend:{show:false},tooltip:shadowTip({formatter:p=>{const y=p[0].axisValue,b=Bchart.find(x=>String(x.year)===y);return `<b>${y}年</b><br/>广义安排 ${fmt(b.total)} 亿元<br/>赤字 ${fmt(b.deficit)}　专项债 ${fmt(b.special)}　特别国债 ${fmt(b.treasury)}<br/>占GDP ${fmt(b.ratio,2)}%　脉冲 ${sgn(b.impulse)} pp`;}}),
    xAxis:{type:'category',data:Bchart.map(b=>String(b.year)),boundaryGap:true,axisLine:{lineStyle:{color:C.rule}},axisTick:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO}},
    yAxis:[valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}}),valAxis({splitLine:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}})],
    series:[{name:'脉冲',type:'bar',data:Bchart.map(b=>b.impulse),barMaxWidth:36,itemStyle:{color:p=>p.value>0?C.zhu:C.indigo},markLine:zero(),label:{show:true,position:'top',fontFamily:MONO,fontSize:10.5,color:C.ink2,formatter:p=>p.value==null?'':sgn(p.value)}},{name:'广义安排占GDP',type:'line',yAxisIndex:1,data:Bchart.map(b=>b.ratio),symbol:'circle',symbolSize:5,lineStyle:{width:1.6,color:C.brass,type:'dashed'},itemStyle:{color:C.brass}}]
  }),true);

  const Wd=win(D.fi_num,D.fi_den,D.fi_gap);
  getChart('c_dec').setOption(base({legend:lgd(['分子效应','分母效应','净脉冲']),tooltip:shadowTip(),xAxis:timeAxis(Wd,{boundaryGap:true}),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}}),series:[
    {name:'分子效应',type:'bar',stack:'a',data:Wd.cut(D.fi_num),itemStyle:{color:C.zhu},barMaxWidth:7},
    {name:'分母效应',type:'bar',stack:'a',data:Wd.cut(D.fi_den),itemStyle:{color:C.indigoL},barMaxWidth:7},
    {name:'净脉冲',type:'line',data:Wd.cut(D.fi_gap),connectNulls:false,symbol:'none',lineStyle:{width:1.8,color:C.ink},itemStyle:{color:C.ink},markLine:zero(),z:5},
    bridgeSeries('净脉冲衔接',Wd.cut(D.fi_gap),C.ink)
  ]}),true);

  const Ww=winAll(D.new_ratio,D.wedge_gap_ratio);
  getChart('c_wedge').setOption(base({legend:lgd(['新增债务口径','收支缺口']),tooltip:shadowTip(),xAxis:timeAxis(Ww),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}}),series:[
    {name:'新增债务口径',type:'line',data:Ww.cut(D.new_ratio),symbol:'none',connectNulls:false,lineStyle:{width:1.8,color:C.indigo,type:'dashed'},itemStyle:{color:C.indigo},z:6},
    {name:'收支缺口',type:'line',data:Ww.cut(D.wedge_gap_ratio),symbol:'none',connectNulls:false,lineStyle:{width:2,color:C.zhu},itemStyle:{color:C.zhu},markLine:zero(),z:9},
    bridgeSeries('收支缺口衔接',Ww.cut(D.wedge_gap_ratio),C.zhu)
  ]}),true);

  const grey=['#CBD0D5','#B9C0C7','#A7B0B9'];
  const months=Array.from({length:11},(_,i)=>(i+2)+'月');
  getChart('c_pub_gap_path').setOption(base({legend:{top:0,left:0,itemWidth:16,itemHeight:2,itemGap:13,textStyle:{fontSize:11.5,color:C.ink2,fontFamily:MONO}},grid:{left:8,right:28,top:36,bottom:4,containLabel:true},tooltip:{trigger:'axis',axisPointer:{type:'line'},backgroundColor:'#fff',borderColor:C.rule,borderWidth:1,textStyle:{color:C.ink,fontSize:12,fontFamily:SANS},formatter:params=>{if(!params?.length)return'';let x='<b>'+params[0].axisValue+'</b>';params.filter(p=>p.value!=null).forEach(p=>x+=`<br/>${p.marker}${p.seriesName}　<b>${(p.value/10000).toFixed(2)}</b> 万亿元`);return x;}},xAxis:{type:'category',data:months,boundaryGap:false,axisLine:{lineStyle:{color:C.rule}},axisTick:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO}},yAxis:valAxis({scale:true,splitNumber:5,axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>(v/10000).toFixed(1)+'万亿'}}),series:D.pubYears.map((y,i)=>{const now=y===Math.max(...D.pubYears),prev=y===Math.max(...D.pubYears)-1,col=now?C.zhu:(prev?C.indigo:grey[i%grey.length]);return {name:String(y),type:'line',data:D.pub_gap_by_year[String(y)].slice(1),connectNulls:false,symbol:'none',lineStyle:{width:now?2.6:(prev?2:1.2),color:col},itemStyle:{color:col},z:now?9:(prev?7:3),emphasis:{focus:'series'},markLine:now?zero():undefined};})}),true);

  const Wgg=win(D.pub_growth_gap);
  getChart('c_pub_growth_gap').setOption(base({legend:{show:false},tooltip:shadowTip({formatter:p=>{const x=p&&p[0];return(!x||x.value==null)?'':`<b>${x.axisValue}</b><br/>支出增速 − 收入增速　<b>${sgn(x.value,1)}</b> pp`;}}),xAxis:timeAxis(Wgg),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>v+'pp'}}),series:[{name:'收支增速差',type:'line',data:Wgg.cut(D.pub_growth_gap),connectNulls:false,symbol:'none',lineStyle:{width:2,color:C.zhu},itemStyle:{color:C.zhu},areaStyle:{color:'rgba(163,43,43,.06)'},markLine:zero(),z:4},bridgeSeries('增速差衔接',Wgg.cut(D.pub_growth_gap),C.zhu)]}),true);

  const Wf=win(D.frev_yoy,D.land_yoy);
  getChart('c_fund').setOption(base({legend:lgd(['政府性基金收入','国有土地出让收入']),xAxis:timeAxis(Wf),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}}),series:[{name:'政府性基金收入',type:'line',data:Wf.cut(D.frev_yoy),connectNulls:false,symbol:'none',lineStyle:{width:1.8,color:C.zhu},itemStyle:{color:C.zhu},markLine:zero()},{name:'国有土地出让收入',type:'line',data:Wf.cut(D.land_yoy),connectNulls:false,symbol:'none',lineStyle:{width:1.6,color:C.brass,type:'dashed'},itemStyle:{color:C.brass}},bridgeSeries('基金衔接',Wf.cut(D.frev_yoy),C.zhu),bridgeSeries('土地衔接',Wf.cut(D.land_yoy),C.brass)]}),true);

  const Wp=win(D.fdep_chg);
  getChart('c_dep').setOption(base({legend:{show:false},tooltip:shadowTip({formatter:p=>p?.[0]?p[0].axisValue+'　<b>'+fmt(p[0].value)+'</b> 亿元':''}),xAxis:timeAxis(Wp,{boundaryGap:true}),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>fmt(v)}}),series:[{name:'月度变动',type:'bar',data:Wp.cut(D.fdep_chg),barMaxWidth:6,itemStyle:{color:p=>p.value>0?C.indigo:C.zhu},markLine:zero()}]}),true);

  const Wl=win(D.lgb_spe,D.lgb_gen,D.lgb_refi,D.lgb_net);
  getChart('c_lgb').setOption(base({legend:lgd(['新增专项债','新增一般债','再融资债','净融资']),tooltip:shadowTip({formatter:params=>{if(!params?.length)return'';let s='<b>'+params[0].axisValue+'</b>';params.filter(p=>p.value!=null).forEach(p=>s+=`<br/>${p.marker}${p.seriesName}　<b>${fmt(p.value)}</b> 亿元`);const vals=Object.fromEntries(params.filter(p=>p.value!=null).map(p=>[p.seriesName,p.value]));if(vals['新增专项债']!=null&&vals['新增一般债']!=null&&vals['再融资债']!=null)s+=`<br/><span style="color:${C.ink3}">当月发行合计　${fmt(vals['新增专项债']+vals['新增一般债']+vals['再融资债'])} 亿元</span>`;return s;}}),xAxis:timeAxis(Wl,{boundaryGap:true}),yAxis:valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>fmt(v)}}),series:[{name:'新增专项债',type:'bar',stack:'a',data:Wl.cut(D.lgb_spe),itemStyle:{color:C.zhu},barMaxWidth:8},{name:'新增一般债',type:'bar',stack:'a',data:Wl.cut(D.lgb_gen),itemStyle:{color:C.brass},barMaxWidth:8},{name:'再融资债',type:'bar',stack:'a',data:Wl.cut(D.lgb_refi),itemStyle:{color:C.indigoL},barMaxWidth:8},{name:'净融资',type:'line',data:Wl.cut(D.lgb_net),connectNulls:false,symbol:'none',lineStyle:{width:1.7,color:C.ink},itemStyle:{color:C.ink},markLine:zero(),z:6}]}),true);

  const Wb=winAll(D.bal,D.bal_yoy);
  getChart('c_bal').setOption(base({legend:lgd(['债务余额','同比']),tooltip:{trigger:'axis',axisPointer:{type:'line'},backgroundColor:'#fff',borderColor:C.rule,borderWidth:1,textStyle:{color:C.ink,fontSize:12,fontFamily:SANS},formatter:params=>{if(!params?.length)return'';const d=params[0].axisValue,idx=D.dates.indexOf(d),b=D.bal[idx],y=D.bal_yoy[idx],r=D.bal_ratio[idx];let s='<b>'+d+'</b>';if(b!=null)s+=`<br/>债务余额　<b>${(b/10000).toFixed(2)}</b> 万亿元`;if(y!=null)s+=`<br/>同比　<b>${sgn(y,1)}</b>%`;if(r!=null)s+=`<br/>余额 / GDP　<b>${fmt(r,1)}</b>%`;return s;}},xAxis:timeAxis(Wb),yAxis:[valAxis({axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>(v/10000).toFixed(0)+'万亿'}}),valAxis({splitLine:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:pct}})],series:[{name:'债务余额',type:'line',data:Wb.cut(D.bal),symbol:'none',lineStyle:{width:1.8,color:C.indigo},itemStyle:{color:C.indigo},areaStyle:{color:'rgba(31,78,121,.09)'}},{name:'同比',type:'line',yAxisIndex:1,data:Wb.cut(D.bal_yoy),connectNulls:false,symbol:'none',lineStyle:{width:1.4,color:C.zhu,type:'dashed'},itemStyle:{color:C.zhu}}]}),true);

  getChart('c_sp_progress').setOption(base({legend:{top:0,left:0,itemWidth:16,itemHeight:2,itemGap:13,textStyle:{fontSize:11.5,color:C.ink2,fontFamily:MONO}},grid:{left:8,right:40,top:36,bottom:4,containLabel:true},tooltip:{trigger:'axis',axisPointer:{type:'line'},backgroundColor:'#fff',borderColor:C.rule,borderWidth:1,textStyle:{color:C.ink,fontSize:12,fontFamily:SANS},formatter:params=>{if(!params?.length)return'';let s='<b>'+params[0].axisValue+'</b>';params.filter(p=>p.value!=null).forEach(p=>{const q=D.spSummary[p.seriesName]?.quota,prog=q?p.value/q*100:null;s+=`<br/>${p.marker}${p.seriesName}　<b>${(p.value/10000).toFixed(2)}</b> 万亿元${q?' ／ 额度 '+(q/10000).toFixed(2)+' 万亿元':''}${prog==null?'':'　'+prog.toFixed(1)+'%'}`;});return s;}},xAxis:{type:'category',data:Array.from({length:12},(_,i)=>(i+1)+'月'),boundaryGap:false,axisLine:{lineStyle:{color:C.rule}},axisTick:{show:false},axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO}},yAxis:valAxis({scale:true,min:0,splitNumber:5,axisLabel:{fontSize:11,color:C.ink3,fontFamily:MONO,formatter:v=>(v/10000).toFixed(1)+'万亿'}}),series:D.spYears.map((y,i)=>{const maxY=Math.max(...D.spYears),now=y===maxY,prev=y===maxY-1,col=now?C.zhu:(prev?C.indigo:grey[i%grey.length]);return {name:String(y),type:'line',data:D.spCum[String(y)],connectNulls:false,symbol:'none',lineStyle:{width:now?2.6:(prev?2:1.2),color:col},itemStyle:{color:col},z:now?9:(prev?7:3)};})}),true);
}

function renderMethods(){
  const DEFS=[
    ['D1','广义收支缺口 D','D = （一般公共预算支出 + 政府性基金支出）\n    −（一般公共预算收入 + 政府性基金收入）','两本账合计口径。该计算在 Excel 结果 Sheet 中完成。'],
    ['D2','12 个月滚动值','X₁₂ₘ(y, m) = X累计(y, m) + X累计(y−1, 12) − X累计(y−1, m)','由年初至今累计值直接构造，无需先还原单月值。该计算在 Excel 中完成。'],
    ['D3','收支口径缺口率 d','d(t) = D₁₂ₘ(t) ÷ Y₁₂ₘ(t)','Y 为名义 GDP 的 12 个月滚动值。'],
    ['D4','财政脉冲 FI','FI(t) = d(t) − d(t−12)','同比变化，单位为百分点。'],
    ['D5','分子分母分解','FI(t) = ΔD ÷ Y(t)  −  d(t−12) × ΔY ÷ Y(t)\n         分子效应          分母效应\n\nΔD = D(t) − D(t−12)　　ΔY = Y(t) − Y(t−12)','恒等式，两项之和等于 FI，无残差项。'],
    ['D6','融资口径','政府债券净融资率 = 政府债券净融资₁₂ₘ ÷ Y₁₂ₘ\n净再融资 = 地方再融资债发行 − 地方债到期偿还\n新增债务口径 = 政府债券净融资 − 净再融资','核心融资口径均由 Excel 结果 Sheet 提供；网页仅用于展示。'],
    ['D7','预算口径（事前）','广义安排 = 中央赤字 + 地方赤字 + 新增专项债 + 特别国债\n预算口径脉冲 = 本年广义安排占GDP比重 − 上年比重','取自历年预算报告；结果直接读取 Excel 的 01_预算口径脉冲。']
  ];
  document.getElementById('defs').innerHTML=DEFS.map(([t,n,e,c])=>`<div class="def"><div class="tag">${t}</div><div class="body"><div class="nm">${n}</div><div class="eq">${e}</div><div class="cm">${c}</div></div></div>`).join('');
  const NOTES=[
    `网页以 ${EXCEL_LABEL} 为唯一业务数据源，直接读取 01—10 结果 Sheet。核心财政计算在 Excel 中完成，HTML 不内嵌固定业务数据。`,
    `覆盖同目录下同名 Excel 后点击“刷新数据”或刷新浏览器即可更新。若直接双击 HTML，浏览器可能限制读取本地同目录文件，可使用“选择 Excel”手动载入；部署到网站或本地服务器后可自动读取。`,
    `当前数据截止 ${D.meta.asof}。图表统一自 2022 年起展示；缺失值保留为空，不补零、不参与计算。个别月度序列的虚线衔接仅作视觉连接。`,
    `一般公共预算收支缺口年内路径、收支增速差、政府性基金与土地出让收入、财政性存款、地方债发行和债务余额均直接取对应结果 Sheet。`,
    `融资口径采用政府债券净融资，并在 Excel 中剔除地方净再融资的近似影响形成新增债务口径；网页的“新增债务口径脉冲”仅对该已算好的比率做同月同比差。`,
    `地方政府债券发行按新增专项债、新增一般债和再融资债拆分；净融资为当月发行额减到期偿还本金。`,
    `新增专项债发行进度读取 10_新增专项债发行进度中的年内累计发行与年度额度；完成率主要用于比较发行节奏。`,
    `地方政府债务余额图的余额/GDP直接读取 09_地方政府债务余额中的核查列，不在网页端重新估算 GDP。`,
    `数据来源与计算底稿以 ${EXCEL_LABEL} 内各结果 Sheet 的说明为准；更新 Wind 底稿并保存 Excel 后，需确保 Excel 公式结果已重新计算并保存。`
  ];
  document.getElementById('notes').innerHTML=NOTES.map((f,i)=>`<div class="note"><div class="n">${i+1}</div><div>${f}</div></div>`).join('');
}

const noLegendCharts=new Set(['c_bud','c_pub_growth_gap','c_dep']);
function applyResponsiveCharts(){
  const mobile=innerWidth<=600,tablet=innerWidth<=1024;
  Object.entries(charts).forEach(([id,c])=>{
    const noLegend=noLegendCharts.has(id); const desktop=(id==='c_pub_gap_path'?{left:8,right:28,top:36,bottom:4}:id==='c_sp_progress'?{left:8,right:40,top:36,bottom:4}:{left:8,right:8,top:34,bottom:4});
    const grid=mobile?{left:4,right:4,top:noLegend?18:52,bottom:6,containLabel:true}:tablet?{left:6,right:6,top:noLegend?28:42,bottom:4,containLabel:true}:Object.assign({},desktop,{containLabel:true});
    c.setOption({grid,tooltip:{confine:mobile}},false);
    if(mobile&&(id==='c_pub_gap_path'||id==='c_sp_progress')) c.setOption({xAxis:{axisLabel:{interval:1,fontSize:10,color:C.ink3,fontFamily:MONO}}},false);
    else if(!mobile&&(id==='c_pub_gap_path'||id==='c_sp_progress')) c.setOption({xAxis:{axisLabel:{interval:0,fontSize:11,color:C.ink3,fontFamily:MONO}}},false);
    c.resize();
  });
}

function renderAll(){
  if(!D) return;
  renderHeaderAndCards(); renderSummary(); renderCharts(); renderMethods(); applyResponsiveCharts();
}
function showStatus(type,title,text,showPicker=false){
  const box=document.getElementById('statusBox'); box.className=`status-box show ${type||''}`;
  document.getElementById('statusTitle').textContent=title; document.getElementById('statusText').textContent=text;
  document.getElementById('fileLabel').style.display=showPicker?'inline-block':'none';
}
function hideStatus(){document.getElementById('statusBox').className='status-box';}
function hideLoading(){document.getElementById('loadingMask').classList.add('hidden');}
async function useWorkbook(workbook, label) {
    D = normalizeWorkbook(workbook);
    renderAll();
}
async function loadExcel(){
  document.getElementById('loadingMask').classList.remove('hidden'); hideStatus();
  try{
    if(typeof XLSX==='undefined') throw new Error('Excel 读取组件未加载，请检查网络连接');
    if(typeof echarts==='undefined') throw new Error('图表组件未加载，请检查网络连接');
    const res=await fetch(cacheBust(EXCEL_FILE),{cache:'no-store'});
    if(!res.ok) throw new Error(`读取 ${EXCEL_LABEL} 失败（HTTP ${res.status}）`);
    const buf=await res.arrayBuffer(); const workbook=XLSX.read(buf,{type:'array',cellDates:false});
    await useWorkbook(workbook,'Excel 已自动读取');
  }catch(err){
    console.error(err); showStatus('error','未能自动读取 Excel',`${err.message}。如果你是直接双击打开 HTML，浏览器可能禁止读取同目录文件；可点击“选择 Excel”载入 fiscal data.xlsx。`,true);
  }finally{hideLoading();}
}

document.getElementById('refreshBtn').addEventListener('click',loadExcel);
document.getElementById('fileInput').addEventListener('change',async e=>{
  const file=e.target.files?.[0]; if(!file) return; document.getElementById('loadingMask').classList.remove('hidden');
  try{const buf=await file.arrayBuffer();const workbook=XLSX.read(buf,{type:'array',cellDates:false});await useWorkbook(workbook,`当前预览文件：${file.name}`);}catch(err){console.error(err);showStatus('error','Excel 读取失败',err.message,true);}finally{hideLoading();}
});
let resizeTimer=null; addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(applyResponsiveCharts,80);});
loadExcel();
