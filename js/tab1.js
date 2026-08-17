const EXCEL_FILE = window.APP_CONFIG?.ICHI_EXCEL_URL || "ichi data.xlsx";
const EXCEL_LABEL = "ichi data.xlsx";
const COLORS = {
  paper:'#E9ECEF', panel:'#FFFFFF', ink:'#14181C', ink2:'#626C76', ink3:'#8D959D', rule:'#D3D9DF',
  zhu:'#A32B2B', zhuL:'#C98B84', indigo:'#1F4E79', indigoL:'#7FA0BE', brass:'#8A6E28'
};
const MONO = 'SF Mono,JetBrains Mono,Roboto Mono,Consolas,monospace';
const SANS = '-apple-system,BlinkMacSystemFont,PingFang SC,Microsoft YaHei,Hiragino Sans GB,sans-serif';
const SERIF = 'Songti SC,STSong,Source Han Serif SC,Noto Serif SC,SimSun,serif';

function cacheBust(url){
  const sep=url.includes('?')?'&':'?';
  return encodeURI(url)+sep+'v='+Date.now();
}

let DATA = [];
let rangeWeeks = 52;
let charts = {};

function num(v){
  if(v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function fmt(v,d=2){return v==null?'—':Number(v).toFixed(d)}
function signed(v,d=2){
  if(v==null) return '—';
  const x = Number(v);
  if(Math.abs(x) < Math.pow(10,-d)/2) return Number(0).toFixed(d);
  return (x>0?'+':'') + x.toFixed(d);
}
function direction(v,threshold=0.005){
  if(v==null || Math.abs(v)<threshold) return {arrow:'→', cls:'delta-flat', word:'持平'};
  return v>0 ? {arrow:'↑',cls:'delta-up',word:'上升'} : {arrow:'↓',cls:'delta-down',word:'下降'};
}
function ma(arr,window){
  return arr.map((_,i)=>{
    if(i < window-1) return null;
    const s = arr.slice(i-window+1,i+1);
    if(s.some(v=>v==null)) return null;
    return s.reduce((a,b)=>a+b,0)/window;
  });
}
function latest(){return DATA[DATA.length-1] || null}
function previous(){return DATA[DATA.length-2] || null}
function getChart(id){
  if(!charts[id]) charts[id] = echarts.init(document.getElementById(id), null, {renderer:'canvas'});
  return charts[id];
}
function baseOption(){
  return {
    animationDuration:500,
    textStyle:{fontFamily:SANS,color:COLORS.ink2},
    tooltip:{
      trigger:'axis',
      backgroundColor:'rgba(20,24,28,.94)',
      borderWidth:0,
      padding:[10,12],
      textStyle:{color:'#fff',fontSize:11,fontFamily:MONO},
      extraCssText:'box-shadow:0 10px 28px rgba(20,24,28,.18);'
    },
    grid:{left:48,right:20,top:28,bottom:35,containLabel:false},
    xAxis:{
      type:'category',boundaryGap:false,
      axisLine:{lineStyle:{color:COLORS.rule}},
      axisTick:{show:false},
      axisLabel:{color:COLORS.ink3,fontSize:10,fontFamily:MONO,hideOverlap:true},
      splitLine:{show:false}
    },
    yAxis:{
      type:'value',scale:true,
      axisLine:{show:false},axisTick:{show:false},
      axisLabel:{color:COLORS.ink3,fontSize:10,fontFamily:MONO,formatter:v=>Number(v).toFixed(1)},
      splitLine:{lineStyle:{color:'#E7EBEE',type:'solid'}}
    }
  };
}
function mark100(){
  return {silent:true,symbol:'none',label:{show:true,formatter:'100',position:'insideEndTop',color:COLORS.ink3,fontFamily:MONO,fontSize:10},lineStyle:{color:COLORS.ink3,type:'dashed',width:1},data:[{yAxis:100}]};
}
function chartWindow(){
  if(rangeWeeks === 'all') return DATA;
  return DATA.slice(-Math.min(Number(rangeWeeks),DATA.length));
}
function renderMainChart(){
  const W = chartWindow();
  const labels = W.map(d=>d.weeknum);
  const values = W.map(d=>d.composite);
  const allMA = ma(DATA.map(d=>d.composite),4);
  const offset = DATA.length-W.length;
  const avg = allMA.slice(offset);
  const chart = getChart('mainChart');
  const opt = baseOption();
  opt.grid = {left:52,right:22,top:36,bottom:38};
  opt.xAxis.data = labels;
  opt.xAxis.axisLabel.interval = rangeWeeks === 26 ? 3 : (rangeWeeks === 52 ? 7 : 'auto');
  opt.legend = {top:0,right:0,itemWidth:18,itemHeight:2,textStyle:{fontFamily:MONO,fontSize:10,color:COLORS.ink3},data:['周度指数','4周均值']};
  opt.series = [
    {
      name:'周度指数',type:'line',data:values,symbol:'none',smooth:false,
      lineStyle:{width:1.35,color:COLORS.zhuL},itemStyle:{color:COLORS.zhuL},
      markLine:mark100(),z:3
    },
    {
      name:'4周均值',type:'line',data:avg,symbol:'none',smooth:.16,
      lineStyle:{width:2.6,color:COLORS.zhu},itemStyle:{color:COLORS.zhu},z:5
    }
  ];
  opt.tooltip.formatter = params=>{
    if(!params || !params.length) return '';
    let s = `<b>${params[0].axisValue}</b>`;
    params.forEach(p=>{ if(p.value!=null) s += `<br/>${p.marker}${p.seriesName}　<b>${fmt(p.value,2)}</b>`; });
    return s;
  };
  chart.setOption(opt,true);
}
function renderMini(id,key,color){
  const W = DATA.slice(-13);
  const labels = W.map(d=>d.weeknum);
  const vals = W.map(d=>d[key]);
  const chart = getChart(id);
  const opt = baseOption();
  opt.grid = {left:44,right:14,top:14,bottom:26};
  opt.xAxis.data = labels;
  opt.xAxis.axisLabel.interval = 3;
  opt.xAxis.axisLabel.fontSize = 9;
  opt.yAxis.axisLabel.fontSize = 9;
  opt.series = [{
    type:'line',data:vals,symbol:'none',smooth:.18,
    lineStyle:{width:2,color},itemStyle:{color},
    areaStyle:{color:color+'12'},
    markLine:mark100(),z:4
  }];
  opt.tooltip.formatter = params=>{
    const p=params&&params[0];
    return p?`<b>${p.axisValue}</b><br/>指数　<b>${fmt(p.value,2)}</b>`:'';
  };
  chart.setOption(opt,true);
}
function renderStructure(){
  const L = latest();
  if(!L) return;
  const items = [
    {name:'消费',value:L.consumption-100},
    {name:'投资',value:L.investment-100},
    {name:'出口',value:L.export-100},
    {name:'生产',value:L.production-100}
  ].sort((a,b)=>b.value-a.value);
  const maxAbs = Math.max(.15,...items.map(x=>Math.abs(x.value))) * 1.28;
  const chart = getChart('structureChart');
  chart.setOption({
    animationDuration:500,
    textStyle:{fontFamily:SANS,color:COLORS.ink2},
    grid:{left:54,right:55,top:18,bottom:26},
    tooltip:{trigger:'axis',axisPointer:{type:'shadow'},backgroundColor:'rgba(20,24,28,.94)',borderWidth:0,textStyle:{color:'#fff',fontSize:11,fontFamily:MONO},formatter:p=>{
      const x=p&&p[0]; return x?`<b>${x.name}</b><br/>相对 100　<b>${signed(x.value,2)}</b>`:'';
    }},
    xAxis:{
      type:'value',min:-maxAbs,max:maxAbs,
      axisLine:{show:true,lineStyle:{color:COLORS.ink3}},axisTick:{show:false},
      axisLabel:{color:COLORS.ink3,fontSize:10,fontFamily:MONO,formatter:v=>signed(v,1)},
      splitLine:{lineStyle:{color:'#E7EBEE'}}
    },
    yAxis:{
      type:'category',data:items.map(x=>x.name),inverse:true,
      axisLine:{show:false},axisTick:{show:false},
      axisLabel:{color:COLORS.ink,fontSize:12,fontFamily:SANS,fontWeight:500}
    },
    series:[{
      type:'bar',data:items.map(x=>x.value),barWidth:16,
      itemStyle:{color:p=>p.value>=0?COLORS.zhu:COLORS.indigo},
      label:{show:true,position:'outside',formatter:p=>signed(p.value,2),color:COLORS.ink2,fontFamily:MONO,fontSize:10},
      markLine:{silent:true,symbol:'none',label:{show:false},lineStyle:{color:COLORS.ink,width:1},data:[{xAxis:0}]}
    }]
  },true);
}
function updateCards(){
  const L = latest(), P = previous();
  if(!L) return;
  const avg4 = ma(DATA.map(d=>d.composite),4).at(-1);
  const defs = [
    {k:'综合景气指数',key:'composite',primary:true},
    {k:'消费景气指数',key:'consumption'},
    {k:'投资景气指数',key:'investment'},
    {k:'出口景气指数',key:'export'},
    {k:'生产景气指数',key:'production'}
  ];
  document.getElementById('readout').innerHTML = defs.map(x=>{
    const v=L[x.key], delta=P? v-P[x.key]:null, dir=direction(delta);
    const d = x.primary
      ? `周环比 <span class="${dir.cls}">${dir.arrow} ${signed(delta,2)}</span><br/>4W MA ${fmt(avg4,2)}`
      : `周环比 <span class="${dir.cls}">${dir.arrow} ${signed(delta,2)}</span><br/>距100 ${signed(v-100,2)}`;
    return `<div class="rd ${x.primary?'primary':''}"><div class="k">${x.k}</div><div class="v">${fmt(v,2)}</div><div class="d">${d}</div></div>`;
  }).join('');
}
function updateComponentMeta(valueId,metaId,key){
  const L=latest(),P=previous(); if(!L) return;
  const v=L[key],delta=P?v-P[key]:null,dir=direction(delta);
  document.getElementById(valueId).textContent=fmt(v,2);
  document.getElementById(metaId).innerHTML=`<span class="${dir.cls}">${dir.arrow} ${signed(delta,2)}</span><br/>距100 ${signed(v-100,2)}`;
}
function updateReading(){
  const L=latest(),P=previous(); if(!L) return;
  const delta=P?L.composite-P.composite:null;
  const dir=direction(delta);
  const comps=[
    {name:'消费',v:L.consumption},{name:'投资',v:L.investment},{name:'出口',v:L.export},{name:'生产',v:L.production}
  ];
  const strongest=[...comps].sort((a,b)=>(b.v-100)-(a.v-100))[0];
  const weakest=[...comps].sort((a,b)=>(a.v-100)-(b.v-100))[0];
  const above=comps.filter(x=>x.v>=100).length;
  let sentence=`${L.weeknum} 综合景气指数为 ${fmt(L.composite,2)}，较上周${dir.word}${Math.abs(delta||0).toFixed(2)}点。`;
  sentence+=`${strongest.name}景气当前最强，为 ${fmt(strongest.v,2)}，高于100景气线 ${Math.abs(strongest.v-100).toFixed(2)}点。`;
  if(weakest.v<100){
    sentence+=`${weakest.name}景气相对偏弱，为 ${fmt(weakest.v,2)}，低于景气线 ${Math.abs(weakest.v-100).toFixed(2)}点。`;
  }else{
    sentence+=`四个分项中有 ${above} 项位于100景气线以上，整体结构偏扩张。`;
  }
  document.getElementById('readingText').textContent=sentence;
  const avg4=ma(DATA.map(d=>d.composite),4).at(-1);
  document.getElementById('readingDetails').innerHTML=`<span>4W MA ${fmt(avg4,2)}</span><span>高于100分项 ${above}/4</span><span>最强分项 ${strongest.name}</span>`;
}
function renderAll(){
  if(!DATA.length) return;
  document.getElementById('asof').textContent=latest().weeknum;
  updateCards();
  updateComponentMeta('vConsumption','mConsumption','consumption');
  updateComponentMeta('vInvestment','mInvestment','investment');
  updateComponentMeta('vExport','mExport','export');
  updateComponentMeta('vProduction','mProduction','production');
  renderMainChart();
  renderMini('consumptionChart','consumption',COLORS.ink);
  renderMini('investmentChart','investment',COLORS.brass);
  renderMini('exportChart','export',COLORS.indigoL);
  renderMini('productionChart','production',COLORS.indigo);
  renderStructure();
  updateReading();
}

function normalizeWorkbook(workbook){
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null});
  let h = -1;
  for(let i=0;i<Math.min(rows.length,20);i++){
    const row=(rows[i]||[]).map(x=>String(x??'').trim());
    if(row.includes('weeknum') && row.includes('综合景气指数')){h=i;break;}
  }
  if(h<0) throw new Error('未找到表头：需要包含 weeknum 和 综合景气指数');
  const headers=(rows[h]||[]).map(x=>String(x??'').trim());
  const idx=name=>headers.indexOf(name);
  const required=['weeknum','消费景气指数','投资景气指数','出口景气指数','生产景气指数','综合景气指数'];
  const missing=required.filter(x=>idx(x)<0);
  if(missing.length) throw new Error('Excel 缺少字段：'+missing.join('、'));
  const out=[];
  for(let i=h+1;i<rows.length;i++){
    const r=rows[i]||[];
    const wk=String(r[idx('weeknum')]??'').trim();
    if(!wk) continue;
    const item={
      weeknum:wk,
      consumption:num(r[idx('消费景气指数')]),
      investment:num(r[idx('投资景气指数')]),
      export:num(r[idx('出口景气指数')]),
      production:num(r[idx('生产景气指数')]),
      composite:num(r[idx('综合景气指数')])
    };
    if([item.consumption,item.investment,item.export,item.production,item.composite].some(v=>v==null)) continue;
    out.push(item);
  }
  if(out.length<4) throw new Error('有效周度数据不足 4 行');
  return out;
}
function showStatus(type,title,text,showPicker=false){
  const box=document.getElementById('statusBox');
  box.className=`status-box show ${type||''}`;
  document.getElementById('statusTitle').textContent=title;
  document.getElementById('statusText').textContent=text;
  document.getElementById('fileLabel').style.display=showPicker?'inline-block':'none';
}
function hideStatus(){document.getElementById('statusBox').className='status-box'}
function hideLoading(){document.getElementById('loadingMask').classList.add('hidden')}

async function loadExcel(){
  document.getElementById('loadingMask').classList.remove('hidden');
  hideStatus();
  try{
    if(typeof XLSX==='undefined') throw new Error('Excel 读取组件未加载，请检查网络连接');
    if(typeof echarts==='undefined') throw new Error('图表组件未加载，请检查网络连接');
    const url=cacheBust(EXCEL_FILE);
    const res=await fetch(url,{cache:'no-store'});
    if(!res.ok) throw new Error(`读取 ${EXCEL_LABEL} 失败（HTTP ${res.status}）`);
    const buf=await res.arrayBuffer();
    const workbook=XLSX.read(buf,{type:'array'});
    DATA=normalizeWorkbook(workbook);
    renderAll();
  }catch(err){
    console.error(err);
    showStatus('error','未能自动读取 Excel',`${err.message}。如果你是直接双击打开 HTML，浏览器可能禁止读取同目录文件；部署到网站/本地服务器后会自动读取，也可暂时手动选择 Excel 预览。`,true);
  }finally{
    hideLoading();
  }
}

document.getElementById('refreshBtn').addEventListener('click',loadExcel);
document.getElementById('fileInput').addEventListener('change',async e=>{
  const file=e.target.files&&e.target.files[0]; if(!file) return;
  document.getElementById('loadingMask').classList.remove('hidden');
  try{
    const buf=await file.arrayBuffer();
    const workbook=XLSX.read(buf,{type:'array'});
    DATA=normalizeWorkbook(workbook);
    renderAll();
  }catch(err){
    showStatus('error','Excel 读取失败',err.message,true);
  }finally{hideLoading();}
});
document.querySelectorAll('.range-btn').forEach(btn=>btn.addEventListener('click',()=>{
  document.querySelectorAll('.range-btn').forEach(x=>x.classList.remove('active'));
  btn.classList.add('active');
  rangeWeeks=btn.dataset.weeks==='all'?'all':Number(btn.dataset.weeks);
  renderMainChart();
}));
window.addEventListener('resize',()=>Object.values(charts).forEach(c=>c&&c.resize()));

loadExcel();
