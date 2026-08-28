const EXCEL_URL='data/budget-data.json';
const ACTUALS_KEY='household-budget-actuals-v1';
let excel=null;
let data=null;
const yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(Number(n)||0);
const actuals=JSON.parse(localStorage.getItem(ACTUALS_KEY)||'{}');
function saveActuals(){localStorage.setItem(ACTUALS_KEY,JSON.stringify(actuals));}
function key(p){return `${p.name}|${p.type}`}
function buildData(){
  const items=(excel.paymentItems||[]).map((p,i)=>({id:i+1,name:p.name,type:p.type,month:p.latestMonth,monthly:Number(p.monthly)||0,baseRemaining:p.remainingAmount==null?null:Number(p.remainingAmount),baseCount:p.remainingCount==null?null:Number(p.remainingCount),paid:Number(actuals[key(p)]||0)}));
  const months=(excel.months||[]).map(m=>({month:m.month,income:Number(m.income)||0,balance:m.balance==null?null:Number(m.balance),fixed:(m.payments||[]).filter(p=>p.category==='固定費').reduce((s,p)=>s+Number(p.amount||0),0),utilities:(m.payments||[]).filter(p=>['電気','ガス','水道'].includes(p.category)).reduce((s,p)=>s+Number(p.amount||0),0),installments:(m.payments||[]).filter(p=>['リボ','分割'].includes(p.category)).reduce((s,p)=>s+Number(p.amount||0),0),other:(m.payments||[]).filter(p=>p.category==='その他').reduce((s,p)=>s+Number(p.amount||0),0),total:Number(m.paymentTotal)||0}));
  return {payments:items,months};
}
function remaining(p){return p.baseRemaining==null?null:Math.max(0,p.baseRemaining-p.paid)}
function countRemaining(p){const r=remaining(p);return r==null?null:Math.max(0,Math.ceil(r/Math.max(1,p.monthly)))}
function paymentAt(p,monthsFromLatest){if(p.type==='分割'&&p.baseCount!=null)return monthsFromLatest<p.baseCount?p.monthly:0;return p.monthly}
function monthLabel(m){const [y,mo]=m.split('-');return `${y}年${Number(mo)}月`}
function addMonths(base,n){const d=new Date(`${base}-01`);d.setMonth(d.getMonth()+n);return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`}
function renderExcel(){const status=document.querySelector('#dataStatus'),box=document.querySelector('#excelSummary');if(!excel){status.textContent='Excelデータを読み込めませんでした。';box.innerHTML='<p>data/budget-data.json を確認してください。</p>';return}status.textContent=`Excelデータ読み込み済み：${excel.source}（${excel.sheetNames.length}シート）`;const rules=excel.rules||{};const chips=Object.entries(excel.categories||{}).map(([k,v])=>`<div class="card"><div class="label">${k}</div><div class="value">${v.length}シート</div><small>${v.length?v.join(' / '):'検出なし'}</small></div>`).join('');box.innerHTML=`<p><strong>生成日時：</strong>${excel.generatedAt}</p><p>収入：A・B列 ／ 残高：${rules.balanceCell||'B7'} ／ 支払額：${rules.paymentAmountColumn||'G'}列 ／ 内訳：${rules.paymentDetailColumn||'E'}列 ／ 分割判定：H・I列</p><div class="cards">${chips}</div><details><summary>読み込んだシート一覧</summary><ul>${excel.sheetNames.map(s=>`<li>${s}</li>`).join('')}</ul></details>`}
function render(){
  if(!data)return;
  const active=data.payments.filter(p=>remaining(p)===null||remaining(p)>0);
  const knownRemaining=active.reduce((s,p)=>s+(remaining(p)||0),0);
  const latest=data.months[data.months.length-1];
  const forecast=[];
  for(let i=0;i<12;i++){const m=addMonths(latest?.month||new Date().toISOString().slice(0,7),i);const income=latest?.income||0;const payments=data.payments.reduce((s,p)=>s+paymentAt(p,i),0);const fixed=latest?.fixed||0,utilities=latest?.utilities||0,other=latest?.other||0;forecast.push({m,income,fixed,utilities,installments:payments,other,total:fixed+utilities+payments+other,balance:income-(fixed+utilities+payments+other)})}
  const allRows=[...(data.months||[]).map(m=>({...m,total:m.total||0,balance:m.balance})),...forecast.slice(1)];
  const max=allRows.reduce((a,b)=>Math.max(a,b.total||0),0);
  document.querySelector('#summary').innerHTML=[['今後の支払残高（分割で算出）',yen(knownRemaining)],['直近月の支払額',yen(latest?.total)],['今後12か月の月間最高必要額',yen(max)],['支払い項目数',data.payments.length+'件']].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="value">${x[1]}</div></div>`).join('');
  document.querySelector('#payments').innerHTML=data.payments.map(p=>{const rem=remaining(p),count=countRemaining(p),done=rem!==null&&rem<=0;return `<tr><td><strong>${p.name}</strong></td><td>${p.type}</td><td>${p.month}</td><td>${yen(p.monthly)}</td><td>${rem==null?'未入力':yen(rem)}</td><td>${count==null?'—':count+'回'}</td><td><input class="input paid-input" data-id="${p.id}" type="number" min="0" step="100" value="${p.paid}"></td><td><span class="status ${done?'paid':'active'}">${done?'完済':(p.type==='分割'?'支払予定':'継続')}</span></td><td></td></tr>`}).join('');
  document.querySelector('#months').innerHTML=allRows.map((m,idx)=>{const actual=idx<data.months.length;const remain=m.balance==null?m.income-m.total:m.balance;return `<tr><td><strong>${monthLabel(m.month||m.m)}</strong>${actual?'':'<small>（予測）</small>'}</td><td>${yen(m.income)}</td><td>${yen(m.fixed)}</td><td>${yen(m.utilities)}</td><td>${yen(m.installments)}</td><td>${yen(m.other)}</td><td><strong>${yen(m.total)}</strong></td><td class="${remain>=0?'positive':'negative'}">${yen(remain)}</td></tr>`}).join('');
  document.querySelector('#milestones').innerHTML=data.payments.map(p=>{const c=countRemaining(p),r=remaining(p);if(r!==null&&r<=0)return `<div class="milestone"><strong>${p.name}</strong>：完済済み</div>`;if(c==null)return `<div class="milestone"><strong>${p.name}</strong>：毎月${yen(p.monthly)}（残高はExcelで別管理）</div>`;const end=new Date(`${p.month}-01`);end.setMonth(end.getMonth()+c-1);return `<div class="milestone"><strong>${p.name}</strong>：あと${c}回・${yen(r)} → ${end.getFullYear()}年${end.getMonth()+1}月ごろ完済</div>`}).join('');
  document.querySelectorAll('.paid-input').forEach(el=>el.onchange=()=>{const p=data.payments.find(x=>x.id==el.dataset.id);p.paid=Math.max(0,Number(el.value)||0);actuals[key(p)]=p.paid;saveActuals();render()});
  renderExcel();
}
async function loadExcel(){try{const r=await fetch(EXCEL_URL,{cache:'no-store'});if(!r.ok)throw new Error(`HTTP ${r.status}`);excel=await r.json();data=buildData();render()}catch(e){console.error(e);document.querySelector('#dataStatus').textContent='Excelデータの読み込みに失敗しました。'}}
loadExcel();
