"use strict";
const KEY="warehouse-mobile-lab-v1", app=document.getElementById("app"), toastBox=document.getElementById("toast");
const icons={home:'<path d="M3 10 12 3l9 7v10a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',list:'<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2M8 10h8M8 14h8M8 18h5"/>',manage:'<rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/>',box:'<path d="m12 2 9 5v10l-9 5-9-5V7zM3 7l9 5 9-5M12 12v10"/>',arrow:'<path d="M4 12h16m-6-6 6 6-6 6"/>',check:'<path d="m4 12 5 5L20 6"/>'};
const icon=n=>`<svg viewBox="0 0 24 24" aria-hidden="true">${icons[n]||icons.box}</svg>`;
const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const defaults=()=>({stock:[],task:null,imports:[]});
function load(){try{const parsed=JSON.parse(localStorage.getItem(KEY)||"null");return parsed&&Array.isArray(parsed.stock)&&Array.isArray(parsed.imports)?parsed:defaults();}catch{return defaults();}}
let state=load();
let ui={view:"home",adminTab:"import",preview:null,itemId:null,filter:"",loading:false};
let toastTimer;
function toast(message,error=false){toastBox.textContent=message;toastBox.className="show"+(error?" error":"");clearTimeout(toastTimer);toastTimer=setTimeout(()=>toastBox.className="",3600);}
function update(fn){
  const next=JSON.parse(JSON.stringify(state));fn(next);
  try{localStorage.setItem(KEY,JSON.stringify(next));state=next;return true;}
  catch{toast("裝置儲存空間不足，這次變更沒有保存。",true);return false;}
}
const taskNames={active:"進行中",review:"待覆核",accounting:"待帳務",closed:"已結案"};
const itemNames={uncounted:"未盤點",counted:"已盤點",exception:"待確認",recount:"需覆盤",approved:"已核准"};
const chip=(text,color="")=>`<span class="chip ${color}">${esc(text)}</span>`;
const taskColor=s=>({active:"green",review:"orange",accounting:"blue",closed:"green"}[s]||"");
const itemColor=s=>({counted:"green",approved:"blue",exception:"red",recount:"orange"}[s]||"");
const heading=(k,title,sub)=>`<div class="heading"><h1>${esc(title)}</h1>${sub?`<p>${esc(sub)}</p>`:""}</div>`;
const stat=(label,value,foot,warn=false)=>`<div class="card stat ${warn?"warn":""}"><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`;
function notice(text,type=""){return `<div class="notice ${type}"><div>${text}</div></div>`;}
function button(label,action,cls="secondary",id=""){return `<button type="button" class="btn ${cls}" data-action="${action}" ${id?`data-id="${esc(id)}"`:""}>${esc(label)}</button>`;}
function formatQty(v){if(v===null||v===undefined||v==="")return "—";return String(Math.round((Number(v)+Number.EPSILON)*1000)/1000);}
function progress(){let t=state.task;if(!t)return {total:0,done:0,approved:0,missing:0};return {total:t.items.length,done:t.items.filter(i=>["counted","approved"].includes(i.state)).length,approved:t.items.filter(i=>i.state==="approved").length,missing:t.items.filter(i=>["uncounted","recount","exception"].includes(i.state)).length};}
function shell(inner){
  let nav=(name,title,ico)=>`<button type="button" data-action="navigate" data-view="${name}" class="${ui.view===name?"active":""}">${icon(ico)}<span>${title}</span></button>`;
  return `<div class="shell"><header class="top"><div class="brand"><strong class="brand-name">SUSS</strong><span class="brand-divider"></span><span class="brand-title">庫房盤點系統</span></div><span class="test-badge">測試版</span></header><div class="banner">僅供範例測試 · 資料保存在本機</div><main class="content">${inner}</main><nav class="bottom" aria-label="主要導覽">${nav("home","總覽","home")}${nav("tasks","盤點","list")}${nav("admin","管理","manage")}</nav></div>`;
}
function home(){
  const p=progress(),missing=state.stock.filter(s=>!s.unit).length;
  const task=state.task;
  return heading("OVERVIEW","工作總覽","")+
   `<div class="hero card"><div class="hero-label">盤點作業</div><h2>${task?esc(task.name):"尚無盤點資料"}</h2><p>${task?esc(taskNames[task.status]):"匯入資料後開始盤點"}</p>${task?`<button type="button" class="btn ghost" data-action="navigate" data-view="tasks">查看盤點</button>`:button("前往管理","go-import","ghost")}</div>
    <div class="section"><h2>作業概況</h2></div><div class="stats">${stat("物料數",state.stock.length,"")}${stat("單位待確認",missing,"",missing>0)}${stat("盤點進度",task?`${p.done} / ${p.total}`:"—","")}${stat("任務狀態",task?taskNames[task.status]:"未建立","")}</div>`;
}
function previewCounts(p){let keys=new Set(state.stock.map(s=>s.sku+"\u0000"+s.bin));return {add:p.entries.filter(e=>!keys.has(e.sku+"\u0000"+e.bin)).length,old:p.entries.filter(e=>keys.has(e.sku+"\u0000"+e.bin)).length};}
function importView(){
  const p=ui.preview;
  let html=`<div class="card panel"><h2>資料匯入</h2><div class="field"><label for="excel-file">選擇 Excel（.xlsx）</label><input id="excel-file" class="file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></div><p class="tiny muted">限範例資料；檔案僅在本機讀取。</p><div class="actions"><a class="btn secondary" href="./warehouse-demo-sample.xlsx?v=20261028-4" download>下載範例</a>${button("載入範例","sample","secondary")}</div></div>`;
  if(ui.loading)html+=notice("正在這台裝置讀取 Excel，請稍候。","");
  if(!p)return html;
  const c=previewCounts(p);
  const stockByKey=new Map(state.stock.map(s=>[s.sku+"\u0000"+s.bin,s]));
  const conflicts=p.entries.filter(e=>{const old=stockByKey.get(e.sku+"\u0000"+e.bin);return e.unit&&old?.unit&&e.unit!==old.unit;});
  const noUnit=p.entries.filter(e=>!e.unit&&!stockByKey.get(e.sku+"\u0000"+e.bin)?.unit).length;
  html+=`<div class="section"><h2>匯入預覽</h2><small>${esc(p.sheet)}</small></div><div class="summary"><div><small>有效列</small><strong>${p.entries.length}</strong></div><div><small>新增</small><strong>${c.add}</strong></div><div><small>既有資料</small><strong>${c.old}</strong></div><div><small>錯誤</small><strong>${p.errors.length+conflicts.length}</strong></div></div>
   ${notice(`來源：${esc(p.source)} · 採用帳面數量`)}
   ${p.errors.length?notice(`<strong>需要修正</strong><br>${p.errors.slice(0,15).map(esc).join("<br>")}${p.errors.length>15?`<br>還有 ${p.errors.length-15} 項錯誤`:""}`,"error"):""}
   ${conflicts.length?notice(`${conflicts.length} 筆物料的 Excel 單位與主檔不同，請先核對。`,"error"):""}
   ${noUnit?notice(`${noUnit} 筆單位未提供，盤點時會顯示未提供。`,"warning"):""}
   ${p.warnings.length?notice(`<strong>請留意</strong><br>${p.warnings.map(esc).join("<br>")}`,"warning"):""}
   <div class="card panel"><div class="row"><h2>資料明細</h2><span class="tiny muted">最多顯示前 20 筆</span></div>${p.entries.slice(0,20).map(e=>`<div class="list-note"><div class="row"><strong>${esc(e.sku)}</strong>${chip(stockByKey.has(e.sku+"\u0000"+e.bin)?"既有／待對帳":"新增","green")}</div><div class="tiny muted">${esc(e.bin)} · 帳面 ${esc(formatQty(e.qty))} · 單位 ${esc(e.unit||"未提供")} · 文中 ${esc(e.reference||"—")}</div></div>`).join("")||'<div class="empty">沒有可匯入的資料</div>'}</div>
   <form id="import-apply" class="card panel" style="margin-top:12px"><div class="field"><label for="source-time">帳面來源時間（必填）</label><input id="source-time" class="input" type="datetime-local" name="source_time" required></div><div class="field"><label><input type="checkbox" name="confirmed" required>已確認匯入資料</label></div>${state.task&&state.task.status!=="closed"?notice("請先完成目前盤點，再匯入。","warning"):""}<button class="btn primary full" type="submit" ${p.errors.length||conflicts.length||!p.entries.length||state.task&&state.task.status!=="closed"?"disabled":""}>確認並開始盤點</button><p class="tiny muted">盤點清單採用本次 Excel 帳面數量；既有主檔不會覆蓋。</p></form>`;
  return html;
}
function stockView(){
  let missing=state.stock.filter(s=>!s.unit).length, term=ui.filter.toLowerCase(),rows=state.stock.filter(s=>[s.sku,s.bin,s.name].some(v=>String(v||"").toLowerCase().includes(term)));
  return `<div class="card panel"><div class="row"><h2>物料主檔</h2>${chip(`單位待補 ${missing}`,missing?"orange":"green")}</div>
  <form id="unit-form"><div class="field"><label for="unit-prefix">料號前綴（留白表示全部待補）</label><input class="input" id="unit-prefix" name="prefix" placeholder="例如 DEMO-1"></div><div class="field"><label for="unit-choice">確認後使用的單位</label><select id="unit-choice" name="unit" required><option value="">請選擇</option><option value="件">件</option><option value="個">個</option><option value="箱">箱</option><option value="米">米</option><option value="公斤">公斤</option></select></div><div class="field"><label><input type="checkbox" name="confirmed" required>我已核對所選料號的單位</label></div><button class="btn secondary full" type="submit" ${!missing?"disabled":""}>為待補資料設定單位</button></form></div>
  <form id="search-form" class="field"><label for="stock-search">搜尋物料</label><div class="actions"><input class="input" id="stock-search" name="term" value="${esc(ui.filter)}" placeholder="料號或儲位"><button class="btn secondary" type="submit">搜尋</button></div></form>
  <div class="section"><h2>物料清單</h2><small>${rows.length} 筆，顯示前 30 筆</small></div><div class="stack">${rows.slice(0,30).map(s=>`<div class="card item"><div class="item-main"><strong>${esc(s.bin)} · ${esc(s.sku)}</strong><small>${esc(s.name)} · 帳面 ${esc(formatQty(s.qty))} ${esc(s.unit||"單位待補")}</small></div>${chip(s.unit?"已確認":"單位待補",s.unit?"green":"orange")}</div>`).join("")||'<div class="card empty">目前沒有符合的物料，請先匯入或調整搜尋。</div>'}</div>`;
}
function taskAdmin(){
  const t=state.task,p=progress();
  if(!t)return `<div class="card empty">尚無盤點清單</div>${button("匯入資料","go-import","primary full")}`;
  let list=t.items.map((i,idx)=>{
    const diff=i.actual===null?null:Math.round((i.actual-i.book)*1000)/1000;
    let actions="";
    if(t.status==="review"&&i.state==="counted")actions=`<div class="field"><label for="reason-${idx}">差異／覆盤原因${diff?"（差異必填）":"（選填）"}</label><input class="input" id="reason-${idx}" placeholder="輸入原因"></div><div class="actions">${button("要求覆盤","recount","secondary",String(idx))}${button("核准此筆","approve","primary",String(idx))}</div>`;
    if(t.status==="active"&&i.state==="exception")actions=button("要求覆盤","recount","secondary",String(idx));
    return `<div class="card panel"><div class="row"><strong>${esc(i.bin)} · ${esc(i.sku)}</strong>${chip(itemNames[i.state],itemColor(i.state))}</div><p class="tiny muted">帳面 ${esc(formatQty(i.book))} ${esc(i.unit||"（單位未提供）")} · 實盤 ${esc(formatQty(i.actual))} ${esc(i.unit||"")}${diff===null?"":` · 差異 ${esc(formatQty(diff))}`}</p>${actions}</div>`;
  }).join("");
  return `<div class="card panel"><div class="row"><h2>${esc(t.name)}</h2>${chip(taskNames[t.status],taskColor(t.status))}</div><p class="tiny muted">已盤 ${p.done} / ${p.total} · 已核准 ${p.approved} / ${p.total}</p></div>
  <div class="section"><h2>任務明細與覆核</h2><small>${p.total} 筆</small></div><div class="stack">${list}</div>
  ${t.status==="accounting"?`<form id="fake-accounting" class="card panel" style="margin-top:12px"><h2>模擬帳務確認</h2><p class="tiny muted">不會寫入 ERP。</p><div class="field"><label for="fake-ref">模擬單據編號</label><input id="fake-ref" class="input" name="reference" placeholder="例如 DEMO-001" required></div><button class="btn primary full" type="submit">記錄並結案</button></form>`:""}
  ${t.status==="closed"?notice(`模擬已結案 · ${esc(t.reference||"無單據")}`)+button("匯入下一次盤點","go-import","secondary full"):""}`;
}
function admin(){
  const tabs=[["import","Excel 匯入"],["stock","物料主檔"],["task","任務與覆核"]];
  return heading("MANAGEMENT","管理後台","")+
   `<div class="tabbar">${tabs.map(([id,label])=>`<button type="button" data-action="admin-tab" data-id="${id}" class="${ui.adminTab===id?"active":""}">${label}</button>`).join("")}</div>`+
   (ui.adminTab==="import"?importView():ui.adminTab==="stock"?stockView():taskAdmin());
}
function taskList(){
 const t=state.task,p=progress();if(!t)return heading("MY TASKS","盤點清單","")+`<div class="card empty">尚未匯入盤點資料</div>${button("前往管理後台","go-import","primary full")}`;
 const missing=t.items.filter(i=>!i.unit).length;
 return heading("INVENTORY TASK",t.name,`${p.done} / ${p.total} 項已盤點 · ${taskNames[t.status]}`)+
  `${missing?notice(`${missing} 筆單位未提供。`,"warning"):""}<div class="section"><h2>儲位與物料</h2><small>${p.total} 項</small></div><div class="stack">${t.items.map((i,idx)=>`<div class="card item"><div class="item-main"><strong>${esc(i.bin)}</strong><small>${esc(i.sku)} · 帳面 ${esc(formatQty(i.book))} ${esc(i.unit||"（單位未提供）")}</small>${i.actual!==null?`<small>實盤 ${esc(formatQty(i.actual))} ${esc(i.unit||"")}</small>`:""}</div><div class="item-side">${chip(itemNames[i.state],itemColor(i.state))}${t.status==="active"&&!['approved','exception'].includes(i.state)?`<button type="button" class="btn secondary" data-action="count-item" data-id="${idx}" aria-label="盤點 ${esc(i.sku)}">盤點</button>`:""}</div></div>`).join("")}</div>
  ${t.status==="active"?`<div style="margin-top:14px">${button("完成盤點，送交覆核","submit-task","primary full")}</div>`:""}`;
}
function countView(){
 const i=state.task?.items[ui.itemId];if(!i)return taskList();
 return `<button class="back" type="button" data-action="back-list">← 返回儲位清單</button>`+heading("COUNT ITEM","開始盤點","")+
  `<div class="card panel"><div class="info-line"><div class="info-icon">${icon("box")}</div><div><strong>${esc(i.bin)}</strong><div class="tiny muted">${esc(i.sku)} · ${esc(i.name)}</div></div></div>
  <hr class="divider"><div class="row"><span class="tiny muted">清點時帳面</span><strong>${esc(formatQty(i.book))} ${esc(i.unit||"（單位未提供）")}</strong></div><form id="count-form"><div class="field quantity"><label for="actual">實際盤點數量</label><input id="actual" class="input" name="actual" inputmode="decimal" type="number" min="0" step="0.001" placeholder="請輸入" required><small>${esc(i.unit||"單位未提供")} · 0 表示無庫存</small></div>
  <div class="actions">${button("回報找不到","exception","secondary")}<button class="btn primary" type="submit">儲存並下一筆</button></div></form></div>`;
}
function render(){app.innerHTML=shell(ui.view==="home"?home():ui.view==="admin"?admin():ui.view==="count"?countView():taskList());}
const sampleEntries=[
 {sku:"DEMO-1001",bin:"A-01-01",name:"合成軸承",qty:12,unit:"件",reference:"12",sourceRow:2},
 {sku:"DEMO-1002",bin:"A-01-02",name:"合成接頭",qty:6,unit:"件",reference:"7",sourceRow:3},
 {sku:"DEMO-1003",bin:"A-02-01",name:"合成固定件",qty:0,unit:"件",reference:"0",sourceRow:4},
 {sku:"DEMO-1004",bin:"B-01-01",name:"合成墊片",qty:20,unit:"件",reference:"20",sourceRow:5}
];
function previewSample(){ui.preview={entries:sampleEntries.map(x=>({...x})),errors:[],warnings:[],sheet:"範例資料",source:"範例資料"};ui.adminTab="import";ui.view="admin";render();toast("已載入範例");}
function go(view){ui.view=view;ui.itemId=null;render();window.scrollTo(0,0);}

document.addEventListener("click",event=>{
  const b=event.target.closest("[data-action]");if(!b)return;
  const action=b.dataset.action,id=Number(b.dataset.id);
  if(action==="navigate")return go(b.dataset.view);
  if(action==="go-import"){ui.adminTab="import";return go("admin");}
  if(action==="admin-tab"){ui.adminTab=b.dataset.id;ui.filter="";render();return;}
  if(action==="sample")return previewSample();
  if(action==="archive-task"){
    if(state.task?.status!=="closed")return;
    if(update(s=>{s.history=s.history||[];s.history.push(s.task);s.task=null;})){render();toast("已保留上一個模擬任務，可建立新任務");}return;
  }
  if(action==="back-list")return go("tasks");
  if(action==="count-item"){ui.itemId=id;ui.view="count";render();window.scrollTo(0,0);return;}
  if(action==="exception"){
    if(update(s=>{let i=s.task.items[ui.itemId];i.state="exception";i.actual=null;})){go("tasks");toast("已回報找不到，等待主管確認");}return;
  }
  if(action==="submit-task"){
    if(!state.task||state.task.items.some(i=>!["counted","approved"].includes(i.state))){toast("仍有未盤或待確認項目，請先處理。",true);return;}
    if(update(s=>{s.task.status="review";})){render();toast("已送交主管覆核");}return;
  }
  if(action==="recount"){
    const item=state.task.items[id],reason=document.getElementById(`reason-${id}`)?.value.trim()||"";
    if(item.state==="counted"&&!reason){toast("請先填寫覆盤原因。",true);return;}
    if(update(s=>{s.task.items[id].state="recount";s.task.items[id].recount_reason=reason||"確認異常後覆盤";s.task.status="active";})){render();toast("已要求重新盤點");}return;
  }
  if(action==="approve"){
    let item=state.task.items[id],reason=document.getElementById(`reason-${id}`)?.value.trim()||"";
    if(item.actual!==item.book&&!reason){toast("有差異時請填寫核准原因。",true);return;}
    if(update(s=>{let x=s.task.items[id],stock=s.stock.find(v=>v.sku===x.sku&&v.bin===x.bin);x.state="approved";x.reason=reason;if(stock)stock.qty=x.actual;if(s.task.items.every(v=>v.state==="approved"))s.task.status="accounting";})){render();toast("已核准模擬盤點結果");}return;
  }
});

document.addEventListener("change",async event=>{
  if(event.target.id!=="excel-file"||!event.target.files[0])return;
  const file=event.target.files[0];
  ui.preview=null;ui.loading=true;render();
  try{ui.preview=await readExcelOnDevice(file);toast("已完成預覽");}
  catch(error){toast(error.message||"讀取 Excel 失敗",true);ui.preview={entries:[],errors:[error.message||"讀取失敗"],warnings:[],sheet:"讀取失敗",source:"選取的 Excel"};}
  finally{ui.loading=false;render();}
});

document.addEventListener("submit",event=>{
  event.preventDefault();const form=event.target,values=new FormData(form);
  if(form.id==="import-apply"){
    const p=ui.preview;if(!p||p.errors.length||!values.get("confirmed")||!values.get("source_time")){toast("請確認資料、來源時間及提示。",true);return;}
    if(state.task&&state.task.status!=="closed"){toast("請先完成目前任務，再套用新匯入。",true);return;}
    const stockByKey=new Map(state.stock.map(s=>[s.sku+"\u0000"+s.bin,s]));
    const conflicts=p.entries.some(e=>{const old=stockByKey.get(e.sku+"\u0000"+e.bin);return e.unit&&old?.unit&&e.unit!==old.unit;});
    if(!p.entries.length||conflicts){toast("匯入資料無效或單位與主檔不同，請核對。",true);return;}
    if(update(s=>{
      if(s.task?.status==="closed"){s.history=s.history||[];s.history.push(s.task);}
      const byKey=new Map(s.stock.map(x=>[x.sku+"\u0000"+x.bin,x]));
      const items=p.entries.map(e=>{
        const key=e.sku+"\u0000"+e.bin;let old=byKey.get(key);
        if(!old){old={...e,unit:e.unit||""};s.stock.push(old);byKey.set(key,old);}
        else if(!old.unit&&e.unit)old.unit=e.unit;
        return {sku:e.sku,bin:e.bin,name:e.name,unit:e.unit||old.unit||"",book:e.qty,actual:null,state:"uncounted",sourceRow:e.sourceRow};
      });
      s.imports.push({source_time:String(values.get("source_time")),total:p.entries.length,source:p.source});
      s.task={name:"本次盤點",prefix:"",status:"active",source_time:String(values.get("source_time")),items};
    })){ui.preview=null;go("tasks");toast("已匯入，請開始盤點");}return;
  }
  if(form.id==="unit-form"){
    const unit=String(values.get("unit")||""),prefix=String(values.get("prefix")||"").trim();
    if(!unit||!values.get("confirmed")){toast("請選擇並確認單位。",true);return;}
    const n=state.stock.filter(s=>!s.unit&&s.sku.startsWith(prefix)).length;if(!n){toast("沒有符合範圍的待補資料。",true);return;}
    if(update(s=>{s.stock.forEach(x=>{if(!x.unit&&x.sku.startsWith(prefix))x.unit=unit;});s.task?.items.forEach(i=>{if(!i.unit&&["uncounted","recount"].includes(i.state)&&i.sku.startsWith(prefix))i.unit=unit;});})){render();toast(`已為 ${n} 筆測試資料設定單位`);}return;
  }
  if(form.id==="search-form"){ui.filter=String(values.get("term")||"").trim();render();return;}
  if(form.id==="task-create"){
    const name=String(values.get("name")||"").trim(),prefix=String(values.get("prefix")||"").trim(),selected=state.stock.filter(s=>s.bin.startsWith(prefix));
    if(!name||!selected.length){toast("請填任務名稱，並確認範圍內有物料。",true);return;}
    if(selected.some(s=>!s.unit)){toast("範圍內有物料尚未確認單位，請先補齊。",true);return;}
    if(update(s=>{s.task={name,prefix,status:"active",items:s.stock.filter(x=>x.bin.startsWith(prefix)).map(x=>({sku:x.sku,bin:x.bin,name:x.name,unit:x.unit,book:x.qty,actual:null,state:"uncounted"}))};})){render();toast("任務已建立，請至盤點頁操作");}return;
  }
  if(form.id==="count-form"){
    const raw=String(values.get("actual")||"").trim(),n=Number(raw);
    if(!/^\d+(\.\d{1,3})?$/.test(raw)||!Number.isFinite(n)||n>1000000000){toast("請輸入有效的非負數量，最多三位小數。",true);return;}
    if(update(s=>{let i=s.task.items[ui.itemId];i.actual=n;i.state="counted";})){
      const next=state.task.items.findIndex(i=>["uncounted","recount"].includes(i.state));
      if(next>=0){ui.itemId=next;render();window.scrollTo(0,0);}else go("tasks");toast("已儲存於這台裝置");
    }return;
  }
  if(form.id==="fake-accounting"){
    const ref=String(values.get("reference")||"").trim();if(!ref){toast("請填寫模擬單據編號。",true);return;}
    if(update(s=>{s.task.status="closed";s.task.reference=ref;})){render();toast("模擬任務已結案");}return;
  }
});
window.addEventListener("storage",event=>{if(event.key===KEY){state=load();render();toast("這個瀏覽器的另一個分頁已更新。");}});
render();
