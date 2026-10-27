/* Browser-only, limited XLSX preview for this demo. No network requests. */
"use strict";

function inspectZipBounds(buffer){
  const view=new DataView(buffer), length=buffer.byteLength, min=Math.max(0,length-65557);
  let eocd=-1;
  for(let pos=length-22;pos>=min;pos--)if(view.getUint32(pos,true)===0x06054b50){eocd=pos;break;}
  if(eocd<0)throw new Error("不是可辨識的 .xlsx 檔案。");
  const count=view.getUint16(eocd+10,true), size=view.getUint32(eocd+12,true), start=view.getUint32(eocd+16,true);
  if(count===65535||count>2000||size===0xffffffff||start===0xffffffff||start+size>length)throw new Error("檔案結構超出示範站可處理範圍。");
  let p=start,total=0;
  for(let i=0;i<count;i++){
    if(p+46>length||view.getUint32(p,true)!==0x02014b50)throw new Error("檔案目錄不完整。");
    const flags=view.getUint16(p+8,true), expanded=view.getUint32(p+24,true);
    const nameLen=view.getUint16(p+28,true), extraLen=view.getUint16(p+30,true), commentLen=view.getUint16(p+32,true);
    if(flags&1)throw new Error("示範站不支援加密的 Excel 檔案。");
    total+=expanded;
    if(expanded===0xffffffff||expanded>25*1024*1024||total>70*1024*1024)throw new Error("解壓後資料過大，請使用較小的測試檔。");
    p+=46+nameLen+extraLen+commentLen;
    if(p>start+size)throw new Error("檔案目錄長度不正確。");
  }
}

function xml(text){
  const doc=new DOMParser().parseFromString(text,"application/xml");
  if(doc.getElementsByTagName("parsererror").length)throw new Error("Excel 內容解析失敗。");
  return doc;
}
function all(node,name){return Array.from(node.getElementsByTagNameNS("*",name));}
function children(node,name){return Array.from(node.children||[]).filter(x=>x.localName===name);}
function colNum(ref){const m=/^([A-Z]+)/i.exec(ref||"");if(!m)return 0;return [...m[1].toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);}
function normalize(v){return String(v??"").replace(/[\s\u3000()（）_\-\/]/g,"").toLowerCase();}
function sheetTarget(target){
  const parts=(target.startsWith("/")?target.slice(1):"xl/"+target).split("/"),out=[];
  for(const part of parts){if(part==="..")out.pop();else if(part&&part!==".")out.push(part);}
  return out.join("/");
}

async function readExcelOnDevice(file){
  if(!file.name.toLowerCase().endsWith(".xlsx"))throw new Error("請選擇 .xlsx 檔案。");
  if(file.size>10*1024*1024)throw new Error("檔案超過 10 MB。");
  const bytes=await file.arrayBuffer();inspectZipBounds(bytes);
  const zip=await JSZip.loadAsync(bytes,{checkCRC32:false});
  async function read(path,optional=false){const f=zip.file(path);if(!f){if(optional)return "";throw new Error("Excel 工作表內容不完整。");}return await f.async("string");}
  const workbook=xml(await read("xl/workbook.xml"));
  const relDoc=xml(await read("xl/_rels/workbook.xml.rels"));
  const relationships=new Map(all(relDoc,"Relationship").map(x=>[x.getAttribute("Id"),x.getAttribute("Target")]));
  const sheets=all(workbook,"sheet").map(x=>({name:x.getAttribute("name"),id:x.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id")||x.getAttribute("r:id")}));
  const detail=sheets.find(s=>s.name==="盤點明細")||sheets[0];
  if(!detail||!relationships.has(detail.id))throw new Error("找不到盤點明細工作表。");
  const sharedText=await read("xl/sharedStrings.xml",true);
  const shared=sharedText?all(xml(sharedText),"si").map(si=>all(si,"t").map(t=>t.textContent||"").join("")):[];
  const stylesText=await read("xl/styles.xml",true);let styleFormats=[];
  if(stylesText){
    const styles=xml(stylesText), custom=new Map(all(styles,"numFmt").map(x=>[x.getAttribute("numFmtId"),x.getAttribute("formatCode")]));
    const block=all(styles,"cellXfs")[0];
    styleFormats=block?children(block,"xf").map(x=>custom.get(x.getAttribute("numFmtId"))||""):[];
  }
  function cellValue(cell){
    const t=cell.getAttribute("t")||"n", value=children(cell,"v")[0]?.textContent??"", formula=children(cell,"f").length>0;
    if(t==="e")return {text:"",error:true,formula};
    if(t==="inlineStr")return {text:all(cell,"t").map(x=>x.textContent||"").join(""),formula};
    if(t==="s")return {text:shared[Number(value)]??"",formula};
    let text=value;
    if(t==="b")text=value==="1"?"TRUE":"FALSE";
    const fmt=styleFormats[Number(cell.getAttribute("s")||0)]||"";
    if(t==="n"&&/^\d+$/.test(text)&&/^0+$/.test(fmt))text=text.padStart(fmt.length,"0");
    return {text:String(text).trim(),formula,missingCache:formula&&value===""};
  }
  function parseRows(doc){
    let rows=all(doc,"row");if(rows.length>10002)throw new Error("示範站最多處理 10,000 列資料。");
    return rows.map(r=>{
      const values={};for(const c of children(r,"c")){const col=colNum(c.getAttribute("r"));if(col&&col<=100)values[col]=cellValue(c);}
      return {row:Number(r.getAttribute("r")||0),values};
    });
  }
  const source=xml(await read(sheetTarget(relationships.get(detail.id))));
  const rows=parseRows(source),header=rows.find(r=>r.row===1)||rows[0];
  if(!header)throw new Error("工作表沒有標題。");
  const aliases={sku:["料號","物料編號","物料號碼","sku"],bin:["儲位編碼bin","儲位編碼","儲位","位置","bin"],book:["帳面數量qty","帳面數量","庫存數量","qty"],reference:["文中數量","參考數量"],name:["品名","物料名稱","名稱"]};
  const columns={};
  for(const [key,names] of Object.entries(aliases))for(const [n,v] of Object.entries(header.values))if(names.includes(normalize(v.text))&&!columns[key])columns[key]=Number(n);
  if(!columns.sku||!columns.bin||!columns.book)throw new Error("示範站需要首列含料號、儲位與帳面數量；其他格式請使用正式系統的欄位對應功能。");
  const entries=[],errors=[],warnings=[],seen=new Set(),skuBins=new Map();let requiredFormula=0;
  for(const row of rows){
    if(row.row<=header.row)continue;
    const get=k=>row.values[columns[k]]||{text:""}, sku=get("sku").text.trim(),bin=get("bin").text.trim(),raw=get("book").text.trim();
    if(!sku&&!bin&&!raw)continue;
    if(!sku||!bin||!raw){errors.push(`第 ${row.row} 列：料號、儲位或帳面數量空白。`);continue;}
    if([get("sku"),get("bin"),get("book")].some(c=>c.error||c.missingCache)){errors.push(`第 ${row.row} 列：必要欄位為 Excel 錯誤或公式沒有快取值。`);continue;}
    if(!/^\d+(\.\d{1,3})?$/.test(raw)||!Number.isFinite(Number(raw))||Number(raw)>1000000000){errors.push(`第 ${row.row} 列：帳面數量無法辨識。`);continue;}
    const key=sku+"\u0000"+bin;if(seen.has(key)){errors.push(`第 ${row.row} 列：料號與儲位重複。`);continue;}seen.add(key);
    if([get("sku"),get("bin"),get("book")].some(c=>c.formula))requiredFormula++;
    entries.push({sku,bin,name:get("name").text||"未提供品名",qty:Number(raw),reference:get("reference").text,sourceRow:row.row});
    if(!skuBins.has(sku))skuBins.set(sku,new Set());skuBins.get(sku).add(bin);
  }
  if(requiredFormula)warnings.push(`${requiredFormula} 筆必要欄位含公式，只讀取儲存的快取值。`);
  const master=sheets.find(s=>s.name==="料架位置主檔");
  if(master&&relationships.has(master.id)){
    const masterRows=parseRows(xml(await read(sheetTarget(relationships.get(master.id))))),bins=new Set();let formulas=0;
    for(const r of masterRows)if(r.row>=6&&r.values[7]){let c=r.values[7];if(c.formula)formulas++;if(c.text&&!c.error)bins.add(c.text);}
    if(formulas)warnings.push("位置主檔含公式；本機只比對 Excel 已儲存的快取，請確認是否最新。");
    const missing=entries.filter(e=>!bins.has(e.bin)).length;
    if(missing)warnings.push(`${missing} 筆儲位不在位置主檔，需由庫房確認。`);
  }else warnings.push("沒有找到位置主檔，無法檢查儲位是否存在。");
  const multi=[...skuBins.values()].filter(b=>b.size>1).length;
  if(multi)warnings.push(`${multi} 個料號分布於多個儲位，請確認每個儲位的來源數量。`);
  if(!entries.length&&!errors.length)errors.push("沒有可匯入的明細。");
  return {entries,errors,warnings,sheet:detail.name,source:"您選取的 Excel（僅在本機解析）"};
}
