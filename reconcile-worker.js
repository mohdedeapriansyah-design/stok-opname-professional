
// V10.5: trim worksheet ranges to the last meaningful cell.
// This avoids iterating Excel's accidental 1,048,576-row !ref on mobile.
function trimWorksheetRange(ws) {
  if (!ws || !ws["!ref"] || typeof XLSX === "undefined") return ws;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let lastR = range.s.r - 1, lastC = range.s.c - 1;

  // V10.6: inspect only cells that actually exist in the worksheet object.
  // This avoids a million-row x column nested scan on Android.
  for (const key of Object.keys(ws)) {
    if (key[0] === "!") continue;
    const cell = ws[key];
    if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === "") continue;
    const pos = XLSX.utils.decode_cell(key);
    if (pos.r > lastR) lastR = pos.r;
    if (pos.c > lastC) lastC = pos.c;
  }

  if (lastR >= range.s.r && lastC >= range.s.c) {
    ws["!ref"] = XLSX.utils.encode_range({
      s: range.s,
      e: { r: lastR, c: lastC }
    });
  }
  return ws;
}

/* Reconcile Stock Pro V10.4
   Template Komper is READ-ONLY. No cell is written back to it.
   Heavy XLSX work runs in this worker to avoid Android UI freezes. */
importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

const aliases = {
  odoo:["productodooCode","productOdooCode","odooCode","odooProductCode","kodeOdoo","kodeOdooProduk"],
  sap:["productSapCode","productsapcode","sapCode","kodeSAP","kodeSap","materialCode","material"],
  item:["kodeItem","kodeBarang","itemCode","item","sku","productCode","kodeproduk"],
  name:["productName","namaProduk","productNama","description","namaBarang","product"],
  qty:["qtyFix","qtyfix","qty","quantity","qtySO","qtyOpname","stockQty","countQty"],
  target:["target","targetsekarang","targetso","targetpenjualan","targetqty","qtytarget"]
};
const norm=v=>String(v??"").trim().toLowerCase().replace(/[\s_\-./()]+/g,"");
const numericKey=v=>{
  if(v===null||v===undefined||v==="") return "";
  const s=String(v).trim().replace(/,/g,"").replace(/\.0+$/,"").replace(/\s/g,"");
  return /^\d+$/.test(s) ? (s.replace(/^0+/,"")||"0") : s.toLowerCase();
};
const qtyNumber=v=>{
  if(v===null||v===undefined||v==="") return null;
  const n=Number(String(v).replace(/,/g,"").trim());
  return Number.isFinite(n)?n:null;
};
function findHeader(rows,type,preferredRow=null){
  const wanted=new Set(aliases[type].map(norm));
  const max=Math.min(rows.length,40);
  if(preferredRow!==null&&rows[preferredRow]){
    for(let c=0;c<rows[preferredRow].length;c++) if(wanted.has(norm(rows[preferredRow][c]))) return {rowIndex:preferredRow,colIndex:c};
  }
  for(let r=0;r<max;r++) for(let c=0;c<(rows[r]||[]).length;c++) if(wanted.has(norm(rows[r][c]))) return {rowIndex:r,colIndex:c};
  return null;
}
function actualRange(ws){
  const ref=ws["!ref"]||"A1:A1",base=XLSX.utils.decode_range(ref);
  let maxR=base.e.r,maxC=base.e.c;
  for(const key of Object.keys(ws)){
    if(key[0]==="!") continue;
    const m=/^([A-Z]+)(\d+)$/.exec(key); if(!m) continue;
    const r=+m[2]-1,c=XLSX.utils.decode_col(m[1]),cell=ws[key];
    if(cell&&(cell.v!==undefined||cell.f!==undefined)){if(r>maxR)maxR=r;if(c>maxC)maxC=c;}
  }
  return {s:{r:base.s.r,c:base.s.c},e:{r:maxR,c:maxC}};
}
function rowsFromSheet(ws){return XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true,range:actualRange(ws)});}
function indexTemplate(wb){
  const all=[];
  for(const sheetName of wb.SheetNames){
    const ws=wb.Sheets[sheetName],rows=rowsFromSheet(ws);
    if(!rows.length) continue;
    const item=findHeader(rows,"item"),odoo=findHeader(rows,"odoo"),sap=findHeader(rows,"sap"),name=findHeader(rows,"name"),target=findHeader(rows,"target");
    const code=item||odoo||sap; if(!code) continue;
    const headerRow=code.rowIndex;
    const entry={sheetName,rows,headerRow,itemCol:item?.colIndex??-1,odooCol:odoo?.colIndex??-1,sapCol:sap?.colIndex??-1,nameCol:name?.colIndex??-1,targetCol:target?.rowIndex===headerRow?target.colIndex:-1,map:new Map()};
    for(let r=headerRow+1;r<rows.length;r++){
      const row=rows[r]||[];
      if(entry.odooCol>=0){const k=numericKey(row[entry.odooCol]);if(k&&!entry.map.has("odoo:"+k))entry.map.set("odoo:"+k,r);}
      if(entry.sapCol>=0){const k=numericKey(row[entry.sapCol]);if(k&&!entry.map.has("sap:"+k))entry.map.set("sap:"+k,r);}
      if(entry.itemCol>=0){const k=numericKey(row[entry.itemCol]);if(k&&!entry.map.has("item:"+k))entry.map.set("item:"+k,r);}
    }
    all.push(entry);
  }
  return all;
}
function globalMap(indexes){
  const maps={odoo:new Map(),sap:new Map(),item:new Map()};
  for(const idx of indexes) for(const [compound,row] of idx.map){
    const p=compound.indexOf(":"),kind=compound.slice(0,p),key=compound.slice(p+1);
    if(!maps[kind].has(key)) maps[kind].set(key,[]);
    maps[kind].get(key).push({idx,row});
  }
  return maps;
}
function findCandidate(x,maps){
  for(const kind of ["odoo","sap","item"]){
    const k=numericKey(x[kind]);
    if(k&&maps[kind].has(k)) return maps[kind].get(k)[0];
  }
  const sap=numericKey(x.sap);
  if(/^\d+$/.test(sap)&&sap.length>4){
    const suffix=numericKey(sap.slice(-6));
    if(maps.item.has(suffix)) return maps.item.get(suffix)[0];
  }
  return null;
}
function parseSO(wb){
  const out=[];
  for(const sheetName of wb.SheetNames){
    const rows=rowsFromSheet(wb.Sheets[sheetName]); if(!rows.length) continue;
    const odoo=findHeader(rows,"odoo"),sap=findHeader(rows,"sap"),item=findHeader(rows,"item"),name=findHeader(rows,"name"),qty=findHeader(rows,"qty");
    if(!qty) continue;
    const headerRow=Math.max(qty.rowIndex,odoo?.rowIndex??0,sap?.rowIndex??0,item?.rowIndex??0);
    for(let r=headerRow+1;r<rows.length;r++){
      const row=rows[r]||[],q=qtyNumber(row[qty.colIndex]); if(q===null) continue;
      const o=odoo?row[odoo.colIndex]:"",s=sap?row[sap.colIndex]:"",i=item?row[item.colIndex]:"";
      if(!numericKey(o)&&!numericKey(s)&&!numericKey(i)) continue;
      out.push({odoo:o,sap:s,item:i,name:name?row[name.colIndex]:"",qty:q,source:sheetName});
    }
  }
  return out;
}
function makeResultWorkbook(rows){
  const headers=["Sheet","Kode Item","Nama Produk","Odoo","SAP","Qty SO","Status"];
  const data=[headers,...rows.map(x=>[
    x.sheet,x.item,x.name,x.odoo,x.sap,x.qty,x.status
  ])];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws["!cols"]=[{wch:18},{wch:18},{wch:18},{wch:18},{wch:32},{wch:14},{wch:14},{wch:14},{wch:16}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Hasil Reconcile");
  const summary=[["Metric","Value"],["Total baris SO",rows.reduce((a,x)=>a+(x.sourceMatched?1:0),0)],["SKU cocok",rows.length],["Qty SO",rows.reduce((a,x)=>a+(Number(x.qty)||0),0)]];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),"Ringkasan");
  return XLSX.write(wb,{bookType:"xlsx",type:"array",compression:true});
}
function buildMissingBytes(rows){
  if(!rows.length)return null;
  const data=[["Odoo","SAP","Kode Item","Product Name","Qty SO","Source"],...rows.map(x=>[x.odoo,x.sap,x.item,x.name,x.qty,x.source])];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),"SKU Tidak Ditemukan");
  return XLSX.write(wb,{bookType:"xlsx",type:"array",compression:true});
}
function postProgress(value,text){postMessage({type:"progress",value,text});}

onmessage=async e=>{
  try{
    if(e.data.type!=="process") return;
    postProgress(3,"Membaca Template Komper...");
    const template=XLSX.read(e.data.templateBuffer,{type:"array",cellFormula:true,cellStyles:false,cellNF:true});
    const indexes=indexTemplate(template);
    if(!indexes.length) throw Error("Tidak ditemukan sheet template dengan Odoo, SAP, atau Kode Item.");
    const maps=globalMap(indexes);
    postProgress(20,`Template siap · ${indexes.length} sheet · mode baca saja`);

    const allSO=[];
    for(let i=0;i<e.data.soBuffers.length;i++){
      postProgress(20+Math.round((i/e.data.soBuffers.length)*30),`Membaca file SO ${i+1}/${e.data.soBuffers.length}...`);
      const wb=XLSX.read(e.data.soBuffers[i],{type:"array",cellFormula:false,cellStyles:false,cellNF:false});
      allSO.push(...parseSO(wb));
    }
    if(!allSO.length) throw Error("Tidak menemukan data SO dengan header Qty.");

    const totals=new Map(), missing=[];
    for(const x of allSO){
      const c=findCandidate(x,maps);
      if(!c){missing.push(x);continue;}
      const key=c.idx.sheetName+"!"+c.row;
      const p=totals.get(key);
      if(p) p.qty+=x.qty;
      else totals.set(key,{idx:c.idx,row:c.row,qty:x.qty,first:x});
    }

    postProgress(65,"Menyusun hasil reconcile di aplikasi...");
    const resultRows=[];
    for(const t of totals.values()){
      const row=t.idx.rows[t.row]||[];
      const target=t.idx.targetCol>=0?row[t.idx.targetCol]:"";
      const qty=t.qty;
      const tn=qtyNumber(target);
      const selisih=tn===null?"":qty-tn;
      const status=tn===null?"Target tidak tersedia":selisih===0?"Sesuai":selisih>0?"Lebih":"Kurang";
      resultRows.push({
        sheet:t.idx.sheetName,
        odoo:t.idx.odooCol>=0?row[t.idx.odooCol]:"",
        sap:t.idx.sapCol>=0?row[t.idx.sapCol]:"",
        item:t.idx.itemCol>=0?row[t.idx.itemCol]:"",
        name:t.idx.nameCol>=0?row[t.idx.nameCol]:(t.first.name||""),
        target,qty,selisih,status,sourceMatched:true
      });
    }
    resultRows.sort((a,b)=>String(a.sheet).localeCompare(String(b.sheet))||String(a.item||a.sap||a.odoo).localeCompare(String(b.item||b.sap||b.odoo)));

    postProgress(82,"Membuat file hasil terpisah...");
    const resultBytes=makeResultWorkbook(resultRows);
    const missingBytes=buildMissingBytes(missing);
    const summary={total:allSO.length,matched:allSO.length-missing.length,missing:missing.length,qtyWritten:resultRows.reduce((a,x)=>a+(Number(x.qty)||0),0),matchedCells:resultRows.length,sheets:indexes.length};
    postProgress(100,"Selesai · Template Komper tidak diubah.");
    const transfers=[resultBytes.buffer]; if(missingBytes) transfers.push(missingBytes.buffer);
    postMessage({type:"done",resultBytes,missingBytes,missingRows:missing,resultRows,summary},transfers);
  }catch(err){postMessage({type:"error",message:err?.message||String(err)});}
};
