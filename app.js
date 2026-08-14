/* V10.6 performance: keep UI responsive; heavy workbook work remains in Worker. */
(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const state = { templateFile:null, soFiles:[], resultBytes:null, missingRows:[], lastResult:null, worker:null, busy:false };

  const aliases = {
    odoo:["productodooCode","productOdooCode","odooCode","odooProductCode","kodeOdoo","kodeOdooProduk"],
    sap:["productSapCode","productsapcode","sapCode","kodeSAP","kodeSap","materialCode","material"],
    item:["kodeItem","kodeBarang","itemCode","item","sku","productCode","kodeproduk"],
    name:["productName","namaProduk","productNama","description","namaBarang","product"],
    qty:["qtyFix","qtyfix","qty","quantity","qtySO","qtyOpname","hasilSO","stockQty","countQty"]
  };

  const norm = v => String(v ?? "").trim().toLowerCase().replace(/[\s_\-./()]+/g, "");
  const numericKey = v => {
    if (v === null || v === undefined || v === "") return "";
    const s = String(v).trim().replace(/,/g, "").replace(/\.0+$/, "").replace(/\s/g, "");
    return /^\d+$/.test(s) ? (s.replace(/^0+/, "") || "0") : s.toLowerCase();
  };
  const qtyNumber = v => {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(String(v).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : null;
  };

  function findHeader(rows, type, preferredRow=null) {
    const wanted = new Set(aliases[type].map(norm));
    const max = Math.min(rows.length, 40);
    if (preferredRow !== null && rows[preferredRow]) {
      for (let c=0;c<rows[preferredRow].length;c++) if (wanted.has(norm(rows[preferredRow][c]))) return {rowIndex:preferredRow,colIndex:c};
    }
    for (let r=0;r<max;r++) for (let c=0;c<(rows[r]||[]).length;c++) if (wanted.has(norm(rows[r][c]))) return {rowIndex:r,colIndex:c};
    return null;
  }
  function findExact(rows, names, preferredRow=null) {
    const wanted = new Set(names.map(norm));
    const scan = preferredRow !== null ? [preferredRow] : Array.from({length:Math.min(rows.length,40)},(_,i)=>i);
    for (const r of scan) {
      const row = rows[r] || [];
      for (let c=0;c<row.length;c++) if (wanted.has(norm(row[c]))) return {rowIndex:r,colIndex:c};
    }
    return null;
  }
  function actualRange(ws){
    const ref=ws["!ref"] || "A1:A1";
    const base=XLSX.utils.decode_range(ref);
    let maxR=base.s.r, maxC=base.s.c;
    for(const key of Object.keys(ws)){
      if(key[0]==="!") continue;
      const m=/^([A-Z]+)(\d+)$/.exec(key);
      if(!m) continue;
      const r=Number(m[2])-1, c=XLSX.utils.decode_col(m[1]);
      const cell=ws[key];
      if(cell && (cell.v!==undefined || cell.f!==undefined)) { if(r>maxR)maxR=r; if(c>maxC)maxC=c; }
    }
    return {s:{r:base.s.r,c:base.s.c},e:{r:maxR,c:maxC}};
  }
  function rowsFromSheet(ws) {
    return XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true,range:actualRange(ws)});
  }

  async function readWorkbook(file){
    const data=await file.arrayBuffer();
    return XLSX.read(data,{type:"array",cellFormula:true,cellStyles:true,cellNF:true});
  }

  function showToast(msg){ const e=$("toast"); e.textContent=msg; e.classList.add("show"); clearTimeout(showToast.t); showToast.t=setTimeout(()=>e.classList.remove("show"),3200); }
  function setStatus(text,type="neutral") { $("statusText").textContent=text; $("engineState").className="status-pill "+type; $("engineState").textContent=type==="success"?"Selesai":type==="error"?"Error":"Proses"; }
  function progress(n){ $("progressBar").style.width=Math.max(0,Math.min(100,n))+"%"; }

  function templateIndexes(wb){
    const all=[];
    for(const sheetName of wb.SheetNames){
      const ws=wb.Sheets[sheetName];
      const rows=rowsFromSheet(ws);
      if(!rows.length) continue;
      const item=findHeader(rows,"item");
      const odoo=findHeader(rows,"odoo");
      const sap=findHeader(rows,"sap");
      const name=findHeader(rows,"name");
      // Source quantity is qtyFix/quantity. Destination must be the explicit HASIL SO column.
      const hasil=findExact(rows,["HASIL SO","HASIL_STOK_OPNAME","HASILSO"]);
      const qtySource=findHeader(rows,"qty");
      // Only treat a sheet as product data when it has a code column and a plausible header row.
      const code=item||odoo||sap;
      if(!code) continue;
      const headerRow=code.rowIndex;
      const targetHeader=hasil && hasil.rowIndex===headerRow ? hasil : null;
      const targetCol=targetHeader ? targetHeader.colIndex : 4; // fallback to historical column E
      const entry={sheetName,ws,rows,headerRow,itemCol:item?.colIndex??-1,odooCol:odoo?.colIndex??-1,sapCol:sap?.colIndex??-1,nameCol:name?.colIndex??-1,targetCol, map:new Map()};
      for(let r=headerRow+1;r<rows.length;r++){
        const row=rows[r]||[];
        const keys=[];
        if(entry.odooCol>=0){const k=numericKey(row[entry.odooCol]); if(k) keys.push(["odoo",k]);}
        if(entry.sapCol>=0){const k=numericKey(row[entry.sapCol]); if(k) keys.push(["sap",k]);}
        if(entry.itemCol>=0){const k=numericKey(row[entry.itemCol]); if(k) keys.push(["item",k]);}
        for(const [kind,key] of keys){ if(!entry.map.has(kind+":"+key)) entry.map.set(kind+":"+key,{row:r,score:1}); }
      }
      all.push(entry);
    }
    return all;
  }

  function globalTemplateMap(indexes){
    const maps={odoo:new Map(),sap:new Map(),item:new Map()};
    for(const idx of indexes){
      for(const [compound,val] of idx.map.entries()){
        const [kind,key]=compound.split(":");
        if(!maps[kind].has(key)) maps[kind].set(key,[]);
        maps[kind].get(key).push({idx,row:val.row});
      }
    }
    return maps;
  }

  function findCandidate(x,maps){
    for(const kind of ["odoo","sap","item"]){
      const k=numericKey(x[kind]);
      if(k && maps[kind].has(k)) return maps[kind].get(k)[0];
    }
    // SAP may be 4000005880 while template has 5880.
    const sap=numericKey(x.sap);
    if(/^\d+$/.test(sap) && sap.length>4){
      const suffix=numericKey(sap.slice(-6));
      if(maps.item.has(suffix)) return maps.item.get(suffix)[0];
    }
    return null;
  }

  function parseSO(wb){
    const out=[];
    for(const sheetName of wb.SheetNames){
      const rows=rowsFromSheet(wb.Sheets[sheetName]);
      if(!rows.length) continue;
      const odoo=findHeader(rows,"odoo"), sap=findHeader(rows,"sap"), item=findHeader(rows,"item"), name=findHeader(rows,"name"), qty=findHeader(rows,"qty");
      if(!qty) continue;
      const headerRow=Math.max(qty.rowIndex,odoo?.rowIndex??0,sap?.rowIndex??0,item?.rowIndex??0);
      for(let r=headerRow+1;r<rows.length;r++){
        const row=rows[r]||[]; const q=qtyNumber(row[qty.colIndex]); if(q===null) continue;
        const o=odoo?row[odoo.colIndex]:"", s=sap?row[sap.colIndex]:"", i=item?row[item.colIndex]:"";
        if(!numericKey(o)&&!numericKey(s)&&!numericKey(i)) continue;
        out.push({odoo:o,sap:s,item:i,name:name?row[name.colIndex]:"",qty:q,source:sheetName});
      }
    }
    return out;
  }

  function setCell(ws,r,c,value){
    const addr=XLSX.utils.encode_cell({r,c});
    const old=ws[addr];
    if(old){ old.v=value; old.t="n"; if(old.f) delete old.f; }
    else ws[addr]={t:"n",v:value};
  }

  function cloneWorkbook(wb){
    const bytes=XLSX.write(wb,{bookType:"xlsx",type:"array",compression:true});
    return XLSX.read(bytes,{type:"array",cellFormula:true,cellStyles:true,cellNF:true});
  }

  function renderResults(rows){
    const body=$("resultBody");
    if(!body) return;
    if(!rows || !rows.length){
      body.innerHTML='<tr><td colspan="7" class="empty-result">Tidak ada SKU yang cocok.</td></tr>';
      return;
    }
    const esc=v=>String(v??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
    body.innerHTML=rows.slice(0,1000).map(x=>{
      const code=x.item||x.sap||x.odoo||"-";
      return `<tr><td>${esc(x.sheet)}</td><td>${esc(code)}</td><td>${esc(x.name||"-")}</td><td>${esc(x.odoo||"-")}</td><td>${esc(x.sap||"-")}</td><td>${esc(x.qty)}</td><td><span class="table-status neutral">${esc(x.status||"Cocok")}</span></td></tr>`;
    }).join("");
  }

  function ensureWorker(){
    if(state.worker) return state.worker;
    state.worker = new Worker("reconcile-worker.js");
    state.worker.onmessage = e => {
      const m=e.data;
      if(m.type==="progress"){ progress(m.value); setStatus(m.text); return; }
      if(m.type==="done"){
        state.busy=false;
        state.resultBytes=m.resultBytes;
        state.missingRows=m.missingRows||[];
        state.lastResult=m.summary;
        renderResults(m.resultRows||[]);
        state.missingBytes=m.missingBytes||null;
        $("totalSo").textContent=m.summary.total.toLocaleString("id-ID");
        $("matched").textContent=m.summary.matched.toLocaleString("id-ID");
        $("notFound").textContent=m.summary.missing.toLocaleString("id-ID");
        $("qtyWritten").textContent=m.summary.qtyWritten.toLocaleString("id-ID");
        $("downloadBtn").disabled=false;
        $("downloadMissingBtn").disabled=!state.missingRows.length;
        $("resultState").className="status-pill success";
        $("resultState").textContent="Berhasil";
        progress(100);
        setStatus(`Selesai · ${m.summary.sheets} sheet · ${m.summary.matchedCells.toLocaleString("id-ID")} item terisi`,"success");
        saveHistory(m.summary);
        showToast("Hasil siap di-download.");
        return;
      }
      if(m.type==="missingDone"){
        state.missingBytes=m.bytes;
        downloadBlob(new Blob([m.bytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
          `SKU_Tidak_Ditemukan_${new Date().toISOString().slice(0,10)}.xlsx`);
        return;
      }
      if(m.type==="error"){
        state.busy=false;
        setStatus(m.message,"error"); showToast("Gagal: "+m.message);
      }
    };
    state.worker.onerror = e => { state.busy=false; setStatus("Worker gagal dijalankan. Coba refresh halaman.","error"); showToast("Proses gagal dijalankan."); };
    return state.worker;
  }

  async function processReconcile(){
    if(state.busy) return showToast("Proses masih berjalan.");
    if(!state.templateFile) throw Error("Template belum dipilih.");
    if(!state.soFiles.length) throw Error("File SO belum dipilih.");

    state.busy=true; state.resultBytes=null; state.missingRows=[];
    $("downloadBtn").disabled=true; $("downloadMissingBtn").disabled=true;
    $("resultState").className="status-pill neutral"; $("resultState").textContent="Memproses";
    progress(2); setStatus("Menyiapkan proses di background...");
    const templateBuffer=await state.templateFile.arrayBuffer();
    const soBuffers=[];
    for(const f of state.soFiles) soBuffers.push(await f.arrayBuffer());
    ensureWorker().postMessage({
      type:"process",
      templateBuffer,
      soBuffers,
      templateName:state.templateFile.name
    }, [templateBuffer, ...soBuffers]);
  }

  function downloadBlob(blob,name){
    const u=URL.createObjectURL(blob),a=document.createElement("a");
    a.href=u;a.download=name;a.rel="noopener";
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(u),5000);
  }

  function downloadResult(){
    if(!state.resultBytes)return showToast("Belum ada hasil.");
    downloadBlob(new Blob([state.resultBytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
      `Hasil_Reconcile_SO_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  function downloadMissing(){
    if(!state.missingRows.length)return showToast("Tidak ada SKU yang tidak ditemukan.");
    if(state.missingBytes){
      downloadBlob(new Blob([state.missingBytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
        `SKU_Tidak_Ditemukan_${new Date().toISOString().slice(0,10)}.xlsx`);
    } else {
      showToast("File daftar tidak ditemukan belum siap.");
    }
  }

  function saveHistory(result){ const h=JSON.parse(localStorage.getItem("rsp_v10_history")||"[]"); h.unshift({date:new Date().toLocaleString("id-ID"),...result}); localStorage.setItem("rsp_v10_history",JSON.stringify(h.slice(0,20))); renderHistory(); }
  function renderHistory(){ const h=JSON.parse(localStorage.getItem("rsp_v10_history")||"[]"); $("historyList").innerHTML=h.map(x=>`<div class="history-item"><strong>${x.date}</strong><br><span class="muted">${x.total} SO · ${x.matched} cocok · ${x.missing} tidak ditemukan · ${x.matchedCells||0} item terisi · Qty ${x.qtyWritten}</span></div>`).join("")||'<div class="empty-history">Belum ada riwayat.</div>'; }

  function reset(){ state.templateFile=null;state.soFiles=[];state.resultBytes=null;state.missingBytes=null;state.missingRows=[];state.lastResult=null;state.busy=false; $("templateInput").value="";$("soInput").value="";$("templateInfo").textContent="Belum ada template";$("soInfo").textContent="Belum ada file SO";$("downloadBtn").disabled=true;$("downloadMissingBtn").disabled=true;["totalSo","matched","notFound","qtyWritten"].forEach(id=>$(id).textContent="0");$("resultState").textContent="Belum diproses";$("resultState").className="status-pill neutral";progress(0);setStatus("Upload template dan file SO untuk memulai."); }

  function init(){
    $("templateInput").addEventListener("change",e=>{const f=e.target.files[0];if(!f)return;try{state.templateFile=f;state.resultBytes=null;state.missingBytes=null;$("templateInfo").textContent=`✓ ${f.name} · ${(f.size/1024).toFixed(1)} KB`;setStatus("Template siap. Akan diproses di background.");}catch(err){showToast("Gagal membaca template: "+err.message);}});
    $("soInput").addEventListener("change",e=>{state.soFiles=[...e.target.files];$("soInfo").textContent=state.soFiles.length?`✓ ${state.soFiles.length} file dipilih`:`Belum ada file SO`;});
    $("processBtn").addEventListener("click",async()=>{try{await processReconcile();}catch(err){console.error(err);setStatus(err.message,"error");showToast(err.message);}});
    $("downloadBtn").addEventListener("click",downloadResult); $("downloadMissingBtn").addEventListener("click",downloadMissing); $("clearBtn").addEventListener("click",reset);
    $("calcOverstock").addEventListener("click",()=>{const a=Number($("ovMonth1").value||0),b=Number($("ovMonth2").value||0),f=Number($("ovFactor").value||1.5),t=Number($("ovTarget").value||0),v=(a+b)*f,sel=v>t;$("ovResult").innerHTML=`<strong>${$("ovBrand").value||"Brand"}</strong><br>(Bulan 1 + Bulan 2) × ${f.toFixed(1)} = <b>${v.toLocaleString("id-ID")}</b><br>Target sekarang: <b>${t.toLocaleString("id-ID")}</b><br><span class="${sel?'danger':'good'}">${sel?'TERINDIKASI OVERSTOK':'Tidak melebihi target'}</span>`;});
    const drawer=$("drawer"),back=$("drawerBackdrop"); const open=()=>{drawer.classList.add("open");back.classList.remove("hidden")},close=()=>{drawer.classList.remove("open");back.classList.add("hidden")}; $("menuBtn").onclick=open;$("closeMenu").onclick=close;back.onclick=close;
    document.querySelectorAll(".nav-item").forEach(btn=>btn.addEventListener("click",()=>{document.querySelectorAll(".nav-item").forEach(x=>x.classList.remove("active"));btn.classList.add("active");document.querySelectorAll(".section").forEach(s=>s.classList.add("hidden"));$("section-"+btn.dataset.section).classList.remove("hidden");close();}));
    renderHistory();
  }
  window.addEventListener("DOMContentLoaded",init);
})();
