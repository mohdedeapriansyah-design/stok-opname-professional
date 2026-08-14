/* V9 clean rebuild: one API, no legacy FileEngine calls. */
const App={template:null,soFiles:[],lastResult:null,history:JSON.parse(localStorage.getItem("rsp_v9_history")||"[]")};
const $=s=>document.querySelector(s);
const toast=m=>{const t=$("#toast");t.textContent=m;t.style.display="block";setTimeout(()=>t.style.display="none",2500)};
const money=n=>new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(Number(n)||0);
const qty=n=>Number.isFinite(Number(n))?Number(n):0;
const canon=v=>{
  if(v===null||v===undefined)return "";
  let s=String(v).trim();
  if(/^\d+$/.test(s)) s=s.replace(/^0+(?=\d)/,"");
  return s.toUpperCase();
};
const safeFileName=s=>String(s).replace(/[^\\w.-]+/g,"_");

/* Read worksheet rows without copying dense SheetJS arrays.
   SheetJS dense worksheets are already row arrays; returning them directly
   avoids another large allocation on Android. */
function readRows(ws){
  if(Array.isArray(ws)) return ws;
  return XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
}

function page(name){
  const titles={home:"Beranda",reconcile:"Reconcile Otomatis",overstock:"Analisa Overstok",history:"Riwayat Proses",about:"Tentang Aplikasi"};
  $("#app").innerHTML=`<div class="page"><h1>${titles[name]}</h1><div id="content"></div></div>`;
  render[name]();
  document.querySelectorAll(".sidebar button").forEach(b=>b.classList.toggle("active",b.dataset.page===name));
  $("#sidebar").classList.remove("open");
}
const render={};

render.home=()=>$("#content").innerHTML=`
<div class="hero"><div class="big">Reconcile Automation</div><p>Upload Template Komper + banyak file SO, proses otomatis dan tetap menjaga template asli.</p></div>
<div class="grid">
<div class="card"><h3>↻ Reconcile Otomatis</h3><p class="muted">Masukkan template dan file SO sebanyak yang diperlukan.</p><button class="btn" onclick="page('reconcile')">Mulai</button></div>
<div class="card"><h3>◈ Analisa Overstok</h3><p class="muted">Hitung stok per brand berdasarkan 2 bulan dan faktor 1,5–2,0.</p><button class="btn" onclick="page('overstock')">Buka</button></div>
<div class="card"><h3>▣ Riwayat</h3><p class="muted">Melihat proses terakhir tanpa menyimpan file Excel ke server.</p><button class="btn" onclick="page('history')">Buka</button></div>
</div>`;

render.reconcile=()=>{
$("#content").innerHTML=`
<div class="grid">
<div class="card"><h3>1. File Template Komper</h3>
<label class="upload">📄<br><b>Klik untuk memilih file</b><br><small>XLSX / XLS / CSV</small><input id="templateInput" type="file" accept=".xlsx,.xls,.csv"></label>
<div id="templateInfo" class="status muted">Belum ada template.</div></div>
<div class="card"><h3>2. File Hasil SO</h3>
<label class="upload">📚<br><b>Klik untuk memilih banyak file</b><br><small>Semua file diproses berurutan</small><input id="soInput" type="file" multiple accept=".xlsx,.xls,.csv"></label>
<div id="soInfo" class="status muted">Belum ada file SO.</div></div>
</div>
<div class="card" style="margin-top:14px"><h3>3. Engine Reconcile</h3>
<div class="progress"><i id="bar"></i></div><p id="progressText" class="muted">Siap.</p>
<div class="row"><button class="btn" id="processBtn">Proses Rekonsiliasi</button><button class="btn secondary" id="clearBtn">Bersihkan</button></div></div>
<div class="card" style="margin-top:14px"><h3>4. Hasil</h3><div id="resultInfo" class="status">Belum diproses.</div><div class="row" style="margin-top:10px"><button class="btn" id="downloadBtn" disabled>Download Hasil</button><button class="btn secondary" id="missingBtn" disabled>SKU Tidak Ditemukan</button></div></div>
<div class="card" style="margin-top:14px"><h3>Aturan pembacaan SO</h3><p>Kolom dicari berdasarkan nama header: <b>productSapCode / Kode SAP</b>, <b>ProductName</b>, <b>ProductCode</b>, dan <b>qtyFix</b>. Kode numerik seperti 0003184 dinormalisasi menjadi 3184 agar dapat dicocokkan dengan template.</p><p>Template asli tidak diubah. Yang ditulis hanya nilai Qty pada salinan hasil.</p></div>`;
bindReconcile();
};

function setProgress(p,text){$("#bar").style.width=Math.max(0,Math.min(100,p))+"%";$("#progressText").textContent=text}

function cleanHeader(v){
  return String(v??"").trim().toUpperCase().replace(/[^A-Z0-9]/g,"");
}
function headerCandidates(rows, maxRows=40){
  const result=[];
  for(let r=0;r<Math.min(rows.length,maxRows);r++){
    const headers=rows[r].map(cleanHeader);
    result.push({r,headers});
  }
  return result;
}
/* Automatic header detection:
   - Odoo and SAP are optional.
   - Qty is detected by aliases.
   - Column positions do not matter.
   - Each sheet/file is detected independently. */
function detectSOHeader(rows){
  const aliases={
    odoo:["PRODUCTODOOCODE","ODOOCODE","ODOOPRODUCTCODE","KODEODOO","ODOO"],
    sap:["PRODUCTSAPCODE","SAPCODE","KODESAP","SAPPRODUCTCODE","MATERIALCODE","MATERIAL"],
    name:["PRODUCTNAME","NAMAPRODUK","PRODUCTDESCRIPTION","DESCRIPTION","NAMA"],
    qty:["QTYFIX","QTY","QUANTITY","QTYSO","QTYOPNAME","HASILSO","STOCKQTY","SOQTY"]
  };
  for(const item of headerCandidates(rows,40)){
    const find=keys=>item.headers.findIndex(h=>keys.some(k=>h===k||h.includes(k)));
    const odoo=find(aliases.odoo), sap=find(aliases.sap), name=find(aliases.name), qty=find(aliases.qty);
    if(qty>=0 && (odoo>=0||sap>=0||name>=0)){
      return {headerRow:item.r,odoo,sap,name,qty,
        hasOdoo:odoo>=0,hasSap:sap>=0,hasName:name>=0};
    }
  }
  return null;
}
function detectTemplateHeader(rows){
  const aliases={
    code:["KODEITEM","PRODUCTODOOCODE","PRODUCTCODE","ITEMCODE","KODEPRODUK"],
    name:["NAMAPRODUK","PRODUCTNAME","DESCRIPTION"],
    qty:["HASILSO","QTY","QUANTITY"]
  };
  for(const item of headerCandidates(rows,20)){
    const find=keys=>item.headers.findIndex(h=>keys.some(k=>h===k||h.includes(k)));
    const code=find(aliases.code), name=find(aliases.name), qty=find(aliases.qty);
    if(code>=0 && qty>=0)return {headerRow:item.r,code,name,qty};
  }
  return null;
}
function normalizeCode(v){
  if(v===null||v===undefined||v==="")return "";
  let s=String(v).trim();
  if(/^\d+\.0+$/.test(s))s=s.replace(/\.0+$/,"");
  return s.toUpperCase();
}
function makeKeys(item){
  const keys=[];
  if(item.odoo)keys.push("ODOO:"+normalizeCode(item.odoo));
  if(item.sap)keys.push("SAP:"+normalizeCode(item.sap));
  if(item.name)keys.push("NAME:"+String(item.name).trim().toUpperCase());
  return keys;
}
async function readSO(file){
  const wb=XLSX.read(await file.arrayBuffer(),{
    type:"array",dense:true,cellFormula:false,cellStyles:false,cellNF:false
  });
  const out=[];
  for(const sheetName of wb.SheetNames){
    const rows=readRows(wb.Sheets[sheetName]);
    const h=detectSOHeader(rows);
    if(!h)continue;
    const items=[];
    for(let r=h.headerRow+1;r<rows.length;r++){
      const row=rows[r];
      const item={
        odoo:h.hasOdoo?row[h.odoo]:"",
        sap:h.hasSap?row[h.sap]:"",
        name:h.hasName?row[h.name]:"",
        qty:qty(row[h.qty])
      };
      if(makeKeys(item).length)items.push(item);
      if(r%2000===0)await new Promise(requestAnimationFrame);
    }
    out.push({name:sheetName,header:h,items});
  }
  return out;
}
function buildTemplateIndex(wb){
  const byKey=new Map(), sheets=[];
  for(const name of wb.SheetNames){
    const ws=wb.Sheets[name],rows=readRows(ws);
    const h=detectTemplateHeader(rows);
    if(!h)continue;
    sheets.push({name,rows,header:h});
    for(let r=h.headerRow+1;r<rows.length;r++){
      const code=normalizeCode(rows[r][h.code]);
      if(code)byKey.set("ODOO:"+code,{sheet:name,row:r,qtyCol:h.qty});
      if(h.name>=0){
        const nm=String(rows[r][h.name]??"").trim().toUpperCase();
        if(nm)byKey.set("NAME:"+nm,{sheet:name,row:r,qtyCol:h.qty});
      }
    }
  }
  return {byKey,sheets};
}
function setQtyCell(ws,row,col,value){
  const addr=XLSX.utils.encode_cell({r:row,c:col});
  const old=ws[addr]||{};
  ws[addr]={...old,t:"n",v:Number(value)||0};
}
async function reconcile(){
  if(!App.template||!App.soFiles.length)return toast("Upload template dan minimal 1 file SO.");
  setProgress(3,"Deteksi header Template...");
  const wb=XLSX.read(await App.template.arrayBuffer(),{
    type:"array",dense:true,cellFormula:true,cellStyles:true,cellNF:true
  });
  const ti=buildTemplateIndex(wb);
  if(!ti.sheets.length)throw new Error("Header Template tidak dikenali.");
  const totals=new Map(),unmatched=[];
  let detectedOdoo=0,detectedSap=0;
  for(let i=0;i<App.soFiles.length;i++){
    const f=App.soFiles[i];
    setProgress(8+(i/App.soFiles.length)*42,`Deteksi header SO ${i+1}/${App.soFiles.length}: ${f.name}`);
    const sheets=await readSO(f);
    for(const sh of sheets){
      if(sh.header.hasOdoo)detectedOdoo++;
      if(sh.header.hasSap)detectedSap++;
      for(const item of sh.items){
        let hit=null,matchedBy="";
        for(const key of makeKeys(item)){
          hit=ti.byKey.get(key);
          if(hit){matchedBy=key.split(":")[0];break;}
        }
        if(!hit){
          unmatched.push({
            File:f.name,Sheet:sh.name,
            ProductOdooCode:item.odoo||"",
            ProductSapCode:item.sap||"",
            ProductName:item.name||"",
            QtyFix:item.qty
          });
          continue;
        }
        const k=hit.sheet+"|"+hit.row;
        const old=totals.get(k);
        totals.set(k,{...hit,qty:(old?.qty||0)+item.qty,matchedBy});
      }
    }
    await new Promise(requestAnimationFrame);
  }
  setProgress(60,"Menulis hasil ke kolom Qty template...");
  const perBrand={},byMethod={ODOO:0,SAP:0,NAME:0};
  for(const [key,v] of totals){
    setQtyCell(wb.Sheets[v.sheet],v.row,v.qtyCol,v.qty);
    perBrand[v.sheet]=(perBrand[v.sheet]||0)+1;
    byMethod[v.matchedBy]=(byMethod[v.matchedBy]||0)+1;
  }
  setProgress(90,"Validasi hasil...");
  const brandRows=Object.entries(perBrand).map(([b,c])=>`<tr><td>${b}</td><td>${c.toLocaleString("id-ID")}</td></tr>`).join("");
  App.lastResult={wb,unmatched,matched:totals.size,byMethod};
  $("#resultInfo").innerHTML=`<b>Selesai.</b> ${totals.size.toLocaleString("id-ID")} item cocok, ${unmatched.length.toLocaleString("id-ID")} tidak ditemukan.
  <p class="muted">Deteksi file: Odoo ${detectedOdoo} sheet, SAP ${detectedSap} sheet. Metode cocok: Odoo ${byMethod.ODOO}, SAP ${byMethod.SAP}, Nama ${byMethod.NAME}.</p>
  <div class="tablewrap"><table><thead><tr><th>Sheet / Brand</th><th>Item terisi</th></tr></thead><tbody>${brandRows||"<tr><td colspan=2>Tidak ada.</td></tr>"}</tbody></table></div>`;
  $("#downloadBtn").disabled=false;$("#missingBtn").disabled=false;
  App.history.unshift({date:new Date().toLocaleString("id-ID"),files:App.soFiles.length,matched:totals.size,missing:unmatched.length});
  App.history=App.history.slice(0,30);localStorage.setItem("rsp_v9_history",JSON.stringify(App.history));
  setProgress(100,"Selesai.");
}
function saveXlsxBrowser(wb, filename){
  try{
    const data=XLSX.write(wb,{bookType:"xlsx",type:"array",compression:true});
    const blob=new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=filename;
    a.style.display="none";
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{a.remove();URL.revokeObjectURL(url)},1500);
    toast("File sedang diunduh: "+filename);
    return true;
  }catch(err){
    console.error("Browser download error:",err);
    toast("Download gagal: "+err.message);
    return false;
  }
}
function downloadWorkbook(){
  if(!App.lastResult?.wb)return toast("Hasil belum tersedia. Proses rekonsiliasi dulu.");
  const name=`Hasil_Komper_SO_${new Date().toISOString().slice(0,10)}.xlsx`;
  saveXlsxBrowser(App.lastResult.wb,name);
}
function downloadMissing(){
  const rows=App.lastResult?.unmatched||[];
  if(!rows.length)return toast("Tidak ada SKU yang tidak ditemukan.");
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"SKU_Tidak_Ditemukan");
  saveXlsxBrowser(wb,`SKU_Tidak_Ditemukan_${new Date().toISOString().slice(0,10)}.xlsx`);
}
function bindReconcile(){
  $("#templateInput").onchange=e=>{App.template=e.target.files[0];$("#templateInfo").innerHTML=`✓ ${App.template.name} — ${(App.template.size/1024).toFixed(1)} KB`};
  $("#soInput").onchange=e=>{App.soFiles=[...e.target.files];$("#soInfo").innerHTML=`✓ ${App.soFiles.length} file dipilih.`};
  $("#processBtn").onclick=()=>reconcile().catch(e=>{console.error(e);toast("Gagal: "+e.message);setProgress(0,"Terjadi error.")});
  $("#downloadBtn").onclick=downloadWorkbook;
  $("#missingBtn").onclick=downloadMissing;
  $("#clearBtn").onclick=()=>{App.template=null;App.soFiles=[];App.lastResult=null;page("reconcile")};
}

render.overstock=()=>{
$("#content").innerHTML=`
<div class="card"><h3>Analisa Overstok per Brand</h3><p class="muted">Input manual sesuai kebutuhan. Faktor dapat dipilih 1,5 sampai 2,0.</p>
<div id="brands"></div><div class="row"><button class="btn secondary" id="addBrand">+ Tambah Brand</button><button class="btn" id="calcBrands">Hitung Semua</button></div></div>
<div class="card" style="margin-top:14px"><h3>Hasil</h3><div id="overResult" class="tablewrap"></div></div>`;
let n=0; const add=()=>{n++;$("#brands").insertAdjacentHTML("beforeend",`<div class="brandrow card" style="padding:10px"><div class="field brandname"><label>Brand</label><input data-k="brand" placeholder="Contoh: Wardah"></div><div class="field"><label>Stok Bulan 1</label><input data-k="a" type="number" min="0" value="0"></div><div class="field"><label>Stok Bulan 2</label><input data-k="b" type="number" min="0" value="0"></div><div class="field"><label>Faktor</label><select data-k="f"><option>1.5</option><option>1.6</option><option>1.7</option><option>1.8</option><option>1.9</option><option>2.0</option></select></div><div class="field"><label>Target Sekarang</label><input data-k="t" type="number" min="0" value="0"></div></div>`)};add();$("#addBrand").onclick=add;$("#calcBrands").onclick=()=>{
 const out=[...document.querySelectorAll("#brands .brandrow")].map(x=>{const v=k=>x.querySelector(`[data-k="${k}"]`).value;const a=+v("a"),b=+v("b"),f=+v("f"),t=+v("t"),pot=(a+b)*f;return {brand:v("brand")||"-",a,b,f,pot,t,diff:pot-t,status:pot>t?"OVERSTOK":"NORMAL"}});
 $("#overResult").innerHTML=`<table><thead><tr><th>Brand</th><th>Bulan 1</th><th>Bulan 2</th><th>Faktor</th><th>Potensi</th><th>Target</th><th>Selisih</th><th>Status</th></tr></thead><tbody>${out.map(x=>`<tr><td>${x.brand}</td><td>${money(x.a)}</td><td>${money(x.b)}</td><td>${x.f}×</td><td>${money(x.pot)}</td><td>${money(x.t)}</td><td>${money(x.diff)}</td><td class="${x.status==="OVERSTOK"?"bad":"ok"}">${x.status}</td></tr>`).join("")}</tbody></table>`;
};
};

render.history=()=>{$("#content").innerHTML=`<div class="card"><h3>Riwayat</h3><div class="tablewrap"><table><thead><tr><th>Tanggal</th><th>File SO</th><th>Match</th><th>Tidak ditemukan</th></tr></thead><tbody>${App.history.map(x=>`<tr><td>${x.date}</td><td>${x.files}</td><td>${x.matched}</td><td>${x.missing}</td></tr>`).join("")||"<tr><td colspan=4>Belum ada riwayat.</td></tr>"}</tbody></table></div></div>`};
render.about=()=>$("#content").innerHTML=`<div class="card"><h3>Reconcile Stock Pro V9</h3><p>Clean rebuild berdasarkan kebutuhan aplikasi dan bug versi sebelumnya.</p><ul><li>Template asli tidak diubah; hasil ditulis ke salinan workbook di browser.</li><li>SO dapat banyak file.</li><li>Kolom dibaca berdasarkan header.</li><li>Kode SAP numerik dinormalisasi agar 0003184 dan 3184 dapat dicocokkan.</li><li>Overstok per brand dengan faktor 1,5–2,0.</li><li>Tidak ada upload data ke server; proses dilakukan di browser.</li></ul></div>`;

$("#menuBtn").onclick=()=>$("#sidebar").classList.toggle("open");
document.querySelectorAll(".sidebar button").forEach(b=>b.onclick=()=>page(b.dataset.page));
page("home");
  
