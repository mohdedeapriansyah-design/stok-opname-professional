(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const state = {
    templateFile: null,
    soFiles: [],
    templateWorkbook: null,
    templateSheetName: null,
    resultWorkbook: null,
    missingRows: [],
    lastResult: null
  };

  const aliases = {
    odoo: ["productodooCode","odooCode","odooProductCode","kodeOdoo","kodeOdooProduk"],
    sap: ["productsapCode","sapCode","kodeSAP","kodeSap","materialCode","material"],
    item: ["kodeItem","kodeBarang","itemCode","item","sku","productCode","productSapCode","productOdooCode"],
    name: ["productName","namaProduk","productNama","description","namaBarang","product"],
    qty: ["qtyFix","qty","quantity","qtySO","qtyOpname","hasilSO","stockQty","countQty"]
  };

  function norm(v) {
    return String(v ?? "").trim().toLowerCase().replace(/[\s_\-./()]+/g, "");
  }

  function numericKey(v) {
    if (v === null || v === undefined || v === "") return "";
    const s = String(v).trim().replace(/\.0+$/,"").replace(/\s/g,"");
    return s.replace(/^0+/, "") || "0";
  }

  function findHeader(row, type) {
    const wanted = new Set(aliases[type].map(norm));
    for (let i = 0; i < row.length; i++) {
      if (wanted.has(norm(row[i]))) return i;
    }
    return -1;
  }

  function findHeaderInRows(rows, type) {
    const limit = Math.min(rows.length, 30);
    for (let r = 0; r < limit; r++) {
      const idx = findHeader(rows[r], type);
      if (idx >= 0) return { rowIndex: r, colIndex: idx };
    }
    return null;
  }

  function rowsFromSheet(ws) {
    return XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:true});
  }

  function cloneWorkbook(wb) {
    // SheetJS creates a separate workbook when reading/writing.
    // The original uploaded file is never modified.
    return XLSX.read(XLSX.write(wb,{bookType:"xlsx",type:"array"}), {type:"array", cellFormula:true, cellStyles:true});
  }

  function showToast(msg) {
    const el = $("toast"); el.textContent = msg; el.classList.add("show");
    clearTimeout(showToast.t); showToast.t = setTimeout(() => el.classList.remove("show"), 3200);
  }

  function setStatus(text, type="neutral") {
    $("statusText").textContent = text;
    $("engineState").className = "status-pill " + type;
    $("engineState").textContent = type === "success" ? "Selesai" : type === "error" ? "Error" : "Proses";
  }

  function progress(n) { $("progressBar").style.width = Math.max(0,Math.min(100,n)) + "%"; }

  async function readWorkbook(file) {
    const data = await file.arrayBuffer();
    return XLSX.read(data, {type:"array", cellFormula:true, cellStyles:true, cellNF:true});
  }

  function firstSheet(wb) { return wb.SheetNames[0]; }

  function getTemplateIndex(rows) {
    const header = findHeaderInRows(rows,"item") || findHeaderInRows(rows,"odoo") || findHeaderInRows(rows,"sap");
    const odoo = findHeaderInRows(rows,"odoo");
    const sap = findHeaderInRows(rows,"sap");
    const item = findHeaderInRows(rows,"item");
    return {header, odoo, sap, item};
  }

  function buildTemplateMaps(rows) {
    const h = getTemplateIndex(rows);
    if (!h.header) throw new Error("Header Kode Item/Odoo/SAP tidak ditemukan pada template.");
    const headerRow = h.header.rowIndex;
    const maps = {odoo:new Map(), sap:new Map(), item:new Map()};
    const itemCol = h.item?.colIndex ?? -1;
    const odooCol = h.odoo?.colIndex ?? -1;
    const sapCol = h.sap?.colIndex ?? -1;

    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const item = itemCol >= 0 ? numericKey(row[itemCol]) : "";
      const odoo = odooCol >= 0 ? numericKey(row[odooCol]) : "";
      const sap = sapCol >= 0 ? numericKey(row[sapCol]) : "";
      if (item) maps.item.set(item, r);
      if (odoo) maps.odoo.set(odoo, r);
      if (sap) maps.sap.set(sap, r);
    }
    return {headerRow, itemCol, maps};
  }

  function findTargetRow(code, maps) {
    const key = numericKey(code);
    if (!key) return -1;
    if (maps.odoo.has(key)) return maps.odoo.get(key);
    if (maps.sap.has(key)) return maps.sap.get(key);
    if (maps.item.has(key)) return maps.item.get(key);

    // Safe SAP -> Kode Item fallback:
    // only accept suffix when it matches exactly one template item.
    if (/^\d+$/.test(key) && key.length > 4) {
      const suffix = key.replace(/^0+/, "");
      const candidates = [];
      for (const [item, row] of maps.item.entries()) {
        if (item === suffix || key.endsWith(item)) candidates.push(row);
      }
      if (candidates.length === 1) return candidates[0];
    }
    return -1;
  }

  function parseSO(wb) {
    const output = [];
    for (const sheetName of wb.SheetNames) {
      const rows = rowsFromSheet(wb.Sheets[sheetName]);
      if (!rows.length) continue;
      const hOdoo = findHeaderInRows(rows,"odoo");
      const hSap = findHeaderInRows(rows,"sap");
      const hItem = findHeaderInRows(rows,"item");
      const hName = findHeaderInRows(rows,"name");
      const hQty = findHeaderInRows(rows,"qty");
      if (!hQty) continue;
      const headerRow = Math.max(hQty.rowIndex, hOdoo?.rowIndex ?? 0, hSap?.rowIndex ?? 0, hItem?.rowIndex ?? 0);
      for (let r = headerRow + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const qtyRaw = hQty.colIndex >= 0 ? row[hQty.colIndex] : "";
        if (qtyRaw === "" || qtyRaw === null || qtyRaw === undefined) continue;
        const qty = Number(String(qtyRaw).replace(/,/g,""));
        if (!Number.isFinite(qty)) continue;
        output.push({
          odoo: hOdoo ? row[hOdoo.colIndex] : "",
          sap: hSap ? row[hSap.colIndex] : "",
          item: hItem ? row[hItem.colIndex] : "",
          name: hName ? row[hName.colIndex] : "",
          qty,
          source: sheetName
        });
      }
    }
    return output;
  }

  function setCellValue(ws, row, col, value) {
    const addr = XLSX.utils.encode_cell({r:row,c:col});
    if (!ws[addr]) ws[addr] = {t:"n",v:value};
    else { ws[addr].v = value; ws[addr].t = "n"; }
  }

  function detectQtyColumn(rows, headerRow) {
    const direct = findHeaderInRows(rows,"qty");
    return direct && direct.rowIndex === headerRow ? direct.colIndex : (direct?.colIndex ?? -1);
  }

  function processReconcile() {
    if (!state.templateWorkbook) throw new Error("Template belum dipilih.");
    if (!state.soFiles.length) throw new Error("File SO belum dipilih.");

    progress(5); setStatus("Membaca template...");
    const templateSheet = state.templateSheetName;
    const ws = state.templateWorkbook.Sheets[templateSheet];
    const rows = rowsFromSheet(ws);
    const ti = buildTemplateMaps(rows);
    const qtyCol = detectQtyColumn(rows, ti.headerRow);

    // Preferred target column: an existing Qty/hasil SO column. If absent, use column E (4),
    // matching the historical Komper requirement.
    const targetCol = qtyCol >= 0 ? qtyCol : 4;
    const soRows = [];
    for (let i=0; i<state.soFiles.length; i++) {
      progress(10 + Math.round((i/state.soFiles.length)*35));
      const wb = state.soFiles[i];
      soRows.push(...parseSO(wb));
    }

    progress(50); setStatus("Mencocokkan Odoo / SAP / Kode Item...");
    const missing = [];
    let matched = 0, qtyWritten = 0;
    for (let i=0; i<soRows.length; i++) {
      const x = soRows[i];
      const row = findTargetRow(x.odoo, ti.maps);
      const finalRow = row >= 0 ? row : findTargetRow(x.sap, ti.maps);
      const finalRow2 = finalRow >= 0 ? finalRow : findTargetRow(x.item, ti.maps);
      if (finalRow2 < 0) {
        missing.push(x); continue;
      }
      setCellValue(ws, finalRow2, targetCol, x.qty);
      matched++; qtyWritten += x.qty;
      if (i % 250 === 0) progress(50 + Math.round((i/Math.max(1,soRows.length))*45));
    }

    progress(98); setStatus("Membuat file hasil...");
    state.resultWorkbook = cloneWorkbook(state.templateWorkbook);
    // Re-apply the modified worksheet to the cloned workbook so the downloaded workbook contains the edits.
    state.resultWorkbook.Sheets[templateSheet] = ws;
    state.missingRows = missing;
    state.lastResult = {total:soRows.length, matched, missing:missing.length, qtyWritten};
    $("totalSo").textContent = soRows.length.toLocaleString("id-ID");
    $("matched").textContent = matched.toLocaleString("id-ID");
    $("notFound").textContent = missing.length.toLocaleString("id-ID");
    $("qtyWritten").textContent = qtyWritten.toLocaleString("id-ID");
    $("downloadBtn").disabled = false;
    $("downloadMissingBtn").disabled = missing.length === 0;
    $("resultState").className = "status-pill success";
    $("resultState").textContent = "Berhasil";
    progress(100); setStatus("Selesai. Hasil berada di salinan template.", "success");
    saveHistory(state.lastResult);
    showToast("Rekonsiliasi selesai.");
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.rel = "noopener";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  }

  function downloadResult() {
    if (!state.resultWorkbook) return showToast("Belum ada hasil.");
    const out = XLSX.write(state.resultWorkbook, {bookType:"xlsx", type:"array", compression:true});
    downloadBlob(new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), `Hasil_Komper_SO_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  function downloadMissing() {
    if (!state.missingRows.length) return showToast("Tidak ada SKU yang tidak ditemukan.");
    const data = state.missingRows.map(x => ({Odoo:x.odoo, SAP:x.sap, Kode_Item:x.item, ProductName:x.name, QtyFix:x.qty, Sheet:x.source}));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "SKU Tidak Ditemukan");
    const out = XLSX.write(wb, {bookType:"xlsx",type:"array",compression:true});
    downloadBlob(new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), `SKU_Tidak_Ditemukan_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  function saveHistory(result) {
    const h = JSON.parse(localStorage.getItem("rsp_v10_history") || "[]");
    h.unshift({date:new Date().toLocaleString("id-ID"), ...result});
    localStorage.setItem("rsp_v10_history", JSON.stringify(h.slice(0,20)));
  }

  function renderHistory() {
    const h = JSON.parse(localStorage.getItem("rsp_v10_history") || "[]");
    $("historyList").innerHTML = h.map(x => `<div class="history-item"><strong>${x.date}</strong><br><span class="muted">${x.total} SO · ${x.matched} cocok · ${x.missing} tidak ditemukan · Qty ${x.qtyWritten}</span></div>`).join("");
  }

  function reset() {
    state.templateFile=null; state.soFiles=[]; state.templateWorkbook=null; state.resultWorkbook=null; state.missingRows=[]; state.lastResult=null;
    $("templateInput").value=""; $("soInput").value="";
    $("templateInfo").textContent="Belum ada template"; $("templateInfo").className="file-info empty";
    $("soInfo").textContent="Belum ada file SO"; $("soInfo").className="file-info empty";
    ["totalSo","matched","notFound","qtyWritten"].forEach(id => $(id).textContent="0");
    $("downloadBtn").disabled=true; $("downloadMissingBtn").disabled=true;
    $("resultState").className="status-pill neutral"; $("resultState").textContent="Belum diproses";
    progress(0); setStatus("Upload template dan file SO untuk memulai."); renderHistory();
  }

  $("templateInput").addEventListener("change", async e => {
    const f=e.target.files[0]; if(!f) return;
    try { state.templateFile=f; state.templateWorkbook=await readWorkbook(f); state.templateSheetName=firstSheet(state.templateWorkbook); $("templateInfo").textContent=`✓ ${f.name} · ${(f.size/1024).toFixed(1)} KB`; $("templateInfo").className="file-info"; setStatus("Template terbaca. Pilih file SO."); showToast("Template siap."); }
    catch(err){ state.templateFile=null; showToast("Gagal membaca template: "+err.message); setStatus("Gagal membaca template.","error"); }
  });

  $("soInput").addEventListener("change", async e => {
    const files=[...e.target.files]; if(!files.length)return;
    try { state.soFiles=[]; for(const f of files) state.soFiles.push(await readWorkbook(f)); $("soInfo").textContent=`✓ ${files.length} file dipilih · ${files.map(f=>f.name).join(", ")}`; $("soInfo").className="file-info"; setStatus("File SO siap diproses."); showToast(`${files.length} file SO siap.`); }
    catch(err){ showToast("Gagal membaca file SO: "+err.message); setStatus("Gagal membaca file SO.","error"); }
  });

  $("processBtn").addEventListener("click", () => {
    try { processReconcile(); } catch(err) { console.error(err); setStatus(err.message,"error"); showToast("Gagal: "+err.message); }
  });
  $("downloadBtn").addEventListener("click", downloadResult);
  $("downloadMissingBtn").addEventListener("click", downloadMissing);
  $("clearBtn").addEventListener("click", reset);

  $("calcOverstock").addEventListener("click", () => {
    const b=$("ovBrand").value.trim() || "Brand";
    const a=Number($("ovMonth1").value), c=Number($("ovMonth2").value), f=Number($("ovFactor").value), t=Number($("ovTarget").value);
    if (![a,c,f,t].every(Number.isFinite) || a<0 || c<0 || t<0 || f<1.5 || f>2) return showToast("Isi angka dengan benar. Faktor 1,5–2,0.");
    const base=a+c, batas=base*f, over=batas>t, selisih=batas-t;
    $("ovResult").innerHTML=`<strong>${b}</strong><br>Total Bulan 1 + Bulan 2: <b>${base.toLocaleString("id-ID")}</b><br>Nilai × faktor ${f.toFixed(1)}: <b>${batas.toLocaleString("id-ID")}</b><br>Target sekarang: <b>${t.toLocaleString("id-ID")}</b><br><br><b>${over ? "⚠ TERINDIKASI OVERSTOK" : "✓ TIDAK TERINDIKASI OVERSTOK"}</b>${over ? `<br>Selisih di atas target: ${selisih.toLocaleString("id-ID")}` : ""}`;
  });

  function openMenu(){ $("drawer").classList.add("open"); $("drawerBackdrop").classList.remove("hidden"); }
  function closeMenu(){ $("drawer").classList.remove("open"); $("drawerBackdrop").classList.add("hidden"); }
  $("menuBtn").onclick=openMenu; $("closeMenu").onclick=closeMenu; $("drawerBackdrop").onclick=closeMenu;
  document.querySelectorAll(".nav-item").forEach(btn => btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-item").forEach(x=>x.classList.remove("active")); btn.classList.add("active");
    document.querySelectorAll(".section").forEach(s=>s.classList.add("hidden"));
    $("section-"+btn.dataset.section).classList.remove("hidden"); if(btn.dataset.section==="history")renderHistory(); closeMenu(); window.scrollTo({top:0,behavior:"smooth"});
  }));
  renderHistory();
})();
