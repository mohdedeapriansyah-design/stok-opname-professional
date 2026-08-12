/* ==========================================================
   RECONCILE ENGINE
   STOK OPNAME PROFESSIONAL
   VERSI TEMPLATE KOMPER
========================================================== */

"use strict";

window.ReconcileEngine = (() => {

  function normalizeHeader(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[_\-\/]/g, "");
  }

  function normalizeSap(value) {
    if (value === null || value === undefined || value === "") return "";
    // Excel sering membaca kode SAP numerik sebagai 12345 atau 12345.0.
    // Normalisasi dibuat konsisten agar kode dari SO dan template bisa bertemu.
    if (typeof value === "number" && Number.isFinite(value)) {
      return Number.isInteger(value) ? String(value) : String(value).replace(/\.0+$/, "");
    }
    let text = String(value).trim();
    if (!text) return "";
    text = text.replace(/\s+/g, "");
    text = text.replace(/^'/, "");
    if (/^\d+\.0+$/.test(text)) text = text.replace(/\.0+$/, "");
    return text;
  }

  function getEffectiveRange(worksheet) {
    const keys = Object.keys(worksheet || {});
    let maxR = 0;
    let maxC = 0;
    for (const key of keys) {
      if (!/^[A-Z]+\\d+$/.test(key)) continue;
      const cell = XLSX.utils.decode_cell(key);
      if (cell.r > maxR) maxR = cell.r;
      if (cell.c > maxC) maxC = cell.c;
    }
    if (maxR === 0 && maxC === 0) return null;
    return XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  }

  function normalizeQty(value) {
    if (value === null || value === undefined || value === "") return 0;

    if (typeof value === "number") {
      if (!Number.isFinite(value)) return 0;
      return Number.isInteger(value) ? value : Number(value.toFixed(6));
    }

    let text = String(value).trim().replace(/\s/g, "");
    if (!text) return 0;

    if (text.includes(",") && text.includes(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else if (text.includes(",")) {
      text = text.replace(",", ".");
    }

    const numberValue = Number(text);
    if (!Number.isFinite(numberValue)) return 0;

    return Number.isInteger(numberValue)
      ? numberValue
      : Number(numberValue.toFixed(6));
  }

  function findSOHeaders(rows) {
    const sapHeaders = [
      "productsapcode",
      "sapcode",
      "kodesap",
      "kodeitem",
      "kodeproduk",
      "productcode"
    ];

    const qtyHeaders = [
      "qtyfix",
      "qtyfixed",
      "quantityfix",
      "quantity",
      "qty"
    ];

    for (let rowIndex = 0; rowIndex < Math.min(rows.length, 30); rowIndex++) {
      const row = rows[rowIndex] || [];
      let sapCol = -1;
      let qtyCol = -1;

      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const header = normalizeHeader(row[colIndex]);

        if (sapCol === -1 && sapHeaders.includes(header)) sapCol = colIndex;
        if (qtyCol === -1 && qtyHeaders.includes(header)) qtyCol = colIndex;
      }

      if (sapCol !== -1 && qtyCol !== -1) {
        return {
          headerRow: rowIndex,
          sapCol,
          qtyCol,
          sapHeader: row[sapCol],
          qtyHeader: row[qtyCol]
        };
      }
    }

    return null;
  }

  async function processMultiple(soEntries) {
    const qtyMap = new Map();       // SAP -> total qtyFix
    const odooMap = new Map();      // Odoo -> total qtyFix, derived from same SAP rows
    const sapToOdoo = new Map();    // SAP -> Odoo
    const odooToSap = new Map();    // Odoo -> SAP
    const sourceMap = new Map();
    const errors = [];
    let rowsRead = 0;
    let sheetsRead = 0;

    if (!Array.isArray(soEntries)) {
      return { qtyMap, odooMap, sapToOdoo, sourceMap, errors: ["Data file SO tidak valid."], rowsRead: 0, sheetsRead: 0 };
    }

    for (const entry of soEntries) {
      if (!entry || !entry.workbook) continue;

      for (const sheetName of entry.workbook.SheetNames) {
        sheetsRead++;
        const worksheet = entry.workbook.Sheets[sheetName];
        if (!worksheet) continue;

        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        const headers = findSOHeaders(rows);

        if (!headers) {
          errors.push(`${entry.file.name} / ${sheetName}: kolom productSapCode atau qtyFix tidak ditemukan.`);
          continue;
        }

        // productOdooCode diperlukan sebagai jembatan untuk Template Komper lama
        // yang menyimpan Kode Odoo di kolom M, sementara Kode SAP tetap menjadi
        // identitas utama dari file SO.
        const headerRow = rows[headers.headerRow] || [];
        let odooCol = -1;
        for (let c = 0; c < headerRow.length; c++) {
          if (normalizeHeader(headerRow[c]) === "productodoocode") {
            odooCol = c;
            break;
          }
        }

        for (let rowIndex = headers.headerRow + 1; rowIndex < rows.length; rowIndex++) {
          if ((rowIndex - headers.headerRow) % 2000 === 0) {
            await new Promise(requestAnimationFrame);
          }

          const row = rows[rowIndex] || [];
          const sap = normalizeSap(row[headers.sapCol]);
          if (!sap) continue;

          const qty = normalizeQty(row[headers.qtyCol]);
          qtyMap.set(sap, (qtyMap.get(sap) || 0) + qty);

          const odoo = odooCol >= 0 ? normalizeSap(row[odooCol]) : "";
          if (odoo) {
            odooMap.set(odoo, (odooMap.get(odoo) || 0) + qty);
            sapToOdoo.set(sap, odoo);
            odooToSap.set(odoo, sap);
          }

          if (!sourceMap.has(sap)) sourceMap.set(sap, []);
          sourceMap.get(sap).push({
            file: entry.file.name,
            sheet: sheetName,
            row: rowIndex + 1,
            qty,
            odoo
          });

          rowsRead++;
        }
      }
    }

    return { qtyMap, odooMap, sapToOdoo, odooToSap, sourceMap, errors, rowsRead, sheetsRead };
  }

  function findTemplateHeaders(rows) {
    // SAP adalah kunci utama. Jangan menganggap "KODE ITEM" sebagai Kode SAP.
    const sapHeaders = [
      "kodesap", "sapcode", "productsapcode", "productsapkode",
      "sapkode", "kodeproduk", "productcode"
    ];
    const odooHeaders = ["productodoocode", "odoocode", "kodeodoo", "odoo"];
    const qtyHeaders = [
      "hasilso", "hasilstok", "hasilopname", "hasilqty",
      "qtyhasil", "qtyfix", "quantityfix"
    ];

    let explicit = null;
    let bridge = null;

    for (let rowIndex = 0; rowIndex < Math.min(rows.length, 80); rowIndex++) {
      const row = rows[rowIndex] || [];
      let sapCol = -1, odooCol = -1, qtyCol = -1;

      for (let c = 0; c < row.length; c++) {
        const h = normalizeHeader(row[c]);
        if (sapCol < 0 && sapHeaders.includes(h)) sapCol = c;
        if (odooCol < 0 && odooHeaders.includes(h)) odooCol = c;
        if (qtyCol < 0 && qtyHeaders.includes(h)) qtyCol = c;
      }

      if (sapCol >= 0 && qtyCol >= 0) {
        explicit = {
          headerRow: rowIndex, sapCol, qtyCol,
          sapHeader: row[sapCol], qtyHeader: row[qtyCol],
          matchType: "SAP"
        };
        break;
      }

      if (odooCol >= 0 && qtyCol >= 0 && !bridge) {
        bridge = {
          headerRow: rowIndex, odooCol, qtyCol,
          odooHeader: row[odooCol], qtyHeader: row[qtyCol],
          matchType: "SAP_VIA_ODOO"
        };
      }
    }

    return explicit || bridge || null;
  }

  function findBestTemplateHeaders(rows) {
    return findTemplateHeaders(rows);
  }

  async function reconcileTemplate(templateWorkbook, result) {
    if (!templateWorkbook) {
      throw new Error("Workbook template tidak tersedia.");
    }

    if (!result || !result.qtyMap) {
      throw new Error("Data hasil SO tidak tersedia.");
    }

    const workbook = templateWorkbook;
    const notFound = [];
    const matched = [];
    const sheetResults = [];

    let totalTemplateRows = 0;
    let totalMatched = 0;
    let totalNotFound = 0;

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) continue;

      const effectiveRange = getEffectiveRange(worksheet);
      if (!effectiveRange) continue;
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        range: effectiveRange
      });

      const headers = findBestTemplateHeaders(rows);

      if (!headers) {
        sheetResults.push({
          sheet: sheetName,
          status: "DILEWATI",
          reason: "KODE ITEM / HASIL SO tidak ditemukan.",
          matched: 0,
          notFound: 0
        });
        continue;
      }

      let sheetMatched = 0;
      let sheetNotFound = 0;

      for (
        let rowIndex = headers.headerRow + 1;
        rowIndex < rows.length;
        rowIndex++
      ) {
        if ((rowIndex - headers.headerRow) % 2000 === 0) {
          await new Promise(requestAnimationFrame);
        }
        const row = rows[rowIndex] || [];
        // Jika template punya SAP, gunakan langsung. Jika template Komper lama,
        // gunakan productOdooCode sebagai bridge dari SAP yang sudah diproses dari file SO.
        const templateKey = headers.matchType === "SAP_VIA_ODOO"
          ? normalizeSap(row[headers.odooCol])
          : normalizeSap(row[headers.sapCol]);

        if (!templateKey) continue;

        totalTemplateRows++;

        let matchedKey = templateKey;
        let sap = headers.matchType === "SAP_VIA_ODOO"
          ? (result.odooToSap && result.odooToSap.get(templateKey)) || templateKey
          : templateKey;

        let qtyValue = result.qtyMap.get(sap);
        let matchMode = "SAP";

        // Template Komper yang dikirim tidak memiliki kolom KODE SAP.
        // Ia menyimpan productOdooCode di kolom M. Karena SAP adalah kunci utama,
        // gunakan relasi productSapCode -> productOdooCode dari file SO sebagai bridge.
        if (headers.matchType === "SAP_VIA_ODOO") {
          matchedKey = normalizeSap(row[headers.odooCol]);
          qtyValue = result.odooMap ? result.odooMap.get(matchedKey) : undefined;
          matchMode = "SAP_VIA_ODOO";
        }

        if (qtyValue !== undefined) {
          const qty = normalizeQty(qtyValue);

          const cellAddress = XLSX.utils.encode_cell({
            r: rowIndex,
            c: headers.qtyCol
          });

          worksheet[cellAddress] = {
            t: "n",
            v: qty
          };
          // Pastikan range worksheet mencakup sel HASIL SO yang baru diisi.
          const cellRef = XLSX.utils.decode_range(worksheet["!ref"] || cellAddress);
          const target = XLSX.utils.decode_cell(cellAddress);
          cellRef.s.r = Math.min(cellRef.s.r, target.r);
          cellRef.s.c = Math.min(cellRef.s.c, target.c);
          cellRef.e.r = Math.max(cellRef.e.r, target.r);
          cellRef.e.c = Math.max(cellRef.e.c, target.c);
          worksheet["!ref"] = XLSX.utils.encode_range(cellRef);

          matched.push({
            sheet: sheetName,
            sap,
            qty,
            row: rowIndex + 1,
            matchMode
          });

          sheetMatched++;
          totalMatched++;
        } else {
          notFound.push({
            sheet: sheetName,
            sap: headers.matchType === "SAP_VIA_ODOO"
              ? ((result.odooToSap && result.odooToSap.get(templateKey)) || "")
              : sap,
            templateKey,
            row: rowIndex + 1
          });

          sheetNotFound++;
          totalNotFound++;
        }
      }

      sheetResults.push({
        sheet: sheetName,
        status: "DIPROSES",
        matched: sheetMatched,
        notFound: sheetNotFound,
        headerRow: headers.headerRow + 1,
        sapColumn: headers.sapCol >= 0 ? XLSX.utils.encode_col(headers.sapCol) : null,
        odooColumn: headers.odooCol >= 0 ? XLSX.utils.encode_col(headers.odooCol) : null,
        qtyColumn: XLSX.utils.encode_col(headers.qtyCol),
        matchType: headers.matchType
      });
    }

    return {
      workbook,
      matched,
      notFound,
      templateRows: totalTemplateRows,
      totalMatched,
      totalNotFound,
      sheetResults
    };
  }

  function inspectTemplate(templateWorkbook) {
    const result = [];

    for (const sheetName of templateWorkbook.SheetNames) {
      const worksheet = templateWorkbook.Sheets[sheetName];

      const effectiveRange = getEffectiveRange(worksheet);
      if (!effectiveRange) continue;
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        range: effectiveRange
      });

      const headers = findTemplateHeaders(rows);

      result.push({
        sheet: sheetName,
        headerFound: !!headers,
        headerRow: headers ? headers.headerRow + 1 : null,
        sapColumn: headers && headers.sapCol >= 0 ? XLSX.utils.encode_col(headers.sapCol) : null,
        sapHeader: headers ? (headers.sapHeader || headers.odooHeader) : null,
        odooColumn: headers && headers.odooCol >= 0 ? XLSX.utils.encode_col(headers.odooCol) : null,
        qtyColumn: headers ? XLSX.utils.encode_col(headers.qtyCol) : null,
        matchType: headers ? headers.matchType : null,
        qtyHeader: headers ? headers.qtyHeader : null
      });
    }

    return result;
  }

  return {
    processMultiple,
    reconcileTemplate,
    findTemplateHeaders,
    findSOHeaders,
    inspectTemplate,
    normalizeQty,
    normalizeSap
  };

})();
