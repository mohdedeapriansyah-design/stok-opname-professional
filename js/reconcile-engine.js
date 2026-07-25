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
    return String(value).trim().replace(/\s+/g, "");
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

  function processMultiple(soEntries) {
    const qtyMap = new Map();
    const sourceMap = new Map();
    const errors = [];
    let rowsRead = 0;
    let sheetsRead = 0;

    if (!Array.isArray(soEntries)) {
      return {
        qtyMap,
        sourceMap,
        errors: ["Data file SO tidak valid."],
        rowsRead: 0,
        sheetsRead: 0
      };
    }

    for (const entry of soEntries) {
      if (!entry || !entry.workbook) continue;

      const workbook = entry.workbook;

      for (const sheetName of workbook.SheetNames) {
        sheetsRead++;

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;

        const rows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: ""
        });

        const headers = findSOHeaders(rows);

        if (!headers) {
          errors.push(
            `${entry.file.name} / ${sheetName}: kolom productSapCode atau qtyFix tidak ditemukan.`
          );
          continue;
        }

        for (
          let rowIndex = headers.headerRow + 1;
          rowIndex < rows.length;
          rowIndex++
        ) {
          const row = rows[rowIndex] || [];
          const sap = normalizeSap(row[headers.sapCol]);

          if (!sap) continue;

          const qty = normalizeQty(row[headers.qtyCol]);
          const oldQty = qtyMap.get(sap) || 0;

          qtyMap.set(sap, oldQty + qty);

          if (!sourceMap.has(sap)) sourceMap.set(sap, []);

          sourceMap.get(sap).push({
            file: entry.file.name,
            sheet: sheetName,
            row: rowIndex + 1,
            qty
          });

          rowsRead++;
        }
      }
    }

    return {
      qtyMap,
      sourceMap,
      errors,
      rowsRead,
      sheetsRead
    };
  }

  function findTemplateHeaders(rows) {
    const sapHeaders = [
      "kodeitem",
      "kodesap",
      "sapcode",
      "productsapcode",
      "kodeproduk",
      "productcode"
    ];

    const qtyHeaders = [
      "hasilso",
      "qty",
      "quantity",
      "qtyfix",
      "quantityfix"
    ];

    for (let rowIndex = 0; rowIndex < Math.min(rows.length, 50); rowIndex++) {
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

  function reconcileTemplate(templateWorkbook, result) {
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

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: ""
      });

      const headers = findTemplateHeaders(rows);

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
        const row = rows[rowIndex] || [];
        const sap = normalizeSap(row[headers.sapCol]);

        if (!sap) continue;

        totalTemplateRows++;

        if (result.qtyMap.has(sap)) {
          const qty = normalizeQty(result.qtyMap.get(sap));

          const cellAddress = XLSX.utils.encode_cell({
            r: rowIndex,
            c: headers.qtyCol
          });

          worksheet[cellAddress] = {
            t: "n",
            v: qty
          };

          matched.push({
            sheet: sheetName,
            sap,
            qty,
            row: rowIndex + 1
          });

          sheetMatched++;
          totalMatched++;
        } else {
          notFound.push({
            sheet: sheetName,
            sap,
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
        sapColumn: XLSX.utils.encode_col(headers.sapCol),
        qtyColumn: XLSX.utils.encode_col(headers.qtyCol)
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

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: ""
      });

      const headers = findTemplateHeaders(rows);

      result.push({
        sheet: sheetName,
        headerFound: !!headers,
        headerRow: headers ? headers.headerRow + 1 : null,
        sapColumn: headers ? XLSX.utils.encode_col(headers.sapCol) : null,
        sapHeader: headers ? headers.sapHeader : null,
        qtyColumn: headers ? XLSX.utils.encode_col(headers.qtyCol) : null,
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
