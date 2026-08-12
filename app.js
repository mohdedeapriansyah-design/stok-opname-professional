(() => {
  "use strict";

  const $ = id => document.getElementById(id);

  const AppState = {
    templateFile: null,
    templateWorkbook: null,
    soFiles: [],
    reconcileResult: null,
    overstokBrands: []
  };

  const menuToggle = $("menuToggle");
  const sidebar = $("sidebar");
  const overlay = $("overlay");

  function openMenu() {
    sidebar.classList.add("open");
    menuToggle.setAttribute("aria-expanded", "true");
  }

  function closeMenu() {
    sidebar.classList.remove("open");
    menuToggle.setAttribute("aria-expanded", "false");
  }

  menuToggle.addEventListener("click", () =>
    sidebar.classList.contains("open") ? closeMenu() : openMenu()
  );

  $("menuClose").addEventListener("click", closeMenu);
  overlay.addEventListener("click", closeMenu);

  function showPage(page) {
    document.querySelectorAll(".page").forEach(x =>
      x.classList.remove("active")
    );

    $(`page-${page}`).classList.add("active");

    document.querySelectorAll(".nav-item").forEach(x =>
      x.classList.toggle("active", x.dataset.page === page)
    );

    closeMenu();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  document.querySelectorAll(".nav-item").forEach(x =>
    x.addEventListener("click", () => showPage(x.dataset.page))
  );

  document.querySelectorAll("[data-go]").forEach(x =>
    x.addEventListener("click", () => showPage(x.dataset.go))
  );

  $("templateFile").addEventListener("change", async e => {
    const file = e.target.files[0];
    if (!file) return;

    AppState.templateFile = file;

    $("templateInfo").className = "file-info";
    $("templateInfo").textContent =
      `✓ ${file.name} — ${(file.size / 1024).toFixed(1)} KB`;

    try {
      AppState.templateWorkbook =
        await FileEngine.readWorkbook(file);

      const inspection =
        ReconcileEngine.inspectTemplate(
          AppState.templateWorkbook
        );

      const validSheets =
        inspection.filter(x => x.headerFound);

      if (validSheets.length === 0) {
        $("statusText").textContent =
          "Template terbaca, tetapi struktur KODE ITEM / HASIL SO belum ditemukan.";
      } else {
        $("statusText").textContent =
          `Template siap. ${validSheets.length} sheet terdeteksi.`;
      }

    } catch (error) {

      $("statusText").textContent =
        "Gagal membaca template: " + error.message;

    }
  });

  $("soFiles").addEventListener("change", e => {

    AppState.soFiles =
      Array.from(e.target.files || []);

    const box = $("soFileList");

    box.className = "file-list";

    box.innerHTML =
      AppState.soFiles.length

        ? AppState.soFiles.map((file, index) => `
            <div class="file-row">
              <span>
                ${index + 1}. ${escapeHtml(file.name)}
              </span>
              <small>
                ${(file.size / 1024).toFixed(1)} KB
              </small>
            </div>
          `).join("")

        : "Belum ada file SO";

  });

  function setProgress(number, text) {

    $("progressBar").style.width =
      `${number}%`;

    $("statusText").textContent =
      text;

  }

  $("btnProcess").addEventListener(
    "click",
    async () => {

      try {

        if (
          !AppState.templateFile ||
          !AppState.templateWorkbook
        ) {

          throw new Error(
            "Upload Template Komper terlebih dahulu."
          );

        }

        if (
          !AppState.soFiles.length
        ) {

          throw new Error(
            "Upload minimal satu file hasil SO."
          );

        }

        $("btnProcess").disabled = true;

        setProgress(
          5,
          "Membaca file SO..."
        );

        const soEntries =
          await FileEngine.readMultiple(
            AppState.soFiles,
            setProgress
          );

        setProgress(
          70,
          "Menggabungkan data QtyFix..."
        );

        const processed =
          await ReconcileEngine.processMultiple(
            soEntries
          );

        if (
          processed.qtyMap.size === 0
        ) {

          throw new Error(
            "Tidak ada data QtyFix yang berhasil dibaca dari file SO."
          );

        }

        setProgress(
          82,
          "Membuat salinan workbook template..."
        );

        const workbook =
          XLSX.read(
            await AppState.templateFile.arrayBuffer(),
            {
              type: "array",
              cellFormula: true,
              cellNF: false,
              cellStyles: false,
              cellHTML: false
            }
          );

        setProgress(
          88,
          "Memasukkan QtyFix ke kolom HASIL SO..."
        );

        const result =
          await ReconcileEngine.reconcileTemplate(
            workbook,
            processed
          );

        AppState.reconcileResult = {
          ...result,
          processed
        };

        $("statTemplate").textContent =
          result.templateRows;

        $("statSO").textContent =
          processed.qtyMap.size;

        $("statMatched").textContent =
          result.matched.length;

        $("statNotFound").textContent =
          result.notFound.length;

        $("resultSummary").textContent =
          `${processed.rowsRead} baris SO dibaca dari ` +
          `${AppState.soFiles.length} file dan ` +
          `${processed.sheetsRead} sheet.`;

        $("resultPanel").classList.remove(
          "hidden"
        );

        const nf =
          $("notFoundBox");

        if (
          result.notFound.length
        ) {

          nf.classList.remove(
            "hidden"
          );

          nf.innerHTML =
            `<b>${result.notFound.length} data KODE ITEM tidak ditemukan di SO.</b>`;

        } else {

          nf.classList.add(
            "hidden"
          );

        }

        StorageEngine.add({

          date:
            new Date().toLocaleString(
              "id-ID"
            ),

          template:
            AppState.templateFile.name,

          files:
            AppState.soFiles.map(
              x => x.name
            ),

          matched:
            result.matched.length,

          notFound:
            result.notFound.length

        });

        renderHistory();

        setProgress(
          100,
          "✓ Proses selesai. Hasil siap di-download."
        );

      } catch (error) {

        console.error(
          "RECONCILE ERROR:",
          error
        );

        alert(
          error.message ||
          String(error)
        );

        setProgress(
          0,
          "Proses gagal."
        );

      } finally {

        $("btnProcess").disabled =
          false;

      }

    }
  );

  $("btnDownload").addEventListener("click", async () => {
    if (!AppState.reconcileResult) {
      alert("Belum ada hasil.");
      return;
    }

    const btn = $("btnDownload");
    if (btn.disabled) return;
    btn.disabled = true;
    const oldText = btn.textContent;
    btn.textContent = "⏳ Menyiapkan file...";

    try {
      await new Promise(requestAnimationFrame);
      await ExportEngine.downloadXLSX(AppState.reconcileResult.workbook);
      $("statusText").textContent = "✓ Download dimulai.";
    } catch (error) {
      console.error("DOWNLOAD ERROR:", error);
      alert("Gagal membuat file download: " + (error.message || error));
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  });

  $("btnPreview").addEventListener(
    "click",
    () => {

      const result =
        AppState.reconcileResult;

      if (!result) return;

      const rows =
        result.matched.slice(
          0,
          500
        );

      $("previewContent").innerHTML = `

        <div class="stats-grid">

          <div class="stat">
            <span>Sheet Diproses</span>
            <b>
              ${
                result.sheetResults
                  .filter(x =>
                    x.status === "DIPROSES"
                  ).length
              }
            </b>
          </div>

          <div class="stat">
            <span>SKU Cocok</span>
            <b>
              ${result.matched.length}
            </b>
          </div>

          <div class="stat">
            <span>Tidak Ditemukan</span>
            <b>
              ${result.notFound.length}
            </b>
          </div>

        </div>

        <table class="data-table">

          <thead>

            <tr>
              <th>Sheet</th>
              <th>Kode Item</th>
              <th>QtyFix</th>
              <th>Baris Template</th>
            </tr>

          </thead>

          <tbody>

            ${
              rows.map(x => `

                <tr>

                  <td>
                    ${escapeHtml(x.sheet)}
                  </td>

                  <td>
                    ${escapeHtml(x.sap)}
                  </td>

                  <td>
                    ${x.qty}
                  </td>

                  <td>
                    ${x.row}
                  </td>

                </tr>

              `).join("")
            }

          </tbody>

        </table>

        <button
          id="exportNF"
          class="secondary-btn"
          style="margin-top:12px"
        >
          Export SKU Tidak Ditemukan
        </button>

      `;

      $("previewModal")
        .classList
        .remove("hidden");

      $("exportNF").onclick =
        () =>
          ExportEngine.exportNotFound(
            result.notFound
          );

    }
  );

  $("btnClose").addEventListener(
    "click",
    () =>
      $("previewModal")
        .classList
        .add("hidden")
  );

  $("previewModal").addEventListener(
    "click",
    e => {

      if (
        e.target ===
        $("previewModal")
      ) {

        $("previewModal")
          .classList
          .add("hidden");

      }

    }
  );

  // =========================
  // ANALISA OVERSTOK PER BRAND
  // Rumus: ((stok bulan 1 + stok bulan 2) * 2) / omset bulan 3
  // =========================
  function getOverstokInput() {
    const brand = $("brandOverstok").value.trim();
    const stok1 = Number($("stokBulan1").value);
    const stok2 = Number($("stokBulan2").value);
    const omset3 = Number($("omsetBulan3").value);

    if (!brand) throw new Error("Nama brand wajib diisi.");
    if (![stok1, stok2, omset3].every(Number.isFinite) || stok1 < 0 || stok2 < 0 || omset3 <= 0) {
      throw new Error("Nominal stok harus >= 0 dan Omset bulan 3 harus lebih dari 0.");
    }
    return { brand, stok1, stok2, omset3 };
  }

  function clearOverstokInputs() {
    $("brandOverstok").value = "";
    $("stokBulan1").value = "";
    $("stokBulan2").value = "";
    $("omsetBulan3").value = "";
    $("brandOverstok").focus();
  }

  function renderOverstokBrands() {
    const box = $("overstokTableBody");
    if (!box) return;
    const target = Number($("targetRasio").value) || 100;

    box.innerHTML = AppState.overstokBrands.length
      ? AppState.overstokBrands.map((x, i) => {
          const r = OverstokEngine.calculateBrand(x.stok1, x.stok2, x.omset3, target);
          return `<tr>
            <td><b>${escapeHtml(x.brand)}</b></td>
            <td>${rupiah(x.stok1)}</td>
            <td>${rupiah(x.stok2)}</td>
            <td>${rupiah(x.omset3)}</td>
            <td>${rupiah(x.stok1 + x.stok2)} × 2 ÷ ${rupiah(x.omset3)}</td>
            <td>${r.ratio.toFixed(2)}%</td>
            <td><span class="status-result ${r.status === "OVERSTOK" ? "overstok" : "aman"}">${r.status}</span>
              <button type="button" class="danger-btn" data-remove-brand="${i}" style="margin-left:6px;padding:5px 8px">×</button>
            </td>
          </tr>`;
        }).join("")
      : `<tr><td colspan="7" class="muted">Belum ada data brand.</td></tr>`;
  }

  $("btnTambahBrand").addEventListener("click", () => {
    try {
      const item = getOverstokInput();
      AppState.overstokBrands.push(item);
      renderOverstokBrands();
      $("overstokResult").classList.remove("hidden");
      clearOverstokInputs();
      $("statusText").textContent = `Data ${item.brand} ditambahkan.`;
    } catch (error) {
      alert(error.message);
    }
  });

  $("btnClearBrand").addEventListener("click", () => {
    AppState.overstokBrands = [];
    renderOverstokBrands();
    $("overstokResult").classList.add("hidden");
  });

  $("overstokTableBody").addEventListener("click", e => {
    const btn = e.target.closest("[data-remove-brand]");
    if (!btn) return;
    AppState.overstokBrands.splice(Number(btn.dataset.removeBrand), 1);
    renderOverstokBrands();
  });

  $("btnAnalisa").addEventListener("click", () => {
    try {
      if (!AppState.overstokBrands.length) {
        // Izinkan satu input langsung tanpa harus menekan Tambah Brand.
        const item = getOverstokInput();
        AppState.overstokBrands.push(item);
        clearOverstokInputs();
      }

      const target = Number($("targetRasio").value);
      if (!Number.isFinite(target) || target < 0) throw new Error("Target rasio tidak valid.");

      const results = AppState.overstokBrands.map(x =>
        OverstokEngine.calculateBrand(x.stok1, x.stok2, x.omset3, target)
      );
      const over = results.filter(x => x.status === "OVERSTOK").length;

      $("totalBrand").textContent = results.length;
      $("jumlahOverstok").textContent = over;
      $("jumlahAman").textContent = results.length - over;
      $("targetResult").textContent = `${target.toFixed(2)}%`;

      renderOverstokBrands();
      $("overstokResult").classList.remove("hidden");
      $("statusText").textContent = `Analisa selesai: ${over} dari ${results.length} brand terindikasi OVERSTOK.`;
    } catch (error) {
      alert(error.message);
    }
  });

  $("btnFilterQty").addEventListener(
    "click",
    () => {

      const data =
        OverstokEngine
          .parseBulk(
            $("bulkData").value
          )
          .filter(
            x => x.qty > 6
          );

      $("qtyAnalysisResult").innerHTML =

        data.length

          ? `

            <table class="data-table">

              <thead>

                <tr>
                  <th>Kode SAP</th>
                  <th>Nama</th>
                  <th>Qty</th>
                </tr>

              </thead>

              <tbody>

                ${
                  data.map(x => `

                    <tr>

                      <td>
                        ${escapeHtml(x.sap)}
                      </td>

                      <td>
                        ${escapeHtml(x.name)}
                      </td>

                      <td>
                        ${x.qty}
                      </td>

                    </tr>

                  `).join("")
                }

              </tbody>

            </table>

          `

          : "<p class='muted'>Tidak ada data dengan Qty lebih dari 6.</p>";

    }
  );

  function renderHistory() {

    const data =
      StorageEngine.get();

    const box =
      $("historyList");

    box.innerHTML =

      data.length

        ? data.map(x => `

            <div class="history-item">

              <strong>
                ${escapeHtml(x.template)}
              </strong>

              <span>
                ${escapeHtml(x.date)}
                · ${x.files.length} file SO
                · ${x.matched} SKU cocok
                · ${x.notFound} tidak ditemukan
              </span>

            </div>

          `).join("")

        : `
          <div class="panel">
            <p class="muted">
              Belum ada riwayat proses.
            </p>
          </div>
        `;

  }

  $("btnClearHistory").addEventListener(
    "click",
    () => {

      if (
        confirm(
          "Hapus semua riwayat?"
        )
      ) {

        StorageEngine.clear();

        renderHistory();

      }

    }
  );

  renderHistory();

  function rupiah(number) {

    return new Intl.NumberFormat(
      "id-ID",
      {
        style: "currency",
        currency: "IDR",
        maximumFractionDigits: 0
      }
    ).format(
      Number(number) || 0
    );

  }

  function escapeHtml(value) {

    return String(
      value ?? ""
    ).replace(
      /[&<>"']/g,
      match =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&#039;"
        }[match])
    );

  }

})();
