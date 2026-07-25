(() => {
  "use strict";

  const $ = id => document.getElementById(id);

  const AppState = {
    templateFile: null,
    templateWorkbook: null,
    soFiles: [],
    reconcileResult: null
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
          ReconcileEngine.processMultiple(
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
              cellNF: true,
              cellStyles: true
            }
          );

        setProgress(
          88,
          "Memasukkan QtyFix ke kolom HASIL SO..."
        );

        const result =
          ReconcileEngine.reconcileTemplate(
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

  $("btnDownload").addEventListener(
    "click",
    () => {

      if (
        !AppState.reconcileResult
      ) {

        alert(
          "Belum ada hasil."
        );

        return;

      }

      ExportEngine.downloadXLSX(
        AppState.reconcileResult.workbook
      );

    }
  );

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

  $("btnAnalisa").addEventListener(
    "click",
    () => {

      const result =
        OverstokEngine.calculate(

          $("omsetJuni").value,

          $("omsetJuli").value,

          $("targetAgustus").value,

          $("faktorOverstok").value

        );

      $("totalOmset").textContent =
        rupiah(result.total);

      $("potensiStok").textContent =
        rupiah(result.potensi);

      $("targetResult").textContent =
        rupiah(result.target);

      $("selisihResult").textContent =
        rupiah(
          Math.abs(
            result.selisih
          )
        );

      const box =
        $("statusOverstok");

      box.className =
        `status-result ${
          result.status === "OVERSTOK"
            ? "overstok"
            : "aman"
        }`;

      box.textContent =
        result.status === "OVERSTOK"

          ? `⚠️ OVERSTOK — Potensi ${rupiah(result.potensi)} lebih besar dari Target Agustus.`

          : `✓ AMAN — Potensi ${rupiah(result.potensi)} tidak lebih besar dari Target Agustus.`;

      $("overstokResult")
        .classList
        .remove("hidden");

    }
  );

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
