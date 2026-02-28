// --- DETECTOR GLOBAL DE ERRORES --- //
window.onerror = function (msg, url, lineNo, columnNo, error) {
  alert("Error interno detectado:\n" + msg + "\nLínea: " + lineNo);
  return false;
};

class DataViewerApp {
  constructor() {
    this.rawData = [];
    this.visibleData = [];
    this.columns = [];
    this.colSettings = {};

    // Pagination params
    this.pageSize = 100;
    this.currentPage = 1;

    // Sort/Filter state
    this.sortCol = null;
    this.sortAsc = true;
    this.searchQuery = "";
    this.activeMenuCol = null;

    // Excel Workbook State
    this.currentWorkbook = null;
    this.tempRawMatrix = [];
    this.tempHeaderIdx = 0;

    // Historial Ctrl+Z
    this.undoStack = [];

    this.initElements();
    this.initEvents();
  }

  initElements() {
    this.filterSummary = document.getElementById("filterSummaryBar");
    this.els = {
      fileInput: document.getElementById("fileInput"),
      emptyState: document.getElementById("emptyState"),
      loadingState: document.getElementById("loadingState"),
      tableWrapper: document.getElementById("tableWrapper"),
      footer: document.getElementById("appFooter"),
      tbody: document.getElementById("tBody"),
      thead: document.getElementById("tHead"),
      tfoot: document.getElementById("tFoot"),
      colMenu: document.getElementById("colMenu"),
      exportMenu: document.getElementById("exportMenu"),
      globalSearch: document.getElementById("globalSearch"),
      dragOverlay: document.getElementById("dragOverlay"),
      ctxMenu: document.getElementById("columnContextMenu"),
      colListContainer: document.getElementById("colListContainer"),
      reportTitle: document.getElementById("reportTitle"),
      reportAuthor: document.getElementById("reportAuthor"),
      sheetModal: document.getElementById("sheetModal"),
      sheetList: document.getElementById("sheetList"),
      exportModal: document.getElementById("exportModal"),
      confirmTitle: document.getElementById("confirmTitle"),
      confirmAuthor: document.getElementById("confirmAuthor"),
      btnConfirmExport: document.getElementById("btnConfirmExport"),
      btnCancelExport: document.getElementById("btnCancelExport"),

      structureModal: document.getElementById("structureModal"),
      previewTable: document.getElementById("previewTable"),
      footerSkipCount: document.getElementById("footerSkipCount"),
      selectedHeaderDisplay: document.getElementById(
        "selectedHeaderIndexDisplay"
      ),
      btnConfirmStructure: document.getElementById("btnConfirmStructure"),
      btnCancelStructure: document.getElementById("btnCancelStructure"),
      csvMapModal: document.getElementById("csvMapModal"),
      mapLocalidad: document.getElementById("mapLocalidad"),
      mapScanCode: document.getElementById("mapScanCode"),
      mapProducto: document.getElementById("mapProducto"),
      mapPedido: document.getElementById("mapPedido"),
      mapOrdenCompra: document.getElementById("mapOrdenCompra"),
      chkAutoOC: document.getElementById("chkAutoOC"),
      previewAutoOC: document.getElementById("previewAutoOC"),
      chkManualLocalidad: document.getElementById("chkManualLocalidad"),
      inputManualLocalidad: document.getElementById("inputManualLocalidad")
    };
  }

  initEvents() {
    // BLINDAJE DEL INPUT DE ARCHIVOS
    if (this.els.fileInput) {
      this.els.fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files.length > 0) {
          this.handleFiles(e.target.files);
        }
      });
      this.els.fileInput.addEventListener("click", (e) => {
        e.target.value = null;
      });
    }

    const btnReset = document.getElementById("btnResetPrefs");
    if (btnReset)
      btnReset.addEventListener("click", () => this.resetPreferences());

    // Drag & Drop
    window.addEventListener("dragover", (e) => {
      e.preventDefault();
      this.els.dragOverlay.classList.add("active");
    });
    window.addEventListener("dragleave", (e) => {
      if (e.target === this.els.dragOverlay)
        this.els.dragOverlay.classList.remove("active");
    });
    window.addEventListener("drop", (e) => {
      e.preventDefault();
      this.els.dragOverlay.classList.remove("active");
      if (e.dataTransfer.files.length) this.handleFiles(e.dataTransfer.files);
    });

    // Pagination
    document
      .getElementById("btnPrev")
      .addEventListener("click", () => this.changePage(-1));
    document
      .getElementById("btnNext")
      .addEventListener("click", () => this.changePage(1));
    document.getElementById("pageSize").addEventListener("change", (e) => {
      this.pageSize = parseInt(e.target.value);
      this.currentPage = 1;
      this.savePreferences();
      this.render();
    });

    // Search
    this.els.globalSearch.addEventListener("input", (e) => {
      this.searchQuery = e.target.value.toLowerCase();
      this.currentPage = 1;
      this.processData();
    });

    // Modals
    document
      .getElementById("btnCloseSheetModal")
      .addEventListener("click", () => {
        this.els.sheetModal.classList.remove("active");
        this.setLoading(false);
        this.resetState();
      });

    this.els.btnCancelExport.addEventListener("click", () => {
      this.els.exportModal.classList.remove("active");
      this.pendingExportFormat = null;
    });

    this.els.btnConfirmExport.addEventListener("click", () => {
      this.els.reportTitle.value = this.els.confirmTitle.value;
      this.els.reportAuthor.value = this.els.confirmAuthor.value;
      this.executeExport(this.pendingExportFormat);
      this.els.exportModal.classList.remove("active");
      this.pendingExportFormat = null;
    });

    this.els.btnCancelStructure.addEventListener("click", () => {
      this.els.structureModal.classList.remove("active");
      this.resetState();
    });

    this.els.btnConfirmStructure.addEventListener("click", () => {
      this.applyStructureAndLoad();
    });

    this.els.footerSkipCount.addEventListener("input", () => {
      this.renderPreviewTableRows();
    });

    // Click Outside
    document.addEventListener("click", (e) => {
      if (!e.target.closest("#btnColumns") && !e.target.closest("#colMenu"))
        this.els.colMenu.classList.remove("show");
      if (!e.target.closest("#btnExport") && !e.target.closest("#exportMenu"))
        this.els.exportMenu.classList.remove("show");
      if (
        !e.target.closest("#columnContextMenu") &&
        !e.target.closest(".btn-col-menu")
      )
        this.els.ctxMenu.classList.remove("show");
    });

    // Ctrl + Z
    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
        if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA")
          return;
        e.preventDefault();
        this.undo();
      }
    });
  }

  // --- CORE FILE HANDLING --- //
  async handleFiles(fileList) {
    try {
      if (!fileList || fileList.length === 0) return;
      const files = Array.from(fileList);

      this.resetState();

      if (files.length === 1) {
        await this.handleSingleFile(files[0]);
      } else {
        await this.handleMultipleFiles(files);
      }
    } catch (err) {
      alert("Error en handleFiles: " + err.message);
    }
  }

 async handleMultipleFiles(files) {
    this.setLoading(true);
    this.showToast(`Iniciando aceleración por hardware para ${files.length} archivos...`, 'info');

    try {
      // Instanciamos el Web Worker
      const worker = new Worker('worker.js');

      // Escuchamos lo que nos dice el Worker
      worker.onmessage = (e) => {
        const response = e.data;

        if (response.type === 'progress') {
          // Actualizar texto de carga
          const loadingText = this.els.loadingState.querySelector("p");
          if(loadingText) loadingText.innerText = response.msg;
        } 
        else if (response.type === 'error') {
          this.showToast(response.msg, 'error');
        } 
        else if (response.type === 'done') {
          if (response.data.length === 0) {
            this.showToast("No se encontraron datos válidos.", "error");
            this.setLoading(false);
            worker.terminate();
            return;
          }

          this.els.reportTitle.value = "Reporte_Combinado";
          
          const loadingText = this.els.loadingState.querySelector("p");
          if(loadingText) loadingText.innerText = "Construyendo tabla interactiva...";

          // Inicializamos los datos
          this.initData(response.data, response.columns);
          this.setLoading(false);

          if (response.structureMismatch) {
              this.showToast(`Carga masiva: ${response.data.length.toLocaleString()} filas. NOTA: Las columnas variaban.`, 'warning');
          } else {
              this.showToast(`¡Completado en tiempo récord! ${response.data.length.toLocaleString()} filas procesadas.`, 'success');
          }

          // Asesinamos al Worker para liberar la RAM instantáneamente
          worker.terminate();
        }
      };

      worker.onerror = (err) => {
        console.error("Error en Worker:", err);
        this.showToast("Ocurrió un error en el procesamiento en segundo plano.", "error");
        this.setLoading(false);
        worker.terminate();
      };

      // Le pasamos los archivos al Worker y que comience la magia
      worker.postMessage({ files: Array.from(files) });

    } catch (err) {
      console.error(err);
      this.showToast(`Error de configuración: ${err.message}`, 'error');
      this.setLoading(false);
    }
  }

  async handleSingleFile(file) {
    if (!file) return;
    this.setLoading(true);
    const fname = file.name.replace(/\.[^/.]+$/, "");
    this.els.reportTitle.value = fname;

    try {
      await this.forceRender();
      const data = await this.readFileAsync(file);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      this.currentWorkbook = workbook;

      if (workbook.Props && workbook.Props.Author) {
        this.els.reportAuthor.value = workbook.Props.Author;
      } else {
        this.els.reportAuthor.value = "Usuario";
      }

      if (workbook.SheetNames.length === 0)
        throw new Error("Archivo vacío o inválido.");

      if (workbook.SheetNames.length > 1) {
        this.showSheetSelection(workbook.SheetNames);
        this.setLoading(false);
      } else {
        await this.loadSheetData(workbook.SheetNames[0]);
      }
    } catch (err) {
      console.error(err);
      this.showToast(`Error de archivo: ${err.message}`, "error");
      this.setLoading(false);
    }
  }

  readFileAsync(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(new Uint8Array(e.target.result));
      reader.onerror = (e) => reject(new Error("Error de lectura de archivo"));
      reader.readAsArrayBuffer(file);
    });
  }

  forceRender() {
    return new Promise((resolve) =>
      requestAnimationFrame(() => setTimeout(resolve, 0))
    );
  }

  resetState() {
    this.rawData = [];
    this.visibleData = [];
    this.columns = [];
    this.currentWorkbook = null;
    this.sortCol = null;
    this.searchQuery = "";
    this.undoStack = [];

    if (this.els.globalSearch) this.els.globalSearch.value = "";
    if (this.els.tableWrapper) this.els.tableWrapper.classList.add("hidden");
    if (this.els.footer) this.els.footer.classList.add("hidden");
    if (this.filterSummary) this.filterSummary.classList.add("hidden");
    if (this.els.emptyState) this.els.emptyState.classList.remove("hidden");
    if (this.els.thead) this.els.thead.innerHTML = "";
    if (this.els.tbody) this.els.tbody.innerHTML = "";
    if (this.els.tfoot) this.els.tfoot.innerHTML = "";

    this.tempRawMatrix = [];
  }

  showSheetSelection(sheets) {
    const list = this.els.sheetList;
    list.innerHTML = "";
    sheets.forEach((sheet) => {
      const btn = document.createElement("div");
      btn.className = "sheet-btn";
      btn.innerHTML = `<span style="font-weight:600">${sheet}</span> <i class="ph ph-caret-right"></i>`;
      btn.onclick = async () => {
        this.els.sheetModal.classList.remove("active");
        this.setLoading(true);
        await this.forceRender();
        await this.loadSheetData(sheet);
      };
      list.appendChild(btn);
    });
    this.els.sheetModal.classList.add("active");
  }

  async loadSheetData(sheetName) {
    try {
      if (!this.currentWorkbook) throw new Error("No hay libro cargado.");
      const sheet = this.currentWorkbook.Sheets[sheetName];
      const rawMatrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      if (rawMatrix.length === 0)
        throw new Error(`La hoja "${sheetName}" está vacía.`);

      this.tempRawMatrix = rawMatrix;
      this.tempHeaderIdx = 0;
      this.els.footerSkipCount.value = 0;
      this.openStructureSelector();
    } catch (err) {
      this.showToast(err.message, "error");
      this.setLoading(false);
    }
  }

  openStructureSelector() {
    this.setLoading(false);
    this.els.sheetModal.classList.remove("active");
    this.els.structureModal.classList.add("active");

    const likelyHeader = this.tempRawMatrix.findIndex(
      (row) => row && row.filter((c) => c).length > 1
    );
    this.tempHeaderIdx = likelyHeader >= 0 ? likelyHeader : 0;
    this.renderPreviewTableRows();
  }

  renderPreviewTableRows() {
    const table = this.els.previewTable;
    table.innerHTML = "";
    const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;
    const totalRows = this.tempRawMatrix.length;

    this.els.selectedHeaderDisplay.innerText = `Fila ${this.tempHeaderIdx + 1}`;
    const limit = Math.min(this.tempRawMatrix.length, 50);

    for (let i = 0; i < limit; i++) {
      this.buildPreviewRow(table, i, totalRows, footerSkip);
    }

    if (totalRows > limit) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="100" style="text-align:center; padding:4px; font-style:italic; color:var(--text-muted)">... ${
        totalRows - limit
      } filas más ...</td>`;
      table.appendChild(tr);

      const startEnd = Math.max(limit, totalRows - 5);
      for (let i = startEnd; i < totalRows; i++) {
        this.buildPreviewRow(table, i, totalRows, footerSkip);
      }
    }
  }

  buildPreviewRow(table, index, totalRows, footerSkip) {
    const rowData = this.tempRawMatrix[index];
    if (!rowData) return;

    const tr = document.createElement("tr");
    const isHeader = index === this.tempHeaderIdx;
    const isIgnoredTop = index < this.tempHeaderIdx;
    const isIgnoredBottom = index >= totalRows - footerSkip;

    if (isHeader) tr.className = "preview-header";
    else if (isIgnoredTop || isIgnoredBottom) tr.className = "preview-ignored";

    tr.onclick = () => {
      this.tempHeaderIdx = index;
      this.renderPreviewTableRows();
    };

    const tdNum = document.createElement("td");
    tdNum.className = "preview-row-num";
    tdNum.innerText = index + 1;
    tr.appendChild(tdNum);

    const colLimit = Math.min(rowData.length, 8);
    for (let j = 0; j < colLimit; j++) {
      const td = document.createElement("td");
      td.innerText = rowData[j] !== undefined ? rowData[j] : "";
      tr.appendChild(td);
    }

    if (rowData.length > colLimit) {
      const td = document.createElement("td");
      td.innerText = "...";
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  applyStructureAndLoad() {
    this.els.structureModal.classList.remove("active");
    this.setLoading(true);

    setTimeout(async () => {
      try {
        const headerIdx = this.tempHeaderIdx;
        const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;
        const headerRow = this.tempRawMatrix[headerIdx];
        if (!headerRow || headerRow.length === 0)
          throw new Error("La fila de encabezado seleccionada está vacía.");

        const columns = [];
        headerRow.forEach((colName, idx) => {
          let safeName =
            colName !== undefined &&
            colName !== null &&
            String(colName).trim() !== ""
              ? String(colName).trim()
              : `Columna_${idx + 1}`;
          if (columns.includes(safeName)) {
            let c = 1;
            while (columns.includes(`${safeName}_${c}`)) c++;
            safeName = `${safeName}_${c}`;
          }
          columns.push(safeName);
        });

        const startIndex = headerIdx + 1;
        const endIndex = this.tempRawMatrix.length - footerSkip;
        const jsonData = [];

        // Procesamiento por Lotes
        const CHUNK_SIZE = 5000;
        for (let i = startIndex; i < endIndex; i += CHUNK_SIZE) {
          const endChunk = Math.min(i + CHUNK_SIZE, endIndex);
          for (let j = i; j < endChunk; j++) {
            const rowArr = this.tempRawMatrix[j];
            if (!rowArr || rowArr.length === 0) continue;
            const rowObj = {};
            let hasData = false;

            columns.forEach((colKey, colIdx) => {
              const cellVal = rowArr[colIdx];
              rowObj[colKey] = cellVal !== undefined ? cellVal : "";
              if (cellVal !== undefined && cellVal !== "" && cellVal !== null)
                hasData = true;
            });
            if (hasData) jsonData.push(rowObj);
          }
          await new Promise((res) => setTimeout(res, 1));
        }

        if (jsonData.length === 0)
          throw new Error(
            "No se encontraron datos válidos con esta estructura."
          );

        this.tempRawMatrix = []; // Limpiar memoria
        this.initData(jsonData);
        this.setLoading(false);
        this.showToast(
          `Estructura aplicada. ${jsonData.length.toLocaleString()} filas cargadas.`,
          "success"
        );
      } catch (err) {
        console.error(err);
        this.showToast("Error procesando estructura: " + err.message, "error");
        this.setLoading(false);
      }
    }, 50);
  }

  initData(data, customColumns = null) {
    this.rawData = data;
    this.columns = customColumns || Object.keys(data[0]);

    const prefs = this.loadPreferences();
    if (prefs && prefs.pageSize) {
      this.pageSize = prefs.pageSize;
      const pageSizeSelect = document.getElementById("pageSize");
      if (pageSizeSelect) pageSizeSelect.value = this.pageSize;
    }

    this.colSettings = {};
    this.columns.forEach((col) => {
      this.colSettings[col] = {
        hidden: false,
        type: this.inferType(col, data),
        activeFilters: null,
        decimals: 2,
        dateStyle: "short",
        currency: "PEN",
        align: "auto",
        textStyle: "none"
      };

      if (prefs && prefs.colSettings && prefs.colSettings[col]) {
        const saved = prefs.colSettings[col];
        this.colSettings[col].hidden =
          saved.hidden !== undefined
            ? saved.hidden
            : this.colSettings[col].hidden;
        this.colSettings[col].type = saved.type || this.colSettings[col].type;
        this.colSettings[col].decimals =
          saved.decimals !== undefined
            ? saved.decimals
            : this.colSettings[col].decimals;
        this.colSettings[col].currency =
          saved.currency || this.colSettings[col].currency;
        this.colSettings[col].dateStyle =
          saved.dateStyle || this.colSettings[col].dateStyle;
        this.colSettings[col].align =
          saved.align || this.colSettings[col].align;
        this.colSettings[col].textStyle =
          saved.textStyle || this.colSettings[col].textStyle;
      }
    });

    this.els.globalSearch.disabled = false;
    this.buildColumnPicker();
    this.processData();

    this.els.emptyState.classList.add("hidden");
    this.els.tableWrapper.classList.remove("hidden");
    this.els.footer.classList.remove("hidden");
  }

  inferType(colName, data) {
    const lower = colName.toLowerCase();
    if (lower.match(/^(id|cod|sku|isbn|ean|item|ref|dni|ruc)/i)) return "text";
    if (lower.match(/(código|codigo|identificador)/i)) return "text";

    const sample = data.slice(0, 100).find((row) => row[colName] !== "");
    if (!sample) return "text";
    const val = sample[colName];

    if (val instanceof Date) return "date";
    if (typeof val === "number") {
      if (lower.match(/(precio|costo|total|valor|importe|venta|compra)/))
        return "currency";
      if (Number.isInteger(val)) return "integer";
      return "number";
    }
    if (String(val).startsWith("http")) return "link";
    return "text";
  }

  processData() {
    let processed = this.rawData.filter((row) => {
      return this.columns.every((col) => {
        const settings = this.colSettings[col];
        if (!settings.activeFilters) return true;
        return settings.activeFilters.has(String(row[col]));
      });
    });

    if (this.searchQuery) {
      processed = processed.filter((row) => {
        return Object.entries(row).some(([key, val]) => {
          if (this.colSettings[key].hidden) return false;
          return String(val).toLowerCase().includes(this.searchQuery);
        });
      });
    }

    if (this.sortCol) {
      processed.sort((a, b) => {
        let va = a[this.sortCol],
          vb = b[this.sortCol];
        if (typeof va === "string") va = va.toLowerCase();
        if (typeof vb === "string") vb = vb.toLowerCase();
        if (va < vb) return this.sortAsc ? -1 : 1;
        if (va > vb) return this.sortAsc ? 1 : -1;
        return 0;
      });
    }

    this.visibleData = processed;
    this.updatePaginationInfo();
    this.renderHeaders();
    this.render();
    this.renderFooterTotals();
    this.renderFilterSummary();
  }

  renderFilterSummary() {
    const bar = this.filterSummary;
    bar.innerHTML = "";
    const activeCols = this.columns.filter(
      (c) => this.colSettings[c].activeFilters !== null
    );

    if (activeCols.length === 0) {
      bar.classList.add("hidden");
      return;
    }

    bar.classList.remove("hidden");
    bar.innerHTML = `<span style="font-size:12px; font-weight:600; color:var(--text-muted)">Filtros activos:</span>`;

    activeCols.forEach((col) => {
      const chip = document.createElement("div");
      chip.className = "filter-chip";
      chip.innerHTML = `<span>${col}</span> <i class="ph ph-x" onclick="app.clearColFilter('${col}')"></i>`;
      bar.appendChild(chip);
    });

    if (activeCols.length > 1) {
      const clearAll = document.createElement("span");
      clearAll.className = "clear-filters-btn";
      clearAll.innerText = "Limpiar Todo";
      clearAll.onclick = () => {
        activeCols.forEach((c) => (this.colSettings[c].activeFilters = null));
        this.processData();
      };
      bar.appendChild(clearAll);
    }
  }

  renderHeaders() {
    this.els.thead.innerHTML = "";
    const tr = document.createElement("tr");

    this.columns.forEach((col) => {
      if (!this.colSettings[col] || this.colSettings[col].hidden) return;

      const th = document.createElement("th");
      const settings = this.colSettings[col];
      const alignClass = this.getAlignClass(settings.type);
      const isSorted = this.sortCol === col;
      const hasFilter = settings.activeFilters !== null;
      const iconClass = isSorted
        ? this.sortAsc
          ? "ph-arrow-up"
          : "ph-arrow-down"
        : "";
      const safeCol = String(col).replace(/'/g, "\\'");

      th.innerHTML = `
        <div class="th-content ${alignClass}">
          <div class="btn-col-menu ${
            hasFilter ? "active" : ""
          }" onclick="app.openColumnMenu(event, '${safeCol}')">
             <i class="ph ${
               hasFilter ? "ph-funnel ph-fill" : "ph-dots-three-vertical"
             }"></i>
          </div>
          <div class="th-title" onclick="app.sortBy('${safeCol}')">
            <span>${col}</span>
            ${isSorted ? `<i class="ph ${iconClass} sort-icon"></i>` : ""}
          </div>
        </div>
      `;
      tr.appendChild(th);
    });
    this.els.thead.appendChild(tr);
  }

  render() {
    this.els.tbody.innerHTML = "";
    const start = (this.currentPage - 1) * this.pageSize;
    const end = start + this.pageSize;
    const pageData = this.visibleData.slice(start, end);
    const fragment = document.createDocumentFragment();

    pageData.forEach((row) => {
      const tr = document.createElement("tr");

      tr.addEventListener("click", (e) => {
        if (e.target.tagName === "INPUT" || e.target.tagName === "A") return;
        const currentlySelected = this.els.tbody.querySelector(".row-selected");
        if (currentlySelected && currentlySelected !== tr)
          currentlySelected.classList.remove("row-selected");
        tr.classList.toggle("row-selected");
      });

      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const td = document.createElement("td");
        const config = this.colSettings[col];

        td.className = this.getAlignClass(config.type);
        td.innerHTML = this.formatValue(row[col], config);

        // APLICAR LAS NUEVAS REGLAS DE FORMATO VISUAL
        if (config.align && config.align !== "auto")
          td.style.textAlign = config.align;
        if (config.textStyle && config.textStyle !== "none")
          td.style.textTransform = config.textStyle;

        td.addEventListener("dblclick", () => this.enableEditing(td, row, col));
        tr.appendChild(td);
      });
      fragment.appendChild(tr);
    });
    this.els.tbody.appendChild(fragment);
    this.updateFooterUI();
  }

  renderFooterTotals() {
    this.els.tfoot.innerHTML = "";
    let hasTotals = false;
    const tr = document.createElement("tr");

    this.columns.forEach((col, idx) => {
      if (this.colSettings[col].hidden) return;
      const td = document.createElement("td");
      const config = this.colSettings[col];
      const type = config.type;

      if (["number", "currency", "integer", "percent"].includes(type)) {
        const sum = this.visibleData.reduce(
          (acc, r) => acc + (parseFloat(r[col]) || 0),
          0
        );
        if (sum !== 0 && type !== "percent") {
          hasTotals = true;
          td.className = "text-right";
          td.innerHTML = this.formatValue(sum, config);
        }
      }
      if (idx === 0 && !hasTotals) td.innerText = "Totales";
      tr.appendChild(td);
    });
    if (hasTotals) this.els.tfoot.appendChild(tr);
  }

  sortBy(col) {
    if (this.sortCol === col) this.sortAsc = !this.sortAsc;
    else {
      this.sortCol = col;
      this.sortAsc = true;
    }
    this.processData();
  }

  changePage(delta) {
    const maxPages = Math.ceil(this.visibleData.length / this.pageSize);
    const newPage = this.currentPage + delta;
    if (newPage >= 1 && newPage <= maxPages) {
      this.currentPage = newPage;
      this.render();
      this.els.tableWrapper.scrollTop = 0;
    }
  }

  updatePaginationInfo() {
    const total = this.visibleData.length;
    document.getElementById(
      "statusMsg"
    ).innerText = `${total.toLocaleString()} registros encontrados`;
    this.updateFooterUI();
  }

  updateFooterUI() {
    const maxPages = Math.ceil(this.visibleData.length / this.pageSize) || 1;
    document.getElementById("currPage").innerText = this.currentPage;
    document.getElementById("totalPages").innerText = maxPages;
    document.getElementById("btnPrev").disabled = this.currentPage <= 1;
    document.getElementById("btnNext").disabled = this.currentPage >= maxPages;
  }

  openColumnMenu(e, col) {
    e.stopPropagation();
    this.activeMenuCol = col;
    const menu = this.els.ctxMenu;

    const rect = e.currentTarget.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.left;
    if (left + 280 > window.innerWidth) left = window.innerWidth - 290;
    if (left < 0) left = 10;

    menu.style.top = top + "px";
    menu.style.left = left + "px";

    this.renderMenuContent(col, menu);
    menu.classList.add("show");

    const menuRect = menu.getBoundingClientRect();
    if (menuRect.bottom > window.innerHeight) {
      menu.style.top = "auto";
      menu.style.bottom = "10px";
    }
  }

  renderMenuContent(col, container) {
    const settings = this.colSettings[col];
    const relevantRows = this.rawData.filter((row) => {
      return this.columns.every((c) => {
        if (c === col) return true;
        const s = this.colSettings[c];
        if (!s.activeFilters) return true;
        return s.activeFilters.has(String(row[c]));
      });
    });
    const uniqueVals = [
      ...new Set(relevantRows.map((r) => String(r[col])))
    ].sort();

    let extraControls = "";
    if (["number", "currency", "percent"].includes(settings.type)) {
      extraControls += `<div style="margin-top:8px; display:flex; align-items:center; justify-content:space-between;"><label class="col-menu-label" style="margin:0">Decimales</label><input type="number" min="0" max="6" class="form-input form-input-sm" style="width:60px" value="${settings.decimals}" onchange="app.changeColDecimal('${col}', this.value)"></div>`;
    }

    if (settings.type === "currency") {
      extraControls += `<div style="margin-top:8px"><label class="col-menu-label" style="margin-bottom:2px">Simbolo</label><select class="form-select" onchange="app.changeColCurrency('${col}', this.value)"><option value="PEN" ${
        settings.currency === "PEN" ? "selected" : ""
      }>S/ (PEN)</option><option value="USD" ${
        settings.currency === "USD" ? "selected" : ""
      }>$ (USD)</option><option value="EUR" ${
        settings.currency === "EUR" ? "selected" : ""
      }>€ (EUR)</option></select></div>`;
    }

    if (["date", "datetime"].includes(settings.type)) {
      extraControls += `<div style="margin-top:8px"><label class="col-menu-label" style="margin-bottom:2px">Estilo</label><select class="form-select" onchange="app.changeColDateStyle('${col}', this.value)">
        <option value="short" ${
          settings.dateStyle === "short" ? "selected" : ""
        }>Corto (DD/MM/YYYY)</option>
        <option value="medium" ${
          settings.dateStyle === "medium" ? "selected" : ""
        }>Medio (04 ene 2026)</option>
        <option value="long" ${
          settings.dateStyle === "long" ? "selected" : ""
        }>Largo (4 de enero...)</option>
        <option value="full" ${
          settings.dateStyle === "full" ? "selected" : ""
        }>Texto (Miércoles...)</option>
        </select></div>`;
    }

    // --- OPCIONES DE FORMATO DE TEXTO Y ALINEACION ---
    extraControls += `
      <div style="margin-top:8px; padding-top:8px; border-top:1px dashed var(--border);">
        <label class="col-menu-label" style="margin-bottom:2px">Alineación</label>
        <select class="form-select" onchange="app.changeColAlign('${col}', this.value)">
           <option value="auto" ${
             settings.align === "auto" || !settings.align ? "selected" : ""
           }>Automática</option>
           <option value="left" ${
             settings.align === "left" ? "selected" : ""
           }>Izquierda</option>
           <option value="center" ${
             settings.align === "center" ? "selected" : ""
           }>Centro</option>
           <option value="right" ${
             settings.align === "right" ? "selected" : ""
           }>Derecha</option>
        </select>
      </div>
      <div style="margin-top:8px">
        <label class="col-menu-label" style="margin-bottom:2px">Mayús / Minús</label>
        <select class="form-select" onchange="app.changeColTextStyle('${col}', this.value)">
           <option value="none" ${
             settings.textStyle === "none" || !settings.textStyle
               ? "selected"
               : ""
           }>Normal</option>
           <option value="uppercase" ${
             settings.textStyle === "uppercase" ? "selected" : ""
           }>MAYÚSCULAS</option>
           <option value="lowercase" ${
             settings.textStyle === "lowercase" ? "selected" : ""
           }>minúsculas</option>
           <option value="capitalize" ${
             settings.textStyle === "capitalize" ? "selected" : ""
           }>Capitalizar</option>
        </select>
      </div>
    `;

    container.innerHTML = `
        <div class="col-menu-section">
          <label class="col-menu-label">Formato</label>
          <select class="form-select" onchange="app.changeColFormat('${col}', this.value)">
            <option value="auto" ${
              settings.type === "auto" ? "selected" : ""
            }>Automático</option>
            <option value="text" ${
              settings.type === "text" ? "selected" : ""
            }>Texto</option>
            <option value="number" ${
              settings.type === "number" ? "selected" : ""
            }>Número</option>
            <option value="integer" ${
              settings.type === "integer" ? "selected" : ""
            }>Entero</option>
            <option value="currency" ${
              settings.type === "currency" ? "selected" : ""
            }>Moneda</option>
            <option value="percent" ${
              settings.type === "percent" ? "selected" : ""
            }>Porcentaje (%)</option>
            <option value="date" ${
              settings.type === "date" ? "selected" : ""
            }>Fecha</option>
            <option value="datetime" ${
              settings.type === "datetime" ? "selected" : ""
            }>Fecha y Hora</option>
            <option value="time" ${
              settings.type === "time" ? "selected" : ""
            }>Hora</option>
            <option value="link" ${
              settings.type === "link" ? "selected" : ""
            }>Enlace (URL)</option>
          </select>
          ${extraControls}
        </div>
        <div class="col-menu-section">
          <label class="col-menu-label">Filtrar (${uniqueVals.length})</label>
          <input type="text" class="form-input form-input-sm" placeholder="Buscar..." oninput="app.filterMenuSearch(this.value)">
          <div class="filter-list" id="filterListContainer"></div>
          <div style="display:flex; justify-content:space-between; margin-top:8px;">
             <button class="btn btn-sm" onclick="app.clearColFilter('${col}')">Limpiar</button>
             <button class="btn btn-sm btn-primary" onclick="app.applyColFilter('${col}')">Aplicar</button>
          </div>
        </div>
      `;

    const filterContainer = container.querySelector("#filterListContainer");
    const allDiv = document.createElement("div");
    allDiv.className = "filter-item";
    allDiv.innerHTML = `<input type="checkbox" id="chkAllFilters" ${
      settings.activeFilters === null ? "checked" : ""
    }> <span>(Seleccionar Todo)</span>`;
    allDiv.onclick = (ev) => {
      if (ev.target.tagName !== "INPUT") {
        const chk = allDiv.querySelector("input");
        chk.checked = !chk.checked;
      }
      const state = allDiv.querySelector("input").checked;
      filterContainer
        .querySelectorAll(".val-chk")
        .forEach((c) => (c.checked = state));
    };
    filterContainer.appendChild(allDiv);

    const maxItems = 2000;
    const isLimited = uniqueVals.length > maxItems;
    const valsToShow = isLimited ? uniqueVals.slice(0, maxItems) : uniqueVals;

    valsToShow.forEach((val) => {
      const div = document.createElement("div");
      div.className = "filter-item val-item";
      div.setAttribute("data-val", val.toLowerCase());
      const displayVal = val === "" ? "(Vacío)" : this.escapeHTML(val);
      const isChecked =
        settings.activeFilters === null
          ? true
          : settings.activeFilters.has(val);
      div.innerHTML = `<input type="checkbox" class="val-chk" value="${this.escapeHTML(
        val
      )}" ${isChecked ? "checked" : ""}> <span>${displayVal}</span>`;
      div.onclick = (ev) => {
        if (ev.target.tagName !== "INPUT") {
          const chk = div.querySelector("input");
          chk.checked = !chk.checked;
        }
        if (!div.querySelector("input").checked) {
          const ac = container.querySelector("#chkAllFilters");
          if (ac) ac.checked = false;
        }
      };
      filterContainer.appendChild(div);
    });
  }

  filterMenuSearch(term) {
    term = term.toLowerCase();
    const items = document.querySelectorAll("#filterListContainer .val-item");
    items.forEach((el) => {
      el.style.display = el.getAttribute("data-val").includes(term)
        ? "flex"
        : "none";
    });
  }

  applyColFilter(col) {
    const inputs = document.querySelectorAll("#filterListContainer .val-chk");
    const allChk = document.getElementById("chkAllFilters");

    if (allChk && allChk.checked) {
      this.colSettings[col].activeFilters = null;
    } else {
      const selected = new Set();
      inputs.forEach((inp) => {
        if (inp.checked) selected.add(inp.value);
      });
      this.colSettings[col].activeFilters = selected;
    }
    this.els.ctxMenu.classList.remove("show");
    this.currentPage = 1;
    this.processData();
  }

  clearColFilter(col) {
    this.colSettings[col].activeFilters = null;
    this.els.ctxMenu.classList.remove("show");
    this.processData();
  }

  changeColFormat(col, type) {
    this.colSettings[col].type = type;
    this.updateAndKeepMenu(col);
  }
  changeColDecimal(col, val) {
    this.colSettings[col].decimals = parseInt(val) || 0;
    this.updateAndKeepMenu(col);
  }
  changeColDateStyle(col, val) {
    this.colSettings[col].dateStyle = val;
    this.updateAndKeepMenu(col);
  }
  changeColCurrency(col, val) {
    this.colSettings[col].currency = val;
    this.updateAndKeepMenu(col);
  }
  changeColAlign(col, val) {
    this.colSettings[col].align = val;
    this.updateAndKeepMenu(col);
  }
  changeColTextStyle(col, val) {
    this.colSettings[col].textStyle = val;
    this.updateAndKeepMenu(col);
  }

  updateAndKeepMenu(col) {
    this.savePreferences();
    this.render();
    this.renderHeaders();
    this.renderFooterTotals();
    this.renderMenuContent(col, this.els.ctxMenu);
  }

  enableEditing(td, row, col) {
    if (td.querySelector("input")) return;
    const currentVal = row[col];
    const config = this.colSettings[col];
    const type = config.type;
    const originalHtml = td.innerHTML;

    td.classList.add("cell-editing");
    td.innerHTML = "";
    const input = document.createElement("input");
    input.className = "table-input";

    if (["number", "currency", "integer", "percent"].includes(type)) {
      input.type = "number";
      input.step = "any";
      input.value = currentVal;
    } else if (type === "date" || type === "datetime") {
      input.type = type === "datetime" ? "datetime-local" : "date";
      try {
        if (currentVal instanceof Date) {
          input.value =
            type === "datetime"
              ? currentVal.toISOString().slice(0, 16)
              : currentVal.toISOString().split("T")[0];
        } else if (currentVal) {
          const d = new Date(currentVal);
          input.value =
            type === "datetime"
              ? d.toISOString().slice(0, 16)
              : d.toISOString().split("T")[0];
        }
      } catch (e) {
        input.value = "";
      }
    } else {
      input.type = "text";
      input.value = currentVal !== undefined ? currentVal : "";
    }

    input.addEventListener("blur", () =>
      this.saveEdit(td, row, col, input.value)
    );
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") input.blur();
      else if (e.key === "Escape") {
        td.classList.remove("cell-editing");
        td.innerHTML = originalHtml;
      }
    });
    td.appendChild(input);
    input.focus();
  }

  saveEdit(td, row, col, newVal) {
    const config = this.colSettings[col];
    const type = config.type;
    let finalVal = newVal;

    if (["number", "currency", "percent"].includes(type))
      finalVal = newVal === "" ? 0 : parseFloat(newVal);
    else if (type === "integer")
      finalVal = newVal === "" ? 0 : parseInt(newVal);
    else if (type === "date" || type === "datetime") {
      if (newVal) {
        // Reconstruir fecha considerando formato
        finalVal = new Date(newVal);
      } else {
        finalVal = "";
      }
    }

    const oldVal = row[col];
    if (oldVal !== finalVal) {
      this.undoStack.push({ row: row, col: col, oldVal: oldVal });
      if (this.undoStack.length > 50) this.undoStack.shift();
    }

    row[col] = finalVal;
    td.classList.remove("cell-editing");
    td.innerHTML = this.formatValue(finalVal, config);

    if (config.align && config.align !== "auto")
      td.style.textAlign = config.align;
    if (config.textStyle && config.textStyle !== "none")
      td.style.textTransform = config.textStyle;

    this.renderFooterTotals();
    td.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
    setTimeout(() => (td.style.backgroundColor = ""), 500);
  }

  undo() {
    if (this.undoStack.length === 0) return;
    const lastAction = this.undoStack.pop();
    lastAction.row[lastAction.col] = lastAction.oldVal;
    this.render();
    this.renderFooterTotals();
    this.showToast("Edición deshecha", "info");
  }

  buildColumnPicker() {
    this.renderColumnList(this.columns);
  }
  renderColumnList(cols) {
    this.els.colListContainer.innerHTML = "";
    cols.forEach((col) => {
      const item = document.createElement("div");
      item.className = "dropdown-item";
      item.innerHTML = `<input type="checkbox" ${
        !this.colSettings[col].hidden ? "checked" : ""
      }><span>${col}</span>`;
      item.onclick = (e) => {
        const chk = item.querySelector("input");
        if (e.target !== chk) chk.checked = !chk.checked;
        this.colSettings[col].hidden = !chk.checked;
        this.savePreferences();
        this.processData();
      };
      this.els.colListContainer.appendChild(item);
    });
  }
  filterColumnList(term) {
    this.renderColumnList(
      this.columns.filter((c) => c.toLowerCase().includes(term.toLowerCase()))
    );
  }
  toggleAllColumns(show) {
    this.columns.forEach((col) => (this.colSettings[col].hidden = !show));
    this.savePreferences();
    this.buildColumnPicker();
    this.processData();
  }

  exportTo(format) {
    this.els.exportMenu.classList.remove("show");
    if (!this.visibleData.length)
      return this.showToast("No hay datos", "error");
    this.pendingExportFormat = format;
    this.els.confirmTitle.value = this.els.reportTitle.value;
    this.els.confirmAuthor.value = this.els.reportAuthor.value;
    this.els.exportModal.classList.add("active");
  }

  executeExport(format) {
    let fname = this.els.reportTitle.value.trim() || "Reporte";
    fname = fname.replace(/[^a-z0-9_\-\sáéíóúñ]/gi, "_");
    const author = this.els.reportAuthor.value.trim();
    const timestamp = new Date().toLocaleString();

    const exportData = this.visibleData.map((row) => {
      const newRow = {};
      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const val = row[col];
        const config = this.colSettings[col];

        if (format === "xlsx") {
          newRow[col] = config.type === "text" ? String(val) : val;
        } else {
          // Extraer solo texto, ignorar HTML de enlaces si existieran
          const d = document.createElement("div");
          d.innerHTML = this.formatValue(val, config);
          newRow[col] = d.textContent.trim();
        }
      });
      return newRow;
    });

    if (format === "xlsx") {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      XLSX.writeFile(wb, `${fname}.xlsx`);
    } else if (format === "pdf") {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: "landscape" });
      doc.text(fname, 14, 15);
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text(`Autor: ${author} | ${timestamp}`, 14, 22);
      doc.autoTable({
        head: [Object.keys(exportData[0])],
        body: exportData.map(Object.values),
        startY: 28,
        theme: "grid",
        styles: { fontSize: 8 },
        headStyles: { fillColor: [59, 130, 246] }
      });
      doc.save(`${fname}.pdf`);
    } else if (format === "html") {
      // Recreamos el HTML sin insignias y con los valores puros formateados
      let rowsHtmlArray = [];
      let totalRowHtml = "";
      const totals = {};

      this.columns.forEach((col) => {
        const type = this.colSettings[col].type;
        if (
          ["number", "currency", "integer", "percent"].includes(type) &&
          !this.colSettings[col].hidden
        ) {
          totals[col] = 0;
        }
      });

      this.visibleData.forEach((row) => {
        let tr = "<tr>";
        this.columns.forEach((col) => {
          if (this.colSettings[col].hidden) return;
          const config = this.colSettings[col];
          const type = config.type;
          let val = row[col];
          let cellHtml = "";
          let alignClass = "text-left";
          let cssClass = "col-text";
          let dataAttrs = "";

          // Aplicar la alineación personalizada de exportación si existe
          if (config.align && config.align !== "auto")
            alignClass = `text-${config.align}`;

          const isNum = typeof val === "number";
          if (isNum && !["date", "datetime", "time"].includes(type)) {
            if (config.align === "auto") alignClass = "text-right";
            cssClass = "col-num";
            dataAttrs = ` data-val="${val}"`;
            if (totals[col] !== undefined) totals[col] += val;
          }

          cellHtml = this.formatValue(val, config);

          // Estilo de texto CSS (mayúsculas/minúsculas)
          let styleAttr = "";
          if (config.textStyle && config.textStyle !== "none") {
            styleAttr = ` style="text-transform: ${config.textStyle};"`;
          }

          tr += `<td class="${cssClass} ${alignClass}"${dataAttrs}${styleAttr}>${cellHtml}</td>`;
        });
        tr += "</tr>";
        rowsHtmlArray.push(tr);
      });

      let rowsHtml = rowsHtmlArray.join("");

      let hasTotals = Object.keys(totals).length > 0;
      if (hasTotals) {
        totalRowHtml = '<tr class="row-total">';
        this.columns.forEach((col, idx) => {
          if (this.colSettings[col].hidden) return;
          let td = "";
          if (idx === 0) td = "<td>TOTAL</td>";
          else {
            if (totals[col] !== undefined) {
              const config = this.colSettings[col];
              const sum = totals[col];
              const fmtVal = this.formatValue(sum, config);
              td = `<td class="text-right" data-sum="1" data-fmt="${this.escapeHTML(
                config.currency || ""
              )}">${fmtVal}</td>`;
            } else {
              td = "<td></td>";
            }
          }
          totalRowHtml += td;
        });
        totalRowHtml += "</tr>";
      }

      let headersHtml = "";
      let colIndex = 0;
      let filterOptions = '<option value="all">Todas las columnas</option>';
      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const config = this.colSettings[col];
        const isNum = ["number", "currency", "integer", "percent"].includes(
          config.type
        );
        let align = isNum ? "text-right" : "text-left";
        if (config.align && config.align !== "auto")
          align = `text-${config.align}`;

        headersHtml += `<th class="${align}" onclick="sortGrid(${colIndex})">${this.escapeHTML(
          col
        )}</th>`;
        filterOptions += `<option value="${colIndex}">${this.escapeHTML(
          col
        )}</option>`;
        colIndex++;
      });

      // El mismo HTML core para interactividad
      const fullHtml = `<!DOCTYPE html>
<html lang='es'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'>
  <title>${this.escapeHTML(fname)}</title>
  <style>
    :root { --bg-body: #f8fafc; --bg-gradient: linear-gradient(135deg, #f0f9ff 0%, #e0e7ff 100%); --bg-card: #ffffff; --text-main: #0f172a; --text-muted: #64748b; --primary: #0f172a; --accent: #0ea5e9; --border: #e2e8f0; --table-head: #f8fafc; --table-head-text: #334155; --row-hover: #f1f5f9; --shadow-soft: 0 10px 30px -10px rgba(0,0,0,0.08); --shadow-float: 0 20px 25px -5px rgba(0, 0, 0, 0.1); --link-color: #0284c7; }
    [data-theme='dark'] { --bg-body: #0f172a; --bg-gradient: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%); --bg-card: #1e293b; --text-main: #f8fafc; --text-muted: #94a3b8; --primary: #818cf8; --accent: #38bdf8; --border: #334155; --table-head: #1e293b; --table-head-text: #cbd5e1; --row-hover: #334155; --shadow-soft: 0 10px 30px -10px rgba(0,0,0,0.5); --link-color: #7dd3fc; }
    *, *::before, *::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; transition: all 0.2s ease; }
    html, body { height: 100%; height: 100dvh; margin: 0; padding: 0; overflow: hidden; font-family: 'Inter', system-ui, -apple-system, sans-serif; background: var(--bg-gradient); color: var(--text-main); }
    a { color: var(--link-color); text-decoration: none; font-weight: 500; }
    a:hover { text-decoration: underline; }
    .page { height: 100%; display: flex; flex-direction: column; padding: 24px; max-width: 2000px; margin: 0 auto; gap: 20px; align-items: center; }
    .header-container { display: flex; flex-direction: column; gap: 20px; flex-shrink: 0; width: 100%; max-width: 100%; }
    .title-area { text-align: center; }
    .title-area h1 { font-size: 28px; font-weight: 800; letter-spacing: -0.5px; margin: 0; background: linear-gradient(to right, var(--primary), var(--accent)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .subtitle-area { font-size: 12px; color: var(--text-muted); font-weight: 400; font-style: italic; }
    .meta-row { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin: 0 auto; }
    .actions-area { display: flex; align-items: center; gap: 12px; }
    .author-pill { display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px; background: rgba(79, 70, 229, 0.1); color: var(--primary); border-radius: 20px; font-size: 11px; font-weight: 600; letter-spacing: 0.5px; }
    .btn-group { display: flex; gap: 8px; }
    .btn { display: inline-flex; align-items: center; justify-content: center; height: 36px; border-radius: 99px; border: 1px solid var(--border); background: var(--bg-card); color: var(--text-muted); cursor: pointer; padding: 0 14px; font-size: 13px; font-weight: 600; gap: 6px; }
    .btn:hover { background: var(--bg-body); color: var(--primary); transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-color: var(--primary); }
    .btn-primary { background: var(--bg-card); color: #fff; border: none; }
    .btn-primary:hover { background: var(--accent); color: #fff; box-shadow: 0 4px 12px rgba(14, 165, 233, 0.3); }
    .search-floater { width: 100%; max-width: 600px; height: 40px; background: var(--bg-card); border-radius: 16px; padding: 3px; align-items: center; box-shadow: var(--shadow-float); display: flex; gap: 8px; border: 1px solid var(--border); animation: floatUp 0.6s ease-out; }
    .select-wrapper { position: relative; border-right: 1px solid var(--border); }
    .filter-select { appearance: none; background: transparent; border: none; padding: 10px 30px 10px 16px; font-size: 13px; font-weight: 600; color: var(--text-main); cursor: pointer; outline: none; height: 100%; }
    .filter-select option { background-color: var(--bg-card); color: var(--text-main); }
    .select-arrow { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); pointer-events: none; color: var(--text-muted); width: 12px; }
    .search-input-wrapper { flex-grow: 1; position: relative; }
    .search-input { width: 100%; border: none; background: transparent; padding: 10px 12px; font-size: 14px; color: var(--text-main); outline: none; }
    .table-card { width: fit-content; max-width: 100%; background: rgba(255, 255, 255, 0.7); backdrop-filter: blur(10px); border-radius: 20px; box-shadow: var(--shadow-soft); flex-grow: 1; overflow: hidden; border: 1px solid rgba(255,255,255,0.5); display: flex; flex-direction: column; }
    [data-theme='dark'] .table-card { background: rgba(30, 41, 59, 0.7); border-color: rgba(255,255,255,0.1); }
    .table-container { overflow: auto; flex-grow: 1; position: relative; width: 100%; }
    table { width: auto; border-collapse: separate; border-spacing: 0; font-size: 13px; }
    th, td { white-space: nowrap; }
    thead th { position: sticky; top: 0; background: var(--bg-card); color: var(--text-muted); padding: 10px 16px; font-weight: 700; text-transform: uppercase; font-size: 14px; letter-spacing: 0.8px; border-bottom: 2px solid var(--border); cursor: pointer; z-index: 20; transition: background 0.2s; }
    thead th:hover { background: var(--row-hover); color: var(--primary); }
    thead th::after { content: ''; display: inline-block; margin-left: 8px; vertical-align: middle; border-left: 4px solid transparent; border-right: 4px solid transparent; opacity: 0; transition: opacity 0.2s; }
    thead th:hover::after { opacity: 0.5; border-top: 4px solid currentColor; }
    thead th.asc::after { opacity: 1; border-bottom: 4px solid var(--accent); border-top: none; }
    thead th.desc::after { opacity: 1; border-top: 4px solid var(--accent); border-bottom: none; }
    .text-left { text-align: left; } .text-right { text-align: right; } .text-center { text-align: center; }
    td { padding: 8px 16px; border-bottom: 1px solid var(--border); color: var(--text-main); font-weight: 500; height: 38px; }
    tbody tr:hover { background-color: var(--row-hover); }
    .row-total td { position: sticky; bottom: 0; background-color: var(--bg-card); border-top: 2px solid var(--primary); color: var(--primary); font-weight: 800; z-index: 30; box-shadow: 0 -4px 20px rgba(0,0,0,0.1); padding: 10px 16px; }
    
    .modal-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(15, 23, 42, 0.6); backdrop-filter: blur(4px); z-index: 9999; opacity: 0; visibility: hidden; transition: all 0.25s ease; display: flex; align-items: center; justify-content: center; padding: 20px; }
    .modal-overlay.active { opacity: 1; visibility: visible; }
    .modal-card { background: var(--bg-card); width: 100%; max-width: 450px; max-height: 85vh; border-radius: 12px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); transform: scale(0.95); transition: all 0.25s; display: flex; flex-direction: column; overflow: hidden; border: 1px solid var(--border); }
    .modal-overlay.active .modal-card { transform: scale(1); }
    .modal-header { padding: 16px 24px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; background: var(--bg-card); z-index: 10; }
    .modal-title { font-size: 18px; font-weight: 700; margin:0; color: var(--primary); }
    .modal-body { padding: 0; overflow-y: auto; display: flex; flex-direction: column; }
    .detail-item { padding: 16px 24px; border-bottom: 1px dashed var(--border); display: flex; flex-direction: column; gap: 4px; }
    .detail-item:last-child { border-bottom: none; }
    .detail-label { font-size: 11px; text-transform: uppercase; color: var(--text-muted); font-weight: 600; letter-spacing: 0.5px; }
    .detail-value { font-size: 15px; color: var(--text-main); font-weight: 500; word-break: break-word; line-height: 1.5; }
    @keyframes floatUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
    @media print {
      html, body, .page, .table-card, .table-container { height: auto !important; overflow: visible !important; width: 100% !important; margin: 0 !important; padding: 0 !important; display: block !important; background: #fff !important; color: #000 !important; }
      .search-floater, .actions-area, .btn, .modal-overlay, .select-wrapper, #filterCol { display: none !important; }
      table { width: 100% !important; border-collapse: collapse !important; }
      th, td { white-space: normal !important; border: 1px solid #000 !important; font-size: 10px !important; color: #000 !important; }
      thead th { background: #eee !important; color: #000 !important; border-bottom: 2px solid #000 !important; }
      .row-total td { background: #f0f0f0 !important; color: #000 !important; border-top: 2px solid #000 !important; }
      .header-container { display: block !important; padding-bottom: 20px; border-bottom: 2px solid #000; margin-bottom: 20px; }
      .title-area h1 { color: #000 !important; -webkit-text-fill-color: initial !important; text-align: left !important; font-size: 18px !important; }
      .table-card { box-shadow: none !important; border: none !important; }
      :root { --bg-body: #ffffff !important; --text-main: #000000 !important; --bg-card: #ffffff !important; }
    }
  </style>
  <script>
    var curSort = { col: -1, dir: 'asc' };
    function printRpt(){window.print()}
    function toggleTheme() { var h = document.documentElement; var c = h.getAttribute('data-theme'); var n = c === 'dark' ? 'light' : 'dark'; h.setAttribute('data-theme', n); localStorage.setItem('theme', n); }
    function updateDatalist() { var colIdx = document.getElementById('filterCol').value; var list = document.getElementById('search-options'); list.innerHTML = ''; var set = new Set(); var rows = document.querySelectorAll('tbody tr:not(.row-total)'); rows.forEach(function(row) { var td; if(colIdx === 'all') {} else { td = row.children[parseInt(colIdx)]; if(td) { var txt = td.innerText.trim(); if(txt && txt.length > 1 && txt.length < 50) set.add(txt); } } }); var arr = Array.from(set).sort().slice(0, 500); arr.forEach(function(val) { var opt = document.createElement('option'); opt.value = val; list.appendChild(opt); }); }
    function onSelectorChange() { document.getElementById('search').value = ''; filterTbl(); updateDatalist(); }
    function filterTbl(){ var term = document.getElementById('search').value.toLowerCase(); var colIdx = document.getElementById('filterCol').value; var tbody = document.querySelector('tbody'); var rows = tbody.querySelectorAll('tr:not(.row-total)'); var totalRow = tbody.querySelector('.row-total'); var sums = []; rows.forEach(function(row) { var visible = false; var tds = row.querySelectorAll('td'); if(colIdx === 'all') { for(var i=0; i<tds.length; i++) { if(tds[i].innerText.toLowerCase().indexOf(term) > -1) { visible = true; break; } } } else { var targetTd = tds[colIdx]; if(targetTd && targetTd.innerText.toLowerCase().indexOf(term) > -1) visible = true; } row.style.display = visible ? '' : 'none'; if(visible) { tds.forEach(function(td, idx) { var val = parseFloat(td.getAttribute('data-val')); if(!isNaN(val)) { if(!sums[idx]) sums[idx] = 0; sums[idx] += val; } }); } }); if(totalRow) { totalRow.querySelectorAll('td').forEach(function(td, n) { if(td.hasAttribute('data-sum')) { var fmt = td.getAttribute('data-fmt')||''; var isPct = fmt.includes('%'); var sum = sums[n] || 0; var txt = ''; if(isPct) txt = (sum * 100).toFixed(2) + '%'; else txt = sum.toLocaleString('es-PE', {minimumFractionDigits:2, maximumFractionDigits:2}); if(fmt && !isPct) txt = fmt + ' ' + txt; td.innerText = txt; } }); } }
    function sortGrid(idx) { var tbody = document.querySelector('tbody'); var rows = Array.from(tbody.querySelectorAll('tr:not(.row-total)')); var totalRow = tbody.querySelector('.row-total'); if (curSort.col === idx) { curSort.dir = curSort.dir === 'asc' ? 'desc' : 'asc'; } else { curSort.col = idx; curSort.dir = 'asc'; } document.querySelectorAll('thead th').forEach(function(th) { th.className = th.className.replace(/ asc| desc/g, ''); }); document.querySelectorAll('thead th')[idx].className += ' ' + curSort.dir; rows.sort(function(a, b) { var cellA = a.children[idx]; var cellB = b.children[idx]; var valA = cellA.hasAttribute('data-val') ? parseFloat(cellA.getAttribute('data-val')) : null; var valB = cellB.hasAttribute('data-val') ? parseFloat(cellB.getAttribute('data-val')) : null; if (valA !== null && valB !== null) return curSort.dir === 'asc' ? valA - valB : valB - valA; var txtA = cellA.innerText.trim().toLowerCase(); var txtB = cellB.innerText.trim().toLowerCase(); return curSort.dir === 'asc' ? txtA.localeCompare(txtB) : txtB.localeCompare(txtA); }); rows.forEach(function(r) { tbody.appendChild(r); }); if(totalRow) tbody.appendChild(totalRow); }
    function dlXLS() { var rows = document.querySelectorAll('table tr'); var csv = []; rows.forEach(function(row) { if(row.style.display !== 'none') { var cols = []; row.querySelectorAll('th, td').forEach(function(cell) { cols.push('"' + cell.innerText.replace(/"/g, '""') + '"'); }); csv.push(cols.join(';')); } }); var blob = new Blob(['\\uFEFF' + csv.join('\\r\\n')], { type: 'text/csv;charset=utf-8;' }); var url = URL.createObjectURL(blob); var a = document.createElement('a'); a.href = url; a.download = 'reporte.csv'; a.click(); }
    document.addEventListener('DOMContentLoaded', function() { updateDatalist(); var saved = localStorage.getItem('theme') || 'light'; document.documentElement.setAttribute('data-theme', saved); var rows = document.querySelectorAll('tbody tr:not(.row-total)'); var modal = document.getElementById('detailModal'); var modalBody = modal.querySelector('.modal-body'); var modalTitle = modal.querySelector('.modal-title'); var headers = Array.from(document.querySelectorAll('thead th')).map(function(th) { return th.innerText; }); function showModal(row) { modalBody.innerHTML = ''; var cells = row.querySelectorAll('td'); modalTitle.innerText = cells[0].innerText || 'Detalle'; cells.forEach(function(cell, index) { var val = cell.innerText; if (cell.querySelector('a')) val = cell.innerHTML; var item = document.createElement('div'); item.className = 'detail-item'; item.innerHTML = '<div class="detail-label">' + headers[index] + '</div><div class="detail-value">' + val + '</div>'; modalBody.appendChild(item); }); modal.classList.add('active'); } rows.forEach(function(row) { row.addEventListener('dblclick', function() { showModal(row); }); }); });
    function closeModal() { document.getElementById('detailModal').classList.remove('active'); }
  </script>
</head>
<body>
  <div id='detailModal' class='modal-overlay' onclick='if(event.target === this) closeModal()'>
    <div class='modal-card'>
      <div class='modal-header'>
        <h3 class='modal-title'>Detalle</h3><button class='btn' style='border:none' onclick='closeModal()'><svg width='20' height='20' fill='none' stroke='currentColor' stroke-width='2' viewBox='0 0 24 24'><path d='M18 6L6 18M6 6l12 12'></path></svg></button>
      </div>
      <div class='modal-body'></div>
    </div>
  </div>
  <div class='page'>
    <div class='header-container'>
      <div class='title-area'>
        <h1>${this.escapeHTML(fname)}</h1>
        <div class='author-pill'>${this.escapeHTML(
          author
        )} <div class='subtitle-area'>Generado: ${timestamp}</div></div>
      </div>
      <div class='meta-row'>
        <datalist id='search-options'></datalist>
        <div class='search-floater'>
          <div class='select-wrapper'><select id='filterCol' class='filter-select' onchange='onSelectorChange()'>${filterOptions}</select></div>
          <div class='search-input-wrapper'><input type='text' id='search' list='search-options' autocomplete='on' class='search-input' onkeyup='filterTbl()' placeholder='Buscar...'></div>
        </div>
        <div class='actions-area'>
          <div class='btn-group'>
            <button class='btn btn-primary' onclick='toggleTheme()' title='Tema'>Tema</button>
            <button class='btn btn-primary' onclick='dlXLS()' title='Exportar CSV'>CSV</button>
            <button class='btn btn-primary' onclick='printRpt()' title='Imprimir'>Imprimir</button>
          </div>
        </div>
      </div>
    </div>
    <div class='table-card'>
      <div class='table-container'>
        <table><thead><tr>${headersHtml}</tr></thead><tbody>${rowsHtml}${totalRowHtml}</tbody></table>
      </div>
    </div>
  </div>
</body>
</html>`;

      const url = URL.createObjectURL(
        new Blob([fullHtml], { type: "text/html" })
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = `${fname}.html`;
      a.click();
    }
    this.showToast(`Exportado a ${format.toUpperCase()}`, "success");
  }

  // --- CSV CUSTOM LOGIC --- //
  openCsvMapper() {
    this.els.exportMenu.classList.remove("show");
    if (this.columns.length === 0)
      return this.showToast("No hay datos cargados", "error");

    const selects = [
      this.els.mapLocalidad,
      this.els.mapScanCode,
      this.els.mapProducto,
      this.els.mapPedido,
      this.els.mapOrdenCompra
    ];

    selects.forEach((sel) => {
      sel.innerHTML = '<option value="">-- Seleccionar Columna --</option>';
      this.columns.forEach((col) => {
        const opt = document.createElement("option");
        opt.value = col;
        opt.innerText = col;
        sel.appendChild(opt);
      });
    });

    const autoSelect = (selectElement, keywords) => {
      const options = Array.from(selectElement.options);
      const found = options.find((opt) =>
        keywords.some((k) => opt.text.toLowerCase().includes(k))
      );
      if (found) selectElement.value = found.value;
    };

    autoSelect(this.els.mapLocalidad, [
      "localidad",
      "nom_tienda",
      "rs comprador",
      "local",
      "cliente",
      "ciudad",
      "sede"
    ]);
    autoSelect(this.els.mapScanCode, [
      "cód. empaque",
      "upc",
      "code",
      "codigo",
      "código",
      "ean",
      "sku"
    ]);
    autoSelect(this.els.mapProducto, [
      "producto",
      "descripcion",
      "descripción",
      "descripcion_larga",
      "sku_name",
      "item"
    ]);
    autoSelect(this.els.mapPedido, [
      "empaques pedidos",
      "pedido",
      "unidades",
      "cant",
      "solicitud"
    ]);
    autoSelect(this.els.mapOrdenCompra, [
      "orden",
      "num_oc",
      "compra",
      "oc",
      "po",
      "numero"
    ]);

    this.els.chkManualLocalidad.checked = false;
    this.toggleLocalidadInput();
    this.els.inputManualLocalidad.value = "";
    this.els.chkAutoOC.checked = false;
    this.toggleAutoOC();
    this.els.csvMapModal.classList.add("active");
  }

  toggleLocalidadInput() {
    if (this.els.chkManualLocalidad.checked) {
      this.els.mapLocalidad.classList.add("hidden");
      this.els.inputManualLocalidad.classList.remove("hidden");
    } else {
      this.els.mapLocalidad.classList.remove("hidden");
      this.els.inputManualLocalidad.classList.add("hidden");
    }
  }

  toggleAutoOC() {
    if (this.els.chkAutoOC.checked) {
      this.els.mapOrdenCompra.classList.add("hidden");
      this.els.previewAutoOC.classList.remove("hidden");
      const now = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      this.els.previewAutoOC.value = `${now.getFullYear()}${pad(
        now.getMonth() + 1
      )}${pad(now.getDate())}_${pad(now.getHours())}${pad(
        now.getMinutes()
      )}${pad(now.getSeconds())}`;
    } else {
      this.els.mapOrdenCompra.classList.remove("hidden");
      this.els.previewAutoOC.classList.add("hidden");
    }
  }

  generateCustomCSV() {
    const isAutoOC = this.els.chkAutoOC.checked;
    let autoOCValue = "";

    if (isAutoOC) {
      const now = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      autoOCValue = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(
        now.getDate()
      )}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(
        now.getSeconds()
      )}`;
    }

    const map = {
      locCol: this.els.mapLocalidad.value,
      scanCol: this.els.mapScanCode.value,
      prodCol: this.els.mapProducto.value,
      pedCol: this.els.mapPedido.value,
      ocCol: this.els.mapOrdenCompra.value,
      isManualLoc: this.els.chkManualLocalidad.checked,
      manualLocVal: this.els.inputManualLocalidad.value.trim().toUpperCase()
    };

    if (!map.isManualLoc && !map.locCol)
      return this.showToast("Falta definir columna Localidad", "error");
    if (map.isManualLoc && !map.manualLocVal)
      return this.showToast("Falta valor manual de Localidad", "error");
    if (!map.scanCol) return this.showToast("Falta columna Scan Code", "error");
    if (!map.prodCol) return this.showToast("Falta columna Producto", "error");
    if (!map.pedCol) return this.showToast("Falta columna Pedido", "error");
    if (!isAutoOC && !map.ocCol)
      return this.showToast("Falta columna Orden de Compra", "error");

    const csvRows = [];
    csvRows.push("LOCALIDAD,SCAN_COD,PRODUCTO X,PEDIDO,ORDEN DE COMPRA");

    const clean = (txt) => {
      if (txt === null || txt === undefined) return "";
      return String(txt)
        .replace(/,/g, " ")
        .replace(/[\r\n]+/g, " ")
        .trim();
    };

    this.visibleData.forEach((row) => {
      let valLoc = map.isManualLoc ? map.manualLocVal : row[map.locCol] || "";
      let valScan = row[map.scanCol] || "";
      let valProd = row[map.prodCol] || "";
      let valPed = row[map.pedCol] || "";
      let valOC = isAutoOC ? autoOCValue : row[map.ocCol] || "";
      csvRows.push(
        `${clean(valLoc)},${clean(valScan)},${clean(valProd)},${clean(
          valPed
        )},${clean(valOC)}`
      );
    });

    const csvContent = "\uFEFF" + csvRows.join("\r\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    const fname = this.els.reportTitle.value.trim() || "Reporte";
    link.setAttribute("href", url);
    link.setAttribute("download", `${fname}_custom.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    this.els.csvMapModal.classList.remove("active");
    this.showToast("CSV generado correctamente", "success");
  }

  escapeHTML(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  toggleMenu(id) {
    document.getElementById(id).classList.toggle("show");
  }

  toggleTheme() {
    const html = document.documentElement;
    const isDark = html.getAttribute("data-theme") === "dark";
    html.setAttribute("data-theme", isDark ? "light" : "dark");
    document.getElementById("themeIcon").className = isDark
      ? "ph ph-moon"
      : "ph ph-sun";
  }

  setLoading(v) {
    if (v) {
      this.els.loadingState.classList.remove("hidden");
      this.els.emptyState.classList.add("hidden");
      this.els.tableWrapper.classList.add("hidden");
    } else {
      this.els.loadingState.classList.add("hidden");
    }
  }

  getAlignClass(type) {
    if (["number", "currency", "integer", "percent"].includes(type))
      return "text-right";
    if (["date", "datetime", "time"].includes(type)) return "text-center";
    return "";
  }

  // --- MOTOR VISUAL CENTRALIZADO Y LIMPIO ---
  formatValue(val, config) {
    const type = config.type;
    const decimals = config.decimals !== undefined ? config.decimals : 2;
    const curr = config.currency || "PEN";

    if (val === null || val === undefined || val === "") return "";

    // 1. Textos y Enlaces (Links) - SIN BADGES
    if (
      type === "link" ||
      (typeof val === "string" && val.startsWith("http"))
    ) {
      return `<a href="${val}" target="_blank" style="color:var(--accent); font-weight:bold; text-decoration:underline;">${val}</a>`;
    }

    if (type === "text" || typeof val === "string") {
      return String(val);
    }

    // 2. Fechas y Horas (Con formato extendido)
    if (
      type === "date" ||
      type === "datetime" ||
      type === "time" ||
      val instanceof Date
    ) {
      try {
        const d = val instanceof Date ? val : new Date(val);
        if (isNaN(d.getTime())) return val;

        if (type === "time")
          return d.toLocaleTimeString("es-PE", {
            hour: "2-digit",
            minute: "2-digit"
          });

        const opts = {};
        const style = config.dateStyle || "short";

        if (style === "short") {
          opts.day = "2-digit";
          opts.month = "2-digit";
          opts.year = "numeric";
        } else if (style === "medium") {
          opts.day = "numeric";
          opts.month = "short";
          opts.year = "numeric";
        } else if (style === "long") {
          opts.day = "numeric";
          opts.month = "long";
          opts.year = "numeric";
        } else if (style === "full") {
          // Formato: "miércoles, 25 de febrero de 2026"
          opts.weekday = "long";
          opts.day = "numeric";
          opts.month = "long";
          opts.year = "numeric";
        }

        // Si es Fecha y Hora, se le agrega el tiempo a cualquier estilo
        if (type === "datetime") {
          opts.hour = "2-digit";
          opts.minute = "2-digit";
        }

        return d.toLocaleString("es-PE", opts);
      } catch (e) {
        return val;
      }
    }

    // 3. Números, Monedas y Porcentajes
    if (typeof val === "number") {
      if (type === "currency")
        return val.toLocaleString("es-PE", {
          style: "currency",
          currency: curr,
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        });

      if (type === "percent") {
        const pctVal = val > 1 ? val / 100 : val;
        return pctVal.toLocaleString("es-PE", {
          style: "percent",
          minimumFractionDigits: decimals,
          maximumFractionDigits: decimals
        });
      }

      if (type === "integer") return parseInt(val).toLocaleString("es-PE");

      return val.toLocaleString("es-PE", {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
      });
    }

    return val;
  }

  showToast(msg, type = "info") {
    const c = document.getElementById("toastContainer");
    const t = document.createElement("div");
    t.className = `toast toast-${type}`;
    const icon =
      type === "success"
        ? "ph-check-circle"
        : type === "error"
        ? "ph-warning-circle"
        : "ph-info";
    t.innerHTML = `<i class="ph ${icon}" style="font-size:20px; color:${
      type === "success" ? "var(--success)" : "var(--danger)"
    }"></i><span>${msg}</span>`;
    c.appendChild(t);
    setTimeout(() => {
      t.style.opacity = "0";
      t.addEventListener("transitionend", () => t.remove());
    }, 3000);
  }

  loadPreferences() {
    try {
      const stored = localStorage.getItem("dataViewerPrefs");
      return stored ? JSON.parse(stored) : null;
    } catch (e) {
      return null;
    }
  }

  savePreferences() {
    try {
      const prefs = { pageSize: this.pageSize, colSettings: {} };
      Object.keys(this.colSettings).forEach((col) => {
        prefs.colSettings[col] = {
          hidden: this.colSettings[col].hidden,
          type: this.colSettings[col].type,
          decimals: this.colSettings[col].decimals,
          currency: this.colSettings[col].currency || "PEN",
          dateStyle: this.colSettings[col].dateStyle,
          align: this.colSettings[col].align,
          textStyle: this.colSettings[col].textStyle
        };
      });
      localStorage.setItem("dataViewerPrefs", JSON.stringify(prefs));
    } catch (e) {}
  }

  resetPreferences() {
    if (confirm("¿Restaurar configuración visual de fábrica?")) {
      localStorage.removeItem("dataViewerPrefs");
      this.pageSize = 100;
      const pageSizeSelect = document.getElementById("pageSize");
      if (pageSizeSelect) pageSizeSelect.value = 100;

      if (this.rawData && this.rawData.length > 0) {
        this.initData(this.rawData, null);
      }
      this.showToast("Configuración restaurada", "success");
    }
  }
}

document.addEventListener("DOMContentLoaded", () => {
  try {
    window.app = new DataViewerApp();
  } catch (e) {
    alert("Error crítico iniciando la app: " + e.message);
  }
});
