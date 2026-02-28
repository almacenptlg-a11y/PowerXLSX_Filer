// --- DETECTOR GLOBAL DE ERRORES --- //
window.onerror = function(msg, url, lineNo, columnNo, error) {
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

    // Control del Web Worker y Vistas Previas
    this.worker = null;
    this.tempPreviewMatrix = [];
    this.tempTotalRows = 0;
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
      selectedHeaderDisplay: document.getElementById("selectedHeaderIndexDisplay"),
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
    if(this.els.fileInput) {
      this.els.fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
          this.handleFiles(e.target.files);
        }
      });
      this.els.fileInput.addEventListener('click', (e) => {
        e.target.value = null; 
      });
    }

    const btnReset = document.getElementById("btnResetPrefs");
    if(btnReset) btnReset.addEventListener("click", () => this.resetPreferences());

    window.addEventListener('dragover', (e) => { e.preventDefault(); this.els.dragOverlay.classList.add('active'); });
    window.addEventListener('dragleave', (e) => { if (e.target === this.els.dragOverlay) this.els.dragOverlay.classList.remove('active'); });
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      this.els.dragOverlay.classList.remove('active');
      if (e.dataTransfer.files.length) this.handleFiles(e.dataTransfer.files);
    });

    document.getElementById("btnPrev").addEventListener("click", () => this.changePage(-1));
    document.getElementById("btnNext").addEventListener("click", () => this.changePage(1));
    document.getElementById("pageSize").addEventListener("change", (e) => {
      this.pageSize = parseInt(e.target.value);
      this.currentPage = 1;
      this.savePreferences();
      this.render();
    });

    this.els.globalSearch.addEventListener("input", (e) => {
      this.searchQuery = e.target.value.toLowerCase();
      this.currentPage = 1;
      this.processData();
    });

    document.getElementById("btnCloseSheetModal").addEventListener("click", () => {
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

    document.addEventListener("click", (e) => {
      if (!e.target.closest("#btnColumns") && !e.target.closest("#colMenu"))
        this.els.colMenu.classList.remove("show");
      if (!e.target.closest("#btnExport") && !e.target.closest("#exportMenu"))
        this.els.exportMenu.classList.remove("show");
      if (!e.target.closest("#columnContextMenu") && !e.target.closest(".btn-col-menu"))
        this.els.ctxMenu.classList.remove("show");
    });

    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return; 
        e.preventDefault(); 
        this.undo();        
      }
    });
  }

  // --- MOTOR WEB WORKER CENTRALIZADO --- //
  initWorker() {
    if (this.worker) this.worker.terminate();
    this.worker = new Worker('worker.js');
    
    this.worker.onmessage = (e) => {
      const res = e.data;

      // Actualizar mensajes en pantalla
      if (res.type === 'progress') {
        const loadingText = this.els.loadingState.querySelector("p");
        if(loadingText) loadingText.innerText = res.msg;
      } 
      // Errores
      else if (res.type === 'error') {
        this.showToast(res.msg, 'error');
        this.setLoading(false);
      }
      // Archivo único analizado (Saber hojas y nombre)
      else if (res.type === 'fileAnalyzed') {
        this.els.reportTitle.value = res.fileName.replace(/\.[^/.]+$/, "");
        this.els.reportAuthor.value = res.author;
        
        if (res.sheets.length > 1) {
          this.showSheetSelection(res.sheets);
          this.setLoading(false);
        } else {
          this.worker.postMessage({ action: 'loadSheet', sheetName: res.sheets[0] });
        }
      }
      // Vista previa de hoja cargada
      else if (res.type === 'sheetLoaded') {
        this.tempPreviewMatrix = res.preview;
        this.tempTotalRows = res.totalRows;
        
        const likelyHeader = this.tempPreviewMatrix.findIndex((row) => row && row.filter((c) => c).length > 1);
        this.tempHeaderIdx = likelyHeader >= 0 ? likelyHeader : 0;
        this.els.footerSkipCount.value = 0;
        
        this.setLoading(false);
        this.els.sheetModal.classList.remove("active");
        this.els.structureModal.classList.add("active");
        this.renderPreviewTableRows();
      }
      // Archivo único procesado completamente (500k filas)
      else if (res.type === 'singleDone') {
        this.els.structureModal.classList.remove("active");
        const loadingText = this.els.loadingState.querySelector("p");
        if(loadingText) loadingText.innerText = "Construyendo tabla interactiva...";
        
        if (res.data.length === 0) {
          this.showToast("No se encontraron datos.", "error");
          this.setLoading(false);
          return;
        }

        this.initData(res.data, res.columns);
        this.setLoading(false);
        this.showToast(`¡Estructura aplicada! ${res.data.length.toLocaleString()} filas cargadas ultra rápido.`, "success");
      }
      // Archivos múltiples combinados
      else if (res.type === 'multipleDone') {
        if (res.data.length === 0) {
          this.showToast("No se encontraron datos válidos.", "error");
          this.setLoading(false);
          return;
        }
        this.els.reportTitle.value = "Reporte_Combinado";
        const loadingText = this.els.loadingState.querySelector("p");
        if(loadingText) loadingText.innerText = "Construyendo tabla interactiva...";
        
        this.initData(res.data, res.columns);
        this.setLoading(false);

        if (res.structureMismatch) {
            this.showToast(`Carga masiva: ${res.data.length.toLocaleString()} filas. NOTA: Las columnas variaban.`, 'warning');
        } else {
            this.showToast(`¡Completado! ${res.data.length.toLocaleString()} filas combinadas de ${res.filesProcessed} archivos.`, 'success');
        }
      }
    };
    
    this.worker.onerror = (err) => {
      console.error("Worker error:", err);
      this.showToast("Error en procesamiento de hardware (Worker).", "error");
      this.setLoading(false);
    };
  }

  async handleFiles(fileList) {
    try {
      if (!fileList || fileList.length === 0) return;
      const files = Array.from(fileList);
      
      this.resetState();
      this.initWorker(); // Iniciamos un trabajador limpio y fresco

      this.setLoading(true);

      if (files.length === 1) {
        this.worker.postMessage({ action: 'analyzeFile', file: files[0] });
      } else {
        this.worker.postMessage({ action: 'processMultiple', files: files });
      }
    } catch (err) {
      alert("Error en handleFiles: " + err.message);
    }
  }

  resetState() {
    this.rawData = [];
    this.visibleData = [];
    this.columns = [];
    this.sortCol = null;
    this.searchQuery = "";
    this.undoStack = [];
    this.tempPreviewMatrix = [];
    this.tempTotalRows = 0;
    
    if(this.els.globalSearch) this.els.globalSearch.value = "";
    if(this.els.tableWrapper) this.els.tableWrapper.classList.add("hidden");
    if(this.els.footer) this.els.footer.classList.add("hidden");
    if(this.filterSummary) this.filterSummary.classList.add("hidden");
    if(this.els.emptyState) this.els.emptyState.classList.remove("hidden");
    if(this.els.thead) this.els.thead.innerHTML = "";
    if(this.els.tbody) this.els.tbody.innerHTML = "";
    if(this.els.tfoot) this.els.tfoot.innerHTML = "";
  }

  showSheetSelection(sheets) {
    const list = this.els.sheetList;
    list.innerHTML = "";
    sheets.forEach((sheet) => {
      const btn = document.createElement("div");
      btn.className = "sheet-btn";
      btn.innerHTML = `<span style="font-weight:600">${sheet}</span> <i class="ph ph-caret-right"></i>`;
      btn.onclick = () => {
        this.els.sheetModal.classList.remove("active");
        this.setLoading(true);
        this.worker.postMessage({ action: 'loadSheet', sheetName: sheet });
      };
      list.appendChild(btn);
    });
    this.els.sheetModal.classList.add("active");
  }

  renderPreviewTableRows() {
    const table = this.els.previewTable;
    table.innerHTML = "";
    const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;
    const totalRows = this.tempTotalRows;

    this.els.selectedHeaderDisplay.innerText = `Fila ${this.tempHeaderIdx + 1}`;
    const limit = Math.min(this.tempPreviewMatrix.length, 50);

    for (let i = 0; i < limit; i++) {
      this.buildPreviewRow(table, i, totalRows, footerSkip);
    }

    if (totalRows > limit) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="100" style="text-align:center; padding:8px; font-weight:bold; color:var(--accent); background: rgba(14, 165, 233, 0.05)">... y ${(totalRows - limit).toLocaleString()} filas más se procesarán a máxima velocidad ...</td>`;
      table.appendChild(tr);
    }
  }

  buildPreviewRow(table, index, totalRows, footerSkip) {
    const rowData = this.tempPreviewMatrix[index];
    if (!rowData) return;

    const tr = document.createElement("tr");
    const isHeader = index === this.tempHeaderIdx;
    const isIgnoredTop = index < this.tempHeaderIdx;

    if (isHeader) tr.className = "preview-header";
    else if (isIgnoredTop) tr.className = "preview-ignored";

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
    
    const footerSkip = parseInt(this.els.footerSkipCount.value) || 0;
    
    // Le decimos al worker que procese todo
    this.worker.postMessage({ 
        action: 'processSingle', 
        headerIdx: this.tempHeaderIdx, 
        footerSkip: footerSkip 
    });
  }

  initData(data, customColumns = null) {
    this.rawData = data;
    this.columns = customColumns || Object.keys(data[0]);

    const prefs = this.loadPreferences();
    if (prefs && prefs.pageSize) {
        this.pageSize = prefs.pageSize;
        const pageSizeSelect = document.getElementById('pageSize');
        if(pageSizeSelect) pageSizeSelect.value = this.pageSize;
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
        this.colSettings[col].hidden = saved.hidden !== undefined ? saved.hidden : this.colSettings[col].hidden;
        this.colSettings[col].type = saved.type || this.colSettings[col].type;
        this.colSettings[col].decimals = saved.decimals !== undefined ? saved.decimals : this.colSettings[col].decimals;
        this.colSettings[col].currency = saved.currency || this.colSettings[col].currency;
        this.colSettings[col].dateStyle = saved.dateStyle || this.colSettings[col].dateStyle;
        this.colSettings[col].align = saved.align || this.colSettings[col].align;
        this.colSettings[col].textStyle = saved.textStyle || this.colSettings[col].textStyle;
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
      if (lower.match(/(precio|costo|total|valor|importe|venta|compra)/)) return "currency";
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
        let va = a[this.sortCol], vb = b[this.sortCol];
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
    const activeCols = this.columns.filter((c) => this.colSettings[c].activeFilters !== null);

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
      const iconClass = isSorted ? (this.sortAsc ? "ph-arrow-up" : "ph-arrow-down") : "";
      const safeCol = String(col).replace(/'/g, "\\'");

      th.innerHTML = `
        <div class="th-content ${alignClass}">
          <div class="btn-col-menu ${hasFilter ? "active" : ""}" onclick="app.openColumnMenu(event, '${safeCol}')">
             <i class="ph ${hasFilter ? "ph-funnel ph-fill" : "ph-dots-three-vertical"}"></i>
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
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'A') return;
        const currentlySelected = this.els.tbody.querySelector('.row-selected');
        if (currentlySelected && currentlySelected !== tr) currentlySelected.classList.remove('row-selected');
        tr.classList.toggle("row-selected");
      });

      this.columns.forEach((col) => {
        if (this.colSettings[col].hidden) return;
        const td = document.createElement("td");
        const config = this.colSettings[col];
        
        td.className = this.getAlignClass(config.type);
        td.innerHTML = this.formatValue(row[col], config);
        
        if (config.align && config.align !== "auto") td.style.textAlign = config.align;
        if (config.textStyle && config.textStyle !== "none") td.style.textTransform = config.textStyle;

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
        const sum = this.visibleData.reduce((acc, r) => acc + (parseFloat(r[col]) || 0), 0);
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
    else { this.sortCol = col; this.sortAsc = true; }
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
    document.getElementById("statusMsg").innerText = `${total.toLocaleString()} registros encontrados`;
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
    const uniqueVals = [...new Set(relevantRows.map((r) => String(r[col])))].sort();

    let extraControls = "";
    if (["number", "currency", "percent"].includes(settings.type)) {
      extraControls += `<div style="margin-top:8px; display:flex; align-items:center; justify-content:space-between;"><label class="col-menu-label" style="margin:0">Decimales</label><input type="number" min="0" max="6" class="form-input form-input-sm" style="width:60px" value="${settings.decimals}" onchange="app.changeColDecimal('${col}', this.value)"></div>`;
    }

    if (settings.type === "currency") {
      extraControls += `<div style="margin-top:8px"><label class="col-menu-label" style="margin-bottom:2px">Simbolo</label><select class="form-select" onchange="app.changeColCurrency('${col}', this.value)"><option value="PEN" ${settings.currency === "PEN" ? "selected" : ""}>S/ (PEN)</option><option value="USD" ${settings.currency === "USD" ? "selected" : ""}>$ (USD)</option><option value="EUR" ${settings.currency === "EUR" ? "selected" : ""}>€ (EUR)</option></select></div>`;
    }

    if (["date", "datetime"].includes(settings.type)) {
      extraControls += `<div style="margin-top:8px"><label class="col-menu-label" style="margin-bottom:2px">Estilo</label><select class="form-select" onchange="app.changeColDateStyle('${col}', this.value)">
        <option value="short" ${settings.dateStyle === "short" ? "selected" : ""}>Corto (DD/MM/YYYY)</option>
        <option value="medium" ${settings.dateStyle === "medium" ? "selected" : ""}>Medio (04 ene 2026)</option>
        <option value="long" ${settings.dateStyle === "long" ? "selected" : ""}>Largo (4 de enero...)</option>
        <option value="full" ${settings.dateStyle === "full" ? "selected" : ""}>Texto (Miércoles...)</option>
        </select></div>`;
    }

    extraControls += `
      <div style="margin-top:8px; padding-top:8px; border-top:1px dashed var(--border);">
        <label class="col-menu-label" style="margin-bottom:2px">Alineación</label>
        <select class="form-select" onchange="app.changeColAlign('${col}', this.value)">
           <option value="auto" ${settings.align === "auto" || !settings.align ? "selected" : ""}>Automática</option>
           <option value="left" ${settings.align === "left" ? "selected" : ""}>Izquierda</option>
           <option value="center" ${settings.align === "center" ? "selected" : ""}>Centro</option>
           <option value="right" ${settings.align === "right" ? "selected" : ""}>Derecha</option>
        </select>
      </div>
      <div style="margin-top:8px">
        <label class="col-menu-label" style="margin-bottom:2px">Mayús / Minús</label>
        <select class="form-select" onchange="app.changeColTextStyle('${col}', this.value)">
           <option value="none" ${settings.textStyle === "none" || !settings.textStyle ? "selected" : ""}>Normal</option>
           <option value="uppercase" ${settings.textStyle === "uppercase" ? "selected" : ""}>MAYÚSCULAS</option>
           <option value="lowercase" ${settings.textStyle === "lowercase" ? "selected" : ""}>minúsculas</option>
           <option value="capitalize" ${settings.textStyle === "capitalize" ? "selected" : ""}>Capitalizar</option>
        </select>
      </div>
    `;

    container.innerHTML = `
        <div class="col-menu-section">
          <label class="col-menu-label">Formato</label>
          <select class="form-select" onchange="app.changeColFormat('${col}', this.value)">
            <option value="auto" ${settings.type === "auto" ? "selected" : ""}>Automático</option>
            <option value="text" ${settings.type === "text" ? "selected" : ""}>Texto</option>
            <option value="number" ${settings.type === "number" ? "selected" : ""}>Número</option>
            <option value="integer" ${settings.type === "integer" ? "selected" : ""}>Entero</option>
            <option value="currency" ${settings.type === "currency" ? "selected" : ""}>Moneda</option>
            <option value="percent" ${settings.type === "percent" ? "selected" : ""}>Porcentaje (%)</option>
            <option value="date" ${settings.type === "date" ? "selected" : ""}>Fecha</option>
            <option value="datetime" ${settings.type === "datetime" ? "selected" : ""}>Fecha y Hora</option>
            <option value="time" ${settings.type === "time" ? "selected" : ""}>Hora</option>
            <option value="link" ${settings.type === "link" ? "selected" : ""}>Enlace (URL)</option>
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
    allDiv.innerHTML = `<input type="checkbox" id="chkAllFilters" ${settings.activeFilters === null ? "checked" : ""}> <span>(Seleccionar Todo)</span>`;
    allDiv.onclick = (ev) => {
      if (ev.target.tagName !== "INPUT") { const chk = allDiv.querySelector("input"); chk.checked = !chk.checked; }
      const state = allDiv.querySelector("input").checked;
      filterContainer.querySelectorAll(".val-chk").forEach((c) => (c.checked = state));
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
