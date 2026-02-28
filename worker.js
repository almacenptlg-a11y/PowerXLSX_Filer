importScripts("https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js");

let currentWorkbook = null;
let tempRawMatrix = [];

self.onmessage = async function(e) {
  const action = e.data.action;

  try {
      // 1. ANALIZAR UN ARCHIVO GIGANTE
      if (action === 'analyzeFile') {
          const file = e.data.file;
          self.postMessage({ type: 'progress', msg: 'Cargando archivo en memoria...' });
          
          const buffer = await file.arrayBuffer();
          currentWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true });
          
          const sheets = currentWorkbook.SheetNames;
          const author = currentWorkbook.Props && currentWorkbook.Props.Author ? currentWorkbook.Props.Author : "Usuario";
          
          self.postMessage({ type: 'fileAnalyzed', sheets: sheets, author: author, fileName: file.name });
      }
      
      // 2. EXTRAER VISTA PREVIA (Para el modal)
      else if (action === 'loadSheet') {
          const sheetName = e.data.sheetName;
          self.postMessage({ type: 'progress', msg: 'Extrayendo vista previa de: ' + sheetName + '...' });
          
          const sheet = currentWorkbook.Sheets[sheetName];
          tempRawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          
          const preview = tempRawMatrix.slice(0, 50);
          self.postMessage({ type: 'sheetLoaded', preview: preview, totalRows: tempRawMatrix.length });
      }

      // 3. PROCESAR EL ARCHIVO GIGANTE INDIVIDUAL (Medio millón de filas)
      else if (action === 'processSingle') {
          const headerIdx = e.data.headerIdx;
          const footerSkip = e.data.footerSkip;
          self.postMessage({ type: 'progress', msg: 'Ensamblando cientos de miles de filas...' });

          const headerRow = tempRawMatrix[headerIdx];
          if (!headerRow) throw new Error("Fila de encabezado inválida");

          const columns = [];
          headerRow.forEach((colName, idx) => {
            let safeName = colName !== undefined && colName !== null && String(colName).trim() !== "" ? String(colName).trim() : 'Columna_' + (idx + 1);
            if (columns.includes(safeName)) {
              let c = 1;
              while (columns.includes(safeName + '_' + c)) c++;
              safeName = safeName + '_' + c;
            }
            columns.push(safeName);
          });

          const startIndex = headerIdx + 1;
          const endIndex = tempRawMatrix.length - footerSkip;
          const jsonData = [];

          for (let i = startIndex; i < endIndex; i++) {
            const rowArr = tempRawMatrix[i];
            if (!rowArr || rowArr.length === 0) continue;
            const rowObj = {};
            let hasData = false;

            columns.forEach((colKey, colIdx) => {
              const cellVal = rowArr[colIdx];
              rowObj[colKey] = cellVal !== undefined ? cellVal : "";
              if (cellVal !== undefined && cellVal !== "" && cellVal !== null) hasData = true;
            });
            if (hasData) jsonData.push(rowObj);
          }

          tempRawMatrix = []; 
          currentWorkbook = null;
          self.postMessage({ type: 'singleDone', data: jsonData, columns: columns });
      }

      // 4. PROCESAR MÚLTIPLES ARCHIVOS
      else if (action === 'processMultiple') {
          const files = e.data.files;
          let combinedJson = [];
          let allColumns = new Set();
          let structureMismatch = false;
          let referenceHeaders = null;
          let filesProcessed = 0;

          for (let i = 0; i < files.length; i++) {
            const file = files[i];
            self.postMessage({ type: 'progress', msg: 'Procesando archivo ' + (i + 1) + ' de ' + files.length + '...\n' + file.name });

            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

            if (rawMatrix.length === 0) continue;

            const headerIdx = rawMatrix.findIndex(row => row && row.filter(c => c !== undefined && String(c).trim() !== "").length >= 1);
            if (headerIdx === -1) continue; 

            const headerRow = rawMatrix[headerIdx];
            const fileCols = [];

            headerRow.forEach((colName, idx) => {
               let safeName = (colName !== undefined && colName !== null && String(colName).trim() !== "") ? String(colName).trim() : 'Columna_' + (idx+1);
               if(fileCols.includes(safeName)) {
                  let c = 1;
                  while(fileCols.includes(safeName + '_' + c)) c++;
                  safeName = safeName + '_' + c;
               }
               fileCols.push(safeName);
               allColumns.add(safeName);
            });

            if (!referenceHeaders) {
                referenceHeaders = fileCols;
            } else if (referenceHeaders.join(',') !== fileCols.join(',')) {
                structureMismatch = true;
            }

            for (let r = headerIdx + 1; r < rawMatrix.length; r++) {
               const rowArr = rawMatrix[r];
               if (!rowArr || rowArr.length === 0 || rowArr.every(c => c === "" || c === undefined)) continue;

               const rowObj = {};
               let hasData = false;

               fileCols.forEach((colKey, colIdx) => {
                  const cellVal = rowArr[colIdx];
                  rowObj[colKey] = cellVal !== undefined ? cellVal : "";
                  if (cellVal !== undefined && cellVal !== "" && cellVal !== null) hasData = true;
               });

               if (hasData) {
                   rowObj['Archivo_Origen'] = file.name; 
                   combinedJson.push(rowObj);
               }
            }
            filesProcessed++;
          }

          if (combinedJson.length > 0) allColumns.add('Archivo_Origen');

          self.postMessage({
            type: 'multipleDone',
            data: combinedJson,
            columns: Array.from(allColumns),
            filesProcessed: filesProcessed,
            structureMismatch: structureMismatch
          });
      }
  } catch (err) {
      self.postMessage({ type: 'error', msg: err.message });
  }
};
