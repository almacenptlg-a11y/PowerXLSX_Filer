// Importamos SheetJS (Para Excel) y PapaParse (Para CSV ultrarrápido)
importScripts("https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js");
importScripts("https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js");

let currentFile = null;
let currentWorkbook = null;
let tempRawMatrix = [];
let isCurrentCSV = false;

self.onmessage = async function(e) {
  const action = e.data.action;

  try {
      // 1. ANALIZAR EL ARCHIVO
      if (action === 'analyzeFile') {
          const file = e.data.file;
          currentFile = file;
          isCurrentCSV = file.name.toLowerCase().endsWith('.csv');
          
          self.postMessage({ type: 'progress', msg: 'Analizando formato del archivo...' });

          if (isCurrentCSV) {
              // Si es CSV, no hay "hojas", creamos una virtual al instante
              self.postMessage({ type: 'fileAnalyzed', sheets: ["Datos_CSV"], author: "Usuario", fileName: file.name });
          } else {
              // Si es Excel, usamos SheetJS
              const buffer = await file.arrayBuffer();
              currentWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true });
              const sheets = currentWorkbook.SheetNames;
              const author = currentWorkbook.Props && currentWorkbook.Props.Author ? currentWorkbook.Props.Author : "Usuario";
              self.postMessage({ type: 'fileAnalyzed', sheets: sheets, author: author, fileName: file.name });
          }
      }
      
      // 2. EXTRAER VISTA PREVIA (Para el modal)
      else if (action === 'loadSheet') {
          self.postMessage({ type: 'progress', msg: 'Extrayendo vista previa...' });
          
          if (isCurrentCSV) {
              // PapaParse extrae las primeras 50 filas del CSV en milisegundos
              Papa.parse(currentFile, {
                  preview: 50,
                  skipEmptyLines: true,
                  complete: function(results) {
                      tempRawMatrix = results.data;
                      // Mandamos 999999 para que la pantalla sepa que hay muchos más datos
                      self.postMessage({ type: 'sheetLoaded', preview: tempRawMatrix, totalRows: 999999 });
                  }
              });
          } else {
              // Vista previa de Excel
              const sheetName = e.data.sheetName;
              const sheet = currentWorkbook.Sheets[sheetName];
              tempRawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
              const preview = tempRawMatrix.slice(0, 50);
              self.postMessage({ type: 'sheetLoaded', preview: preview, totalRows: tempRawMatrix.length });
          }
      }

      // 3. PROCESAR UN SOLO ARCHIVO MASIVO
      else if (action === 'processSingle') {
          const headerIdx = e.data.headerIdx;
          const footerSkip = e.data.footerSkip;
          self.postMessage({ type: 'progress', msg: 'Ensamblando cientos de miles de filas...' });

          if (isCurrentCSV) {
              // PROCESAMIENTO STREAMING PARA CSV (0% colapso de RAM)
              let rowCount = 0;
              let headerRow = null;
              let columns = [];
              let jsonData = [];

              Papa.parse(currentFile, {
                  skipEmptyLines: true,
                  step: function(results) {
                      const rowArr = results.data;
                      
                      if (rowCount === headerIdx) {
                          headerRow = rowArr;
                          headerRow.forEach((colName, idx) => {
                              let safeName = colName !== undefined && colName !== null && String(colName).trim() !== "" ? String(colName).trim() : 'Columna_' + (idx + 1);
                              if (columns.includes(safeName)) {
                                  let c = 1;
                                  while (columns.includes(safeName + '_' + c)) c++;
                                  safeName = safeName + '_' + c;
                              }
                              columns.push(safeName);
                          });
                      } else if (rowCount > headerIdx) {
                          const rowObj = {};
                          let hasData = false;
                          columns.forEach((colKey, colIdx) => {
                              const cellVal = rowArr[colIdx];
                              rowObj[colKey] = cellVal !== undefined ? cellVal : "";
                              if (cellVal !== undefined && cellVal !== "" && cellVal !== null) hasData = true;
                          });
                          if (hasData) jsonData.push(rowObj);
                      }
                      rowCount++;
                  },
                  complete: function() {
                      if (footerSkip > 0 && jsonData.length > footerSkip) {
                          jsonData = jsonData.slice(0, jsonData.length - footerSkip);
                      }
                      currentFile = null;
                      self.postMessage({ type: 'singleDone', data: jsonData, columns: columns });
                  },
                  error: function(err) {
                      self.postMessage({ type: 'error', msg: "Error al procesar el CSV: " + err.message });
                  }
              });

          } else {
              // PROCESAMIENTO EXCEL (El que ya conoces, de 40 segs)
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
      }

      // 4. PROCESAR MÚLTIPLES ARCHIVOS COMBINADOS
      else if (action === 'processMultiple') {
          const files = e.data.files;
          let combinedJson = [];
          let allColumns = new Set();
          let structureMismatch = false;
          let referenceHeaders = null;
          let filesProcessed = 0;

          // Función Promesa para extraer CSV limpiamente dentro de un loop
          const parseCSVAsync = (f) => {
              return new Promise((resolve, reject) => {
                  let headerRow = null;
                  let columns = [];
                  let fileData = [];
                  Papa.parse(f, {
                      skipEmptyLines: true,
                      step: function(results) {
                          const rowArr = results.data;
                          if (!headerRow) {
                              headerRow = rowArr;
                              headerRow.forEach((colName, idx) => {
                                  let safeName = colName !== undefined && colName !== null && String(colName).trim() !== "" ? String(colName).trim() : 'Columna_' + (idx + 1);
                                  if (columns.includes(safeName)) {
                                      let c = 1;
                                      while (columns.includes(safeName + '_' + c)) c++;
                                      safeName = safeName + '_' + c;
                                  }
                                  columns.push(safeName);
                              });
                          } else {
                              const rowObj = {};
                              let hasData = false;
                              columns.forEach((colKey, colIdx) => {
                                  const cellVal = rowArr[colIdx];
                                  rowObj[colKey] = cellVal !== undefined ? cellVal : "";
                                  if (cellVal !== undefined && cellVal !== "" && cellVal !== null) hasData = true;
                              });
                              if (hasData) {
                                  rowObj['Archivo_Origen'] = f.name;
                                  fileData.push(rowObj);
                              }
                          }
                      },
                      complete: function() { resolve({ columns, fileData }); },
                      error: reject
                  });
              });
          };

          for (let i = 0; i < files.length; i++) {
            const file = files[i];
            self.postMessage({ type: 'progress', msg: 'Procesando archivo ' + (i + 1) + ' de ' + files.length + '...\n' + file.name });
            const isCSVFile = file.name.toLowerCase().endsWith('.csv');

            let fileCols = [];
            let fileData = [];

            if (isCSVFile) {
                // Lectura Streaming de CSV
                const csvResult = await parseCSVAsync(file);
                fileCols = csvResult.columns;
                fileData = csvResult.fileData;
                fileCols.forEach(c => allColumns.add(c));
                combinedJson = combinedJson.concat(fileData);
            } else {
                // Lectura de Excel
                const buffer = await file.arrayBuffer();
                const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                if (rawMatrix.length === 0) continue;

                const headerIdx = rawMatrix.findIndex(row => row && row.filter(c => c !== undefined && String(c).trim() !== "").length >= 1);
                if (headerIdx === -1) continue; 
                const headerRow = rawMatrix[headerIdx];

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
            }

            if (!referenceHeaders) {
                referenceHeaders = fileCols;
            } else if (referenceHeaders.join(',') !== fileCols.join(',')) {
                structureMismatch = true;
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
