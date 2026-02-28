// Importamos la librería de Excel directamente dentro del trabajador en segundo plano
importScripts("https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js");

self.onmessage = async function(e) {
  const files = e.data.files;
  let combinedJson = [];
  let allColumns = new Set();
  let structureMismatch = false;
  let referenceHeaders = null;
  let filesProcessed = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    
    // Avisamos a la pantalla frontal cómo vamos
    self.postMessage({ type: 'progress', msg: `Procesando archivo ${i + 1} de ${files.length}...\n${file.name}` });

    try {
      // Leemos el archivo en crudo (ArrayBuffer es ultra rápido)
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
         let safeName = (colName !== undefined && colName !== null && String(colName).trim() !== "")
                        ? String(colName).trim() : `Columna_${idx+1}`;
         if(fileCols.includes(safeName)) {
            let c = 1;
            while(fileCols.includes(`${safeName}_${c}`)) c++;
            safeName = `${safeName}_${c}`;
         }
         fileCols.push(safeName);
         allColumns.add(safeName);
      });

      if (!referenceHeaders) {
          referenceHeaders = fileCols;
      } else if (referenceHeaders.join(',') !== fileCols.join(',')) {
          structureMismatch = true;
      }

      // Procesamiento al 100% de la CPU (sin pausas artificiales)
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
    } catch (err) {
      self.postMessage({ type: 'error', msg: `Error en ${file.name}: ${err.message}` });
    }
  }

  if (combinedJson.length > 0) {
      allColumns.add('Archivo_Origen');
  }

  // Devolvemos el paquete de datos procesados al hilo principal
  self.postMessage({
    type: 'done',
    data: combinedJson,
    columns: Array.from(allColumns),
    filesProcessed: filesProcessed,
    structureMismatch: structureMismatch
  });
};
