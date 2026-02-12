// CONFIGURACIÓN
const CONFIG = {
  // ID del archivo fuente (quita la parte https://docs.google.com/spreadsheets/d/ y /edit...)
  SOURCE_SPREADSHEET_ID: '1tYDuVZFmaZOxUDqUfEjgap4PWSlvYSYwnZnZNFW_CgM',
  
  // Nombres de las hojas que quieres copiar (pueden ser una o varias)
  SHEETS_TO_COPY: ['Leads', 'Tareas', 'Viajes', 'Agenda', 'Descartes', 'Comentarios', 'Aprobacion Leads'],
  
  // Intervalo de actualización en minutos
  UPDATE_INTERVAL_MINUTES: 30
};

// Función principal para copiar las hojas
function copySheetsFromSource() {
  try {
    // Abrir el archivo fuente
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Copiar cada hoja especificada
    CONFIG.SHEETS_TO_COPY.forEach(sheetName => {
      copySingleSheet(sourceSpreadsheet, targetSpreadsheet, sheetName);
    });
    
    // Registrar la última actualización
    updateLastSyncTime();
    
    console.log('✅ Hojas copiadas exitosamente: ' + CONFIG.SHEETS_TO_COPY.join(', '));
  } catch (error) {
    console.error('❌ Error al copiar hojas:', error.toString());
    // Puedes agregar notificaciones por email aquí si lo necesitas
  }
}

// Versión mejorada que preserva columnas adicionales
function copySingleSheet(sourceSpreadsheet, targetSpreadsheet, sheetName) {
  try {
    console.log(`🔄 Procesando "${sheetName}" (modo VALORES)...`);
    
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sourceSheet) {
      console.warn(`⚠️ "${sheetName}" no existe en fuente`);
      return;
    }
    
    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();
    
    if (lastRow === 0 || lastColumn === 0) {
      console.log(`ℹ️ "${sheetName}" está vacía`);
      return;
    }
    
    console.log(`📏 Tamaño: ${lastRow} filas x ${lastColumn} columnas`);
    
    // Crear hoja destino si no existe
    let targetSheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      targetSheet = targetSpreadsheet.insertSheet(sheetName);
    }
    
    // 🔴 IMPORTANTE: NO usar clear() en toda la hoja
    // En su lugar, limpiamos SOLO el área que vamos a actualizar
    const currentLastRow = targetSheet.getLastRow();
    const currentLastColumn = targetSheet.getLastColumn();
    
    // Limpiar solo el área donde pondremos los nuevos datos
    if (currentLastRow > 0 && currentLastColumn > 0) {
      // Limpiar contenido y formato del área fuente
      const cleanRange = targetSheet.getRange(1, 1, currentLastRow, currentLastColumn);
      cleanRange.clear(); // clear() en un rango específico, no en toda la hoja
    }
    
    // Obtener y escribir los nuevos valores
    const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
    const values = sourceRange.getValues();
    
    // Escribir SOLO en el área que corresponde a los datos fuente
    targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    
    // Copiar formato SOLO en el área actualizada
    copyFormattingWithValues(sourceSheet, targetSheet, values.length, values[0].length);
    
    console.log(`✅ "${sheetName}" copiada con VALORES calculados (columnas adicionales preservadas)`);
    
  } catch (error) {
    console.error(`❌ Error en "${sheetName}":`, error.toString());
  }
}

function copyFormattingWithValues(sourceSheet, targetSheet, rows, cols) {
  try {
    // Copiar ancho de columnas
    for (let col = 1; col <= cols; col++) {
      targetSheet.setColumnWidth(col, sourceSheet.getColumnWidth(col));
    }
    
    // Copiar formato básico (colores, alineación)
    const sourceRange = sourceSheet.getRange(1, 1, rows, cols);
    const targetRange = targetSheet.getRange(1, 1, rows, cols);
    
    // Copiar formatos
    targetRange.setBackgrounds(sourceRange.getBackgrounds());
    targetRange.setFontFamilies(sourceRange.getFontFamilies());
    targetRange.setFontSizes(sourceRange.getFontSizes());
    targetRange.setFontWeights(sourceRange.getFontWeights());
    targetRange.setFontStyles(sourceRange.getFontStyles());
    targetRange.setHorizontalAlignments(sourceRange.getHorizontalAlignments());
    targetRange.setVerticalAlignments(sourceRange.getVerticalAlignments());
    targetRange.setNumberFormats(sourceRange.getNumberFormats());
    targetRange.setWraps(sourceRange.getWraps());
    
    // Copiar bordes si es necesario
    targetRange.setBorder(true, true, true, true, true, true);
    
  } catch (error) {
    console.warn('⚠️ Error copiando formato:', error.toString());
  }
}

// Función para copiar formato básico (opcional)
function copyBasicFormatting(sourceSheet, targetSheet) {
  try {
    const lastRow = sourceSheet.getLastRow();
    const lastColumn = sourceSheet.getLastColumn();
    
    if (lastRow > 0 && lastColumn > 0) {
      // Copiar ancho de columnas
      for (let col = 1; col <= lastColumn; col++) {
        const width = sourceSheet.getColumnWidth(col);
        targetSheet.setColumnWidth(col, width);
      }
      
      // Copiar alto de filas (opcional)
      for (let row = 1; row <= lastRow; row++) {
        const height = sourceSheet.getRowHeight(row);
        targetSheet.setRowHeight(row, height);
      }
    }
  } catch (error) {
    console.warn('⚠️ No se pudo copiar el formato:', error.toString());
  }
}

// Función para registrar última actualización
function updateLastSyncTime() {
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = targetSpreadsheet.getSheetByName('Log_Sincronizacion');
  
  if (!logSheet) {
    logSheet = targetSpreadsheet.insertSheet('Log_Sincronizacion');
    logSheet.getRange('A1:B1').setValues([['Fecha/Hora', 'Estado']]);
  }
  
  const timestamp = new Date();
  logSheet.appendRow([timestamp, 'Sincronización exitosa']);
  
  // Mantener solo las últimas 100 entradas
  const lastRow = logSheet.getLastRow();
  if (lastRow > 100) {
    logSheet.deleteRows(2, lastRow - 100);
  }
}

// Configurar el trigger para ejecución automática
function setupTrigger() {
  // Eliminar triggers existentes para este script
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'copySheetsFromSource') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger
  ScriptApp.newTrigger('copySheetsFromSource')
    .timeBased()
    .everyMinutes(CONFIG.UPDATE_INTERVAL_MINUTES)
    .create();
  
  console.log(`✅ Trigger configurado para ejecutar cada ${CONFIG.UPDATE_INTERVAL_MINUTES} minutos`);
  
  // Ejecutar inmediatamente la primera vez
  copySheetsFromSource();
}

// Función para detener la actualización automática
function stopTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'copySheetsFromSource') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  console.log(`⏹️ ${removed} trigger(s) eliminado(s)`);
}

// Función para ver triggers activos
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    console.log('ℹ️ No hay triggers activos');
    return;
  }
  
  console.log('📋 Triggers activos:');
  triggers.forEach(trigger => {
    console.log(`- Función: ${trigger.getHandlerFunction()}, 
      Tipo: ${trigger.getEventType()},
      Fuente: ${trigger.getTriggerSource()}`);
  });
}