// ============================================
// MÓDULO: USUARIOS
// ============================================

var Usuarios = (function () {

  /**
   * Obtiene todos los usuarios de la hoja 'Usuarios'
   * @returns {Array} Array de objetos {id, nombre}
   */
  function getUsuarios() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Usuarios');

      if (!sheet) throw new Error('No se encontró la hoja "Usuarios"');

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];

      // Obtener Mail (Col A) y Nombre (Col B)
      const dataRange = sheet.getRange(2, 1, lastRow - 1, 2);
      const data = dataRange.getValues();

      const usuarios = data
        .filter(row => row[1] && row[1].toString().trim() !== '')
        .map(row => ({
          id: row[0] ? row[0].toString().trim() : '',
          nombre: row[1].toString().trim()
        }))
        .filter(usuario => usuario.id && usuario.nombre);

      console.log(`✅ Usuarios.getUsuarios: ${usuarios.length} usuarios`);
      return usuarios;

    } catch (error) {
      console.error('❌ Error en Usuarios.getUsuarios:', error);
      throw error;
    }
  }

  /**
   * Obtiene el nombre de un usuario por su ID
   * @param {string} usuarioId - Email del usuario
   * @returns {string} Nombre del usuario o string por defecto
   */
  function getNombrePorId(usuarioId) {
    try {
      const usuarios = getUsuarios();
      const usuario = usuarios.find(u => u.id.toLowerCase() === usuarioId.toLowerCase());
      return usuario ? usuario.nombre : `Usuario ${usuarioId}`;
    } catch (error) {
      console.error('❌ Error en Usuarios.getNombrePorId:', error);
      return `Usuario ${usuarioId}`;
    }
  }

  /**
   * Obtiene un usuario por su ID
   * @param {string} usuarioId - Email del usuario
   * @returns {Object|null} Objeto usuario o null
   */
  function getUsuarioPorId(usuarioId) {
    try {
      const usuarios = getUsuarios();
      return usuarios.find(u => u.id.toLowerCase() === usuarioId.toLowerCase()) || null;
    } catch (error) {
      console.error('❌ Error en Usuarios.getUsuarioPorId:', error);
      return null;
    }
  }

  /**
   * Obtiene estadísticas básicas de usuarios
   * @returns {Object} Estadísticas de usuarios
   */
  function getEstadisticas() {
    try {
      const usuarios = getUsuarios();
      return {
        total: usuarios.length,
        activos: usuarios.length,
        lista: usuarios
      };
    } catch (error) {
      console.error('❌ Error en Usuarios.getEstadisticas:', error);
      return { total: 0, activos: 0, lista: [] };
    }
  }

  // API Pública
  return {
    getUsuarios: getUsuarios,
    getNombrePorId: getNombrePorId,
    getUsuarioPorId: getUsuarioPorId,
    getEstadisticas: getEstadisticas
  };

})();

// ================================
// Extensiones: Última actividad
// ================================

/**
 * Busca la última fecha de actividad asociada a un `usuarioId` en hojas relevantes.
 * Revisa `Leads` (columna `COL_ULTIMA_GESTION`) y las hojas: `Comentarios`, `Agenda`, `Tareas`.
 * Devuelve una fecha (Date) o null.
 */
/**
 * Busca la última fecha de actividad asociada a un `usuarioId`.
 * PRIORIZA lectura de índice `Ultima_Actividad` para velocidad.
 * Luego revisa `Leads` (columna `COL_ULTIMA_GESTION`) como fallback rápido.
 * Devuelve una fecha formateada (string) o null.
 */
/**
 * Busca la última fecha de actividad asociada a un `usuarioId` en columnas específicas.
 * Lógica solicitada:
 * - Leads: Usuario en Col AE (31), Fecha en Col B (2)
 * - Tareas: Usuario en Col I (9), Fecha en Col F (6)
 * - Agenda: Usuario en Col I (9), Fecha en Col J (10)
 */
function getUltimaActividadAsociado(usuarioId) {
  try {
    if (!usuarioId) return null;
    var usuarioIdStr = usuarioId.toString().trim().toLowerCase();
    var maxDate = null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Helper para escanear columnas específicas
    function checkSheetCols(sheetName, userColIdx, dateColIdx) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      // Leer solo las dos columnas necesarias (Rápido)
      try {
        // getRange(row, col, numRows, numCols)
        // userColIdx y dateColIdx son 1-based (como pide Apps Script)
        var userVals = sheet.getRange(2, userColIdx, lastRow - 1, 1).getValues();
        var dateVals = sheet.getRange(2, dateColIdx, lastRow - 1, 1).getValues();

        for (var i = 0; i < userVals.length; i++) {
          var u = userVals[i][0];
          if (u && u.toString().trim().toLowerCase() === usuarioIdStr) {
            var d = parsePossiblySheetDate(dateVals[i][0]);
            if (d && (!maxDate || d.getTime() > maxDate.getTime())) {
              maxDate = d;
            }
          }
        }
      } catch (e) {
        console.warn('Error leyendo columnas en ' + sheetName, e);
      }
    }

    // 1. Leads (AE=31, B=2)
    checkSheetCols('Leads', 31, 2);

    // 2. Tareas (I=9, F=6)
    checkSheetCols('Tareas', 9, 6);

    // 3. Agenda (I=9, J=10)
    checkSheetCols('Agenda', 9, 10);

    // 4. Comentarios (G=7, E=5)
    checkSheetCols('Comentarios', 7, 5);

    // Formatear resultado
    if (maxDate) return Fechas && Fechas.formatear ? Fechas.formatear(maxDate) : maxDate.toLocaleString();
    return "Sin actividad reciente";

  } catch (error) {
    console.error('❌ Error en getUltimaActividadAsociado:', error);
    return null;
  }
}

/**
 * Intenta convertir distintos formatos de fecha (Date, número de Sheets, string) a Date.
 */
function parsePossiblySheetDate(value) {
  try {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
      // Número típico de Google Sheets (días desde 1899-12-30 / 1900)
      if (value > 25000) {
        return new Date((value - 25569) * 86400 * 1000);
      }
      return null;
    }
    // Intentar parsear string
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Reconstruye la hoja `Ultima_Actividad` con la última fecha por LeadID.
 * Escanea Tareas, Agenda, Comentarios usando columnas fijas.
 */
function buildUltimaActividadIndex() {
  try {
    var start = new Date().getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Función de ayuda para obtener hoja ignorando espacios
    function getSheetFlexible(name) {
      var s = ss.getSheetByName(name);
      if (s) return s;
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim() === name.trim()) return sheets[i];
      }
      return null;
    }

    var sheetName = 'Ultima_Actividad';

    // Mapa temporal
    var index = {};

    /** Normalización interna */
    function normalizarId(val) {
      if (val === null || val === undefined) return "";
      var sid = val.toString().trim().replace(/\.0$/, "");
      return sid;
    }

    // Configuración estricta de columnas (LeadID Col, Date Col, Actor Col)
    // Tareas: LeadID B(2), Date F(6), Actor I(9)
    // Agenda: LeadID B(2), Date J(10), Actor I(9)
    // Comentarios: LeadID B(2), Date E(5), Actor G(7)
    var sheetsConfig = [
      { name: 'Tareas', leadCol: 2, dateCol: 6, actorCol: 9 },
      { name: 'Agenda', leadCol: 2, dateCol: 10, actorCol: 9 },
      { name: 'Comentarios', leadCol: 2, dateCol: 5, actorCol: 7 }
    ];

    sheetsConfig.forEach(function (conf) {
      var sheet = getSheetFlexible(conf.name);
      if (!sheet) return;

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      // Leer columnas LeadID, Fecha y Actor
      try {
        var leadVals = sheet.getRange(2, conf.leadCol, lastRow - 1, 1).getValues();
        var dateVals = sheet.getRange(2, conf.dateCol, lastRow - 1, 1).getValues();
        var actorVals = sheet.getRange(2, conf.actorCol, lastRow - 1, 1).getValues();

        for (var i = 0; i < leadVals.length; i++) {
          var leadVal = leadVals[i][0];
          if (!leadVal) continue;
          var leadId = normalizarId(leadVal);
          if (!leadId) continue;

          var dateVal = dateVals[i][0];
          var parsed = parsePossiblySheetDate(dateVal);
          if (!parsed) continue;

          var actorVal = actorVals[i][0] ? actorVals[i][0].toString().trim() : 'N/A';

          var existing = index[leadId];
          if (!existing || parsed.getTime() > new Date(existing.ultimaISO).getTime()) {
            index[leadId] = { ultimaISO: parsed.toISOString(), origen: conf.name, actor: actorVal, fila: i + 2 };
          }
        }
      } catch (e) {
        console.warn('Error escaneando ' + conf.name + ' para índice:', e);
      }
    });

    // Escribir hoja
    var outSheet = getSheetFlexible(sheetName);
    if (!outSheet) outSheet = ss.insertSheet(sheetName);
    outSheet.clear();
    var rows = [['LeadID', 'ultimaISO', 'origen', 'actor', 'fila']];
    Object.keys(index).forEach(function (k) { rows.push([k, index[k].ultimaISO, index[k].origen, index[k].actor, index[k].fila]); });
    if (rows.length > 0) outSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    return { success: true, total: Object.keys(index).length, tiempoMs: new Date().getTime() - start };
  } catch (e) {
    return { success: false, error: e && e.toString() };
  }
}

/**
 * Lee la hoja `Ultima_Actividad` y devuelve un mapa { leadId: {ultimaISO, origen, actor, fila} }
 */
function getUltimaActividadMap() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'Ultima_Actividad';

    // Búsqueda flexible
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim() === sheetName) {
          sheet = sheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      console.warn('⚠️ No se encontró la hoja ' + sheetName);
      return {};
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var map = {};
    data.forEach(function (row) {
      var k = row[0] ? row[0].toString().trim().replace(/\.0$/, "") : '';
      if (!k) return;
      map[k] = { ultimaISO: row[1] || null, origen: row[2] || null, actor: row[3] || null, fila: row[4] || null };
    });
    console.log('✅ Mapa de actividad cargado: ' + Object.keys(map).length + ' leads');
    return map;
  } catch (e) { return {}; }
}