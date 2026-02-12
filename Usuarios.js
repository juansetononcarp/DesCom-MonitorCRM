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
 * Escanea una hoja buscando filas asociadas a `usuarioIdStr` y devuelve la última fecha encontrada (Date) o null.
 * Usa heurísticas sobre encabezados para encontrar columnas de usuario y fecha.
 */
function scanSheetForUserDates(sheet, usuarioIdStr) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  // Buscar columna usuario y columna fecha por encabezado
  var userCol = -1;
  var dateCol = -1;
  var userRegex = /(asesor|ac asignado|ac|asociad|usuario|mail|email|responsable|creado por)/i;
  var dateRegex = /(fecha|date|created|timestamp|hora|time|datetime|creado)/i;

  for (var c = 0; c < headers.length; c++) {
    var h = headers[c] ? headers[c].toString() : '';
    if (userCol === -1 && userRegex.test(h)) userCol = c;
    if (dateCol === -1 && dateRegex.test(h)) dateCol = c;
  }

  // Priorizar columnas conocidas para actor según el nombre de la hoja
  try {
    var sheetName = sheet.getName();
    var sheetActorCols = {
      'Tareas': 8,   // columna I -> índice 8
      'Agenda': 8,   // columna I -> índice 8
      'Comentarios': 6 // columna G -> índice 6
    };
    if (sheetActorCols[sheetName] !== undefined) {
      userCol = sheetActorCols[sheetName];
    }
  } catch (e) {
    // ignore
  }

  // Leer datos (desde fila 2)
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var maxDate = null;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];

    // Si se encontraron columnas por encabezado
    if (userCol >= 0 && dateCol >= 0) {
      var cellUser = row[userCol] ? row[userCol].toString().trim().toLowerCase() : '';
      if (cellUser && (cellUser === usuarioIdStr || cellUser.indexOf(usuarioIdStr) !== -1)) {
        var d = parsePossiblySheetDate(row[dateCol]);
        if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
      }
      continue;
    }

    // Si no hay encabezados detectados, hacer heurística: buscar usuario en cualquier columna
    var foundUser = false;
    for (var cc = 0; cc < row.length; cc++) {
      if (row[cc] && row[cc].toString().toLowerCase().indexOf(usuarioIdStr) !== -1) {
        foundUser = true; break;
      }
    }
    if (foundUser) {
      // intentar encontrar alguna celda en la fila que parezca fecha
      for (var dd = 0; dd < row.length; dd++) {
        var candidate = parsePossiblySheetDate(row[dd]);
        if (candidate && (!maxDate || candidate.getTime() > maxDate.getTime())) maxDate = candidate;
      }
    }
  }

  return maxDate;
}

/**
 * Escanea una hoja buscando filas donde el actor sea `usuarioIdStr` o donde aparezca algún `leadId` del conjunto.
 * Devuelve la última fecha encontrada (Date) o null.
 */
function scanSheetForUserOrLeadDates(sheet, usuarioIdStr, leadIds) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  var userCol = -1;
  var dateCol = -1;
  var userRegex = /(asesor|ac asignado|ac|asociad|usuario|mail|email|responsable|creado por)/i;
  var dateRegex = /(fecha|date|created|timestamp|hora|time|datetime|creado)/i;

  for (var c = 0; c < headers.length; c++) {
    var h = headers[c] ? headers[c].toString() : '';
    if (userCol === -1 && userRegex.test(h)) userCol = c;
    if (dateCol === -1 && dateRegex.test(h)) dateCol = c;
  }

  // Priorizar columnas conocidas según hoja
  try {
    var sheetName = sheet.getName();
    var sheetActorCols = {
      'Tareas': 8,
      'Agenda': 8,
      'Comentarios': 6
    };
    if (sheetActorCols[sheetName] !== undefined) userCol = sheetActorCols[sheetName];
  } catch (e) { }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var maxDate = null;

  var leadSet = new Set((leadIds || []).map(function (x) { return x.toString().toLowerCase(); }));

  for (var r = 0; r < data.length; r++) {
    var row = data[r];

    var matched = false;

    // 1) actor coincide
    if (userCol >= 0) {
      var cellUser = row[userCol] ? row[userCol].toString().trim().toLowerCase() : '';
      if (cellUser && (cellUser === usuarioIdStr || cellUser.indexOf(usuarioIdStr) !== -1)) matched = true;
    }

    // 2) si no matched, buscar LeadID en cualquier celda
    if (!matched && leadSet.size > 0) {
      for (var cc = 0; cc < row.length; cc++) {
        if (!row[cc]) continue;
        var text = row[cc].toString().toLowerCase();
        // buscar si algun leadId está contenido
        for (var lid of leadSet) {
          if (text.indexOf(lid) !== -1) { matched = true; break; }
        }
        if (matched) break;
      }
    }

    if (matched) {
      if (dateCol >= 0) {
        var d = parsePossiblySheetDate(row[dateCol]);
        if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
      } else {
        for (var dd = 0; dd < row.length; dd++) {
          var cand = parsePossiblySheetDate(row[dd]);
          if (cand && (!maxDate || cand.getTime() > maxDate.getTime())) maxDate = cand;
        }
      }
    }
  }

  return maxDate;
}

/**
 * Diagnóstico rápido: devuelve estadísticas resumidas sin escanear hojas completas.
 * Escanea como máximo `maxRows` por hoja y limita columnas para ser rápido.
 */
/**
 * Diagnóstico rápido: devuelve estadísticas resumidas y VERIFICACIÓN DE COLUMNAS.
 * Escanea como máximo `maxRows` por hoja y limita columnas para ser rápido.
 */
function getDiagnosticoRapido(usuarioId, maxRows) {
  try {
    maxRows = typeof maxRows === 'number' ? maxRows : 1000;
    var start = new Date().getTime();
    var result = {
      usuario: usuarioId,
      tiempoInicio: new Date().toISOString(),
      leadsCount: 0,
      leadsSheetInfo: {},
      sampleLeadIds: [],
      sheets: [],
      tiempoMs: 0
    };

    // 1. Diagnóstico de Hoja Leads (Columnas)
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetLeads = ss.getSheetByName('Leads');
      if (sheetLeads) {
        var headers = sheetLeads.getRange(1, 1, 1, Math.min(sheetLeads.getLastColumn(), 50)).getValues()[0];
        // Verificar columnas clave
        // Leads.CONFIG.COL_ID_USUARIO is 2 (3rd col)
        // Leads.CONFIG.COL_ESTADO is 37 (38th col)

        var colUsuarioVal = headers.length > 2 ? headers[2] : 'N/A';
        var colEstadoVal = headers.length > 37 ? headers[37] : 'N/A';

        // Muestra de datos
        var sampleData = [];
        if (sheetLeads.getLastRow() > 1) {
          sampleData = sheetLeads.getRange(2, 1, Math.min(5, sheetLeads.getLastRow() - 1), 40).getValues().map(function (r) {
            return {
              col_C_Usuario: r[2],
              col_AL_Estado: r[37],
              col_E_Nombre: r[4]
            };
          });
        }

        result.leadsSheetInfo = {
          exists: true,
          headersCount: headers.length,
          header_Col_C: colUsuarioVal,
          header_Col_AL: colEstadoVal,
          sampleRows: sampleData
        };
      } else {
        result.leadsSheetInfo = { exists: false };
      }
    } catch (e) {
      result.leadsSheetInfo = { error: e.toString() };
    }

    // Obtener leads del usuario (usando lógica actual)
    var usuarioLeads = [];
    try { usuarioLeads = Leads.getLeadsPorUsuario(usuarioId) || []; } catch (e) { usuarioLeads = []; }
    result.leadsCount = usuarioLeads.length;
    result.sampleLeadIds = usuarioLeads.slice(0, 5).map(function (l) { return l.leadId; });

    var leadIds = result.sampleLeadIds.map(function (x) { return x ? x.toString().toLowerCase() : ''; }).filter(Boolean);

    // Escanear otras hojas (limitado)
    var candidateSheets = ['Comentarios', 'Agenda', 'Tareas', 'Viajes'];

    candidateSheets.forEach(function (sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        result.sheets.push({ name: sheetName, exists: false });
        return;
      }

      var lastRow = sheet.getLastRow();
      var lastCol = Math.min(sheet.getLastColumn(), 20); // limitar columnas para velocidad
      var rowsToScan = Math.min(Math.max(0, lastRow - 1), maxRows);
      var stats = { name: sheetName, exists: true, lastRow: lastRow, scannedRows: rowsToScan, actorMatches: 0, leadMatches: 0 };

      if (rowsToScan > 0) {
        var data = sheet.getRange(2, 1, rowsToScan, lastCol).getValues();

        // columnas actor heurísticas
        var actorCols = { 'Tareas': 8, 'Agenda': 8, 'Comentarios': 6 };
        var actorCol = actorCols[sheetName] !== undefined ? actorCols[sheetName] : -1;

        var leadSet = new Set(leadIds);

        for (var r = 0; r < data.length; r++) {
          var row = data[r];
          // actor match
          if (actorCol >= 0 && actorCol < row.length) {
            var cellUser = row[actorCol] ? row[actorCol].toString().trim().toLowerCase() : '';
            if (cellUser && cellUser.indexOf((usuarioId || '').toString().toLowerCase()) !== -1) stats.actorMatches++;
          }
          // lead match
          if (leadSet.size > 0 && row.length > 0) {
            for (var c = 0; c < Math.min(row.length, 8); c++) {
              if (!row[c]) continue;
              var text = row[c].toString().toLowerCase();
              for (var lid of leadSet) {
                if (lid && text.indexOf(lid) !== -1) { stats.leadMatches++; break; }
              }
            }
          }
        }
      }

      result.sheets.push(stats);
    });

    result.tiempoMs = new Date().getTime() - start;
    return result;
  } catch (e) {
    return { error: e && e.toString() };
  }
}

/**
 * Reconstruye la hoja `Ultima_Actividad` con la última fecha por LeadID.
 * Escanea las hojas de actividad y guarda: LeadID | ultimaISO | origen | actor | fila
 */
function buildUltimaActividadIndex() {
  try {
    var start = new Date().getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'Ultima_Actividad';

    // Mapa temporal
    var index = {};

    var candidateSheets = ['Comentarios', 'Agenda', 'Tareas', 'Viajes', 'Leads'];

    candidateSheets.forEach(function (sName) {
      var sheet = ss.getSheetByName(sName);
      if (!sheet) return;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < 2) return;

      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
      // buscar columna leadId
      var leadCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = headers[c] ? headers[c].toString().toLowerCase() : '';
        if (/lead\s*id|leadid|lead|id\b/.test(h)) { leadCol = c; break; }
      }
      // buscar columna fecha
      var dateCol = -1;
      for (var c = 0; c < headers.length; c++) {
        var h = headers[c] ? headers[c].toString().toLowerCase() : '';
        if (/(fecha|date|created|timestamp|hora|time|datetime|creado)/.test(h)) { dateCol = c; break; }
      }
      // actor col heurística
      var actorCols = { 'Tareas': 8, 'Agenda': 8, 'Comentarios': 6 };
      var actorCol = actorCols[sName] !== undefined ? actorCols[sName] : -1;

      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

      for (var r = 0; r < data.length; r++) {
        var row = data[r];
        var leadVal = null;
        if (leadCol >= 0 && leadCol < row.length) leadVal = row[leadCol];
        if (!leadVal) {
          // intentar detectar en fila
          for (var c = 0; c < Math.min(row.length, 10); c++) {
            if (row[c] && row[c].toString().match(/[A-Za-z0-9\-]{2,}/)) { leadVal = row[c]; break; }
          }
        }
        if (!leadVal) continue;
        var leadId = leadVal.toString().trim();

        var dateVal = null;
        if (dateCol >= 0 && dateCol < row.length) dateVal = row[dateCol];
        // fallback: buscar cualquier fecha en la fila
        if (!dateVal) {
          for (var c = 0; c < row.length; c++) {
            var d = parsePossiblySheetDate(row[c]);
            if (d) { dateVal = d; break; }
          }
        }
        var parsed = parsePossiblySheetDate(dateVal);
        if (!parsed) continue;

        var actor = (actorCol >= 0 && actorCol < row.length && row[actorCol]) ? row[actorCol].toString() : '';

        var existing = index[leadId];
        if (!existing || parsed.getTime() > new Date(existing.ultimaISO).getTime()) {
          index[leadId] = { ultimaISO: parsed.toISOString(), origen: sName, actor: actor, fila: r + 2 };
        }
      }
    });

    // Escribir hoja
    var outSheet = ss.getSheetByName(sheetName);
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
    var sheet = ss.getSheetByName('Ultima_Actividad');
    if (!sheet) return {};
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var map = {};
    data.forEach(function (row) {
      var k = row[0] ? row[0].toString() : '';
      if (!k) return;
      map[k.toString()] = { ultimaISO: row[1] || null, origen: row[2] || null, actor: row[3] || null, fila: row[4] || null };
    });
    return map;
  } catch (e) { return {}; }
}