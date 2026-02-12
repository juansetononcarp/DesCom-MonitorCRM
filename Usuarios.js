// ============================================
// MÓDULO: USUARIOS
// ============================================

var Usuarios = (function() {
  
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
function getUltimaActividadAsociado(usuarioId) {
  try {
    if (!usuarioId) return null;
    var usuarioIdStr = usuarioId.toString().trim().toLowerCase();

    var maxDate = null;

    // 1) Revisar Leads (usa la función existente para obtener leads por usuario)
    try {
      var leads = Leads.getLeadsPorUsuario(usuarioId);
      leads.forEach(function(lead) {
        if (lead && lead.ultimaGestion) {
          var d = parsePossiblySheetDate(lead.ultimaGestion);
          if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
        }
      });
    } catch (e) {
      console.warn('Advertencia: no se pudo leer leads para última actividad:', e && e.toString());
    }

    // 2) Otras hojas que pueden contener actividades
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var candidateSheets = ['Comentarios', 'Agenda', 'Tareas', 'Viajes'];

    candidateSheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      try {
        var last = scanSheetForUserDates(sheet, usuarioIdStr);
        if (last && (!maxDate || last.getTime() > maxDate.getTime())) maxDate = last;
      } catch (e) {
        console.warn('Error escaneando hoja ' + sheetName + ': ' + (e && e.toString()));
      }
    });

    // 3) Además, escanear por LeadID en hojas de actividad (por si la acción la hizo otro actor)
    try {
      var leadIds = [];
      try {
        var usuarioLeads = Leads.getLeadsPorUsuario(usuarioId) || [];
        leadIds = usuarioLeads.map(function(l) { return (l.leadId !== undefined && l.leadId !== null) ? l.leadId.toString().toLowerCase() : ''; }).filter(Boolean);
      } catch (e) {
        console.warn('No se pudieron obtener leadIds del usuario:', e && e.toString());
      }

      if (leadIds.length > 0) {
        candidateSheets.forEach(function(sheetName) {
          var sheet = ss.getSheetByName(sheetName);
          if (!sheet) return;
          try {
            var lastByLead = scanSheetForUserOrLeadDates(sheet, usuarioIdStr, leadIds);
            if (lastByLead && (!maxDate || lastByLead.getTime() > maxDate.getTime())) maxDate = lastByLead;
          } catch (e) {
            console.warn('Error escaneando por lead en hoja ' + sheetName + ': ' + (e && e.toString()));
          }
        });
      }
    } catch (e) {
      // ignore
    }

    // Formatear resultado usando Fechas.formatear si existe
    if (maxDate) return Fechas && Fechas.formatear ? Fechas.formatear(maxDate) : maxDate.toString();
    return null;
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
  } catch (e) {}

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var maxDate = null;

  var leadSet = new Set((leadIds || []).map(function(x){return x.toString().toLowerCase();}));

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