// ============================================
// MÓDULO: USUARIOS
// ============================================

var Usuarios = (function () {

  /**
   * Obtiene todos los usuarios de la hoja 'Usuarios'
   */
  function getUsuarios() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Usuarios');
      if (!sheet) return [];
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      return data
        .filter(row => row[1])
        .map(row => ({ id: row[0].toString().trim(), nombre: row[1].toString().trim() }))
        .filter(u => u.id && u.nombre);
    } catch (e) { return []; }
  }

  function getNombrePorId(usuarioId) {
    var us = getUsuarios();
    var u = us.find(x => x.id.toLowerCase() === usuarioId.toLowerCase());
    return u ? u.nombre : usuarioId;
  }

  function getUsuarioPorId(usuarioId) {
    var us = getUsuarios();
    return us.find(x => x.id.toLowerCase() === usuarioId.toLowerCase()) || null;
  }

  function getEstadisticas() {
    var us = getUsuarios();
    return { total: us.length, activos: us.length, lista: us };
  }

  function parsePossiblySheetDate(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'number' && value > 25000) return new Date((value - 25569) * 86400 * 1000);
    var p = new Date(value);
    return isNaN(p.getTime()) ? null : p;
  }

  function buildUltimaActividadIndex() {
    try {
      var start = new Date().getTime();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetName = 'Ultima_Actividad';
      var index = {};

      function normalizarId(val) {
        if (val === null || val === undefined) return "";
        return val.toString().trim().replace(/\.0$/, "");
      }

      var configs = [
        { name: 'Tareas', leadCol: 2, dateCol: 6, actorCol: 9 },
        { name: 'Agenda', leadCol: 2, dateCol: 10, actorCol: 9 },
        { name: 'Comentarios', leadCol: 2, dateCol: 5, actorCol: 7 }
      ];

      configs.forEach(function (conf) {
        var sheet = ss.getSheetByName(conf.name);
        if (!sheet) return;
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        var leadVals = sheet.getRange(2, conf.leadCol, lastRow - 1, 1).getValues();
        var dateVals = sheet.getRange(2, conf.dateCol, lastRow - 1, 1).getValues();
        var actorVals = sheet.getRange(2, conf.actorCol, lastRow - 1, 1).getValues();

        for (var i = 0; i < leadVals.length; i++) {
          var lid = normalizarId(leadVals[i][0]);
          if (!lid) continue;
          var d = parsePossiblySheetDate(dateVals[i][0]);
          if (!d) continue;
          var actor = actorVals[i][0] ? actorVals[i][0].toString().trim() : 'N/A';
          var existing = index[lid];
          if (!existing || d.getTime() > new Date(existing.ultimaISO).getTime()) {
            index[lid] = { ultimaISO: d.toISOString(), origen: conf.name, actor: actor, fila: i + 2 };
          }
        }
      });

      var out = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
      out.clear();
      var rows = [['LeadID', 'ultimaISO', 'origen', 'actor', 'fila']];
      Object.keys(index).forEach(function (k) { rows.push([k, index[k].ultimaISO, index[k].origen, index[k].actor, index[k].fila]); });
      if (rows.length > 0) out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      return { success: true, count: Object.keys(index).length };
    } catch (e) { return { success: false, error: e.toString() }; }
  }

  function getUltimaActividadMap() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var s = ss.getSheetByName('Ultima_Actividad');
      if (!s) return {};
      var lastRow = s.getLastRow();
      if (lastRow < 2) return {};
      var data = s.getRange(2, 1, lastRow - 1, 5).getValues();
      var map = {};
      data.forEach(function (row) {
        var k = row[0] ? row[0].toString().trim().replace(/\.0$/, "") : '';
        if (k) map[k] = { ultimaISO: row[1], origen: row[2], actor: row[3], fila: row[4] };
      });
      return map;
    } catch (e) { return {}; }
  }

  function getUltimaActividadAsociado(usuarioId) {
    try {
      if (!usuarioId) return "Sin datos";
      var uid = usuarioId.toString().trim().toLowerCase();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var maxDate = null;

      function check(name, uCol, dCol) {
        var s = ss.getSheetByName(name);
        if (!s) return;
        var last = s.getLastRow();
        if (last < 2) return;
        var uVs = s.getRange(2, uCol, last - 1, 1).getValues();
        var dVs = s.getRange(2, dCol, last - 1, 1).getValues();
        for (var i = 0; i < uVs.length; i++) {
          if (uVs[i][0] && uVs[i][0].toString().trim().toLowerCase() === uid) {
            var d = parsePossiblySheetDate(dVs[i][0]);
            if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
          }
        }
      }

      check('Leads', 31, 2);
      check('Tareas', 9, 6);
      check('Agenda', 9, 10);
      check('Comentarios', 7, 5);

      return maxDate ? (Fechas.formatear ? Fechas.formatear(maxDate) : maxDate.toLocaleString()) : "Sin actividad reciente";
    } catch (e) { return "Error"; }
  }

  return {
    getUsuarios: getUsuarios,
    getNombrePorId: getNombrePorId,
    getUsuarioPorId: getUsuarioPorId,
    getEstadisticas: getEstadisticas,
    buildUltimaActividadIndex: buildUltimaActividadIndex,
    getUltimaActividadMap: getUltimaActividadMap,
    getUltimaActividadAsociado: getUltimaActividadAsociado
  };

})();