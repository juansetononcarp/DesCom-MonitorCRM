// ============================================
// MÓDULO: LEADS
// ============================================

var Leads = (function () {

  // ===== CONFIGURACIÓN DE COLUMNAS =====
  const CONFIG = {
    // Columnas principales
    COL_ID_USUARIO: 2,     // Columna C - "AC asignado"
    COL_ESTADO: 37,        // Columna AL - "ESTADO"
    COL_ID_LEAD: 0,        // Columna A - "LeadID"
    COL_NOMBRE: 4,         // Columna E - "Nombre"
    COL_APELLIDO: 5,       // Columna F - "Apellido"
    COL_FECHA: 1,          // Columna B - "Fecha de asignación"

    // Columnas adicionales (para futuras métricas)
    COL_CREADO_POR: 30,    // Columna AE - "Creado por"
    COL_K: 33,             // Columna AH - "K"
    COL_KV: 34,            // Columna AI - "Kv"
    COL_ULTIMA_GESTION: 38 // Columna AM - "Ult Coment AC"
  };

  // ===== MAPEO DE ESTADOS - VERSIÓN COMPLETA CON 6 ESTADOS =====
  const MAPA_ESTADOS = {
    'NUEVO': ['NUEVO', 'NUEVA'],
    'ABIERTO': ['ABIERTO', 'ABIERTA'],
    'CONVERTIDO': ['CONVERTIDO', 'CONVERTIDA'],
    'ACTIVO': ['ACTIVO', 'ACTIVA'],
    'OPERANDO': ['OPERANDO'],
    'DESCARTADO': ['DESCARTADO', 'DESCARTADA', 'PERDIDO', 'PERDIDA'],
    'OTROS': []
  };

  /**
   * Obtiene todos los leads de la hoja 'Leads'
   * @returns {Array} Array de objetos lead
   */
  function getLeads() {
    try {
      // Intentar leer de caché primero (ScriptMem o PropertiesService para persistencia corta)
      // Usamos CacheService para mejorar velocidad entre llamadas consecutivas del dashboard
      const cache = CacheService.getScriptCache();
      const cached = cache.get('LEADS_CACHE');
      if (cached) {
        console.log('✅ Leads.getLeads: Recuperado de caché');
        return JSON.parse(cached);
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Leads');

      if (!sheet) throw new Error('No se encontró la hoja "Leads"');

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();

      if (lastRow < 2) return [];

      const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

      const leads = data.map((row, index) => ({
        // Datos básicos
        idUsuario: row[CONFIG.COL_ID_USUARIO] ? row[CONFIG.COL_ID_USUARIO].toString().trim() : '',
        estado: row[CONFIG.COL_ESTADO] ? row[CONFIG.COL_ESTADO].toString().trim() : '',
        leadId: row[CONFIG.COL_ID_LEAD],
        nombre: row[CONFIG.COL_NOMBRE] || '',
        apellido: row[CONFIG.COL_APELLIDO] || '',
        fecha: row[CONFIG.COL_FECHA],

        // Datos adicionales
        creadoPor: row[CONFIG.COL_CREADO_POR] ? row[CONFIG.COL_CREADO_POR].toString().trim() : '',
        k: row[CONFIG.COL_K] || '',
        kv: row[CONFIG.COL_KV] || '',
        ultimaGestion: row[CONFIG.COL_ULTIMA_GESTION] || null,

        // Fila original para acceso futuro (reducimos data para cache)
        _filaNumero: index + 2
      }));

      console.log(`✅ Leads.getLeads: ${leads.length} leads obtenidos (origen)`);

      // Guardar en caché (try-catch por si excede 100KB)
      try {
        cache.put('LEADS_CACHE', JSON.stringify(leads), 300); // 5 minutos
      } catch (e) {
        console.warn('⚠️ Cache leads fulled:', e);
      }

      return leads;

    } catch (error) {
      console.error('❌ Error en Leads.getLeads:', error);
      throw error;
    }
  }

  /**
   * Obtiene leads filtrados por usuario
   * @param {string} usuarioId - Email del usuario
   * @returns {Array} Leads del usuario
   */
  function getLeadsPorUsuario(usuarioId) {
    const todosLosLeads = getLeads();
    const usuarioIdStr = usuarioId.toString().trim().toLowerCase();

    // Obtener nombre del usuario para búsqueda flexible (por si en la hoja Leads pusieron el Nombre en vez del Mail)
    let usuarioNombreStr = "";
    try {
      const usuario = Usuarios.getUsuarioPorId(usuarioId);
      if (usuario && usuario.nombre) {
        usuarioNombreStr = usuario.nombre.toString().trim().toLowerCase();
      }
    } catch (e) {
      console.warn('Leads.getLeadsPorUsuario: No se pudo obtener nombre para fallback match', e);
    }

    const leadsFiltrados = todosLosLeads.filter(lead => {
      if (!lead.idUsuario) return false;
      const val = lead.idUsuario.toString().trim().toLowerCase();
      // Match por ID (Mail) O por Nombre
      return val === usuarioIdStr || (usuarioNombreStr && val === usuarioNombreStr);
    });

    console.log(`✅ Leads.getLeadsPorUsuario: ${leadsFiltrados.length} leads para ${usuarioId} (Matching: ID=${usuarioIdStr}, Nombre=${usuarioNombreStr})`);
    return leadsFiltrados;
  }

  /**
   * Clasifica un estado según el mapeo (VERSIÓN ACTUALIZADA CON 6 ESTADOS)
   * @param {string} estadoRaw - Estado original
   * @returns {string} Estado clasificado
   */
  function clasificarEstado(estadoRaw) {
    if (!estadoRaw) return 'OTROS';

    const estado = estadoRaw.toString().trim().toUpperCase();

    for (const [categoria, palabras] of Object.entries(MAPA_ESTADOS)) {
      for (const palabra of palabras) {
        if (estado.includes(palabra)) {
          return categoria;
        }
      }
    }

    return 'OTROS';
  }

  /**
   * Formatea un lead para el frontend
   * @param {Object} lead - Objeto lead
   * @returns {Object} Lead formateado
   */
  function formatearLead(lead, indexMap) {
    const nombreCompleto = `${lead.nombre || ''} ${lead.apellido || ''}`.trim() || 'Sin nombre';

    // Obtener última actividad del indexMap
    let ultimaActividadStr = '—';
    let ultimaActividadOrigen = '';
    let ultimaActividadActor = '';
    if (indexMap && lead.leadId && indexMap[lead.leadId] && indexMap[lead.leadId].ultimaISO) {
      ultimaActividadStr = Fechas.formatear(new Date(indexMap[lead.leadId].ultimaISO));
      ultimaActividadOrigen = indexMap[lead.leadId].origen || '';
      ultimaActividadActor = indexMap[lead.leadId].actor || '';
    }

    return {
      id: lead.leadId !== undefined && lead.leadId !== null ? lead.leadId.toString() : 'N/A',
      nombre: nombreCompleto,
      fecha: Fechas.formatear(lead.fecha),
      estado: lead.estado || 'Sin estado',
      estadoClasificado: clasificarEstado(lead.estado),
      creadoPor: lead.creadoPor || null,
      k: lead.k || '',
      kv: lead.kv || '',
      ultimaActividad: ultimaActividadStr,
      ultimaActividadOrigen: ultimaActividadOrigen,
      ultimaActividadActor: ultimaActividadActor
    };
  }

  /**
   * Formatea múltiples leads
   * @param {Array} leads - Array de leads
   * @returns {Array} Leads formateados
   */
  function formatearLeads(leads) {
    if (!leads || !leads.length) return [];
    // allow optional indexMap as last arg
    var indexMap = null;
    if (arguments.length > 1) indexMap = arguments[1];
    return leads.map(lead => formatearLead(lead, indexMap));
  }

  /**
   * Obtiene leads formateados por usuario
   * @param {string} usuarioId - Email del usuario
   * @returns {Array} Leads formateados
   */
  function getLeadsFormateadosPorUsuario(usuarioId) {
    const leads = getLeadsPorUsuario(usuarioId);
    // obtener índice de última actividad si existe
    var indexMap = {};
    try { indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { indexMap = {}; }

    return formatearLeads(leads, indexMap).sort((a, b) => {
      const fechaA = a.fecha === 'Sin fecha' ? 0 : new Date(a.fecha).getTime() || 0;
      const fechaB = b.fecha === 'Sin fecha' ? 0 : new Date(b.fecha).getTime() || 0;
      return fechaB - fechaA;
    });
  }

  /**
   * Obtiene estadísticas de estados para un conjunto de leads (VERSIÓN ACTUALIZADA)
   * @param {Array} leads - Array de leads
   * @returns {Object} Conteo de estados
   */
  function contarEstados(leads) {
    const conteo = {
      'NUEVO': 0,
      'ABIERTO': 0,
      'CONVERTIDO': 0,
      'ACTIVO': 0,
      'OPERANDO': 0,
      'DESCARTADO': 0,
      'OTROS': 0
    };

    const otrosDetalles = {};

    leads.forEach(lead => {
      const estadoClasificado = clasificarEstado(lead.estado);

      if (conteo.hasOwnProperty(estadoClasificado)) {
        conteo[estadoClasificado]++;
      } else {
        conteo['OTROS']++;
        otrosDetalles[lead.estado] = (otrosDetalles[lead.estado] || 0) + 1;
      }
    });

    return { conteo, otrosDetalles };
  }

  /**
   * Calcula porcentajes de estados
   * @param {Object} conteo - Objeto con conteos
   * @param {number} total - Total de leads
   * @returns {Object} Porcentajes
   */
  function calcularPorcentajes(conteo, total) {
    const porcentajes = {};
    Object.keys(conteo).forEach(key => {
      porcentajes[key] = total > 0 ? Math.round((conteo[key] / total) * 100) : 0;
    });
    return porcentajes;
  }

  // API Pública
  return {
    CONFIG: CONFIG,
    MAPA_ESTADOS: MAPA_ESTADOS,
    getLeads: getLeads,
    getLeadsPorUsuario: getLeadsPorUsuario,
    getLeadsFormateadosPorUsuario: getLeadsFormateadosPorUsuario,
    clasificarEstado: clasificarEstado,
    formatearLead: formatearLead,
    formatearLeads: formatearLeads,
    contarEstados: contarEstados,
    calcularPorcentajes: calcularPorcentajes
  };

})();

// ================================
// Utilidades para última actividad por Lead
// ================================

/**
 * Intenta convertir distintos formatos de fecha (Date, número de Sheets, string) a Date.
 */
function parsePossiblySheetDate(value) {
  try {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
      if (value > 25000) return new Date((value - 25569) * 86400 * 1000);
      return null;
    }
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Escanea una hoja buscando filas que mencionen `leadId` y devuelve la última fecha encontrada (Date) o null.
 * Usa heurísticas parecidas a las de `Usuarios.scanSheetForUserDates`.
 */
function scanSheetForLeadDates(sheet, leadId) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var dateCol = -1;
  var dateRegex = /(fecha|date|created|timestamp|hora|time|datetime|creado)/i;
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c] ? headers[c].toString() : '';
    if (dateCol === -1 && dateRegex.test(h)) dateCol = c;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var maxDate = null;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var found = false;
    for (var cc = 0; cc < row.length; cc++) {
      if (row[cc] && row[cc].toString().toLowerCase().indexOf((leadId || '').toString().toLowerCase()) !== -1) {
        found = true; break;
      }
    }
    if (found) {
      if (dateCol >= 0) {
        var d = parsePossiblySheetDate(row[dateCol]);
        if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
      } else {
        // intentar detectar alguna celda con fecha en la fila
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
 * Obtiene la última actividad para un lead consultando `Leads.ultimaGestion` y escaneando hojas: Comentarios, Agenda, Tareas, Viajes.
 */
function getUltimaActividadPorLead(leadId) {
  try {
    if (!leadId && leadId !== 0) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var candidateSheets = ['Comentarios', 'Agenda', 'Tareas', 'Viajes'];
    var maxDate = null;

    candidateSheets.forEach(function (sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      try {
        var found = scanSheetForLeadDates(sheet, leadId);
        if (found && (!maxDate || found.getTime() > maxDate.getTime())) maxDate = found;
      } catch (e) {
        console.warn('Error escaneando ' + sheetName + ': ' + (e && e.toString()));
      }
    });

    return maxDate;
  } catch (e) {
    return null;
  }
}