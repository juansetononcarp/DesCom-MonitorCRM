// ============================================
// ORQUESTADOR PRINCIPAL - DASHBOARD DE LEADS
// ============================================

// ========== SERVIDOR WEB ==========

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Monitor CRM v1.3.11 [PRECISION]')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ========== API V1 - COMPATIBILIDAD HACIA ATRÁS ==========

/**
 * @deprecated Usar Usuarios.getUsuarios()
 */
function getUsuarios() {
  return Usuarios.getUsuarios();
}

/**
 * @deprecated Usar Leads.getLeadsFormateadosPorUsuario()
 */
function getEstadisticasUsuario(usuarioId) {
  try {
    console.log(`🔄 [API V1] getEstadisticasUsuario: ${usuarioId}`);

    // 1. Obtener leads del usuario
    const leads = Leads.getLeadsPorUsuario(usuarioId);

    // 2. Calcular métricas
    const metricas = Metricas.calcular(leads, { usuarioId });

    // 3. Formatear leads para tabla
    var indexMap = {};
    try { indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { indexMap = {}; }
    const todosLosLeads = leads
      .map(lead => Leads.formatearLead(lead, indexMap))
      .sort((a, b) => {
        const fechaA = a.fecha === 'Sin fecha' ? 0 : new Date(a.fecha).getTime() || 0;
        const fechaB = b.fecha === 'Sin fecha' ? 0 : new Date(b.fecha).getTime() || 0;
        return fechaB - fechaA;
      });

    // 4. Obtener nombre del usuario
    const nombreUsuario = Usuarios.getNombrePorId(usuarioId);

    return {
      usuarioId: usuarioId,
      usuarioNombre: nombreUsuario,
      total: leads.length,
      estados: metricas.estados,
      porcentajes: metricas.porcentajes,
      todosLosLeads: todosLosLeads,
      ultimosLeads: todosLosLeads.slice(0, 5),
      otrosEstados: metricas.otrosEstados || {},
      fechaActualizacion: new Date().toLocaleString('es-ES'),
      resumen: `📊 ${leads.length} leads para ${nombreUsuario}`
    };

  } catch (error) {
    console.error(`❌ [API V1] Error en getEstadisticasUsuario:`, error);
    throw error;
  }
}

/**
 * @deprecated Usar Leads.getLeads() + Metricas.calcular()
 */
function getEstadisticasGenerales() {
  try {
    console.log(`🔄 [API V1] getEstadisticasGenerales`);

    const todosLosLeads = Leads.getLeads();
    const metricas = Metricas.calcular(todosLosLeads);
    const usuarios = Usuarios.getUsuarios();

    return {
      total: todosLosLeads.length,
      estados: metricas.estados,
      usuariosActivos: usuarios.length,
      fechaActualizacion: new Date().toLocaleString('es-ES')
    };

  } catch (error) {
    console.error(`❌ [API V1] Error en getEstadisticasGenerales:`, error);
    throw error;
  }
}

// ========== API V2 - NUEVAS FUNCIONES CON FILTROS ==========

/**
 * Obtiene años disponibles de todos los leads
 */
function getAniosAbiertos() {
  try {
    const todosLosLeads = Leads.getLeads();
    return Fechas.getAniosDisponibles(todosLosLeads);
  } catch (error) {
    console.error('Error en getAniosAbiertos:', error);
    return [];
  }
}

/**
 * Versión mejorada de getEstadisticasUsuario con filtros
 */
function getEstadisticasUsuarioConFiltros(usuarioId, filtros = {}) {
  try {
    console.log(`🔄 getEstadisticasUsuarioConFiltros: ${usuarioId}`, filtros);

    // 1. Obtener leads del usuario
    const leads = Leads.getLeadsPorUsuario(usuarioId);

    // 2. Aplicar filtros de fecha y DEDUPLICAR por LeadID
    const uniqueLeadsMap = new Map();
    leads.forEach(lead => {
      // Filtrar por fecha antes de guardar en el mapa de únicos
      let matchAnio = !filtros.anio || filtros.anio === "" || Fechas.getAnio(lead.fecha) === parseInt(filtros.anio);
      let matchMes = !filtros.mes || filtros.mes === "" || Fechas.getMes(lead.fecha) === parseInt(filtros.mes);

      if (matchAnio && matchMes) {
        const lid = lead.leadId || ("TEMP_" + Math.random());
        // Si hay duplicados, conservamos el primero o el último (aquí el último)
        uniqueLeadsMap.set(lid, lead);
      }
    });

    const leadsFiltradosUnicos = Array.from(uniqueLeadsMap.values());

    // 3. Calcular métricas sobre leads únicos
    const metricas = Metricas.calcular(leadsFiltradosUnicos, { usuarioId });
    const metricasCuentas = Metricas.calcularMetricasCuentas(leadsFiltradosUnicos);

    // 4. Formatear leads (aquí podemos mostrar todos o solo los únicos, optamos por únicos para consistencia)
    var indexMap = {};
    try { indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { indexMap = {}; }
    const leadsFormateados = leadsFiltradosUnicos
      .map(lead => Leads.formatearLead(lead, indexMap))
      .sort((a, b) => {
        // Ordenar por fecha ISO si existe, si no por fecha normal
        const valA = a.fechaISO || '';
        const valB = b.fechaISO || '';
        if (valA < valB) return 1;
        if (valA > valB) return -1;
        return 0;
      });

    const nombreUsuario = Usuarios.getNombrePorId(usuarioId);

    return {
      usuarioId: usuarioId,
      usuarioNombre: nombreUsuario,
      total: leadsFiltradosUnicos.length,
      estados: metricas.estados,
      porcentajes: metricas.porcentajes,
      todosLosLeads: leadsFormateados,
      ultimosLeads: leadsFormateados.slice(0, 5),
      metricasCuentas: metricasCuentas,
      otrosEstados: metricas.otrosEstados || {},
      fechaActualizacion: new Date().toLocaleString('es-ES'),
      resumen: `📊 ${leadsFiltradosUnicos.length} leads únicos para ${nombreUsuario}`
    };

  } catch (error) {
    console.error('❌ Error en getEstadisticasUsuarioConFiltros:', error);
    return {
      usuarioId: usuarioId,
      usuarioNombre: 'Error',
      total: 0,
      estados: {},
      porcentajes: {},
      todosLosLeads: [],
      ultimosLeads: [],
      metricasCuentas: {},
      otrosEstados: {},
      fechaActualizacion: new Date().toLocaleString('es-ES'),
      resumen: 'Error al obtener estadísticas'
    };
  }
}

/**
 * Estadísticas generales con filtros
 */
function getEstadisticasGeneralesConFiltros(filtros = {}) {
  try {
    // 1. Obtener lista de mails válidos desde 'Usuarios'
    const usuarios = Usuarios.getUsuarios();
    const mailsValidos = new Set(usuarios.map(u => u.id.toLowerCase()));

    // 2. Obtener todos los leads
    const todosLosLeads = Leads.getLeads();

    // 3. Filtrar leads por asesor válido y fecha, y DEDUPLICAR por LeadID
    const uniqueLeadsMap = new Map();
    todosLosLeads.forEach(lead => {
      const idUsu = lead.idUsuario ? lead.idUsuario.toString().trim().toLowerCase() : "";

      // Solo si el asesor es válido
      if (mailsValidos.has(idUsu)) {
        // Filtrar por fecha
        let matchAnio = !filtros.anio || filtros.anio === "" || Fechas.getAnio(lead.fecha) === parseInt(filtros.anio);
        let matchMes = !filtros.mes || filtros.mes === "" || Fechas.getMes(lead.fecha) === parseInt(filtros.mes);

        if (matchAnio && matchMes) {
          const lid = lead.leadId || ("TEMP_" + Math.random());
          uniqueLeadsMap.set(lid, lead);
        }
      }
    });

    const leadsUnicos = Array.from(uniqueLeadsMap.values());
    const metricas = Metricas.calcular(leadsUnicos);

    console.log(`✅ getEstadisticasGeneralesConFiltros: ${leadsUnicos.length} leads únicos, ${usuarios.length} asesores válidos.`);

    return {
      total: leadsUnicos.length,
      estados: metricas.estados,
      usuariosActivos: usuarios.length, // Total de mails únicos en la hoja Usuarios
      fechaActualizacion: new Date().toLocaleString('es-ES')
    };

  } catch (error) {
    console.error('❌ Error en getEstadisticasGeneralesConFiltros:', error);
    return {
      total: 0,
      estados: {},
      usuariosActivos: 0,
      fechaActualizacion: new Date().toLocaleString('es-ES')
    };
  }
}

/**
 * Obtiene datos completos del dashboard con filtros
 */
function getDashboardData(filtros = {}) {
  try {
    console.log(`🔄 [API V2] getDashboardData con filtros:`, filtros);

    // 1. Obtener datos base
    const todosLosLeads = Leads.getLeads();
    const usuarios = Usuarios.getUsuarios();

    // 2. Aplicar filtros
    const leadsFiltrados = Filtros.aplicar(todosLosLeads, filtros);

    // 3. Calcular métricas generales
    const metricasGenerales = Metricas.calcular(leadsFiltrados);

    // 4. Si hay usuario específico, calcular métricas para ese usuario
    let metricasUsuario = null;
    let leadsUsuario = [];
    let leadsUsuarioFiltrados = [];

    if (filtros.usuarioId) {
      leadsUsuario = Leads.getLeadsPorUsuario(filtros.usuarioId);
      leadsUsuarioFiltrados = Filtros.aplicar(leadsUsuario, filtros);
      metricasUsuario = Metricas.calcular(leadsUsuarioFiltrados, {
        usuarioId: filtros.usuarioId
      });
    }

    // 5. Obtener años disponibles para filtros
    const aniosDisponibles = Fechas.getAniosDisponibles(todosLosLeads);

    return {
      success: true,
      timestamp: new Date().toLocaleString('es-ES'),
      filtrosAplicados: filtros,

      // Datos básicos
      usuarios: usuarios,
      totalUsuarios: usuarios.length,

      // Métricas
      metricasGenerales: metricasGenerales,
      metricasUsuario: metricasUsuario,

      // Datos para filtros
      aniosDisponibles: aniosDisponibles,
      estadosDisponibles: Object.keys(Leads.MAPA_ESTADOS),

      // Leads formateados (si hay usuario)
      leadsUsuario: (function () {
        if (!filtros.usuarioId) return [];
        try { var indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { indexMap = {}; }
        return Leads.formatearLeads(leadsUsuarioFiltrados || [], indexMap);
      })(),
    };

  } catch (error) {
    console.error(`❌ [API V2] Error en getDashboardData:`, error);
    return {
      success: false,
      error: error.toString(),
      timestamp: new Date().toLocaleString('es-ES'),
      usuarios: [],
      totalUsuarios: 0,
      metricasGenerales: {},
      metricasUsuario: {},
      aniosDisponibles: [],
      estadosDisponibles: [],
      leadsUsuario: []
    };
  }
}

// Global wrappers for triggers
function buildUltimaActividadIndex() {
  return Usuarios.buildUltimaActividadIndex();
}

// ============================================
// MÓDULO: DEBUG
// ============================================

function getDebugInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    leads: { headers: [], row1: [], row2: [] },
    usuarios: { headers: [], row1: [], row2: [] }
  };

  var sheet = ss.getSheetByName('Leads');
  if (sheet) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    result.leads.headers = headers.map(function (c, i) { return i + ': ' + c; });

    var sampleData = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    result.leads.row1 = sampleData.map(function (c, i) { return i + ': ' + (c && c.toString().substring(0, 20)); });
  }

  var sheetU = ss.getSheetByName('Usuarios');
  if (sheetU) {
    var rangeU = sheetU.getDataRange().getValues();
    if (rangeU.length > 0) result.usuarios.headers = rangeU[0];
    if (rangeU.length > 1) result.usuarios.row1 = rangeU[1];
  }

  var sheetUA = ss.getSheetByName('Ultima_Actividad');
  if (sheetUA) {
    var dataUA = sheetUA.getRange(1, 1, Math.min(sheetUA.getLastRow(), 10), 5).getValues();
    result.ultimaActividad = {
      data: dataUA.map(function (row) { return row.join(' | '); })
    };
  }

  return result;
}

function verificarCruceDeDatos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lSheet = ss.getSheetByName('Leads');
    var uaSheet = ss.getSheetByName('Ultima_Actividad');

    if (!lSheet || !uaSheet) return "Error: No se encuentran las hojas";

    var leadsIdsRaw = lSheet.getRange(2, 1, Math.min(lSheet.getLastRow() - 1, 20), 1).getValues().flat();
    var uaIdsAll = uaSheet.getRange(2, 1, uaSheet.getLastRow() - 1, 1).getValues().flat().map(function (id) {
      return id ? id.toString().trim().replace(/\.0$/, "") : "";
    });

    var report = "--- DIAGNÓSTICO DE CRUCE REAL ---\n";
    report += "Buscando los primeros 20 leads en TODA la hoja de Ultima Actividad...\n\n";

    leadsIdsRaw.forEach(function (lid, i) {
      var normL = lid ? lid.toString().trim().replace(/\.0$/, "") : "VACÍO";
      var match = uaIdsAll.indexOf(normL) !== -1;
      report += "ID [" + lid + "] -> Corregido: [" + normL + "] -> ¿Está en Actividad?: " + (match ? "SÍ ✅" : "NO ❌") + "\n";
    });

    return report;
  } catch (e) {
    return "Error en diagnóstico: " + e.toString();
  }
}

// ============================================
// MÓDULO: FECHAS
// ============================================

var Fechas = (function () {
  function formatear(fecha) {
    if (!fecha) return 'Sin fecha';
    try {
      if (fecha instanceof Date) {
        if (!isNaN(fecha.getTime())) return fecha.toLocaleDateString('es-ES');
      }
      if (typeof fecha === 'string' || typeof fecha === 'number') {
        if (typeof fecha === 'number' && fecha > 25569) {
          const date = new Date((fecha - 25569) * 86400 * 1000);
          if (!isNaN(date.getTime())) return date.toLocaleDateString('es-ES');
        }
        const date = new Date(fecha);
        if (!isNaN(date.getTime())) return date.toLocaleDateString('es-ES');
      }
      return fecha.toString();
    } catch (e) { return 'Sin fecha'; }
  }

  function getAnio(fecha) {
    try {
      if (!fecha) return null;
      let date;
      if (fecha instanceof Date) date = fecha;
      else if (typeof fecha === 'number') date = new Date((fecha - 25569) * 86400 * 1000);
      else date = new Date(fecha);
      return !isNaN(date.getTime()) ? date.getFullYear() : null;
    } catch (e) { return null; }
  }

  function getMes(fecha) {
    try {
      if (!fecha) return null;
      let date;
      if (fecha instanceof Date) date = fecha;
      else if (typeof fecha === 'number') date = new Date((fecha - 25569) * 86400 * 1000);
      else date = new Date(fecha);
      return !isNaN(date.getTime()) ? date.getMonth() + 1 : null;
    } catch (e) { return null; }
  }

  function coincidePeriodo(fecha, anio, mes) {
    if (!fecha) return false;
    const fechaAnio = getAnio(fecha);
    const fechaMes = getMes(fecha);
    if (anio && fechaAnio !== parseInt(anio)) return false;
    if (mes && fechaMes !== parseInt(mes)) return false;
    return true;
  }

  function getAniosDisponibles(leads) {
    const anios = leads.map(lead => getAnio(lead.fecha)).filter(anio => anio !== null);
    return [...new Set(anios)].sort((a, b) => b - a);
  }

  function getNombreMes(mes) {
    const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    return meses[mes - 1] || mes;
  }

  return { formatear, getAnio, getMes, coincidePeriodo, getAniosDisponibles, getNombreMes };
})();

// ============================================
// MÓDULO: FILTROS
// ============================================

var Filtros = (function () {
  function aplicar(leads, filtros = {}) {
    if (!leads || !leads.length) return [];
    let resultado = [...leads];
    if (filtros.anio) resultado = filtrarPorAnio(resultado, filtros.anio);
    if (filtros.mes) resultado = filtrarPorMes(resultado, filtros.mes);
    if (filtros.estados && filtros.estados.length > 0) resultado = filtrarPorEstados(resultado, filtros.estados);
    if (filtros.usuarioId) resultado = filtrarPorUsuario(resultado, filtros.usuarioId);
    return resultado;
  }

  function filtrarPorAnio(leads, anio) {
    const anioNum = parseInt(anio);
    if (isNaN(anioNum)) return leads;
    return leads.filter(lead => Fechas.getAnio(lead.fecha) === anioNum);
  }

  function filtrarPorMes(leads, mes) {
    const mesNum = parseInt(mes);
    if (isNaN(mesNum) || mesNum < 1 || mesNum > 12) return leads;
    return leads.filter(lead => Fechas.getMes(lead.fecha) === mesNum);
  }

  function filtrarPorEstados(leads, estados) {
    if (!estados || !estados.length) return leads;
    return leads.filter(lead => {
      const estadoClasificado = Leads.clasificarEstado(lead.estado);
      return estados.includes(estadoClasificado);
    });
  }

  function filtrarPorUsuario(leads, usuarioId) {
    if (!usuarioId) return leads;
    const usuarioIdStr = usuarioId.toString().trim().toLowerCase();
    return leads.filter(lead => lead.idUsuario && lead.idUsuario.toLowerCase() === usuarioIdStr);
  }

  function filtrarPorRangoFechas(leads, fechaInicio, fechaFin) {
    if (!fechaInicio && !fechaFin) return leads;
    return leads.filter(lead => {
      if (!lead.fecha) return false;
      const fechaLead = new Date(lead.fecha);
      if (isNaN(fechaLead.getTime())) return false;
      if (fechaInicio && fechaLead < fechaInicio) return false;
      if (fechaFin && fechaLead > fechaFin) return false;
      return true;
    });
  }

  return { aplicar, filtrarPorAnio, filtrarPorMes, filtrarPorEstados, filtrarPorUsuario, filtrarPorRangoFechas };
})();

// ============================================
// MÓDULO: LEADS
// ============================================

var Leads = (function () {
  const CONFIG = {
    COL_ID_LEAD: 0,      // Columna A (LeadID)
    COL_FECHA: 1,        // Columna B (Fecha Registro)
    COL_ID_USUARIO: 2,   // Columna C (Asesor - Mail)
    COL_NOMBRE: 4,       // Columna E
    COL_APELLIDO: 5,     // Columna F
    COL_CREADO_POR: 30,  // Columna AF (Probablemente)
    COL_K: 33,           // Columna AH
    COL_KV: 34,          // Columna AI
    COL_ESTADO: 37,      // Columna AL (Estado)
    COL_ULTIMA_GESTION: 38 // Columna AM
  };

  const MAPA_ESTADOS = {
    'NUEVO': ['NUEVO', 'NUEVA'], 'ABIERTO': ['ABIERTO', 'ABIERTA'], 'CONVERTIDO': ['CONVERTIDO', 'CONVERTIDA'],
    'ACTIVO': ['ACTIVO', 'ACTIVA'], 'OPERANDO': ['OPERANDO'],
    'DESCARTADO': ['DESCARTADO', 'DESCARTADA', 'PERDIDO', 'PERDIDA'], 'OTROS': []
  };

  function getLeads() {
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'LEADS_CACHE_V4'; // Bump version
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Leads');
      if (!sheet) throw new Error('No se encontró la hoja "Leads"');

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow < 2) return [];

      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      let colLeadId = CONFIG.COL_ID_LEAD;
      if (headers[colLeadId] && headers[colLeadId].toString().toUpperCase().indexOf('LEAD') === -1) {
        const foundIndex = headers.findIndex(h => h.toString().toUpperCase().indexOf('LEAD') !== -1);
        if (foundIndex !== -1) colLeadId = foundIndex;
      }

      const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

      // 1. Obtener usuarios válidos para filtrar inmediatamente
      const usuariosValidos = Usuarios.getUsuarios();
      const mailsValidos = new Set(usuariosValidos.map(u => u.id.toLowerCase()));

      // 2. Mapa para deduplicación por LeadID
      const uniqueLeadsMap = new Map();

      data.forEach((row, index) => {
        const idUsu = row[CONFIG.COL_ID_USUARIO] ? row[CONFIG.COL_ID_USUARIO].toString().trim().toLowerCase() : '';

        // Solo si el asesor está en la hoja Usuarios
        if (mailsValidos.has(idUsu)) {
          const valRaw = row[colLeadId];
          const lid = (valRaw !== undefined && valRaw !== null) ? valRaw.toString().trim().replace(/\.0$/, "") : '';

          if (lid) {
            // Guardar en el mapa (si ya existe, se pisa con el último, cumpliendo el requisito de valores únicos)
            uniqueLeadsMap.set(lid, {
              idUsuario: idUsu,
              estado: row[CONFIG.COL_ESTADO] ? row[CONFIG.COL_ESTADO].toString().trim() : '',
              leadId: lid,
              nombre: row[CONFIG.COL_NOMBRE] || '',
              apellido: row[CONFIG.COL_APELLIDO] || '',
              fecha: row[CONFIG.COL_FECHA],
              creadoPor: row[CONFIG.COL_CREADO_POR] ? row[CONFIG.COL_CREADO_POR].toString().trim() : '',
              k: row[CONFIG.COL_K] || '',
              kv: row[CONFIG.COL_KV] || '',
              ultimaGestion: row[CONFIG.COL_ULTIMA_GESTION] || null,
              _filaNumero: index + 2
            });
          }
        }
      });

      const leads = Array.from(uniqueLeadsMap.values());
      console.log(`✅ Leads.getLeads: ${leads.length} leads únicos y válidos obtenidos`);

      try { cache.put(cacheKey, JSON.stringify(leads), 600); } catch (e) { }
      return leads;
    } catch (error) { throw error; }
  }

  function getLeadsPorUsuario(usuarioId) {
    const todosLosLeads = getLeads();
    const usuarioIdStr = usuarioId.toString().trim().toLowerCase();
    let usuarioNombreStr = "";
    try {
      const usuario = Usuarios.getUsuarioPorId(usuarioId);
      if (usuario && usuario.nombre) usuarioNombreStr = usuario.nombre.toString().trim().toLowerCase();
    } catch (e) { }
    return todosLosLeads.filter(lead => {
      if (!lead.idUsuario) return false;
      const val = lead.idUsuario.toString().trim().toLowerCase();
      return val === usuarioIdStr || (usuarioNombreStr && val === usuarioNombreStr);
    });
  }

  function clasificarEstado(estadoRaw) {
    if (!estadoRaw) return 'OTROS';
    const estado = estadoRaw.toString().trim().toUpperCase();
    for (const [categoria, palabras] of Object.entries(MAPA_ESTADOS)) {
      for (const palabra of palabras) { if (estado.includes(palabra)) return categoria; }
    }
    return 'OTROS';
  }

  function formatearLead(lead, indexMap) {
    const nombreCompleto = `${lead.nombre || ''} ${lead.apellido || ''}`.trim() || 'Sin nombre';
    let ultimaActividadStr = '—', ultimaActividadOrigen = '', ultimaActividadActor = '';
    const lid = lead.leadId ? lead.leadId.toString().trim().replace(/\.0$/, "") : '';
    if (indexMap && lid && indexMap[lid]) {
      const info = indexMap[lid];
      ultimaActividadStr = Fechas.formatear(info.ultimaISO);
      ultimaActividadOrigen = info.origen || '';
      ultimaActividadActor = info.actor || '';
    }
    return {
      id: lead.leadId !== undefined && lead.leadId !== null ? lead.leadId.toString() : 'N/A',
      nombre: nombreCompleto, fecha: Fechas.formatear(lead.fecha),
      fechaISO: lead.fecha instanceof Date ? lead.fecha.toISOString() : (new Date(lead.fecha).toISOString() || ''),
      estado: lead.estado || 'Sin estado', estadoClasificado: clasificarEstado(lead.estado),
      creadoPor: lead.creadoPor || null, k: lead.k || '', kv: lead.kv || '',
      ultimaActividad: ultimaActividadStr, ultimaActividadISO: (indexMap && lid && indexMap[lid]) ? indexMap[lid].ultimaISO : '',
      ultimaActividadOrigen: ultimaActividadOrigen, ultimaActividadActor: ultimaActividadActor
    };
  }

  function formatearLeads(leads, indexMap) {
    return leads.map(lead => formatearLead(lead, indexMap));
  }

  function getLeadsFormateadosPorUsuario(usuarioId) {
    const leads = getLeadsPorUsuario(usuarioId);
    var indexMap = {};
    try { indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { }
    return formatearLeads(leads, indexMap).sort((a, b) => {
      const fechaA = a.fecha === 'Sin fecha' ? 0 : new Date(a.fecha).getTime() || 0;
      const fechaB = b.fecha === 'Sin fecha' ? 0 : new Date(b.fecha).getTime() || 0;
      return fechaB - fechaA;
    });
  }

  function contarEstados(leads) {
    const conteo = { 'NUEVO': 0, 'ABIERTO': 0, 'CONVERTIDO': 0, 'ACTIVO': 0, 'OPERANDO': 0, 'DESCARTADO': 0, 'OTROS': 0 };
    const otrosDetalles = {};
    leads.forEach(lead => {
      const estadoClasificado = clasificarEstado(lead.estado);
      if (conteo.hasOwnProperty(estadoClasificado)) conteo[estadoClasificado]++;
      else { conteo['OTROS']++; otrosDetalles[lead.estado] = (otrosDetalles[lead.estado] || 0) + 1; }
    });
    return { conteo, otrosDetalles };
  }

  function calcularPorcentajes(conteo, total) {
    const porcentajes = {};
    Object.keys(conteo).forEach(key => { porcentajes[key] = total > 0 ? Math.round((conteo[key] / total) * 100) : 0; });
    return porcentajes;
  }

  return { CONFIG, MAPA_ESTADOS, getLeads, getLeadsPorUsuario, getLeadsFormateadosPorUsuario, clasificarEstado, formatearLead, formatearLeads, contarEstados, calcularPorcentajes };
})();

// ============================================
// MÓDULO: MÉTRICAS
// ============================================

var Metricas = (function () {
  function calcular(leads, opciones = {}) {
    if (!leads || !leads.length) return { estados: Leads.contarEstados([]).conteo, porcentajes: {}, total: 0, metricasComerciales: {}, metricasCuentas: {} };
    const { conteo, otrosDetalles } = Leads.contarEstados(leads);
    const porcentajes = Leads.calcularPorcentajes(conteo, leads.length);
    const metricasCuentas = calcularMetricasCuentas(leads);
    const metricasComerciales = opciones.usuarioId ? calcularMetricasComerciales(leads, opciones.usuarioId) : {};
    return { estados: conteo, porcentajes, total: leads.length, otrosEstados: otrosDetalles, metricasCuentas, metricasComerciales };
  }

  function calcularMetricasCuentas(leads) {
    const totalCuentas = leads.length;
    const cuentasConRelevancia = leads.map(lead => ({ ...lead, relevancia: calcularRelevancia(lead) }));
    const cuentasConGestion = leads.filter(lead => lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== '').length;
    const porcentajeGestion = totalCuentas > 0 ? Math.round((cuentasConGestion / totalCuentas) * 100) : 0;
    const ranking = cuentasConRelevancia.sort((a, b) => b.relevancia.score - a.relevancia.score).slice(0, 10).map(lead => ({
      id: lead.leadId, nombre: `${lead.nombre || ''} ${lead.apellido || ''}`.trim(), score: lead.relevancia.score, nivel: lead.relevancia.nivel, ultimaGestion: lead.ultimaGestion
    }));
    return {
      totalCuentas, cuentasConGestion, porcentajeGestion, rankingRelevancia: ranking,
      promedioScoreRelevancia: totalCuentas > 0 ? Math.round(cuentasConRelevancia.reduce((sum, c) => sum + c.relevancia.score, 0) / totalCuentas) : 0
    };
  }

  function calcularRankingRelevancia(leads) {
    return leads.map(lead => ({
      id: lead.leadId, nombre: `${lead.nombre || ''} ${lead.apellido || ''}`.trim(),
      score: calcularRelevancia(lead).score, nivel: calcularRelevancia(lead).nivel, ultimaGestion: lead.ultimaGestion, estado: lead.estado
    })).sort((a, b) => b.score - a.score);
  }

  function calcularRelevancia(lead) {
    let score = 0; const razones = [];
    if (lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== '') { score += 20; razones.push('tiene_ultima_gestion'); }
    const estadoClasificado = Leads.clasificarEstado(lead.estado);
    if (['ACTIVO', 'OPERANDO', 'ABIERTO'].includes(estadoClasificado)) { score += 15; razones.push('estado_activo'); }
    else if (['CONVERTIDO'].includes(estadoClasificado)) { score += 25; razones.push('estado_convertido'); }
    else if (['NUEVO'].includes(estadoClasificado)) { score += 10; razones.push('estado_nuevo'); }
    if (lead.fecha) {
      const fechaLead = new Date(lead.fecha);
      if (!isNaN(fechaLead.getTime())) {
        const diasDesdeAsignacion = Math.floor((new Date() - fechaLead) / (1000 * 60 * 60 * 24));
        if (diasDesdeAsignacion <= 30) { score += 10; razones.push('asignacion_reciente'); }
      }
    }
    let nivel = 'BAJA'; if (score >= 40) nivel = 'ALTA'; else if (score >= 20) nivel = 'MEDIA';
    return { score, nivel, razones };
  }

  function calcularMetricasComerciales(leads, usuarioId) {
    const totalLeads = leads.length;
    const cuentasGestionadas = leads.filter(lead => lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== '').length;
    const porcentajeGestion = totalLeads > 0 ? Math.round((cuentasGestionadas / totalLeads) * 100) : 0;
    const sociedadesCreadas = leads.filter(lead => {
      if (!lead.creadoPor) return false;
      const creadoPorStr = lead.creadoPor.toString().trim().toLowerCase();
      const usuarioIdStr = usuarioId.toString().trim().toLowerCase();
      return creadoPorStr === usuarioIdStr || creadoPorStr.includes(usuarioIdStr);
    }).length;
    const { conteo } = Leads.contarEstados(leads);
    return { totalLeads, cuentasGestionadas, porcentajeGestion, sociedadesCreadas, distribucionEstados: conteo, promedioGestionDiaria: totalLeads > 0 ? (cuentasGestionadas / 30).toFixed(1) : 0 };
  }

  return { calcular, calcularMetricasCuentas, calcularMetricasComerciales, calcularRelevancia, calcularRankingRelevancia };
})();

// ============================================
// MÓDULO: USUARIOS
// ============================================

var Usuarios = (function () {
  function getUsuarios() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Usuarios');
      if (!sheet) return [];
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

      const uniqueUsersMap = new Map();
      data.forEach(row => {
        const mail = row[0] ? row[0].toString().trim() : '';
        const nombre = row[1] ? row[1].toString().trim() : (mail.split('@')[0] || 'Sin Nombre');
        if (mail && !uniqueUsersMap.has(mail.toLowerCase())) {
          uniqueUsersMap.set(mail.toLowerCase(), { id: mail, nombre: nombre });
        }
      });

      return Array.from(uniqueUsersMap.values());
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
      function normalizarId(val) { if (val === null || val === undefined) return ""; return val.toString().trim().replace(/\.0$/, ""); }
      var configs = [{ name: 'Tareas', leadCol: 2, dateCol: 6, actorCol: 9 }, { name: 'Agenda', leadCol: 2, dateCol: 10, actorCol: 9 }, { name: 'Comentarios', leadCol: 2, dateCol: 5, actorCol: 7 }];
      configs.forEach(function (conf) {
        var sheet = ss.getSheetByName(conf.name); if (!sheet) return;
        var lastRow = sheet.getLastRow(); if (lastRow < 2) return;
        var leadVals = sheet.getRange(2, conf.leadCol, lastRow - 1, 1).getValues();
        var dateVals = sheet.getRange(2, conf.dateCol, lastRow - 1, 1).getValues();
        var actorVals = sheet.getRange(2, conf.actorCol, lastRow - 1, 1).getValues();
        for (var i = 0; i < leadVals.length; i++) {
          var lid = normalizarId(leadVals[i][0]); if (!lid) continue;
          var d = parsePossiblySheetDate(dateVals[i][0]); if (!d) continue;
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

      // Columnas 0-indexed: LeadID es 0, Mail es depende la hoja
      function check(name, uCol, dCol) {
        var s = ss.getSheetByName(name); if (!s) return;
        var last = s.getLastRow(); if (last < 2) return;
        var data = s.getRange(2, 1, last - 1, Math.max(uCol, dCol) + 1).getValues();
        for (var i = 0; i < data.length; i++) {
          if (data[i][uCol] && data[i][uCol].toString().trim().toLowerCase() === uid) {
            var d = parsePossiblySheetDate(data[i][dCol]);
            if (d && (!maxDate || d.getTime() > maxDate.getTime())) maxDate = d;
          }
        }
      }

      // Ajuste de columnas según descripción (Leads: C=2 es asesor, B=1 es fecha)
      check('Leads', 2, 1);
      check('Tareas', 8, 5); // Tareas: Col I (8) es actor, Col F (5) es fecha
      check('Agenda', 8, 9); // Agenda: Col I (8) es actor, Col J (9) es fecha
      check('Comentarios', 6, 4); // Comentarios: Col G (6) es actor, Col E (4) es fecha

      return maxDate ? Fechas.formatear(maxDate) : "Sin actividad reciente";
    } catch (e) {
      console.error('Error en getUltimaActividadAsociado:', e);
      return "Error";
    }
  }

  return { getUsuarios, getNombrePorId, getUsuarioPorId, getEstadisticas, buildUltimaActividadIndex, getUltimaActividadMap, getUltimaActividadAsociado };
})();

// ==========================================
// CONFIGURACIÓN DE DISPARADORES (TRIGGERS)
// ==========================================

function instalarDisparadores() {
  try {
    var functionName = 'buildUltimaActividadIndex';
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === functionName) ScriptApp.deleteTrigger(triggers[i]);
    }
    ScriptApp.newTrigger(functionName).timeBased().everyHours(1).create();
    console.log('✅ Disparador instalado: ' + functionName + ' (Cada 1 hora)');
  } catch (e) { console.error('❌ Error instalando disparadores:', e); }
}