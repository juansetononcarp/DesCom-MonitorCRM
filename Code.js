// ============================================
// ORQUESTADOR PRINCIPAL - DASHBOARD DE LEADS
// ============================================

// ========== SERVIDOR WEB ==========

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Monitor CRM v1.3.2 [TEST]')
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
function getAniosDisponibles() {
  try {
    const todosLosLeads = Leads.getLeads();
    return Fechas.getAniosDisponibles(todosLosLeads);
  } catch (error) {
    console.error('Error en getAniosDisponibles:', error);
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

    // 2. Aplicar filtros de fecha
    let leadsFiltrados = [...leads];

    if (filtros.anio) {
      leadsFiltrados = Filtros.filtrarPorAnio(leadsFiltrados, filtros.anio);
    }
    if (filtros.mes) {
      leadsFiltrados = Filtros.filtrarPorMes(leadsFiltrados, filtros.mes);
    }

    // 3. Calcular métricas
    const metricas = Metricas.calcular(leadsFiltrados, { usuarioId });
    const metricasCuentas = Metricas.calcularMetricasCuentas(leadsFiltrados);

    // 4. Formatear leads
    var indexMap = {};
    try { indexMap = Usuarios.getUltimaActividadMap() || {}; } catch (e) { indexMap = {}; }
    const todosLosLeads = leadsFiltrados
      .map(lead => Leads.formatearLead(lead, indexMap))
      .sort((a, b) => {
        const fechaA = a.fecha === 'Sin fecha' ? 0 : new Date(a.fecha).getTime() || 0;
        const fechaB = b.fecha === 'Sin fecha' ? 0 : new Date(b.fecha).getTime() || 0;
        return fechaB - fechaA;
      });

    const nombreUsuario = Usuarios.getNombrePorId(usuarioId);

    return {
      usuarioId: usuarioId,
      usuarioNombre: nombreUsuario,
      total: leadsFiltrados.length,
      estados: metricas.estados,
      porcentajes: metricas.porcentajes,
      todosLosLeads: todosLosLeads,
      ultimosLeads: todosLosLeads.slice(0, 5),
      metricasCuentas: metricasCuentas,
      otrosEstados: metricas.otrosEstados || {},
      fechaActualizacion: new Date().toLocaleString('es-ES'),
      resumen: `📊 ${leadsFiltrados.length} leads para ${nombreUsuario}`
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
    const todosLosLeads = Leads.getLeads();

    let leadsFiltrados = [...todosLosLeads];

    if (filtros.anio) {
      leadsFiltrados = Filtros.filtrarPorAnio(leadsFiltrados, filtros.anio);
    }
    if (filtros.mes) {
      leadsFiltrados = Filtros.filtrarPorMes(leadsFiltrados, filtros.mes);
    }

    const metricas = Metricas.calcular(leadsFiltrados);
    const usuarios = Usuarios.getUsuarios();

    console.log(`✅ getEstadisticasGeneralesConFiltros: ${leadsFiltrados.length} leads filtrados para métricas globales.`);

    return {
      total: leadsFiltrados.length,
      estados: metricas.estados,
      usuariosActivos: usuarios.length,
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
 * @param {Object} filtros - { anio, mes, estados, usuarioId }
 * @returns {Object} Datos completos del dashboard
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

// ========== FUNCIONES DE PRUEBA ==========

function testAllModules() {
  console.log("🧪 TEST: Todos los módulos");
  console.log("==========================");

  try {
    // Test Usuarios
    const usuarios = Usuarios.getUsuarios();
    console.log(`✅ Usuarios: ${usuarios.length} usuarios`);

    // Test Leads
    const leads = Leads.getLeads();
    console.log(`✅ Leads: ${leads.length} leads`);

    // Test Fechas
    if (leads.length > 0) {
      const fechaFormateada = Fechas.formatear(leads[0].fecha);
      console.log(`✅ Fechas: ${fechaFormateada}`);
    }

    // Test Filtros
    const filtrados = Filtros.aplicar(leads, { anio: 2024 });
    console.log(`✅ Filtros: ${filtrados.length} leads en 2024`);

    // Test Metricas
    const metricas = Metricas.calcular(leads);
    console.log(`✅ Metricas: ${metricas.total} leads procesados`);
    console.log(`   Estados:`, metricas.estados);

    // Test API V2
    const dashboardData = getDashboardData();
    console.log(`✅ API V2: Dashboard data OK`);

    console.log("==========================");
    console.log("🎉 Todos los módulos funcionan correctamente!");

  } catch (error) {
    console.error("❌ Error en test:", error);
  }
}

// ==========================================
// CONFIGURACIÓN DE DISPARADORES (TRIGGERS)
// ==========================================

/**
 * Instala el disparador automático para actualizar el índice de actividad.
 * Ejecutar esta función UNA VEZ manualmente.
 */
function instalarDisparadores() {
  try {
    var functionName = 'buildUltimaActividadIndex';

    // 1. Eliminar disparadores existentes para evitar duplicados
    var triggers = ScriptApp.getProjectTriggers();
    var deleted = 0;
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(triggers[i]);
        deleted++;
      }
    }
    console.log('🗑️ Disparadores anteriores eliminados: ' + deleted);

    // 2. Crear nuevo disparador (cada 1 hora)
    // Nota: buildUltimaActividadIndex está en Usuarios.js.
    // Al ser un proyecto de Apps Script, las funciones top-level son globales.
    // Sin embargo, buildUltimaActividadIndex está DENTRO del módulo Usuarios (closure)?
    // REVISIÓN: Usuarios.js define var Usuarios = (function(){...}) pero 
    // buildUltimaActividadIndex está DEFINIDA FUERA del IIFE en las últimas líneas de Usuarios.js
    // Así que es accesible globalmente.

    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyHours(1)
      .create();

    console.log('✅ Disparador instalado: ' + functionName + ' (Cada 1 hora)');

  } catch (e) {
    console.error('❌ Error instalando disparadores:', e);
  }
}