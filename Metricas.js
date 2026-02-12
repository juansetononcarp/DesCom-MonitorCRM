// ============================================
// MÓDULO: MÉTRICAS
// ============================================

var Metricas = (function() {
  
  /**
   * Calcula todas las métricas para un conjunto de leads
   * @param {Array} leads - Array de leads
   * @param {Object} opciones - Opciones adicionales
   * @returns {Object} Todas las métricas calculadas
   */
  function calcular(leads, opciones = {}) {
    if (!leads || !leads.length) {
      return {
        estados: Leads.contarEstados([]).conteo,
        porcentajes: {},
        total: 0,
        metricasComerciales: {},
        metricasCuentas: {}
      };
    }
    
    // Métricas básicas
    const { conteo, otrosDetalles } = Leads.contarEstados(leads);
    const porcentajes = Leads.calcularPorcentajes(conteo, leads.length);
    
    // Métricas de cuentas
    const metricasCuentas = calcularMetricasCuentas(leads);
    
    // Métricas comerciales
    const metricasComerciales = opciones.usuarioId ? 
      calcularMetricasComerciales(leads, opciones.usuarioId) : {};
    
    return {
      estados: conteo,
      porcentajes: porcentajes,
      total: leads.length,
      otrosEstados: otrosDetalles,
      metricasCuentas: metricasCuentas,
      metricasComerciales: metricasComerciales
    };
  }
  
  /**
   * Calcula métricas de cuentas (relevancia, gestión)
   * @param {Array} leads - Array de leads
   * @returns {Object} Métricas de cuentas
   */
  function calcularMetricasCuentas(leads) {
    const totalCuentas = leads.length;
    
    // Calcular relevancia para cada cuenta
    const cuentasConRelevancia = leads.map(lead => ({
      ...lead,
      relevancia: calcularRelevancia(lead)
    }));
    
    // Estadísticas de gestión
    const cuentasConGestion = leads.filter(lead => 
      lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== ''
    ).length;
    
    const porcentajeGestion = totalCuentas > 0 ? 
      Math.round((cuentasConGestion / totalCuentas) * 100) : 0;
    
    // Ranking de relevancia
    const ranking = cuentasConRelevancia
      .sort((a, b) => b.relevancia.score - a.relevancia.score)
      .slice(0, 10)
      .map(lead => ({
        id: lead.leadId,
        nombre: `${lead.nombre || ''} ${lead.apellido || ''}`.trim(),
        score: lead.relevancia.score,
        nivel: lead.relevancia.nivel,
        ultimaGestion: lead.ultimaGestion
      }));
    
    return {
      totalCuentas: totalCuentas,
      cuentasConGestion: cuentasConGestion,
      porcentajeGestion: porcentajeGestion,
      rankingRelevancia: ranking,
      promedioScoreRelevancia: totalCuentas > 0 ?
        Math.round(cuentasConRelevancia.reduce((sum, c) => sum + c.relevancia.score, 0) / totalCuentas) : 0
    };
  }
  
  /**
   * Calcula ranking de relevancia para todas las cuentas
   * @param {Array} leads - Array de leads
   * @returns {Array} Ranking completo
   */
  function calcularRankingRelevancia(leads) {
    return leads
      .map(lead => ({
        id: lead.leadId,
        nombre: `${lead.nombre || ''} ${lead.apellido || ''}`.trim(),
        score: calcularRelevancia(lead).score,
        nivel: calcularRelevancia(lead).nivel,
        ultimaGestion: lead.ultimaGestion,
        estado: lead.estado
      }))
      .sort((a, b) => b.score - a.score);
  }
  
  /**
   * Calcula el score de relevancia para un lead
   * @param {Object} lead - Objeto lead
   * @returns {Object} Score y nivel de relevancia
   */
  function calcularRelevancia(lead) {
    let score = 0;
    const razones = [];
    
    // Tiene gestión reciente (último comentario)
    if (lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== '') {
      score += 20;
      razones.push('tiene_ultima_gestion');
    }
    
    // Estado activo o en proceso
    const estadoClasificado = Leads.clasificarEstado(lead.estado);
    if (['ACTIVO', 'OPERANDO', 'ABIERTO'].includes(estadoClasificado)) {
      score += 15;
      razones.push('estado_activo');
    } else if (['CONVERTIDO'].includes(estadoClasificado)) {
      score += 25;
      razones.push('estado_convertido');
    } else if (['NUEVO'].includes(estadoClasificado)) {
      score += 10;
      razones.push('estado_nuevo');
    }
    
    // Fecha reciente (últimos 30 días)
    if (lead.fecha) {
      const fechaLead = new Date(lead.fecha);
      if (!isNaN(fechaLead.getTime())) {
        const diasDesdeAsignacion = Math.floor((new Date() - fechaLead) / (1000 * 60 * 60 * 24));
        if (diasDesdeAsignacion <= 30) {
          score += 10;
          razones.push('asignacion_reciente');
        }
      }
    }
    
    // Determinar nivel
    let nivel = 'BAJA';
    if (score >= 40) nivel = 'ALTA';
    else if (score >= 20) nivel = 'MEDIA';
    
    return {
      score: score,
      nivel: nivel,
      razones: razones
    };
  }
  
  /**
   * Calcula métricas comerciales para un usuario específico
   * @param {Array} leads - Leads del comercial
   * @param {string} usuarioId - ID del comercial
   * @returns {Object} Métricas comerciales
   */
  function calcularMetricasComerciales(leads, usuarioId) {
    const totalLeads = leads.length;
    
    // Porcentaje de gestión (cuentas con comentario/agenda)
    const cuentasGestionadas = leads.filter(lead => 
      lead.ultimaGestion && lead.ultimaGestion.toString().trim() !== ''
    ).length;
    
    const porcentajeGestion = totalLeads > 0 ? 
      Math.round((cuentasGestionadas / totalLeads) * 100) : 0;
    
    // Sociedades creadas por el comercial
    const sociedadesCreadas = leads.filter(lead => {
      if (!lead.creadoPor) return false;
      const creadoPorStr = lead.creadoPor.toString().trim().toLowerCase();
      const usuarioIdStr = usuarioId.toString().trim().toLowerCase();
      return creadoPorStr === usuarioIdStr || creadoPorStr.includes(usuarioIdStr);
    }).length;
    
    // Distribución por estado
    const { conteo } = Leads.contarEstados(leads);
    
    return {
      totalLeads: totalLeads,
      cuentasGestionadas: cuentasGestionadas,
      porcentajeGestion: porcentajeGestion,
      sociedadesCreadas: sociedadesCreadas,
      distribucionEstados: conteo,
      promedioGestionDiaria: totalLeads > 0 ? 
        (cuentasGestionadas / 30).toFixed(1) : 0 // Últimos 30 días
    };
  }
  
  // API Pública
  return {
    calcular: calcular,
    calcularMetricasCuentas: calcularMetricasCuentas,
    calcularMetricasComerciales: calcularMetricasComerciales,
    calcularRelevancia: calcularRelevancia,
    calcularRankingRelevancia: calcularRankingRelevancia
  };
  
})();