// ============================================
// MÓDULO: FILTROS
// ============================================

var Filtros = (function() {
  
  /**
   * Aplica filtros a un array de leads
   * @param {Array} leads - Array de leads
   * @param {Object} filtros - Objeto con filtros
   * @returns {Array} Leads filtrados
   */
  function aplicar(leads, filtros = {}) {
    if (!leads || !leads.length) return [];
    
    let resultado = [...leads];
    
    // Filtro por año
    if (filtros.anio) {
      resultado = filtrarPorAnio(resultado, filtros.anio);
    }
    
    // Filtro por mes
    if (filtros.mes) {
      resultado = filtrarPorMes(resultado, filtros.mes);
    }
    
    // Filtro por estados
    if (filtros.estados && filtros.estados.length > 0) {
      resultado = filtrarPorEstados(resultado, filtros.estados);
    }
    
    // Filtro por usuario
    if (filtros.usuarioId) {
      resultado = filtrarPorUsuario(resultado, filtros.usuarioId);
    }
    
    return resultado;
  }
  
  /**
   * Filtra leads por año
   * @param {Array} leads - Array de leads
   * @param {number|string} anio - Año a filtrar
   * @returns {Array} Leads filtrados
   */
  function filtrarPorAnio(leads, anio) {
    const anioNum = parseInt(anio);
    if (isNaN(anioNum)) return leads;
    
    return leads.filter(lead => {
      const fechaAnio = Fechas.getAnio(lead.fecha);
      return fechaAnio === anioNum;
    });
  }
  
  /**
   * Filtra leads por mes
   * @param {Array} leads - Array de leads
   * @param {number|string} mes - Mes a filtrar (1-12)
   * @returns {Array} Leads filtrados
   */
  function filtrarPorMes(leads, mes) {
    const mesNum = parseInt(mes);
    if (isNaN(mesNum) || mesNum < 1 || mesNum > 12) return leads;
    
    return leads.filter(lead => {
      const fechaMes = Fechas.getMes(lead.fecha);
      return fechaMes === mesNum;
    });
  }
  
  /**
   * Filtra leads por estados
   * @param {Array} leads - Array de leads
   * @param {Array} estados - Lista de estados permitidos
   * @returns {Array} Leads filtrados
   */
  function filtrarPorEstados(leads, estados) {
    if (!estados || !estados.length) return leads;
    
    return leads.filter(lead => {
      const estadoClasificado = Leads.clasificarEstado(lead.estado);
      return estados.includes(estadoClasificado);
    });
  }
  
  /**
   * Filtra leads por usuario
   * @param {Array} leads - Array de leads
   * @param {string} usuarioId - ID del usuario
   * @returns {Array} Leads filtrados
   */
  function filtrarPorUsuario(leads, usuarioId) {
    if (!usuarioId) return leads;
    
    const usuarioIdStr = usuarioId.toString().trim().toLowerCase();
    return leads.filter(lead => 
      lead.idUsuario && lead.idUsuario.toLowerCase() === usuarioIdStr
    );
  }
  
  /**
   * Filtra leads por rango de fechas
   * @param {Array} leads - Array de leads
   * @param {Date} fechaInicio - Fecha inicial
   * @param {Date} fechaFin - Fecha final
   * @returns {Array} Leads filtrados
   */
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
  
  // API Pública
  return {
    aplicar: aplicar,
    filtrarPorAnio: filtrarPorAnio,
    filtrarPorMes: filtrarPorMes,
    filtrarPorEstados: filtrarPorEstados,
    filtrarPorUsuario: filtrarPorUsuario,
    filtrarPorRangoFechas: filtrarPorRangoFechas
  };
  
})();