// ============================================
// MÓDULO: FECHAS
// ============================================

var Fechas = (function() {
  
  /**
   * Formatea una fecha desde cualquier formato
   * @param {any} fecha - Fecha en cualquier formato
   * @returns {string} Fecha formateada DD/MM/YYYY
   */
  function formatear(fecha) {
    if (!fecha) return 'Sin fecha';
    
    try {
      // Si es objeto Date
      if (fecha instanceof Date) {
        if (!isNaN(fecha.getTime())) {
          return fecha.toLocaleDateString('es-ES');
        }
      }
      
      // Si es string o número
      if (typeof fecha === 'string' || typeof fecha === 'number') {
        // Timestamp de Google Sheets (días desde 1900)
        if (typeof fecha === 'number' && fecha > 25569) {
          const date = new Date((fecha - 25569) * 86400 * 1000);
          if (!isNaN(date.getTime())) {
            return date.toLocaleDateString('es-ES');
          }
        }
        
        // Parsear string
        const date = new Date(fecha);
        if (!isNaN(date.getTime())) {
          return date.toLocaleDateString('es-ES');
        }
      }
      
      return fecha.toString();
    } catch (e) {
      return 'Sin fecha';
    }
  }
  
  /**
   * Extrae el año de una fecha
   * @param {any} fecha 
   * @returns {number|null} Año o null
   */
  function getAnio(fecha) {
    try {
      if (!fecha) return null;
      
      let date;
      if (fecha instanceof Date) {
        date = fecha;
      } else if (typeof fecha === 'number') {
        date = new Date((fecha - 25569) * 86400 * 1000);
      } else {
        date = new Date(fecha);
      }
      
      return !isNaN(date.getTime()) ? date.getFullYear() : null;
    } catch (e) {
      return null;
    }
  }
  
  /**
   * Extrae el mes de una fecha (1-12)
   * @param {any} fecha 
   * @returns {number|null} Mes o null
   */
  function getMes(fecha) {
    try {
      if (!fecha) return null;
      
      let date;
      if (fecha instanceof Date) {
        date = fecha;
      } else if (typeof fecha === 'number') {
        date = new Date((fecha - 25569) * 86400 * 1000);
      } else {
        date = new Date(fecha);
      }
      
      return !isNaN(date.getTime()) ? date.getMonth() + 1 : null;
    } catch (e) {
      return null;
    }
  }
  
  /**
   * Compara si dos fechas están en el mismo período
   * @param {any} fecha - Fecha a evaluar
   * @param {number|null} anio - Año a comparar
   * @param {number|null} mes - Mes a comparar
   * @returns {boolean} True si coincide
   */
  function coincidePeriodo(fecha, anio, mes) {
    if (!fecha) return false;
    
    const fechaAnio = getAnio(fecha);
    const fechaMes = getMes(fecha);
    
    if (anio && fechaAnio !== parseInt(anio)) return false;
    if (mes && fechaMes !== parseInt(mes)) return false;
    
    return true;
  }
  
  /**
   * Obtiene lista de años disponibles en los leads
   * @param {Array} leads - Array de leads
   * @returns {Array} Años únicos ordenados
   */
  function getAniosDisponibles(leads) {
    const anios = leads
      .map(lead => getAnio(lead.fecha))
      .filter(anio => anio !== null);
    
    return [...new Set(anios)].sort((a, b) => b - a);
  }
  
  /**
   * Obtiene nombre del mes en español
   * @param {number} mes - Número de mes (1-12)
   * @returns {string} Nombre del mes
   */
  function getNombreMes(mes) {
    const meses = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ];
    return meses[mes - 1] || mes;
  }
  
  // API Pública
  return {
    formatear: formatear,
    getAnio: getAnio,
    getMes: getMes,
    coincidePeriodo: coincidePeriodo,
    getAniosDisponibles: getAniosDisponibles,
    getNombreMes: getNombreMes
  };
  
})();