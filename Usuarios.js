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