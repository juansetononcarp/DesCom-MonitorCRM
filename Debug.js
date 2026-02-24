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

/**
 * Función técnica para diagnosticar por qué no coinciden los IDs
 */
function verificarCruceDeDatos() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var lSheet = ss.getSheetByName('Leads');
        var uaSheet = ss.getSheetByName('Ultima_Actividad');

        if (!lSheet || !uaSheet) return "Error: No se encuentran las hojas";

        var leadsIdsRaw = lSheet.getRange(2, 1, Math.min(lSheet.getLastRow(), 10), 1).getValues().flat();
        var uaIdsRaw = uaSheet.getRange(2, 1, Math.min(uaSheet.getLastRow(), 10), 1).getValues().flat();

        var report = "--- DIAGNÓSTICO DE CRUCE ---\n";
        report += "Leads (Primeros 10): " + JSON.stringify(leadsIdsRaw) + "\n";
        report += "Última Act. (Primeros 10): " + JSON.stringify(uaIdsRaw) + "\n\n";

        leadsIdsRaw.forEach(function (lid, i) {
            var normL = lid ? lid.toString().trim().replace(/\.0$/, "") : "VACÍO";
            var match = uaIdsRaw.some(function (uaid) {
                return (uaid ? uaid.toString().trim().replace(/\.0$/, "") : "") === normL;
            });
            report += "Lead [" + lid + "] -> Normalizado: [" + normL + "] -> ¿Coincide en UA?: " + (match ? "SÍ ✅" : "NO ❌") + "\n";
        });

        return report;
    } catch (e) {
        return "Error en diagnóstico: " + e.toString();
    }
}
