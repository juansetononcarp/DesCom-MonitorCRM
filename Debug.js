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
