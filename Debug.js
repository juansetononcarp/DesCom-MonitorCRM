function getDebugInfo() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
        leads: { headers: [], row1: [], row2: [] },
        usuarios: { headers: [], row1: [], row2: [] }
    };

    var sheet = ss.getSheetByName('Leads');
    if (sheet) {
        var maxCol = Math.min(sheet.getLastColumn(), 40); // Inspect first 40 cols (covers AL=37)
        // Get headers (Row 1) and first 2 rows of data (Row 2, 3)
        var range = sheet.getRange(1, 1, 3, maxCol).getValues();
        result.leads.headers = range[0].map(function (c, i) { return i + ': ' + c; });
        result.leads.row1 = range[1].map(function (c, i) { return i + ': ' + c; });
        result.leads.row2 = range[2].map(function (c, i) { return i + ': ' + c; });
    }

    var sheetU = ss.getSheetByName('Usuarios');
    if (sheetU) {
        var rangeU = sheetU.getDataRange().getValues();
        if (rangeU.length > 0) result.usuarios.headers = rangeU[0];
        if (rangeU.length > 1) result.usuarios.row1 = rangeU[1];
    }

    return result;
}
