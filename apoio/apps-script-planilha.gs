// Google Apps Script: grava 1 linha por evento (lead_inicio e diagnostico_completo) em uma planilha.
// Como usar: planilha > Extensões > Apps Script > cole este código > Implantar > Nova implantação >
// Tipo "App da Web" > Executar como: você > Acesso: Qualquer pessoa > copie a URL para WEBHOOK_URL no HTML.
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('respostas') ||
            SpreadsheetApp.getActiveSpreadsheet().insertSheet('respostas');
  var cab = aba.getLastRow() ? aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0] : [];
  d.data_hora = new Date();
  Object.keys(d).forEach(function (k) {
    if (cab.indexOf(k) < 0) { cab.push(k); aba.getRange(1, cab.length).setValue(k); }
  });
  aba.appendRow(cab.map(function (k) { return d[k] === undefined ? '' : d[k]; }));
  return ContentService.createTextOutput('ok');
}
