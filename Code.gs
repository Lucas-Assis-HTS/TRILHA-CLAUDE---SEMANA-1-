/*
  Trilha Claude · registro de progresso numa planilha do Google
  Cole este código em Extensões > Apps Script da planilha e implante como "App da Web"
  (Executar como: você · Quem pode acessar: Qualquer pessoa). O endereço gerado (termina em /exec)
  vai na linha SHEET_URL do index.html.
*/
var ABA = 'Respostas';
var COLS = ['id','quando','tipo','semana','apelido','area','personagem','uso_antes','pontos','patente','missoes','medalhas','desafio','concluido','missao_atual'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = aba();
    var data = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.id) && String(data[i][3]) === String(d.semana)) { row = i + 1; break; }
    }
    var vals = COLS.map(function (c) {
      if (c === 'quando') return new Date();
      var v = d[c];
      return (v === undefined || v === null) ? '' : v;
    });
    if (row > 0) sh.getRange(row, 1, 1, COLS.length).setValues([vals]);
    else sh.appendRow(vals);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.acao === 'placar') {
    var sh = aba();
    var data = sh.getDataRange().getValues();
    var sem = String(p.semana || '1');
    var n = 0, c = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]) === sem) {
        n++;
        if (data[i][13] === true || String(data[i][13]) === 'true') c++;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ entraram: n, concluiram: c }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Trilha Claude: registro ativo');
}

function aba() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}
