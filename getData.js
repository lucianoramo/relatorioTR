function getDataFromDoneSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('to_done')
  return ws.getRange(2, 1, ws.getLastRow() - 1, 10).getDisplayValues()
}

function getDataFromPreparadoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('to_preparado')
  return ws.getRange(2, 1, ws.getLastRow() - 1, 10).getDisplayValues()
}

function getDetailedCardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('cards')
  return ws.getRange(2, 1, ws.getLastRow() - 1, 11).getValues()
}

function getDataFromProjetosSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('projetos')
  return ws.getRange(2, 1, ws.getLastRow() - 1, 15).getValues()
}
function getDatesFromReport(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Relatorio')
  const dates = [ws.getRange("D3").getValue().toISOString().slice(0, 10),ws.getRange("D4").getValue().toISOString().slice(0, 10)]
  //console.log(dates)
  return dates
}