/* Script run on open */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1]
  var cell = sheet.getRange("J1")
  cell.setValue(getRand())
  SpreadsheetApp.flush()
}

/* Generate random number for refresh */
function getRand() { 
  return Math.random()
}

/* Fetch stock data by code */
function getStockPrice(code, rand) {
  let response = UrlFetchApp.fetch(`https://query1.finance.yahoo.com/v8/finance/chart/${code}?region=US&lang=en-US&includePrePost=false&interval=2m&range=1d&corsDomain=finance.yahoo.com&.tsrc=finance`)
  let json = JSON.parse(response.getContentText())
  return json.chart.result[0].meta.regularMarketPrice
}
