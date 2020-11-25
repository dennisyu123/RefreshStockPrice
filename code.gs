/* Script run on open */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var entries = [{
    name : "Stock Price",
    functionName : "refreshStockPrice"
  }];
  sheet.addMenu("Refresh", entries)
}

/* Fetch stock data by code */
function getStockPrice(code) {
  let response = UrlFetchApp.fetch(`https://query1.finance.yahoo.com/v8/finance/chart/${code}?region=US&lang=en-US&includePrePost=false&interval=2m&range=1d&corsDomain=finance.yahoo.com&.tsrc=finance`)
  let json = JSON.parse(response.getContentText())
  return json.chart.result[0].meta.regularMarketPrice
}

/* Refresh stock data */
function refreshStockPrice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[1] // number of your sheet
  var range = sheet.getRange("B9:9") // select range
  var formulas = range.getFormulas()
  
  for (var i in formulas) {
    for (var j in formulas[i]) {
      if(formulas[i][j] === "") {
        break
      }
      let col = parseInt(j) + 1
      let cell = range.getCell(1,col)
      cell.setValue("Loading")
      cell.setFormula(formulas[i][j])
    }
  }
}
