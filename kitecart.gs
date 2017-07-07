/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{ name : "Place Basket Order", functionName : "basketOrder" }]
  sheet.addMenu("Kite", entries)
}

function basketOrder(){
  var htmlOutput = HtmlService.createHtmlOutput(buildHTML()).setWidth(200).setHeight(50)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Please wait...")
}

function buildHTML(){
  return String.concat(
    orders().reduce(function(prev, curr){
    return prev += curr
    }, '<script src="https://kite.zerodha.com/publisher.js"></script>;' +
       '<script> KiteConnect.ready(function() { var kite = new KiteConnect("mpqxfuz4emmqg9q6");'),
    "kite.connect(); setTimeout(function(){ google.script.host.close(); }, 2000); })</script>")
}

function orders(){
  return SpreadsheetApp.getActiveSpreadsheet().getRange("A2:B").getValues()
  .filter(function(x) { return x.filter(Boolean).length !==0 })
  .map(function(x) {
    return orderJson(x[0], x[1])
  })
}

function orderJson(symbol, quantity){
  return String.concat("kite.add(",
                       JSON.stringify({"exchange": "NSE",
                                       "tradingsymbol": symbol,
                                       "quantity": Math.abs(quantity),
                                       "transaction_type": quantity >= 0 ? "BUY" : "SELL",
                                       "order_type": "MARKET",
                                       "product": "CNC"}),
                       ");"
                      )
}
