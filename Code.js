var HEADER_ROWS = 2;

var SYMBOL_COL = 2;
var PRICE_COL = 5;
var CHANGE_1h_COL = 7;
var CHANGE_1d_COL = 9;
var CHANGE_1w_COL = 11;

function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Coins options')
       .addItem('Recalculate rates', 'fillMe')
       .addToUi();
 }

function fillMe() {
  var markets = getMarkets();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Current');
  
  var columns = sheet.getLastColumn();
  
  var range = sheet.getRange(HEADER_ROWS + 1, 1, sheet.getLastRow()-1, columns);
  
  var values = range.getValues();
  for(y in values) {
    var row = values[y];
    if(markets[row[SYMBOL_COL-1]]) {
      fillRow(sheet, parseInt(y) + HEADER_ROWS + 1, markets[row[SYMBOL_COL-1]]);
    }
  }
}

function fillRow(sheet, row, market) {
  fillCell(sheet, row, PRICE_COL, parseFloat(market["price_czk"]));
  fillCell(sheet, row, CHANGE_1h_COL, parseFloat(market["percent_change_1h"])/100);
  fillCell(sheet, row, CHANGE_1d_COL, parseFloat(market["percent_change_24h"])/100);
  fillCell(sheet, row, CHANGE_1w_COL, parseFloat(market["percent_change_7d"])/100);
}

function fillCell(sheet, row, col, value) {
  var range = sheet.getRange(row, col);
  range.setValue(value);
}

function getMarkets() {
  var raw = getValuesFromApi();
  
  var markets = {};
  for(i in raw) {
    var row = raw[i];
    markets[row['symbol']] = row;
  }
  
  return markets;
}

function getValuesFromApi() {
  var url = "https://api.coinmarketcap.com/v1/ticker/?convert=CZK";
  var json = UrlFetchApp.fetch(url);
  var raw = JSON.parse(json);
  return raw;
}
