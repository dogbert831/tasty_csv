//initial credit for this project goes to Arthur Squires
//https://firebyarthur.com
//thanks for giving me the basis to build this out to suite my own needs

var csvFileName = "tastyworks.csv";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Trading Integrations')
  .addItem('Load Positions CSV','importCSVFromGoogleDrive')
  .addToUi();
}

//Create spreadsheet column numbers
var symbolSheet = 1;
var statusSheet = 2;
var expireDateSheet = 3;
var originalCreditSheet = 4;
var costSheet = 5;
var netLiqSheet = 6;
var plSheet = 7;
var plPercentSheet = 8;
var deltaSheet = 9;
var adjust1Sheet = 10;
var adjust2Sheet = 11;
var adjust3Sheet = 12;
var adjust4Sheet = 13;
var adjust5Sheet = 14;
var targetNetLiqSheet = 15;
var quantityNumSheet = 16;
var targetMarketPriceSheet = 17;
var openDateSheet = 18;
var openDaysInTradeSheet = 19;
var startingDTESheet = 20;
var biggestDeltaOpenSheet = 21;
var closedDateSheet = 22;
var daysInTradeSheet = 23;

//create headers for open positions sheet
function createOpenHeaders(sheet) {
  sheet.getRange(1,symbolSheet).setValue("Ticker");
  sheet.getRange(1,statusSheet).setValue("Status");
  sheet.getRange(1,expireDateSheet).setValue("DTE");
  sheet.getRange(1,originalCreditSheet).setValue("Original Credit");
  sheet.getRange(1,costSheet).setValue("Total Cost Basis");
  sheet.getRange(1,netLiqSheet).setValue("Open Net Liq");
  sheet.getRange(1,plSheet).setValue("Position P/L");
  sheet.getRange(1,plPercentSheet).setValue("Position P/L Percent");
  sheet.getRange(1,deltaSheet).setValue("Biggest Delta");
  sheet.getRange(1,adjust1Sheet).setValue("1st Adj Credit");
  sheet.getRange(1,adjust2Sheet).setValue("2nd Adj Credit");
  sheet.getRange(1,adjust3Sheet).setValue("3rd Adj Credit");
  sheet.getRange(1,adjust4Sheet).setValue("4th Adj Credit");
  sheet.getRange(1,adjust5Sheet).setValue("5th Adj Credit");
  sheet.getRange(1,targetNetLiqSheet).setValue("Target Net Liq");
  sheet.getRange(1,quantityNumSheet).setValue("Quantity");
  sheet.getRange(1,targetMarketPriceSheet).setValue("Target Market Price");
  sheet.getRange(1,openDateSheet).setValue("Open Date");
  sheet.getRange(1,openDaysInTradeSheet).setValue("Days In Trade");
  sheet.getRange(1,startingDTESheet).setValue("Starting DTE");
  sheet.getRange(1,biggestDeltaOpenSheet).setValue("Biggest Delta On Open");
}

//create headers for closed positions sheet
function createClosedHeaders(sheet) {
  sheet.getRange(1,symbolSheet).setValue("Ticker");
  sheet.getRange(1,statusSheet).setValue("Status");
  sheet.getRange(1,expireDateSheet).setValue("DTE");
  sheet.getRange(1,originalCreditSheet).setValue("Original Credit");
  sheet.getRange(1,costSheet).setValue("Total Cost Basis");
  sheet.getRange(1,netLiqSheet).setValue("Open Net Liq");
  sheet.getRange(1,plSheet).setValue("Position P/L");
  sheet.getRange(1,plPercentSheet).setValue("Position P/L Percent");
  sheet.getRange(1,deltaSheet).setValue("Biggest Delta");
  sheet.getRange(1,adjust1Sheet).setValue("1st Adj Credit");
  sheet.getRange(1,adjust2Sheet).setValue("2nd Adj Credit");
  sheet.getRange(1,adjust3Sheet).setValue("3rd Adj Credit");
  sheet.getRange(1,adjust4Sheet).setValue("4th Adj Credit");
  sheet.getRange(1,adjust5Sheet).setValue("5th Adj Credit");
  sheet.getRange(1,targetNetLiqSheet).setValue("Target Net Liq");
  sheet.getRange(1,quantityNumSheet).setValue("Quantity");
  sheet.getRange(1,targetMarketPriceSheet).setValue("Target Market Price");
  sheet.getRange(1,openDateSheet).setValue("Open Date");
  sheet.getRange(1,openDaysInTradeSheet).setValue("Days In Trade");
  sheet.getRange(1,startingDTESheet).setValue("Starting DTE");
  sheet.getRange(1,biggestDeltaOpenSheet).setValue("Biggest Delta On Open");
  sheet.getRange(1, closedDateSheet).setValue("Closed Date");
  sheet.getRange(1, daysInTradeSheet).setValue("Days In Trade");
}

function importCSVFromGoogleDrive() {
  if (!DriveApp.getFilesByName(csvFileName).hasNext()) throw new Error("Please make sure a CSV file named "+csvFileName+" has been saved in your Google drive.");
  
  var file = DriveApp.getFilesByName(csvFileName).next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  Logger.log(csvData);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Open Positions");
  var closedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Closed Positions 2018");
  if (sheet == null) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Open Positions");
  if (closedSheet == null) closedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Closed Positions 2018");
  
  var netLiqHash = {};
  var costHash = {};
  var expireHash = {};
  var spreadsheetContainHash = {};
  var biggestDeltaHash = {};
  var quantityNumHash = {};
  
  //Create the cvs format variables
  var symbolColumn = 1;
  var typeColumn = 2;
  var strikeColumn = 6;
  var deltaColumn = 16;
  var netLiqColumn = 15;
  var costColumn = 13;
  var expireColumn = 5;
  var quantityColumn = 3;
  //End create file format
  
  //End create spreadhseet column numbers
  
  checkCSVFormat();
  //Loop through csv and build hashes
  for ( var i=1, lenCsv=csvData.length; i<lenCsv; i++ ) {
    var typeArray = ((csvData[i][symbolColumn]).split(" "));
    var tradeType = csvData[i][typeColumn];
    var symbol = typeArray[0];
    var tmpType = typeArray[typeArray.length-1];
    var strike = csvData[i][strikeColumn];
    var quantity = csvData[i][quantityColumn];
    var tmpNL = netLiqHash[symbol];
    
    if (tmpNL == null) tmpNL = 0;
    var tmpCost = costHash[symbol];
    if (tmpCost == null) tmpCost = 0;
    
    //get biggest delta
    var tmpDelta = biggestDeltaHash[symbol];
    
    var deltaStr = csvData[i][deltaColumn];
    
    //Check if delta is a valid number
    if(!isNaN(parseFloat(deltaStr)) && isFinite(deltaStr)) {
      
      var newDelta = Math.abs(csvData[i][deltaColumn])*100;
      if ( Math.abs(quantity) > 0 ) { // If there is still a closed position in the cvs, we don't want it's values muddling us.
        
        if (tmpDelta != null) {
          if (newDelta > tmpDelta) biggestDeltaHash[symbol] = newDelta;
        }
        else {
          biggestDeltaHash[symbol] = newDelta;
        }
      }
      //end getting biggest delta
      
    } //end check if delta is valid number
    
    var quantityNum = Math.abs(quantity);
    if (quantityNum > 0) quantityNumHash[symbol] = quantityNum;
    
    //only import date if it's an option NOT a stock
    if (tradeType == "OPTION") {
       tmpNL = tmpNL + parseFloat((csvData[i][netLiqColumn]).replace(",",""));
       tmpCost = tmpCost + parseFloat((csvData[i][costColumn]).replace(",",""));
    }
    var tmpExpire = csvData[i][expireColumn];
    tmpExpire = tmpExpire.toString().replace("d","");
    var dte = Math.abs(tmpExpire);
    var oldDTE = expireHash[symbol];
    
    if (oldDTE != null && oldDTE < dte) dte = oldDTE;
    
    netLiqHash[symbol] = tmpNL;
    costHash[symbol] = tmpCost;
    if (quantityNum > 0) expireHash[symbol] = dte;
    
  }
  
  //Loop through all spreadsheet rows, find open positions and update
  //values
  for (var i=2, rowCount=sheet.getLastRow(); i<= rowCount; i++) {
    var tmpStatus = sheet.getRange(i, statusSheet);
    var tmpSymbol = sheet.getRange(i,symbolSheet);
    var tmpStatusVal = tmpStatus.getValue();
    var tmpSymbolVal = tmpSymbol.getValue();
    
    //If the position is open, we can see if we have a net liq
    if (tmpStatusVal == "Open") {
      var tmpNetLiq = Math.abs(netLiqHash[tmpSymbolVal]);
      if (!isNaN(tmpNetLiq) && tmpNetLiq != null && tmpNetLiq != 0) {
        var tmpNetLiqCell = sheet.getRange(i,netLiqSheet);
        tmpNetLiqCell.setValue(tmpNetLiq);
      }
      
      //if there is no net liq or it's zero, we'll mark this position closed
      if (isNaN(tmpNetLiq) || tmpNetLiq == null || tmpNetLiq == 0) tmpStatus.setValue("Closed")
      else { //If we are setting this to closed, there is no need to fill in the rest of the data because
        //it will just mess up the history.
        
        // var tmpDeltaCell = sheet.getRange(i,deltaSheet);
        
        var biggestHash = biggestDeltaHash[tmpSymbolVal];
        var expireValue = expireHash[tmpSymbolVal];
        var expireCell = sheet.getRange(i,expireDateSheet);
        expireCell.setValue(expireValue);
        
        var quantityValue = quantityNumHash[tmpSymbolVal];
        var quantityNumCell = sheet.getRange(i,quantityNumSheet);
        quantityNumCell.setValue(quantityValue);
      }
      spreadsheetContainHash[tmpSymbolVal] = tmpSymbolVal;
    }
  }
  
  //Now loop through all symbols from import and add new rows for
  //the ones that don't have open rows in the spreadsheet
  for (var key in netLiqHash) {
    if (spreadsheetContainHash[key] == null) {
      //Make sure it's not just a closed position
      if (costHash[key] > 0) {
        var lastRow = sheet.getLastRow();
        if (lastRow == 0) {
          createOpenHeaders(sheet);
          lastRow = 1;
        }
        sheet.insertRowAfter(lastRow);
        lastRow++;
        
        sheet.getRange(lastRow,symbolSheet).setValue(key);
        sheet.getRange(lastRow,statusSheet).setValue("Open");
        sheet.getRange(lastRow,expireDateSheet).setValue(expireHash[key]);
        sheet.getRange(lastRow,originalCreditSheet).setValue(costHash[key]);
        sheet.getRange(lastRow,costSheet).setFormula("=D"+lastRow+"+J"+lastRow+"+K"+lastRow+"+L"+lastRow+"+M"+lastRow+"+N"+lastRow);
        sheet.getRange(lastRow,netLiqSheet).setValue(Math.abs(netLiqHash[key]));
        sheet.getRange(lastRow,plSheet).setFormula("=E"+lastRow+"-F"+lastRow);
        sheet.getRange(lastRow,plPercentSheet).setNumberFormat("#.##%");
        sheet.getRange(lastRow,plPercentSheet).setFormula("=G"+lastRow+"/D"+lastRow);
        sheet.getRange(lastRow,deltaSheet).setValue(biggestDeltaHash[key]);
        sheet.getRange(lastRow,quantityNumSheet).setValue(quantityNumHash[key]);
        sheet.getRange(lastRow,targetNetLiqSheet).setFormula("=((D"+lastRow+"/2)-E"+lastRow+")*-1");
        sheet.getRange(lastRow,targetMarketPriceSheet).setFormula("=O"+lastRow+"/P"+lastRow);
        sheet.getRange(lastRow,openDateSheet).setValue(Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy"));
        sheet.getRange(lastRow,openDaysInTradeSheet).setFormula("=DATEDIF(R"+lastRow+",TODAY(),\"D\")");
        //column 23 left blank for IFTTT data
        sheet.getRange(lastRow,startingDTESheet).setValue(expireHash[key]);
        sheet.getRange(lastRow,biggestDeltaOpenSheet).setValue(biggestDeltaHash[key]);
      }
    }
  }
  var finalRow = sheet.getLastRow();
  if (finalRow == 0) finalRow == 2;
  
  if (closedSheet.getLastRow() == 0) {
    createClosedHeaders(closedSheet);
  }
  
  for (var i=2; i<= finalRow; i++) {
    var tmpStatus = sheet.getRange(i, statusSheet);
    if (tmpStatus.getValue() == "Closed") {
      var targetRange = closedSheet.getRange(closedSheet.getLastRow() + 1, 1);
      sheet.getRange(tmpStatus.getRow(), 1, 1, sheet.getLastColumn()).moveTo(targetRange);
      var closedDateRange = closedSheet.getRange(targetRange.getRow(),closedDateSheet);
      closedDateRange.setValue(Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy"));
      closedSheet.getRange(targetRange.getRow(), daysInTradeSheet).setFormula("=DATEDIF(R"+targetRange.getRow()+",V"+targetRange.getRow()+",\"D\")");
      sheet.deleteRow(tmpStatus.getRow());
      //Since we deleted one, we need to go back one row in the loop
      i--;
    }
  }
  //Internal checkCSVFormat option
  
  function checkCSVFormat() {
    //Create the cvs format variables
    var foundSymbol = false;
    var foundDelta = false;
    var foundNL = false;
    var foundCost = false;
    var foundExpDate = false;
    var foundStrike = false;
    var foundQuantity = false;
    
    for ( var i=0, lenCsv=csvData[0].length; i<lenCsv; i++ ) {
      var columnHeader = csvData[0][i];
      if (columnHeader == "Symbol") {
        symbolColumn = i;
        foundSymbol = true;
      }
      else if (columnHeader == "/ Delta") {
        deltaColumn = i;
        foundDelta = true;
      }
      else if (columnHeader == "NetLiq") {
        netLiqColumn = i;
        foundNL = true;
      }
      else if (columnHeader == "Cost") {
        costColumn = i;
        foundCost = true;
      }
      else if (columnHeader == "DTE") {
        expireColumn = i;
        foundExpDate = true;
      }
      else if (columnHeader == "Strike Price") {
        strikeColumn = i;
        foundStrike = true;
      }
      else if (columnHeader == "Quantity") {
        quantityColumn = i;
        foundQuantity = true;
      }
      
    }
    
    if (!foundSymbol) throw new Error("Symbol column not found in CSV import");
    if (!foundDelta) throw new Error("Delta column not found in CSV import");
    if (!foundNL) throw new Error("Net Liq column not found in CSV import");
    if (!foundCost) throw new Error("Cost column not found in CSV import");
    if (!foundExpDate) throw new Error("Expire Date column not found in CSV import");
    if (!foundStrike) throw new Error("Strike Price column not found in CSV import");
    if (!foundQuantity) throw new Error("Quantity column not found in CSV import");
    //End create file format
  }

}
