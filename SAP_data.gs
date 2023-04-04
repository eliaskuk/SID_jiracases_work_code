// Calculates how many days ago has the shipment been created and calculate pending status
function calculateDaysBetweenDates() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("SAP_DATA");
  var lastRow = sheet.getLastRow();
  var dateRange = sheet.getRange("AC2:AC" + lastRow);
  var dateValues = dateRange.getValues();
  var results = [];
  var today = new Date();
  
  for (var i = 0; i < dateValues.length; i++) {
    var date = dateValues[i][0];
    
    if (date == "") {
      results.push([null]);
      continue;
    }
    var formattedDate = new Date(date.split(".").reverse().join("-"));
    
    var diff = today.getTime() - formattedDate.getTime();
    var diffInDays = diff / (1000 * 60 * 60 * 24);
    var roundedDiffInDays = Math.round(diffInDays);
  
    // set pending status
    var pending_status = "";

    if (roundedDiffInDays < 14) {
      pending_status = " Less than 2 weeks"
    }
    else if (roundedDiffInDays < 32) {
      pending_status = "Less than 1 month"
    }
    else if (roundedDiffInDays < 63) {
      pending_status = "Less than 2 months"
    }
    else if (roundedDiffInDays < 94) {
      pending_status = "Less than 3 months"
    }
    else {
      pending_status = "More than 3 months"
  }
  sheet.getRange(i + 2, 38).setValue(roundedDiffInDays);
  sheet.getRange(i + 2, 44).setValue(pending_status);
  }
}
// Converts currencies to euro from column AH appends to AM
function convertToEuro() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sapSheet = ss.getSheetByName("SAP_DATA");
  var currencySheet = ss.getSheetByName("Currency");
  var sapData = sapSheet.getDataRange().getValues();
  var currencyData = currencySheet.getDataRange().getValues();

  for (var i = 1; i < sapData.length; i++) {
    var netPrice = sapData[i][33];
    var amount = sapData[i][11];

    // Making the number to be float
    if (typeof sapData[i][33] === 'string') {
      netPrice = parseFloat(sapData[i][33].replace(/\s+/g, ''));
    }   
    else {
      netPrice = parseFloat(String(sapData[i][33]).replace(/\s+/g, ''));
    }
    var currency = sapData[i][36];
    
    for (var j = 0; j < currencyData.length; j++) {
      if (currencyData[j][0] == currency) {
        var exchangeRate = currencyData[j][1];
        var euroValue = netPrice * exchangeRate * amount;
        sapSheet.getRange(i + 1, 39).setValue(euroValue.toFixed(0));       
      }
    }
  }
}
// Check which region the shipment is from
function getRegion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sapSheet = ss.getSheetByName("SAP_DATA");
  var regionSheet = ss.getSheetByName("Region");
  var sapData = sapSheet.getDataRange().getValues();
  var regionData = regionSheet.getDataRange().getValues();

  for (var i = 1; i < sapData.length; i++) {
    var region = sapData[i][13];
    var found = false;
    
    for (var j = 1; j < regionData.length; j++) {
      
      if (regionData[j][0] == region) {
        sapSheet.getRange(i + 1, 40).setValue(regionData[j][1]);
        found = true;
        break;
      }
    }
    if (!found) {
      sapSheet.getRange(i + 1, 40).setValue("other");
    }
  }
}
// Gets the jira num, jira responsible and jira status
function matchJiraNumbers() {
  var sapDataSheet = SpreadsheetApp.getActive().getSheetByName("SAP_DATA");
  var shipNumsSheet = SpreadsheetApp.getActive().getSheetByName("ship_nums");
  
  var sapData = sapDataSheet.getRange(1, 2, sapDataSheet.getLastRow(), 1).getValues();
  var shipNums = shipNumsSheet.getRange(1, 2, shipNumsSheet.getLastRow(), 7).getValues();
  
  
  for (var i = 1; i < sapData.length; i++) {
    var sapNum = sapData[i][0];
    if (sapNum) {
      for (var j = 1; j < shipNums.length; j++) {
        var shipNum = shipNums[j][1];
      
        if (sapNum == shipNum) {
      
          // set hyperlink to the Jira case
          var linkName = shipNums[j][0]
          
          var linkURL = "https://kalmarws.atlassian.net/browse/" + linkName

          sapDataSheet.getRange(i + 1, 41).setFormula('=HYPERLINK("' + linkURL + '","' + linkName + '")');
          sapDataSheet.getRange(i + 1, 42).setValue(shipNums[j][4]);
          sapDataSheet.getRange(i + 1, 43).setValue(shipNums[j][5]);

          break;
        }
      }
    }
  }
}

  