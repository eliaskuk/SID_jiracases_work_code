function printWeekNumber() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Summary");
  var sheet2 = SpreadsheetApp.getActive().getSheetByName("Sums");
  
  var lastRow = sheet.getLastRow();
  var sum_shipments = sheet2.getRange("B3").getValue();
  var sum_line_value = sheet2.getRange("C3").getValue();
  var lastWeekNumber = sheet.getRange("A" + lastRow).getValue();
  var currentWeekNumber = lastWeekNumber ? lastWeekNumber + 1 : 1;

  sheet.getRange("A" + (lastRow + 1)).setValue(currentWeekNumber);
  sheet.getRange("B" + (lastRow + 1)).setValue(sum_shipments);
  sheet.getRange("C" + (lastRow + 1)).setValue(sum_line_value);
}

