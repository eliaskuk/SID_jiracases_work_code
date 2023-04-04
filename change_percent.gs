function calculatePercentageChange() {
  // Get the needed sheets
  var sheet = SpreadsheetApp.getActive().getSheetByName("Summary");
  var sheet_import = SpreadsheetApp.getActive().getSheetByName("Change_percent");

  var lastRow = sheet.getLastRow();

  // SID
  var lastValueship = sheet.getRange("B" + lastRow).getValue();
  var secondToLastRowship = lastRow - 1;
  var secondLastValueship = sheet.getRange("B" + secondToLastRowship).getValue();
  var change_percentship = (lastValueship - secondLastValueship) / secondLastValueship
  var change_percent_2dship = change_percentship.toFixed(3)

  // HU
  var lastValueValue = sheet.getRange("C" + lastRow).getValue();
  var secondToLastRowValue = lastRow - 1;
  var secondLastValueValue = sheet.getRange("C" + secondToLastRowValue).getValue();
  var change_percentValue = (lastValueValue - secondLastValueValue) / secondLastValueValue
  var change_percent_2dValue = change_percentValue.toFixed(3)

  // importing the percents to sheet
  sheet_import.getRange("A2").setValue(change_percent_2dship);
  sheet_import.getRange("B2").setValue(change_percent_2dValue);
}
