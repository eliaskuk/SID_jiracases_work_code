// read shipmentnumbers from column summary
function getShipmentNumbers() {
  var sheet1 = SpreadsheetApp.getActive().getSheetByName("EDC_JIRA_DATA");
  var sheet2 = SpreadsheetApp.getActive().getSheetByName("ship_nums");
  var values = sheet1.getRange(1, 3, sheet1.getLastRow(), 1).getValues();
  var ship_nums = [];
  var jiras = [];
  var jira_resb = [];
  var jira_sta = [];
  var counter = 0;
  
  for (var i = 0; i < values.length; i++) {
    var cell = values[i][0];

    if (cell) {
      var words = cell.split(" ");

      for (var j = 0; j < words.length; j++) {
        if (words[j].toUpperCase().indexOf("FI9") != -1) {
          ship_nums.push(words[j]);

          var jira = sheet1.getRange(i + 1, 2).getValue();
          var resb = sheet1.getRange(i + 1, 7).getValue();
          var status = sheet1.getRange(i + 1, 4).getValue();
          jiras.push(jira)
          jira_resb.push(resb)
          jira_sta.push(status)

          counter++;
        }
      }
    }
  }
  // Print numbers to sheet Ship_nums
  for (var i = 1; i < ship_nums.length; i++) {
    sheet2.getRange(i + 1, 1).setValue(ship_nums[i]);
    sheet2.getRange(i + 1, 2).setValue(jiras[i]);
    sheet2.getRange(i + 1, 6).setValue(jira_resb[i]);
    sheet2.getRange(i + 1, 7).setValue(jira_sta[i]);
  }
}
