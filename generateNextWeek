function generateNextWeek() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('This Week_writeonce'), true);
  spreadsheet.deleteActiveSheet(); // Deletes current weeks sheets
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('This Week'), true);
  spreadsheet.deleteActiveSheet();
  spreadsheet.getSheetByName('Template').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setTabColor('#5b95f9');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Next Week'), true);
  spreadsheet.getActiveSheet().setName('This Week');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Copy of Template'), true);
  spreadsheet.getActiveSheet().setName('Next Week');
  var protection = spreadsheet.getActiveSheet().protect();
  protection.setUnprotectedRanges([spreadsheet.getRange('D6:D'), spreadsheet.getRange('G6:J')]);
  spreadsheet.moveActiveSheet(3);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template'), true);
  spreadsheet.getActiveSheet().hideSheet();
// Creates Helper Sheet based on "Write Once Read Many.gs"
  spreadsheet.getSheetByName('This Week').showSheet()
  .activate();
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('This Week_writeonce');
  spreadsheet.moveActiveSheet(4);
  var protection = spreadsheet.getActiveSheet().protect();
  spreadsheet.getActiveSheet().setTabColor('#ff0000');
  spreadsheet.getRange('A1:J5').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: false, skipFilteredRows: true})
  .setBackground('#ff0000');
  dateRange();

  spreadsheet.getSheetByName('Next Week').showSheet()
  .activate();
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('Next Week_writeonce');
  spreadsheet.moveActiveSheet(5);
  var protection = spreadsheet.getActiveSheet().protect();
  spreadsheet.getActiveSheet().setTabColor('#A020F0');
  spreadsheet.getRange('A1:J5').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: false, skipFilteredRows: true})
  .setBackground('#A020F0');
  dateRange();
};

function dateRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var date = new Date();
  var endDate = new Date();
  var days = 6;
  var startDays = 7;
  date.setDate(date.getDate() + startDays)
  endDate.setDate(date.getDate() + days);
  var stringDate1 = Utilities.formatDate(date, tz, 'M/dd');
  var stringDate2 = Utilities.formatDate(endDate, tz, 'M/dd');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Next Week").getRange("D4").setValue("OPENINGS " + stringDate1 + " - " + stringDate2);
}
