function onOpen() {
 SpreadsheetApp
   .getUi()
   .createMenu("Add Holes")
   .addItem("Add Holes for This Week", "showThisWeekSidebar")
   .addItem("Add Holes for Next Week", "showNextWeekSidebar")
  //  .addSeparator()
  //  .addItem("Add Holes for Holidays", "showHolidaySidebar")
  //  .addSeparator()
  //  .addItem("Synchronize Account", "showSynchronizeSidebar")
   .addToUi();
}
// ==========================================================================================================
function showThisWeekSidebar() {
  var html = HtmlService
      .createTemplateFromFile('This Week');

  // Add the dropdown lists to the template
  html.namedRangesDPDWN = SpreadsheetApp.getActiveSheet().getRange("Named Ranges!NamedRanges").getValues();

  // Keep adding the variables you need based on the ranges containing your dropdown values
  // ...

  // Prepares the template to be shown in the UI
  html = html.evaluate()
      .setTitle('Nordic Security Services')
      .setWidth(200);

  SpreadsheetApp.getUi().showSidebar(html);
}
// ==========================================================================================================
function showNextWeekSidebar() {
  var html = HtmlService
      .createTemplateFromFile('Next Week');

  // Add the dropdown lists to the template
  html.namedRangesDPDWN = SpreadsheetApp.getActiveSheet().getRange("Named Ranges!NamedRanges").getValues();

  // Keep adding the variables you need based on the ranges containing your dropdown values
  // ...

  // Prepares the template to be shown in the UI
  html = html.evaluate()
      .setTitle('Nordic Security Services')
      .setWidth(200);

  SpreadsheetApp.getUi().showSidebar(html);
}
// ==========================================================================================================
function showSynchronizeSidebar() {
  var html = HtmlService
      .createTemplateFromFile('Sync Sheets');

  // Add the dropdown lists to the template
  html.namedRangesDPDWN = SpreadsheetApp.getActiveSheet().getRange("Named Ranges!NamedRanges").getValues();

  // Keep adding the variables you need based on the ranges containing your dropdown values
  // ...

  // Prepares the template to be shown in the UI
  html = html.evaluate()
      .setTitle('Nordic Security Services')
      .setWidth(200);

  SpreadsheetApp.getUi().showSidebar(html);
}
// ==========================================================================================================
function showHolidaySidebar() {
  var html = HtmlService
      .createTemplateFromFile('Holiday');

  // Add the dropdown lists to the template
  html.namedRangesDPDWN = SpreadsheetApp.getActiveSheet().getRange("Named Ranges!NamedRanges").getValues();

  // Keep adding the variables you need based on the ranges containing your dropdown values
  // ...

  // Prepares the template to be shown in the UI
  html = html.evaluate()
      .setTitle('Nordic Security Services')
      .setWidth(200);

  SpreadsheetApp.getUi().showSidebar(html);
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowThis(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('This Week');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'This Week!'+account;
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
      }
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
 }
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowThis_writeonce(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('This Week_writeonce');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'This Week_writeonce!'+account;
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
      
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
 }
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowNext(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('Next Week');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'Next Week!'+account;
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
      
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
  }
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowNext_writeonce(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('Next Week_writeonce');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'Next Week_writeonce!'+account;
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
      
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
  }
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowHoliday(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('Holiday Openings');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'Holiday Openings!'+account+'_Holiday';
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
      }
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
 }
}
// ==========================================================================================================
// Take the Account and number of holes from
// the sidebar and insert rows to the proper named range
function insertRowHoliday_writeonce(account,n_rows) {  
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var ws = ss.getSheetByName('Holiday Openings_writeonce');  // Change to your sheet name
    
    //Replace space with underscore
    account=account.replace(/ /g,"_");
    var nameRange = 'Holiday Openings_writeonce!'+account;
    var range = ss.getRangeByName(nameRange);
    var rangeRows = range.getNumRows();
    n_rows-=1;
    
    if(rangeRows == 2){
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
      }
    }
    else if (rangeRows > 2) {
      //add row      
      ws.insertRowsBefore(range.getLastRow(),n_rows+1);
      //Show all rows in the namedRange
      ws.showRows(range.getRow(),range.getNumRows()+n_rows+1);
      //Get formula in the header row in column E
      var formula = ws.getRange(range.getRow(),5).getFormula();
      //Add formula in the newly added rows
      for (var i = 0; i<n_rows+1; i++){
      ws.getRange(range.getLastRow()+i,5).setFormula(formula);
    }
 }
}
// ==========================================================================================================
// function synchronizeAccounts(account) {  
//     var ss = SpreadsheetApp.getActiveSpreadsheet(); 
//     var ws = ss.getSheetByName('This Week');  // Change to your sheet name
    
//     //Replace space with underscore
//     account=account.replace(/ /g,"_");
//     var nameRange = 'This Week!'+account;
//     var range = ws.getRangeByName(nameRange);

//     ws.getRange(range).clearContent;
// }
