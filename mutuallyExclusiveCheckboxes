/**
* Unticks checkboxes when a checkbox is ticked on the same row.
*
* To take this script into use:
* 
*  - make a backup of your spreadsheet through File > Make a copy
*  - select all the text in this script, starting at the first "/**" line above,
*    and ending at the last "}"
*  - copy the script to the clipboard with Control+C (on a Mac, ⌘C)
*  - open the spreadsheet where you want to use the function
*  - choose Tools > Script editor > Blank (this opens a new tab in the browser)
*  - if you see just the 'function myFunction() {}' placeholder, press Control+A
*    (on a Mac, ⌘A), followed by Control+V (⌘V) to paste the script in
*  - otherwise, choose File > New > Script file, then press Control+A (⌘A)
*    followed by Control+V (⌘V) to paste the script in
*  - if you have an existing onEdit(e) function, add the following line as the
*    first line after the initial '{' in that onEdit(e) function:
*      mutuallyExclusiveCheckboxes_(e);
*    ...and then delete the onEdit(e) function below
*  - modify the settings under "START modifiable parameters" as necessary
*  - press Control+S (⌘S) to save the script
*  - when prompted, name the project and file 'Mutually exclusive checkboxes'
*  - close the script editor tab and go back to the spreadsheet tab
*  - the script will run automatically when you edit a cell
*
* see https://support.google.com/docs/thread/27755440
*/

/**
* Simple trigger that runs each time the user edits the spreadsheet.
*
* @param {Object} e The onEdit() event object.
*/


/**
* Unticks checkboxes when a checkbox is ticked on the same row.
*
* @param {Object} e The onEdit() event object.
*/
function mutuallyExclusiveCheckboxes_(e) {
  
  try {
    ////////////////////////////////
    // [START modifiable parameters]
    var sheetsToWatch = /./i; 
    //                  this is a regular expression
    //                  /./i means every sheet
    //                  /^Sheet1$|^Sheet2$/i means Sheet1 and Sheet2
    //                  see http://en.wikipedia.org/wiki/Regular_expression
    //                  see https://github.com/google/re2/wiki/Syntax
    var mutuallyExclusiveCheckboxRows = 'G6:H';
    //                  checkboxes in this range on sheetsToWatch will be watched
    //                  when one checkbox on a row is checked, other checkboxes on that row are automatically unchecked
    var uncheckedValues = [0, false];
    //                  specify an "unchecked" value for each checkbox on a row
    //                  by default, the unchecked value of a checkbox is "false" (enter it in the list without quotes)
    //                  there must be as many unchecked values as there are columns in mutuallyExclusiveCheckboxRows
    //                  when a checkbox on the row is automatically unchecked, its value is set to this value
    // [END modifiable parameters]
    ////////////////////////////////
    var tickedValue = e.source.getActiveCell().getValue();
    var checkboxes = {};
    checkboxes.range = e.source.getActiveSheet().getRange(mutuallyExclusiveCheckboxRows);
    checkboxes.rowStart = checkboxes.range.getRow();
    checkboxes.columnStart = checkboxes.range.getColumn();
    checkboxes.numColumns = checkboxes.range.getWidth();
    if (uncheckedValues.length !== checkboxes.numColumns) {
      throw new Error('The number of values in uncheckedValues differs from the number of columns in mutuallyExclusiveCheckboxRows.');
    }
    if (!tickedValue // assuming that checkbox values are like [true, false] or [8, 0] instead of ['Checked', 'Unchecked']
        || !e.range.getSheet().getName().match(sheetsToWatch)
        || e.range.rowStart < checkboxes.rowStart
        || e.range.columnStart < checkboxes.columnStart
        || e.range.rowEnd > checkboxes.range.getLastRow()
        || e.range.columnEnd > checkboxes.range.getLastColumn()) {
      return;
    }
    var tickedColumn = e.range.columnStart - checkboxes.columnStart;
    for (var column = 0; column < checkboxes.numColumns; column++) {
      if (column === tickedColumn) { // skip the checkbox cell that was ticked
        continue;
      }
      var checkboxCell = checkboxes.range.offset(e.range.rowStart - checkboxes.rowStart, column, 1, 1);
      var rule = checkboxCell.getDataValidation();
      if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        checkboxCell.setValue(uncheckedValues[column]);
      }
    }
  } catch (error) {
    SpreadsheetApp.getActive().toast(error.stack + error.message, 'Error in checkbox script', 30);
    throw error;
  }
}
