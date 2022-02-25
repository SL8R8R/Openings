function onEdit(e) {
  try {
    autoCaps(e); //Re-case any text to UPPERCASE for uniformity/asthetics.
    mutuallyExclusiveCheckboxes_(e); //Run the mutuallyExclusiveCheckboxes script each time an edit is made
    
//Checks the _writeonce sheet to see if the cell is empty, if it is empty it writes the data to the original and _writeonce sheet.     
      if (String(e.oldValue).match(/\?/) || e.range.getSheet().getRange(e.range.getRow(), 8).getValue() === 4) {
      e.source.getSheetByName(e.range.getSheet().getName() + '_writeonce').getRange(e.range.getRow(), e.range.getColumn()).setValue(e.value);
    } else { //Overwrite cell just modified to match same cell on _writeonce sheet.
      // writeOnceReadMany();
    }
  } catch (error) {
    SpreadsheetApp.getActive().alert(error.message + ', stack:' + error.stack, 'Error in script', 5);
    throw error;
  }
}
