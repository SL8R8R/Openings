/**
* This script lets users enter values on a sheet but not edit them afterwards.
* 
* This script prevents cells from being updated. Empty cells can be filled in, but once 
* there is a value in a cell, it can no longer be edited.  Runs on an edit trigger.
* 
* This in effect protects sheets in a "write once" manner, albeit with some limitations.
* The script does *not* provide protection against edits by a determined user.
* For example, a user that has edit access to the spreadsheet can easily disable this script.
* 
* The script uses additional helper sheets (tabs) in the spreadsheet to store backup copies of 
* the values entered in cells. When a user edits a cell on any sheet, it is checked against the 
* same cell on the helper sheet, and:
* 
*   - if the value on the helper sheet is empty, the new value is stored on both sheets
*   - if the value on the helper sheet is not empty, it is copied back to the cell on
*     the source sheet, undoing the change
* 
* Helper sheets are created automatically when an edit is first made, one helper sheet
* per each source sheet. For a source sheet named "Sheet1", the helper sheet is "Sheet1_writeonce".
* Helper sheets are automatically hidden when created to not clutter the display, but
* they can be unhidden by any user with "can edit" rights to the spreadsheet.
* Users with edit rights can also disable this script at will.
* 
* To change a value that was entered previously, empty the corresponding cell on the helper sheet,
* then edit the cell on the source sheet.
* In the event you rename a source sheet, remember to rename the helper sheet as well.
* Choose "View > Hidden sheets" to show the helper sheet, then rename it using the pop-up
* menu at the sheet's tab at the tab bar at bottom of the browser window.
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
*  - modify the settings under "START modifiable parameters" as necessary
*  - press Control+S (⌘S) to save the script
*  - when prompted, name the project 'Write Once Read Many'
*  - choose Run > Run function > setTrigger_writeOnceReadMany > 
*    Review Permissions > [choose account] > Advanced >
*    Go to Write Once Read Many (unsafe) > Allow
*  - close the script editor tab and go back to the spreadsheet tab
*  - the script will run automatically when you edit a cell
*
* The script will from then on watch updates on all the sheets and only allow edits
* when the cell is empty to start with. The script will run under your account, which means that 
* you can make helper sheets read-only which will prevent others from tampering with the saved values.
* To protect the helper sheets, click the arrow triangle on a helper sheet's tab in the tab bar at 
* the bottom of the browser window and choose Protect sheet > Set permissions.
* 
* Note that the script only protects _values_ rather than _formulas_.
* To protect formulas, use Data > Named and protected ranges.
*  
* If your sheets that you would like to protect already have data on them, create helper
* sheets manually by choosing the Duplicate command from the sheet's tab menu at the tab bar
* at the bottom of the browser window. Rename the new sheet so that it ends in "_writeonce":
* for example, when you duplicate "Sheet1",you get "Copy of Sheet1", and must rename that to 
* "Sheet1_writeonce".
* 
* The range where edits are of this "write once" type can be limited by changing the values
* assigned to the firstDataRow, lastDataRow, firstDataColumn and lastDataColumn variables below.
* The range defined by these values is global and will apply to all the sheets the same.
* 
* You can exclude some columns and sheets from being watched by putting them on the 
* freeToEditColumns or freeToEditSheetNames lists. See below for more info.
* 
*/
function writeOnceReadMany() {
  
  try {
    var protectionModeEnum = {
      PROTECTED_BY_DEFAULT: 1,
      UNPROTECTED_BY_DEFAULT: 2,
      PROTECTED_WITH_EXCEPTIONS: 3,
      UNPROTECTED_WITH_EXCEPTIONS: 4
    };
    
    ////////////////////////////////
    // [START modifiable parameters]
    // Modify the following variables per your requirements.
    
    // Specify the range where edits are "write once". 
    // For example, to watch only the range A1:D100:
    //   set firstDataColumn="A", lastDataColumn="D", firstDataRow=1, lastDataRow=100
    // To watch only the range M20:V30:
    //   set firstDataColumn="M", lastDataColumn="V", firstDataRow=20, lastDataRow=30
    
    var firstDataColumn = "A"; // only take into account edits on or to the right of this column
    var lastDataColumn = "J"; // only take into account edits on or to the left of this column
    var firstDataRow = 6; // only take into account edits on or below this row
    var lastDataRow = 99999; // only take into account edits on or above this row
    
    // To make an exception and allow editing of certain columns, modify the list below.
    var freeToEditColumns = [[]]; // columns to exclude from protection; use [[]] to not exclude any columns
    
    // To allow users to clear a cell by selecting it and pressing Delete, set this true.
    // To not allow users to clear cells, set this false.
    var clearingValuesAllowed = false;
    
    // Choose between one of these modes:
    //
    // PROTECTED_BY_DEFAULT:
    //     1. sheets whose name matches freeToEditSheetNames remain free to edit
    //     2. other sheets will by default be protected
    //
    // UNPROTECTED_BY_DEFAULT:
    //     1. sheets whose name matches protectedSheetNames are protected
    //     2. other sheets will by default remain free to edit without protection
    //
    // PROTECTED_WITH_EXCEPTIONS:
    //     1. sheets whose name matches freeToEditSheetNames remain free to edit (but see step 3)
    //     2. other sheets will by default be protected
    //     3. additionally, sheets whose name matches protectedSheetNames will be protected,
    //        even if their name matched freeToEditSheetNames in step 1
    //
    // UNPROTECTED_WITH_EXCEPTIONS:
    //     1. sheets whose name matches protectedSheetNames are protected (but see step 3)
    //     2. other sheets will by default remain free to edit without protection
    //     3. additionally, sheets whose name matches freeToEditSheetNames remain free to edit,
    //        even if their name matched protectedSheetNames in step 1
    
    var protectionMode = protectionModeEnum.PROTECTED_BY_DEFAULT;
    
    // Names of sheets that are to be protected
    // NOTE: this setting has no effect when protectionMode = protectionModeEnum.UNPROTECTED_BY_DEFAULT
    var protectedSheetNames = []; // use ["."] to protect all sheets
    
    // Exceptions: sheets that are free to edit with no protection
    // NOTE: this setting has no effect when protectionMode = protectionModeEnum.UNPROTECTED_BY_DEFAULT
   
    var freeToEditSheetNames = ["Template","DropDown","# Shifts for This Week","Emails","Named Ranges"];
    
    // You can use regular expressions in sheet names. The match is not case-sensitive.
    //
    // Wildcards include:
    //     .           matches any single character
    //     .*          matches any number of any characters, including zero characters
    //     \d+         matches one or more digits
    //     \.          matches the period character (.)
    //     \*          matches the asterisk character (*)
    //     ^           matches the beginning of a name
    //     $           matches the end of a name
    //
    // Examples:
    //     ".*"        matches any sheet name
    //     "^Data$"    matches "Data" but not "Data1" nor "Top Data"
    //     "^Data"     matches "Data", "Data1", "Datasheet", "Data August 2018" but not "Top Data"
    //     "Data$"     matches "Data", "Top Data", "SheetData" but not "Data1" nor "Datasheet"
    //     "^2018.*F$" matches "2018 set1 F", "2018-08-30F", "2018F" but not "2019F" nor "2018F1"
    //
    // You can list several sheet name patterns by separating the patterns with commas like this:
    //     ["^Sheet1$", "^Sheet2$", "^Sheet3$"]
    //
    // Sheet names always need to be quoted (") and the list needs to be enclosed in square
    // brackets ([]) even if there is just one name in the list.
    //
    // See these sites for more info on regular expressions:
    //   - http://en.wikipedia.org/wiki/Regular_expression
    //   - https://github.com/google/re2/blob/master/doc/syntax.txt
    
    // The default suffix for helper sheets where values are copied for later checking
    var helperSheetNameSuffix = "_writeonce";
    
    // [END modifiable parameters]
    ////////////////////////////////
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getActiveSheet();
    var masterSheetName = masterSheet.getName();
    var masterRange = masterSheet.getActiveRange();
    var protectedSheetNamesRegExp;
    var freeToEditSheetNamesRegExp;
    
    // add helper sheets to the list of freeToEditSheetNames to ensure that changes to
    // a helper sheet do not trigger the creation of another _writeonce_writeonce sheet.
    freeToEditSheetNames.push(helperSheetNameSuffix + "$");
    
    // find out whether the edited sheet should be protected or not
    var sheetIsFreeToEdit = undefined;
    if (protectionMode === protectionModeEnum.PROTECTED_BY_DEFAULT) {
      sheetIsFreeToEdit = false;
      for (var sheet in freeToEditSheetNames) {
        freeToEditSheetNamesRegExp = new RegExp(freeToEditSheetNames[sheet], "i");
        if (freeToEditSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = true;
      }
    } else if (protectionMode === protectionModeEnum.UNPROTECTED_BY_DEFAULT) {
      sheetIsFreeToEdit = true;
      for (var sheet in protectedSheetNames) {
        protectedSheetNamesRegExp = new RegExp(protectedSheetNames[sheet], "i");
        if (protectedSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = false;
      }
    } else if (protectionMode === protectionModeEnum.PROTECTED_WITH_EXCEPTIONS) {
      sheetIsFreeToEdit = false;
      for (var sheet in freeToEditSheetNames) {
        freeToEditSheetNamesRegExp = new RegExp(freeToEditSheetNames[sheet], "i");
        if (freeToEditSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = true;
      }
      for (var sheet in protectedSheetNames) {
        protectedSheetNamesRegExp = new RegExp(protectedSheetNames[sheet], "i");
        if (protectedSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = false;
      }
    } else if (protectionMode === protectionModeEnum.UNPROTECTED_WITH_EXCEPTIONS) {
      sheetIsFreeToEdit = true;
      for (var sheet in protectedSheetNames) {
        protectedSheetNamesRegExp = new RegExp(protectedSheetNames[sheet], "i");
        if (protectedSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = false;
      }
      for (var sheet in freeToEditSheetNames) {
        freeToEditSheetNamesRegExp = new RegExp(freeToEditSheetNames[sheet], "i");
        if (freeToEditSheetNamesRegExp.test(masterSheetName)) sheetIsFreeToEdit = true;
      }
    } else {
      Browser.msgBox('Write once read many', 
                     'Incorrect definition of protectionMode. Please use one of:\\n' +
                     '  PROTECTED_BY_DEFAULT\\n' +
                     '  UNPROTECTED_BY_DEFAULT\\n' +
                     '  PROTECTED_WITH_EXCEPTIONS\\n' +
                     '  UNPROTECTED_WITH_EXCEPTIONS',
                     Browser.Buttons.OK);
      return;
    }
    
    if (sheetIsFreeToEdit) {
      return;
    }
    
    // convert textual column labels into column numbers
    firstDataColumn = columnLabelsToColumnNumbers_(firstDataColumn);
    lastDataColumn = columnLabelsToColumnNumbers_(lastDataColumn);
    freeToEditColumns = [].concat.apply([], columnLabelsToColumnNumbers_(freeToEditColumns)); // convert and flatten the array
    
    if (masterRange.getRow() < firstDataRow || masterRange.getColumn() < firstDataColumn || 
      masterRange.getRow() > lastDataRow || masterRange.getColumn() > lastDataColumn || 
        freeToEditColumns.indexOf(masterRange.getColumn()) >= 0) return;
    
    // find helper sheet
    var helperSheetName = masterSheetName + helperSheetNameSuffix;
    var helperSheet = ss.getSheetByName(helperSheetName);
    if (!helperSheet) { // helper sheet does not exist yet, create it as the last sheet in the spreadsheet
      helperSheet = ss.insertSheet(helperSheetName, ss.getNumSheets());
      Utilities.sleep(2000); // give time for the new sheet to render before going back
      ss.setActiveSheet(masterSheet);
      helperSheet.hideSheet();
      ss.setActiveRange(masterRange);
    }
    
    var helperRange = helperSheet.getRange(masterRange.getA1Notation());
    var newValue = masterRange.getValues();
    var oldValue = helperRange.getValues();
    var oldValueIsEmpty = true;
    
    // find out whether the edit should be allowed or not
    oldValueIteration: {
      for (var i = 0; i < oldValue.length; i++) {
        for (var j = 0; j < oldValue[i].length; j++) {
          // first, deal with fill down (Control+D) and fill right (Control+R)
          if (!i && !j && oldValue.length + oldValue[i].length > 2 && newValue[i][j] == oldValue[i][j]) {
            ; // do nothing when the value in the start cell [0][0] matches oldValue _and_ there are more values to look at
          } else if (oldValue[i][j] != "") {
            oldValueIsEmpty = false;
            break oldValueIteration;
          }
        }
      }
    }
    
    if (oldValueIsEmpty || (clearingValuesAllowed && newValue[0][0] === '')) {
      helperRange.setValues(newValue);
    } else {
      masterRange.setValues(oldValue);
      ss.toast('Overwriting values in these cells is not allowed.', 'Write once read many');
    }
  } catch (error) {
    ss.toast(error.message, 'Write once read many', 30);
    throw error;
  }
}

/**
* Converts textual column labels such as "A", "Z" and "AA" into column numbers like 1, 26 and 27.
* Returns an array of numbers. 
* Supports single values and 2-dimensional arrays like [["A", "Z", "AA"]]
* or [["A", "B"], ["M", "N"], ["AA", "AB"]].
* If the input array contains numbers or blank values, they are returned as is.
* Throws an error when the input contain illegal labels such as "1A".
*
* @param {"AA"} labels A textual column label or an array labels to convert to column numbers.
* @return {Array} Column numbers.
*/
function columnLabelsToColumnNumbers_(labels) {

  if (labels instanceof Array) {
    return labels.map(columnLabelsToColumnNumbers_);
  } else if (typeof labels == 'number') {
    return labels;
  } else if (labels == "") {
    return null;
  } else {
    var match = labels.toUpperCase().match(/(^[A-Z]+)/gmi);
    if (!match || match.length != 1)
      throw new Error('columnLabelsToColumnNumbers_ expected a textual column label like "A" or "A1", but got invalid column label "' +  labels + '".');
    labels = match[0];
    var alphabetStart = "A".charCodeAt(0);
    var alphabetSize = "Z".charCodeAt(0) - alphabetStart + 1;
    var alphaValue = 0;
    var colNumber = 0;
    for (var i = 0; i < labels.length; i++) {
      alphaValue = labels.charCodeAt(i) - alphabetStart + 1;
      colNumber += alphaValue * Math.pow(alphabetSize, labels.length - i - 1);
    }
  }
  return colNumber;
}

/**
* Creates an onEdit trigger for the writeOnceReadMany function.
*
*/
function setTrigger_writeOnceReadMany() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('writeOnceReadMany')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}
