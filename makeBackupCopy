// backup service
function makeBackupCopy() {
  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var timeZone = Session.getScriptTimeZone();
  var formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd' 'HH:mm:ss");
  
  // gets the name of the original file and appends the word "copy" followed by the timestamp stored in formattedDate
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + formattedDate;
  
  // gets the destination folder by their ID. REPLACE xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx with your folder's ID that you can get by opening the folder in Google Drive and checking the URL in the browser's address bar
  var destination = DriveApp.getFolderById("1dvyEjX6FzhlOaKn87kcfDqAT7SqtvWU0");
  
  // gets the current Google Sheet file
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
  
  // makes copy of "file" with "name" at the "destination"
  file.makeCopy(name, destination);
  
  // FOR FORMS...
  // where spreadsheet is located (current one)
  // var spreadsheetId =  SpreadsheetApp.getActiveSpreadsheet().getId();
  // var spreadsheetFile =  DriveApp.getFileById(spreadsheetId);
  // var ssparents = spreadsheetFile.getParents().next();
  
  // create backup folder for forms
  // var fdr_name = "Forms Backup";
  // var forms_backup_folder = "";
  
  // try {
  //   var newFdr = destination.getFoldersByName(fdr_name).next();
  //   forms_backup_folder = newFdr.getId();
    
  // }
  // catch(e) {
  //   var newFdr = destination.createFolder(fdr_name);
  //   forms_backup_folder = newFdr.getId();
    
  // }
  
  // create new folder for forms...
  // var new_folder = createFolderBasic(forms_backup_folder, name);
  
  // Loop through all the files and add the values to the spreadsheet.
  // var folder = ssparents;
  // var files = folder.getFiles();
  // var i=1;
  // while(files.hasNext()) {
  //   var file = files.next();
    
  //   //
  //   if (file.getName().indexOf("Copy of") == 0) {
  //     // that's our file!
  //     new_folder.addFile(file);
  //     ssparents.removeFile(file);
      
  //     var filename = file.getName();
  //     filename = filename.replace("Copy of", "");
  //     file.setName(filename);
  //   }
  // }
}

function createFolderBasic(folderID, folderName) {
  var folder = DriveApp.getFolderById(folderID);
  var newFolder = folder.createFolder(folderName);
  return newFolder;
};
