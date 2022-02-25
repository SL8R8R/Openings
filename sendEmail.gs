// function sendEmail() {
//   var SheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("# Shifts for This Week");
//   var Range = SheetName.getRange(2,4,100,1);       //2,4,100,1);
//   var Value = Range.getValues().toString();
//   // Check against range
//   if (Value > 39){
//     // Fetch the email address
//     var emailRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails").getRange("A1:A2");
//     var emailAddress = emailRange.getValue();
//     // Send Alert Email.
//     var message = 'An officer has gone over the number of hours limit! Please check the "Number of Shifts for This Week" sheet in the NSS Openings spreadsheet. "https://docs.google.com/spreadsheets/d/1MUMDZFe1Z0wI26NZOn-r4Zja6aqk4tzRh2UefhZL6fU/edit?usp=sharing". They have ' + Range ;
//     var subject = 'Officer Over Alotted Shifts';
//     MailApp.sendEmail(emailAddress, subject, message);
//   }
// }

// function SendEmail() {
//   // var ui = SpreadsheetApp.getUi();
//   var file = SpreadsheetApp.getActive();
//   var sheet = file.getSheetByName("# Shifts for This Week");  //Change as needed

//   if(sheet.getRange(99,4).getValue()>39){ 
//     Logger.log(value)     //change row and column in get range to match what you need
//     MailApp.sendEmail("srubin@nordicsec.com", "subttttject", "treeeeeeeee");
//   }
// }

// function sendEmail() {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = ss.getSheetByName("# Shifts for This Week");  
//     var values = sheet.getRange("Amount_of_Hours").getValues();
//     var value1s = 39;
//     var results = [];
//     for(var i=0;i<values.length;i++){
//       if(values[i]>value1s[i]){
//         results.push("alert on line: "+(i+2)); // +2 because the loop start at zero and first line is the second one (D2)
//       }
//     }
//     MailApp.sendEmail('srubin@nordicsec.com', 'subject', results.join("\n"));
// };
