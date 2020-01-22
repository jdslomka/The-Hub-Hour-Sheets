/** Library for all employee sheets */

// CREATE SPREADSHEET FUNCTION
function CreateNextSchedule() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Set current sheet to active them duplicate
  spreadsheet.getRange('C20').activate();
  spreadsheet.duplicateActiveSheet();
  
  // Clear Sheet
  spreadsheet.getRange('F20').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('F4').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('F11').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B4:E17').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('G4:G17').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
  // Set the proper dates
  spreadsheet.getRange('A3').activate();
  spreadsheet.getRange('A17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // Paste date
  spreadsheet.getActiveRangeList().setBackground('#f3f3f3'); // Change colour to grey
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A17'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // Drag and autofill
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true}); // Clear top
  
  // Move sheet to front
  spreadsheet.moveActiveSheet(1);
  
  // Set sheet name to first and last dates
  var date1 = spreadsheet.getRange('A4').getDisplayValue().split(" ");
  var date2 = spreadsheet.getRange('A17').getDisplayValue().split(" ");
  var name = date1[1] + " " + date1[2] + " - " + date2[1] + " " + date2[2]
  spreadsheet.renameActiveSheet(name);
  spreadsheet.getRange('B5').activate();
  
  // Notification
  SpreadsheetApp.getActiveSpreadsheet().toast("Pay period: " + name, 'Schedule Created', 10);
};

// Function for Button
function pressMe() {
  CreateNextSchedule();
};


// MAILING FUNCTION

function sendSheetToPdfwithA1MailAdress(){ // this is the function to call
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet(); // Sends current activated sheet
  
  // if you change the number, change it also in the parameters below
  var shName = sh.getName()
  var ssName = ss.getName()
  
  
  // EMAIL RECIPIENTS
  // "jdslomka@icloud.com"
  // "mana6250@mylaurier.ca, kech4630@mylaurier.ca"
  var emails = "mana6250@mylaurier.ca, sdehoop@wlu.ca"
  //var emails = 'jdslomka@icloud.com'
  sendSpreadsheetToPdf(0, ssName + " " + "("+ shName + ")", emails, ssName + "'s schedule " + shName + " <TheHub>", "");
}

function sendSpreadsheetToPdf(sheetNumber, pdfName, email, subject, htmlbody) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId()  
  var sheetId = spreadsheet.getActiveSheet().getSheetId()
  var url_base = spreadsheet.getUrl().replace(/edit$/,'');
  
  var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
      + (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId))  
      // following parameters are optional...
      + '&size=A4'      // paper size
      + '&portrait=false'    // orientation, false for landscape
      + '&fitw=true'        // fit to width, false for actual size
      + '&sheetnames=true&printtitle=true&pagenumbers=false'  //hide optional headers and footers
      + '&gridlines=false'  // hide gridlines
      + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

  var options = { 
      'muteHttpExceptions': true,
      headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),}
  }
  
  DriveApp.getRootFolder()
   
   try {
     var response = UrlFetchApp.fetch(url_base + url_ext, options);
  } catch (e) {
    Logger.log(e)
  }
  
  var blob = response.getBlob().setName(pdfName + '.pdf');
  if (email) {
    var mailOptions = {
      attachments:blob, htmlBody:htmlbody
    }
    
MailApp.sendEmail(
      email, 
      subject, 
      "", 
      mailOptions);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Don't forget to log your hours on LORIS.","Email Sent!", 5);
}
