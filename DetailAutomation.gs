//Project: Detail Spreadsheet Management
//Author: Dax Gerts
//Date: 24 October 2015
//Description: the following functions are use to maintain the CLO work details spreadsheets.



//Function: onEdit
//Description: 
  // 1) Runs every time the spreadsheet is changed and checks to see if a resident signed off on a detail.
  //    If it is confirmed that a resident signed off, sends an email to the manager notifying them of detail completion.
  // 2) Runs every time the spreadsheet is changed and checks to see if manager has signed off on a detail
  //    If it is confirmed that if manager signed off, sends an email to resident with fines/comments
//Requirements: must include a trigger in order for emails to be sent properly.
  //To setup a trigger
     //1) go to the "Resources" tab and click "current project triggers"
     //2) then click "add a new trigger" and select the following, "onEdit", "from spreadsheet", and "On edit"
     //3) then click "save" to confirm the changes made

function onEdit() {
  //Check if active sheet is detail sheet (default name is "Sheet1")
  var sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() != "Sheet1") {
    return;
  }
  
  //Retrieve active cell and check if it is in the correct column
  var activeCell = sheet.getActiveCell();
  
  //Currently the "Res. Initial" column is column #4 (change as necessary)
  if(activeCell.getColumn() == "4") {
   
    //Retrive information about the active row
    var activeData = sheet.getRange(activeCell.getRow(),1,1,4)
    var data = activeData.getValues();
    var row = data[0];
    
    var residentName = row[0];
    var residentEmail = row[1];
    var residentDetail = row[2];
    
    //Browser.msgBox("Checking for detail sign off (message for testing purpose only)")
    
    //Only sends notification email if cell was previously empty and a resident is signing off for the first time
    if(activeCell.isBlank() == false) {
      
      //The following subject and message text can be changed as necessary
      var subject = residentName + " completed a detail";
      var message = residentName + " compelted a detail " + residentDetail;
      
      //House manager email (adjust this if necessary)
      var managerEmail = "gerts.dax1@gmail.com"
      
      //Sends email
      MailApp.sendEmail(managerEmail,subject,message);
      Browser.msgBox("Detail Complete! (notification email sent)")
   }
  }
   //Currently the "Manag. Initial" column is column #6 (change as necessary)
   if(activeCell.getColumn() == "6") {
      var activeData = sheet.getRange(activeCell.getRow(),1,1,9)
      var data = activeData.getValues();
      var row = data[0];
      
      var residentName = row[0];
      var residentEmail = row[1];
      var detail = row[2];
      var residentInitial = row[3];
      var timeComplete = row[4];
      var comments = row[5];
      var extension = row[6];
      var fined = row[7];
      var managerInitial = row[8];
      
      //Browser.msgBox("Checking for manager sign off (message for testing purpose only)")
      
      //Only sends notification email if cell was previously empty and a resident is signing off for the first time
      if(activeCell.isBlank() == false) {
        
        //The following subject and message text can be changed as necessary
        if(fined == false) {
          var subject = "Detail completed";
          var message = "Detail " + detail + " completed.\nComments: " + comments;
          
        } else {
          var subject = "Detail not completed, fined";
          var message = "Detail " + detail + " not completed, fined.\nComments: " + comments;
        }
        
        //Sends out resident emails, leave commented until ready
        MailApp.sendEmail(residentEmail, subject, message);
        
        
        //Save current records to archive sheet
        var targetSheet = ss.getSheetByName("Sheet2");
        activeData.copyTo("Sheet2");
     }
     
    //WARNING!!! Method not used, all records pushed to archive sheet
     /*
    Browser.msgBox("Locking cells now")
    //Lock cells from further access
    var protection = dataRange.protect().setDescription("Past deadline to complete details");
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
   }
   */
   
  }
  
  //WARNING!!! User Validation Will Not Work! (code retained for further testing or the event that policy changes)
  //Reason: Google security policy does not allow the method Session.getActiveUser().getEmail() to return
  //anything other than an empty string if user is not owner of the sheet or under the same domain as the owner.
  
  /*
  //Check valid user
  if(residentEmail == Session.getEffectiveUser().getEmail() && activeCell.isBlank() == false) {
    var subject = residentName + " completed a detail";
    var message = residentName + " completed detail: " + residentDetail;
 
    //send email
    MailApp.sendEmail(residentEmail,subject,message);
    Browser.msgBox("Detail Complete! (notification email sent)")
  } else {
    sheet.getActiveCell().setValue(null);
    Browser.msgBox("Curr: " + Session.getActiveUser().getEmail() + " TabUser: " + residentEmail);
    Browser.msgBox("Error: Please enter information in the correct row");
    return;
  }
  */
}



//Function: detailReminder
//Description: generate an automated reminder of upcoming details and emails residents in advance
//Requirements: must include a trigger in order for emails to be sent properly in advance
  //To setup a trigger
     //1) go to the "Resources" tab and click "current project triggers"
     //2) then click "add a new trigger" and select the following, "detailReminder", "time driven", "specific data and time", and enter the specific time in the field below
     //3) then click "save" to confirm the changes made

//Detail reminder
function detailReminder() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 4; // First row to confirm
  var numRows = 2;
  
  //Likely to need debugging, be wary
  var dataRange = sheet.getRange(startRow, 1, numRows, 9);
  
  var data = dataRange.getValues();
  
  for(i in data) {
    var row = data[i];
    var residentName = row[0];
    var residentEmail = row[1];
    var detail = row[2];
    var residentInitial = row[3];
    var timeComplete = row[4];
    var managerInitial = row[5];
    var extension = row[6];
    var fined = row[7];
    var comments = row[8];
    
    var subject = "Detail reminder"
    var message = residentName + ", you have a " + detail + " detail coming up soon. Please make sure to check the schedule to confirm."
      
    //Sends out resident emails, leave commented until ready
    MailApp.sendEmail(residentEmail, subject, message);
    }   
}

