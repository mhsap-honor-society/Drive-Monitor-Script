/*************************************************

original script created by amit

contact  :  amit@labnol.org
twitter  :  @labnol
tutorial :  http://www.labnol.org/internet/google-drive-activity-report/13857/
written  :  November 219, 2014

***************************************************/

/* Edited for the use of the MHSAP Honor Society by Isaac Petersen
contact: ispete1@outlook.com
edited: May 18, 2019
*/
function generateReports() {
  
  // Get the current working spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getActiveSheet();
  
  // Retrieve emails from the given range
  var emails = ss.getRangeByName("EmailList"); // The range values are from E2 to E8 right now, but this can be changed in the named ranges setting in the spreadsheet
  Logger.log(emails.getValues())
  var emaillist = emails.getValues();
  
  // Gets the current time zone
  var timezone = ss.getSpreadsheetTimeZone();
  
  // Gets the current time
  var today     = new Date();
  var oneDayAgo = new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000);  
  var startTime = oneDayAgo.toISOString();
  
  // Searches the Incoming and Outgoing folders for new or modified files that were added within the past 24 hours
  var search = '(trashed = true or trashed = false) and (modifiedDate > "' + startTime + '")';
  var incomingfolder  = DriveApp.getFoldersByName("Incoming");
  var incomingfiles = incomingfolder.next().searchFiles(search);
  var outgoingfolder = DriveApp.getFoldersByName("Outgoing");
  var outgoingfiles = outgoingfolder.next().searchFiles(search);
  
  var row = "", incount=0, outcount=0;
  
  // This loop adds the files found in the search to the spreadsheet
  while( incomingfiles.hasNext() ) {
    
    var infile = incomingfiles.next();
    
    var infileName = infile.getName();
    var infileURL  = infile.getUrl();
    var inlastUpdated =  Utilities.formatDate(infile.getLastUpdated(), timezone, "yyyy-MM-dd HH:mm");
    var indateCreated =  Utilities.formatDate(infile.getDateCreated(), timezone, "yyyy-MM-dd HH:mm")
    
    row += "<li>" + inlastUpdated + " <a href='" + infileURL + "'>" + infileName + "</a></li>";
    
    sheet.appendRow([indateCreated, inlastUpdated, infileName, infileURL]);
    
    incount++;
  }
  
  // This loop theoretically will notify the original author that his paper has been reviewed
  while( outgoingfiles.hasNext() ) {
    
    var outfile = outgoingfiles.next();
    
    var outfileName = outfile.getName();
    var outfileURL  = outfile.getUrl();
    var outfileEditors = outfile.getEditors();
    var outlastUpdated =  Utilities.formatDate(outfile.getLastUpdated(), timezone, "yyyy-MM-dd HH:mm");
    var outdateCreated =  Utilities.formatDate(outfile.getDateCreated(), timezone, "yyyy-MM-dd HH:mm");
    
    // This set of loops check each of the editors and removes the ones that are peer reviewers
    for (var i = 0; i < outfileEditors.length; i++)
    {
      for (var j = 0; j < emaillist.length; j++)
      {
        if (outfileEditors[i] == emaillist[j])
        {
          outfileEditors[i] = "";
        }
      }
    }
    
    // This loop sends emails to each of the remaining editors of the document
    for (var i = 0; i < outfileEditors.length; i++)
    {
      var email = outfileEditors[i];
      if (email != "")
      {
        MailApp.sendEmail(email, "Honor Society Peer Review Completion Notification", "", {htmlBody: "<p>This is a notification for the author of " + outfileName + " that the peer review by the MHSAP Honor Society has been completed. Your paper is located at " + outfileURL + ""});
      }
    }
    
    outcount++;
  }
  
  // This statement notifies each of the peer reviewers that there have been file changes in the Incoming directory
  if (row !== "") {
    row = "<p>" + incount + " file(s) have changed in Iota Beta's Incoming Directory in the past 24 hours. Here's the list:</p><ol>" + row + "</ol>";
    row +=  "<br><small>Please contact an Honor Society Cabinet member to move your completed Peer Review paper to the outgoing folder once you've finished.</small><br><small>To stop these notifications, please contact an Honor Society Cabinet member to remove you from this list</small>";
    for (var i = 0; i < emaillist.length; i++)
    {
      var email = emaillist[i];
      if (email != "")
      {
        MailApp.sendEmail(email, "Honor Society Writing - File Activity Report", "", {htmlBody: row});
      }
    }
  }
  
}

// These last functions were not changed and should not be changed
function onOpen() {  
  var menu = [    
    { name: "☎ Help and Support »",    functionName: "help"},
    null,
    { name: "Step 1: Authorize",   functionName: "init"      },
    { name: "Step 2: Schedule Reports", functionName: "configure" },
    null,
    { name: "✖ Uninstall (Stop)",    functionName: "reset"     },
    null
  ];  
  SpreadsheetApp.getActiveSpreadsheet()
  .addMenu("➪ Drive Activity Report", menu);
}

function help() {
  var html = HtmlService.createHtmlOutputFromFile('help')
  .setTitle("Google Scripts Support")
  .setWidth(400)
  .setHeight(160);
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

function configure() {
  
  try {
    
    reset(true);
    
    var ss = SpreadsheetApp.getActive();
    
    var email = ss.getRange("E1").getValue();
    
    if (email == "amit@labnol.org") {
      Browser.msgBox("Please put your email address in cell E1 where you wish to receive the daily reports.");
      return;
    }
    
    ScriptApp.newTrigger("generateReports").timeBased().everyDays(1).create();
    
    generateReports();
    
    ss.toast("The program is now running. You can close this sheet.", "Success", -1);
    
  } catch (e) {
    Browser.msgBox(e.toString());
  }
  
}

function init() {  
  
  SpreadsheetApp.getActive().toast("The program is now initialized. Please run Step #2");
  
}

function reset(e) {
  
  var triggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);    
  }
  
  if (!e) {
    SpreadsheetApp.getActive().toast("The script is no longer active. You can re-initialize anytime later.", "Stopped", -1);
  }
  
}


