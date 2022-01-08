/*************************************************
*
* UPDATE THESE VARIABLES
*
*************************************************/
// Comma-separated Email addresses of owner and any additional recipients for notification when the audit completes
const NOTIFICATION_RECIPIENTS = "your-email@domain.com";
// Google Sheet URL that you have access to edit (should be blank to begin with)
const GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/your-file-id/edit#gid=0";
// tab/sheet name to house the list of File IDs for everything in your Google Drive
const GOOGLE_SHEET_RESULTS_TAB_NAME = "Sheet1";
// change TIMEOUT VALUE
//var TIMEOUT_VALUE_MS = 270000; // 4.5 minutes
const TIMEOUT_VALUE_MS = 210000; // 3.5 mins (so we can run this every 5 minutes on a trigger if you want)
// new file name suffix - put onto the end of the renamed file (after the created time and contents string)
const NEW_FILE_NAME_SUFFIX = '___(RENAMED BY AUTOMATION)'; // IE: '___(RENAMED BY AUTOMATION)'
// number of characters to use from the file contents/header as part of the new file name
const NUM_CHARACTERS_FROM_FILE_CONTENTS = 30; // IE: 15 or 30 or whateer fits in the file name length max limit
// the names of the files you want to find and rename
const SEARCH_TERM = 'Untitled Document'; // IE: Untitled Document or Untitled Spreadsheet
/*

/*
*
* ONLY RUN THIS ONCE TO SET THE HEADER ROWS FOR THE GOOGLE SHEETS
* Later should probably have some logic to look up to see if the first row is set, then run this automatically (or not)
*
*/
function job_one_time_set_sheet_headers() {
  
  var results_sheet = SpreadsheetApp.openByUrl(GOOGLE_SHEET_URL).getSheetByName(GOOGLE_SHEET_RESULTS_TAB_NAME);
  results_sheet.appendRow(["AUDIT_DATE", "ID", "OLD_NAME", "RENAMED_TO", "URL"]);
  
}
/*
*
* run this when you want to clear the tokens so you can run the loop
*
*/
function job_delete_token_and_reset_run_history() {

  var scriptProperties = PropertiesService.getScriptProperties();
  
  PropertiesService.getScriptProperties().deleteProperty('continuationToken'); 
  scriptProperties.setProperty('alreadyRun', 'false');
}

/*
*
* big function, that should probably be split into multiple smaller things
* This is the one to run on a schedule (or ad-hoc) and takes care of everything.
*
*/
function job_find_and_rename_files() {
  
  Logger.log('starting lookup'); 
  
  var scriptProperties = PropertiesService.getScriptProperties();
  
  var alreadyRun = scriptProperties.getProperty('alreadyRun');
  
  if(alreadyRun == "true") {
    Logger.log('already run: ' + alreadyRun); 
    return;
  } 
  
  var start = new Date();
  var audit_timestamp = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  
  var continuationToken,files,filesFromToken,fileIterator,thisFile;//Declare variable names
  
  var arrayAllFileNames = [];//Create an empty array and assign it to this variable name
  
  var existingToken = scriptProperties.getProperty('continuationToken');
  
  Logger.log(existingToken);
  
  if(existingToken == null) {

    // Use the searchFiles() method and use the equivalent search of "owner:me" to only find files owned by the person running this script
    var files = DriveApp.getFilesByName(SEARCH_TERM);
    continuationToken = files.getContinuationToken();//Get the continuation token
    Logger.log("Token (Continuation): " + continuationToken);

  } else {
    continuationToken = existingToken; //Get the continuation token that was already stored
    Logger.log("Token (Existing): " + continuationToken);
    
  }
  
  scriptProperties.setProperty('continuationToken', continuationToken);
  
  //Utilities.sleep(1);//Pause the code for 1ms seconds
  
  filesFromToken = DriveApp.continueFileIterator(scriptProperties.getProperty('continuationToken'));//Get the original files stored in the token
  files = null;//Delete the files that were stored in the original variable, to prove that the continuation token is working
  
  var newRow = [];
  var rowsToWrite = [];
  
  var ss = SpreadsheetApp.openByUrl(GOOGLE_SHEET_URL).getSheetByName(GOOGLE_SHEET_RESULTS_TAB_NAME);
  
  while (filesFromToken.hasNext()) {//If there is a next file, then continue looping
    
    if (isTimeUp_(start)) {
      Logger.log("Time up");
      break;
    }
    
    var thisFile = filesFromToken.next();//Get the next file

    // get the file type, so if it's a DOC or a SHEET (or something we can open and get the contents of), we can process and rename it.
    var thisFileId = thisFile.getId();
    var name = thisFile.getName();
    var type = thisFile.getMimeType();
    var created = thisFile.getDateCreated();
    var created_formatted_time = Utilities.formatDate(created, "UTC", "yyyy-MM-dd");
    var url = thisFile.getUrl();
    var new_file_name = '';
    var new_file_name_prefix = created_formatted_time + '___';
    var new_file_name_middle = '';

    // per https://developers.google.com/drive/api/v3/mime-types

    switch(type) {
      case 'application/vnd.google-apps.document':
            // go into the DOC and rename based on the first heading/contents etc
            var doc = DocumentApp.openById(thisFileId);
            var header = doc.getHeader();
            var body = doc.getBody();
            if(header == null) {
              Logger.log('Header is null');
              if(body == null) {
                  Logger.log('Body is NULL');
                  new_file_name_middle = 'BLANK DOCUMENT';
                } else {
                  Logger.log('Getting Body Text');
                  var body_text = body.getText();
                  new_file_name_middle = body_text.substring(0,NUM_CHARACTERS_FROM_FILE_CONTENTS);
                  Logger.log(new_file_name_middle);
                }
            } else {
              Logger.log('Getting Header Text');
              var header_text = header.getText();
              new_file_name_middle = header_text.substring(0,NUM_CHARACTERS_FROM_FILE_CONTENTS);
            }
            break;
      case 'application/vnd.google-apps.spreadsheet':
            // go into the Sheet and rename based on the first heading/contents etc
            new_file_name_middle = 'sheet';
            // go into the DOC and rename based on the first heading/contents etc
            var thisSheet = SpreadsheetApp.openById(thisFileId).getActiveSheet();
            var first_cell_value = thisSheet.getRange('A1').getValue().trim();
            Logger.log(first_cell_value);
            if(first_cell_value == null || first_cell_value == '') {
              Logger.log('A1 is empty, use tab/sheet name instead');
              new_file_name_middle = thisSheet.getName().substring(0,NUM_CHARACTERS_FROM_FILE_CONTENTS);
            } else {
              Logger.log('A1 is not empty so use this value');
              new_file_name_middle = first_cell_value.substring(0,NUM_CHARACTERS_FROM_FILE_CONTENTS);
            }
            break;
      default: 
        // do not know what kind of file this is, or it's a type we can't access and use the contents
        Logger.log('Unknown or cannot open file type to use its contents.');
        new_file_name_middle = 'Untitled';

    }

    new_file_name = new_file_name_prefix + new_file_name_middle + NEW_FILE_NAME_SUFFIX;
    Logger.log(new_file_name);

    // set new file name
    thisFile.setName(new_file_name);
    //write change info to sheet for auditing (and undoing if needed)
    var newRow = [audit_timestamp, thisFileId, name, new_file_name, url];
    
    // add to row array instead of append because append is SLOOOOOWWWWW
    rowsToWrite.push(newRow);
    
    // Save our place by setting the token in our script properties
    // this is the magic that allows us to set this to run every minute/hour depending on the timeout value
    if(filesFromToken.hasNext()){
      var continuationToken = filesFromToken.getContinuationToken();
      scriptProperties.setProperty('continuationToken', continuationToken);
    } else {
      // Delete the token and store that we are complete
      PropertiesService.getScriptProperties().deleteProperty('continuationToken');
      scriptProperties.setProperty('alreadyRun', "true");
    }
    
  };
  
  ss.getRange(ss.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  
  if(!filesFromToken.hasNext()) {
   
    var templ = HtmlService
      .createTemplateFromFile('email');
  
    templ.sheetUrl = GOOGLE_SHEET_URL;
    
    var message = templ.evaluate().getContent();
    
    MailApp.sendEmail({
      to: NOTIFICATION_RECIPIENTS,
      subject: 'Google Drive Rename Untitled Files Complete',
      htmlBody: message
    });
    
  }
  
};

/*
* quick function to see if the timeout value has been reached (therefore stop the loop)
*/
function isTimeUp_(start) {
  var now = new Date();
  return now.getTime() - start.getTime() > TIMEOUT_VALUE_MS; // milliseconds
}
