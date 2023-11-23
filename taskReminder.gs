var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var configSheet = spreadSheet.getSheetByName('Config');
var sheeturl = spreadSheet.getUrl();

function taskReminder() {
  var dataSheet = spreadSheet.getSheetByName('Data');
  var lastRow = dataSheet.getLastRow();
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')

  //Fetching config informations
  var reminder1 = configSheet.getRange('B3').getValue();
  var reminder2 = configSheet.getRange('B4').getValue();
  var mattermostChannelID = configSheet.getRange('B8').getValue();
  var emailId = configSheet.getRange('B11').getValue();
  var mattermostAccessToken = configSheet.getRange('B9').getValue();

  var emailID = verifyEmails(emailId)

  //Iterate over data in the Data sheet
  for(var i=2 ; i <= lastRow ; i++){

    // Fetching row values
    var row = dataSheet.getRange(i, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    var itemName = row[1];
    var expiryDate = row[2];
    var comments = row[3];
    var status = row[4];

    var remindertimeDiff = expiryDate.getTime() - currentDate.getTime();
    var reminderDays = Math.ceil(remindertimeDiff / (1000 * 60 * 60 * 24))
    console.log(reminderDays);

    if ((status === '' && reminderDays > reminder1) || (reminderDays > 0 && status === 'Expired')){
      //New or Renewed Items
      dataSheet.getRange(i, 5).setValue('Watching Expiry Date');
    }

    else if (reminderDays > 0 && reminderDays <= reminder1 && (status === '' || status === 'Watching Expiry Date')) {
      //Remainder 1
      createMattermostPost(mattermostAccessToken,mattermostChannelID,itemName,reminderDays,comments)
      sendEmail(emailID,itemName,reminderDays,comments);
      dataSheet.getRange(i, 5).setValue('First Reminder sent');
    }
    
    else if (reminderDays <= reminder2 && status === 'First Reminder sent') {
      //Remainder 2
      createMattermostPost(mattermostAccessToken,mattermostChannelID,itemName,reminderDays,comments)
      sendEmail(emailID,itemName,reminderDays,comments);
      dataSheet.getRange(i, 5).setValue('Second Reminder sent');
    }

    else if (reminderDays == 0 || reminderDays == -7 || reminderDays == -14 || reminderDays == -21){
      //Expired
      createMattermostPost(mattermostAccessToken,mattermostChannelID,itemName,reminderDays,comments)
      sendEmail(emailID,itemName,reminderDays,comments);
      dataSheet.getRange(i, 5).setValue('Expired');
    }
  }
}


function createMattermostPost(mattermost_access_token,mattermostChannel_ID,itemname,expirebefore,comment) {

  var mattermostUrl = configSheet.getRange('B7').getValue();

  if(mattermostChannel_ID != '' || mattermost_access_token != '' || mattermostUrl != '' ){
    var channelID = mattermostChannel_ID; // Replace with the ID of the channel where you want to post
    var message_content = messageContent(itemname,expirebefore,comment,sheeturl);
    
    var payload = {
      channel_id: channelID,
      message: message_content.msg
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + mattermost_access_token
      },
      payload: JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(mattermostUrl, options);
    
    //Logger.log(response.getContentText());
  }
}


function sendEmail(emails, itemname, expirebefore, comment){
  if (emails.length > 0){
    var toEmail = emails[0];
    var ccEmail = emails.slice(1).join(',');
    var message_content = messageContent(itemname,expirebefore,comment,sheeturl);
    MailApp.sendEmail({
      to: `${toEmail}`,
      subject: message_content.sub,
      body: message_content.msg,
      cc: `${ccEmail}`
    })
  }
}


function createDailyTimeTrigger() {
  var emailID = configSheet.getRange('B11').getValue();
  var invalidEmails = verifyEmails(emailID, true)

  if(invalidEmails.length > 0){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      "Invalid Emails Detected",
      "The following emails are not in a valid format:\n\n" +
        invalidEmails.join("\n") +
        "\n\nDo you want to continue? NOTE: If you click 'No', the trigger won't be updated",
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.NO) {
      return;
  }
  }  
  
  checkAndRemoveTrigger();
  var triggerHour = configSheet.getRange('B5').getValue();
  ScriptApp.newTrigger('taskReminder')
  .timeBased()
  .atHour(triggerHour)
  .nearMinute(30)
  .everyDays(1)
  .create();
}


function checkAndRemoveTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  // Loop through all triggers and check if the specified function is associated with any trigger
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if (trigger.getHandlerFunction() === 'taskReminder') {
      // Trigger with the specified function name exists
      ScriptApp.deleteTrigger(trigger);
    }
  }
}


function messageContent(itemname,expirebefore,comment,sheeturl){
if(expirebefore <= 0){
    var subject =  `${itemname} expired before ${-expirebefore} day(s) ago`;
    var message = `The ${itemname} has been expired before ${-expirebefore} day(s) ago.`;
  }
  else{
    var subject =  `${itemname} expires in ${expirebefore} day(s)`;
    var message = `The ${itemname} is going to expire in ${expirebefore} day(s).`;
  }

  message = message + `Please renew.\n${comment}\nUpdate the spreadsheet with new date after renewal ${sheeturl}`;

  return {
    sub: subject,
    msg: message};
}


function verifyEmails(emailList,returnInvalid = false) {
  var emails = emailList.split(",");

  // Regular expression for email validation
  var emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;

  var invalidEmails = [];
  var validEmails = [];

  // Iterate through the emails and check if they match the regex
  for (var i = 0; i < emails.length; i++) {
    if (!emailRegex.test(emails[i].trim())) {
      invalidEmails.push(emails[i]);
    }
    else {
      validEmails.push(emails[i]);
    }
  }

  if(!returnInvalid){
    return validEmails;
  }
  else{
    return invalidEmails;
  }
}
