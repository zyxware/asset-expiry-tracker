var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var configSheet = spreadSheet.getSheetByName('Config');
var sheeturl = spreadSheet.getUrl();

function taskReminder() {
  var dataSheet = spreadSheet.getSheetByName('Data');
  var lastRow = dataSheet.getLastRow();
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')

  //Fetching config informations
  var reminder1 = configSheet.getRange('B2').getValue();
  var reminder2 = configSheet.getRange('B3').getValue();
  var mattermostChannelID = configSheet.getRange('B5').getValue();
  var emailID = configSheet.getRange('B6').getValue();
  var mattermostAccessToken = configSheet.getRange('B7').getValue();
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

    else if (reminderDays <= 0){
      //Expired
      createMattermostPost(mattermostAccessToken,mattermostChannelID,itemName,reminderDays,comments)
      sendEmail(emailID,itemName,reminderDays,comments);
      dataSheet.getRange(i, 5).setValue('Expired');
    }
  }
}


function createMattermostPost(mattermost_access_token,mattermostChannel_ID,itemname,expirebefore,comment) {
  var mattermostUrl = configSheet.getRange('B4').getValue();; // Replace with your Mattermost API URL
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
  
  Logger.log(response.getContentText());
}


function sendEmail(emailid, itemname, expirebefore, comment){
  var emails = emailid.split(',');
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


function createDailyTimeTrigger() {
  checkAndRemoveTrigger();
  var triggerHour = configSheet.getRange('B8').getValue();
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
