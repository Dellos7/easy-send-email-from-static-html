/*******************************************************************************************
 * Based on https://github.com/dwyl/html-form-send-email-via-google-script-without-server  *
*******************************************************************************************/

// if you want to store your email server-side (hidden), uncomment the next line
//var TO_ADDRESS = "";

function formatMailBody(obj, emailTemplateName) {
  var _obj = JSON.parse(JSON.stringify(obj));
  var avoid = {
    'success_url': _obj['success_url'],
    'error_url':  _obj['error_url'],
    'send_email_to': _obj['send_email_to'],
    'send_email_copy': _obj['send_email_copy'],
    'send_email_to_name': _obj['send_email_to_name'],
    'bitc_hp': _obj['bitc_hp'],
    'contact_form_url': _obj['contact_form_url']
  };
  var htmlTemplate = HtmlService.createTemplateFromFile(emailTemplateName);
  htmlTemplate.data = _obj;
  htmlTemplate.avoid = avoid;
  return htmlTemplate.evaluate();
}

function doPost(e) {

  try {
    Logger.log(e); // the Google Script version of console.log see: Class Logger
    record_data(e);
    
    // shorter name for form data
    var mailData = e.parameters;
    
    if( String(mailData.bitc_hp) ) {
      throw "No spam please!";
    }
    
    // determine recepient of the email
    // if you have your email uncommented above, it uses that `TO_ADDRESS`
    // otherwise, it defaults to the email provided by the form's data attribute
    var sendEmailTo = (typeof TO_ADDRESS !== "undefined") ? TO_ADDRESS : ( mailData.send_email_to ? String(mailData.send_email_to) : Session.getEffectiveUser().getEmail() );
    
    Logger.log(sendEmailTo);
    
    var sendEmailData = { 
      to: String(sendEmailTo),
      subject: "New message from your contact form at " + new Date().toLocaleTimeString() + " " + new Date().toLocaleDateString(),
      replyTo: ( mailData.email ? String(mailData.email) : sendEmailTo ),
      htmlBody: formatMailBody(mailData, "email-contact-template").getContent()
    };
    
    if( mailData.name ) {
      sendEmailData.name = String(mailData.name);
    }
    
    MailApp.sendEmail(sendEmailData);
    
    if( mailData.send_email_copy ) {
      sendEmailCopy( mailData, sendEmailTo ); 
    }
    
    if( mailData['success_url'] ) {
      return HtmlService.createHtmlOutput(
        "<script>window.top.location.href='" + mailData['success_url'] +"';</script>"
      )
    }

    return ContentService    // return json success results
          .createTextOutput(
            JSON.stringify({"result":"success",
                            "data": JSON.stringify(e.parameters), "htmlOutput": formatMailBody(mailData, "email-contact-template").getContent() }))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(error) { // if error return this
    Logger.log(error);
    if( mailData['error_url'] ) {
      return HtmlService.createHtmlOutput(
        "<script>window.top.location.href='" + mailData['error_url'] +"';</script>"
      )
    }
    return ContentService
    .createTextOutput(JSON.stringify({"result":"error", "error": error, "data": e.parameters, "tcontent": formatMailBody(mailData, "email-contact-template").getContent()}))
          .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * record_data inserts the data received from the html form submission
 * e is the data received from the POST
 */
function record_data(e) {
  Logger.log(JSON.stringify(e)); // log the POST data in case we need to debug it
  try {
    var doc     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = doc.getSheetByName('responses'); // select the 'responses' sheet by default
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row     = [ new Date() ]; // first element in the row should always be a timestamp
    // loop through the header columns
    for (var i = 1; i < headers.length; i++) { // start at 1 to avoid Timestamp column
      if(headers[i].length > 0) {
        row.push(e.parameter[headers[i]]); // add data to row
      }
    }
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
  }
  catch(error) {
    Logger.log(e);
  }
  finally {
    return;
  }

}

function sendEmailCopy(mailData, replyTo) {
  var sendEmailData = { 
      to: String(mailData.email),
      subject: "Form contact submitted at " + new Date().toLocaleTimeString() + " " + new Date().toLocaleDateString(),
      replyTo: replyTo,
      htmlBody: formatMailBody(mailData, "response-email-contact-template").getContent()
    };
    
    if( mailData.send_email_to_name ) {
      sendEmailData.name = String(mailData.send_email_to_name);
    }
    
    MailApp.sendEmail(sendEmailData);
}


function formatMailBody_email_contact_template_Test() {
  var obj = {
    name: 'David',
    email: 'example@gmail.com',
    message: 'Hello!',
    success_url: 'http://example.com/success',
    error_url: 'http://example.com/error',
    send_email_to_name: 'BITC ;)'
  };
  var htmlOutput = formatMailBody(obj, "email-contact-template");
  var bp = 'bp';
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, 'Test format HTML Body');
}

function formatMailBody_response_email_contact_template_Test() {
  var obj = {
    name: 'David',
    email: 'example@gmail.com',
    message: 'Hello!',
    success_url: 'http://example.com/success',
    error_url: 'http://example.com/error',
    send_email_to_name: 'BITC ;)'
  };
  var htmlOutput = formatMailBody(obj, "response-email-contact-template");
  var bp = 'bp';
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, 'Test format HTML Body');
}