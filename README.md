# Googlesheet-email-spread


/**
 * Sends emails with data from the current spreadsheet.
 */
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // Fetch the range of cells A2:C
  var dataRange = sheet.getRange("Form Responses 1!A94:D94");
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i in data) {
    var row = data[i];
    var emailSent = row[0];
    
    if (!emailSent) {
      var status = row[1];
      var timestamp = row[2];
      var emailAddress = row[3];
      var name = emailAddress.split("@")[0];
      var message =
          "Dear " + name + "<br/><br/>" +
          "Kindly note that your medical refund transaction you submitted at: " + timestamp + ", has been: <b>" + status + "</b>.<br/>" +
          "If you need any further help, don’t hestitate to cotact us.<br/><br/>" +
          "Regards,<br/>" +
          "Incorta HR Team";
      var subject = "Medical Refund Status - " + timestamp;
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message
      });
    }
  }
}  

