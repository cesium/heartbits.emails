// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';
var FILE_NOT_FOUND = 'FILE_NOT_FOUND';


var EMAIL_SUBJECT = 'Certificado de participação na Hackathon HeartBits 2020';

var LOGO = UrlFetchApp.fetch('https://heartbits.pt/assets/banner.png').getBlob().setName("LOGO");

var BODY = 'Caro(a) participante,\n\nMuito obrigado por teres participado na Hackathon Heartbits 2020!\nEsperamos que tenhas gostado do evento e que tenha sido um bom momento de aprendizem e diversão.\n\nPara terminar a edição deste ano e oficializar a tua participação, enviamos, em anexo, o teu certificado de participação!\n\nBrevemente, iremos contactar-te com informações relativas à entrega/envio do teu Kit.\n\n\nAté já,\nA equipa organizadora da Hackathon HeartBits 2020';

var HTML_BODY = HtmlService.createHtmlOutputFromFile('certificates-body').getBlob().getDataAsString();


function getAttachment(cell, cc){
  cell.setValue(cc + '.pdf');
  var files = DriveApp.getFilesByName(cc + '.pdf');
  var file = files.next();
  if (files.hasNext()){
    throw "More than one file";
  }
  return file;
}

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 70; // Number of rows to process
  var startColumn = 1;
  var numColumns = 6;
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    try {
      var row = data[i];
      var emailAddress = row[1];
      var ccidadao = row[2];
      var emailSent = row[3];
      var fileNotFound = row[4];
      var file = getAttachment(sheet.getRange(startRow + i, 6), ccidadao);
      if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates 
        var message = {
          to: emailAddress,
          cc: 'comunicacao@anem.pt',
          subject: EMAIL_SUBJECT,
          body: BODY,
          htmlBody: HTML_BODY,
          inlineImages: {logo: LOGO},
          attachments: [file.getBlob()]
        };
        MailApp.sendEmail(message);
        sheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);
      }
    } catch (error) {
        sheet.getRange(startRow + i, 5).setValue(error);
    } finally {
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();  
    }
  }
}