
var savePDF = true; // le PDF sera enregistré dans google drive en + d'être envoyé par mail.
var saveToRootFolder = true; // l'enregistrement se fait dans le dossier root.
var mailingList = ["mailopif@yopmail.com", "mechaRaspoutine@protonmail.com"]; // liste des mails ou le PDF sera envoyé.





/**
 * utils
 */
function padded(n, min=2, pad='0') {
	return n.toString().padStart(min, pad);
}
function formatDate(d) {
  return `${d.getFullYear()}-${padded(d.getMonth()+1)}-${padded(d.getDate())}_${padded(d.getHours())}-${padded(d.getMinutes())}-${padded(d.getSeconds())}`;
}

/**
 * 
 */
function _exportBlob(blob, fileName, spreadsheet) {
  //blob = blob.setName(fileName);
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next();
  var pdfFile = folder.createFile(blob);
  
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful');
}

/**
 * 
 */
function _getAsBlob(url, sheet, range) {
  var rangeParam = '';
  var sheetParam = '';
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId();
  }

  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam;
      
  Logger.log('exportUrl=' + exportUrl);
  var response;
  for (var i = 0; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    });
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000);
    } else {
      break;
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.');
  }
  
  return response.getBlob();
}

function _sendBlobMail(blob, pdfName, email, subject, htmlbody) {
  if (!email) return;
  var mailOptions = { attachments:blob, htmlBody:htmlbody };
  MailApp.sendEmail(
      email, 
      subject + " (" + pdfName +")", 
      "html content only", 
      mailOptions);
  MailApp.sendEmail(
      Session.getActiveUser().getEmail(), 
      "FRWD " + subject + " (" + pdfName +")", 
      "html content only", 
      mailOptions);
}

//sendSpreadsheetToPdf(sheetNumber, pdfName, email,subject, htmlbody) {
//sendSpreadsheetToPdf(0, shName, sh.getRange('A1').getValue(),"test email with the adress in cell A1 ", "This is it !")


/**
 * 
 */
function exportAsPDF() {
  var dateString = formatDate(new Date());
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var name = spreadsheet.getName() + "_" + dateString;
  var blob = _getAsBlob(spreadsheet.getUrl());
  blob = blob.setName(name);
  

  for (var i = 0; i < mailingList.length; i++) {
    var email = mailingList[i];
    _sendBlobMail(blob, name, email, "Test d'envoi de mail depuis google spreadsheets", 'Voici le PDF "' + name + '" en pièce jointe.');
  }
  if (savePDF) {
    _exportBlob(blob, name, spreadsheet);
  }
}


/**
 * 
 */
function exportCurrentSheetAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet);
  _exportBlob(blob, currentSheet.getName(), spreadsheet);
}


/**
 * 
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('script PDF')
      .addItem('Envoyer le document par mail', 'sendMail')
      .addToUi();
}

function sendMail() {
  exportAsPDF();
  Browser.msgBox("Mail envoyé!");
}




/** 
 * references:
 * https://xfanatical.com/blog/print-google-sheet-as-pdf-using-apps-script/
 * https://stackoverflow.com/questions/45781031/google-script-send-active-sheet-as-pdf-to-email-listed-in-cell
 */


