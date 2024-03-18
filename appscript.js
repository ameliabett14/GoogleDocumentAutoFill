function onOpen() {
const ui = SpreadsheetApp.getUi();
const menu = ui.createMenu("AutoFill Docs");
menu.addItem('Create New Docs', 'createNewGoogleDocs');
menu.addToUi();

}

//This is to make sure that the ZIP code stays in format. It may not be necessary in your sheet.
function setNumberFormat() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet.getRange("F:F");
  range.setNumberFormat("00000");
}

function createNewGoogleDocs() {

const docTemplate = DriveApp.getFileById('INSERT TEMPLATE ID HERE');
const destinationFolder = DriveApp.getFolderById('INSERT TARGET FOLDER ID FOR STORAGE HERE')
//Your sheet should be named Current for data stored
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Current');

const rows = sheet.getDataRange().getValues();

Logger.log(rows);

rows.forEach(function(row,index){
if (index === 0) return;
//This will set the condition for a row to be skipped when a document is made
if (row[8]) return;

const copy = docTemplate.makeCopy(row[1]+' Document', destinationFolder);
const doc = DocumentApp.openById(copy.getId())
const body = doc.getBody();
const date = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy");

//These variables will be different based on what you use/need  
body.replaceText('{{date}}', date);
body.replaceText('{{Company}}', row[1]);
body.replaceText('{{Address}}', row[2]);
body.replaceText('{{City}}', row[3]);
body.replaceText('{{StateABV}}', row[4]);
body.replaceText('{{Zip}}', row[5]);
body.replaceText('{{Superintendent}}', row[6]);
body.replaceText('{{Website}}', row[7]);

doc.saveAndClose();
const identification = doc.getId();
//This allows you to click the url to download the file directly from the sheet
sheet.getRange(index + 1, 9).setValue("https://docs.google.com/document/d/"+identification+"/export?format=pdf");

})

}
