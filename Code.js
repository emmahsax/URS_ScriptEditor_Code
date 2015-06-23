/**
 * This code was originally used in UMM's CSci 3601, Fall 2014 class as an experimentation with Google Drive and 
 * the spreadsheet script editor. Now, this is to be used by Emma Sax during her implementation of the automatic
 * making of google documents whenever a google form is submitted. Much of (or most of) this code was put together
 * by Hongya Zhou, with bits and pieces of other resources as well:
 * - Johninio's code http://www.google.com/support/forum/p/apps-script/thread?tid=032262c2831acb66&hl=en
 * - https://developers.google.com/apps-script/service_spreadsheet
 * - and a couple other places, too
 *
 * If the form ever changes or if anything ever happens the things that will need to be changed are the 
 * ID of the current spreadsheet (can use getID function for this), possibly templateID, column numbers and row
 * numbers, and the submissionFolderID if that's being used.
 *
 * Things for Emma to ask Stephanie:
 * - changing form to ask for all names in one line
 * - asking for co-presenter emails separately
 * - asking for t-shirt information simply as one question
 * - what parameters to have on the google documents and in what order
 */

// ID of this spreadsheet: 1IkgJOguzDdZnF7MtzWWER1uXCgsBlbors7pQLeT_cHg

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {name : "Get Spreadsheet ID", functionName : "getID"},
    {name : "Create Doc from Last Row", functionName : "createDocFromSheet"}
  ];
  spreadsheet.addMenu("Script Center Menu", entries);
}

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 */
// function readRows() {
//   var sheet = SpreadsheetApp.getActiveSheet();
//   var rows = sheet.getDataRange();
//   var numRows = rows.getNumRows();
//   var values = rows.getValues();

//   for (var i = 0; i <= numRows - 1; i++) {
//     var row = values[i];
//     Logger.log(row);
//   }
// }

/**
 * Gathers the ID of the spreadsheet and can return it in a message box
 */
function getID() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
}

/**
 * Returns the value of the last column in the last row, weirdly appending tons of commas to the end
 */
function getLastRow() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = data.getDataRange().getNumRows();
  var cols = data.getDataRange().getNumColumns();
  var lastRow = data.getRange(rows, cols, 1, cols).getValues();
  return lastRow;
}

/**
 * Uses the current spreadsheet and a template document to make a new document per row (upon submission of form) with all categories filled in.
 * The column numbers might have to change a fair amount depending on the actual spreadsheet when implemented. Also, whenever the form or spreadsheet columns change,
 * there is hardcoding involved, so someone would need to come in to change it in the code, but it isn't rocket science. Lastly, this code is under the assumption
 * that the form asks for names of people in one box, and then the emails following in one box.
 *
 * In form, change co-presenter and presenter naming to be one box for name, one for email for every person
 * In form, change t-shirt to simply be checkboxes for everyone regardless of their role in the process
 */
function createDocFromSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // gets current spreadsheet
  var currentSpreadsheet = spreadsheet.getActiveSheet(); // gets another version of current spreadsheet
  var spreadsheetData = currentSpreadsheet.getRange(currentSpreadsheet.getLastRow(), 1, currentSpreadsheet.getLastRow(), currentSpreadsheet.getLastColumn()).getValues(); // gathers spreadsheet data of last row

  var column = spreadsheetData[0]; // names column a column of spreadsheet data, so to be used with an input of a number

  var newDoc = DocumentApp.create("URS Submission - " + column[8] + " (" + column[3] + ")"); // name of new document
  //var newDocFile = DocsList.getFileById(newDoc.getId()); // getting file ID from the newDoc
  var setID = currentSpreadsheet.getDataRange().getCell(currentSpreadsheet.getLastRow(), currentSpreadsheet.getLastColumn()); //doing something not quite sure what

  setID.setValue(newDoc.getId()); // getting ID of newDoc

  var templateID = "1C2Q20zjkCgqo8s5p6-6Wa_KPVqYF8zXIwwDsHCi3Np4"; // this ID is for the template of the documents
  var newDocFromTemplateID = DocsList.getFileById(templateID).makeCopy().getId(); // making a copy of the template to be used with the newDoc
  var docFromTemplate = DocumentApp.openById(newDocFromTemplateID); // opening template
  var body = docFromTemplate.getActiveSection(); // getting the body of the new template
 
  // adding appropriate column values to the newDoc
  // IMPORTANT: think of timestamp column as column[0]
  body.insertParagraph(2, column[8] + ", " + column[9]); // primary presenter information
  body.insertParagraph(6, column[10] + ", " + column[11]); // co-presenter 1 information
  body.insertParagraph(10, column[12] + ", " + column[13]); // co-presenter 2 information
  body.insertParagraph(14, column[16] + ", " + column[17]); // faculty sponsor information
  body.insertParagraph(18, column[3]); // title
  body.insertParagraph(22, column[5]); // abstract
  body.insertParagraph(26, column[14]); // discipline
  body.insertParagraph(30, column[4]); // format for submission
  body.insertParagraph(34, column[6]); // type of presentation
  body.insertParagraph(38, column[18]); // opening ceremony
  body.insertParagraph(42, column[7]); // willingness to change presentation type
  body.insertParagraph(46, column[15]); // sponsoring funds
  body.insertParagraph(50, column[19]); // media services
  body.insertParagraph(54, column[20]); // room location
  body.insertParagraph(58, column[21] + ", " + column[22]); // t-shirt information
  body.insertParagraph(62, column[23]); // additional comments

  docFromTemplate.saveAndClose(); // save and close newDoc
  
  appendToDoc(docFromTemplate, newDoc); // append template copy to newDoc

  DocsList.getFileById(newDocFromTemplateID).setTrashed(true); // delete temporary template file

  //var submissionFolder = DocsList.getFolderById("0B6MWcSznaGDRci1GWUF5WUZNNGM"); // this ID is for the folder of the generated documents, found at the end of the URL
  //newDocFile.addToFolder(submissionFolder); // adds the newDoc to the submissionFolder, so not in some random place
  //newDocFile.removeFromFolder(newDocFile.getParents()[0]); // remove copy from root of Drive
  
  spreadsheet.toast("Document Created"); // show message on spreadsheet that this function is over
}

/**
 * Iterates across the elements in the template source document, and then calls appendElementToDoc to
 * append each element to the new destination document
 */
function appendToDoc(source, destination) {
  for (var i = 0; i < source.getNumChildren(); i++) {
    appendElementToDoc(destination, source.getChild(i));
  }
}
 
/**
 * Takes a document and an object, and appends the object to the document, under the assumption
 * that the object is of type paragraph. The original code has options for handling paragraphs, tables, etc 
 * differently. For the purposes of this function, we only need to work with paragraphs.
 */
function appendElementToDoc(document, object) {
  var element = object.copy(); // need to remove or can't append
  document.appendParagraph(element);
}
