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
 */

// ID of this spreadsheet: 1QoZFE6tM14d4CeJq23LiRvzGZnmSIv_gH33-mB0yJH0

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {name : "Read Data", functionName : "readRows"},
    {name : "Get Spreadsheet ID", functionName : "getID"},
    {name : "Get Last Row", functionName : "getLastRow"},
    {name : "Create Doc from Last Row", functionName : "createDocFromSheet"}
  ];
  spreadsheet.addMenu("Script Center Menu", entries);
};

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

/**
 * Gathers the ID of the spreadsheet and can return it in a message box
 */
function getID() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
};

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
 * Uses the current spreadsheet and a template document to make a new document per row (upon submission of form) with
 * all categories filled in. The column numbers might have to change a fair amount depending on the actual spreadsheet
 * when implemented. Also, whenever the form or spreadsheet columns change, there is hard-coding involved, so someone
 * would need to come in to change it in the code, but it isn't rocket science. Lastly, this code is under the assumption
 * that the form asks for names of people in one box, and then the emails following in one box.
 */
function createDocFromSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // current spreadsheet
  var lastRow = spreadsheet.getLastRow(); // last row of spreadsheet
  var lastColumn = spreadsheet.getLastColumn(); // last column of spreadsheet
  var spreadsheetData = spreadsheet.getRange(lastRow, 1, lastRow, lastColumn).getValues(); // entire data of last row

  var newDoc = DocumentApp.create("2016 URS - " + column[9] + " " + column[8]); // new document to be created
  var newDocFile = DriveApp.getFileById(newDoc.getId()); // ID of new file
  // IMPORTANT: hard-coded, do not change unless folder changes; ID found at end of URL
  // var submissionFolder = DriveApp.getFolderById("0B-4Ru4UajECXdGFhMmJ0N1I5R0U"); // ID of folder for generated documents

  // IMPORTANT: hard-coded, do not change unless template document changes; ID found at end of URL 
  var templateID = "1FTXNICzBXEhFUExSZqdS9jxGenD2RcNDCYZSZd8XLsk"; // ID of the template of the documents
  var newDocFromTemplateID = DriveApp.getFileById(templateID).makeCopy().getId(); // copy of the template
  var docFromTemplate = DocumentApp.openById(newDocFromTemplateID); // opened template copy
  var body = docFromTemplate.getActiveSection(); // body of the template copy

  var column = spreadsheetData[0]; // makes column a column of the spreadsheet (to be used with an input of a number)
 
  // adding appropriate column values to the template copy; think of timestamp column as column[0]
  body.insertParagraph(2, column[16]); // discipline
  body.insertParagraph(6, column[9] + " " + column[8]); // primary presenter name
  body.insertParagraph(10, column[11] + " " + column[10] + ", " + column[13] + " " + column[12]); // co-presenter names
  body.insertParagraph(14, column[18]); // faculty sponsor name
  body.insertParagraph(18, column[3]); // title
  body.insertParagraph(22, column[4]); // format for proposal
  body.insertParagraph(26, column[5]); // abstract
  body.insertParagraph(30, column[20]); // feature presentation
  body.insertParagraph(34, column[7]); // willingness to change presentation type
  body.insertParagraph(38, column[25]); // additional comments
  /* Optional other columns to add
   * body.insertParagraph(26, column[14]); // type of presentation
   * body.insertParagraph(46, column[15]); // sponsoring funds
   * body.insertParagraph(50, column[19]); // media services
   * body.insertParagraph(54, column[20]); // room location
   * body.insertParagraph(58, column[21] + ", " + column[22]); // t-shirt information
   */
  
  docFromTemplate.saveAndClose(); // saving and closing template copy
  appendToDoc(docFromTemplate, newDoc); // merging template copy with newDoc
  DriveApp.getFileById(newDocFromTemplateID).setTrashed(true); // deleting template copy
  
  /* IMPORTANT: the code below is finicky and only sometimes works; use with caution and test carefully
   * spreadsheet.getDataRange().getCell(lastRow, lastColumn).setValue(newDoc.getId()); // putting ID into spreadsheet
   * newDocFile.addToFolder(submissionFolder); // adding the newDoc to the submissionFolder
   * newDocFile.removeFromFolder(newDocFile.getParents()[0]); // removing newDoc from root of Drive
   */
  
  spreadsheet.toast("Document Created"); // show message on spreadsheet that this function is over
}

/**
 * Iterates across the elements in the template source document, and then calls appendElementToDoc to
 * append each element to the new destination document
 */
function appendToDoc(source, destination) {
  for (var i = 0; i < source.getNumChildren(); i++) {
    appendElementToDoc(destination, source.getChild(i));
  };
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
