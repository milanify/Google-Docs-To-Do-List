/**
 * Instructions on how to use this add-on:
 * Select 'Initialize' from the Add-ons menu to begin.
 * Make sure formatting is preserved by only using the 'New note' and 'Delete' buttons from the sidebar,
 * which is shown by clicking 'View notes' in the Add-ons menu.
 */

/**
 * Declare constants
 */
var body = DocumentApp.getActiveDocument().getBody();
var style = {};
style[DocumentApp.Attribute.SPACING_BEFORE] = 15;
style[DocumentApp.Attribute.SPACING_AFTER] = 15;

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when document is opened
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Initialize', 'showInitializationAlert')
      .addItem('View notes', 'showSidebar')
      .addToUi();
}

/**
 * Show dialog box that appears when running the add-on for the first time
 */
function showInitializationAlert() {
  var ui = DocumentApp.getUi();

  var result = ui.alert(
    'Clear all contents of this document?',
     'Selecting \'Yes\' is required to initialize this add-on for the first time.',
      ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    clearAllContents();
    insertHorizontalLine();
    ui.alert('Contents cleared. \n\nBegin typing above the horizontal line.');
  } else {
    ui.alert('You selected \'No\', the contents of this document will be preserved. \n\nCreate a new blank word document, and then run this add-on.');
  }
}

/**
 * Delete everything in the document
 */
function clearAllContents() {
  body.clear();
}

/**
 * Insert a horizontal line and a blank line on top of it, using the style specified
 */
function insertHorizontalLine() {
  body.insertHorizontalRule(0);
  var par = body.insertParagraph(0, '');
  par.setAttributes(style);
  showSidebar();
}

/**
 * Show HTML sidebar containing note list and buttons
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
  .evaluate()
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .setTitle('To-Do List')

  DocumentApp.getUi().showSidebar(html);
}

/**
 * Retrieve all the text content in the document
 *
 * @return The document text
 */
function getAllDocumentText() {
  return DocumentApp.getActiveDocument().getBody().getParagraphs();
}

/**
 * Organize the each note (text between horizontal lines) into an array
 * From top to bottom, add each line of text as an item in a temporary array
 * Once a horizontal line is detected, join all the text in the temp array
 *
 * @return {Array} All notes between horizontal lines
 */
function getAllTextBetweenHorizontalRules() {
  var textBetweenHorizontalRules = [];
  var tempArray = [];

  pars = getAllDocumentText();
  pars.forEach(function(e) {
    var isHorizontalRule = e.findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
    if (!isHorizontalRule && tempArray.length == 0) {
      tempArray.push(e.getText());
    } else if (!isHorizontalRule && tempArray.length > 0) {
      tempArray.push('\n' + e.getText());
    } else {
      textBetweenHorizontalRules.push(tempArray.join(""));
      tempArray = [];
    }
  })
  return textBetweenHorizontalRules;
}

/**
 * Delete notes based on which checkboxes on the HTML sidebar were selected
 * Special case for the  first note, which is always deleted last
 * If the above is not done, then deleting does not work due to deleting the head, 0th index
 * Keeps track of how many horizontal lines were encountered and checks if each were selected
 *
 * @param {Array} checkboxData contains booleans of which checkboxes are selected
 */
function deleteCheckboxSelection(checkboxData) {
  pars = getAllDocumentText();
  var horizontalLineCount = 0;
  var isDeleteFirst = false;
  var indexOfDeleteFirst = 0;

  for (var i = 0; i < pars.length; i++) {
    var isHorizontalRule = pars[i].findElement(DocumentApp.ElementType.HORIZONTAL_RULE);

    if (isHorizontalRule && checkboxData[horizontalLineCount] && horizontalLineCount == 0) {
      isDeleteFirst = true;
      indexOfDeleteFirst = i;
      horizontalLineCount++;
    } else if (isHorizontalRule && checkboxData[horizontalLineCount]) {
      deleteNote(pars, i);
      horizontalLineCount++;
    } else if (isHorizontalRule) {
      horizontalLineCount++;
    }
 }

  if (isDeleteFirst) {
    deleteNote(pars, indexOfDeleteFirst);
  }
}

/**
 * Delete the note by removing text
 * Always remove the startIndex, which is the horizontal line separator below each note
 * Keep on deleting each line of text until the next horizontal line is detected or the top of the document is reached
 * Using [startIndex-i] because we are deleting from bottom to top, the very first line of the document is index 0
 *
 * @param {Array} textData contains all of the document's text
 * @param {Integer} startIndex is where to start deleting from, which is always a horizontal line
 */
function deleteNote(textData, startIndex) {
  var i = 1;
  textData[startIndex].editAsText().removeFromParent();

  while (!(textData[startIndex-i].findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) && (startIndex-i >= 0)) {
    textData[startIndex-i].editAsText().removeFromParent();
    i++;
  }
}
