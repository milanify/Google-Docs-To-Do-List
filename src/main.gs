var doc = DocumentApp.getActiveDocument();
var body = doc.getBody();
var style = {};
style[DocumentApp.Attribute.SPACING_BEFORE] = 15;
style[DocumentApp.Attribute.SPACING_AFTER] = 15;

function clearAllContents() {
  body.clear();
}

function onOpen() {
  DocumentApp.getUi()
      .createMenu('Add-On Testing')
      .addItem('Show initial message', 'showInitializationAlert')
      .addItem('Insert Horizontal Line', 'insertHorizontalLine')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showInitializationAlert() {
  var ui = DocumentApp.getUi(); // Same variations.

  var result = ui.alert(
    'Clear all contents of this document?',
     'Selecting \'Yes\' is required to initialize this add-on for the first time.',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    clearAllContents();
    insertHorizontalLine();
    ui.alert('Contents cleared.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('You selected \'No\', the contents of this document will be preserved. \nCreate a new blank document, and then run this add-on.');
  }
}

function insertHorizontalLine() {
  body.insertHorizontalRule(0);
  var par = body.insertParagraph(0, '');
  par.setAttributes(style);
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
  .evaluate()
  .setTitle('To-Do List')

  DocumentApp.getUi().showSidebar(html);
}

function getAllDocumentText() {
  return DocumentApp.getActiveDocument().getBody().getParagraphs();
}

function getAllTextBetweenHorizontalRules() {
  var textBetweenHorizontalRules = [];

  pars = getAllDocumentText();
  pars.forEach(function(e) {
    var isHorizontalRule = e.findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
    if (!isHorizontalRule) {
      textBetweenHorizontalRules.push(e.getText());
    }
  })
  return textBetweenHorizontalRules;
}
