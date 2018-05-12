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

  showSidebar();
}

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

function insertHorizontalLine() {
  body.insertHorizontalRule(0);
  var par = body.insertParagraph(0, '');
  par.setAttributes(style);
  showSidebar();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
  .evaluate()
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .setTitle('To-Do List')

  DocumentApp.getUi().showSidebar(html);
}

function getAllDocumentText() {  
  return DocumentApp.getActiveDocument().getBody().getParagraphs();
}

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

function deleteText(textData) {
  body.insertParagraph(0, textData);
}
