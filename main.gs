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
      .addToUi();
}

function showInitializationAlert() {
  var ui = DocumentApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm, are you sure you want to clear all contents of this document?',
     'Selecting \'Yes\' is required to initialize this add-on for the first time.',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    clearAllContents();
    ui.alert('Contents cleared.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('You selected \'No\', the contents of this document will be preserved. Create a new blank document, and then run this add-on.');
  }
}


function insertHorizontalLine() {
  body.insertHorizontalRule(0);
  var par = body.insertParagraph(0, '');
  par.setHeading(DocumentApp.ParagraphHeading.HEADING4);
  par.setAttributes(style);
}
