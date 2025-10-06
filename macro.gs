/** macro.gs — Menu & sidebar launchers (removed "Part & Price Info") */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SMT')
    .addItem('Open Quote Form', 'showQuoteForm')
    .addItem('Generate Quote', 'generateQuoteFromActiveForm') // convenience
    .addToUi();
}

function showQuoteForm() {
  const html = HtmlService.createHtmlOutputFromFile('QuoteForm')
    .setTitle('Somerset Mobile Towbars — Quote');
  SpreadsheetApp.getUi().showSidebar(html);
}

// convenience wrapper to call same backend as the form
function generateQuoteFromActiveForm() {
  // read last selections from a hidden place if you store them, or just call the generator which reads Operations sheet.
  processQuote(); // Keep parity with QuoteOutput.gs entry point
}
