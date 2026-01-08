function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('mBOM')
    .addItem('Open App (Window)', 'ui_openAppWindow')
    .addSeparator()
    .addItem('Refresh Agile Index', 'agile_refreshIndex')
    .addItem('Refresh Files Index (Drive)', 'files_refreshIndexFromDrive')
    .addSeparator()
    .addItem('Open Copy Monitoring (Jobs)', 'ui_openCopyMonitoring')
    .addToUi();
}

function ui_openAppWindow() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('VERDON mBOM App')
    .setWidth(1320)
    .setHeight(880);

  SpreadsheetApp.getUi().showModelessDialog(html, 'VERDON mBOM App');
}

function ui_openCopyMonitoring() {
  // Opens the main app; monitoring is a dedicated tab in the UI.
  ui_openAppWindow();
}

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.__WEB_MODE__ = true;
  tpl.__QUERY__ = (e && e.parameter) ? e.parameter : {};
  return tpl.evaluate().setTitle('VERDON mBOM App');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
