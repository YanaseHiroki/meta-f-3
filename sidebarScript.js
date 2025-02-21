function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Facebook広告取得');
  SpreadsheetApp.getUi().showSidebar(html);
}
