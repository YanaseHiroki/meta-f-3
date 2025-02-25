function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('データ取得・レポート作成');
  SpreadsheetApp.getUi().showSidebar(html);
}
