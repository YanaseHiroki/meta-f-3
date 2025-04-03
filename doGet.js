function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('【WLB様】Meta広告')
    .setWidth(300);
}