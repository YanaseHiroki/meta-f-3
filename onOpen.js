// ファイルを開いたときにApps Scriptのリンクを表示(開発用)
function showScriptDialog() {
var scriptId = ScriptApp.getScriptId();
var url = "https://script.google.com/d/" + scriptId + "/edit";
var html = '<button onclick="window.open(\'' + url + '\', \'_blank\'); google.script.host.close();"><h1>Apps Script 開く</h1></button>';
var userInterface = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(250);
SpreadsheetApp.getUi().showModalDialog(userInterface, "Apps Script");
}

// スプレッドシートを開いたときに実行される関数
function onOpenProcess() {
    showSidebar();                  // サイドバーを表示
    addOriginalMenu();              // カスタムメニューを表示
    showScriptDialog();             // Apps Scriptのリンクを表示【開発用】
}

// サイドバーを表示(sideBar.htmlを読み込む)
function showSidebar() {
    Logger.log('showSidebar() start');

    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('データ取得・レポート作成');

    SpreadsheetApp.getUi().showSidebar(html);

    Logger.log('showSidebar() end');
}

// カスタムメニュー「初期設定」を表示
function addOriginalMenu() {
    Logger.log('addOriginalMenu() start');

    var ui = SpreadsheetApp.getUi();

    ui.createMenu('初期設定')
        .addItem('Metaトークン登録', 'registerMetaLongToken')
        .addToUi();

        // .addItem('', '') // メニュー追加
        // .addSeparator() // セパレーター追加

    Logger.log('addOriginalMenu() end');
}
