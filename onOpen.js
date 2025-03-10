// スプレッドシートを開いたときに実行される関数
function onOpenProcess() {
    showSidebar();                  // サイドバーを表示
    addOriginalMenu();              // カスタムメニューを表示
}

// サイドバーを表示(sideBar.htmlを読み込む)
function showSidebar() {
    Logger.log('showSidebar() start');

    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('データ取得・レポート作成');

    SpreadsheetApp.getUi().showSidebar(html);

    Logger.log('showSidebar() end');
}

// カスタムメニューを表示(「初期設定」)
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
