// スプレッドシートを開いたときに実行される関数
function onOpen () {
    
    initializeScriptProperties();   // スクリプトプロパティを初期化
    showSidebar();                  // サイドバーを表示
    addOriginalMenu();              // カスタムメニューを追加
}

// サイドバーを表示(sideBar.htmlを読み込む)
function showSidebar() {

    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('データ取得・レポート作成');

    SpreadsheetApp.getUi().showSidebar(html);
}

// カスタムメニューを追加
function addOriginalMenu() {

    var ui = SpreadsheetApp.getUi();

    ui.createMenu('初期設定')
        .addItem('Metaトークン登録', 'registerMetaLongToken')
        .addToUi();

        // .addItem('', '') // メニュー追加
        // .addSeparator() // セパレーター追加
}

// Metaアプリの長期トークンを取得してプロパティに登録する関数
function registerMetaLongToken() {

    // クライアントIDを入力してもらう
    const appId = promptUserInput("META_APP_ID_PROMPT");
    if (appId) return;

    // クライアントシークレットを入力してもらう
    const appSecret = promptUserInput("META_APP_SECRET_PROMPT");
    if (appSecret) return;
    
    // 短期アクセストークンを入力してもらう
    const accessToken = promptUserInput("META_ACCESS_TOKEN_PROMPT");
    if (accessToken) return;

    // 長期アクセストークンを取得
    getLongAccessToken(appId, appSecret, accessToken);

}

// ユーザ入力を取得する関数
// 引数：プロンプトのプロパティキー
function promptUserInput(promptKey) {
    const prompt = properties.getProperty(promptKey);
    Logger.log(`プロンプト: ${prompt}`);

    const input = Browser.inputBox(prompt, Browser.Buttons.OK_CANCEL);
    Logger.log(`ユーザ入力: ${clientId}`);

    if (input == 'cancel') {
        return null;
    }
    return input;
}

// 長期アクセストークンを取得する関数
// 引数：アプリID, シークレット, 短期アクセストークン
getLongAccessToken(clientId, clientSecret, shortLivedToken) {
