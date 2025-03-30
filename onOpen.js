// ファイルを開いたときにApps Scriptのリンクを表示(開発用)
function showScriptDialog() {
var scriptId = ScriptApp.getScriptId();
var url = "https://script.google.com/d/" + scriptId + "/edit";
var html = '<br><br><button onclick="window.open(\'' + url + '\', \'_blank\'); google.script.host.close();"><h1>Apps Script を開く</h1></button>';
var userInterface = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(200);
SpreadsheetApp.getUi().showModalDialog(userInterface, "Apps Script");
}

// スプレッドシートを開いたときに実行される関数
function onOpenProcess() {
    showSidebar();                  // サイドバーを表示
    addOriginalMenu();              // カスタムメニューを表示
    showScriptDialog();             // Apps Scriptのリンクを表示【開発用】
    checkScriptProperties();        // スクリプトプロパティのチェック
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

    ui.createMenu('スクリプト実行')
        .addItem('サイドメニュー表示', 'showSidebar')
        .addItem('CRTレポート再作成', 'makeCreativeReport')
        .addSeparator()
        .addItem('APIトークン更新', 'refreshAccessToken')
        .addToUi();

        // .addItem('', '') // メニュー追加
        // .addSeparator() // セパレーター追加

    Logger.log('addOriginalMenu() end');
}


// スクリプトプロパティが設定されているかどうかをチェックする関数
function checkScriptProperties() {

    // 必要なスクリプトプロパティのキーの配列
    var requiredKeys = 
        [
            'META_API_VERSION',
            'DIFY_APP_ID',
            'META_AD_ACCOUNT_ID',
            'META_APP_ID',
            'META_APP_SECRET',
            'META_ACCESS_TOKEN'
        ];

    var scriptProperties = PropertiesService.getScriptProperties();
    var allKeysSet = true;

    // 必要なスクリプトプロパティが設定されているかどうかをチェック
    requiredKeys.forEach(function(key) {
        var value = scriptProperties.getProperty(key);

        // スクリプトプロパティが設定されていない場合、ユーザに入力してもらう
        if (!value) {
            value = promptUserInput(key + '_PROMPT');
            if (value) {
                scriptProperties.setProperty(key, value);
            } else {
                Logger.log('Missing script property: ' + key);
                allKeysSet = false;
            }
        }
    });

    // すべてのスクリプトプロパティが設定されているかどうかをログに出力
    if (allKeysSet) {
        Logger.log('All required script properties are set.');
    } else {
        Logger.log('Some required script properties are missing.');
    }
}
