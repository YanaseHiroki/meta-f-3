// 導入時に手動実行する関数(トリガー追加は手動実行によってしか行えないため)

function aSettingAtDay1() {
    Logger.log('初期設定を開始します');
    SpreadsheetApp.getActiveSpreadsheet().toast("必要な設定値を入力してください。", "初期設定", 10);
     
    initializeScriptProperties();   // スクリプトプロパティを初期化
    addTrigger();                   // トリガーを追加
    onOpenProcess();                // カスタムメニューを表示
    refreshAccessToken();           // 長期トークンを取得

    SpreadsheetApp.getActiveSpreadsheet().toast("初期設定を終了します。", "初期設定", 10);
    Logger.log('初期設定を終了します');
}

// スクリプトプロパティの初期値を設定する関数
function initializeScriptProperties() {
    
    // スクリプトプロパティサービスを取得
    const properties = PropertiesService.getScriptProperties();
    
    // 各種プロパティの値を設定する(初回は追加、2回目以降は更新)
        properties.setProperties({

            // MetaアプリのAPIのバージョン
            'META_API_VERSION': 'v22.0',
            'META_API_VERSION_PROMPT': 'MetaアプリのAPIのバージョンを入力してください。',
            
            // DifyのアプリケーションID
            'DIFY_APP_ID': 'app-6gipSAA9nO3Uy8gBolsqRwGb', // 1~5個の広告を分析するCF (Claude)
            'DIFY_APP_ID_PROMPT': 'DifyのアプリケーションIDを入力してください。',

            // 広告アカウントID
            'META_AD_ACCOUNT_ID_PROMPT': 'Metaの広告アカウントIDを入力してください。',
            
            // MetaアプリのアプリID
            'META_APP_ID_PROMPT': 'MetaアプリのアプリIDを入力してください。',

            // Metaアプリのapp secret
            'META_APP_SECRET_PROMPT': 'Metaアプリのapp secretを入力してください。',

            // Metaアプリのアクセストークン
            'META_ACCESS_TOKEN_PROMPT': 'Metaアプリのアクセストークンを入力してください(ads_read権限が必要です)。',

            // コンバージョンのタイプ
            'META_CONVERSION_TYPE_PROMPT': '広告のコンバージョンのタイプを入力してください。',
        });
}

//   // すべてのKey-Valueを取得
//   properties.getProperties("");
//   // 取得
//   properties.getProperty("");
//   // 更新
//   properties.setProperty("","");


// トリガーを追加
function addTrigger() {

    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    if(triggers) {
        for (let i = 0; i < triggers.length; i++) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }

    // スプレッドシートを開いたときにonOpenProcess()を実行するトリガーを追加
    ScriptApp.newTrigger('onOpenProcess')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onOpen()
        .create();
        
    Logger.log('起動時にサイドバーとメニューを表示する処理のトリガーを追加しました');

    // 1か月ごとにrefreshAccessToken()を実行するトリガーを追加
    ScriptApp.newTrigger('refreshAccessToken')
        .timeBased()
        .onMonthDay(1)
        .atHour(5)
        .create();

    Logger.log('1か月ごとにrefreshAccessToken()を実行するトリガーを追加しました');

    // 毎日午前5時にfacebook_getAdSetsForYesterday()を実行するトリガーを追加
    ScriptApp.newTrigger('facebook_getAdSetsForYesterday')
        .timeBased()
        .everyDays(1)
        .atHour(5)
        .create();

    Logger.log('長期トークンを自動的に更新するためのトリガーを追加しました');
}
