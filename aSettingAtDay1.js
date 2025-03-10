// 導入時に手動実行する関数

function aSettingAtDay1() {
    Logger.log('初期設定を開始します');
    
    initializeScriptProperties();   // スクリプトプロパティを初期化
    addTrigger();                   // トリガーを追加
    onOpen();                       // サイドバーを表示, カスタムメニューを表示
    registerMetaLongToken();        // メタアプリのトークンを登録

    Logger.log('初期設定を終了します');
}

// スクリプトプロパティを取得する関数
function initializeScriptProperties() {
    
    // スクリプトプロパティを取得
    const properties = PropertiesService.getScriptProperties();
    
    // 各種プロパティの値を設定する(初回は追加、2回目以降は更新)
        properties.setProperties({
            
            // DifyのアプリケーションID
            'DIFY_APP_ID': 'app-R2f9xzr4VeznYFxX4fagkQGl', // 1~5個の広告を分析するCF (access url)

            // MetaアプリのAPIのバージョン
            'META_API_VERSION': 'v22.0',
            
            // Metaアプリの長期トークンを取得してプロパティに登録するモーダルのプロンプト
            'META_APP_ID_PROMPT': 'Metaアプリの長期トークンを登録します。1/3 アプリIDを入力してください。',
            'META_APP_SECRET_PROMPT': '2/3 app secretを入力してください。',
            'META_ACCESS_TOKEN_PROMPT': '3/3 ads_readを持つアクセストークンを入力してください。',
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
        
    // 1か月ごとにrefreshAccessToken()を実行するトリガーを追加
    ScriptApp.newTrigger('refreshAccessToken')
    .timeBased()
    .onMonthDay(1)
    .atHour(5)
    .create();

    Logger.log('長期トークンを自動的に更新するトリガーを追加しました');
}
