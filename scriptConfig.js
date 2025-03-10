// GASの設定に関するファイル

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
            'META_APP_ID_PROMPT': 'Metaアプリの長期トークンを登録します。/n/nアプリIDを入力してください。',
            'META_APP_SECRET_PROMPT': '次にapp secretを入力してください。',
            'META_ACCESS_TOKEN_PROMPT': '次にads_readを持つアクセストークンを入力してください。',
        });
}

//   // すべてのKey-Valueを取得
//   properties.getProperties("");

//   // 取得
//   properties.getProperty("");

//   // 更新
//   properties.setProperty("","");

