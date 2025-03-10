
// Metaアプリの長期トークンを取得してプロパティに登録する関数
function registerMetaLongToken() {
    //Logger.log('registerMetaLongToken() start');

    // クライアントIDを入力してもらう
    const appId = promptUserInput("META_APP_ID_PROMPT");
    if (!appId) return;

    // クライアントシークレットを入力してもらう
    const appSecret = promptUserInput("META_APP_SECRET_PROMPT");
    if (!appSecret) return;

    // 短期アクセストークンを入力してもらう
    const accessToken = promptUserInput("META_ACCESS_TOKEN_PROMPT");
    if (!accessToken) return;

    // 長期アクセストークンを取得
    const longAccessToken = getLongAccessToken(appId, appSecret, accessToken);
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty("META_ACCESS_TOKEN", longAccessToken);

    // メッセージを表示
    Browser.msgBox("Metaアプリの長期トークンを登録しました。");

    //Logger.log('registerMetaLongToken() end');
}

// ユーザ入力を取得する関数
// 引数：プロンプトのプロパティキー
function promptUserInput(promptKey) {

    const properties = PropertiesService.getScriptProperties();
    const prompt = properties.getProperty(promptKey);
    Logger.log(`プロンプト: ${prompt}`);

    const input = Browser.inputBox(prompt, Browser.Buttons.OK_CANCEL);
    Logger.log(`ユーザ入力: ${input}`);

    if (input == 'cancel') {
        return null;
    }
    return input;
}

// 長期アクセストークンを取得する関数
// 引数：アプリID, シークレット, 短期アクセストークン
function getLongAccessToken(clientId, clientSecret, shortLivedToken) {
    Logger.log('getLongAccessToken(clientId, clientSecret, shortLivedToken) start');

    // URLとクエリパラメータの構築
    const apiVersion = PropertiesService.getScriptProperties().getProperty("META_API_VERSION");
    const baseUrl = `https://graph.facebook.com/${apiVersion}/oauth/access_token`;
    const queryParams = {
        grant_type: "fb_exchange_token",
        client_id: clientId,
        client_secret: clientSecret,
        fb_exchange_token: shortLivedToken,
    };

    // クエリパラメータをURL形式にエンコード
    const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join("&");
    const urlWithParams = `${baseUrl}?${queryString}`;

    // APIリクエスト用オプションの設定
    const options = {
        method: "get",
        muteHttpExceptions: true, // エラー発生時に例外をスローしない
    };

    try {
        // APIリクエストを送信
        const response = UrlFetchApp.fetch(urlWithParams, options);
        const responseCode = response.getResponseCode();

        if (responseCode === 200) {
            // 正常にデータを取得
            const json = JSON.parse(response.getContentText());
            const newToken = json.access_token;

            if (!newToken) {
                throw new Error(`トークンがレスポンスに含まれていません: ${JSON.stringify(json)}`);
            }

            // 新しいトークンを保存
            saveAccessToken(newToken);
            Logger.log("新しいトークンを取得しました: " + newToken);
            return newToken;
        } else {
            // エラーハンドリング
            const errorDetails = response.getContentText();
            const errorMessage = `トークン更新に失敗しました。ステータスコード: ${responseCode}, エラー詳細: ${errorDetails}`;
            Logger.log(errorMessage);
            throw new Error(errorMessage);
        }
    } catch (error) {
        // 例外処理
        const errorMessage = `トークン更新中にエラーが発生しました。エラー内容: ${error.message}`;
        Logger.log(errorMessage);
        throw error; // 必要に応じて再スロー
    }
}

// 【定期実行用】長期アクセストークンを更新する関数
function refreshAccessToken() {
    Logger.log('refreshAccessToken() start');

    // 更新前のトークンを取得
    const properties = PropertiesService.getScriptProperties();
    const currentToken = properties.getProperty("META_ACCESS_TOKEN");

    // トークンが取得できない場合はエラー
    if (!currentToken) {
        const errorMessage = "更新前のトークンが取得できませんでした。";
        Logger.log(errorMessage);
        throw new Error(errorMessage);
    }

    // トークンを更新
    const appId = properties.getProperty("META_APP_ID");
    const appSecret = properties.getProperty("META_APP_SECRET");
    const newToken = getLongAccessToken(appId, appSecret, currentToken);

    // 更新後のトークンが取得できない場合はエラー
    if (!newToken) {
        const errorMessage = "更新後のトークンが取得できませんでした。";
        Logger.log(errorMessage);
        throw new Error(errorMessage);
    }

    // トークンを保存
    properties.setProperty("META_ACCESS_TOKEN", newToken);

    Logger.log('refreshAccessToken() end');
}