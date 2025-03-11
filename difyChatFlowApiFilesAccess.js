// 　1つ目の広告情報を取得する(動作確認用)
function testDifyChatflowApiFilesAccess() {

    // Extracting additional parameters from the sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRTレポート");
    const data = sheet.getRange("B5:B9").getValues(); // B5:N9の範囲を取得（最大5行）
    difyChatflowApi(data);
}


// Difyの「1~5個の広告を分析するCF (access url)」にAPIで接続する
// DifyのワークフローにAPIで接続する
// 接続先：1つの広告を分析するWF
function difyChatflowApiFilesAccess(data, adSetName) {

    // スクリプトプロパティを取得
    const properties = PropertiesService.getScriptProperties();

    // ヘッダー情報
    const headers = {
        'Authorization': "Bearer " + properties.getProperty("DIFY_APP_ID"),
        'Content-Type': 'application/json'
    };

    let addArr = [];
    let imagesForPayload = [];

    data.forEach(row => {
        if (row[0]) { // B列に値がある場合
            addArr.push({
                'ad_name': row[0],
                'image_url': row[1],
                'spend': row[4],
                'impression': row[5],
                'cpm': row[6],
                'click': row[7],
                'ctr': row[8],
                'cpc': row[9],
                'conversion': row[10],
                'cvr': row[11],
                'cpa': row[12]
            });
            
            imagesForPayload.push({
                "type": "image",
                "transfer_method": "remote_url",
                "url": row[1]
            });
        }
    });

    const addsForPayloard = JSON.stringify(addArr);
    Logger.log('addsForPayloard: ' + addsForPayloard);

    const payload = JSON.stringify({
        "user": "gas-difyChatflowApi",
        'response_mode': 'blocking',
        'conversation_id': searchAdSetMapFromScriptProperties(adSetName),
        'inputs': {
            "adds": addsForPayloard,
        },
        'query': 'この広告を分析してください。',
        'files': imagesForPayload
    });

    const options = {
        "method": "post",
        "payload": payload,
        "headers": headers,
        "muteHttpExceptions": true
    };

    const requestUrl = "https://api.dify.ai/v1/chat-messages";
    const response = UrlFetchApp.fetch(requestUrl, options);
    const responseText = response.getContentText();

    // 帰ってきたレスポンスを表示
    Logger.log("responseText: " + responseText);
    
    // StatusCodeによって処理分岐
    if (response.getResponseCode() === 200) {
    const responseJson = JSON.parse(responseText);
        const answerJson = JSON.parse(responseJson.answer);
        const conversationId = responseJson.conversation_id;

        // スクリプトプロパティのMapにconversationIdと広告セット名を追加
        addAdSetMapToScriptProperties(conversationId, adSetName);

        Logger.log('会話ID: ' + conversationId);
        Logger.log('現状整理: ' + answerJson.current_status);
        Logger.log('今後の示唆: ' + answerJson.future_implications);
        Logger.log('画像情報: ' + answerJson.img_info);
        console.log('difyChatflowApi return answerJson: ' + JSON.stringify(answerJson));
        
        return answerJson;
    } else {
        // エラー処理

        Logger.log("difyChatflowApi Error Code: " + response.getResponseCode());
    }
}

// スクリプトプロパティのconversationAdSetMapに追加する関数
function addAdSetMapToScriptProperties(conversationId, adSetName) {
    const properties = PropertiesService.getScriptProperties();
    const conversationAdSetMap = JSON.parse(properties.getProperty("conversationAdSetMap"));
    conversationAdSetMap[adSetName] = conversationId;
    properties.setProperty("conversationAdSetMap", JSON.stringify(conversationAdSetMap));
    Logger.log('set conversationAdSetMap: ' + JSON.stringify(conversationAdSetMap));
}

// スクリプトプロパティのconversationAdSetMapを検索する関数
function searchAdSetMapFromScriptProperties(adSetName) {
    const properties = PropertiesService.getScriptProperties();

    // スクリプトプロパティのconversationAdSetMapが存在しない場合は空文字を返す
    const conversationAdSetString = properties.getProperty("conversationAdSetMap");
    if(!conversationAdSetString) {
        Logger.log('conversationAdSetMap is not found in script properties');
        return "";
    }

    // スクリプトプロパティのconversationAdSetMap
    const conversationAdSetMap = JSON.parse(conversationAdSetString);
    const conversationId = conversationAdSetMap[adSetName];
    if(!conversationId) {
        Logger.log('conversationId for adSetName is not found');
        return "";
    }

    Logger.log('searched conversationId: ' + conversationId);
    return conversationId;
}
