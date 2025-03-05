// Difyの「1~5個の広告を分析するCF」にAPIで接続する
// DifyのワークフローにAPIで接続する
// 接続先：1つの広告を分析するWF
function difyChatflowApi() {
    const headers = {
        'Authorization': "Bearer app-45HKkeZwPDkHlkKez0J8fsqD",  //api key　(1~5個の広告を分析するCF-file試す)
        'Content-Type': 'application/json'              //必須
    };
    
    // Extracting additional parameters from the sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRTレポート");
    const data = sheet.getRange("B5:N9").getValues(); // B5:N9の範囲を取得（最大5行）

    let addArr = [];
    let imagesForPayload = [];

    data.forEach(row => {
        if (row[0]) { // B列に値がある場合
            addArr.push({
                'ad_name': row[0],
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
    Logger.log('adds: ' + addsForPayloard);

    const payload = JSON.stringify({
        "user": "gas-difyChatflowApi",
        'response_mode': 'blocking',
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
    const responseJson = JSON.parse(responseText);

    // 帰ってきたレスポンスを表示
    Logger.log("responseText: " + responseText);
    
    // StatusCodeによって処理分岐
    if (response.getResponseCode() === 200) {
        const answerJson = JSON.parse(responseJson.answer);
        Logger.log('現状整理: ' + answerJson.current_status);
        Logger.log('今後の示唆: ' + answerJson.future_implications);
    } else {
        Logger.log("Error");
    }
}
