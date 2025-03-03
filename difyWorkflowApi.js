// DifyのワークフローにAPIで接続する
// 接続先：1つの広告を分析するWF
function difyWorkflowApi() {
    const headers = {
        'Authorization': "Bearer app-EV39iJ5fu9bjLTvpwhqvdm5S",  //api key発行してペースト
        'Content-Type': 'application/json'              //必須
    };
    
    // Extracting additional parameters from the sheet
    const adName = sheet.getRange("B5").getValue();
    const imageUrl = sheet.getRange("C5").getValue();
    const costNet = sheet.getRange("F5").getValue();
    const imp = sheet.getRange("G5").getValue();
    const cpm = sheet.getRange("H5").getValue();
    const click = sheet.getRange("I5").getValue();
    const ctr = sheet.getRange("J5").getValue();
    const cpc = sheet.getRange("K5").getValue();
    const cv = sheet.getRange("L5").getValue();
    const cvr = sheet.getRange("M5").getValue();
    const cpa = sheet.getRange("N5").getValue();

    // prepare payload
    const inputs = {
            'ad_name': adName,
            'image_url': imageUrl,
            'spend': costNet,
            'impression': imp,
            'click': click,
            'conversion': cv
        };

    const data = {
        "user": "gas-difyWorkflowApi",                            // 任意の文字列で可能（監視で表示される）
        "response_mode": "blocking",                    // streaming or blocking 
        'inputs': inputs
    };

    const options = {
        "method": "post",
        "payload": JSON.stringify(data),
        "headers": headers,
        "muteHttpExceptions": true                      // エラーを平文で返してもらう
    };

    const requestUrl = "https://api.dify.ai/v1/workflows/run";
    const response = UrlFetchApp.fetch(requestUrl, options);
    const responseText = response.getContentText()

    // 帰ってきたレスポンスを表示
    Logger.log(responseText);                         // レスポンス内容をログに出力

    // StatusCodeによって処理分岐
    if (response.getResponseCode() === 200) {
        const responseJson = JSON.parse(responseText);
        Logger.log(responseJson.data.outputs.PR)
    } else {
        Logger.log("Error"); // エラー発生時のログ出力
    }
}
