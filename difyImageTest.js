// DifyのワークフローにAPIで接続する
// 接続先：1つの広告を分析するWF
function difyImageTest() {
    const headers = {
        'Authorization': "Bearer app-EV39iJ5fu9bjLTvpwhqvdm5S",  //api key発行してペースト
        'Content-Type': 'application/json'              //必須
    };
    
    // Extracting additional parameters from the sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRTレポート");
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

    // inputs for payload
    const inputs = {
            'ad_name': adName,
            'image_url': imageUrl,
            'spend': costNet,
            'impression': imp,
            'cpm': cpm,
            'click': click,
            'ctr': ctr,
            'cpc': cpc,
            'conversion': cv,
            'cvr': cvr,
            'cpa': cpa
        };

    const payload = JSON.stringify({
        "user": "gas-difyImageTest",
        "response_mode": "blocking",
    });

    const options = {
        "method": "post",
        "payload": payload,
        "headers": headers,
        "muteHttpExceptions": true
    };

    const requestUrl = "https://api.dify.ai/v1/workflows/run";
    const response = UrlFetchApp.fetch(requestUrl, options);
    const responseText = response.getContentText();
    const responseJson = JSON.parse(responseText);

    // 帰ってきたレスポンスを表示
    Logger.log("responseText: " + responseText);
    
    // StatusCodeによって処理分岐
    if (response.getResponseCode() === 200) {
        Logger.log('responseJson.data.outputs.text: ' + responseJson.data.outputs.text); // レスポンスのdata部分をログ出力
    } else {
        Logger.log("Error"); // エラー発生時のログ出力
    }
}
