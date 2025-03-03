// difyApiCall.js
// This file is used to make API calls to the Dify API

/**
 * Calls the Dify API to process an add info and return the analysis and the insight.
 * The request data is an add info in the "CRTレポート" sheet.
 **/
function processDifyAgentMessage() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRTレポート");
    var range = sheet.getRange("B5:N5");
    var userMessage = 'この広告を分析してください。';
    var conversationId = Utilities.getUuid(); // Generate a valid UUID
    var userId = "user-gas-dev";

    // Extracting additional parameters from the sheet
    var imageUrl = sheet.getRange("C5").getValue();
    var costGross = sheet.getRange("E5").getValue();
    var costNet = sheet.getRange("F5").getValue();
    var imp = sheet.getRange("G5").getValue();
    var cpm = sheet.getRange("H5").getValue();
    var click = sheet.getRange("I5").getValue();
    var ctr = sheet.getRange("J5").getValue();
    var cpc = sheet.getRange("K5").getValue();
    var cv = sheet.getRange("L5").getValue();
    var cvr = sheet.getRange("M5").getValue();
    var cpa = sheet.getRange("N5").getValue();

    var options = {
        'method': 'post',
        'headers': {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + 'app-ZHfWqlTBDM1pbYvFc44gO3Us'
        },
        'payload': JSON.stringify({
            'inputs': {
                'image_url': imageUrl,
                'cost_gross': costGross,
                'cost_net': costNet,
                'imp': imp,
                'cpm': cpm,
                'click': click,
                'ctr': ctr,
                'cpc': cpc,
                'cv': cv,
                'cvr': cvr,
                'cpa': cpa
            },
            'query': userMessage,
            'response_mode': 'streaming',
            'user': userId,
            'conversation_id': conversationId
        }),
        'muteHttpExceptions': true
    };

    try {
        const DIFY_API_BASE_URL = 'https://api.dify.ai/v1';
        var response = UrlFetchApp.fetch(DIFY_API_BASE_URL + "/chat-messages", options);
        var result = "";
        var lines = response.getContentText().split("\n");

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (line.startsWith("data: ")) {
                var data = JSON.parse(line.substring(6));
                if (data.event === "agent_message" || data.event === "message") {
                    result += data.answer;
                } else if (data.event === "message_end") {
                    Logger.log("conversation: " + data.conversation_id);
                    return result;
                } else if (data.event === "error") {
                    return result + data.event + String(data);
                }
            }
        }
        var json = JSON.parse(response);
        if (json.code === "not_found") {
            Logger.log(json.message);
            return json.message;
        }
    } catch (e) {
        Logger.log(e.toString());
    }
}