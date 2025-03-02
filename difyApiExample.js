/**
 * Dify APIを使用してユーザーメッセージを処理し、エージェントの応答を返します。
 * この関数はストリーミングレスポンスを処理し、行ごとに解析します。
 * 
 * @param {string} userMessage - 処理するユーザーからのメッセージ。
 * @param {string} conversationId - 会話のID。新しい会話の場合は空文字列を指定。
 * @param {string} userId - メッセージを送信するユーザーのID。
 * @returns {string|Object} エージェントの応答を文字列で返します。エラーが発生した場合はエラーメッセージを含むオブジェクトを返します。
 * 
 * @throws APIリクエストまたはレスポンス処理中に発生したエラーはログに記録されます。
 * 
 * @example
 * var response = processDifyAgentMessage("こんにちは、お元気ですか？", "conv123", "user456");
 * Logger.log(response);
 * 
 * @note 対話のスタート時にはconversationIdを空の文字列を渡せばいいです。
 *       対話を継続するときには、conversationIdを指定すればよいです。
 */
function processDifyAgentMessage(userMessage, conversationId, userId) {
    var options = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + DIFY_AGENT_API_KEY
      },
      'payload': JSON.stringify({
        'inputs': {},
        'query': userMessage,
        'response_mode': 'streaming',
        'user': userId,
        'conversation_id': conversationId
      }),
      'muteHttpExceptions': true
    };
  
    try {
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
        return json.message;
      }
    } catch (e) {
      Logger.log(e.toString());
      return { error: 'リクエストの処理中にエラーが発生しました。' };
    }
  }
  
  function testAgentConversation() {
    var message = processDifyAgentMessage("自己紹介して", "", "your_user_name");
    // var message = processDifyAgentMessage("カフェに行こう", "your_conversation_id", "your_user_name");
    Logger.log(message);
  }