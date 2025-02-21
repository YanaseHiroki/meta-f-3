   // キャンペーン情報を取得する関数
function facebook_getCampaign() {
  console.log("facebook_getCampaign()");

  var sheetName="キャンペーン";
  var endpoint = "campaigns";
  console.log(sheetName + "情報 取得開始");

  // キャンペーンで使用可能なフィールドについては、以下の公式ドキュメントを参考にしてください：
  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign-group/insights?locale=ja_JP
  // ただし、記載されているフィールドの中には一部使用できないものがある点にご注意ください。
  // 利用可能なフィールドについては、実際にAPIリクエストを行って確認することをおすすめします。
  var fields = "date_start,date_stop,account_id,account_name,campaign_id,campaign_name,impressions,inline_link_clicks,conversions,spend";
  
  refreshAccessToken()
  facebook_writeFacebookAdsDataToSheet(sheetName,endpoint, fields);
}

// 広告セット情報を取得する関数
function facebook_getAdSets() {
  console.log("facebook_getAdSets()");

  var sheetName="広告セット";
  var endpoint = "adsets";
  console.log(sheetName + "情報 取得開始");

  // 広告セットで使用可能なフィールドについては、以下の公式ドキュメントを参考にしてください：
  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign/insights?locale=ja_JP  
  // ただし、記載されているフィールドの中には一部使用できないものがある点にご注意ください。
  // 利用可能なフィールドについては、実際にAPIリクエストを行って確認することをおすすめします。  
  var fields = "account_id,adset_id,adset_name,date_start,date_stop,impressions,spend,cpm,clicks,ctr,cpc,conversions,conversion_values";

  refreshAccessToken()
  facebook_writeFacebookAdsDataToSheet(sheetName,endpoint, fields);
  
  console.log(sheetName + "情報 取得完了");
}

// 広告情報を取得する関数
function facebook_getAds(daySince, dayUntil) {
  console.log(`facebook_getAds(${daySince},${dayUntil})`);

  var sheetName="広告";
  var endpoint = "ads";
  console.log(sheetName + "情報 取得開始");

  // 広告で使用可能なフィールドについては、以下の公式ドキュメントを参考にしてください：
  // https://developers.facebook.com/docs/marketing-api/reference/adgroup/insights/
  // ただし、記載されているフィールドの中には一部使用できないものがある点にご注意ください。
  // 利用可能なフィールドについては、実際にAPIリクエストを行って確認することをおすすめします。
  var fields = "ad_id,campaign_name,adset_name,ad_name,spend,impressions,cpm,clicks,ctr,cpc,actions,date_start,date_stop";
  // account_id,conversion_values,image_url,thumbnail_url,conversions,conversion_values,action_values,


  refreshAccessToken();

  facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil);

  console.log(sheetName + "情報 取得完了");
  
}

// スプレッドシートからアクセストークンを取得する関数
function facebook_getAccessToken() {
  // console.log("facebook_getAccessToken()");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('トークン管理');
  if (sheet) {
    return sheet.getRange(1, 1).getValue(); // 1行1列目からトークンを取得
  } else {
    Logger.log('シート "トークン管理" が見つかりません。');
    return null; // シートが見つからなかった場合、nullを返す
  }
}

function saveAccessToken(token) {
  // console.log(`saveAccessToken(${token})`);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('トークン管理');
  if (sheet) {
    sheet.getRange(1, 1).setValue(token); // 1行1列目にトークンを保存
  } else {
    Logger.log('シート "トークン管理" が見つかりません。');
  }
}

function loadAccessToken() {
  // console.log("loadAccessToken()");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('トークン管理');
  if (sheet) {
    return sheet.getRange(1, 1).getValue(); // 1行1列目からトークンを取得
  } else {
    Logger.log('シート "トークン管理" が見つかりません。');
    return null; // シートが見つからなかった場合、nullを返す
  }
}

// アクセストークンを生成する関数
function refreshAccessToken() {
  console.log("refreshAccessToken()");

  // 必要な定数を取得
  const clientId = '984540190212296'; // アプリケーションID
  const clientSecret = '70ece718923048ca46698fb66fc1e90f'; // アプリシークレットキー
  const shortLivedToken = loadAccessToken(); // 現在の短期間トークンを取得

  // URLとクエリパラメータの構築
  const apiVersion = 'v22.0'; // 使用するAPIバージョン
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

// 任意の日数前の日付を計算する共通関数
function facebook_getDateNDaysAgo(daysAgo) {
  var today = new Date();
  today.setDate(today.getDate() - daysAgo); // 指定の日数前に設定
  return Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Facebook Ads APIからデータを取得する汎用関数
function facebook_getData(endpoint) {
  // Facebook広告アカウントIDとアクセストークンを設定
  var accessToken = facebook_getAccessToken(); // スプレッドシートからアクセストークンを取得
  var adAccountId = '1362620894448891'; // Facebook広告アカウントID
  var apiVersion = 'v22.0'; // 使用するAPIのバージョン

  // APIのURLを構築
  var apiUrl = `https://graph.facebook.com/${apiVersion}/act_${adAccountId}/${endpoint}`;
  var queryParams = {
    fields: "id,name", // リクエストで取得するフィールド
    access_token: accessToken // アクセストークン
  };

  // クエリパラメータをURLエンコードして追加
  var queryString = "";
  for (var key in queryParams) {
    if (queryParams.hasOwnProperty(key)) {
      var encodedKey = encodeURIComponent(key);
      var encodedValue = encodeURIComponent(queryParams[key]);
      queryString += `${encodedKey}=${encodedValue}&`;
    }
  }

  // 最後の&を削除
  queryString = queryString.slice(0, -1);

  // 完全なURLを構築
  var urlWithParams = apiUrl + "?" + queryString;

  // APIリクエスト用オプションの設定
  var options = {
    method: "get",
    muteHttpExceptions: true // エラーが発生した場合でも例外をスローしない
  };

  try {
    // APIリクエストを送信
    var response = UrlFetchApp.fetch(urlWithParams, options);
    var responseCode = response.getResponseCode(); // レスポンスコードを取得

    if (responseCode === 200) {
      // 正常にデータを取得
      var jsonData = JSON.parse(response.getContentText());
      return jsonData.data;
    } else {
      // エラーハンドリング（ステータスコードが200以外の場合）
      var errorDetails = response.getContentText();
      var errorMessage = `Meta広告データ取得に失敗しました。エンドポイント: ${endpoint}, ステータスコード: ${responseCode}, エラー詳細: ${errorDetails}`;
      Logger.log(errorMessage);
      return null;
    }
  } catch (error) {
    // その他のエラーハンドリング
    var errorMessage = `Meta広告データ取得中にエラーが発生しました。エンドポイント: ${endpoint}, エラー内容: ${error.message}`;
    Logger.log(errorMessage);
    return null;
  }
}


// キャンペーンごとに広告データを取得する関数
function getFacebookAdsDataForCampaign(campaignId,fields,argDaySince,argDayUince) {
  console.log(`getFacebookAdsDataForCampaign(${campaignId},${fields},${argDaySince},${argDayUince})`);

  // 必要な変数と設定を準備
  var apiVersion = 'v22.0';  // APIのバージョン
  var accessToken = facebook_getAccessToken(); // スプレッドシートからアクセストークンを取得

  // APIリクエスト用のURLを構築
  var apiUrl = `https://graph.facebook.com/${apiVersion}/${campaignId}/insights`;
  var queryParams = {
    level: "ad",
    fields: fields,
    // "time_range[since]": argDaySince,
    // "time_range[until]": argDayUince,
  date_preset: "last_90d",
    access_token: accessToken
  };

  // クエリパラメータをURLに追加
  var queryString = "";
  for (var key in queryParams) {
    if (queryParams.hasOwnProperty(key)) {
      // パラメータをエンコードして、&でつなげる
      var encodedKey = encodeURIComponent(key);
      var encodedValue = encodeURIComponent(queryParams[key]);
      queryString += `${encodedKey}=${encodedValue}&`;
    }
  }

  // 最後の&を削除
  queryString = queryString.slice(0, -1);

  // 完全なURLを構築
  var urlWithParams = apiUrl + "?" + queryString;
  console.log(`urlWithParams in getFacebookAdsDataForCampaign: ${urlWithParams}`);

  // APIリクエスト用オプションの設定
  var options = {
    method: "get",
    muteHttpExceptions: true // エラーが発生した場合でも例外をスローしない
  };

  try {
    // APIリクエストを送信
    var response = UrlFetchApp.fetch(urlWithParams, options);
    var responseCode = response.getResponseCode(); // レスポンスコードを取得

    if (responseCode === 200) {
      // 正常にデータを取得
      var jsonData = JSON.parse(response.getContentText());

      console.log(`getFacebookAdsDataForCampaign return: ${JSON.stringify(response, null, 2)}`);
      return jsonData.data;
    } else {
      // エラーハンドリング（ステータスコードが200以外の場合）
      var errorDetails = response.getContentText();
      var errorMessage = `Meta広告キャンペーンデータ取得に失敗しました。ステータスコード: ${responseCode}, エラー詳細: ${errorDetails}`;
      Logger.log(errorMessage);
      return null;
    }
  } catch (error) {
    // その他のエラーハンドリング
    var errorMessage = `Meta広告キャンペーンデータ取得中にエラーが発生しました: ${error.message}`;
    Logger.log(errorMessage);
    return null;
  }
}

// 取得した広告データをスプレッドシートにキャンペーンごとに書き込む関数
function facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil) {
  console.log(`facebook_writeFacebookAdsDataToSheet(${sheetName}, ${endpoint}, ${fields}, ${daySince}, ${dayUntil})`);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシートを取得

  // シートが存在する場合は、そのシートを削除
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    // シートのすべてをクリアする
    sheet.clear();
  } else {
    // シートが存在しない場合は新しいシートを作成
    sheet = spreadsheet.insertSheet(sheetName);
  }

  var lastRow = sheet.getLastRow();  // 最終行を取得

  // もしシートにデータがない場合は、最初の行をヘッダーとして設定

  // ------------------------------------------------------------
  var daysBefore = 40; // ◆◆◆◆◆◆◆◆◆◆何日前から前を見るか◆◆◆◆◆◆◆◆◆◆
  var yesterday = facebook_getDateNDaysAgo(daysBefore); // ◆◆◆◆◆◆◆◆◆◆何日前から◆◆◆◆◆◆◆◆◆◆
  var oneDaysAgo = facebook_getDateNDaysAgo(daysBefore + 120); // ◆◆◆◆◆◆◆◆◆◆何日前まで◆◆◆◆◆◆◆◆◆◆

  // 最長期間は37か月です。それ以前のものは取得できません。

  // 参考文献
  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign-group/ads?locale=ja_JP
  // ------------------------------------------------------------

  if (lastRow === 0) {
    // キャンペーンのデータ構造を取得
    var campaigns = facebook_getData(endpoint);
    
    if (!campaigns || campaigns.length === 0) {
      Logger.log("キャンペーンデータが取得できませんでした。");
      return;
    }

    var firstCampaignData = null;
    for (var i = 0; i < campaigns.length; i++) {
      firstCampaignData = getFacebookAdsDataForCampaign(campaigns[i].id, fields, oneDaysAgo, yesterday);
      if (firstCampaignData && firstCampaignData.length > 0) {
        break;
      }
    }
    if (!firstCampaignData || firstCampaignData.length === 0) {
      Logger.log("ヘッダーのサンプルデータが取得できませんでした。");
      return;
    }

    var sampleAd = firstCampaignData[0];

    if (!sampleAd || Object.keys(sampleAd).length === 0) {
      Logger.log("ヘッダーに使用するデータが取得できませんでした。");
      return;
    }

    var header = Object.keys(sampleAd);
    header = header.map(h => h === 'actions' ? 'conversion_purchase' : h);
    header.push("image_url");
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    lastRow = 1;
  }

  // 書き込むデータを格納する配列
  var dataToWrite = [];

  // 各キャンペーンのデータを取得して配列に格納
  for (var c = 0; c < campaigns.length; c++) {
    var currentDate = new Date(oneDaysAgo); // 開始日
    var endDate = new Date(yesterday); // 終了日

    while (currentDate <= endDate) {
      // 日付をフォーマット（YYYY-MM-DD形式）に変換
      var argDaySince = Utilities.formatDate(currentDate, "GMT", "yyyy-MM-dd");
      var nextDay = new Date(currentDate);
      nextDay.setDate(currentDate.getDate() + 30);    // ←  ◆◆◆◆◆◆◆◆◆◆何日間隔でAPIを呼び出すか◆◆◆◆◆◆◆◆◆◆

      // 指定した1日のデータを取得
      var adsData = getFacebookAdsDataForCampaign(campaigns[c].id, fields, argDaySince, argDaySince);

      // データをオブジェクト形式で格納
      if (adsData && adsData.length > 0) {
        for (var i = 0; i < adsData.length; i++) {
          var adData = adsData[i];
          var rowData = {};

          // 各フィールドに対応する値をキーとともに格納
          var keys = Object.keys(adData);
          for (var j = 0; j < keys.length; j++) {
            var key = keys[j];

          // adDataオブジェクトの各キーについて処理

          // actionsからconversionを取り出す
          if (Array.isArray(adData[key]) && key === 'actions') {

            var purchase = adData[key].find(action => action.action_type === 'offsite_conversion.fb_pixel_purchase');
            rowData['conversion_purchase'] = purchase ? purchase.value : '';

          } else if (key === 'ad_id' && adData[key]) {

          // 広告IDを書き込む
          rowData[key] = adData[key];

          // ad_idからimage_urlを取得
            var image_url = getAdImageUrl(adData[key]);
            rowData['image_url'] = image_url ? image_url : '';

          } else if (Array.isArray(adData[key])) { // 値が配列の場合

              var conversionsArray = adData[key]; // 配列を取得
              var formattedConversions = conversionsArray
                .map(function (conversion) { // 配列の各要素を処理

                  // キーがconversionsで、なおかつaction_typeがcontact_totalの場合のみ特定の処理を実行
                  if (key === "conversions" && conversion.action_type === "contact_total") {
                    return `${conversion.value}`; // JSON文字列に変換
                  } else  if(key !== "conversions") {
                    // 他の場合の処理（必要に応じて調整）
                    return `${conversion}`; // JSON文字列に変換
                  }
                })
                .join(""); // 空文字区切りで結合して1つの文字列にする
              rowData[key] = formattedConversions; // 結果をrowDataオブジェクトに格納
              
            } else {
              rowData[key] = adData[key]; // 値が配列でない場合はそのまま格納
            }
          }

          // 不足しているキーを補足して空文字を設定（ヘッダーと一致させる）
          if (header) {
            for (var k = 0; k < header.length; k++) {
              var key = header[k];
              if (!(key in rowData)) {
                rowData[key] = ""; // ヘッダーに存在するがデータにないキーに空文字を設定
              }
            }
          }

          // 配列にオブジェクトを追加
          dataToWrite.push(rowData);
        }
      }

      // リクエスト間で一定時間待機（例えば500ms）
      Utilities.sleep(5);

      // 日付を1日進める
      currentDate = nextDay;
    }
  }

  // 一度に全データをシートに書き込む
  if (dataToWrite.length > 0) {
    var formattedData = [];

    // ヘッダーの順番に沿ってデータを並べ替え
    if(header) {
      for (var row = 0; row < dataToWrite.length; row++) {
        var rowObject = dataToWrite[row];
        var formattedRow = [];

        // ヘッダー順にデータを並べ替える
        for (var h = 0; h < header.length; h++) {
          var key = header[h];
          var cellValue = rowObject[key] || ""; // データが存在しない場合は空文字
          formattedRow.push(cellValue);
        }

        // シングルコーテーションを数値文字列に付与
        for (var col = 0; col < formattedRow.length; col++) {
          if (typeof formattedRow[col] === "string" && /^[+-]?\d+(\.\d+)?$/.test(formattedRow[col])) {
            formattedRow[col] = "'" + formattedRow[col]; // 数値文字列にシングルコーテーションを追加
          }
        }
        formattedData.push(formattedRow);
      }
    } 

    // シートに書き込む
    sheet.getRange(lastRow + 1, 1, formattedData.length, header.length).setValues(formattedData);
  }
}

// 画像URL取得
function getAdImageUrl(adId) {
    console.log(`getAdImageUrl(${adId})`);
  var apiVersion = 'v22.0';
  var accessToken = facebook_getAccessToken();  // 必要に応じてアクセストークンを設定
  var url = `https://graph.facebook.com/${apiVersion}/${adId}?fields=creative&access_token=${accessToken}`;

  // Adオブジェクトから「creative」を取得
  var response = UrlFetchApp.fetch(url);
  var ad = JSON.parse(response.getContentText());
  var creativeId = ad.creative.id;

  // creativeの情報を取得
  var fields = "image_url,thumbnail_url";
  // full_image_url,thumbnail_url

  // サムネイルのサイズを拡大するパラメータ
  var thumbnail_height = "1080";
  var thumbnail_width = "1080";
  var thumbnail_size_param = `thumbnail_height=${thumbnail_height}&thumbnail_width=${thumbnail_width}`

  var creativeUrl = `https://graph.facebook.com/${apiVersion}/${creativeId}?fields=${fields}&${thumbnail_size_param}&access_token=${accessToken}`;

  var creativeResponse = UrlFetchApp.fetch(creativeUrl);
  var creative = JSON.parse(creativeResponse.getContentText());

  if (creative.image_url) {
    console.log(`getAdImageUrl 終了（返却値：${creative.image_url}）`);
    return creative.image_url;

  } else if (creative.video_thumbnail_url) {
    console.log(`getAdImageUrl 終了（返却値：${creative.video_thumbnail_url}）`);
    return creative.video_thumbnail_url;

  } else if (creative.thumbnail_url) {
    console.log(`getAdImageUrl return: ${creative.thumbnail_url}`);
    return creative.thumbnail_url;
  }

  console.log(`getAdImageUrl return: null`);
  return null;
}
