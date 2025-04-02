// キャンペーン情報を取得する関数
function facebook_getCampaign(daySince, dayUntil) {
  console.log(`facebook_getCampaign(${daySince},${dayUntil})`);

  // 引数が渡されていなければロギングして終了
  if (!daySince || !dayUntil) {
    console.log('日付が指定されていません。');
    return;
  }

  var sheetName = "キャンペーン";
  var endpoint = "campaigns";

  console.log(sheetName + "情報 取得開始");
  SpreadsheetApp.getActiveSpreadsheet().toast("取得処理を開始しました。", sheetName + "取得", 10);

  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign-group/insights?locale=ja_JP
  var fields = "date_start,date_stop,account_id,account_name,campaign_id,campaign_name,impressions,inline_link_clicks,conversions,spend";
  facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil);

  // メッセージを表示
  console.log(sheetName + "情報 取得完了");
  SpreadsheetApp.getUi().alert(sheetName + "情報を取得しました。");
}

// 広告セット情報を取得する関数
function facebook_getAdSets(daySince, dayUntil) {
  console.log(`facebook_getAdSets(${daySince},${dayUntil})`);

  // 引数が渡されていなければロギングして終了
  if (!daySince || !dayUntil) {
    console.log('日付が指定されていません。');
    return;
  }

  var sheetName = "広告セット";
  var endpoint = "adsets";

  console.log(sheetName + "情報 取得開始");

  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign/insights?locale=ja_JP  
  var fields = "date_start,date_stop,adset_name,impressions,inline_link_clicks,inline_link_click_ctr,cpc,spend,actions";

  // 広告セットを取得
  var adSetsCount = getAdSetsAndWriteSheet(sheetName, endpoint, fields, daySince, dayUntil);

  // 運用レポートに追記
  makeOperationReport();

  // メッセージを表示
  console.log(sheetName + "情報 取得完了");
}

// 指定された期間の1日ごとに facebook_getAdSets 関数を呼び出して運用レポートに追記
// 引数：開始日, 終了日
function getAdSetsAndMakeOperationReport(daySince, dayUntil) {
  console.log(`getAdSetsAndMakeOperationReport(${daySince}, ${dayUntil})`);

  // 引数が渡されていなければロギングして終了
  if (!daySince || !dayUntil) {
    console.log('日付が指定されていません。');
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("取得処理を開始しました。", "運用レポート記入", 10);

  var currentDate = new Date(daySince);
  var endDate = new Date(dayUntil);

  // 運用レポートに記入済みの日付のリスト
  var operationReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運用レポート');
  var operationReportData = operationReportSheet.getRange("B13:B" + operationReportSheet.getLastRow()).getValues();
  var operationReportDates = operationReportData.map(row => {
    if (row[0] instanceof Date) {
      return Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return row[0];
  });
  console.log(`operationReportDates: ${operationReportDates.join(',')}`);

  // 日数を計算
  var totalDays = Math.ceil((endDate - currentDate) / (1000 * 60 * 60 * 24)) + 1;
  var currentDayIndex = 0;

  while (currentDate <= endDate) {
    currentDayIndex++;
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    console.log(`Processing date: ${formattedDate}`);

    // 進捗をtoastで表示
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${formattedDate} のデータを取得中`,
      `${currentDayIndex}日目 / ${totalDays}日間`,
      5
    );

    // すでに運用レポートに記入済みの日付の場合はスキップ
    if (operationReportDates.includes(formattedDate)) {
      console.log(`${formattedDate}はすでに運用レポートに記入済みです。`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`${formattedDate}は運用レポートに記入済みです。`, "運用レポート記入", 5);
      currentDate.setDate(currentDate.getDate() + 1);
      continue;
    }

    // 1日分の広告セットを取得して運用レポートに追記
    facebook_getAdSets(formattedDate, formattedDate);

    // 次の日に進む
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // 「運用レポート」シートを表示
  operationReportSheet.activate();

  console.log("getAdSetsAndMakeOperationReport 完了");
  SpreadsheetApp.getUi().alert("広告セット情報を取得して運用レポートを作成しました。");
}

// 前日の広告情報を取得する関数（動作確認用）
function facebook_getAdsForYesterday() {
  console.log("facebook_getAdsForYesterday()");

  // 昨日の日付を取得
  var daySince = facebook_getDateNDaysAgo(1); // 開始日
  var dayUntil = facebook_getDateNDaysAgo(1); // 終了日

  var sheetName = "広告";
  var endpoint = "ads";

  console.log(sheetName + "情報 取得開始");

  // 広告データを取得してスプレッドシートに書き込む
  var adsCount = getAdsAndWriteSheet(daySince, dayUntil);

  if (adsCount > 0) {
    console.log(sheetName + "情報 取得完了");
  } else {
    console.log(sheetName + "情報がありませんでした。");
  }

  // CRTレポート作成
  makeCreativeReport();
}

// 広告情報を取得してスプレッドシートに書き込む関数
// 引数：開始日, 終了日
// 戻り値：取得した広告の件数
function getAdsAndWriteSheet(daySince, dayUntil) {
  console.log(`getAdsToSheet(${daySince}, ${dayUntil})`);

  // 引数が渡されていなければロギングして終了
  if (!daySince || !dayUntil) {
    console.log('日付が指定されていません。');
    SpreadsheetApp.getUi().alert('日付が指定されていません。');
    return 0;
  }

  // 処理開始のトーストを表示
  var sheetName = "広告";
  SpreadsheetApp.getActiveSpreadsheet().toast("取得処理を開始しました。", sheetName + "取得", 10);

  // 広告シートがあれば表示
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet) {
    sheet.activate();
  } else {
    // 広告シートがなければ作成
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    sheet.activate();
  }

  // Meta APIから広告データを取得する
  const fields = "campaign_id,campaign_name,adset_id,adset_name,ad_id,ad_name,impressions,cpm,inline_link_clicks,inline_link_click_ctr,cpc,actions,spend,date_start,date_stop";
  var adsData = getAdsData(fields, daySince, dayUntil);

  if (!adsData || adsData.length === 0) {
    console.log("広告データが取得できませんでした。");
    return 0;
  }

  // 広告データをスプレッドシートに書き込む
  var sheetName = "広告";
  var endpoint = "ads";
  const writtenDataCount = writeDataToSheet(sheetName, adsData, fields, endpoint, daySince, dayUntil);

  if (writtenDataCount) {
    console.log("取得した広告データをスプレッドシートに書き込みました。");
    SpreadsheetApp.getUi().alert(`${writtenDataCount}件の広告データを取得しました。内容をご確認ください。`);

  } else {
    console.log("取得した広告データの書き込みに失敗しました。");
    SpreadsheetApp.getUi().alert("広告データの書き込みに失敗しました。しばらくしてから再度お試しください。");
  }

  return writtenDataCount;
}

// 広告データを取得する関数
function getAdsData(fields, daySince, dayUntil) {
  console.log(`getAdsData(${daySince}, ${dayUntil})`);

  // スクリプトプロパティから設定値を取得
  const properties = PropertiesService.getScriptProperties();

  // URLのパスまでを作成
  const META_API_VERSION = properties.getProperty("META_API_VERSION");
  const META_AD_ACCOUNT_ID = properties.getProperty("META_AD_ACCOUNT_ID");
  const urlPath = `https://graph.facebook.com/${META_API_VERSION}/act_${META_AD_ACCOUNT_ID}/insights`;

  // URLのクエリパラメータを作成
  const META_ACCESS_TOKEN = properties.getProperty("META_ACCESS_TOKEN");
  const params = {
    access_token: META_ACCESS_TOKEN,
    level: "ad",
    fields: fields,
    sort: "spend_descending",
    limit: 10000,
    time_range: JSON.stringify({ since: daySince, until: dayUntil })
  };

  // APIを呼び出してデータを取得
  const response = UrlFetchApp.fetch(urlPath + '?' + concatUrlParams(params));

  // レスポンスコードを確認
  if (response.getResponseCode() !== 200) {
    console.log(`getAdsData APIリクエストに失敗しました。レスポンスコード: ${response.getResponseCode()}`);
    console.log(`getAdsData レスポンス: ${response.getContentText()}`);
    return null;
  }

  const responseData = JSON.parse(response.getContentText());
  const adsData = responseData.data || [];

  // spendが0のデータを除外
  const filteredAdsData = adsData.filter(ad => ad.spend > 0);

  console.log(`getAdsData 取得した広告データ件数: ${filteredAdsData.length}`);
  return filteredAdsData;
}

// 広告セット情報を取得してスプレッドシートに書き込む関数
// 引数：開始日, 終了日
// 戻り値：取得した広告セットの件数
function getAdSetsAndWriteSheet(sheetName, endpoint, fields, daySince, dayUntil) {
  console.log(`getAdSetsAndWriteSheet(${sheetName}, ${endpoint}, ${fields}, ${daySince}, ${dayUntil})`);

  // スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // シートが存在する場合は取得、存在しない場合は作成
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear(); // 既存のデータをクリア
  }

  // シートをアクティブにする
  sheet.activate();

  // Meta APIから広告セットデータを取得
  const adSetsData = getAdSetsData(fields, daySince, dayUntil);

  if (!adSetsData || adSetsData.length === 0) {
    console.log("広告セットデータが取得できませんでした。");
    sheet.getRange(1, 1).setValue("データなし");
    sheet.getRange(1, 2).setValue(daySince);
    sheet.getRange(1, 3).setValue(dayUntil);
    return 0;
  }

  // 広告セットデータを `adset_name` の昇順にソート
  adSetsData.sort((a, b) => {
    const nameA = a.adset_name ? a.adset_name.toLowerCase() : '';
    const nameB = b.adset_name ? b.adset_name.toLowerCase() : '';
    return nameA.localeCompare(nameB);
  });

// 広告セットデータをスプレッドシートに書き込む
var sheetName = "広告セット";
var endpoint = "adsets";
const writtenDataCount = writeDataToSheet(sheetName, adSetsData, fields, endpoint, daySince, dayUntil);

if (writtenDataCount) {
  console.log(writtenDataCount + "件の広告セットデータをスプレッドシートに書き込みました。");
} else {
  console.log("取得した広告セットデータの書き込みに失敗しました。");
  SpreadsheetApp.getUi().alert("広告セットデータの書き込みに失敗しました。しばらくしてから再度お試しください。");
}

return writtenDataCount;
}

// URLのクエリパラメータを作成する関数
function concatUrlParams(params) {

  const urlParams = Object.keys(params).map(key => {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }
  ).join('&');

  console.log(`URLパラメータ: ${urlParams}`);
  return urlParams;
}

// 広告セットデータを取得する関数
function getAdSetsData(fields, daySince, dayUntil) {
  console.log(`getAdSetsData(${fields}, ${daySince}, ${dayUntil})`);

  // スクリプトプロパティから設定値を取得
  const properties = PropertiesService.getScriptProperties();

  // URLのパスまでを作成
  const META_API_VERSION = properties.getProperty("META_API_VERSION");
  const META_AD_ACCOUNT_ID = properties.getProperty("META_AD_ACCOUNT_ID");
  const urlPath = `https://graph.facebook.com/${META_API_VERSION}/act_${META_AD_ACCOUNT_ID}/insights`;

  // URLのクエリパラメータを作成
  const META_ACCESS_TOKEN = properties.getProperty("META_ACCESS_TOKEN");
  const params = {
    access_token: META_ACCESS_TOKEN,
    level: "adset",
    fields: fields,
    sort: "spend_descending",
    limit: 10000,
    time_range: JSON.stringify({ since: daySince, until: dayUntil })
  };

  // APIを呼び出してデータを取得
  const response = UrlFetchApp.fetch(urlPath + '?' + concatUrlParams(params));

  // レスポンスコードを確認
  if (response.getResponseCode() !== 200) {
    console.log(`getAdSetsData APIリクエストに失敗しました。レスポンスコード: ${response.getResponseCode()}`);
    console.log(`getAdSetsData レスポンス: ${response.getContentText()}`);
    return null;
  }

  const responseData = JSON.parse(response.getContentText());
  const adSetsData = responseData.data || [];

  // spendが0のデータを除外
  const filteredAdSetsData = adSetsData.filter(adSet => adSet.spend > 0);

  console.log(`getAdSetsData 取得した広告セットデータ件数: ${filteredAdSetsData.length}`);
  return filteredAdSetsData;
}

// 前日の広告セット情報を取得する関数（定期実行用）
function facebook_getAdSetsForYesterday() {
  console.log("facebook_getAdSetsForYesterday()");

  var sheetName = "広告セット";
  var endpoint = "adsets";
  console.log(sheetName + "情報 取得開始");

  // https://developers.facebook.com/docs/marketing-api/reference/ad-campaign/insights?locale=ja_JP
  var fields = "date_start,date_stop,adset_name,impressions,inline_link_clicks,inline_link_click_ctr,cpc,spend,actions";

  // 昨日の日付を取得
  var daySince = facebook_getDateNDaysAgo(1); // 開始日
  var dayUntil = facebook_getDateNDaysAgo(1); // 終了日

  var writtenAdSetsCount = getAdSetsAndWriteSheet(sheetName, endpoint, fields, daySince, dayUntil);

  // facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil);
  // 運用レポートに広告セットシートのデータを整形して書き込む
  makeOperationReport();

  // メッセージを表示
  console.log(`${writtenAdSetsCount}件の${sheetName}情報を取得しました`);
}

// 広告情報を取得する関数
function facebook_getAds(daySince, dayUntil) {
  console.log(`facebook_getAds(${daySince},${dayUntil})`);

  // 引数が渡されていなければロギングして終了
  if (!daySince || !dayUntil) {
    console.log('日付が指定されていません。');
    return;
  }

  var sheetName = "広告";
  var endpoint = "ads";

  console.log(sheetName + "情報 取得開始");
  SpreadsheetApp.getActiveSpreadsheet().toast("取得処理を開始しました。", sheetName + "取得", 10);

  // https://developers.facebook.com/docs/marketing-api/reference/adgroup/insights/
  var fields = "campaign_id,campaign_name,adset_id,adset_name,ad_id,ad_name,impressions,cpm,inline_link_clicks,inline_link_click_ctr,cpc,actions,spend,date_start,date_stop";

  // 広告シート作成
  facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil);

  // メッセージを表示
  console.log(sheetName + "情報 取得完了");

  // CRTレポート作成
  makeCreativeReport();
}

// 「広告」シートをもとにクリエイティブレポート（「CRTレポート」シート）を作成する関数
function makeCreativeReport() {
  console.log("makeCreativeReport()");

  var reportSheetName = 'CRTレポート';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var adSheet = spreadsheet.getSheetByName('広告');
  var reportSheet = spreadsheet.getSheetByName(reportSheetName);

  SpreadsheetApp.getActiveSpreadsheet().toast("作成を開始しました。", reportSheetName + "作成", 10);

  // 「CRTレポート」シートがなければ作成、あれば取得する
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(reportSheetName);
  } else {
    reportSheet.clear(); // 既存のデータをクリア
  }

  // シートを表示
  reportSheet.activate();

  // A列の幅を35に設定
  reportSheet.setColumnWidth(1, 35);

  // 「CRTレポート」シートのB1に「※ＣＶ数で降順ソート」という値を入れる
  reportSheet.getRange('B1').setValue('※ＣＶ数で降順ソート');

  // 「【テンプレート】CRTレポート」シートがあれば取得して後続で使用する
  var template = spreadsheet.getSheetByName('【テンプレート】CRTレポート');
  if (!template) {
    Logger.log('シート "【テンプレート】CRTレポート" が見つかりません。');
    return;
  }

  // 「広告」シートからデータを取得し、広告セットごとに分類します
  var data = adSheet.getDataRange().getValues();
  var headers = data[0];
  var adSetIndex = headers.indexOf('adset_name');
  var adSetIdIndex = headers.indexOf('adset_id');
  var adNameIndex = headers.indexOf('ad_name');
  var conversionIndex = headers.indexOf('conversions');
  var spendIndex = headers.indexOf('spend');
  var imageUrlIndex = headers.indexOf('image_url');

  if (adSetIndex === -1 || conversionIndex === -1 || spendIndex === -1 || imageUrlIndex === -1) {
    console.log('「広告」シートに必要な項目が不足しています。広告データ取得を実施いただき再度ご確認ください。');
    return;
  }

  var adSets = {};

  // データを広告セットごとに分類
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var adSetName = row[adSetIndex];
    var adSetId = row[adSetIdIndex];
    var adName = row[adNameIndex];
    var conversion = parseFloat(row[conversionIndex]) || 0;
    var spend = parseFloat(row[spendIndex]) || 0;

    if (!adSets[adSetId]) {
      adSets[adSetId] = {
        adSetName: adSetName,
        ads: []
      };
    }

    adSets[adSetId].ads.push({
      adName: adName,
      conversion: conversion,
      spend: spend,
      imageUrl: row[imageUrlIndex],
      row: row
    });
  }

  // adSets を adset_name の昇順にソート
  var sortedAdSetIds = Object.keys(adSets).sort((a, b) => {
    var nameA = adSets[a].adSetName.toLowerCase();
    var nameB = adSets[b].adSetName.toLowerCase();
    return nameA.localeCompare(nameB);
  });

  var startRow = 3;
  var totalAdSets = sortedAdSetIds.length;

  for (var i = 0; i < sortedAdSetIds.length; i++) {
    var adSetId = sortedAdSetIds[i];
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "各広告セットの分析結果を記入中", `${i + 1} 件目 / ${totalAdSets} 件中`, 5);

    var adSet = adSets[adSetId];
    var ads = adSet.ads;

    // conversions、spendのトップ5を取得
    ads.sort(function (a, b) {
      return b.conversion - a.conversion || b.spend - a.spend;
    });

    var topAds = ads.slice(0, 5);

    // テンプレートをコピーして「CRTレポート」シートの最終行の下に1行あけて貼り付ける
    var templateRange = template.getRange(1, 1, 10, template.getLastColumn());
    var destinationRange = reportSheet.getRange(startRow, 1, 10, template.getLastColumn());
    templateRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    // 行の高さを設定
    reportSheet.setRowHeight(startRow + 2, 86); // 3行目の高さを86に設定
    reportSheet.setRowHeight(startRow + 3, 86); // 4行目の高さを86に設定
    reportSheet.setRowHeight(startRow + 4, 86); // 5行目の高さを86に設定
    reportSheet.setRowHeight(startRow + 5, 86); // 6行目の高さを86に設定
    reportSheet.setRowHeight(startRow + 6, 86); // 7行目の高さを86に設定
    reportSheet.setRowHeight(startRow + 8, 42); // 9行目の高さを42に設定
    reportSheet.setRowHeight(startRow + 9, 42); // 10行目の高さを42に設定

    // 広告セット名を設定
    reportSheet.getRange(startRow, 2).setValue(adSet.adSetName);

    // 広告名とデータを設定
    for (var j = 0; j < topAds.length; j++) {
      var ad = topAds[j];
      var rowIndex = startRow + 2 + j;

      reportSheet.getRange(rowIndex, 2).setValue(ad.adName);

      // C~P列に対応するデータを設定
      var imageUrl = ad.imageUrl;
      const cvr = ad.row[headers.indexOf('inline_link_clicks')] ? ad.conversion / ad.row[headers.indexOf('inline_link_clicks')] : 0;
      const cpa = ad.conversion ? ad.spend / ad.conversion : 0;
      reportSheet.getRange(rowIndex, 3).setValue(imageUrl);
      reportSheet.getRange(rowIndex, 4).setFormula(`=IMAGE("${imageUrl}")`);
      reportSheet.getRange(rowIndex, 5).setValue(ad.spend / 0.7);
      reportSheet.getRange(rowIndex, 6).setValue(ad.spend);
      reportSheet.getRange(rowIndex, 7).setValue(ad.row[headers.indexOf('impressions')]);
      reportSheet.getRange(rowIndex, 8).setValue(ad.row[headers.indexOf('cpm')]);
      reportSheet.getRange(rowIndex, 9).setValue(ad.row[headers.indexOf('inline_link_clicks')]);
      reportSheet.getRange(rowIndex, 10).setValue(ad.row[headers.indexOf('inline_link_click_ctr')]);
      reportSheet.getRange(rowIndex, 11).setValue(ad.row[headers.indexOf('cpc')]);
      reportSheet.getRange(rowIndex, 12).setValue(ad.conversion);
      reportSheet.getRange(rowIndex, 13).setValue(cvr);
      reportSheet.getRange(rowIndex, 14).setValue(cpa);
      reportSheet.getRange(rowIndex, 15).setValue(ad.row[headers.indexOf('date_start')]);
      reportSheet.getRange(rowIndex, 16).setValue(ad.row[headers.indexOf('date_stop')]);
    }

    // // 広告情報の範囲を取得してdifyChatflowApiFilesAccessを呼び出す                               ◆◆◆◆◆【Difyによる分析を無効化】
    // var adDataRange = reportSheet.getRange(startRow + 2, 2, topAds.length, 15).getValues();
    // var answerJson = difyChatflowApiFilesAccess(adDataRange, adSetId, adSet.adSetName);
    
    // // answerJsonの内容をシートに書き込む
    // if(answerJson) {
    //   reportSheet.getRange(startRow + 8, 3).setValue(answerJson.current_status);
    //   reportSheet.getRange(startRow + 9, 3).setValue(answerJson.future_implications);
    // } 

    startRow += 11; // 次の広告セットのために11行下に移動
  }

  // シートを表示
  reportSheet.activate();

  // メッセージを表示
  console.log(reportSheetName + "情報 取得完了");
  SpreadsheetApp.getUi().alert("CRTレポート作成が完了しました。内容をご確認ください。");
}

// 運用レポートに広告セットシートのデータを整形して書き込む関数
function makeOperationReport() {
  console.log("makeOperationReport()");

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId();
  console.log(`Spreadsheet ID: ${spreadsheetId}`);

  // 広告セットシートを取得
  var adSetSheet = spreadsheet.getSheetByName('広告セット');
  if (!adSetSheet) {
    console.log('広告セットシートが見つかりません。');
    return;
  }

  // 運用レポートシートを取得
  var operationReportSheetName = '運用レポート';
  var operationReportSheet = spreadsheet.getSheetByName(operationReportSheetName);
  if (!operationReportSheet) {
    console.log('運用レポートシートが見つかりません。');
    return;
  }

  // 表のタイトル（商材名）の行番号
  const tableTopRow = 10;
  const adSetWidth = 11; // 広告セット1つ分の列数

  // シート情報の読み込み
  const lastRow = adSetSheet.getLastRow();
  const lastColumn = adSetSheet.getLastColumn();
  const adSetData = lastRow > 1 ? adSetSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];
  let noDataDate = lastRow === 1 ? adSetSheet.getRange(1, 3).getValue() : '';
  noDataDate = noDataDate ? Utilities.formatDate(noDataDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  const adSetMap = {};

  // 広告セットの情報を合算
  adSetData.forEach(row => {
    const [date_start, date_stop, adset_name, impressions, inline_link_clicks, inline_link_click_ctr, cpc, spend, conversions] = row;
    if (!adSetMap[adset_name]) {
      adSetMap[adset_name] = {
        impressions: 0,
        inline_link_clicks: 0,
        cpc: 0,
        spend: 0,
        conversions: 0,
        date_stop: Utilities.formatDate(date_stop, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      };
    }
    adSetMap[adset_name].impressions += parseFloat(impressions) || 0;
    adSetMap[adset_name].inline_link_clicks += parseFloat(inline_link_clicks) || 0;
    adSetMap[adset_name].cpc += parseFloat(cpc) || 0;
    adSetMap[adset_name].spend += parseFloat(spend) || 0;
    adSetMap[adset_name].conversions += parseFloat(conversions) || 0;
  });

  // 合計値を計算
  let totalImpressions = 0;
  let totalClicks = 0;
  let totalCtr = 0;
  let totalCpc = 0;
  let totalSpend = 0;
  let totalConversions = 0;
  let dateStop = '';

  for (const adset_name in adSetMap) {
    const adSet = adSetMap[adset_name];
    totalImpressions += adSet.impressions;
    totalClicks += adSet.inline_link_clicks;
    totalSpend += adSet.spend;
    totalConversions += adSet.conversions;
    if (!dateStop) {
      dateStop = adSet.date_stop;
    }
  }

  // 全広告セットのCTR・CPCを計算
  totalCtr = totalImpressions ? (totalClicks / totalImpressions) : 0;
  totalCpc = totalSpend ? (totalSpend / totalClicks) : 0;

  const totalRow = [
    dateStop,
    totalSpend,
    totalImpressions,
    totalClicks,
    totalCtr,
    totalCpc,
    totalConversions,
    totalClicks ? totalConversions / totalClicks : 0, // 媒体CVR
    totalConversions ? totalSpend / totalConversions : 0, // 媒体CPA
    0, // 実CV
    0, // 実CVR
    0  // 実CPA
  ];

  // 運用レポートシートのデータを取得
  const operationReportData = operationReportSheet.getDataRange().getValues();
  let startRow = tableTopRow + 3;

  // 既存のdate_stopをチェック
  let existingRow = -1;
  for (let i = startRow - 1; i < operationReportData.length; i++) {
    const formattedDate = Utilities.formatDate(new Date(operationReportData[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (formattedDate === '1970-01-01') {
      continue;
    }
    Logger.log(`formattedDate: ${formattedDate}, dateStop: ${dateStop}`);
    if (formattedDate === dateStop) {
      console.log(`Date ${dateStop} already exists in the report. Skipping.`);
      return; // 日付が既に存在する場合は処理をスキップ
    }
  }

  if (existingRow === -1) {
    // 新しい行にデータを追加
    for (let i = startRow - 1; i < operationReportData.length; i++) {
      if (!operationReportData[i][1]) {
        existingRow = i + 1;
        break;
      }
    }
    if (existingRow === -1) {
      existingRow = operationReportData.length + 1;
    }
  }

  operationReportSheet.insertRowBefore(existingRow);

  if (adSetData.length === 0) {
    // 広告セットシートのデータが0行の場合、最終行の次に1列挿入してB列に日付を記入
    if (noDataDate) {
    const dateCell = operationReportSheet.getRange(existingRow, 2);
    dateCell.setValue(noDataDate);
    dateCell.setBackground('#d9d9d9'); // 背景色を#d9d9d9に設定
      dateStop = noDataDate; // dateStopにA3の値を代入
    }
  } else {
    // 全広告セットの合計値をC列からM列に入れる
    operationReportSheet.getRange(existingRow, 2, 1, totalRow.length).setValues([totalRow]);

  // B列の日付セルの背景色を#d9d9d9に設定
  const dateCell = operationReportSheet.getRange(existingRow, 2);
  dateCell.setBackground('#d9d9d9');

    // 各項目の形式を指定
    operationReportSheet.getRange(existingRow, 3).setNumberFormat('"¥"#,##0'); // spend
    operationReportSheet.getRange(existingRow, 4).setNumberFormat('#,##0'); // impressions
    operationReportSheet.getRange(existingRow, 5).setNumberFormat('#,##0'); // inline_link_clicks
    operationReportSheet.getRange(existingRow, 6).setNumberFormat('0.00%'); // inline_link_click_ctr
    operationReportSheet.getRange(existingRow, 7).setNumberFormat('"¥"#,##0'); // cpc
    operationReportSheet.getRange(existingRow, 8).setNumberFormat('#,##0'); // conversions
    operationReportSheet.getRange(existingRow, 9).setNumberFormat('0.00%'); // cvr
    operationReportSheet.getRange(existingRow, 10).setNumberFormat('"¥"#,##0'); // cpa
    operationReportSheet.getRange(existingRow, 11).setNumberFormat('#,##0'); // 実CV
    operationReportSheet.getRange(existingRow, 12).setNumberFormat('0.00%'); // 実CVR
    operationReportSheet.getRange(existingRow, 13).setNumberFormat('"¥"#,##0'); // 実CPA

    // 広告セットごとの情報を追加
    let colIndex = 14;
    for (const adset_name in adSetMap) {
      const adSet = adSetMap[adset_name];
      const adSetCtr = adSet.impressions ? (adSet.inline_link_clicks / adSet.impressions) : 0; // 各広告セットのCTRを計算
      const adSetRow = [
        adSet.spend,
        adSet.impressions,
        adSet.inline_link_clicks,
        adSetCtr,
        adSet.cpc,
        adSet.conversions,
        adSet.inline_link_clicks ? adSet.conversions / adSet.inline_link_clicks : 0, // 媒体CVR
        adSet.conversions ? adSet.spend / adSet.conversions : 0, // 媒体CPA
        0, // 実CV
        0, // 実CVR
        0  // 実CPA
      ];

      // 媒体CV、媒体CVR、媒体CPAに#NUM!が入る場合は0にする
      for (let i = 6; i <= 8; i++) {
        if (isNaN(adSetRow[i]) || !isFinite(adSetRow[i])) {
          adSetRow[i] = 0;
        }
      }

      // 広告セット名をN11, Y11, AJ11などに設定
      const adSetNameRange = operationReportSheet.getRange(tableTopRow + 1, colIndex, 1, adSetWidth);
      adSetNameRange.setValue(adset_name);
      adSetNameRange.setBackground('#ADD8E6'); // 水色背景
      adSetNameRange.merge();

      // C12:M12をN12:X12, Y12:AI12, AJ12:AT12などにコピー
      const headerRange = operationReportSheet.getRange(tableTopRow + 2, 3, 1, 11);
      headerRange.copyTo(operationReportSheet.getRange(tableTopRow + 2, colIndex, 1, adSetWidth));

      // 広告セットの値をN13:X13, Y13:AI13, AJ13:AT13などに設定
      operationReportSheet.getRange(existingRow, colIndex, 1, adSetRow.length).setValues([adSetRow]);

      // 各項目の形式を指定
      operationReportSheet.getRange(existingRow, colIndex).setNumberFormat('"¥"#,##0'); // spend
      operationReportSheet.getRange(existingRow, colIndex + 1).setNumberFormat('#,##0'); // impressions
      operationReportSheet.getRange(existingRow, colIndex + 2).setNumberFormat('#,##0'); // inline_link_clicks
      operationReportSheet.getRange(existingRow, colIndex + 3).setNumberFormat('0.00%'); // inline_link_click_ctr
      operationReportSheet.getRange(existingRow, colIndex + 4).setNumberFormat('"¥"#,##0'); // cpc
      operationReportSheet.getRange(existingRow, colIndex + 5).setNumberFormat('#,##0'); // conversions
      operationReportSheet.getRange(existingRow, colIndex + 6).setNumberFormat('0.00%'); // cvr
      operationReportSheet.getRange(existingRow, colIndex + 7).setNumberFormat('"¥"#,##0'); // cpa
      operationReportSheet.getRange(existingRow, colIndex + 8).setNumberFormat('#,##0'); // 実CV
      operationReportSheet.getRange(existingRow, colIndex + 9).setNumberFormat('0.00%'); // 実CVR
      operationReportSheet.getRange(existingRow, colIndex + 10).setNumberFormat('"¥"#,##0'); // 実CPA

      // 罫線を引く
      const rangeToBorder = operationReportSheet.getRange(tableTopRow + 1, 2, existingRow - tableTopRow, colIndex + 9);
      rangeToBorder.setBorder(true, true, true, true, true, true);

      // 背景色を設定
      const lastRow = operationReportSheet.getLastRow() - 12;
      operationReportSheet.getRange(existingRow, colIndex, 1, adSetWidth).setBackground('#FFFFFF'); // 値を入れた範囲を白色
      operationReportSheet.getRange(existingRow, colIndex + 7, 1, 1).setBackground('#fce4d6'); // 媒体CPAをオレンジ
      operationReportSheet.getRange(existingRow, colIndex + 10, 1, 1).setBackground('#fce4d6'); // 実CPAをオレンジ

      colIndex += 11; // 広告セット同士の間に不要な空の列がないようにする
    }
  }

  // データをB列の値の昇順にソート
  const rangeToSort = operationReportSheet.getRange(
    tableTopRow + 3, 2,
    operationReportSheet.getLastRow() - tableTopRow - 2, operationReportSheet.getLastColumn() - 1);
  const sortedData = rangeToSort.getValues().sort((a, b) => {
    const dateA = new Date(a[0]);
    const dateB = new Date(b[0]);
    return dateA - dateB;
  });
  rangeToSort.setValues(sortedData);

  // シートを表示
  operationReportSheet.activate();

  // メッセージを表示
  console.log(operationReportSheetName + "に広告セット情報を追記完了");
}

// スプレッドシートからアクセストークンを取得する関数
function facebook_getAccessToken() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty("META_ACCESS_TOKEN");
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
  const properties = PropertiesService.getScriptProperties();
  var accessToken = properties.getProperty("META_ACCESS_TOKEN"); // アクセストークンを取得
  var adAccountId = properties.getProperty("META_AD_ACCOUNT_ID"); // Facebook広告アカウントID
  var apiVersion = properties.getProperty("META_API_VERSION"); // 使用するAPIのバージョン

  // APIのURLを構築
  var apiUrl = `https://graph.facebook.com/${apiVersion}/act_${adAccountId}/${endpoint}`;
  var queryParams = {
    fields: "id,name", // リクエストで取得するフィールド
    access_token: accessToken, // アクセストークン
    limit: 10000                       // 取得するデータの上限
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
function getFacebookAdsDataForCampaign(campaignId, fields, argDaySince, argDayUntil) {
  console.log(`getFacebookAdsDataForCampaign(${campaignId},${fields},${argDaySince},${argDayUntil})`);

  // 必要な変数と設定を準備
  const properties = PropertiesService.getScriptProperties();
  var apiVersion = properties.getProperty("META_API_VERSION"); // 使用するAPIのバージョン
  var accessToken = facebook_getAccessToken(); // アクセストークンを取得

  // APIリクエスト用のURLを構築
  var apiUrl = `https://graph.facebook.com/${apiVersion}/${campaignId}/insights`;
  var queryParams = {
    level: "ad",
    fields: fields,
    "time_range[since]": argDaySince,
    "time_range[until]": argDayUntil,
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

// データを取得してスプレッドシートに書き込む関数
function facebook_writeFacebookAdsDataToSheet(sheetName, endpoint, fields, daySince, dayUntil) {
  console.log(`facebook_writeFacebookAdsDataToSheet(${sheetName}, ${endpoint}, ${fields}, ${daySince}, ${dayUntil})`);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシートを取得


  // 指定された種類のデータをAPIで取得
  var data = facebook_getData(endpoint);

  if (!data) {
    console.log(`データが取得できませんでした。エンドポイント: ${endpoint}`);
    return;
  }

  // データをスプレッドシートに書き込む
  const writtenData = writeDataToSheet(sheetName, data, fields, endpoint, daySince, dayUntil);

  // sheetNameシートを表示する
  sheet.activate();

  // 書き込んだデータの件数を返却
  return writtenData;
}

// データをスプレッドシートに書き込む関数
function writeDataToSheet(sheetName, data, fields, endpoint, daySince, dayUntil) {

  // シートが存在する場合は、そのシートを削除
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    // シートのすべてをクリアする
    sheet.clear();
  } else {
    // シートが存在しない場合は新しいシートを作成
    sheet = spreadsheet.insertSheet(sheetName);
  }

  // シートを表示
  sheet.activate();

  if (!data || data.length === 0) {
    Logger.log("データが取得できませんでした。");
    sheet.getRange(1, 1).setValue("データなし");
    sheet.getRange(1, 2).setValue(daySince);
    sheet.getRange(1, 3).setValue(dayUntil);
    return 0;
  }

  // ヘッダーに使用するデータを取得
  var sampleData = data[0];

  if (!sampleData || Object.keys(sampleData).length === 0) {
    Logger.log("ヘッダーに使用するデータが取得できませんでした。");
    sheet.getRange(1, 1).setValue("データなし");
    sheet.getRange(1, 2).setValue(daySince);
    sheet.getRange(1, 3).setValue(dayUntil);
    return 0;
  }

  var header = Object.keys(sampleData);

  // 広告の場合はヘッダーに画像URLを追加
  if (endpoint === 'ads') {
    header.push("image_url");
  }

  // ヘッダーのactionsをconversionsに変更
  header = header.map(h => h === 'actions' ? 'conversions' : h);

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  lastRow = 1;

  // 書き込むデータを格納する配列
  var dataToWrite = [];
  var dataCount = data.length; // データの件数

  // 各データの値を整形して配列に格納
  for (var i = 0; i < dataCount; i++) {

    // 何件目の広告データを処理するかをtoastで表示
    if (endpoint === 'ads') {
      SpreadsheetApp.getActiveSpreadsheet().toast(
      sheetName + "データを処理中", `${i + 1} 件目 / ${dataCount} 件`, 5);
    }

    var adData = data[i];
    var rowData = {};

    // 各フィールドに対応する値をキーとともに格納
    var keys = Object.keys(adData);
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];

      // adDataオブジェクトの各キーについて処理
      if (Array.isArray(adData[key]) && key === 'actions') { // 広告のconversions
        var purchase = adData[key].find(action => action.action_type === 'offsite_conversion.fb_pixel_purchase');
        rowData['conversions'] = purchase ? purchase.value : '';

      } else if (key === 'actions' && adData[key].action_type === 'web_in_store_purchase') { // 広告セットのconversions
        var purchase = adData[key].value;
        rowData['conversions'] = purchase ? purchase.value : '';

      } else if (key === 'ad_id' && adData[key]) {
        rowData[key] = adData[key];
        if (endpoint === 'ads') {
          var image_url = getAdImageUrl(adData[key]);
          rowData['image_url'] = image_url ? image_url : '';
        }
      } else if (Array.isArray(adData[key])) {
        var conversionsArray = adData[key];
        var formattedConversions = conversionsArray
          .map(function (conversion) {
            if (key === "conversions" && conversion.action_type === "contact_total") {
              return `${conversion.value}`;
            } else if (key !== "conversions") {
              return `${conversion}`;
            }
          })
          .join("");
        rowData[key] = formattedConversions;
      } else {
        rowData[key] = adData[key];
      }
    }

    if (header) {
      for (var k = 0; k < header.length; k++) {
        var key = header[k];
        if (!(key in rowData)) {
          rowData[key] = "";
        }
      }
    }

    dataToWrite.push(rowData);
  }

    if (dataToWrite.length > 0) {
    var formattedData = [];

    if (header) {
      for (var row = 0; row < dataToWrite.length; row++) {
        var rowObject = dataToWrite[row];
        var formattedRow = [];

        for (var h = 0; h < header.length; h++) {
          var key = header[h];
          var cellValue = rowObject[key] || "";
          formattedRow.push(cellValue);
        }

        for (var col = 0; col < formattedRow.length; col++) {
          if (typeof formattedRow[col] === "string" && /^[+-]?\d+(\.\d+)?$/.test(formattedRow[col])) {
            formattedRow[col] = "'" + formattedRow[col];
          }
        }
        formattedData.push(formattedRow);
      }
    }

    // データをスプレッドシートに書き込む
    sheet.getRange(lastRow + 1, 1, formattedData.length, header.length).setValues(formattedData);

  } else {
    // データが0件の場合
    sheet.getRange(1, 1).setValue("データなし");
    sheet.getRange(1, 2).setValue(daySince);
    sheet.getRange(1, 3).setValue(dayUntil);
  }

  return formattedData.length; // 書き込んだデータの件数を返却
}

// 画像URL取得
function getAdImageUrl(adId) {
  console.log(`getAdImageUrl(${adId})`);
  const properties = PropertiesService.getScriptProperties();
  var apiVersion = properties.getProperty("META_API_VERSION");  // 使用するAPIのバージョン
  var accessToken = facebook_getAccessToken();  // アクセストークンを取得
  var url = `https://graph.facebook.com/${apiVersion}/${adId}?fields=creative&access_token=${accessToken}`;

  // Adオブジェクトから「creative」を取得
  var response = UrlFetchApp.fetch(url);
  var ad = JSON.parse(response.getContentText());
  var creativeId = ad.creative.id;

  // creativeの情報を取得
  var fields = "image_url,thumbnail_url";
  // full_image_url,thumbnail_url

  // サムネイルをオリジナルサイズにするパラメータ
  var thumbnail_height = "0";
  var thumbnail_width = "0";
  var thumbnail_size_param = `thumbnail_height=${thumbnail_height}&thumbnail_width=${thumbnail_width}`

  console.log(`getAdImageUrl（creativeId：${creativeId}）`);

  var creativeImageUrl = `https://graph.facebook.com/${apiVersion}/${creativeId}?fields=${fields}&${thumbnail_size_param}&access_token=${accessToken}`;

  var creativeResponse = UrlFetchApp.fetch(creativeImageUrl);
  var creative = JSON.parse(creativeResponse.getContentText());

  if (creative.image_url) {
    console.log(`getAdImageUrl 終了（返却値：${creative.image_url}）`);
    return creative.image_url;

  } else if (creative.thumbnail_url) {
    console.log(`getAdImageUrl 終了（返却値: ${creative.thumbnail_url}`);
    return creative.thumbnail_url;
  }

  console.log(`getAdImageUrl 終了（返却値: null`);
  return null;
}
