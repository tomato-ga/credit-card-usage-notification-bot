function processEmails() {
  console.log('processEmails開始');
  
  // 今日の日付を取得
  var today = new Date();
  var todayString = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');

  // 昨日の日付を取得
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  var yesterdayString = Utilities.formatDate(yesterday, 'JST', 'yyyy/MM/dd');

  // 明日の日付を取得（クエリの終端として使用）
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowString = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  console.log('昨日の日付: ' + yesterdayString);
  console.log('今日の日付: ' + todayString);
  console.log('明日の日付: ' + tomorrowString);

  // Gmailから昨日と今日の日付で、件名に「ご利用のお知らせ」または「ご利用明細のお知らせ」と「三井住友カード」が含まれるメールを検索
  var query = 'subject:("ご利用のお知らせ" OR "ご利用明細のお知らせ") "三井住友カード" after:' + yesterdayString + ' before:' + tomorrowString;
  var threads = GmailApp.search(query);
  
  console.log('該当するスレッド数: ' + threads.length);
  
  var spreadsheetUrl = 'スプレッドシートURL';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheetName = getMonthSheetName();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  console.log('スプレッドシートURL: ' + spreadsheetUrl);
  console.log('シート名: ' + sheetName);
  
  // 指定した月のシートが存在しない場合は作成する
  if (!sheet) {
    console.log('新しいシートを作成します');
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(['利用日時', '決済場所', '利用金額', 'メッセージID']);
  }
  
  // 「メッセージID」列が存在しない場合は追加する
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('メッセージID') === -1) {
    console.log('「メッセージID」列を追加します');
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue('メッセージID');
  }
  
  try {
    threads.forEach(function(thread) {
      console.log('スレッド処理開始');
      thread.getMessages().forEach(function(message) {
        console.log('メッセージ処理開始');
        var messageId = message.getId();
        if (!isProcessed(sheet, messageId)) {
          console.log('未処理のメッセージ');
          var body = message.getPlainBody(); // HTMLではなくプレーンテキストを取得
          console.log('メール本文:', body); // デバッグ出力
          var data = extractData(body);
          console.log('抽出データ: ', data);
          data.forEach(function(item) {
            sheet.appendRow([item.date, item.location, item.amount, messageId]);
            console.log('データを抽出: ' + item.location + ', ' + item.amount + '円');
          });
        } else {
          console.log('処理済みのメッセージ');
        }
      });
    });
  } catch (e) {
    console.error('エラーが発生しました: ' + e);
  }
  
  // データ追加処理の後に呼び出し
  sortSheetByDate();
  
  console.log('processEmails完了');
}

function extractData(body) {
  var data = [];

  // 正規表現パターンを若干緩和して、より多くのフォーマットに対応できるようにします
  // var recordRegex = /ご利用日時：\s*(\d{4}\/\d{2}\/\d{2})(?: \d{2}:\d{2})?.*?\n(.*?)（買物）\s*([\d,]+)円/g;
  // var recordRegex = /利用日：\s*(\d{4}\/\d{2}\/\d{2})(?: \d{2}:\d{2})?.*?\n(.*?)（買物）\s*([\d,]+)円/g;
  // var recordRegex = /◇利用日：\s*(\d{4}\/\d{2}\/\d{2})(?: \d{2}:\d{2})?.*?◇利用先：\s*(.*?)\s*◇利用取引：.*?\s*◇利用金額：\s*([\d,]+)円/g;
  // var recordRegex = /◇利用日：\s*(\d{4}\/\d{2}\/\d{2} \d{2}:\d{2})\s*◇利用先：\s*(.*?)\s*◇利用取引：.*?\s*◇利用金額：\s*([\d,]+)円/g;
  // var recordRegex = /◇利用日：\s*(\d{4}\/\d{2}\/\d{2})(?: \d{2}:\d{2})?\s*◇利用先：\s*(.*?)\s*◇利用取引：.*?\s*◇利用金額：\s*([\d,]+)円/g; // OK正規表現
  var recordRegex = /◇利用日：\s*(\d{4}\/\d{2}\/\d{2}(?: \d{2}:\d{2})?)\s*◇利用先：\s*(.*?)\s*◇利用取引：.*?\s*◇利用金額：\s*([\d,]+)円/g;


  // マッチング前にログ出力を追加
  console.log("マッチングを開始します。正規表現：", recordRegex);
  console.log("対象のメール本文：", body);

  var match;
  while ((match = recordRegex.exec(body)) !== null) {
      console.log("マッチング成功：", match[0]); // マッチした全体の文字列をログ出力

      var date = match[1]; // 'YYYY/MM/DD' 形式
      var location = match[2].trim();
      var amount = parseInt(match[3].replace(/,/g, ''), 10); // カンマを除去し、整数に変換

      data.push({
          date: date,
          location: location,
          amount: amount
      });
  }

  if (data.length === 0) {
      console.log("抽出データが見つかりませんでした。正規表現とメール本文を確認してください。");
  }

  return data;
}

function sortSheetByDate() {
  var spreadsheetUrl = 'スプレッドシートURL';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheetName = getMonthSheetName(); // 現在の月のシートを取得する関数
  var sheet = spreadsheet.getSheetByName(sheetName);

  // データが存在する範囲の最終行を取得
  var lastRow = sheet.getLastRow();
  
  // データが存在する範囲の最終列を取得
  var lastColumn = sheet.getLastColumn();

  // ヘッダー行を除いた範囲を指定（2行目から開始）
  // 注意: getRangeの引数は (開始行, 開始列, 行数, 列数)
  var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  
  // 列A（インデックスは1から始まる）を基準に昇順でソート
  range.sort({column: 1, ascending: true});
  
  console.log('シートを時系列でソートしました。');
}


function isProcessed(sheet, messageId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }
  var processedIds = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  return processedIds.includes(messageId);
}

function sendDailyReport() {
  console.log('sendDailyReport開始');
  
  var spreadsheetUrl = 'スプレッドシートURL';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheetName = getMonthSheetName(); // 現在の月のシートを取得する関数
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // 一日の利用金額と月の合計金額を計算
  var dailyTotal = getDailyTotal(sheet);
  var monthlyTotal = getMonthlyTotal(sheet);
  
  console.log('一日の利用金額: ' + dailyTotal + '円');
  console.log('月の合計金額: ' + monthlyTotal + '円');
  
  // LINEにメッセージを送信
  sendLineMessage(dailyTotal, monthlyTotal);
  
  console.log('sendDailyReport完了');
}

function getMonthSheetName() {
  var now = new Date();
  return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月');
}

function getDailyTotal(sheet) {
  var today = new Date();
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  
  var dailyTotal = 0;
  values.forEach(function(row) {
    var date = new Date(row[0]);
    if (date.toDateString() === today.toDateString()) {
      dailyTotal += row[2];
    }
  });
  
  return dailyTotal;
}

function getMonthlyTotal(sheet) {
  var values = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  return values.reduce(function(total, row) {
    return total + row[0];
  }, 0);
}

function sendLineMessage(dailyTotal, monthlyTotal) {
  var token = 'LINE Messaging API token';
  var message = '本日の利用金額: ￥' + dailyTotal + '\n今月の合計利用金額: ￥' + monthlyTotal;
  
  var payload = JSON.stringify({
    'to': 'チャネルID QR CODE', // 宛先を指定
    'messages': [{
      'type': 'text',
      'text': message
    }]
  });
  
  var options = {
    'method': 'post',
    'contentType': 'application/json', // コンテンツタイプを指定
    'headers': {
      'Authorization': 'Bearer ' + token
    },
    'payload': payload,
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
  console.log('LINEメッセージ送信結果: ' + response.getContentText());
}
