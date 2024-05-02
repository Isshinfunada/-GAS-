function calendar(year, month, sheetName) {
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  let sh1 = sp.getSheetByName(sheetName);

  if (!sh1) {
    sh1 = sp.insertSheet(sheetName);
    Logger.log("New sheet created: " + sheetName);
  } else {
    Logger.log("Sheet already exists: " + sheetName);
  }

  // スクリプトプロパティからメールアドレスを取得
  const scriptProperties = PropertiesService.getScriptProperties();
  const myEmailAddress = scriptProperties.getProperty('myMailAddress');

  // カレンダーを読み込む
  const cal = CalendarApp.getCalendarById(myEmailAddress);
  if (!cal) {
    Logger.log("Error: Calendar not found for the provided email address.");
    return; // カレンダーが見つからない場合、処理を停止
  }

  // カレンダーのイベントの期間を指定
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextYear = month === 12 ? year + 1 : year;
  const startTime = new Date(year, month - 1, 1); // JavaScriptの月は0から始まるため、month-1
  const endTime = new Date(nextYear, nextMonth - 1, 1);
  const events = cal.getEvents(startTime, endTime);

  Logger.log("Number of events fetched: " + events.length);

  // イベントをスプレッドシートへ書き出す
  let rowIndex = 1;
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const title = event.getTitle();
    const start = event.getStartTime();
    const end = event.getEndTime();
    const duration = (end - start) / (1000 * 60 * 60); // 時間単位で期間を計算

    // 期間が24時間のイベントと、タイトルに「ランチ」を含むイベントを除外
    if (duration === 24 || duration === 0 || title.includes("ランチ")) {
      continue;
    }

    sh1.getRange('A' + rowIndex).setValue(title);
    sh1.getRange('B' + rowIndex).setValue(start);
    sh1.getRange('C' + rowIndex).setValue(end);
    sh1.getRange('D' + rowIndex).setValue("=round((C" + rowIndex + "-B" + rowIndex + ")*24,2)");
    rowIndex++;
  }

  // すべてのデータ入力後、A列をあいうえお順にソート
  if (rowIndex > 1) { // データがある場合のみソートを実行
    sh1.getRange(1, 1, rowIndex - 1, 4).sort(1);
  }
}
