function summarizeEvents(eventSheetName, summarySheetName) {
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sh1 = sp.getSheetByName(eventSheetName);
  let summarySheet = sp.getSheetByName(summarySheetName);

  if (!summarySheet) {
    summarySheet = sp.insertSheet(summarySheetName);
  } else {
    summarySheet.clear();
  }

  const rows = sh1.getDataRange().getValues();
  let eventSummary = {}; // イベントごとの集計を格納するオブジェクト

  rows.forEach((row, index) => {
    if (index === 0) return; // ヘッダー行をスキップ
    const title = row[0];
    const duration = parseFloat(row[3]);

    // タイトルごとに集計
    if (!eventSummary[title]) eventSummary[title] = { count: 0, totalHours: 0 };

    eventSummary[title].count += 1;
    eventSummary[title].totalHours += duration;
  });

  // 新しいシートに結果を出力
  summarySheet.appendRow(['Event Title', 'Occurrences', 'Average Workload', 'Total Hours']);
  for (let title in eventSummary) {
    const data = eventSummary[title];
    const averageWorkload = data.count > 0 ? data.totalHours / data.count : 0;
    summarySheet.appendRow([title, data.count, averageWorkload, data.totalHours]);
  }
}
