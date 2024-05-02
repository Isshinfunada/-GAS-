// 与えられた日付の週のラベルを生成（月曜始まり）
function getWeekLabel(date) {
  const day = date.getDay(); // 曜日を取得（0が日曜、1が月曜...）
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // 月曜日を基準とする
  const monday = new Date(date.setDate(diff)); // その週の月曜日の日付を取得
  const weekNumber = Math.ceil(monday.getDate() / 7); // その月の何週目かを計算
  return `${monday.getFullYear()}年${monday.getMonth() + 1}月第${weekNumber}週`;
}


function summarizeEvents(eventSheetName, summarySheetName) {
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sh1 = sp.getSheetByName(eventSheetName); // 既存のイベントが記録されているシート名をパラメータから取得
  let summarySheet = sp.getSheetByName(summarySheetName);

  // シートが既に存在する場合はそのシートを使い、存在しない場合は新規作成
  if (!summarySheet) {
    summarySheet = sp.insertSheet(summarySheetName);
  } else {
    // 既存のデータをクリアする
    summarySheet.clear();
  }

  const rows = sh1.getDataRange().getValues(); // すべてのデータを取得

  let weeklySummary = {}; // 週ごとの集計を格納するオブジェクト

  // イベントのデータを処理
  rows.forEach((row, index) => {
    if (index === 0) return; // ヘッダー行をスキップ
    const title = row[0];
    const start = new Date(row[1]);
    const duration = parseFloat(row[3]);
    const week = getWeekLabel(start);

    // タイトルと週ごとに集計
    if (!weeklySummary[week]) weeklySummary[week] = {};
    if (!weeklySummary[week][title]) weeklySummary[week][title] = { count: 0, totalHours: 0 };

    weeklySummary[week][title].count += 1;
    weeklySummary[week][title].totalHours += duration;
  });

  // 新しいシートに結果を出力
  summarySheet.appendRow(['Event Title', 'Week', 'Occurrences',  'Average Workload', 'Total Hours',]);
  for (let week in weeklySummary) {
    for (let title in weeklySummary[week]) {
      const data = weeklySummary[week][title];
      const averageWorkload = data.count > 0 ? data.totalHours / data.count : 0; // 0で割るのを避ける
      summarySheet.appendRow([title, week, data.count,  averageWorkload, data.totalHours]);
    }
  }
}
