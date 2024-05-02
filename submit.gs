function createSubmissionSheet(summarySheetName, submissionSheetName) {
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = sp.getSheetByName(summarySheetName);
  let submissionSheet = sp.getSheetByName(submissionSheetName);

  // シートが既に存在する場合はそのシートを使い、存在しない場合は新規作成
  if (!submissionSheet) {
    submissionSheet = sp.insertSheet(submissionSheetName);
  } else {
    // 既存のデータをクリアする
    submissionSheet.clear();
  }

  const data = summarySheet.getDataRange().getValues();
  let uniqueTitles = {};
  let totalOccurrences = {}; // 各タイトルの発生回数の合計を格納するオブジェクト

  // イベントタイトルとその平均工数（単工数）を集計
  data.forEach((row, index) => {
    if (index === 0) return; // ヘッダー行をスキップ
    const title = row[0];
    const averageWorkload = row[3]; // 'Average Workload'は4番目の列にある前提
    const occurrences = parseFloat(row[2]); // 'Occurrences'は3番目の列にある

    if (!uniqueTitles[title]) {
      uniqueTitles[title] = averageWorkload;
      totalOccurrences[title] = 0;
    }
    totalOccurrences[title] += occurrences;
  });

  // 新しいシートに結果を出力
  submissionSheet.appendRow(['Event Title', 'Average Workload', 'Weekly Frequency']);
  for (let title in uniqueTitles) {
    const weeklyFrequency = totalOccurrences[title] / 4; // 週回数の計算
    submissionSheet.appendRow([title, uniqueTitles[title], weeklyFrequency.toFixed(2)]);
  }
}
