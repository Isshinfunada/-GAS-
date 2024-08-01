function createSubmissionSheet(summarySheetName, submissionSheetName) {
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = sp.getSheetByName(summarySheetName);
  let submissionSheet = sp.getSheetByName(submissionSheetName);

  if (!submissionSheet) {
    submissionSheet = sp.insertSheet(submissionSheetName);
  } else {
    submissionSheet.clear();
  }

  const data = summarySheet.getDataRange().getValues();

  // 新しいシートに結果を出力
  submissionSheet.appendRow(['案件', 'カテゴリ', 'タスク名', 'Average Workload', 'Weekly Frequency']); // ヘッダー行
  data.forEach((row, index) => {
    if (index === 0) return; // ヘッダー行をスキップ

    let titles = row[0].split("︙").map(title => title.trim()); // タイトルを分割
    const averageWorkload = row[2];
    const occurrences = parseFloat(row[1]);
    const weeklyFrequency = occurrences / 4;

    // 「︙」で区切られていない場合は、最初の要素をタスク名とする
    if (titles.length === 1) {
      titles = ['', '', titles[0]]; // 案件とカテゴリを空文字にする
    }

    // 空のセルを必要な数だけ追加
    while (titles.length < 3) {
      titles.push('');
    }

    // 分割したタイトルと他のデータを結合して出力
    submissionSheet.appendRow([...titles, averageWorkload, weeklyFrequency.toFixed(2)]);
  });
}
