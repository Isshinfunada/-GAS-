function main() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const sheetName = `${year}_${month}`;
  const summarySheetName = `${sheetName}_Summary`;
  const submissionSheetName = `${sheetName}_submit`;

  calendar(year, month, sheetName); // イベントデータを処理
  summarizeEvents(sheetName, summarySheetName); // サマリーを作成
  createSubmissionSheet(summarySheetName, submissionSheetName); // 提出用シートを作成
}
