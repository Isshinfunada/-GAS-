function callender() {
//1　スプレッドシートを読み込む
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sh1 = sp.getSheetByName('202405');

//2　カレンダーをIDで読み込む
  const cal=CalendarApp.getCalendarById('isshin.funada@zuuonline.com'); 

//3　カレンダーのイベントの期間を指定
  const startTime = new Date('2024/05/01 00:00:00');
  const endTime = new Date('2024/05/31 00:00:00');
  const event = cal.getEvents(startTime,endTime); 

//4　イベントをスプレッドシートへ書き出す
  for(var i=1;i<event.length+1; i++){
    sh1.getRange('a'+i).setValue(event[i-1].getTitle());//イベントタイトル
    sh1.getRange('b'+i).setValue(event[i-1].getStartTime());//イベント開始時刻　　
    sh1.getRange('c'+i).setValue(event[i-1].getEndTime());//イベント終了時刻
    sh1.getRange('d'+i).setValue("=round((rc[-1]-rc[-2])*24,2)");//所要時間　　
  }
}