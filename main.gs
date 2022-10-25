/**
 * スプレッドシート表示の際に呼出し
 */
function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  var subMenus = [];
  subMenus.push({
    name: "実行",
    functionName: "createSchedule"  //実行で呼び出す関数を指定
  });
  ss.addMenu("カレンダー連携", subMenus);
}

//global変数
//var calendar;

/**
 * 予定を作成
 */
function createSchedule() {

  // 連携するアカウント
  const gAccount = "";

  
  // 読み取り範囲（表の始まり行と終わり列）
  const topRow = 4;
  const lastCol = 11;

  // 0始まりで列を指定
  const nameNum = 0;
  const dateCellNum = 1;
  const dayCellNum = 3;
  const periodCellNum = 5;
  const jugyoukoumokuCellNum = 6;
  const jugyounaiyouCellNum = 7;
  const teacherCellNum = 8;
  const locationCellNum = 9;
  const statusCellNum = 10;

  // シートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 予定の最終行を取得
  var lastRow = sheet.getLastRow();
  
  //予定の一覧を取得(getValues二次元配列で取得)
  var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();

  // googleカレンダーの取得
  var calendar = CalendarApp.getCalendarById(gAccount);

  //単位初めの初期値をセット
  var day = new Date(contents[0][dateCellNum]);
  var startPeriodNum = contents[0][periodCellNum];
  var startTime = new Date(startTimes[startPeriodNum - 1]);
  var descriptionString = ""
  var name = contents[0][nameNum];
  
  
  //順に予定を作成
  for (i = 0; i <= lastRow - topRow; i++) {

    //description書き込み
    descriptionString = descriptionString.concat(contents[i][jugyoukoumokuCellNum] + "\n" + contents[i][jugyounaiyouCellNum] + "\n" + contents[i][teacherCellNum] + "\n" + "\n" );

    var status = contents[i][statusCellNum];

    //済がついているか、単位終わりではないときは飛ばす
    if (
      status == "済" || contents[i][periodCellNum] + 1 == contents[i+1][periodCellNum] && contents[i][dayCellNum] === contents[i+1][dayCellNum]
    ) {
      continue;
    }

    //済がついていないかつ単位終わりのとき
    // 単位終わりの変数ををセット
    var endPeriodNum = contents[i][periodCellNum];
    var endTime = new Date(endTimes[endPeriodNum - 1]);
    var title = startPeriodNum + "-" + endPeriodNum + name;
    var options = {location: contents[i][locationCellNum], description: descriptionString};
    
    //decideColor関数別ファイルに実装済み
    var color = decideColor(contents[i][locationCellNum], contents[i][jugyoukoumokuCellNum]);

    try {
        // 開始日時をフォーマット
        var startDate = new Date(day);
        startDate.setHours(startTime.getHours());
        startDate.setMinutes(startTime.getMinutes());
        // 終了日時をフォーマット
        var endDate = new Date(day);
        endDate.setHours(endTime.getHours());
        endDate.setMinutes(endTime.getMinutes());
        // 予定を作成
        calendar.createEvent(
          title,
          startDate,
          endDate,
          options
        ).setColor(CalendarApp.EventColor[color]);

      //予定が作成されたら「済」にする
      sheet.getRange(topRow + i, 11).setValue("済");

      //単位初めの変数を次回へ上書き
       var day = new Date(contents[i+1][dateCellNum]);
       var startPeriodNum = contents[i+1][periodCellNum];
       var startTime = new Date(startTimes[startPeriodNum - 1]);
       var descriptionString = "";
       var name = contents[i+1][nameNum];


    // エラーの場合 (ログ出力)
    } catch(e) {
      Logger.log(e);
    }
    
  }
  // ブラウザへ完了通知
  Browser.msgBox("完了");
}


//let startTimes = ["08:50:00", "10:00:00", "11:10:00", "13:00:00", "14:10:00", "15:20:00", "16:30:00"];
//let endTimes = ["9:50:00", "11:00:00", "12:10:00", "14:00:00", "15:10:00", "16:20:00", "17:30:00"];

const startTimes = [-600000, 3600000, 7800000, 14400000, 18600000, 22800000, 27000000];
const endTimes = [3000000, 7200000, 11400000, 18000000, 22200000, 26400000, 30600000];
