// 実行ボタンをスプレッドシート側に作成
function onOpen() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  const subMenus = [];
  subMenus.push({
    name: '実行',
    functionName: 'createSchedule'
  });
  sheet.addMenu('カレンダー連携', subMenus);
}

function createSchedule() {

  // 連携するアカウント
  const calendarId = CALENDAR_ID; 

  // 現在日時の取得
  const currentDay = new Date()
  const currentYear = currentDay.getFullYear()
  const currentMonth = currentDay.getMonth() + 1

  // シートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${currentYear}年${currentMonth}月`);

  // メンバー数を取得
  const memberNum = Math.floor((sheet.getLastRow() - 1) / 6)

  // 長さがメンバー数の配列を作成
  // schedules[人][行][列]
  const schedules = Array.from({ length: memberNum })

  // 配列に一人ずつの予定を入れる
  schedules.forEach((_schedule, index) => {
    // ひとり６行あるので index * 6 すると次の人の予定が入る
    const initRow = 2 + index * 6
    const initCol = 3
    const rowRange = 4
    const colRange = sheet.getLastColumn() - 4
    schedules[index] = sheet.getRange(initRow, initCol, rowRange, colRange).getValues()
  })

  // カレンダーの取得
  const calendar = CalendarApp.getCalendarById(calendarId)

  // 日付の取得
  const days = sheet.getRange(1, 3, 1, sheet.getLastColumn() - 4).getValues()

  // 予定を作成
  for (let i = 0; i < memberNum; i++) {
    for(let j = 0; j < sheet.getLastColumn() - 4; j++) {
      
      // 誰がいつ出社するかを取得
      const day = new Date(days[0][j])
      const startTime = schedules[i][2][j]
      const endTime = schedules[i][3][j]
      const name = sheet.getRange(2 + i * 6, 1).getValue()
      
      // カレンダーの予定を取得
      const events = calendar.getEventsForDay(day)
      
      // 自分の出社予定が入っていたら削除する
      events.forEach(event => {
        const eventTitle = event.getTitle()
        if (eventTitle === name) {
          event.deleteEvent()
        }
      })

      // 出社しない場合は、飛ばす。
      if (schedules[i][0][j] === '') continue;
      
      // 開始日時をフォーマット
      const startDate = new Date(day)
      startDate.setHours(startTime.getHours())
      startDate.setMinutes(startTime.getMinutes());

      // 終了日時をフォーマット
      const endDate = new Date(day);
      endDate.setHours(endTime.getHours())
      endDate.setMinutes(endTime.getMinutes());
      
      // 予定を作成
      calendar.createEvent(
        name,
        startDate,
        endDate
      );
    }
  }  
}
