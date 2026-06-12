/**
 * スプレッドシートの予定をGoogleカレンダーに同期するメイン関数
 */
function syncGoogleCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 1. 1行目・B1セルからカレンダーIDを取得
  const calendarId = sheet.getRange(1, 2).getValue().toString().trim();
  if (!calendarId || !calendarId.includes('@')) {
    SpreadsheetApp.getUi().alert("エラー: 1行目B列(B1セル)に正しいGoogleカレンダーIDを入力してください。");
    return;
  }
  
  let calendar;
  try {
    calendar = CalendarApp.getCalendarById(calendarId);
  } catch(e) {
    SpreadsheetApp.getUi().alert("エラー: カレンダーが見つかりません。アクセス権限を確認してください。");
    return;
  }

  // 2. データ範囲の取得（3行目から最終行まで）
  const startRow = 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert("処理対象のデータがありません。");
    return;
  }

  // A3からI列（9列分）のデータを一括取得して処理を高速化
  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 9);
  const data = dataRange.getValues();
  const currentDate = new Date();

  // 3. カレンダーの予定をあらかじめ一括取得（過去3ヶ月〜未来15ヶ月）
  const startOfRange = new Date(currentDate.getFullYear(), currentDate.getMonth() - 3, 1);
  const endOfRange   = new Date(currentDate.getFullYear(), currentDate.getMonth() + 15, 0, 23, 59, 59);
  const existingEvents = calendar.getEvents(startOfRange, endOfRange);

  // 4. ループ処理
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentRowNum = startRow + i; // 実際の行番号

    var status      = row[0].toString().trim(); // A列
    var title       = row[1].toString().trim(); // B列
    var startDateStr = row[2];                   // C列
    var startTimeStr = row[3];                   // D列
    var endDateStr   = row[4];                   // E列
    var endTimeStr   = row[5];                   // F列
    var description  = row[6] || '';             // G列
    var location     = row[7] || '';             // H列
    var tg_id        = row[8].toString().trim(); // I列: 管理ID

    // タイトルが空ならスクラップ行としてスキップ
    if (!title) continue;

    // 日付オブジェクトの解析（曜日入り文字列にも対応）
    const startDateTime = parseDateTime_(startDateStr, startTimeStr);
    const endDateTime = parseDateTime_(endDateStr, endTimeStr);

    if (!startDateTime || !endDateTime) {
      console.log(`行 ${currentRowNum}: 日付のパースに失敗したためスキップしました。`);
      continue;
    }

    // 過去のイベントなら、ステータスに関わらず一律で背景色を灰色にする
    if (startDateTime < currentDate) {
      sheet.getRange(currentRowNum, 1, 1, sheet.getLastColumn()).setBackground('#d9d9d9');
    }

    // すでにステータスが「完了」の行はカレンダー処理をスキップ
    if (status === '完了') continue;

    // 管理ID（I列）が空欄なら、プログラムが被らないIDを自動生成
    if (tg_id === '') {
      const now = new Date();
      tg_id = 'id_' + now.getFullYear() + 
              ('0' + (now.getMonth() + 1)).slice(-2) + 
              ('0' + now.getDate()).slice(-2) + '_' + 
              ('0' + now.getHours()).slice(-2) + 
              ('0' + now.getMinutes()).slice(-2) + 
              ('0' + now.getSeconds()).slice(-2) + '_' + i;
              
      sheet.getRange(currentRowNum, 9).setValue(tg_id); // シートのI列に書き込み
      console.log(`行 ${currentRowNum}: 新しい管理ID [${tg_id}] を自動発行しました。`);
    }

    // メモリ上の既存カレンダー予定から隠しタグが一致するものを探す
    let existingEvent = null;
    for (let j = 0; j < existingEvents.length; j++) {
      if (existingEvents[j].getTag('searchId') === tg_id) {
        existingEvent = existingEvents[j];
        break;
      }
    }

    // 時刻の指定があるかないかで終日イベントかを判定
    const isAllDay = (startTimeStr === "" && endTimeStr === "");
    const options = { description: description, location: location };

    if (existingEvent) {
      // 【更新処理】すでにカレンダーに予定が存在する場合
      existingEvent.setTitle(title);
      existingEvent.setDescription(description);
      existingEvent.setLocation(location);
      
      if (isAllDay) {
        // 終日の場合は終了日を+1日する（Googleカレンダーの仕様対策）
        const adjustedEndDate = new Date(endDateTime.getFullYear(), endDateTime.getMonth(), endDateTime.getDate() + 1);
        existingEvent.setAllDayDates(startDateTime, adjustedEndDate);
      } else {
        // ★改善ポイント: 一度既存の予定が終日モードになっていた場合を想定し、
        // 1分間だけのダミー時刻を指定して「終日モード」を確実に強制解除してから、正しい時刻をセットする
        existingEvent.setTime(startDateTime, new Date(startDateTime.getTime() + 60000)); 
        existingEvent.setTime(startDateTime, endDateTime); // 正しい開始・終了時刻をセット
      }
      console.log(`行 ${currentRowNum}: 予定を更新しました。 (ID: ${tg_id})`);

    } else {
      // 【新規作成】カレンダーに予定がない場合
      let newEvent;
      if (isAllDay) {
        const adjustedEndDate = new Date(endDateTime.getFullYear(), endDateTime.getMonth(), endDateTime.getDate() + 1);
        newEvent = calendar.createAllDayEvent(title, startDateTime, adjustedEndDate, options);
      } else {
        newEvent = calendar.createEvent(title, startDateTime, endDateTime, options);
      }
      
      // 作成した予定の「裏データ」に管理IDを隠しタグとして埋め込む
      newEvent.setTag("searchId", tg_id);
      console.log(`行 ${currentRowNum}: 新規予定を追加しました。 (ID: ${tg_id})`);
    }

    // 処理が完了した行のA列を「完了」にする
    sheet.getRange(currentRowNum, 1).setValue("完了");
  }
  SpreadsheetApp.getUi().alert("カレンダーとの同期が完了しました！");
}

/**
 * 曜日入り文字列や日付オブジェクトを安全にDateオブジェクトへ変換する内部関数
 */
function parseDateTime_(dateInput, timeInput) {
  if (!dateInput) return null;
  
  if (dateInput instanceof Date) {
    const year = dateInput.getFullYear();
    const month = dateInput.getMonth();
    const day = dateInput.getDate();
    if (timeInput && timeInput instanceof Date) {
      return new Date(year, month, day, timeInput.getHours(), timeInput.getMinutes(), timeInput.getSeconds());
    } else if (timeInput && timeInput.toString().trim() !== '') {
      const tParts = timeInput.toString().split(':');
      return new Date(year, month, day, parseInt(tParts[0], 10) || 0, parseInt(tParts[1], 10) || 0, 0);
    }
    return new Date(year, month, day);
  }

  let dateStr = dateInput.toString().replace(/\([^)]+\)/g, "").replace(/\/+/g, "/").trim();
  if (dateStr.endsWith("/")) dateStr = dateStr.slice(0, -1);

  const dateParts = dateStr.split("/");
  if (dateParts.length !== 3) return null;

  const y = parseInt(dateParts[0], 10);
  const m = parseInt(dateParts[1], 10) - 1;
  const d = parseInt(dateParts[2], 10);

  if (timeInput && timeInput.toString().trim() !== "") {
    const timeParts = timeInput.toString().split(":");
    const hh = parseInt(timeParts[0], 10) || 0;
    const mm = parseInt(timeParts[1], 10) || 0;
    return new Date(y, m, d, hh, mm, 0);
  }
  return new Date(y, m, d);
}

/**
 * スプレッドシートを開いたときに、上部メニューに実行ボタンを追加する関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 カレンダー連携')
    .addItem('カレンダーと同期する', 'syncGoogleCalendar')
    .addToUi();
}
