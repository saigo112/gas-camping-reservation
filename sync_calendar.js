/**
 * sync_calendar.gs
 * スプレッドシートの内容をGoogleカレンダーに反映させる（追加・削除）
 */

function syncCalendarEvents() {
  const CFG = getAppConfig_();
  const SHEET_ID = CFG.SHEET_ID;
  const SHEET_NAME = CFG.extractor.SHEET_NAME; // 環境に合わせてシート名を切替
  const CAL_ID = CFG.CALENDAR_ID;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    // ヘッダーから列番号を特定
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i + 1; });

    // 必要な列があるかチェック
    const required = ['予約ID', 'ステータス', 'チェックイン日時', 'チェックアウト日時', '名前', 'カレンダーイベントID'];
    if (!required.every(k => colMap[k])) {
      Logger.log('必要な列が見つかりません。シートのヘッダーを確認してください。');
      return;
    }

    const calendar = CalendarApp.getCalendarById(CAL_ID);
    if (!calendar) throw new Error(`カレンダーが見つかりません ID: ${CAL_ID}`);

    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const data = dataRange.getValues();

    // 更新内容を記録する配列
    const updates = [];
    const now = new Date();

    for (let i = 0; i < data.length; i++) {
      const rowNum = i + 2;
      const rowData = data[i];

      const rId = String(rowData[colMap['予約ID'] - 1]);
      const platform = String(rowData[colMap['予約元'] - 1] || '楽天トラベル');
      const status = String(rowData[colMap['ステータス'] - 1]).trim();
      const eventId = String(rowData[colMap['カレンダーイベントID'] - 1]);
      const checkIn = rowData[colMap['チェックイン日時'] - 1];
      const checkOut = rowData[colMap['チェックアウト日時'] - 1];
      const name = String(rowData[colMap['名前'] - 1]);
      const siteName = String(rowData[colMap['サイト名'] - 1] || '');

      // === パターンA：新規作成（予約中 ＆ IDなし ＆ 未来の日付） ===
      if (status === '予約中' && eventId === '' && isValidDate_(checkIn) && checkIn > now) {

        const title = `【${platform}】【予約ID:${rId}】${name}様 (${siteName})`;
        const desc = `予約元: ${platform}\n予約ID: ${rId}\nサイト: ${siteName}\n名前: ${name}\n自動連携により作成`;

        // 終了時間の補正（なければ開始1時間後）
        let endDt = isValidDate_(checkOut) ? checkOut : new Date(checkIn.getTime() + (60 * 60 * 1000));

        try {
          const event = calendar.createEvent(title, checkIn, endDt, { description: desc });
          // 作成したイベントIDを記録
          updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: event.getId() });
          Logger.log(`カレンダー追加: ${title}`);
        } catch (e) {
          Logger.log(`カレンダー追加失敗 Row${rowNum}: ${e.message}`);
        }
      }

      // === パターンB：削除（キャンセル済み ＆ IDあり） ===
      else if (status === 'キャンセル済み' && eventId !== '') {
        try {
          const event = calendar.getEventById(eventId);
          if (event) {
            event.deleteEvent();
            Logger.log(`カレンダー削除: ID ${eventId}`);
          }
          // シートのIDを空にする
          updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: '' });
        } catch (e) {
          Logger.log(`カレンダー削除失敗 Row${rowNum}: ${e.message}`);
          // イベントが既にない場合もIDは消しておく
          updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: '' });
        }
      }
    }

    // シートへの書き込み（まとめて実行）
    if (updates.length > 0) {
      updates.forEach(u => {
        sheet.getRange(u.row, u.col).setValue(u.value);
      });
      SpreadsheetApp.flush();
    }

  } catch (e) {
    Logger.log('Sync Error: ' + e.stack);
  } finally {
    lock.releaseLock();
  }
}

// 日付判定用ヘルパー関数
function isValidDate_(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

/**
 * ▼▼▼ これをトリガーに設定してください ▼▼▼
 * 全体のメイン処理（抽出 → カレンダー同期）
 */
function mainSequence() {
  // 1. Gmailから抽出 & チェックイン完了などのステータス更新
  extractRakutenTravelEmails();

  // 2. カレンダーへの反映（追加・削除）
  syncCalendarEvents();
}