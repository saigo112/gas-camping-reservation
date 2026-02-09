/**
 * webapp.gs
 * Webアプリのバックエンド処理
 */

// 1. Webページを表示する関数（必須）
function doGet() {
  // index.html というファイルを表示する
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('キャンプ場予約管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. スプレッドシートのデータを取得してHTMLに送る関数
function getDataForWeb() {
  const CFG = getAppConfig_();
  const sheet = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(CFG.extractor.SHEET_NAME);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

  const displayData = data.map(row => {
    // 日付オブジェクトを取得
    const startObj = row[colMap['チェックイン日時']];
    const endObj = row[colMap['チェックアウト日時']];

    return {
      id: row[colMap['予約ID']],
      platform: row[colMap['予約元']] || '楽天トラベル', // 予約元（デフォルトは楽天）
      status: row[colMap['ステータス']],
      // 表示用の短い形式（既存）
      checkIn: formatDate_(startObj),
      checkOut: formatDate_(endObj),
      // ★追加：カレンダー描画用の完全な日付データ（ISO文字列）
      startIso: (startObj instanceof Date) ? startObj.toISOString() : null,
      endIso: (endObj instanceof Date) ? endObj.toISOString() : null,

      name: row[colMap['名前']],
      site: row[colMap['サイト名']],
      siteCount: row[colMap['サイト数']],
      phone: row[colMap['電話番号']],
      email: row[colMap['メールアドレス']],
      price: row[colMap['料金']],
      adult: row[colMap['大人']],
      child: row[colMap['子供']],
      infant: row[colMap['幼児']],
      remarks: row[colMap['備考']]
    };
  });

  return displayData;
}

// 日付フォーマット用ヘルパー
function formatDate_(date) {
  if (!date || !(date instanceof Date)) return date;
  return Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd HH:mm');
}

// 3. ステータスを更新する関数（画面から呼ばれる）
function updateReservationStatus(reservationId, newStatus) {
  const CFG = getAppConfig_();
  const sheet = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(CFG.extractor.SHEET_NAME);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = {};
  headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

  // 予約IDの列番号（0始まり）
  const idColIndex = colMap['予約ID'];
  const statusColIndex = colMap['ステータス'];

  // IDが一致する行を探す
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(reservationId)) {
      // ステータスを更新（行番号は i+1）
      sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);

      // ★ここでカレンダー同期も実行してしまう（即時反映）
      // sync_calendar.gs の関数を呼ぶ
      if (typeof syncCalendarEvents === 'function') {
        syncCalendarEvents();
      }

      return { success: true, message: `ID:${reservationId} を「${newStatus}」に更新しました` };
    }
  }

  return { success: false, message: 'IDが見つかりませんでした' };
}

// 4. メール設定を読み込む関数
function getEmailSettings() {
  const CFG = getAppConfig_();
  // メール設定シート名を取得（configになければデフォルト値）
  const sheetName = CFG.mailer.SETTING_SHEET_NAME || 'メール設定';
  const sheet = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(sheetName);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const settings = [];

  // 1行目はヘッダーなので2行目から
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (!key) continue;

    settings.push({
      key: key,
      subject: data[i][1],
      body: data[i][2],
      attachments: data[i][3]
    });
  }
  return settings;
}

// 5. メール設定を保存する関数
function saveEmailSettings(newSettings) {
  const CFG = getAppConfig_();
  const sheetName = CFG.mailer.SETTING_SHEET_NAME || 'メール設定';
  const sheet = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(sheetName);

  if (!sheet) return { success: false, message: '設定シートが見つかりません' };

  const data = sheet.getDataRange().getValues();

  // シートの行をループして、キーが一致する行を更新
  // newSettings は [{key:..., subject:..., body:..., attachments:...}, ...] の配列

  newSettings.forEach(setting => {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === setting.key) {
        // 更新 (行は i+1, 列は B=2, C=3, D=4)
        sheet.getRange(i + 1, 2).setValue(setting.subject);
        sheet.getRange(i + 1, 3).setValue(setting.body);
        sheet.getRange(i + 1, 4).setValue(setting.attachments);
        break; // 見つかったら次の設定へ
      }
    }
  });

  return { success: true, message: 'メール設定を保存しました' };
}