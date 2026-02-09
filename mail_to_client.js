/**
 * mail_to_client.gs (統合版)
 * 前提:
 * - 同一プロジェクト内に config.gs を置き、getAppConfig_() が使えること
 * - config.gs の APP_MODE = 'prod' | 'test' を Script Properties で切り替え推奨
 * * 仕様:
 * ① 受信翌日：南京錠番号のお知らせ (テンプレート: reserve_next_day)
 * ② 宿泊日前日：注意事項リマインド (テンプレート: checkin_prev_day)
 * * 機能:
 * - スプレッドシート「メール設定」から件名・本文・添付ファイルを読み込み
 * - 本文内の {タグ} を自動置換
 * - テストモード時の誤送信防止機能つき
 * 送信条件:
 * - ステータス == 「予約中」 の行のみ対象
 * - ① 予約日時が「昨日」の行で、南京錠送信済み != 送信済み のものだけ送る
 * - ② チェックイン日前日が「今日」の行で、前日案内送信済み != 送信済み のものだけ送る
 * 安全装置:
 * - test モード時は、CONFIG TEST_FORCE_TOで宛先を必ず 919saigo@gmail.com に差し替えている（誤爆防止）
 * - test モード時は、config.gsのDRY_RUNがtrueになっていると送信されない
 */

function sendAllReminders() {
  const CFG = getAppConfig_();

  const CONFIG = {
    MODE: CFG.mode, 
    SHEET_ID: CFG.SHEET_ID,
    SHEET_NAME: CFG.mailer.SHEET_NAME,
    SETTING_SHEET_NAME: CFG.mailer.SETTING_SHEET_NAME || 'メール設定',
    JST: CFG.JST,
    DRY_RUN: !!CFG.mailer.DRY_RUN,
    LOCK_CODE: CFG.mailer.LOCK_CODE,
    FROM_NAME: CFG.mailer.FROM_NAME,
    REPLY_TO: CFG.mailer.REPLY_TO,
    ADMIN_SIGNATURE: CFG.mailer.ADMIN_SIGNATURE, // デフォルト値
    NOW_OVERRIDE: CFG.mailer.NOW_OVERRIDE || '',
    TEST_FORCE_TO: '919saigo@gmail.com'
  };

  const FLAG_COLS = {
    nextDay: '南京錠送信済み',
    dayBefore: '前日案内送信済み'
  };

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  const settingSheet = spreadsheet.getSheetByName(CONFIG.SETTING_SHEET_NAME);

  if (!sheet) throw new Error('メインシートが見つかりません: ' + CONFIG.SHEET_NAME);
  if (!settingSheet) throw new Error('設定シートが見つかりません: ' + CONFIG.SETTING_SHEET_NAME);

  // ★ テンプレート読み込み
  const templates = getMailTemplates_(settingSheet);

  // ★★★ 追加変更点：シートに署名設定があれば、CONFIGの署名を上書きする ★★★
  if (templates['common_signature'] && templates['common_signature'].body) {
    CONFIG.ADMIN_SIGNATURE = templates['common_signature'].body;
  }

  // 必要ヘッダー
  const REQUIRED_HEADERS = [
    '予約日時', '予約ID', 'チェックイン日時', 'チェックアウト日時', '名前', 'メールアドレス', 'ステータス', 
    'サイト名', '料金', '備考',
    FLAG_COLS.nextDay, FLAG_COLS.dayBefore
  ];

  ensureHeaders_(sheet, REQUIRED_HEADERS);

  const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = indexHeaders_(headerValues, REQUIRED_HEADERS);
  const fullColMap = {};
  headerValues.forEach((h, i) => { fullColMap[String(h).trim()] = i; });

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    Logger.log('データ行がありません。');
    return;
  }

  const now = CONFIG.NOW_OVERRIDE ? new Date(CONFIG.NOW_OVERRIDE) : new Date();
  const todayJST0 = toJST0_(now);
  const yesterdayJST0 = addDays_(todayJST0, -1);

  const cnt = {
    mode: CONFIG.MODE,
    dryRun: CONFIG.DRY_RUN,
    totalRows: values.length - 1,
    noEmail: 0,
    skippedNotActive: 0,
    nextDay: { ok: 0, notYesterday: 0, alreadySent: 0, invalidReceived: 0 },
    dayBefore: { ok: 0, invalidCheckin: 0, notDayBefore: 0, alreadySent: 0 }
  };

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const received = row[col['予約日時']];
    const checkin  = row[col['チェックイン日時']];
    const emailRaw = safeStr(row[col['メールアドレス']]);
    const status   = safeStr(row[col['ステータス']]);

    if (status !== '予約中') {
      cnt.skippedNotActive++;
      continue;
    }

    if (!emailRaw) {
      cnt.noEmail++;
      continue;
    }

    const to = (CONFIG.MODE === 'test') ? CONFIG.TEST_FORCE_TO : emailRaw;
    const nextDayFlag = safeStr(row[col[FLAG_COLS.nextDay]]);
    const dayBeforeFlag = safeStr(row[col[FLAG_COLS.dayBefore]]);

    // ① 受信翌日送信 (reserve_next_day)
    if (isValidDate_(received)) {
      if (isSameDay_(toJST0_(received), yesterdayJST0)) {
        if (nextDayFlag === '送信済み') {
          cnt.nextDay.alreadySent++;
        } else {
          const tmpl = templates['reserve_next_day'];
          if (tmpl) {
            cnt.nextDay.ok++;
            const sent = sendEmail_(to, tmpl, row, fullColMap, CONFIG);
            if (sent) sheet.getRange(r + 1, col[FLAG_COLS.nextDay] + 1).setValue('送信済み');
          } else {
            Logger.log('[Error] reserve_next_day テンプレート不足');
          }
        }
      } else {
        cnt.nextDay.notYesterday++;
      }
    } else {
      cnt.nextDay.invalidReceived++;
    }

    // ② 宿泊日前日送信 (checkin_prev_day)
    if (isValidDate_(checkin)) {
      const dayBeforeJST0 = addDays_(toJST0_(checkin), -1);
      if (isSameDay_(dayBeforeJST0, todayJST0)) {
        if (dayBeforeFlag === '送信済み') {
          cnt.dayBefore.alreadySent++;
        } else {
          const tmpl = templates['checkin_prev_day'];
          if (tmpl) {
            cnt.dayBefore.ok++;
            const sent = sendEmail_(to, tmpl, row, fullColMap, CONFIG);
            if (sent) sheet.getRange(r + 1, col[FLAG_COLS.dayBefore] + 1).setValue('送信済み');
          } else {
            Logger.log('[Error] checkin_prev_day テンプレート不足');
          }
        }
      } else {
        cnt.dayBefore.notDayBefore++;
      }
    } else {
      cnt.dayBefore.invalidCheckin++;
    }
  }

  Logger.log('処理統計: ' + JSON.stringify(cnt, null, 2));
}

function sendEmail_(to, template, rowData, colMap, config) {
  if (!template.subject || !template.body) return false;

  const subject = replaceTags_(template.subject, rowData, colMap, config);
  let body = replaceTags_(template.body, rowData, colMap, config);

  // 署名の追加
  body += '\n\n' + '--------------------------------------------------\n' + config.ADMIN_SIGNATURE;

  const blobs = [];
  if (template.attachmentIds && template.attachmentIds.length > 0) {
    template.attachmentIds.forEach(id => {
      try {
        const file = DriveApp.getFileById(id.trim());
        blobs.push(file.getBlob());
      } catch (e) {
        Logger.log(`添付エラー ID:${id} - ${e.message}`);
      }
    });
  }

  const options = {
    name: config.FROM_NAME,
    replyTo: config.REPLY_TO,
    attachments: blobs
  };

  if (config.DRY_RUN) {
    Logger.log(`[DRY_RUN] To: ${to}\nSubject: ${subject}\nAttachments: ${blobs.length}個`);
    return true; 
  } else {
    try {
      GmailApp.sendEmail(to, subject, body, options);
      Logger.log(`送信成功: ${to} (${subject})`);
      return true;
    } catch (e) {
      Logger.log(`送信失敗: ${to} - ${e.message}`);
      return false;
    }
  }
}

function getMailTemplates_(sheet) {
  const data = sheet.getDataRange().getValues();
  const templates = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const subject = String(data[i][1]).trim();
    const body = String(data[i][2]).trim();
    const attachIdsRaw = String(data[i][3]).trim();

    if (key) {
      templates[key] = {
        subject: subject,
        body: body,
        attachmentIds: attachIdsRaw ? attachIdsRaw.split(',') : []
      };
    }
  }
  return templates;
}

function replaceTags_(text, rowData, colMap, config) {
  let res = text;
  const getVal = (name) => {
    const idx = colMap[name];
    return (idx !== undefined) ? rowData[idx] : '';
  };

  const map = {
    '{名前}': getVal('名前'),
    '{予約ID}': getVal('予約ID'),
    '{チェックイン日}': formatDateJP_(getVal('チェックイン日時')),
    '{チェックアウト日}': formatDateJP_(getVal('チェックアウト日時')),
    '{サイト名}': getVal('サイト名'),
    '{料金}': getVal('料金'),
    '{備考}': getVal('備考'),
    '{南京錠}': config.LOCK_CODE
  };

  for (const [key, val] of Object.entries(map)) {
    res = res.split(key).join(val || '');
  }
  return res;
}

function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return headers;
  }
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const missing = headers.filter(h => !current.includes(h));
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    return current.concat(missing);
  }
  return current;
}

function indexHeaders_(headerRow, names) {
  const idx = {};
  names.forEach(n => {
    const i = headerRow.indexOf(n);
    if (i === -1) Logger.log(`注意: ヘッダー ${n} が見つかりません`);
    idx[n] = i; 
  });
  return idx;
}

function safeStr(v) { return v == null ? '' : String(v).trim(); }
function isValidDate_(d) { return d instanceof Date && !isNaN(d.getTime()); }
function toJST0_(d) {
  const ymd = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return new Date(ymd + 'T00:00:00+09:00');
}
function addDays_(d, delta) { return new Date(d.getTime() + delta * 24 * 60 * 60 * 1000); }
function isSameDay_(d1, d2) { return d1.getTime() === d2.getTime(); }
function formatDateJP_(d) {
  if (!(d instanceof Date)) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}