/**
 * extract_all_platforms.gs（マルチプラットフォーム版 / config.gs連携版）
 *
 * 前提:
 * - 同一プロジェクト内に config.gs を置き、getAppConfig_() が使えること
 * - config.gs の extractor 設定で prod/test を切り替える
 *
 * 仕様:
 * - Gmailから楽天トラベルキャンプ・なっぷの予約確定/キャンセルメールを抽出してシートへ追記
 * - 予約ID重複はスキップ（予約元+予約IDの組み合わせで判定）
 * - config.gsのSEARCH_PERIODで抽出期間を設定
 * - キャンセルメールは既存行のステータスを「キャンセル済み」に更新
 * - 二重抽出防止のため、gmailフォルダでラベルがつくので再抽出する場合はラベルを削除する必要あり
 * - チェックイン/アウトは Date 型でシートへ書き込み
 * ステータス仕様:
 * 1) 件名「予約が確定しました」「ご予約ありがとうございます」 → 新規行追加、ステータス「予約中」
 * 2) 件名「予約がキャンセルされました」「キャンセル」 → 予約ID一致行のステータスを「キャンセル済み」に更新
 * 3) ステータス「予約中」かつ チェックイン日時 <= 現在 → 「チェックイン完了」に更新
 */

function extractAllPlatformEmails() {
  const CFG = getAppConfig_();

  const SHEET_ID = CFG.SHEET_ID;
  const SHEET_NAME = CFG.extractor.SHEET_NAME;
  const PROCESSED_LABEL = CFG.extractor.PROCESSED_LABEL;
  const MAX_THREADS = CFG.extractor.MAX_THREADS;
  const ADD_LABEL = !!CFG.extractor.ADD_LABEL;
  // ★設定ファイルから期間を取得（デフォルトは7d）
  const SEARCH_PERIOD = CFG.extractor.SEARCH_PERIOD || '7d';
  const PLATFORMS = CFG.extractor.PLATFORMS;

  const EXPECTED_HEADERS = [
    '予約日時', '予約ID', '予約元', 'チェックイン日時', 'チェックアウト日時', 'サイト名', 'サイト数',
    '大人', '子供', '幼児', '名前', '電話番号', 'メールアドレス', '備考', '料金',
    'ステータス'
  ];

  const lock = LockService.getScriptLock();
  // 30秒待機
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('他のプロセスが実行中のためスキップしました');
    return;
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

    // ヘッダー整備
    const header = ensureHeaders_(sheet, EXPECTED_HEADERS);
    const col = Object.fromEntries(header.map((h, i) => [String(h).trim(), i + 1])); // 1-based
    const totalCols = sheet.getLastColumn();

    // Gmailラベル
    const label = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
    const labeledIds = getLabeledThreadIdSet_(label);

    // シート現状を一度だけ読み、「予約元+予約ID」→行番号（1-based）のMapを作る
    const all = sheet.getDataRange().getValues();
    const idToRow = new Map(); // "platform:reservationId" -> rowNumber (1-based)
    for (let r = 1; r < all.length; r++) {
      const platform = String(all[r][col['予約元'] - 1] || '').trim();
      const id = String(all[r][col['予約ID'] - 1] || '').trim();
      if (id) {
        const key = platform ? `${platform}:${id}` : id; // 予約元がある場合は組み合わせ
        idToRow.set(key, r + 1);
      }
    }

    // Gmail検索（両プラットフォーム対応）
    const query = [
      '(',
      `from:${PLATFORMS.rakuten.FROM}`,
      'OR',
      `from:${PLATFORMS.nap.FROM}`,
      ')',
      `newer_than:${SEARCH_PERIOD}`
    ].join(' ');

    Logger.log(`検索クエリ: ${query}`);
    const threads = GmailApp.search(query, 0, MAX_THREADS);

    // ===== 実行内整合用バッファ =====
    const canceledIds = new Set();     // 「platform:予約ID」のキャンセルセット
    const confirmMap = new Map();      // 「platform:予約ID」 -> extracted info（確定メールの抽出結果）
    const threadsToLabel = new Set();  // 今回処理対象になったスレッドID

    // まず全件スキャン（書き込みしない）
    threads.forEach(th => {
      const isLabeled = labeledIds.has(th.getId());

      th.getMessages().forEach(msg => {
        const subject = msg.getSubject() || '';
        const body = msg.getPlainBody() || msg.getBody() || '';

        // プラットフォーム検出
        const platform = detectPlatform_(msg, PLATFORMS);
        if (!platform) return; // 未対応プラットフォーム

        // プラットフォーム名の日本語表記
        const platformName = platform === 'rakuten' ? '楽天トラベル' : 'なっぷ';

        // 予約IDの抽出（プラットフォームごとに異なる）
        let reservationId = '';
        if (platform === 'rakuten') {
          reservationId = pick_(body, /(?:予約ID|予約ＩＤ)[^\S\r\n]*[:：]?\s*([A-Z0-9-]+)/i);
        } else if (platform === 'nap') {
          // なっぷは「予約詳細番号」または件名から取得
          reservationId = pick_(body, /予約詳細番号[^:：]*[:：]\s*([A-Z0-9-]+)/);
          if (!reservationId) {
            const subjMatch = subject.match(/([A-Z0-9]+-\d+)/);
            reservationId = subjMatch ? subjMatch[1] : '';
          }
        }

        if (!reservationId) return;

        // 「プラットフォーム:予約ID」の組み合わせをキーとする
        const idKey = `${platformName}:${reservationId}`;

        // ラベル済みスレッドは「キャンセル検出だけ」で十分
        const isCancel = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CANCEL_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CANCEL_SUBJECT));
        const isConfirm = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CONFIRM_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CONFIRM_SUBJECT));

        if (isLabeled && !isCancel) return;

        threadsToLabel.add(th.getId());

        // キャンセルメール処理
        if (isCancel) {
          canceledIds.add(idKey);
          return;
        }

        // 確定メール処理
        if (isConfirm) {
          const msgDate = msg.getDate();
          const existing = confirmMap.get(idKey);
          if (existing && existing.msgDate && existing.msgDate.getTime() >= msgDate.getTime()) return;

          // プラットフォームごとに適切な抽出関数を使用
          const extracted = platform === 'rakuten'
            ? extractReservationDataFromBody_(body)
            : extractNapReservationData_(body);

          if (!extracted || !isValidDate_(extracted.checkInDate) || !extracted.name) return;

          confirmMap.set(idKey, {
            platform: platformName,
            reservationId,
            msgDate: extracted.msgDate || msgDate,
            thread: th,
            ...extracted
          });
        }
      });
    });

    // ===== 2) キャンセルを先に反映（既存行があるなら即キャンセル済み） =====
    if (col['ステータス']) {
      canceledIds.forEach(id => {
        const rowNum = idToRow.get(id);
        if (rowNum) {
          sheet.getRange(rowNum, col['ステータス']).setValue('キャンセル済み');
        }
      });
    }

    // ===== 1) 確定データを挿入 =====
    const rowsToInsert = [];

    confirmMap.forEach(info => {
      const id = info.reservationId;
      const platform = info.platform;
      const idKey = `${platform}:${id}`;

      // `idKey`でチェック (platform+ID の組み合わせ)
      if (idToRow.has(idKey)) return;

      const status = canceledIds.has(idKey) ? 'キャンセル済み' : '予約中';

      const row = new Array(totalCols).fill('');
      if (col['予約日時']) row[col['予約日時'] - 1] = info.msgDate;
      if (col['予約ID']) row[col['予約ID'] - 1] = id;
      if (col['予約元']) row[col['予約元'] - 1] = platform;
      if (col['チェックイン日時']) row[col['チェックイン日時'] - 1] = info.checkInDate;
      if (col['チェックアウト日時']) row[col['チェックアウト日時'] - 1] = info.checkOutDate || '';
      if (col['サイト名']) row[col['サイト名'] - 1] = info.siteName || '';
      if (col['サイト数']) row[col['サイト数'] - 1] = info.siteCount || '';
      if (col['大人']) row[col['大人'] - 1] = info.adult;
      if (col['子供']) row[col['子供'] - 1] = info.child;
      if (col['幼児']) row[col['幼児'] - 1] = info.infant;
      if (col['名前']) row[col['名前'] - 1] = info.name.trim();
      // ★電話番号の修正: ""を付けて0落ちを防止
      if (col['電話番号']) row[col['電話番号'] - 1] = info.phone ? '"' + String(info.phone).trim() + '"' : '';
      if (col['メールアドレス']) row[col['メールアドレス'] - 1] = info.email ? String(info.email).trim() : '';
      if (col['備考']) row[col['備考'] - 1] = info.remarks || '';
      if (col['料金']) row[col['料金'] - 1] = info.totalPrice || '';
      if (col['ステータス']) row[col['ステータス'] - 1] = status;

      rowsToInsert.push({ thread: info.thread, row });
    });

    if (rowsToInsert.length) {
      sheet.insertRowsBefore(2, rowsToInsert.length);
      const valuesToWrite = rowsToInsert.slice().reverse().map(i => i.row);
      sheet.getRange(2, 1, valuesToWrite.length, totalCols).setValues(valuesToWrite);

      // 予約IDはテキストにしておく
      if (col['予約ID']) sheet.getRange(2, col['予約ID'], sheet.getLastRow() - 1, 1).setNumberFormat('@');

      // ★電話番号も明示的にテキスト形式にしておく（念のため）
      if (col['電話番号']) sheet.getRange(2, col['電話番号'], sheet.getLastRow() - 1, 1).setNumberFormat('@');
    }

    // ===== 3) チェックイン完了更新 =====
    updateCheckinCompleted_(sheet, col);

    // ===== ソート（予約日時 降順） =====
    if (sheet.getLastRow() > 1 && col['予約日時']) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, totalCols).sort({ column: col['予約日時'], ascending: false });
    }

    // ===== ラベル付与 =====
    if (ADD_LABEL) {
      threads.forEach(th => {
        if (!threadsToLabel.has(th.getId())) return;
        try { th.addLabel(label); } catch (e) { Logger.log('ラベル付与失敗: ' + e); }
      });
    }

  } finally {
    lock.releaseLock();
  }
}

/* ================== プラットフォーム検出 & 抽出ロジック ================== */

/**
 * メールからプラットフォームを検出する
 * @param {GmailMessage} msg - Gmailメッセージオブジェクト
 * @param {Object} platforms - プラットフォーム設定オブジェクト
 * @return {string|null} - 'rakuten' または 'nap'、検出できない場合はnull
 */
function detectPlatform_(msg, platforms) {
  const from = msg.getFrom() || '';
  const subject = msg.getSubject() || '';

  // 楽天トラベルの検出
  if (from.includes(platforms.rakuten.FROM)) {
    return 'rakuten';
  }

  // なっぷの検出
  if (from.includes(platforms.nap.FROM)) {
    return 'nap';
  }

  return null;
}

/**
 * なっぷのメール本文から予約データを抽出
 * @param {string} body - メール本文
 * @return {Object} - 抽出された予約データ
 */
function extractNapReservationData_(body) {
  // 予約日時: 予約日時　　　　　　 : 2026年02月08日(日) 23時19分
  const reservedDateMatch = body.match(/予約日時[^:：]*[:：]\s*(\d{4})年(\d{2})月(\d{2})日[^0-9]*(\d{1,2})時(\d{2})分/);
  const msgDate = reservedDateMatch
    ? new Date(+reservedDateMatch[1], +reservedDateMatch[2] - 1, +reservedDateMatch[3], +reservedDateMatch[4], +reservedDateMatch[5], 0, 0)
    : new Date();

  // チェックイン日時: 2026年02月14日(土) 13時00分
  const checkinMatch = body.match(/チェックイン日時[^:：]*[:：]\s*(\d{4})年(\d{2})月(\d{2})日[^0-9]*(\d{1,2})時(\d{2})分/);
  const checkInDate = checkinMatch
    ? new Date(+checkinMatch[1], +checkinMatch[2] - 1, +checkinMatch[3], +checkinMatch[4], +checkinMatch[5], 0, 0)
    : null;

  // チェックアウト日時: 2026年02月15日(日) 10時00分
  const checkoutMatch = body.match(/チェックアウト日時[^:：]*[:：]\s*(\d{4})年(\d{2})月(\d{2})日[^0-9]*(\d{1,2})時(\d{2})分/);
  const checkOutDate = checkoutMatch
    ? new Date(+checkoutMatch[1], +checkoutMatch[2] - 1, +checkoutMatch[3], +checkoutMatch[4], +checkoutMatch[5], 0, 0)
    : null;

  // 予約施設（サイト名）
  const siteName = pick_(body, /予約施設[^:：]*[:：]\s*([^\n]+)/);

  // サイト数（なっぷには無いことが多いので空の場合もあり）
  const siteCount = '';

  // 人数: 人数　　　　　　　　 : 大人 1人
  const adultMatch = pick_(body, /人数[^\n]*大人\s*(\d+)\s*人/);
  const childMatch = pick_(body, /人数[^\n]*子供\s*(\d+)\s*人/);
  const infantMatch = pick_(body, /人数[^\n]*幼児\s*(\d+)/);
  const adult = toNumOrZero_(adultMatch);
  const child = toNumOrZero_(childMatch);
  const infant = toNumOrZero_(infantMatch);

  // 代表者氏名
  const name = pick_(body, /代表者氏名[^:：]*[:：]\s*([^\n]+)/);

  // 代表者連絡先（電話番号）
  const phone = pick_(body, /代表者連絡先[^:：]*[:：]\s*(0\d{9,10})/);

  // お客様のメールアドレス
  const email = pick_(body, /お客様のメールアドレス[^:：]*[:：]\s*\[?([^\]\s\n]+@[^\]\s\n]+)\]?/);

  // ご要望
  const remarks = pick_(body, /■\s*ご要望\s*\n\s*([^\n]+)/);

  // 利用料総額: 利用料総額　　　　　 :   1,600円
  const totalPrice = pick_(body, /利用料総額[^:：]*[:：]\s*([￥]?\s*[\d,]+円?)/);

  return {
    msgDate,
    checkInDate,
    checkOutDate,
    siteName,
    siteCount,
    adult,
    child,
    infant,
    name,
    phone,
    email,
    remarks,
    totalPrice
  };
}

/**
 * 楽天トラベルのメール本文から予約データを抽出（既存関数）
 */
function extractReservationDataFromBody_(body) {
  const periodMatch = body.match(/[▼\s]*(宿泊期間|利用日)[^\S\r\n]*[:：]?\s*([0-9\/.\-年月日 　:\n～~]+)/i);
  let checkInDate = null;
  let checkOutDate = null;

  if (periodMatch) {
    const kind = String(periodMatch[1] || '').trim();
    const raw = String(periodMatch[2] || '').replace(/\r/g, '').trim();
    const parts = raw.split(/[～~]/).map(p => p.replace(/\n/g, ' ').trim()).filter(Boolean);

    checkInDate = parseToDate_(parts[0]);

    if (kind === '利用日') {
      if (isValidDate_(checkInDate)) {
        checkOutDate = new Date(checkInDate.getTime());
        checkOutDate.setHours(18, 0, 0, 0);
      }
    } else if (kind === '宿泊期間') {
      if (isValidDate_(checkInDate) && parts[1]) {
        const endRaw = parts[1];
        const hasDate = /\d{4}\s*[\/年]/.test(endRaw) || /\d{1,2}\s*[\/月]\s*\d{1,2}/.test(endRaw);
        checkOutDate = hasDate ? parseToDate_(endRaw) : parseToDate_(endRaw, checkInDate);

        if (!hasDate && isValidDate_(checkOutDate) && checkOutDate.getTime() <= checkInDate.getTime()) {
          checkOutDate.setDate(checkOutDate.getDate() + 1);
        }
      }
    }
  }

  const siteName = pick_(body, /サイト名[^\S\r\n]*[:：]?\s*([^\n<]+)/i);
  const siteCount = pick_(body, /予約サイト数[^\S\r\n]*[:：]?\s*(\d+)/i);
  const ppl = body.match(/大人[^\d]*(\d+)\s*名.*?子供[^\d]*(\d+)\s*名.*?幼児[^\d]*(\d+)\s*名/si);
  const adult = toNumOrZero_(ppl ? ppl[1] : pick_(body, /大人[^\d]*(\d+)\s*名/));
  const child = toNumOrZero_(ppl ? ppl[2] : pick_(body, /子供[^\d]*(\d+)\s*名/));
  const infant = toNumOrZero_(ppl ? ppl[3] : pick_(body, /幼児[^\d]*(\d+)\s*名/));
  const name = pick_(body, /お名前[^\S\r\n]*[:：]\s*([^\n<]+)/i);
  const phone = pick_(body, /電話番号[^\S\r\n]*[:：]\s*(0\d{9,10})/);
  const email = pick_(body, /メールアドレス[^\S\r\n]*[:：]\s*([\w._%+\-]+@[\w.\-]+\.[A-Za-z]{2,})/);
  const bodyNorm = body.replace(/\r/g, '');
  const remarks = pick_(bodyNorm, /[▼\s]*備考[^\S\n\r]*[:：]?[^\S\n\r]*(.+?)(?=\n[^\S\n\r]*[▼\s]*(?:利用料金|お支払い済み料金|サイト名|予約者情報|予約詳細|$)|\n{2,}|$)/is);
  const priceAll = body.match(/(利用料金|合計料金|お支払い済み料金)[^\d￥]*([￥]?\s*[\d,]+)/i);
  const totalPrice = priceAll ? String(priceAll[2]).replace(/\s+/g, '') : '';

  return { checkInDate, checkOutDate, siteName, siteCount, adult, child, infant, name, phone, email, remarks, totalPrice };
}

function updateCheckinCompleted_(sheet, col) {
  if (!col['ステータス'] || !col['チェックイン日時']) return;
  const now = new Date();
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const status = String(data[r][col['ステータス'] - 1] || '').trim();
    if (status !== '予約中') continue;
    const checkin = data[r][col['チェックイン日時'] - 1];
    if (checkin instanceof Date && !isNaN(checkin.getTime())) {
      if (checkin.getTime() <= now.getTime()) {
        sheet.getRange(r + 1, col['ステータス']).setValue('チェックイン完了');
      }
    }
  }
}

function isValidDate_(d) { return d instanceof Date && !isNaN(d.getTime()); }
function parseToDate_(str, baseDate) {
  if (!str) return null;
  const s = String(str).replace(/　/g, ' ').replace(/[年月]/g, '/').replace(/日/g, '').replace(/[.\-]/g, '/').replace(/\s+/g, ' ').trim();
  let m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], 0, 0);
  m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0);
  m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m && baseDate instanceof Date && !isNaN(baseDate.getTime())) {
    const d = new Date(baseDate.getTime());
    d.setHours(+m[1], +m[2], 0, 0);
    return d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function toNumOrZero_(v) {
  if (v === undefined || v === null || v === '') return 0;
  const num = Number(String(v).replace(/[^0-9]/g, ''));
  return isNaN(num) ? 0 : num;
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
function getLabeledThreadIdSet_(label) {
  const ids = new Set();
  for (let start = 0; ; start += 100) {
    const page = label.getThreads(start, 100);
    if (!page.length) break;
    page.forEach(t => ids.add(t.getId()));
  }
  return ids;
}
function pick_(s, re) {
  const m = s.match(re);
  return m ? String(m[1] || m[2] || '').trim() : '';
}

/**
 * ▼▼▼ メイン実行関数 ▼▼▼
 * トリガー設定はこれを「1時間おき」等に設定してください。
 * 1. Gmailから抽出・更新
 * 2. Googleカレンダーへ同期
 */
function mainSequence() {
  extractAllPlatformEmails();
  if (typeof syncCalendarEvents === 'function') {
    syncCalendarEvents();
  } else {
    Logger.log('syncCalendarEvents 関数が見つかりません。sync_calendar.gs があるか確認してください。');
  }
}