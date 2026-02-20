/**
 * extract_all_platforms.gs（マルチプラットフォーム版 / config.gs連携版）
 *
 * 前提:
 * - 同一プロジェクト内に config.gs を置き、getAppConfig_() が使えること
 * - config.gs の extractor 設定で prod/test を切り替える
 *
 * 仕様:
 * - Gmailから楽天トラベルキャンプ・なっぷの予約確定/キャンセルメールを抽出して該当シートへ追記
 * - 予約ID重複はスキップ（予約元+予約IDの組み合わせで判定）
 * - config.gsのSEARCH_PERIODで抽出期間を設定
 * - キャンセルメールは既存行のステータスを「キャンセル済み」に更新
 * - 二重抽出防止のため、gmailフォルダでラベルがつくので再抽出する場合はラベルを削除する必要あり
 * - チェックイン/アウトは Date 型でシートへ書き込み
 */

function extractAllPlatformEmails() {
  const CFG = getAppConfig_();

  const SHEET_ID = CFG.SHEET_ID;
  const PROCESSED_LABEL = CFG.extractor.PROCESSED_LABEL;
  const MAX_THREADS = CFG.extractor.MAX_THREADS;
  const ADD_LABEL = !!CFG.extractor.ADD_LABEL;
  const SEARCH_PERIOD = CFG.extractor.SEARCH_PERIOD || '7d';
  const PLATFORMS = CFG.extractor.PLATFORMS;

  const EXPECTED_HEADERS = [
    '予約日時', '予約ID', '予約元', 'チェックイン日時', 'チェックアウト日時', 'サイト名', 'サイト数',
    '大人', '子供', '幼児', '名前', '電話番号', 'メールアドレス', '備考', '料金',
    'ステータス'
  ];

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('他のプロセスが実行中のためスキップしました');
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // シートの取得と初期化
    const sheets = {};
    const cols = {};
    const totalCols = {};
    const idToRow = new Map(); // "platformObjKey:reservationId" -> rowNumber (1-based)

    ['rakuten', 'nap'].forEach(plat => {
      const sheetName = PLATFORMS[plat].SHEET_NAME;
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`シートが見つかりません: ${sheetName}`);

      sheets[plat] = sheet;
      const header = ensureHeaders_(sheet, EXPECTED_HEADERS);
      cols[plat] = Object.fromEntries(header.map((h, i) => [String(h).trim(), i + 1]));
      totalCols[plat] = sheet.getLastColumn();

      const all = sheet.getDataRange().getValues();
      for (let r = 1; r < all.length; r++) {
        const id = String(all[r][cols[plat]['予約ID'] - 1] || '').trim();
        if (id) {
          idToRow.set(`${plat}:${id}`, r + 1);
        }
      }
    });

    const label = GmailApp.getUserLabelByName(PROCESSED_LABEL) || GmailApp.createLabel(PROCESSED_LABEL);
    const labeledIds = getLabeledThreadIdSet_(label);

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

    Logger.log(`検索結果スレッド数: ${threads.length}`);

    const canceledIds = new Set();
    const confirmMap = new Map();
    const threadsToLabel = new Set();

    threads.forEach(th => {
      const isLabeled = labeledIds.has(th.getId());

      th.getMessages().forEach(msg => {
        const subject = msg.getSubject() || '';
        const body = msg.getPlainBody() || msg.getBody() || '';

        const platform = detectPlatform_(msg, PLATFORMS);
        if (!platform) {
          // Logger.log(`プラットフォーム未検出: From=${msg.getFrom()}, Subject=${subject}`);
          return;
        }

        const platformName = platform === 'rakuten' ? '楽天トラベル' : 'なっぷ';

        let reservationId = '';
        if (platform === 'rakuten') {
          reservationId = pick_(body, /(?:予約ID|予約ＩＤ)[^\S\r\n]*[:：]?\s*([A-Z0-9-]+)/i);
        } else if (platform === 'nap') {
          reservationId = pick_(body, /予約詳細番号[^:：]*[:：]\s*([A-Z0-9-]+)/);
          if (!reservationId) {
            const subjMatch = subject.match(/([A-Z0-9]+-\d+)/);
            reservationId = subjMatch ? subjMatch[1] : '';
          }
        }

        if (!reservationId) {
          Logger.log(`予約ID抽出失敗 [${platformName}]: ${subject}`);
          return;
        }

        const idKey = `${platform}:${reservationId}`;

        const isCancel = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CANCEL_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CANCEL_SUBJECT));
        const isConfirm = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CONFIRM_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CONFIRM_SUBJECT));

        if (isLabeled && !isCancel) return;

        threadsToLabel.add(th.getId());

        if (isCancel) {
          Logger.log(`キャンセルメール検出: ${idKey} (${subject})`);
          canceledIds.add(idKey);
          return;
        }

        if (isConfirm) {
          Logger.log(`確定メール検出: ${idKey} (${subject})`);
          const msgDate = msg.getDate();
          const existing = confirmMap.get(idKey);
          if (existing && existing.msgDate && existing.msgDate.getTime() >= msgDate.getTime()) return;

          const extracted = platform === 'rakuten'
            ? extractReservationDataFromBody_(body)
            : extractNapReservationData_(body);

          if (!extracted || !isValidDate_(extracted.checkInDate) || !extracted.name) {
            Logger.log(`抽出不十分 [${idKey}]: Name=${extracted ? extracted.name : 'N/A'}, CheckIn=${extracted ? extracted.checkInDate : 'N/A'}`);
            return;
          }

          confirmMap.set(idKey, {
            platformKey: platform,
            platformName: platformName,
            reservationId,
            msgDate: extracted.msgDate || msgDate,
            thread: th,
            ...extracted
          });
        }
      });
    });

    // 2) キャンセルを先に反映
    canceledIds.forEach(idKey => {
      const rowNum = idToRow.get(idKey);
      if (rowNum) {
        const plat = idKey.split(':')[0];
        const statusCol = cols[plat]['ステータス'];
        if (statusCol) {
          sheets[plat].getRange(rowNum, statusCol).setValue('キャンセル済み');
        }
      }
    });

    // 1) 確定データを挿入（プラットフォームごとに振り分け）
    const rowsToInsert = { rakuten: [], nap: [] };

    confirmMap.forEach(info => {
      const plat = info.platformKey;
      const idKey = `${plat}:${info.reservationId}`;

      if (idToRow.has(idKey)) return;

      const status = canceledIds.has(idKey) ? 'キャンセル済み' : '予約中';
      const c = cols[plat];

      const row = new Array(totalCols[plat]).fill('');
      if (c['予約日時']) row[c['予約日時'] - 1] = info.msgDate;
      if (c['予約ID']) row[c['予約ID'] - 1] = info.reservationId;
      if (c['予約元']) row[c['予約元'] - 1] = info.platformName;
      if (c['チェックイン日時']) row[c['チェックイン日時'] - 1] = info.checkInDate;
      if (c['チェックアウト日時']) row[c['チェックアウト日時'] - 1] = info.checkOutDate || '';
      if (c['サイト名']) row[c['サイト名'] - 1] = info.siteName || '';
      if (c['サイト数']) row[c['サイト数'] - 1] = info.siteCount || '';
      if (c['大人']) row[c['大人'] - 1] = info.adult;
      if (c['子供']) row[c['子供'] - 1] = info.child;
      if (c['幼児']) row[c['幼児'] - 1] = info.infant;
      if (c['名前']) row[c['名前'] - 1] = info.name.trim();
      if (c['電話番号']) row[c['電話番号'] - 1] = info.phone ? '"' + String(info.phone).trim() + '"' : '';
      if (c['メールアドレス']) row[c['メールアドレス'] - 1] = info.email ? String(info.email).trim() : '';
      if (c['備考']) row[c['備考'] - 1] = info.remarks || '';
      if (c['料金']) row[c['料金'] - 1] = info.totalPrice || '';
      if (c['ステータス']) row[c['ステータス'] - 1] = status;

      rowsToInsert[plat].push({ thread: info.thread, row });
    });

    ['rakuten', 'nap'].forEach(plat => {
      if (rowsToInsert[plat].length) {
        const sheet = sheets[plat];
        const rows = rowsToInsert[plat];
        sheet.insertRowsBefore(2, rows.length);
        const valuesToWrite = rows.slice().reverse().map(i => i.row);
        sheet.getRange(2, 1, valuesToWrite.length, totalCols[plat]).setValues(valuesToWrite);
      }

      // 3) チェックイン完了更新
      updateCheckinCompleted_(sheets[plat], cols[plat]);

      // ソート
      if (sheets[plat].getLastRow() > 1 && cols[plat]['予約日時']) {
        sheets[plat].getRange(2, 1, sheets[plat].getLastRow() - 1, totalCols[plat]).sort({ column: cols[plat]['予約日時'], ascending: false });
      }
    });

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