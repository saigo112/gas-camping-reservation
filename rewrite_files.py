import re

# 1. gmail_to_sheets.js
with open('gmail_to_sheets.js', 'r') as f:
    text = f.read()

new_extract = """function extractAllPlatformEmails() {
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

    const sheets = {};
    const cols = {};
    const totalCols = {};
    const idToRow = new Map();
    
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

    const canceledIds = new Set();
    const confirmMap = new Map();
    const threadsToLabel = new Set();

    threads.forEach(th => {
      const isLabeled = labeledIds.has(th.getId());

      th.getMessages().forEach(msg => {
        const subject = msg.getSubject() || '';
        const body = msg.getPlainBody() || msg.getBody() || '';

        const platform = detectPlatform_(msg, PLATFORMS);
        if (!platform) return;

        const platformName = platform === 'rakuten' ? '楽天トラベル' : 'なっぷ';

        let reservationId = '';
        if (platform === 'rakuten') {
          reservationId = pick_(body, /(?:予約ID|予約ＩＤ)[^\\S\\r\\n]*[:：]?\\s*([A-Z0-9-]+)/i);
        } else if (platform === 'nap') {
          reservationId = pick_(body, /予約詳細番号[^:：]*[:：]\\s*([A-Z0-9-]+)/);
          if (!reservationId) {
            const subjMatch = subject.match(/([A-Z0-9]+-\\d+)/);
            reservationId = subjMatch ? subjMatch[1] : '';
          }
        }

        if (!reservationId) return;

        const idKey = `${platform}:${reservationId}`;

        const isCancel = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CANCEL_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CANCEL_SUBJECT));
        const isConfirm = (platform === 'rakuten' && subject.includes(PLATFORMS.rakuten.CONFIRM_SUBJECT)) ||
          (platform === 'nap' && subject.includes(PLATFORMS.nap.CONFIRM_SUBJECT));

        if (isLabeled && !isCancel) return;

        threadsToLabel.add(th.getId());

        if (isCancel) {
          canceledIds.add(idKey);
          return;
        }

        if (isConfirm) {
          const msgDate = msg.getDate();
          const existing = confirmMap.get(idKey);
          if (existing && existing.msgDate && existing.msgDate.getTime() >= msgDate.getTime()) return;

          const extracted = platform === 'rakuten'
            ? extractReservationDataFromBody_(body)
            : extractNapReservationData_(body);

          if (!extracted || !isValidDate_(extracted.checkInDate) || !extracted.name) return;

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

    if (cols['rakuten']['ステータス']) {
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
    }

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

        if (cols[plat]['予約ID']) sheet.getRange(2, cols[plat]['予約ID'], sheet.getLastRow() - 1, 1).setNumberFormat('@');
        if (cols[plat]['電話番号']) sheet.getRange(2, cols[plat]['電話番号'], sheet.getLastRow() - 1, 1).setNumberFormat('@');
      }
      
      updateCheckinCompleted_(sheets[plat], cols[plat]);

      if (sheets[plat].getLastRow() > 1 && cols[plat]['予約日時']) {
        sheets[plat].getRange(2, 1, sheets[plat].getLastRow() - 1, totalCols[plat]).sort({ column: cols[plat]['予約日時'], ascending: false });
      }
    });

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
"""
text = re.sub(r'function extractAllPlatformEmails\(\) \{.*?(?=\/\* ===)', new_extract, text, flags=re.DOTALL)
with open('gmail_to_sheets.js', 'w') as f:
    f.write(text)

# 2. webapp.js
with open('webapp.js', 'r') as f:
    text = f.read()

new_get_data = """function getDataForWeb() {
  const CFG = getAppConfig_();
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  
  let allDisplayData = [];

  ['rakuten', 'nap'].forEach(plat => {
    const sheetName = CFG.extractor.PLATFORMS[plat].SHEET_NAME;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    const displayData = data.map(row => {
      const startObj = row[colMap['チェックイン日時']];
      const endObj = row[colMap['チェックアウト日時']];

      return {
        id: row[colMap['予約ID']],
        platform: row[colMap['予約元']] || (plat === 'rakuten' ? '楽天トラベル' : 'なっぷ'),
        status: row[colMap['ステータス']],
        checkIn: formatDate_(startObj),
        checkOut: formatDate_(endObj),
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
    
    allDisplayData = allDisplayData.concat(displayData);
  });

  allDisplayData.sort((a, b) => {
    if (!a.startIso) return 1;
    if (!b.startIso) return -1;
    return new Date(a.startIso) - new Date(b.startIso);
  });

  return allDisplayData;
}"""

new_update_status = """function updateReservationStatus(reservationId, newStatus) {
  const CFG = getAppConfig_();
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);

  for (const plat of ['rakuten', 'nap']) {
    const sheetName = CFG.extractor.PLATFORMS[plat].SHEET_NAME;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    if (!headers) continue;
    
    const colMap = {};
    headers.forEach((h, i) => { colMap[String(h).trim()] = i; });

    const idColIndex = colMap['予約ID'];
    const statusColIndex = colMap['ステータス'];

    if (idColIndex === undefined || statusColIndex === undefined) continue;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(reservationId)) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
        if (typeof syncCalendarEvents === 'function') {
          syncCalendarEvents();
        }
        return { success: true, message: `ID:${reservationId} を「${newStatus}」に更新しました` };
      }
    }
  }

  return { success: false, message: 'IDが見つかりませんでした' };
}"""

text = re.sub(r'function getDataForWeb\(\) \{.*?(?=\n// 日付フォーマット)', new_get_data, text, flags=re.DOTALL)
text = re.sub(r'function updateReservationStatus\(.*?\).*?(?=\n// 4)', new_update_status, text, flags=re.DOTALL)
with open('webapp.js', 'w') as f:
    f.write(text)

# 3. sync_calendar.js
with open('sync_calendar.js', 'r') as f:
    text = f.read()

new_sync_cal = """function syncCalendarEvents() {
  const CFG = getAppConfig_();
  const SHEET_ID = CFG.SHEET_ID;
  const CAL_ID = CFG.CALENDAR_ID;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  try {
    const calendar = CalendarApp.getCalendarById(CAL_ID);
    if (!calendar) throw new Error(`カレンダーが見つかりません ID: ${CAL_ID}`);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const now = new Date();

    ['rakuten', 'nap'].forEach(plat => {
      const sheetName = CFG.extractor.PLATFORMS[plat].SHEET_NAME;
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colMap = {};
      headers.forEach((h, i) => { colMap[String(h).trim()] = i + 1; });

      const required = ['予約ID', 'ステータス', 'チェックイン日時', 'チェックアウト日時', '名前', 'カレンダーイベントID'];
      if (!required.every(k => colMap[k])) {
        Logger.log(`${sheetName} に必要な列が見つかりません。`);
        return;
      }

      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      const data = dataRange.getValues();
      const updates = [];

      for (let i = 0; i < data.length; i++) {
        const rowNum = i + 2;
        const rowData = data[i];

        const rId = String(rowData[colMap['予約ID'] - 1]);
        const platform = String(rowData[colMap['予約元'] - 1] || (plat === 'rakuten' ? '楽天トラベル' : 'なっぷ'));
        const status = String(rowData[colMap['ステータス'] - 1]).trim();
        const eventId = String(rowData[colMap['カレンダーイベントID'] - 1]);
        const checkIn = rowData[colMap['チェックイン日時'] - 1];
        const checkOut = rowData[colMap['チェックアウト日時'] - 1];
        const name = String(rowData[colMap['名前'] - 1]);
        const siteName = String(rowData[colMap['サイト名'] - 1] || '');

        if (status === '予約中' && eventId === '' && isValidDate_(checkIn) && checkIn > now) {
          const title = `【${platform}】【予約ID:${rId}】${name}様 (${siteName})`;
          const desc = `予約元: ${platform}\\n予約ID: ${rId}\\nサイト: ${siteName}\\n名前: ${name}\\n自動連携により作成`;
          let endDt = isValidDate_(checkOut) ? checkOut : new Date(checkIn.getTime() + (60 * 60 * 1000));

          try {
            const event = calendar.createEvent(title, checkIn, endDt, { description: desc });
            updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: event.getId() });
            Logger.log(`カレンダー追加: ${title}`);
          } catch (e) {
            Logger.log(`カレンダー追加失敗 Row${rowNum}: ${e.message}`);
          }
        }
        else if (status === 'キャンセル済み' && eventId !== '') {
          try {
            const event = calendar.getEventById(eventId);
            if (event) {
              event.deleteEvent();
              Logger.log(`カレンダー削除: ID ${eventId}`);
            }
            updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: '' });
          } catch (e) {
            Logger.log(`カレンダー削除失敗 Row${rowNum}: ${e.message}`);
            updates.push({ row: rowNum, col: colMap['カレンダーイベントID'], value: '' });
          }
        }
      }

      if (updates.length > 0) {
        updates.forEach(u => {
          sheet.getRange(u.row, u.col).setValue(u.value);
        });
        SpreadsheetApp.flush();
      }
    });

  } catch (e) {
    Logger.log('Sync Error: ' + e.stack);
  } finally {
    lock.releaseLock();
  }
}"""
text = re.sub(r'function syncCalendarEvents\(\) \{.*?(?=\n// 日付判定用)', new_sync_cal, text, flags=re.DOTALL)
with open('sync_calendar.js', 'w') as f:
    f.write(text)

print("Done")
