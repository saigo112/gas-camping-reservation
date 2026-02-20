
function debugLineProps() {
  const sp = PropertiesService.getScriptProperties();
  const all = sp.getProperties();
  Logger.log('All Script Properties:\n' + JSON.stringify(all, null, 2));
  Logger.log('Has CHANNEL_TOKEN? ' + Boolean(sp.getProperty('LINE_CHANNEL_ACCESS_TOKEN')));
  Logger.log('Has ADMIN_USER_ID? ' + Boolean(sp.getProperty('LINE_ADMIN_USER_ID')));
  Logger.log('Has NOTIFY_TOKEN?  ' + Boolean(sp.getProperty('LINE_NOTIFY_ACCESS_TOKEN')));
  Logger.log('Script ID: ' + ScriptApp.getScriptId());
}

function setAdminUserIdAndVerify() {
  const KEY = 'LINE_ADMIN_USER_ID';
  const VALUE = 'U3416fc5085d25c254fe6164ff3e20cb6'; // ← 実際の U から始まるID

  const sp = PropertiesService.getScriptProperties();

  // 上書き保存
  sp.setProperty(KEY, String(VALUE).trim());

  // すぐ読み戻して確認
  const after = sp.getProperty(KEY);
  Logger.log('Saved ADMIN_USER_ID? ' + (after ? 'YES: ' + after : 'NO'));

  // 登録済みキー一覧も確認
  const all = sp.getProperties();
  Logger.log('All keys: ' + Object.keys(all).join(', '));
}

/**
 * なっぷメール抽出のテスト実行
 * @param {string} body - テストするメール本文（省略時はサンプルを使用）
 */
function testNapExtraction(body) {
  // サンプル本文（nap_email_sample.txt より）
  const sampleBody = `
件名：ご予約ありがとうございます。【BAMPO CAMP SITE】 NAPRSV-113764989 - なっぷ経由
差出人：なっぷ 予約 <rsv@nap-camp.com>

本文：
(本メールは、送信専用メールです。※返信不可)

----------------------------
▼なっぷ管理画面はこちら
https://adm.nap-camp.com/campsite/
----------------------------

お客様のメールアドレス:[sachiyou98@gmail.com]

種市 祥夫 様

この度は【BAMPO CAMP SITE】をご予約頂き誠にありがとうございます。
＜ご予約を下記にて承りました。＞

■ 予約通知

キャンプ場　　　　　 : BAMPO CAMP SITE
予約日時　　　　　　 : 2026年02月08日(日) 23時19分
予約詳細番号　　　　 : NAPRSV-113764989

■ ご予約プラン[ 1 ]

予約施設　　　　　　 : １人のみ オートフリーサイト区画A
チェックイン日時　　 : 2026年02月14日(土) 13時00分
チェックアウト日時　 : 2026年02月15日(日) 10時00分
キャンプ種別　　　　 : 宿泊キャンプ
宿泊日数　　　　　　 : 1泊2日
人数　　　　　　　　 : 大人 1人
利用料総額　　　　　 :   1,600円
【顧客支払額　　　　 :   1,600円】

代表者氏名　　　　　 : 種市 祥夫
代表者カナ　　　　　 : タネイチ サチオ
代表者住所　　　　　 : 669-1545 兵庫県三田市狭間が丘５－５サンディパークス２－１１１４
代表者連絡先　　　　 : 090-8465-9383

■ 予約時アンケート


■ ご要望



■ お支払い方法

オンラインカード決済

■ 顧客支払額総計

【1,600円】


お支払いや予約・キャンセルについてはこちらをご確認ください。
https://www.nap-camp.com/faq
  `;
  
  const targetBody = body || sampleBody;
  Logger.log('--- なっぷ抽出テスト開始 ---');
  try {
    const result = extractNapReservationData_(targetBody);
    Logger.log('抽出結果:\n' + JSON.stringify(result, null, 2));
    
    // 簡易検証ログ
    if (result.reservationId) {
      Logger.log('✅ 予約ID: ' + result.reservationId);
    } else {
      Logger.log('❌ 予約IDが抽出できませんでした');
    }
  } catch (e) {
    Logger.log('❌ エラー発生: ' + e.toString());
  }
  Logger.log('--- なっぷ抽出テスト完了 ---');
}
