/**
 * config.gs
 * - 概要：gmailから楽天トラベルの予約情報抽出(シート楽天トラベルorテスト)、メールを予約日の次の日、チェックイン日時前日に送信(シート楽天トラベルorテスト)
 * - 抽出/送信の「本番/テスト切替」をここだけで管理する
 * - ※Script Properties で APP_MODE を切り替える(prodをtestにしないとテストモードにならない)
 * - テスト時は安全装置でDRY_RUN: trueになっているので、gmailでの確認時はfalseにする
 * 使い方（おすすめ）
 * 1) Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」
 *    APP_MODE = test   もしくは  prod  を設定
 * 2) 既存コード側は getAppConfig_() を呼んで設定を取る
 * 3) 関数プルダウンから setModeProd()実行で本番モード、setModeTest()実行でテストモード
 */

/** ==== 環境モード ==== **/
const DEFAULT_APP_MODE = 'prod';

function getAppMode_() {
  const v = PropertiesService.getScriptProperties().getProperty('APP_MODE');
  const mode = (v || DEFAULT_APP_MODE).trim();
  if (!['prod', 'test'].includes(mode)) throw new Error('APP_MODE は prod か test を指定してください: ' + mode);
  return mode;
}

function setModeProd() { PropertiesService.getScriptProperties().setProperty('APP_MODE', 'prod'); }
function setModeTest() { PropertiesService.getScriptProperties().setProperty('APP_MODE', 'test'); }

/** ==== ここに全設定を集約 ==== **/
function getAppConfig_() {
  const mode = getAppMode_();

  // 共通設定
  const COMMON = {
    SHEET_ID: '1mEX0QX0KZAQqYocNIKk0P0sBXIKz4ZONBFjFvsj1rHY',
    JST: 'Asia/Tokyo',
    CALENDAR_ID: 'ec5e5887f78caf546266792071b04a84c549fdef877914591ce021a8bfea746e@group.calendar.google.com'
  };

  // 環境別設定
  const ENV = {
    prod: {
      extractor: {
        MODE: 'prod',
        SHEET_NAME: '楽天トラベル',
        PROCESSED_LABEL: 'キャンプ予約',
        MAX_THREADS: 500,
        ADD_LABEL: true,
        CLEANUP_TEST_LABEL: false,
        SEARCH_PERIOD: '30d',
        // プラットフォーム別設定
        PLATFORMS: {
          rakuten: {
            SHEET_NAME: '楽天トラベル',
            FROM: 'no-reply@camp.travel.rakuten.co.jp',
            CONFIRM_SUBJECT: '予約が確定しました',
            CANCEL_SUBJECT: '予約がキャンセルされました'
          },
          nap: {
            SHEET_NAME: 'なっぷ',
            FROM: 'rsv@nap-camp.com',
            CONFIRM_SUBJECT: 'ご予約ありがとうございます',
            CANCEL_SUBJECT: 'キャンセル' // 要確認：実際のキャンセルメール件名
          }
        }
      },
      mailer: {
        SHEET_NAME: '楽天トラベル',
        SETTING_SHEET_NAME: 'メール設定',
        DRY_RUN: false,
        LOCK_CODE: '2727',
        FROM_NAME: 'BAMPO CAMP SITE',
        REPLY_TO: '919saigo@gmail.com',
        ADMIN_SIGNATURE: '【管理人 西郷】\n電話番号: 090-9753-7103\nメール: 919saigo@gmail.com',
        NOW_OVERRIDE: ''
      }
    },
    test: {
      extractor: {
        MODE: 'test',
        SHEET_NAME: 'テスト',
        PROCESSED_LABEL: 'テスト',
        MAX_THREADS: 200,
        ADD_LABEL: true,
        CLEANUP_TEST_LABEL: true,
        SEARCH_PERIOD: '30d',
        // プラットフォーム別設定（本番と同じ構造でテスト用シートを指定）
        PLATFORMS: {
          rakuten: {
            SHEET_NAME: 'テスト',
            FROM: 'no-reply@camp.travel.rakuten.co.jp',
            CONFIRM_SUBJECT: '予約が確定しました',
            CANCEL_SUBJECT: '予約がキャンセルされました'
          },
          nap: {
            SHEET_NAME: 'テスト（なっぷ）',
            FROM: 'rsv@nap-camp.com',
            CONFIRM_SUBJECT: 'ご予約ありがとうございます',
            CANCEL_SUBJECT: 'キャンセル' // 要確認：実際のキャンセルメール件名
          }
        }
      },
      mailer: {
        SHEET_NAME: 'テスト',
        SETTING_SHEET_NAME: 'メール設定',
        DRY_RUN: false,
        LOCK_CODE: '2727',
        FROM_NAME: 'BAMPO CAMP SITE (TEST)',
        REPLY_TO: '919saigo@gmail.com',
        ADMIN_SIGNATURE: '【管理人 西郷】\n電話番号: 090-9753-7103\nメール: 919saigo@gmail.com',
        NOW_OVERRIDE: ''
      }
    }
  };

  return {
    mode,
    ...COMMON,
    extractor: ENV[mode].extractor,
    mailer: ENV[mode].mailer
  };
}