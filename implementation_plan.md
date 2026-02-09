# なっぷ (Nap) 予約対応の追加 - 実装計画

現在、楽天トラベルキャンプからの予約メールのみに対応しているシステムを拡張し、「なっぷ」からの予約メールにも対応できるようにします。

## ユーザー確認事項

> [!IMPORTANT]
> **なっぷからの実際のメール形式を確認する必要があります**
> 
> この実装計画では、Web検索で得られた一般的ななっぷのメール形式を基に設計していますが、実際に受信されるメールの形式（件名、送信元アドレス、本文内のフィールド名など）を確認してから実装を進めることを強く推奨します。
> 
> 確認すべき項目：
> - 送信元メールアドレス（例: `noreply@nap-camp.com` など）
> - 予約確認メールの件名パターン
> - キャンセル通知メールの件名パターン
> - 本文内のフィールド名（予約番号、チェックイン日、チェックアウト日、人数、料金など）

> [!WARNING]
> **既存の楽天トラベル機能への影響**
> 
> この変更により、既存の楽天トラベル予約処理に影響が出ないよう、慎重にテストする必要があります。特にテストモードで十分に動作確認を行ってから本番環境へ適用してください。

## 変更内容

### コアファイル

#### [MODIFY] [gmail_to_sheets.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/gmail_to_sheets.js)

**変更内容：**
マルチプラットフォーム対応のための抽出ロジックのリファクタリング

1. **プラットフォーム検出機能の追加**
   - メールの送信元、件名から「楽天トラベル」か「なっぷ」かを判定する関数 `detectPlatform_(msg)` を追加
   - 戻り値: `'rakuten'` または `'nap'`

2. **メイン抽出関数のリファクタリング**
   - 現在の `extractRakutenTravelEmails()` を `extractAllPlatformEmails()` にリネーム
   - Gmail検索クエリを両プラットフォームに対応
   ```javascript
   const query = [
     '(',
     'from:no-reply@camp.travel.rakuten.co.jp',
     'OR from:noreply@nap-camp.com', // なっぷの送信元（要確認）
     ')',
     '(subject:予約)',
     `newer_than:${SEARCH_PERIOD}`
   ].join(' ');
   ```

3. **プラットフォーム別抽出関数の分離**
   - 既存の `extractReservationDataFromBody_()` を楽天専用に
   - 新規に `extractNapReservationData_(body)` を追加
   - 共通インターフェースで統一されたデータ構造を返す

4. **データ構造の拡張**
   - スプレッドシートに「予約元」列を追加（`'楽天トラベル'` または `'なっぷ'`）
   - 予約IDの重複チェックを「予約元+予約ID」の組み合わせで実施

5. **なっぷ用の抽出ロジック**
   ```javascript
   function extractNapReservationData_(body) {
     // なっぷのメール形式に合わせた抽出
     // 予約番号、チェックイン/アウト日時、人数、名前、電話、メール、料金など
     // 実際のメールサンプルを基に調整が必要
   }
   ```

---

### 設定

#### [MODIFY] [config.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/config.js)

**変更内容：**
なっぷ用の設定を追加

1. **プラットフォーム設定の追加**
   ```javascript
   const ENV = {
     prod: {
       extractor: {
         // 既存設定に追加
         PLATFORMS: {
           rakuten: {
             FROM: 'no-reply@camp.travel.rakuten.co.jp',
             CONFIRM_SUBJECT: '予約が確定しました',
             CANCEL_SUBJECT: '予約がキャンセルされました'
           },
           nap: {
             FROM: 'noreply@nap-camp.com', // 要確認
             CONFIRM_SUBJECT: 'ご予約完了のお知らせ', // 要確認
             CANCEL_SUBJECT: 'キャンセル' // 要確認
           }
         }
       }
     }
   }
   ```

---

### Webインターフェース

#### [MODIFY] [webapp.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/webapp.js)

**変更内容：**
予約元情報の表示対応

1. **`getDataForWeb()`関数の拡張**
   - 返すデータに `platform` フィールドを追加
   ```javascript
   return {
     id: row[colMap['予約ID']],
     platform: row[colMap['予約元']] || '楽天トラベル', // 新規列
     status: row[colMap['ステータス']],
     // ... 既存のフィールド
   };
   ```

#### [MODIFY] [index.html](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/index.html)

**変更内容：**
UIでの予約元表示

1. **カード表示の拡張**
   - 予約元（楽天トラベル/なっぷ）をバッジで表示
   ```html
   <span class="badge bg-platform">${row.platform}</span>
   ```

2. **フィルター機能の追加**
   - 予約元でのフィルタリング機能を追加（オプション）

3. **詳細モーダルの拡張**
   - 詳細情報に予約元を表示

---

### カレンダーとメール

#### [MODIFY] [sync_calendar.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/sync_calendar.js)

**変更内容：**
カレンダーイベントのタイトルに予約元を含める

1. **イベントタイトルの拡張**
   ```javascript
   const platform = String(rowData[colMap['予約元'] - 1] || '楽天');
   const title = `【${platform}】【予約ID:${rId}】${name}様 (${siteName})`;
   ```

#### [MODIFY] [mail_to_client.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/mail_to_client.js)

**変更内容：**
メール送信ロジックは予約元に関わらず共通で使用可能（変更不要、または予約元を本文に含めるオプション追加）

---

## 検証計画

### 自動テスト

現在、このプロジェクトには自動テストが存在しないため、手動テストで検証します。

### 手動検証

#### 1. テストモードでの動作確認

```
前提条件:
- Script Properties で APP_MODE = 'test' を設定
- config.js の test 設定で DRY_RUN = false（または Gmail確認用にtrue）
- テスト用のスプレッドシート「テスト」シートを使用
```

**ステップ 1: なっぷの予約メールをGmailに送信（または転送）**
- ユーザーに、実際になっぷから受信した予約確認メールを用意してもらう
- 必要に応じて、件名や送信元を調整してテスト用メールを作成

**ステップ 2: Gmail検索クエリの確認**
```
1. Apps Script エディタで gmail_to_sheets.js を開く
2. extractAllPlatformEmails() 関数を実行
3. ログを確認し、検索クエリが正しく両プラットフォームを対象にしているか確認
4. 実行ログで「なっぷ」のメールが検出されているか確認
```

**ステップ 3: 抽出結果の確認**
```
1. スプレッドシート「テスト」シートを開く
2. 以下を確認：
   - なっぷの予約データが新規行として追加されているか
   - 「予約元」列に「なっぷ」と記録されているか
   - 予約ID, 名前, チェックイン/アウト日時, 料金などが正しく抽出されているか
   - 既存の楽天トラベルのデータに影響がないか
```

**ステップ 4: Web画面での表示確認**
```
1. WebアプリのURLにアクセス
2. リストビューで「なっぷ」からの予約が表示されるか確認
3. カードに予約元バッジが表示されているか確認
4. 詳細モーダルを開いて予約元情報が表示されているか確認
5. カレンダービューで「なっぷ」の予約がイベントとして表示されているか確認
```

**ステップ 5: カレンダー連携の確認**
```
1. sync_calendar.js の syncCalendarEvents() 関数を実行
2. Googleカレンダーを開く
3. なっぷの予約がイベントとして作成されているか確認
4. イベントタイトルに「なっぷ」と表示されているか確認
```

**ステップ 6: キャンセル処理の確認**
```
1. なっぷからキャンセルメールを受信（またはテスト用に用意）
2. extractAllPlatformEmails() 関数を実行
3. スプレッドシートで該当予約のステータスが「キャンセル済み」に更新されているか確認
4. Googleカレンダーからイベントが削除されているか確認
```

#### 2. 本番環境での段階的適用

```
1. Script Properties で APP_MODE = 'prod' に変更
2. 本番シート「楽天トラベル」を使用
3. 小規模なテスト期間を設けて、なっぷの予約が1-2件発生した後に動作確認
4. 問題がなければ、そのまま運用継続
```

#### 3. ユーザー確認事項

以下の点についてユーザーに確認を依頼：

1. **なっぷから実際に受信する予約確認メールのサンプルを提供してもらう**
   - 送信元アドレス
   - 件名
   - 本文の形式（特に予約番号、日付、料金などのフィールド名）

2. **キャンセルメールのサンプルも同様に提供してもらう**

3. **スプレッドシートに「予約元」列を追加することに同意してもらう**

4. **Web画面のデザイン（予約元バッジの表示位置や色）について要望があれば確認**
