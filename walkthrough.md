# なっぷ (Nap) プラットフォーム対応実装 - ウォークスルー

## 概要

キャンプ場予約管理システムを拡張し、楽天トラベルキャンプに加えて「なっぷ」プラットフォームからの予約メールに対応しました。システムは予約元のプラットフォームを自動検出し、各プラットフォーム固有のデータを抽出し、アプリケーション全体で予約元を表示するようになりました。

## 変更内容

### 1. 設定 ([config.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/config.js))

プラットフォームごとのメールパターンを含む `PLATFORMS` 設定を追加しました：

```javascript
PLATFORMS: {
  rakuten: {
    FROM: 'no-reply@camp.travel.rakuten.co.jp',
    CONFIRM_SUBJECT: '予約が確定しました',
    CANCEL_SUBJECT: '予約がキャンセルされました'
  },
  nap: {
    FROM: 'rsv@nap-camp.com',
    CONFIRM_SUBJECT: 'ご予約ありがとうございます',
    CANCEL_SUBJECT: 'キャンセル'
  }
}
```

この設定は `prod`（本番）と `test`（テスト）の両方の環境設定に追加されました。

---

### 2. マルチプラットフォームメール抽出 ([gmail_to_sheets.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/gmail_to_sheets.js))

#### 関数名の変更
- `extractRakutenTravelEmails()` → `extractAllPlatformEmails()` にリネーム
- `mainSequence()` 内の呼び出しも更新

#### 新機能の追加

**プラットフォーム検出 (`detectPlatform_`)**:
```javascript
function detectPlatform_(msg, platforms) {
  const from = msg.getFrom() || '';
  // 送信元アドレスに基づいて 'rakuten' または 'nap' を返す
}
```

**なっぷデータ抽出 (`extractNapReservationData_`)**:
- なっぷのメール形式から予約データを抽出
- 抽出項目: 予約詳細番号, チェックイン日時, 代表者氏名 など
- なっぷ固有の日時形式（例: `2026年02月14日(土) 13時00分`）をパース

#### 処理ロジックの更新
- Gmail検索クエリが両プラットフォームに対応: `from:no-reply@camp.travel.rakuten.co.jp OR from:rsv@nap-camp.com`
- プラットフォーム間のID衝突を防ぐため、予約IDマッピングを `reservationId` → `"platform:reservationId"` に変更
- メッセージ処理ループ:
  1. 各メールのプラットフォームを検出
  2. プラットフォーム固有の予約IDパターンを抽出
  3. 適切な抽出関数を使用（楽天: `extractReservationDataFromBody_`, なっぷ: `extractNapReservationData_`）
  4. プラットフォーム識別子とともにデータを保存

#### ヘッダーの更新
スプレッドシートのヘッダーに「予約元」列を追加しました：
```javascript
const EXPECTED_HEADERS = [
  '予約日時','予約ID','予約元', // ← 追加
  'チェックイン日時','チェックアウト日時','サイト名',...
];
```

---

### 3. Webアプリ バックエンド ([webapp.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/webapp.js))

`getDataForWeb()` 関数を更新し、platformフィールドを含めるようにしました：

```javascript
return {
  id: row[colMap['予約ID']],
  platform: row[colMap['予約元']] || '楽天トラベル', // ← 追加
  status: row[colMap['ステータス']],
  // ... その他のフィールド
}
```

---

### 4. カレンダー同期 ([sync_calendar.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/sync_calendar.js))

カレンダーイベントの作成時に、タイトルと説明にプラットフォーム名を含めるように更新しました：

**変更前:**
```javascript
const title = `【予約ID:${rId}】${name}様 (${siteName})`;
```

**変更後:**
```javascript
const platform = String(rowData[colMap['予約元'] - 1] || '楽天トラベル');
const title = `【${platform}】【予約ID:${rId}】${name}様 (${siteName})`;
const desc = `予約元: ${platform}\n予約ID: ${rId}\n...`;
```

---

### 5. Web UI ([index.html](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/index.html))

#### CSSの追加
プラットフォームバッジのスタイルを追加しました：
```css
.bg-platform { background: #1976D2; }          /* 楽天トラベル (青) */
.bg-platform-nap { background: #43A047; }      /* なっぷ (緑) */
```

#### カード表示の更新
予約カードにプラットフォームバッジを追加しました：

```html
<div class="card-header">
  <div style="display:flex; gap:5px;">
    <span class="badge bg-platform">楽天トラベル</span>  <!-- ← 追加 -->
    <span class="badge bg-reserve">予約中</span>
  </div>
  <small>13:00 IN</small>
</div>
```

プラットフォームバッジの色は以下のように判定されます：
```javascript
const platformBadgeClass = row.platform === 'なっぷ' ? 'bg-platform-nap' : 'bg-platform';
```

---

## デプロイ

すべての変更は `clasp push` を使用して Google Apps Script に正常にプッシュされました：

```
Pushed 8 files.
└─ config.js
└─ gmail_to_sheets.js
└─ webapp.js
└─ sync_calendar.js
└─ index.html
└─ mail_to_client.js
└─ debug_tool.js
└─ appsscript.json
```

---

## テストに関する推奨事項

> [!IMPORTANT]
> 本番データを使用するため、まずは `test` モードでのテストを推奨します。

### 手動テスト手順

1. **テストモードへの切り替え**
   - スクリプトプロパティを設定: `APP_MODE = 'test'`
   - またはデバッグツールを使用してモードを切り替え

2. **なっぷメールの転送**
   - 実際のなっぷ予約メールをGmailに転送
   - 送信元が `rsv@nap-camp.com` であることを確認

3. **抽出の実行**
   ```javascript
   extractAllPlatformEmails()
   ```

4. **スプレッドシートの確認**
   - `テスト` シートに新しい行が追加されたことを確認
   - `予約元` 列が **なっぷ** となっていることを確認
   - すべてのフィールドが正しく抽出されていることを確認

5. **カレンダーの確認**
   ```javascript
   syncCalendarEvents()
   ```
   - イベントタイトルに `【なっぷ】` が含まれていることを確認

6. **Webアプリの確認**
   - WebアプリのURLを開く
   - 予約カードに緑色の `なっぷ` バッジが表示されていることを確認
   - フィルタリングと詳細ビューをテスト

---

## 注意点と考慮事項

### 既知の制限事項

1. **なっぷキャンセルメール件名の未確認**
   - 現在の設定ではなっぷ用キャンセル件名として `CANCEL_SUBJECT: 'キャンセル'` を仮設定しています
   - **実際のキャンセルメールの件名を確認する必要があります**
   - 確認でき次第、[config.js:L58](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/config.js#L58) を更新してください

2. **後方互換性**
   - 既存の楽天のみの予約データのプラットフォームはデフォルトで `'楽天トラベル'` として扱われます
   - 移行スクリプトは不要で、自動的に処理されます

3. **ID衝突の防止**
   - 内部的に `platform:reservationId` キーを使用しています
   - 異なるプラットフォームで同じ予約IDが存在しても衝突しません

### 使用したメールサンプル

実装はユーザーから提供されたなっぷメールサンプルに基づいています：
- [nap_email_sample.txt](file:///Users/host/.gemini/antigravity/brain/00b60508-8eec-47b8-8d76-28bdd6255461/nap_email_sample.txt)
- 予約詳細番号: `NAPRSV-113764989`
- 件名: `ご予約ありがとうございます。【BAMPO CAMP SITE】 NAPRSV-113764989 - なっぷ経由`

なっぷがメール形式を変更した場合、`extractNapReservationData_()` の抽出ロジックの更新が必要になる可能性があります。

---

## 次のステップ

1. **実際のなっぷメールでのテスト**
   - 複数のなっぷ予約確認メールで抽出精度を検証
   - キャンセルメールの処理テスト（フォーマット判明後）

2. **メール送信互換性の監視**
   - [mail_to_client.js](file:///Users/host/.gemini/antigravity/playground/blazing-galaxy/gas-project/mail_to_client.js) は変更されていません
   - なっぷの顧客へのメール送信が正しく動作することを確認
   - プラットフォーム固有のメッセージが必要な場合はテンプレートを更新

3. **本番運用開始**
   - テスト成功後、`APP_MODE` を `'prod'` に切り替え
   - 抽出エラーがないかログを監視

---

## 実装リファレンス

完全な実装計画: [implementation_plan.md](file:///Users/host/.gemini/antigravity/brain/00b60508-8eec-47b8-8d76-28bdd6255461/implementation_plan.md)
