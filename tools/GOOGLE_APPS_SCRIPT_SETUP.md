# Google Apps Script セットアップガイド

## 概要
このガイドに従って、Google Apps Script を作成・デプロイしてください。デプロイ URL を app.html と index.html に設定することで、「Google Docs で編集」ボタンが機能するようになります。

---

## 📝 Step 1: Google Apps Script を作成

### 1-1. Google Drive を開く
- https://drive.google.com にアクセス

### 1-2. 新規 Google Apps Script を作成
- **左側の「+ 新規」ボタン** をクリック
- **「その他」** を選択
- **「Google Apps Script」** をクリック

### 1-3. コードを貼り付け
Google Apps Script エディタが開いたら：

1. デフォルトのコードをすべて削除
2. 以下のコード全体をコピーして貼り付け：

```javascript
function doPost(e) {
  try {
    const content = e.parameter.content || '';

    if (!content.trim()) {
      return createJsonResponse(false, 'Content is empty');
    }

    // 新しいドキュメントを作成
    const doc = DocumentApp.create('【企画書】' + new Date().toLocaleDateString('ja-JP'));
    const docId = doc.getId();

    // ページサイズを A4 に設定
    doc.setPageSize(DocumentApp.PageSize.A4);

    // マージンを設定（標準的な設定）
    doc.setMarginBottom(36);   // 0.5 inch = 約1.27cm
    doc.setMarginTop(36);
    doc.setMarginLeft(36);
    doc.setMarginRight(36);

    // 本文を取得して設定
    const body = doc.getBody();
    body.clear();

    // 内容を複数の段落として挿入
    const lines = content.split('\n');
    lines.forEach((line, index) => {
      if (line.trim()) {
        body.appendParagraph(line);
      } else if (index < lines.length - 1) {
        // 空行も保持
        body.appendParagraph('');
      }
    });

    // ドキュメントを保存
    doc.saveAndClose();

    // ドキュメント URL を作成
    const docUrl = 'https://docs.google.com/document/d/' + docId + '/edit';

    return createJsonResponse(true, null, docUrl);

  } catch (error) {
    return createJsonResponse(false, error.toString());
  }
}

function createJsonResponse(success, error, url) {
  const response = {
    success: success,
    error: error || null,
    url: url || null
  };

  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 1-4. ファイル名を設定
- 上部の「無題のプロジェクト」をクリック
- **ファイル名: `企画書ドキュメント生成`** などに変更
- エンターキーで保存

---

## 🚀 Step 2: Google Apps Script をデプロイ

### 2-1. デプロイボタンをクリック
右上の **「デプロイ」** ボタンをクリック

### 2-2. 「新しいデプロイ」を選択
**「新しいデプロイ」** をクリック

### 2-3. デプロイ設定
以下の設定で進めます：

| 項目 | 設定値 |
|---|---|
| **種類** | **Web アプリ** |
| **実行ユーザー** | **自分** |
| **アクセス権限** | **誰でも** |

- **種類の横の「▼」** をクリック → **「Web アプリ」** を選択
- **実行ユーザー** は通常「自分」のまま
- **アクセス権限** は「誰でも」に設定（重要！）

### 2-4. デプロイ実行
**「デプロイ」** ボタンをクリック

### 2-5. 承認画面
Google アカウントの承認画面が表示されたら：
- Google アカウントを選択
- **「詳細を表示」** をクリック
- **「企画書ドキュメント生成（安全ではないページ）に移動」** をクリック
- **「許可」** をクリック

---

## 🔗 Step 3: デプロイ URL をコピー

### 3-1. デプロイ完了画面
デプロイが完了すると、以下のような情報が表示されます：

```
デプロイID: AKfycbw...xxxxx
Web アプリ URL: https://script.google.com/macros/s/AKfycbw...xxxxx/usercontent
```

### 3-2. URL をコピー
**「Web アプリ URL」** をすべてコピー

例：
```
https://script.google.com/macros/s/AKfycbw7M1dTrBxcGBLjOz9lWlVK_CKPvGkKxKxKxKx/usercontent
```

---

## 📌 Step 4: app.html と index.html に URL を設定

### 4-1. app.html を編集
ファイル `/Users/aiuego/Desktop/AI CEO_tea/clients/hinomarutraffic/tools/app.html`

検索: `const GAS_URL='https://script.google.com/macros/s/AKfycbw7M1dTrBxcGBLjOz9lWlVK_CKPvGkKxKxKxKxKx/usercontent';`

この行を見つけて、**URL 部分だけ置き換え**：

```javascript
const GAS_URL='YOUR_COPIED_URL_HERE';
```

↓

```javascript
const GAS_URL='https://script.google.com/macros/s/AKfycbw...YOUR_ACTUAL_URL.../usercontent';
```

### 4-2. index.html も同じように編集

---

## ✅ Step 5: テスト

### 5-1. 提案書に内容を入力
- ホーム → 企画 → 企画書生成
- 何か内容を入力（例：タイトル「テスト」）

### 5-2. 「📄 Google Docs で編集」ボタンをクリック
- **「Google Docs を A4 で作成中...」** アラートが表示
- 数秒後、新しい Google Docs タブが自動オープン
- **「✅ Google Docs が作成されました！」** メッセージが表示

### 5-3. Google Docs で確認
- A4 サイズで開かれているか確認
- 企画書の内容が自動で挿入されているか確認

---

## 🐛 トラブルシューティング

### URL を設定しても動かない場合

1. **URL をコピーしたか確認**
   - Web アプリ URL を正確にコピーしたか確認
   - 末尾の `/usercontent` は含めたか確認

2. **デプロイが Web アプリ型か確認**
   - Google Apps Script に戻る
   - 左側の「デプロイ」 → Web アプリが表示されているか確認

3. **キャッシュをクリア**
   - ブラウザで Ctrl+Shift+R（強制更新）を実行
   - GitHub Pages の URL を再度読み込み

### エラーメッセージが表示される場合

エラーに以下が含まれている場合：
- **「Google Docs の作成に失敗しました」**
  → GAS_URL が正しく設定されているか確認
  → Google Apps Script にエラーがないか確認（GAS のエディタでコンソールを確認）

---

## 📞 サポート

問題が発生した場合：
1. Google Apps Script エディタの **実行ログ** を確認
2. エラーメッセージをスクリーンショットで撮影
3. オーナーに報告

---

*企画書 Google Docs 自動生成 / google-apps-script-setup.md / 2026-05-25*
