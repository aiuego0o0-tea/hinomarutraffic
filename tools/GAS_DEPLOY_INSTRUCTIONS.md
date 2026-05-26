# GAS デプロイ手順（Cloud sync 対応版）

## 目的
企画書作成 + Cloud sync（JSONBin）を GAS で一括処理
CORS エラーを完全に回避

## 手順

### 1. 既存の Google Apps Script を開く
- https://script.google.com にアクセス
- 「企画書ドキュメント生成」（または既存のプロジェクト）を開く

### 2. コードを置き換え
**現在の doPost() を全て削除して、以下をコピー貼り付け：**

```javascript
// Google Apps Script - 企画書作成 + Cloud sync 中継
// doPost でアクション（action パラメータ）に基づいて処理を分ける

function doPost(e) {
  try {
    const action = e?.parameter?.action || 'create_doc';
    const content = e?.parameter?.content || '';

    // アクション分岐
    if (action === 'create_doc') {
      return handleCreateDoc(content);
    } else if (action === 'pull_cloud') {
      return handlePullCloud(e?.parameter);
    } else if (action === 'push_cloud') {
      return handlePushCloud(e?.parameter);
    } else {
      return buildResponse(false, 'Unknown action: ' + action, null);
    }

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ 企画書作成 ━━━ */
function handleCreateDoc(content) {
  if (!content.trim()) {
    return buildResponse(false, 'Content is empty', null);
  }

  try {
    const templateId = '1EUyeyxEVYmtX751Sii5uqKtqaoEXjxMziZmcuCHuVXI';

    const fileName = '【企画書】' + Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy/MM/dd HH:mm'
    );

    const copiedFile = DriveApp.getFileById(templateId).makeCopy(fileName);

    copiedFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.EDIT
    );

    const doc = DocumentApp.openById(copiedFile.getId());
    const body = doc.getBody();

    body.clear();

    content.split('\n').forEach(line => {
      if (line.trim()) {
        body.appendParagraph(line.trim());
      }
    });

    doc.saveAndClose();

    const docUrl = 'https://docs.google.com/document/d/' + copiedFile.getId() + '/edit';

    return buildResponse(true, null, docUrl);

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ Cloud Sync: JSONBin から pull ━━━ */
function handlePullCloud(params) {
  try {
    const binId = params?.binId || '';
    const apiKey = params?.apiKey || '';

    if (!binId || !apiKey) {
      return buildResponse(false, 'Missing binId or apiKey', null);
    }

    const url = `https://api.jsonbin.io/v3/b/${binId}/latest`;
    const options = {
      method: 'get',
      headers: {
        'X-Master-Key': apiKey,
        'X-Bin-Meta': 'false'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();

    if (status !== 200) {
      return buildResponse(false, 'JSONBin error: HTTP ' + status, null);
    }

    const data = JSON.parse(response.getContentText());
    const record = data.record || data;

    return buildResponse(true, null, JSON.stringify(record));

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ Cloud Sync: JSONBin に push ━━━ */
function handlePushCloud(params) {
  try {
    const binId = params?.binId || '';
    const apiKey = params?.apiKey || '';
    const data = params?.data || '{}';

    if (!binId || !apiKey) {
      return buildResponse(false, 'Missing binId or apiKey', null);
    }

    const url = `https://api.jsonbin.io/v3/b/${binId}`;
    const options = {
      method: 'put',
      headers: {
        'X-Master-Key': apiKey,
        'Content-Type': 'application/json'
      },
      payload: data,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();

    if (status !== 200) {
      return buildResponse(false, 'JSONBin error: HTTP ' + status, null);
    }

    return buildResponse(true, null, 'Pushed to JSONBin');

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ レスポンス構築 ━━━ */
function buildResponse(success, error, data) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: success,
      error: error || null,
      data: data || null
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 3. 保存 & デプロイ
- Ctrl+S（または Cmd+S）で保存
- 右上「デプロイ」をクリック
- 「新規デプロイ」を選択
- 設定：
  - **種類：** Web アプリ
  - **実行ユーザー：** 自分
  - **アクセス権限：** 誰でも
- 「デプロイ」をクリック

### 4. デプロイ URL を確認
新しい Web アプリ URL が表示されます：
```
https://script.google.com/macros/s/AKfycbXXXXX.../usercontent
```

**注意：末尾は `/usercontent` ですが、フロント側の `GAS_URL` は `/exec` のままです**（既存の設定を変更しないで OK）

---

## テスト内容

デプロイ後、以下をテストしてください：

1. **企画書作成** - Google Docs で編集ボタン をクリック
   - 期待：新タブで Google Docs が開く（変化なし）

2. **Cloud sync がない場合** - ページロード時
   - 期待：コンソールに CORS エラーが出ない ✅

3. **Cloud sync を有効にした場合** - チーム接続で syncCfg を設定
   - 期待：「今すぐ同期」クリック時に GAS 経由で JSONBin と通信
   - コンソールに CORS エラーが出ない ✅

---

## 重要ポイント

- **GAS_URL は変わらない** - app.html / index.html の `GAS_URL` は修正済み
- **アクション分岐** - `action` パラメータで自動判別
- **CORS 回避** - サーバーサイド（GAS）で JSONBin にアクセスするので、ブラウザの CORS チェックを受けない

