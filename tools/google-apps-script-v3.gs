// Google Apps Script - 企画書作成 + Cloud sync (Google Sheets + JSONP)
// doPost/doGet でアクション（action パラメータ）に基づいて処理を分ける

const SHEET_ID_PROPERTY = 'CLOUD_SHEET_ID';

/* ━━━ POST リクエスト処理 ━━━ */
function doPost(e) {
  try {
    const action = e?.parameter?.action || 'create_doc';
    const content = e?.parameter?.content || '';

    // アクション分岐
    if (action === 'create_doc') {
      return handleCreateDoc(content);
    } else if (action === 'create_doc_redirect') {
      return handleCreateDocRedirect(content);
    } else if (action === 'push_sheet') {
      return handlePushSheet(e?.parameter);
    } else if (action === 'pull_sheet') {
      return handlePullSheet();
    } else {
      return buildResponse(false, 'Unknown action: ' + action, null);
    }

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ GET リクエスト処理（JSONP対応） ━━━ */
function doGet(e) {
  try {
    const action = e?.parameter?.action || '';
    const callback = e?.parameter?.callback || '';

    if (action === 'pull_sheet_jsonp') {
      const data = handlePullSheet();
      if (callback) {
        // JSONP形式で返す
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      } else {
        return buildResponse(false, 'callback parameter required', null);
      }
    } else if (action === 'pull_sheet') {
      return buildResponse(true, null, handlePullSheet());
    } else {
      return buildResponse(false, 'Unknown action: ' + action, null);
    }

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ 企画書作成（Google Docs） ━━━ */
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
      const p = body.appendParagraph(line.trim());
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      p.setIndentStart(0);
      p.setIndentFirstLine(0);
      p.setSpacingBefore(0);
      p.setSpacingAfter(6);
    });

    doc.saveAndClose();

    const docUrl = 'https://docs.google.com/document/d/' + copiedFile.getId() + '/edit';

    return buildResponse(true, null, docUrl);

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ 企画書作成＆リダイレクト ━━━ */
function handleCreateDocRedirect(content) {
  try {
    const result = handleCreateDoc(content);
    if (!result.includes('"success":true')) {
      return result;
    }

    const jsonMatch = result.match(/"data":"([^"]+)"/);
    const docUrl = jsonMatch ? jsonMatch[1] : '';

    const html = '<script>window.location.href = "' + docUrl + '";</script>';
    return ContentService
      .createTextOutput(html)
      .setMimeType(ContentService.MimeType.HTML);

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ Google Sheets シートの取得または作成 ━━━ */
function getOrCreateSheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let sheetId = scriptProperties.getProperty(SHEET_ID_PROPERTY);

  if (!sheetId) {
    // シート作成
    const fileName = 'hinomaru-cloud-sync-' + new Date().getTime();
    const sheet = SpreadsheetApp.create(fileName);
    sheetId = sheet.getId();

    // プロパティに保存（全ユーザー共有）
    scriptProperties.setProperty(SHEET_ID_PROPERTY, sheetId);

    // 初期化：ヘッダー行 + データ行
    const ss = SpreadsheetApp.openById(sheetId);
    const activeSheet = ss.getActiveSheet();
    activeSheet.getRange('A1:C1').setValues([['key', 'json', 'updatedAt']]);
    activeSheet.getRange('A2:C2').setValues([['main', '{}', new Date().toISOString()]]);
  }

  return SpreadsheetApp.openById(sheetId);
}

/* ━━━ Cloud Sync: JSONBin に push ━━━ */
function handlePushSheet(params) {
  try {
    const jsonData = params?.data || params?.json || '{}';
    const apiKey = '$2a$10$izhs6odWY1yOkkqy2BBAfe3MudJPZZ73X3t9xwLUnkT6AnVvBFYte';
    const binId = '6a0da0cf6610dd3ae876b58a';

    const payload = typeof jsonData === 'string' ? JSON.parse(jsonData) : jsonData;
    const res = UrlFetchApp.fetch(`https://api.jsonbin.io/v3/b/${binId}`, {
      method: 'put',
      headers: {
        'X-Master-Key': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() === 200) {
      return buildResponse(true, null, 'Pushed to JSONBin');
    } else {
      return buildResponse(false, 'JSONBin push failed: ' + res.getResponseCode(), null);
    }

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ Cloud Sync: JSONBin から pull ━━━ */
function handlePullSheet() {
  try {
    const apiKey = '$2a$10$izhs6odWY1yOkkqy2BBAfe3MudJPZZ73X3t9xwLUnkT6AnVvBFYte';
    const binId = '6a0da0cf6610dd3ae876b58a';

    const res = UrlFetchApp.fetch(`https://api.jsonbin.io/v3/b/${binId}/latest`, {
      headers: {
        'X-Master-Key': apiKey,
        'X-Bin-Meta': 'false'
      },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return { success: false, error: 'JSONBin pull failed: ' + res.getResponseCode(), data: null };
    }

    const json = JSON.parse(res.getContentText());
    const data = json.record || json;
    return { success: true, data: data };

  } catch (error) {
    return { success: false, error: error.toString(), data: null };
  }
}

/* ━━━ レスポンス構築（JSON） ━━━ */
function buildResponse(success, error, data) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: success,
      error: error || null,
      data: data || null
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
