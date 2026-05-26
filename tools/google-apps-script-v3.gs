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

/* ━━━ Cloud Sync: Google Sheet に push ━━━ */
function handlePushSheet(params) {
  try {
    const jsonData = params?.json || '{}';

    const ss = getOrCreateSheet();
    const sheet = ss.getActiveSheet();

    // row 2（メイン）に保存
    const range = sheet.getRange('B2:C2');
    range.setValues([[jsonData, new Date().toISOString()]]);

    return buildResponse(true, null, 'Pushed to Google Sheets');

  } catch (error) {
    return buildResponse(false, error.toString(), null);
  }
}

/* ━━━ Cloud Sync: Google Sheet から pull ━━━ */
function handlePullSheet() {
  try {
    const ss = getOrCreateSheet();
    const sheet = ss.getActiveSheet();

    // row 2 column B（メイン）から読み込み
    const range = sheet.getRange('B2');
    const jsonData = range.getValue();

    if (!jsonData || jsonData.trim() === '') {
      return { success: true, data: {} };
    }

    const parsed = JSON.parse(jsonData);
    return { success: true, data: parsed };

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
