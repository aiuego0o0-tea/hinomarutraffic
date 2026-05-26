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
