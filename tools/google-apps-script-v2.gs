// Google Apps Script - 企画書を Google Docs に自動作成
// A4テンプレートをコピーして、内容を自動挿入

function doPost(e) {
  const content = e?.parameter?.content || '';

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
      DriveApp.Permission.VIEW
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

function buildResponse(success, error, url) {
  return ContentService
    .createTextOutput(JSON.stringify({ success, error, url }))
    .setMimeType(ContentService.MimeType.JSON);
}
