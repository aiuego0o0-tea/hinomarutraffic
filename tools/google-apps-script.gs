// Google Apps Script - 企画書を Google Docs に自動作成
// このコードを Google Apps Script エディタに貼り付けてください

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
