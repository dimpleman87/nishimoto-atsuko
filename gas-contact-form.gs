const SHEET_NAME = '相談フォーム';
const NOTIFY_TO = 'atsukonishimoto245@gmail.com';
const SITE_URL = 'https://dimpleman87.github.io/nishimoto-atsuko/#contact';

function doGet() {
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8">' +
    `<meta http-equiv="refresh" content="0; url=${SITE_URL}">` +
    '<p>問い合わせフォームに戻ります。</p>'
  );
}

function doPost(e) {
  const params = e.parameter || {};
  const sheet = getOrCreateSheet_();
  const now = new Date();

  const row = [
    now,
    params['お名前'] || '',
    params['団体名・所属'] || '',
    params['メールアドレス'] || '',
    params['電話番号'] || '',
    params['開催予定地域'] || '',
    params['希望形式'] || '',
    params['ご相談内容'] || '',
    params.source_url || '',
    '',
    '',
  ];

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  const notificationResult = sendNotification_(params, now);
  sheet.getRange(rowNumber, 10, 1, 2).setValues([[
    notificationResult.ok ? '送信済み' : '送信失敗',
    notificationResult.message,
  ]]);

  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8">' +
    `<meta http-equiv="refresh" content="0; url=${SITE_URL}">` +
    '<p>送信ありがとうございました。ページに戻ります。</p>'
  );
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  ensureHeader_(sheet);
  return sheet;
}

function ensureHeader_(sheet) {
  const headers = [
    '受付日時',
    'お名前',
    '団体名・所属',
    'メールアドレス',
    '電話番号',
    '開催予定地域',
    '希望形式',
    'ご相談内容',
    '送信元URL',
    '通知メール',
    '通知結果メモ',
  ];

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some((header, index) => currentHeaders[index] !== header);

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function sendNotification_(params, now) {
  const subject = '西本敦子公式サイトから講座相談が届きました';
  const body = [
    '西本敦子公式サイトの相談フォームから送信がありました。',
    '',
    `受付日時: ${Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}`,
    `お名前: ${params['お名前'] || ''}`,
    `団体名・所属: ${params['団体名・所属'] || ''}`,
    `メールアドレス: ${params['メールアドレス'] || ''}`,
    `電話番号: ${params['電話番号'] || ''}`,
    `開催予定地域: ${params['開催予定地域'] || ''}`,
    `希望形式: ${params['希望形式'] || ''}`,
    '',
    'ご相談内容:',
    params['ご相談内容'] || '',
  ].join('\n');

  const options = {
    name: '西本敦子公式サイト',
  };

  if (params['メールアドレス']) {
    options.replyTo = params['メールアドレス'];
  }

  try {
    MailApp.sendEmail(NOTIFY_TO, subject, body, options);
    return {
      ok: true,
      message: `送信先: ${NOTIFY_TO}`,
    };
  } catch (error) {
    return {
      ok: false,
      message: error && error.message ? error.message : String(error),
    };
  }
}
