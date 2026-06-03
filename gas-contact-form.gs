const SHEET_NAME = '相談フォーム';
const NOTIFY_TO = 'atsukonishimoto245@gmail.com';

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
  ];

  sheet.appendRow(row);
  sendNotification_(params, now);

  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8">' +
    '<meta http-equiv="refresh" content="0; url=https://dimpleman87.github.io/nishimoto-atsuko/#contact">' +
    '<p>送信ありがとうございました。ページに戻ります。</p>'
  );
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow([
      '受付日時',
      'お名前',
      '団体名・所属',
      'メールアドレス',
      '電話番号',
      '開催予定地域',
      '希望形式',
      'ご相談内容',
      '送信元URL',
    ]);
  }

  return sheet;
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

  GmailApp.sendEmail(NOTIFY_TO, subject, body, {
    replyTo: params['メールアドレス'] || undefined,
    name: '西本敦子公式サイト',
  });
}
