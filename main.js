// =============================
// freee API × GAS 連携メインスクリプト
// =============================

// --- 設定値 ---
const CLIENT_ID = 'YOUR_CLIENT_ID'; // freeeアプリのクライアントID
const CLIENT_SECRET = 'YOUR_CLIENT_SECRET'; // freeeアプリのクライアントシークレット
const REDIRECT_URI = 'YOUR_REDIRECT_URI'; // リダイレクトURI
const COMPANY_ID = 'YOUR_COMPANY_ID'; // freee会計の会社ID

// OAuth2ライブラリのライブラリID: 1B7Jc7wG2QbFJi7dUo6g9gK2kT9gkQ4v1b1Jg1j1Jg1Jg1Jg1Jg1Jg1Jg1

// --- 1. OAuth2認証 ---
function getOAuthService() {
  return OAuth2.createService('freee')
    .setAuthorizationBaseUrl('https://accounts.secure.freee.co.jp/public_api/authorize')
    .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('read')
    .setParam('response_type', 'code')
    .setParam('redirect_uri', REDIRECT_URI);
}

function authorize() {
  const service = getOAuthService();
  const authorizationUrl = service.getAuthorizationUrl();
  Logger.log('以下のURLにアクセスして認証してください: ' + authorizationUrl);
  // 認証URLを表示（GASの場合はログ出力やUI表示など）
}

function authCallback(request) {
  const service = getOAuthService();
  const isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証に成功しました。画面を閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しました。');
  }
}

// --- 2. freee APIリクエスト（取引一覧） ---
function fetchFreeeData(token) {
  const url = `https://api.freee.co.jp/api/1/deals?company_id=${COMPANY_ID}&limit=50`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('freee APIリクエスト失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  return json.deals || [];
}

// --- 3. データ整形 ---
function formatData(rawData) {
  if (!rawData || rawData.length === 0) return [];
  // 例: 取引ID、日付、金額、取引先名、メモ
  const headers = ['取引ID', '日付', '金額', '取引先名', 'メモ'];
  const rows = rawData.map(deal => [
    deal.id,
    deal.issue_date,
    deal.amount,
    deal.partner_name || '',
    deal.description || ''
  ]);
  return [headers, ...rows];
}

// --- 4. スプレッドシート転記 ---
function writeToSpreadsheet(formattedData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  if (formattedData.length > 0) {
    sheet.getRange(1, 1, formattedData.length, formattedData[0].length).setValues(formattedData);
  }
}

// --- メイン処理 ---
function main() {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    authorize();
    return;
  }
  const token = service.getAccessToken();
  const rawData = fetchFreeeData(token);
  const formattedData = formatData(rawData);
  writeToSpreadsheet(formattedData);
}
