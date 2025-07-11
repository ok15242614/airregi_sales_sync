// =============================
// freee API × GAS 連携メインスクリプト
// =============================

// --- 設定値 ---
const CLIENT_ID = '615194859329540'; // freeeアプリのクライアントID
const CLIENT_SECRET = 'Av6z0oGCBVkKPn33Eco7mkr400E4WlEX25o2JDt9vxIpCYe1fszn_tBXrhPDympOBSs6NS_MmxZUPgWyXWwmwA'; // freeeアプリのクライアントシークレット
const REDIRECT_URI = 'https://script.google.com/macros/d/1ogJV1edpGEJ0qeGrJFsLAjU0FJ6F2jA989tsRZHMZ3KSzmjGd0lO7Nyw/usercallback'; // リダイレクトURI

// OAuth2ライブラリのライブラリID: 1ogJV1edpGEJ0qeGrJFsLAjU0FJ6F2jA989tsRZHMZ3KSzmjGd0lO7Nyw

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

// --- 会社ID取得 ---
function fetchCompanyId(token) {
  const url = 'https://api.freee.co.jp/api/1/companies';
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('会社ID取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  if (!json.companies || json.companies.length === 0) {
    throw new Error('会社情報が取得できませんでした');
  }
  return json.companies[0].id; // 複数ある場合は最初の会社IDを返す
}

// --- 2. freee APIリクエスト（取引一覧） ---
function fetchFreeeData(token, companyId) {
  const url = `https://api.freee.co.jp/api/1/deals?company_id=${companyId}&limit=50`;
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
  const companyId = fetchCompanyId(token);
  const rawData = fetchFreeeData(token, companyId);
  const formattedData = formatData(rawData);
  writeToSpreadsheet(formattedData);
}
