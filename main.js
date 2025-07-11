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
  Logger.log('OAuth2サービスを初期化します');
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
  Logger.log('認証が必要です。以下のURLにアクセスして認証してください: ' + authorizationUrl);
  // 認証URLを表示（GASの場合はログ出力やUI表示など）
}

function authCallback(request) {
  Logger.log('authCallbackが呼び出されました');
  const service = getOAuthService();
  const isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    Logger.log('認証に成功しました');
    return HtmlService.createHtmlOutput('認証に成功しました。画面を閉じてください。');
  } else {
    Logger.log('認証に失敗しました');
    return HtmlService.createHtmlOutput('認証に失敗しました。');
  }
}

// --- 会社ID取得 ---
function fetchCompanyId(token) {
  Logger.log('会社IDを取得します');
  const url = 'https://api.freee.co.jp/api/1/companies';
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('会社情報APIレスポンス: ' + response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error('会社ID取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  if (!json.companies || json.companies.length === 0) {
    throw new Error('会社情報が取得できませんでした');
  }
  Logger.log('取得した会社ID: ' + json.companies[0].id);
  return json.companies[0].id; // 複数ある場合は最初の会社IDを返す
}

// --- 2. freee APIリクエスト（取引一覧） ---
function fetchFreeeData(token, companyId) {
  Logger.log('取引一覧を取得します（company_id=' + companyId + '）');
  const url = `https://api.freee.co.jp/api/1/deals?company_id=${companyId}&limit=50`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('取引一覧APIレスポンス: ' + response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error('freee APIリクエスト失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  Logger.log('取得した取引件数: ' + (json.deals ? json.deals.length : 0));
  return json.deals || [];
}

// --- 3. データ整形 ---
function formatData(rawData) {
  Logger.log('データ整形処理を開始します');
  if (!rawData || rawData.length === 0) {
    Logger.log('取得データが空です');
    return [];
  }
  // 例: 取引ID、日付、金額、取引先名、メモ
  const headers = ['取引ID', '日付', '金額', '取引先名', 'メモ'];
  const rows = rawData.map(deal => [
    deal.id,
    deal.issue_date,
    deal.amount,
    deal.partner_name || '',
    deal.description || ''
  ]);
  Logger.log('データ整形完了。行数: ' + rows.length);
  return [headers, ...rows];
}

// --- 4. スプレッドシート転記 ---
function writeToSpreadsheet(formattedData) {
  Logger.log('スプレッドシートに書き込みます');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  if (formattedData.length > 0) {
    sheet.getRange(1, 1, formattedData.length, formattedData[0].length).setValues(formattedData);
    Logger.log('スプレッドシートへの書き込み完了: ' + formattedData.length + '行');
  } else {
    Logger.log('書き込むデータがありません');
  }
}

// --- メイン処理 ---
function main() {
  Logger.log('main処理を開始します');
  const service = getOAuthService();
  if (!service.hasAccess()) {
    Logger.log('アクセストークンがありません。認証を開始します');
    authorize();
    return;
  }
  Logger.log('アクセストークン取得済み。API処理を開始します');
  const token = service.getAccessToken();
  const companyId = fetchCompanyId(token);
  const rawData = fetchFreeeData(token, companyId);
  const formattedData = formatData(rawData);
  writeToSpreadsheet(formattedData);
  Logger.log('main処理が完了しました');
}
