// =============================
// freee API 認証・共通処理モジュール
// =============================

// --- 設定値 ---
const CLIENT_ID = '615194859329540'; // freeeアプリのクライアントID
const CLIENT_SECRET = 'Av6z0oGCBVkKPn33Eco7mkr400E4WlEX25o2JDt9vxIpCYe1fszn_tBXrhPDympOBSs6NS_MmxZUPgWyXWwmwA'; // freeeアプリのクライアントシークレット
const REDIRECT_URI = 'https://script.google.com/macros/d/1ogJV1edpGEJ0qeGrJFsLAjU0FJ6F2jA989tsRZHMZ3KSzmjGd0lO7Nyw/usercallback'; // リダイレクトURI

// OAuth2ライブラリのライブラリID: 1ogJV1edpGEJ0qeGrJFsLAjU0FJ6F2jA989tsRZHMZ3KSzmjGd0lO7Nyw

// --- OAuth2サービス生成 ---
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

// --- アクセストークン取得 ---
function getAccessTokenOrAuthorize() {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    Logger.log('アクセストークンがありません。認証を開始します');
    authorize();
    throw new Error('認証が必要です。ログに表示されたURLで認証してください。');
  }
  Logger.log('アクセストークン取得済み');
  return service.getAccessToken();
}

// --- 認証フロー開始 ---
function authorize() {
  const service = getOAuthService();
  const authorizationUrl = service.getAuthorizationUrl();
  Logger.log('認証が必要です。以下のURLにアクセスして認証してください: ' + authorizationUrl);
}

// --- コールバック関数 ---
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
function getCompanyId(token) {
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