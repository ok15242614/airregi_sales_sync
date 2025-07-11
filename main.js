// =============================
// freee API × GAS 取引一覧取得メインスクリプト
// =============================

// --- 部門一覧取得 ---
function fetchDepartments(token, companyId) {
  Logger.log('部門一覧を取得します（company_id=' + companyId + '）');
  const url = `https://api.freee.co.jp/api/1/departments?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('部門一覧APIレスポンス: ' + response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error('部門一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  if (!json.departments || json.departments.length === 0) {
    Logger.log('部門データがありません');
    return [];
  }
  json.departments.forEach(dep => {
    Logger.log('部門ID: ' + dep.id + ' / 部門名: ' + dep.name);
  });
  return json.departments;
}

// --- 取引一覧取得 ---
function fetchDeals(token, companyId) {
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

// --- データ整形 ---
function formatDealsData(rawData) {
  Logger.log('データ整形処理を開始します');
  if (!rawData || rawData.length === 0) {
    Logger.log('取得データが空です');
    return [];
  }
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

// --- スプレッドシート転記 ---
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
  const token = getAccessTokenOrAuthorize(); // freeeAuth.jsの関数を利用
  const companyId = getCompanyId(token);      // freeeAuth.jsの関数を利用
  // 部門一覧を取得してログ出力
  fetchDepartments(token, companyId);
  // 必要に応じて取引取得や他の処理を続けて実行可能
  Logger.log('main処理が完了しました');
}
