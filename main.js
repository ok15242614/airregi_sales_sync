// =============================
// freee API × GAS 特定部門の現金売上一覧取得スクリプト
// =============================

// --- 部門一覧取得 ---
function fetchDepartments(token, companyId) {
  Logger.log('部門一覧を取得します（company_id=' + companyId + '）');
  // エンドポイントを新形式に修正
  const url = `https://api.freee.co.jp/api/1/companies/${companyId}/departments`;
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

// --- 特定部門の売上（現金）取得 ---
function fetchDeals(token, companyId, departmentId) {
  Logger.log('取引一覧を取得します（company_id=' + companyId + ', department_id=' + departmentId + '）');
  const url = `https://api.freee.co.jp/api/1/deals?company_id=${companyId}&department_id=${departmentId}&type=income&limit=50`;
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

// --- 現金売上のみ抽出（売上高debit＋現金credit） ---
function filterCashSales(deals) {
  const cashSales = [];
  deals.forEach(deal => {
    // 売上高（debit）明細を抽出
    const salesDetails = deal.details.filter(
      detail => detail.account_item_name === '売上高' && detail.entry_side === 'debit'
    );
    // 現金（credit）明細があるか
    const hasCashCredit = deal.details.some(
      detail => detail.account_item_name === '現金' && detail.entry_side === 'credit'
    );
    // 両方揃っていれば現金売上として抽出
    if (salesDetails.length > 0 && hasCashCredit) {
      salesDetails.forEach(sale => {
        cashSales.push({
          deal_id: deal.id,
          date: deal.issue_date,
          amount: sale.amount,
          partner_name: deal.partner_name || '',
          description: deal.description || ''
        });
      });
    }
  });
  Logger.log('現金売上件数: ' + cashSales.length);
  return cashSales;
}

// --- データ整形 ---
function formatCashSalesData(cashSales) {
  Logger.log('現金売上データ整形処理を開始します');
  if (!cashSales || cashSales.length === 0) {
    Logger.log('現金売上データが空です');
    return [];
  }
  const headers = ['取引ID', '日付', '金額', '取引先名', 'メモ'];
  const rows = cashSales.map(sale => [
    sale.deal_id,
    sale.date,
    sale.amount,
    sale.partner_name,
    sale.description
  ]);
  Logger.log('現金売上データ整形完了。行数: ' + rows.length);
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
  const token = getAccessTokenOrAuthorize();
  const companyId = getCompanyId(token);
  Logger.log('取得したcompany_id: ' + companyId);

  // まず部門一覧を取得してログ出力
  fetchDepartments(token, companyId);

  // ↓部門IDが分かったら、departmentIdにセットして再実行
  // const departmentId = ここに部門IDを入力;
  // const deals = fetchDeals(token, companyId, departmentId);
  // const cashSales = filterCashSales(deals);
  // const formattedData = formatCashSalesData(cashSales);
  // writeToSpreadsheet(formattedData);

  Logger.log('main処理が完了しました');
}
