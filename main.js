// =============================
// freee API × GAS 特定部門の現金売上一覧取得スクリプト
// =============================

// --- 勘定科目一覧取得 ---
function fetchAccountItems(token, companyId) {
  Logger.log('勘定科目一覧を取得します（company_id=' + companyId + '）');
  const url = `https://api.freee.co.jp/api/1/account_items?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('勘定科目一覧APIレスポンス: ' + response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error('勘定科目一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  if (!json.account_items) {
    Logger.log('APIレスポンスにaccount_itemsキーがありません。レスポンス内容: ' + response.getContentText());
    return [];
  }
  return json.account_items;
}

// --- 売上高の勘定科目ID取得 ---
function getSalesAccountItemId(accountItems) {
  const salesItem = accountItems.find(item => item.name === '売上高');
  if (!salesItem) {
    Logger.log('売上高の勘定科目が見つかりません');
    return null;
  }
  Logger.log('売上高の勘定科目ID: ' + salesItem.id);
  return salesItem.id;
}

// --- 部門一覧取得 ---
function fetchDepartments(token, companyId) {
  Logger.log('部門一覧を取得します（company_id=' + companyId + '）');
  const url = `https://api.freee.co.jp/api/1/sections?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('部門一覧APIレスポンス: ' + response.getContentText());
  Logger.log('部門一覧APIステータス: ' + response.getResponseCode());
  if (response.getResponseCode() !== 200) {
    throw new Error('部門一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  if (!json.sections) {
    Logger.log('APIレスポンスにsectionsキーがありません。レスポンス内容: ' + response.getContentText());
    return [];
  }
  if (json.sections.length === 0) {
    Logger.log('部門データがありません（sectionsは空配列）');
    return [];
  }
  json.sections.forEach(sec => {
    Logger.log('部門ID: ' + sec.id + ' / 部門名: ' + sec.name);
  });
  return json.sections;
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
function filterCashSales(deals, salesAccountItemId) {
  const cashSales = [];
  deals.forEach(deal => {
    // 売上高（debit, account_item_id一致）明細を抽出
    const salesDetails = deal.details.filter(
      detail => detail.account_item_id === salesAccountItemId && detail.entry_side === 'debit'
    );
    // 現金（credit）明細があるか（account_item_name: '現金' で判定）
    const hasCashCredit = deal.details.some(
      detail => detail.account_item_name === '現金' && detail.entry_side === 'credit'
    );
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

  // 勘定科目一覧を取得し、売上高IDを特定
  const accountItems = fetchAccountItems(token, companyId);
  const salesAccountItemId = getSalesAccountItemId(accountItems);
  if (!salesAccountItemId) {
    Logger.log('売上高の勘定科目IDが取得できなかったため処理を終了します');
    return;
  }

  // 部門一覧を取得してログ出力
  const sections = fetchDepartments(token, companyId);
  // ここで特定の部門名を指定してください（例: '店舗A'）
  const targetSectionName = 'すなば文衛門';
  const targetSection = sections.find(sec => sec.name === targetSectionName);
  if (!targetSection) {
    Logger.log('指定した部門名が見つかりません: ' + targetSectionName);
    return;
  }
  const departmentId = targetSection.id;
  Logger.log('現金売上抽出対象の部門ID: ' + departmentId + ' / 部門名: ' + targetSectionName);

  // 売上取得
  const deals = fetchDeals(token, companyId, departmentId);

  // 現金売上のみ抽出（売上高IDで判定）
  const cashSales = filterCashSales(deals, salesAccountItemId);

  // 整形してスプレッドシート出力
  const formattedData = formatCashSalesData(cashSales);
  writeToSpreadsheet(formattedData);

  Logger.log('main処理が完了しました');
}
