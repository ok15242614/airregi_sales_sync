// =============================
<<<<<<< HEAD
// freee API × GAS 特定部門の売上一覧取得スクリプト（リファクタリング版）
=======
// freee API × GAS 特定部門の売上一覧取得スクリプト
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
// =============================

// --- 設定値 ---
const TARGET_SECTION_NAME = 'すなば文衛門'; // 対象部門名
const DATE_FORMAT = 'yyyy-MM-dd'; // 日付フォーマット
const DAYS_AGO = 0; // 0:今日, 1:昨日, ...

// --- 日付取得ユーティリティ ---
function getTargetDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), DATE_FORMAT);
}

// --- freee API: 勘定科目一覧取得 ---
function fetchAccountItems(token, companyId) {
  const url = `https://api.freee.co.jp/api/1/account_items?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('勘定科目一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  return json.account_items || [];
}

// --- freee API: 部門一覧取得 ---
function fetchSections(token, companyId) {
  const url = `https://api.freee.co.jp/api/1/sections?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('部門一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  return json.sections || [];
}

<<<<<<< HEAD
// --- freee API: 取引一覧取得（type=income, 日付範囲） ---
function fetchIncomeDeals(token, companyId, startDate, endDate, offset = 0, limit = 50) {
=======
// --- 取引一覧取得（日付範囲・type=incomeで絞り込み） ---
function fetchDeals(token, companyId, startDate, endDate, offset = 0, limit = 50) {
  Logger.log('取引一覧を取得します（company_id=' + companyId + ', 日付範囲=' + startDate + '～' + endDate + '）');
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
  const url = `https://api.freee.co.jp/api/1/deals?company_id=${companyId}&type=income&start_issue_date=${startDate}&end_issue_date=${endDate}&limit=${limit}&offset=${offset}`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('取引一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  return json.deals || [];
}

<<<<<<< HEAD
// --- ユーティリティ: 売上高の勘定科目ID取得 ---
function getSalesAccountItemId(accountItems) {
  const item = accountItems.find(i => i.name === '売上高');
  if (!item) throw new Error('売上高の勘定科目が見つかりません');
  return item.id;
}

// --- ユーティリティ: 指定部門ID取得 ---
function getSectionIdByName(sections, sectionName) {
  const section = sections.find(sec => sec.name && sec.name.trim().includes(sectionName.trim()));
  if (!section) throw new Error('指定した部門名が見つかりません: ' + sectionName);
  return section.id;
}

// --- データ抽出: 特定部門・売上高のみ ---
function extractSalesBySection(deals, salesAccountItemId, sectionId) {
  const sales = [];
  deals.forEach(deal => {
    deal.details.forEach(detail => {
      if (
        detail.account_item_id === salesAccountItemId &&
        detail.entry_side === 'debit' &&
        detail.section_id === sectionId
      ) {
        sales.push({
          deal_id: deal.id,
          date: deal.issue_date,
          amount: detail.amount,
          partner_name: deal.partner_name || '',
          description: deal.description || ''
        });
      }
    });
  });
=======
// --- 特定部門の売上のみ抽出 ---
function filterSalesBySection(deals, salesAccountItemId, departmentId) {
  const sales = [];
  deals.forEach(deal => {
    // 指定部門IDの売上高（debit, account_item_id一致, section_id一致）明細を抽出
    const salesDetails = deal.details.filter(
      detail =>
        detail.account_item_id === salesAccountItemId &&
        detail.entry_side === 'debit' &&
        detail.section_id === departmentId
    );
    salesDetails.forEach(sale => {
      sales.push({
        deal_id: deal.id,
        date: deal.issue_date,
        amount: sale.amount,
        partner_name: deal.partner_name || '',
        description: deal.description || ''
      });
    });
  });
  Logger.log('売上件数: ' + sales.length);
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
  return sales;
}

// --- データ整形 ---
function formatSalesData(sales) {
<<<<<<< HEAD
=======
  Logger.log('売上データ整形処理を開始します');
  if (!sales || sales.length === 0) {
    Logger.log('売上データが空です');
    return [];
  }
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
  const headers = ['取引ID', '日付', '金額', '取引先名', 'メモ'];
  const rows = sales.map(sale => [
    sale.deal_id,
    sale.date,
    sale.amount,
    sale.partner_name,
    sale.description
  ]);
<<<<<<< HEAD
=======
  Logger.log('売上データ整形完了。行数: ' + rows.length);
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
  return [headers, ...rows];
}

// --- スプレッドシート転記 ---
function writeToSpreadsheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

// --- メイン処理 ---
function main() {
  try {
    Logger.log('main処理を開始します');
    const token = getAccessTokenOrAuthorize();
    const companyId = getCompanyId(token);

<<<<<<< HEAD
    // 日付範囲（例：今日のみ）
    const date = getTargetDate(DAYS_AGO);

    // マスタ取得
    const accountItems = fetchAccountItems(token, companyId);
    const salesAccountItemId = getSalesAccountItemId(accountItems);
    const sections = fetchSections(token, companyId);
    const sectionId = getSectionIdByName(sections, TARGET_SECTION_NAME);

    // 取引取得
    const deals = fetchIncomeDeals(token, companyId, date, date);

    // データ抽出・整形・出力
    const sales = extractSalesBySection(deals, salesAccountItemId, sectionId);
    const formatted = formatSalesData(sales);
    writeToSpreadsheet(formatted);
    Logger.log('main処理が完了しました');
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    throw e;
  }
=======
  // 日付範囲を指定（例：今日のみ）
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const startDate = today;
  const endDate = today;

  // 勘定科目一覧を取得し、売上高IDを特定
  const accountItems = fetchAccountItems(token, companyId);
  const salesAccountItemId = getSalesAccountItemId(accountItems);
  if (!salesAccountItemId) {
    Logger.log('売上高の勘定科目IDが取得できなかったため処理を終了します');
    return;
  }

  // 部門一覧を取得してログ出力
  const sections = fetchDepartments(token, companyId);
  const targetSectionName = 'すなば文衛門'; // ここに部門名を入力
  const targetSection = sections.find(sec => sec.name && sec.name.trim().includes(targetSectionName.trim()));
  if (!targetSection) {
    Logger.log('指定した部門名が見つかりません: ' + targetSectionName);
    return;
  }
  const departmentId = targetSection.id;
  Logger.log('売上抽出対象の部門ID: ' + departmentId + ' / 部門名: ' + targetSection.name);

  // 売上取得（APIでtype=income, 日付範囲で絞り込み）
  const deals = fetchDeals(token, companyId, startDate, endDate);

  // 特定部門の売上のみ抽出（売上高ID・部門IDで判定）
  const sales = filterSalesBySection(deals, salesAccountItemId, departmentId);

  // 整形してスプレッドシート出力
  const formattedData = formatSalesData(sales);
  writeToSpreadsheet(formattedData);

  Logger.log('main処理が完了しました');
>>>>>>> 4ff7330c8d4a555951d22fe6bbf158d170c22357
}
