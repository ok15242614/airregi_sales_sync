// =============================
// freee API × GAS 特定部門の売上一覧取得スクリプト（リファクタリング版）
// =============================

// --- 設定値 ---
const TARGET_SECTION_NAME = 'すなば文衛門'; // 対象部門名
const DATE_FORMAT = 'yyyy-MM-dd'; // 日付フォーマット
const DAYS_AGO = 0; // 0:今日, 1:昨日, ...
// const SALES_ACCOUNT_ITEM_ID = 499030063; // 売上高のaccount_item_idを直接指定

// --- 日付取得ユーティリティ ---
function getTargetDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);
  const formatted = Utilities.formatDate(date, Session.getScriptTimeZone(), DATE_FORMAT);
  Logger.log(`【日付取得】${daysAgo}日前: ${formatted}`);
  return formatted;
}

// --- freee API: 勘定科目一覧取得 ---
function fetchAccountItems(token, companyId) {
  Logger.log(`【APIリクエスト】勘定科目一覧 company_id=${companyId}`);
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
  Logger.log(`【APIレスポンス】勘定科目件数: ${json.account_items ? json.account_items.length : 0}`);
  return json.account_items || [];
}

// --- freee API: 部門一覧取得 ---
function fetchSections(token, companyId) {
  Logger.log(`【APIリクエスト】部門一覧 company_id=${companyId}`);
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
  Logger.log(`【APIレスポンス】部門件数: ${json.sections ? json.sections.length : 0}`);
  return json.sections || [];
}

// --- freee API: 取引一覧取得（type=income, 日付範囲） ---
function fetchIncomeDeals(token, companyId, startDate, endDate, offset = 0, limit = 50) {
  Logger.log(`【APIリクエスト】取引一覧 company_id=${companyId} 日付: ${startDate}～${endDate} offset=${offset} limit=${limit}`);
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
  Logger.log(`【APIレスポンス】取引件数: ${json.deals ? json.deals.length : 0}`);
  return json.deals || [];
}

// --- ユーティリティ: 売上高の勘定科目ID取得 ---
function getSalesAccountItemId(accountItems) {
  const item = accountItems.find(i => i.name === '売上高');
  if (!item) throw new Error('売上高の勘定科目が見つかりません');
  Logger.log(`【売上高ID特定】売上高 account_item_id=${item.id}`);
  return item.id;
}

// --- ユーティリティ: 指定部門ID取得 ---
function getSectionIdByName(sections, sectionName) {
  const section = sections.find(sec => sec.name && sec.name.trim().includes(sectionName.trim()));
  if (!section) throw new Error('指定した部門名が見つかりません: ' + sectionName);
  Logger.log(`【部門ID特定】${sectionName} section_id=${section.id}`);
  return section.id;
}

// --- データ抽出: 特定部門・売上高のみ ---
function extractSalesBySection(deals, salesAccountItemId, sectionId) {
  Logger.log(`【データ抽出】売上高ID=${salesAccountItemId} 部門ID=${sectionId}`);
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
  Logger.log(`【抽出結果】売上明細件数: ${sales.length}`);
  if (sales.length > 0) {
    Logger.log(`【抽出サンプル】1件目: ${JSON.stringify(sales[0])}`);
  }
  return sales;
}

// --- データ整形 ---
function formatSalesData(sales) {
  const headers = ['取引ID', '日付', '金額', '取引先名', 'メモ'];
  const rows = sales.map(sale => [
    sale.deal_id,
    sale.date,
    sale.amount,
    sale.partner_name,
    sale.description
  ]);
  Logger.log(`【データ整形】行数: ${rows.length + 1}`);
  return [headers, ...rows];
}

// --- スプレッドシート転記 ---
function writeToSpreadsheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log(`【スプレッドシート】クリア前 行数: ${sheet.getLastRow()}`);
  sheet.clear();
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    Logger.log(`【スプレッドシート】書き込み後 行数: ${data.length}`);
  } else {
    Logger.log('【スプレッドシート】書き込むデータがありません');
  }
}

// --- メイン処理 ---
function main() {
  try {
    Logger.log('main処理を開始します');
    const token = getAccessTokenOrAuthorize();
    Logger.log(`【認証】アクセストークン取得: ${token ? 'OK' : 'NG'}`);
    const companyId = getCompanyId(token);
    Logger.log(`【会社ID】company_id=${companyId}`);
    // 日付範囲（過去1カ月分）
    const startDate = getTargetDate(30);
    const endDate = getTargetDate(0);
    Logger.log(`【日付範囲】${startDate} ～ ${endDate}`);
    // マスタ取得
    const accountItems = fetchAccountItems(token, companyId); // 勘定科目一覧を取得
    const salesAccountItemId = getSalesAccountItemId(accountItems); // 売上高のIDを取得
    const sections = fetchSections(token, companyId);
    const sectionId = getSectionIdByName(sections, TARGET_SECTION_NAME);
    // 取引取得
    const deals = fetchIncomeDeals(token, companyId, startDate, endDate);
    // データ抽出・整形・出力
    const sales = extractSalesBySection(deals, salesAccountItemId, sectionId);
    const formatted = formatSalesData(sales);
    writeToSpreadsheet(formatted);
    Logger.log('main処理が完了しました');
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    throw e;
  }
}