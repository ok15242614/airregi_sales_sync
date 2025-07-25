// =============================
// freee API × GAS 特定部門の売上一覧取得スクリプト（リファクタリング版）
// =============================

// --- 設定値 ---
const TARGET_SECTION_NAME = 'すなば文衛門'; // 対象部門名
const DATE_FORMAT = 'yyyy-MM-dd'; // 日付フォーマット
const DAYS_AGO = 0;
const CASH_ITEM_KEYWORD = '【Airレジ】001現金'; // 現金売上の品目に含まれるキーワード

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
  
  // 完全なJSONをログ出力
  if (json.deals && json.deals.length > 0) {
    // サンプルとして最初の取引の詳細情報をログ出力
    const firstDeal = json.deals[0];
    Logger.log(`【取引サンプル】最初の取引: id=${firstDeal.id}, issue_date=${firstDeal.issue_date}`);
    Logger.log(`【JSON】最初の取引詳細:\n${JSON.stringify(firstDeal, null, 2)}`);
    
    if (firstDeal.details && firstDeal.details.length > 0) {
      const firstDetail = firstDeal.details[0];
      Logger.log(`【詳細サンプル】最初の取引の最初の詳細: account_item_id=${firstDetail.account_item_id}, entry_side=${firstDetail.entry_side}, section_id=${firstDetail.section_id || 'なし'}`);
    }
  }
  
  // 追加ページがある場合は再帰的に取得
  const deals = json.deals || [];
  if (json.meta && json.meta.total > offset + deals.length && deals.length > 0) {
    Logger.log(`【ページング】次のページを取得 offset=${offset + limit}`);
    const nextDeals = fetchIncomeDeals(token, companyId, startDate, endDate, offset + limit, limit);
    return deals.concat(nextDeals);
  }
  
  return deals;
}

// --- freee API: 品目一覧取得 ---
function fetchItems(token, companyId) {
  Logger.log(`【APIリクエスト】品目一覧 company_id=${companyId}`);
  const url = `https://api.freee.co.jp/api/1/items?company_id=${companyId}`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('品目一覧取得失敗: ' + response.getContentText());
  }
  const json = JSON.parse(response.getContentText());
  Logger.log(`【APIレスポンス】品目件数: ${json.items ? json.items.length : 0}`);
  
  // 完全なJSONをログ出力
  if (json.items && json.items.length > 0) {
    const firstItem = json.items[0];
    Logger.log(`【品目サンプル】最初の品目: id=${firstItem.id}, name=${firstItem.name}`);
    Logger.log(`【JSON】最初の品目詳細:\n${JSON.stringify(firstItem, null, 2)}`);
    
    // 特定のキーワードを含む品目を検索してログ出力
    const matchingItems = json.items.filter(item => 
      item.name && item.name.includes(CASH_ITEM_KEYWORD));
    if (matchingItems.length > 0) {
      Logger.log(`【検索結果】"${CASH_ITEM_KEYWORD}"を含む品目が${matchingItems.length}件見つかりました`);
      matchingItems.forEach(item => {
        Logger.log(`【マッチング品目】id=${item.id}, name=${item.name}`);
      });
    } else {
      Logger.log(`【検索結果】"${CASH_ITEM_KEYWORD}"を含む品目は見つかりませんでした`);
    }
    
    // 全品目一覧をログ出力
    Logger.log(`【JSON】全品目一覧:\n${JSON.stringify(json.items, null, 2)}`);
  }
  
  return json.items || [];
}

// --- ユーティリティ: 売上高の勘定科目ID取得 ---
function getSalesAccountItemId(accountItems) {
  const item = accountItems.find(i => i.name === '売上高');
  if (!item) throw new Error('売上高の勘定科目が見つかりません');
  Logger.log(`【売上高ID特定】売上高 account_item_id=${item.id}`);
  return item.id;
}

// --- ユーティリティ: 現金の勘定科目ID取得 ---
function getCashAccountItemId(accountItems) {
  const item = accountItems.find(i => i.name === '現金');
  if (!item) throw new Error('現金の勘定科目が見つかりません');
  Logger.log(`【現金ID特定】現金 account_item_id=${item.id}`);
  return item.id;
}

// --- ユーティリティ: 指定部門ID取得 ---
function getSectionIdByName(sections, sectionName) {
  const section = sections.find(sec => sec.name && sec.name.trim().includes(sectionName.trim()));
  if (!section) throw new Error('指定した部門名が見つかりません: ' + sectionName);
  Logger.log(`【部門ID特定】${sectionName} section_id=${section.id}`);
  return section.id;
}

// --- データ抽出: 特定部門・売上高のみ（借方が現金のみ） ---
function extractSalesBySection(deals, salesAccountItemId, sectionId, cashAccountItemId) {
  Logger.log(`【データ抽出】売上高ID=${salesAccountItemId} 部門ID=${sectionId} 現金ID=${cashAccountItemId}`);
  Logger.log(`【データ抽出】売上高ID型=${typeof salesAccountItemId} 部門ID型=${typeof sectionId} 現金ID型=${typeof cashAccountItemId}`);
  
  // 型変換を確実に行う
  const salesAccountItemIdNum = Number(salesAccountItemId);
  const sectionIdNum = Number(sectionId);
  const cashAccountItemIdNum = Number(cashAccountItemId);
  
  const salesByDate = {}; // 日付ごとの売上を集計するためのオブジェクト
  let totalDetails = 0;
  
  deals.forEach(deal => {
    if (!deal.details) {
      Logger.log(`【警告】取引ID=${deal.id}の詳細がありません`);
      return;
    }
    
    totalDetails += deal.details.length;
    
    // 取引に部門IDが設定されているか確認
    const hasDealSection = deal.section_id !== undefined && deal.section_id !== null;
    const dealSectionId = hasDealSection ? Number(deal.section_id) : null;
    const dealSectionMatches = hasDealSection && dealSectionId === sectionIdNum;
    
    if (hasDealSection) {
      Logger.log(`【取引部門】取引ID=${deal.id} 部門ID=${dealSectionId} マッチ=${dealSectionMatches}`);
    }
    
    // この取引に現金の詳細があるか確認
    const hasCashDetail = deal.details.some(d => 
      Number(d.account_item_id) === cashAccountItemIdNum && d.entry_side === 'debit');
    
    if (!hasCashDetail) {
      Logger.log(`【現金なし】取引ID=${deal.id}に現金の借方がありません`);
      return; // 現金の借方がない取引はスキップ
    }
    
    deal.details.forEach(detail => {
      // 売上高の勘定科目IDかチェック
      const detailAccountItemId = Number(detail.account_item_id);
      const isAccountItemMatch = detailAccountItemId === salesAccountItemIdNum;
      
      // 貸方（収入）かチェック - freeeでは売上が貸方(credit)に記録されている
      const isCredit = detail.entry_side === 'credit';
      
      // 部門IDをチェック（詳細レベル、または取引レベル）
      let sectionMatches = false;
      
      // 詳細に部門IDがある場合はそれを優先
      if (detail.section_id !== undefined && detail.section_id !== null) {
        const detailSectionId = Number(detail.section_id);
        sectionMatches = detailSectionId === sectionIdNum;
        Logger.log(`【詳細部門】取引ID=${deal.id} 詳細部門ID=${detailSectionId} マッチ=${sectionMatches}`);
      } 
      // 詳細に部門IDがなく、取引に部門IDがある場合
      else if (hasDealSection) {
        sectionMatches = dealSectionMatches;
        Logger.log(`【取引部門採用】取引ID=${deal.id} 取引部門ID=${dealSectionId} マッチ=${sectionMatches}`);
      }
      // どちらにも部門IDがない場合
      else {
        Logger.log(`【部門なし】取引ID=${deal.id}の詳細と取引に部門指定がありません`);
        sectionMatches = false;
      }
      
      // デバッグ情報
      Logger.log(`【詳細】取引ID=${deal.id} account_item_id=${detail.account_item_id}(${isAccountItemMatch}), entry_side=${detail.entry_side}(${isCredit}), section=${sectionMatches}`);
      
      // 条件に一致する場合
      if (isAccountItemMatch && isCredit && sectionMatches) {
        Logger.log(`【一致】取引ID=${deal.id}の詳細がマッチしました`);
        
        // 日付ごとに金額を集計
        const date = deal.issue_date;
        if (!salesByDate[date]) {
          salesByDate[date] = {
            date: date,
            amount: 0,
            partner_name: '日計',
            description: `${date}の売上集計`,
            deals: [] // 集計対象の取引IDを記録
          };
        }
        
        salesByDate[date].amount += detail.amount;
        salesByDate[date].deals.push(deal.id);
        Logger.log(`【集計】日付=${date} 金額=${detail.amount} 累計=${salesByDate[date].amount}`);
      }
    });
  });
  
  // 日付ごとの集計結果を配列に変換
  const sales = Object.values(salesByDate);
  
  Logger.log(`【処理詳細】総取引数=${deals.length}, 総明細数=${totalDetails}`);
  Logger.log(`【抽出結果】売上明細件数: ${sales.length}`);
  
  if (sales.length > 0) {
    Logger.log(`【抽出サンプル】1件目: ${JSON.stringify(sales[0])}`);
  } else {
    // 抽出結果が0件の場合、参考情報を出力
    if (deals.length > 0 && deals[0].details && deals[0].details.length > 0) {
      const sampleDeal = deals[0];
      const sampleDetail = sampleDeal.details[0];
      Logger.log(`【参考】最初の取引: id=${sampleDeal.id}, section_id=${sampleDeal.section_id || 'なし'}`);
      Logger.log(`【参考】最初の取引の最初の詳細: account_item_id=${sampleDetail.account_item_id}, entry_side=${sampleDetail.entry_side}, section_id=${sampleDetail.section_id || 'なし'}`);
    }
  }
  return sales;
}

// --- データ抽出: 特定部門・売上高のみ（指定品目のもの） ---
function extractSalesByItem(deals, salesAccountItemId, sectionId, targetItemKeyword) {
  Logger.log(`【データ抽出】売上高ID=${salesAccountItemId} 部門ID=${sectionId} 品目キーワード="${targetItemKeyword}"`);
  Logger.log(`【データ抽出】売上高ID型=${typeof salesAccountItemId} 部門ID型=${typeof sectionId}`);
  
  // 型変換を確実に行う
  const salesAccountItemIdNum = Number(salesAccountItemId);
  const sectionIdNum = Number(sectionId);
  
  const salesByDate = {}; // 日付ごとの売上を集計するためのオブジェクト
  let totalDetails = 0;
  let matchingDetails = 0;
  let detailsWithItemInfo = [];  // 品目情報を含む詳細を保存
  
  deals.forEach(deal => {
    if (!deal.details) {
      Logger.log(`【警告】取引ID=${deal.id}の詳細がありません`);
      return;
    }
    
    totalDetails += deal.details.length;
    
    // 取引に部門IDが設定されているか確認
    const hasDealSection = deal.section_id !== undefined && deal.section_id !== null;
    const dealSectionId = hasDealSection ? Number(deal.section_id) : null;
    const dealSectionMatches = hasDealSection && dealSectionId === sectionIdNum;
    
    if (hasDealSection) {
      Logger.log(`【取引部門】取引ID=${deal.id} 部門ID=${dealSectionId} マッチ=${dealSectionMatches}`);
    }
    
    deal.details.forEach(detail => {
      // 売上高の勘定科目IDかチェック
      const detailAccountItemId = Number(detail.account_item_id);
      const isAccountItemMatch = detailAccountItemId === salesAccountItemIdNum;
      
      // 貸方（収入）かチェック - freeeでは売上が貸方(credit)に記録されている
      const isCredit = detail.entry_side === 'credit';
      
      // 部門IDをチェック（詳細レベル、または取引レベル）
      let sectionMatches = false;
      
      // 詳細に部門IDがある場合はそれを優先
      if (detail.section_id !== undefined && detail.section_id !== null) {
        const detailSectionId = Number(detail.section_id);
        sectionMatches = detailSectionId === sectionIdNum;
      } 
      // 詳細に部門IDがなく、取引に部門IDがある場合
      else if (hasDealSection) {
        sectionMatches = dealSectionMatches;
      }
      
      // 品目名をチェックして詳細情報を保存
      let itemMatches = false;
      let itemName = '';
      
      if (detail.item_id && detail.item) {
        itemName = detail.item.name || '';
        itemMatches = itemName.includes(targetItemKeyword);
        
        // 品目情報を含む詳細を保存
        if (detail.item) {
          detailsWithItemInfo.push({
            deal_id: deal.id,
            detail_id: detail.id,
            account_item_id: detail.account_item_id,
            entry_side: detail.entry_side,
            item: detail.item
          });
        }
        
        Logger.log(`【品目チェック】取引ID=${deal.id} 品目名=${itemName} マッチ=${itemMatches}`);
      } else if (detail.item_name) { // 一部のAPIレスポンスではitem_nameで直接返る場合も
        itemName = detail.item_name;
        itemMatches = itemName.includes(targetItemKeyword);
        
        // 品目名情報を含む詳細を保存
        detailsWithItemInfo.push({
          deal_id: deal.id,
          detail_id: detail.id,
          account_item_id: detail.account_item_id,
          entry_side: detail.entry_side,
          item_name: detail.item_name
        });
        
        Logger.log(`【品目チェック】取引ID=${deal.id} 品目名=${itemName} マッチ=${itemMatches}`);
      }
      
      // デバッグ情報
      Logger.log(`【詳細】取引ID=${deal.id} account_item_id=${detail.account_item_id}(${isAccountItemMatch}), entry_side=${detail.entry_side}(${isCredit}), section=${sectionMatches}, item=${itemMatches}`);
      
      // 条件に一致する場合
      if (isAccountItemMatch && isCredit && sectionMatches && itemMatches) {
        Logger.log(`【一致】取引ID=${deal.id}の詳細がマッチしました`);
        matchingDetails++;
        
        // 日付ごとに金額を集計
        const date = deal.issue_date;
        if (!salesByDate[date]) {
          salesByDate[date] = {
            date: date,
            amount: 0,
            partner_name: `${date}の${targetItemKeyword}売上`,
            description: `品目「${targetItemKeyword}」を含む売上`,
            deals: [] // 集計対象の取引IDを記録
          };
        }
        
        salesByDate[date].amount += detail.amount;
        if (!salesByDate[date].deals.includes(deal.id)) {
          salesByDate[date].deals.push(deal.id);
        }
        Logger.log(`【集計】日付=${date} 金額=${detail.amount} 累計=${salesByDate[date].amount}`);
      }
    });
  });
  
  // 品目情報をJSON形式でログ出力
  if (detailsWithItemInfo.length > 0) {
    Logger.log(`【JSON】品目情報を含む明細 (${detailsWithItemInfo.length}件):\n${JSON.stringify(detailsWithItemInfo, null, 2)}`);
  } else {
    Logger.log('【警告】品目情報を含む明細は見つかりませんでした');
  }
  
  // 日付ごとの集計結果を配列に変換
  const sales = Object.values(salesByDate);
  
  Logger.log(`【処理詳細】総取引数=${deals.length}, 総明細数=${totalDetails}, 一致明細=${matchingDetails}`);
  Logger.log(`【抽出結果】売上明細件数: ${sales.length}`);
  
  if (sales.length > 0) {
    Logger.log(`【抽出サンプル】1件目: ${JSON.stringify(sales[0])}`);
  } else {
    // 抽出結果が0件の場合、参考情報を出力
    if (deals.length > 0 && deals[0].details && deals[0].details.length > 0) {
      const sampleDeal = deals[0];
      const sampleDetail = sampleDeal.details[0];
      Logger.log(`【参考】最初の取引: id=${sampleDeal.id}, section_id=${sampleDeal.section_id || 'なし'}`);
      Logger.log(`【参考】品目情報: ${sampleDetail.item ? JSON.stringify(sampleDetail.item) : '品目情報なし'}`);
      Logger.log(`【参考】item_name: ${sampleDetail.item_name || '未設定'}`);
      
      // サンプル取引の詳細情報をJSON形式で出力
      Logger.log(`【JSON】サンプル取引の完全な情報:\n${JSON.stringify(sampleDeal, null, 2)}`);
    }
  }
  return sales;
}

// --- データ整形 ---
function formatSalesData(sales) {
  const headers = ['日付', '金額', '取引先名', 'メモ', '対象取引ID'];
  const rows = sales.map(sale => [
    sale.date,
    sale.amount,
    sale.partner_name,
    sale.description,
    sale.deals ? sale.deals.join(', ') : ''
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
    
    // 品目一覧取得
    try {
      const items = fetchItems(token, companyId);
      Logger.log(`【品目一覧取得成功】${items.length}件の品目があります`);
    } catch (e) {
      Logger.log(`【品目一覧取得エラー】${e.message}`);
    }
    
    // 取引取得
    const deals = fetchIncomeDeals(token, companyId, startDate, endDate);
    Logger.log(`【取引総数】${deals.length}件の取引を取得しました`);
    
    // データ抽出・整形・出力（品目で絞り込み）
    const sales = extractSalesByItem(deals, salesAccountItemId, sectionId, CASH_ITEM_KEYWORD);
    const formatted = formatSalesData(sales);
    writeToSpreadsheet(formatted);
    Logger.log('main処理が完了しました');
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    throw e;
  }
}