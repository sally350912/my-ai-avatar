const SPREADSHEET_ID = '17YWF_Bmppap9RyTutVWTsysl1qC5vl0JPmbU6FGKujQ';
const TZ = 'Asia/Taipei';

const SHEET = {
  PRODUCTS: '產品建檔',
  CUSTOMERS: '客戶建檔',
  INVENTORY: '庫存總覽',
  PURCHASE: '進貨單紀錄',
  QUOTE: '報價單紀錄',
  SALES: '銷售訂單紀錄',
  AP: '應付帳款',
  AR: '應收帳款',
  RECONCILIATION: '帳務對帳總表'
};

// ===== 郵件設定 =====
const ADMIN_EMAIL = 'sally350912@gmail.com';
const SUPPLIER_EMAIL = 'pinxinselect@gmail.com';

// ===== 郵件發送函數 =====
function sendPurchaseEmail(docId, items) {
  try {
    const subject = '品新文創進貨點收通知 - ' + docId;

    let emailBody = '團隊您好：\n\n';
    emailBody += '品新文創已成系統完成進貨點收，明細如下：\n\n';
    emailBody += '進貨單號：' + docId + '\n';

    items.forEach(function(item, index) {
      emailBody += '\n商品 ' + (index + 1) + '：\n';
      emailBody += '產品名稱：' + item.product + '\n';
      emailBody += '產品批號：' + (item.batch || '無') + '\n';
      emailBody += '點收數量：' + item.qty + '\n';
    });

    emailBody += '\n此為品新文創存系統自動發送之通知信件，謝謝！\n\n';
    emailBody += '品新文創 敬上';

    MailApp.sendEmail(ADMIN_EMAIL, subject, emailBody);
    MailApp.sendEmail(SUPPLIER_EMAIL, subject, emailBody);

    return { success: true, message: '通知信已發送' };
  } catch(e) {
    Logger.log('郵件發送失敗：' + e.toString());
    return { success: false, message: '郵件發送失敗：' + e.toString() };
  }
}

function doGet() {
  initializeSheets();
  return HtmlService.createHtmlOutputFromFile('Index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('品新文創 ERP 進階版');
}

function initializeSheets() {
  const ss = getSS();
  const requiredSheets = [
    { name: '應付帳款', headers: ['單據日期', '來源單號', '供應商', '單據總額', '已沖金額', '未沖餘額', '帳齡(天)', '狀態', '發票號碼'] },
    { name: '應收帳款', headers: ['單據日期', '來源單號', '客戶名稱', '單據總額', '已收金額', '未收餘額', '帳齡(天)', '狀態', '發票號碼'] },
    { name: '帳務對帳總表', headers: ['帳務類型', '單據日期', '來源單號', '客戶/供應商', '單據總額', '已收沖金額', '未收餘額', '帳齡(天)', '狀態'] }
  ];

  requiredSheets.forEach(function(sheetConfig) {
    if (!ss.getSheetByName(sheetConfig.name)) {
      ss.insertSheet(sheetConfig.name);
      const sheet = ss.getSheetByName(sheetConfig.name);
      sheet.getRange(1, 1, 1, sheetConfig.headers.length).setValues([sheetConfig.headers]);
    }
  });
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheetOrThrow(sheetName) {
  const sheet = getSS().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('找不到工作表：' + sheetName);
  }
  return sheet;
}

function normalizeText(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function normalizeNumber(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function makeDocId(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
}

function getNextDocId(prefix) {
  return makeDocId(prefix);
}

// ===== AP/AR 函數 =====
function getProductPrice(productName, priceTier) {
  const ss = getSS();
  const productSheet = ss.getSheetByName(SHEET.PRODUCTS);
  if (!productSheet || productSheet.getLastRow() <= 1) {
    return 0;
  }

  const rows = productSheet.getRange(2, 5, productSheet.getLastRow() - 1, 10).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (normalizeText(rows[i][0]) === normalizeText(productName)) {
      const prices = {
        '原廠進貨成本': normalizeNumber(rows[i][1]),
        '品新對外成本': normalizeNumber(rows[i][2]),
        '品新銷售價': normalizeNumber(rows[i][3]),
        '經銷商團購價': normalizeNumber(rows[i][4]),
        '經銷商拿貨價': normalizeNumber(rows[i][5])
      };
      return prices[normalizeText(priceTier)] || 0;
    }
  }
  return 0;
}

function getCustomerPriceTier(customerName) {
  const ss = getSS();
  const customerSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  if (!customerSheet || customerSheet.getLastRow() <= 1) {
    return '散客';
  }

  const rows = customerSheet.getRange(2, 2, customerSheet.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (normalizeText(rows[i][0]) === normalizeText(customerName)) {
      const tier = normalizeText(rows[i][7]);
      return tier || '散客';
    }
  }
  return '散客';
}

function saveAPRecord(docId, items, docDate) {
  try {
    const sheet = getSheetOrThrow(SHEET.AP);
    const rows = items.map(function(item) {
      const unitPrice = getProductPrice(item.product, '原廠進貨成本');
      const totalAmount = unitPrice * item.qty;
      return [
        docDate,
        normalizeText(docId),
        SUPPLIER_EMAIL,
        totalAmount,
        0,
        totalAmount,
        0,
        '待付款',
        ''
      ];
    });

    rows.forEach(function(row) {
      sheet.appendRow(row);
    });

    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    Logger.log('AP 記錄失敗：' + e.toString());
    return false;
  }
}

function saveARRecord(docId, items, customerName, docDate) {
  try {
    const sheet = getSheetOrThrow(SHEET.AR);

    items.forEach(function(item) {
      const totalAmount = normalizeNumber(item.total);
      sheet.appendRow([
        docDate,
        normalizeText(docId),
        normalizeText(customerName),
        totalAmount,
        0,
        totalAmount,
        0,
        '待收款',
        ''
      ]);
    });

    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    Logger.log('AR 記錄失敗：' + e.toString());
    return false;
  }
}

function getERPData() {
  const ss = getSS();
  const result = {
    customers: [],
    products: [],
    inventory: {},
    nextIds: {
      inId: makeDocId('IN'),
      quoteId: makeDocId('QT'),
      saleId: makeDocId('SO')
    }
  };

  const inventorySheet = ss.getSheetByName(SHEET.INVENTORY);
  if (inventorySheet && inventorySheet.getLastRow() > 1) {
    const rows = inventorySheet.getRange(2, 1, inventorySheet.getLastRow() - 1, Math.min(4, inventorySheet.getLastColumn())).getValues();
    rows.forEach(function(row) {
      const productName = normalizeText(row[0]);
      if (productName) {
        result.inventory[productName] = normalizeNumber(row[3]);
      }
    });
  }

  const productSheet = ss.getSheetByName(SHEET.PRODUCTS);
  if (productSheet && productSheet.getLastRow() > 1) {
    const rows = productSheet.getRange(2, 5, productSheet.getLastRow() - 1, 6).getValues();
    result.products = rows
      .filter(function(row) { return normalizeText(row[0]); })
      .map(function(row) {
        const name = normalizeText(row[0]);
        return {
          name: name,
          factoryCost: normalizeNumber(row[1]),
          companyCost: normalizeNumber(row[2]),
          salePrice: normalizeNumber(row[3]),
          distributorGroupPrice: normalizeNumber(row[4]),
          distributorPickupPrice: normalizeNumber(row[5]),
          buyPrice: normalizeNumber(row[1]),
          sellPrice: normalizeNumber(row[3]),
          stock: normalizeNumber(result.inventory[name])
        };
      });
  }

  const customerSheet = ss.getSheetByName(SHEET.CUSTOMERS);
  if (customerSheet && customerSheet.getLastRow() > 1) {
    const rows = customerSheet.getRange(2, 2, customerSheet.getLastRow() - 1, 8).getValues();
    result.customers = rows
      .filter(function(row) { return normalizeText(row[0]); })
      .map(function(row) {
        return {
          name: normalizeText(row[0]),
          priceTier: normalizeText(row[7])
        };
      });
  }

  return result;
}

function buildSalesLikeRow(now, docId, status, item, meta) {
  return [
    now,
    normalizeText(docId),
    normalizeText(status),
    normalizeText(meta.customer),
    normalizeText(item.product),
    normalizeNumber(item.qty),
    normalizeNumber(item.unitPrice),
    normalizeNumber(item.total),
    normalizeText(meta.salesperson),
    normalizeText(meta.contact),
    normalizeText(meta.phone)
  ];
}

function buildPurchaseRow(now, docId, item) {
  return [
    now,
    normalizeText(docId),
    normalizeText(item.product),
    normalizeText(item.factoryId),
    normalizeText(item.batch),
    normalizeText(item.expiry),
    normalizeNumber(item.qty),
    normalizeNumber(item.unitPrice),
    normalizeNumber(item.total)
  ];
}

function validateDocId(docId) {
  if (!normalizeText(docId)) {
    throw new Error('單號未產生，請重新整理頁面後再試。');
  }
}

function validateSalesMeta(meta, docId) {
  validateDocId(docId);
  if (!normalizeText(meta.customer)) {
    throw new Error('請先選擇客戶。');
  }
}

function validateLineItems(items) {
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('沒有可儲存的明細資料。');
  }

  const validItems = items
    .map(function(item) {
      return {
        product: normalizeText(item.product),
        qty: normalizeNumber(item.qty),
        unitPrice: normalizeNumber(item.unitPrice),
        total: normalizeNumber(item.total),
        factoryId: normalizeText(item.factoryId),
        batch: normalizeText(item.batch),
        expiry: normalizeText(item.expiry)
      };
    })
    .filter(function(item) {
      return item.product && item.qty > 0;
    });

  if (validItems.length === 0) {
    throw new Error('請至少輸入一筆完整資料，且數量必須大於 0。');
  }

  validItems.forEach(function(item, index) {
    if (item.unitPrice < 0) {
      throw new Error('第 ' + (index + 1) + ' 筆資料的單價不可小於 0。');
    }
  });

  return validItems;
}

function saveQuote(payload) {
  payload = payload || {};
  const meta = payload.meta || {};
  const docId = normalizeText(payload.docId);
  validateSalesMeta(meta, docId);
  const items = validateLineItems(payload.items || []);
  const sheet = getSheetOrThrow(SHEET.QUOTE);
  const now = new Date();
  const rows = items.map(function(item) {
    return buildSalesLikeRow(now, docId, '報價中', item, meta);
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: '報價單已儲存：' + docId,
    docId: docId,
    nextDocId: makeDocId('QT')
  };
}

function saveSale(payload) {
  payload = payload || {};
  const meta = payload.meta || {};
  const docId = normalizeText(payload.docId);
  validateSalesMeta(meta, docId);
  const items = validateLineItems(payload.items || []);
  const sheet = getSheetOrThrow(SHEET.SALES);
  const now = new Date();
  const rows = items.map(function(item) {
    return buildSalesLikeRow(now, docId, '待出貨', item, meta);
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  const sourceQuoteId = normalizeText(payload.sourceQuoteId);
  if (sourceQuoteId) {
    updateQuoteStatus(sourceQuoteId, '已轉單');
  }

  SpreadsheetApp.flush();

  // ===== 建立 AR 記錄 =====
  saveARRecord(docId, items, meta.customer, now);

  // 庫存已改為公式計算，不再自動更新

  return {
    ok: true,
    message: '銷售訂單已建立：' + docId,
    docId: docId,
    nextDocId: makeDocId('SO')
  };
}

function savePurchase(payload) {
  payload = payload || {};
  const docId = normalizeText(payload.docId);
  validateDocId(docId);
  const items = validateLineItems(payload.items || []);
  const sheet = getSheetOrThrow(SHEET.PURCHASE);
  const now = new Date();
  const rows = items.map(function(item) {
    return buildPurchaseRow(now, docId, item);
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  // ===== 發送郵件通知 =====
  const emailResult = sendPurchaseEmail(docId, items);

  // ===== 建立 AP 記錄 =====
  saveAPRecord(docId, items, now);

  // 庫存已改為公式計算，不再自動更新

  return {
    ok: true,
    message: '進貨單已儲存：' + docId + ' | ' + emailResult.message,
    docId: docId,
    nextDocId: makeDocId('IN'),
    emailSent: emailResult.success
  };
}

function getQuoteData(quoteId) {
  const qid = normalizeText(quoteId);
  if (!qid) {
    throw new Error('請輸入報價單號。');
  }

  const sheet = getSheetOrThrow(SHEET.QUOTE);
  const rows = sheet.getDataRange().getValues();
  const matched = rows.filter(function(row, index) {
    return index > 0 && normalizeText(row[1]) === qid;
  });

  if (matched.length === 0) {
    return {
      found: false,
      docId: qid,
      meta: {},
      items: []
    };
  }

  return {
    found: true,
    docId: qid,
    meta: {
      customer: normalizeText(matched[0][3]),
      salesperson: normalizeText(matched[0][8]),
      contact: normalizeText(matched[0][9]),
      phone: normalizeText(matched[0][10])
    },
    items: matched.map(function(row) {
      return {
        product: normalizeText(row[4]),
        qty: normalizeNumber(row[5]),
        unitPrice: normalizeNumber(row[6]),
        total: normalizeNumber(row[7])
      };
    })
  };
}

function updateQuoteStatus(quoteId, newStatus) {
  const qid = normalizeText(quoteId);
  if (!qid) return false;

  const sheet = getSheetOrThrow(SHEET.QUOTE);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const values = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
  let updated = false;

  values.forEach(function(row, idx) {
    if (normalizeText(row[0]) === qid) {
      sheet.getRange(idx + 2, 3).setValue(normalizeText(newStatus));
      updated = true;
    }
  });

  return updated;
}

function getAPData() {
  try {
    const sheet = getSheetOrThrow(SHEET.AP);
    const lastRow = sheet.getLastRow();
    Logger.log('AP Sheet Last Row: ' + lastRow);

    if (lastRow <= 1) {
      Logger.log('AP Sheet is empty (only headers)');
      return [];
    }

    const rows = sheet.getDataRange().getValues();
    Logger.log('AP Rows count: ' + rows.length);
    const result = [];
    const today = new Date();

    for (let i = 1; i < rows.length; i++) {
      const docDate = rows[i][0];
      const dateStr = docDate instanceof Date
        ? Utilities.formatDate(docDate, TZ, 'yyyy-MM-dd')
        : String(docDate);
      const agingDays = Math.floor((today - new Date(docDate)) / (1000 * 60 * 60 * 24));

      result.push({
        date: dateStr,
        docId: normalizeText(rows[i][1]),
        supplier: normalizeText(rows[i][2]),
        totalAmount: normalizeNumber(rows[i][3]),
        paidAmount: normalizeNumber(rows[i][4]),
        balanceAmount: normalizeNumber(rows[i][5]),
        agingDays: agingDays,
        status: normalizeText(rows[i][7]),
        invoiceNo: normalizeText(rows[i][8]),
        rowIndex: i + 1
      });
    }

    Logger.log('AP Records returned: ' + result.length);
    return result;
  } catch(e) {
    Logger.log('讀取 AP 失敗：' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return [];
  }
}

function getARData() {
  try {
    const sheet = getSheetOrThrow(SHEET.AR);
    const lastRow = sheet.getLastRow();
    Logger.log('AR Sheet Last Row: ' + lastRow);

    if (lastRow <= 1) {
      Logger.log('AR Sheet is empty (only headers)');
      return [];
    }

    const rows = sheet.getDataRange().getValues();
    Logger.log('AR Rows count: ' + rows.length);
    const result = [];
    const today = new Date();

    for (let i = 1; i < rows.length; i++) {
      const docDate = rows[i][0];
      const dateStr = docDate instanceof Date
        ? Utilities.formatDate(docDate, TZ, 'yyyy-MM-dd')
        : String(docDate);
      const agingDays = Math.floor((today - new Date(docDate)) / (1000 * 60 * 60 * 24));

      result.push({
        date: dateStr,
        docId: normalizeText(rows[i][1]),
        customer: normalizeText(rows[i][2]),
        totalAmount: normalizeNumber(rows[i][3]),
        receivedAmount: normalizeNumber(rows[i][4]),
        balanceAmount: normalizeNumber(rows[i][5]),
        agingDays: agingDays,
        status: normalizeText(rows[i][7]),
        invoiceNo: normalizeText(rows[i][8]),
        rowIndex: i + 1
      });
    }

    Logger.log('AR Records returned: ' + result.length);
    return result;
  } catch(e) {
    Logger.log('讀取 AR 失敗：' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return [];
  }
}

function updateAPRecord(rowIndex, paidAmount, status, invoiceNo) {
  try {
    const sheet = getSheetOrThrow(SHEET.AP);
    const row = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
    const totalAmount = normalizeNumber(row[3]);
    const balanceAmount = totalAmount - normalizeNumber(paidAmount);

    sheet.getRange(rowIndex, 5).setValue(normalizeNumber(paidAmount));
    sheet.getRange(rowIndex, 6).setValue(balanceAmount);
    sheet.getRange(rowIndex, 8).setValue(normalizeText(status));
    sheet.getRange(rowIndex, 9).setValue(normalizeText(invoiceNo));

    SpreadsheetApp.flush();
    return { ok: true, message: 'AP 記錄已更新' };
  } catch(e) {
    Logger.log('更新 AP 失敗：' + e.toString());
    return { ok: false, message: '更新失敗：' + e.toString() };
  }
}

function updateARRecord(rowIndex, receivedAmount, status, invoiceNo) {
  try {
    const sheet = getSheetOrThrow(SHEET.AR);
    const row = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
    const totalAmount = normalizeNumber(row[3]);
    const balanceAmount = totalAmount - normalizeNumber(receivedAmount);

    sheet.getRange(rowIndex, 5).setValue(normalizeNumber(receivedAmount));
    sheet.getRange(rowIndex, 6).setValue(balanceAmount);
    sheet.getRange(rowIndex, 8).setValue(normalizeText(status));
    sheet.getRange(rowIndex, 9).setValue(normalizeText(invoiceNo));

    SpreadsheetApp.flush();
    return { ok: true, message: 'AR 記錄已更新' };
  } catch(e) {
    Logger.log('更新 AR 失敗：' + e.toString());
    return { ok: false, message: '更新失敗：' + e.toString() };
  }
}

function getReconciliationData() {
  try {
    const result = {
      ap: [],
      ar: [],
      totalAP: 0,
      totalAR: 0
    };

    const apData = getAPData();
    apData.forEach(function(record) {
      result.ap.push({
        type: '應付',
        date: record.date,
        docId: record.docId,
        party: record.supplier,
        totalAmount: record.totalAmount,
        paidAmount: record.paidAmount,
        balanceAmount: record.balanceAmount,
        agingDays: record.agingDays,
        status: record.status
      });
      result.totalAP += record.balanceAmount;
    });

    const arData = getARData();
    arData.forEach(function(record) {
      result.ar.push({
        type: '應收',
        date: record.date,
        docId: record.docId,
        party: record.customer,
        totalAmount: record.totalAmount,
        receivedAmount: record.receivedAmount,
        balanceAmount: record.balanceAmount,
        agingDays: record.agingDays,
        status: record.status
      });
      result.totalAR += record.balanceAmount;
    });

    return result;
  } catch(e) {
    Logger.log('讀取對帳數據失敗：' + e.toString());
    return { ap: [], ar: [], totalAP: 0, totalAR: 0 };
  }
}

function getDebugInfo() {
  try {
    const ss = getSS();
    const sheetNames = ss.getSheetNames();
    const debug = {
      spreadsheetId: SPREADSHEET_ID,
      allSheets: sheetNames,
      apSheetExists: sheetNames.includes(SHEET.AP),
      arSheetExists: sheetNames.includes(SHEET.AR),
      apLastRow: 0,
      arLastRow: 0,
      apData: null,
      arData: null
    };

    try {
      const apSheet = ss.getSheetByName(SHEET.AP);
      if (apSheet) {
        debug.apLastRow = apSheet.getLastRow();
        debug.apHeaders = apSheet.getLastRow() > 0 ? apSheet.getRange(1, 1, 1, 9).getValues()[0] : [];
        if (apSheet.getLastRow() > 1) {
          const rows = apSheet.getRange(2, 1, Math.min(3, apSheet.getLastRow() - 1), 9).getValues();
          debug.apSampleRows = rows;
        }
      }
    } catch(e) {
      debug.apError = e.toString();
    }

    try {
      const arSheet = ss.getSheetByName(SHEET.AR);
      if (arSheet) {
        debug.arLastRow = arSheet.getLastRow();
        debug.arHeaders = arSheet.getLastRow() > 0 ? arSheet.getRange(1, 1, 1, 9).getValues()[0] : [];
        if (arSheet.getLastRow() > 1) {
          const rows = arSheet.getRange(2, 1, Math.min(3, arSheet.getLastRow() - 1), 9).getValues();
          debug.arSampleRows = rows;
        }
      }
    } catch(e) {
      debug.arError = e.toString();
    }

    Logger.log(JSON.stringify(debug));
    return debug;
  } catch(e) {
    return { error: e.toString() };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
