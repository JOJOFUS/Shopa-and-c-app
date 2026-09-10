const STOCK_SHEET = 'Stock';
const USERS_SHEET = 'Users';
const HISTORY_SHEET = 'History';

function doPost(e) {
  try {
    const p = e.parameter || {};
    let result;
    switch (p.action) {
      case 'login':
        result = handleLogin_(p);
        break;
      case 'updateQty':
        result = handleUpdateQty_(p);
        break;
      case 'transferStock':
        result = handleTransfer_(p);
        break;
      default:
        throw new Error('Unknown or missing action: ' + p.action);
    }
    return jsonOut_(Object.assign({ status: 'success' }, result));
  } catch (err) {
    return jsonOut_({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action !== 'getHistory') {
      throw new Error('Unknown or missing action: ' + p.action);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hist = ss.getSheetByName(HISTORY_SHEET);
    if (!hist) return jsonOut_({ status: 'success', rows: [] });
    const values = hist.getDataRange().getValues();
    const rows = values.slice(1).map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      user: r[1], action: r[2], brand: r[3], size: r[4],
      shop: r[5], toShop: r[6], change: r[7],
      newQtyFrom: r[8], newQtyTo: r[9]
    })).reverse();
    return jsonOut_({ status: 'success', rows: rows });
  } catch (err) {
    return jsonOut_({ status: 'error', message: err.toString() });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet tab not found: "' + name + '". Available tabs: ' +
      ss.getSheets().map(s => s.getName()).join(', '));
  }
  return sheet;
}

function handleLogin_(p) {
  const username = (p.username || '').trim();
  const password = (p.password || '').trim();
  if (!username || !password) throw new Error('Username and password required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const userCol = headers.indexOf('username');
  const passCol = headers.indexOf('password');
  if (userCol === -1 || passCol === -1) {
    throw new Error('Users sheet needs Username and Password columns. Found: ' + headers.join(', '));
  }

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][userCol]).trim() === username && String(values[i][passCol]).trim() === password) {
      return { user: username };
    }
  }
  throw new Error('Invalid username or password');
}

function handleUpdateQty_(p) {
  const user = (p.user || '').trim();
  const shop = (p.shop || '').trim();
  const brand = (p.brand || '').trim();
  const size = (p.size || '').trim();
  const qty = parseInt(p.qty, 10);
  const delta = p.delta !== undefined ? parseInt(p.delta, 10) : 0;

  if (!user || !shop || !brand || !size || isNaN(qty)) {
    throw new Error('Missing or invalid fields: ' + JSON.stringify(p));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, STOCK_SHEET);
  const found = findStockRow_(sheet, shop, brand, size);

  if (found.rowIndex === -1) {
    const newRow = [];
    newRow[found.cols.shopCol] = shop;
    newRow[found.cols.brandCol] = brand;
    newRow[found.cols.sizeCol] = size;
    newRow[found.cols.qtyCol] = qty;
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(found.rowIndex + 1, found.cols.qtyCol + 1).setValue(qty);
  }

  logHistory_(ss, {
    user: user, action: 'Update', brand: brand, size: size,
    shop: shop, toShop: '', change: delta,
    newQtyFrom: qty, newQtyTo: ''
  });

  return { action: found.rowIndex === -1 ? 'inserted' : 'updated' };
}

function handleTransfer_(p) {
  const user = (p.user || '').trim();
  const fromShop = (p.fromShop || '').trim();
  const toShop = (p.toShop || '').trim();
  const brand = (p.brand || '').trim();
  const size = (p.size || '').trim();
  const amount = parseInt(p.qty, 10);

  if (!user || !fromShop || !toShop || !brand || !size || isNaN(amount) || amount <= 0) {
    throw new Error('Missing or invalid fields: ' + JSON.stringify(p));
  }
  if (fromShop === toShop) throw new Error('From and To shop must be different');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss, STOCK_SHEET);

  const from = findStockRow_(sheet, fromShop, brand, size);
  const currentFromQty = from.rowIndex === -1 ? 0 :
    (Number(sheet.getRange(from.rowIndex + 1, from.cols.qtyCol + 1).getValue()) || 0);

  if (from.rowIndex === -1 || currentFromQty < amount) {
    throw new Error('Not enough stock at ' + fromShop + ' (has ' + currentFromQty + ', tried to move ' + amount + ')');
  }

  const newFromQty = currentFromQty - amount;
  sheet.getRange(from.rowIndex + 1, from.cols.qtyCol + 1).setValue(newFromQty);

  const to = findStockRow_(sheet, toShop, brand, size);
  let newToQty;
  if (to.rowIndex === -1) {
    newToQty = amount;
    const newRow = [];
    newRow[to.cols.shopCol] = toShop;
    newRow[to.cols.brandCol] = brand;
    newRow[to.cols.sizeCol] = size;
    newRow[to.cols.qtyCol] = newToQty;
    sheet.appendRow(newRow);
  } else {
    const currentToQty = Number(sheet.getRange(to.rowIndex + 1, to.cols.qtyCol + 1).getValue()) || 0;
    newToQty = currentToQty + amount;
    sheet.getRange(to.rowIndex + 1, to.cols.qtyCol + 1).setValue(newToQty);
  }

  logHistory_(ss, {
    user: user, action: 'Transfer', brand: brand, size: size,
    shop: fromShop, toShop: toShop, change: amount,
    newQtyFrom: newFromQty, newQtyTo: newToQty
  });

  return { action: 'transferred' };
}

function findStockRow_(sheet, shop, brand, size) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const shopCol = headers.indexOf('shop');
  const brandCol = headers.indexOf('brand');
  const sizeCol = headers.indexOf('size');
  const qtyCol = headers.indexOf('qty');
  if (shopCol === -1 || brandCol === -1 || sizeCol === -1 || qtyCol === -1) {
    throw new Error('Stock sheet needs Shop, Brand, Size, Qty columns. Found: ' + headers.join(', '));
  }
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][shopCol]).trim() === shop &&
        String(values[i][brandCol]).trim() === brand &&
        String(values[i][sizeCol]).trim() === size) {
      rowIndex = i;
      break;
    }
  }
  return { rowIndex: rowIndex, cols: { shopCol: shopCol, brandCol: brandCol, sizeCol: sizeCol, qtyCol: qtyCol } };
}

function logHistory_(ss, entry) {
  let hist = ss.getSheetByName(HISTORY_SHEET);
  if (!hist) {
    hist = ss.insertSheet(HISTORY_SHEET);
    hist.appendRow(['Timestamp', 'User', 'Action', 'Brand', 'Size', 'Shop', 'ToShop', 'Change', 'NewQtyFrom', 'NewQtyTo']);
  }
  hist.appendRow([new Date(), entry.user, entry.action, entry.brand, entry.size,
    entry.shop, entry.toShop, entry.change, entry.newQtyFrom, entry.newQtyTo]);
}
