const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// Prices stored in user data folder
const PRICES_PATH = path.join(app.getPath('userData'), 'b2b_prices.json');

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 960,
    minHeight: 640,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
    title: 'B2B Order Processor — CRM Tool',
    backgroundColor: '#F4F7F4',
  });
  win.loadFile('index.html');
  win.setMenuBarVisibility(false);
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());

// ── Load saved prices ─────────────────────────────────────────────────────────
ipcMain.handle('load-prices', () => {
  if (fs.existsSync(PRICES_PATH)) {
    try { return JSON.parse(fs.readFileSync(PRICES_PATH, 'utf8')); }
    catch { return null; }
  }
  return null;
});

// ── Save prices ───────────────────────────────────────────────────────────────
ipcMain.handle('save-prices', (_, products) => {
  fs.writeFileSync(PRICES_PATH, JSON.stringify(products, null, 2));
  return true;
});

// ── Browse for forms file ─────────────────────────────────────────────────────
ipcMain.handle('browse-file', async () => {
  const { filePaths } = await dialog.showOpenDialog(win, {
    title: 'Select Microsoft Forms Export',
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  });
  return filePaths[0] || null;
});

// ── Browse for price list upload ──────────────────────────────────────────────
ipcMain.handle('browse-price-list', async () => {
  const { filePaths } = await dialog.showOpenDialog(win, {
    title: 'Select Price List (Excel)',
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] }],
    properties: ['openFile'],
  });
  return filePaths[0] || null;
});

// ── Parse price list file ─────────────────────────────────────────────────────
ipcMain.handle('parse-price-list', (_, filePath) => {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  // Try to detect columns: look for headers
  // Expected: Item Code | Product Name | Price | MOQ (any order, any row)
  let headerRow = -1;
  let colCode = -1, colName = -1, colPrice = -1, colMoq = -1;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r].map(c => String(c).toLowerCase().trim());
    const codeIdx  = row.findIndex(c => c.includes('code') || c.includes('sku') || c.includes('item'));
    const nameIdx  = row.findIndex(c => c.includes('name') || c.includes('product') || c.includes('description'));
    const priceIdx = row.findIndex(c => c.includes('price') || c.includes('rate') || c.includes('unit'));
    const moqIdx   = row.findIndex(c => c.includes('moq') || c.includes('minimum'));
    if (nameIdx >= 0 && priceIdx >= 0) {
      headerRow = r;
      colCode  = codeIdx;
      colName  = nameIdx;
      colPrice = priceIdx;
      colMoq   = moqIdx;
      break;
    }
  }

  if (headerRow < 0) {
    // Fallback: assume col 0=code, 1=name, 2=price, 3=moq
    headerRow = 0;
    colCode = 0; colName = 1; colPrice = 2; colMoq = 3;
  }

  const products = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    const name  = String(row[colName] || '').trim();
    const price = parseFloat(String(row[colPrice] || '0').replace(/[^0-9.]/g, ''));
    const moq   = parseInt(String(row[colMoq] || '1').replace(/[^0-9]/g, '')) || 1;
    const code  = colCode >= 0 ? String(row[colCode] || '').trim() : '';
    if (name && price > 0) {
      products.push({ sku: code, name, price, moq });
    }
  }
  return products;
});

// ── Parse MS Forms export ─────────────────────────────────────────────────────
ipcMain.handle('parse-forms', (_, filePath, products) => {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rows.length < 2) return [];

  const headers = rows[0];
  const prodCols = headers.slice(11);  // columns 11+ are products

  const orders = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const company  = String(row[6] || '').trim();
    const contact  = String(row[7] || '').trim();
    const phone    = String(row[8] || '').trim();
    const location = String(row[9] || '').trim();
    let   rawDate  = String(row[10] || '').trim();

    if (!company) continue;

    // Parse Excel serial date
    const serial = parseFloat(rawDate);
    if (!isNaN(serial) && serial > 40000) {
      const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
      rawDate = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    }

    const items = [];
    for (let c = 0; c < Math.min(prodCols.length, products.length); c++) {
      const rawQty = String(row[11 + c] || '').trim();
      const qty = parseFloat(rawQty.replace(/[^0-9.]/g, '')) || 0;
      if (qty > 0) {
        const prod = products[c];
        items.push({
          sku:        prod.sku || '',
          name:       prod.name,
          qty:        Math.round(qty),
          unit_price: prod.price,
          moq:        prod.moq,
          net_amount: Math.round(qty * prod.price * 100) / 100,
        });
      }
    }

    const net_sales = items.reduce((s, it) => s + it.net_amount, 0);
    orders.push({ company, contact, phone, location, delivery_date: rawDate, items, net_sales: Math.round(net_sales * 100) / 100 });
  }
  return orders;
});

// ── Export B2B Excel ──────────────────────────────────────────────────────────
ipcMain.handle('export-excel', async (_, orders, reportDate) => {
  const { filePath } = await dialog.showSaveDialog(win, {
    title: 'Save B2B Order Sheet',
    defaultPath: `B2B_${reportDate.replace(/\//g,'_')}.xlsx`,
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (!filePath) return null;

  // Build worksheet data as array of arrays
  const aoa = [];

  // Row 0: date in col B (index 1)
  aoa.push(['', reportDate, '', '', '']);
  // Row 1: empty
  aoa.push(['', '', '', '', '']);

  let firstDataRow = null;

  for (const order of orders) {
    // Customer header row
    aoa.push(['', order.company.toUpperCase(), 'QUANTITY', 'UNIT PRICE', 'NET AMOUNT']);

    // Item rows
    for (const item of order.items) {
      if (firstDataRow === null) firstDataRow = aoa.length + 1; // 1-based
      aoa.push([item.sku, item.name.toUpperCase(), item.qty, item.unit_price, { f: `D${aoa.length + 1}*C${aoa.length + 1}` }]);
    }
    // Blank separator
    aoa.push(['', '', '', '', { f: `D${aoa.length + 1}*C${aoa.length + 1}` }]);
  }

  // Grand total
  const lastRow = aoa.length;
  aoa.push(['', '', '', '', { f: `SUM(E${firstDataRow || 4}:E${lastRow})` }]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Column widths
  ws['!cols'] = [
    { wch: 14 },   // A - Item Code
    { wch: 52 },   // B - Product / Company
    { wch: 14 },   // C - Quantity
    { wch: 16 },   // D - Unit Price
    { wch: 18 },   // E - Net Amount
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, filePath);
  return filePath;
});
