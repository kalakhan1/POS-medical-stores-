/*****************************************************************************
 * ═══════════════════════════════════════════════════════════════════════════
 *  HNK MEDICAL STORE - BACKEND
 *  Modular Architecture - Multi-Branch Medical Store
 *  A PROJECT BY HNK TECH & STUDIO
 * ═══════════════════════════════════════════════════════════════════════════
 *****************************************************************************/

/*****************************************************************************
 * SECTION 1: CONFIGURATION
 *****************************************************************************/
const CONFIG = {
  SS_ID: 'get from sheet url',
  ADMIN_USER: 'admin',
  ADMIN_PASS: 'admin123',
  COMPANY_NAME: 'HNK Medical Store',
  COMPANY_TYPE: 'medical',
  DEFAULT_CURRENCY: 'Rs.',
  PAGE_SIZE: 10,
  CACHE_DURATION: 300
};

/*****************************************************************************
 * SECTION 2: SPREADSHEET HELPERS
 *****************************************************************************/
function getSS() { return SpreadsheetApp.openById(CONFIG.SS_ID); }

function getSheet(name) {
  try {
    const ss = getSS();
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); initHeaders(sh, name); }
    return sh;
  } catch(e) { Logger.log('getSheet error: ' + e); return null; }
}

function initHeaders(sh, name) {
  const headers = {
    'Products': ['id','name','barcode','productType','packetPrice','itemsPerPacket','price','purchasePrice','salePrice','quantity','alertQty','unit','supplier','company','branch','created'],
    'Suppliers': ['id','name','contact','email','address','city','phone','ntn','custom1','custom2','custom3','created'],
    'Users': ['id','username','password','role','name','branch','created'],
    'Purchases': ['id','supplier','invoiceNo','items','total','paid','due','date','user','branch','notes','created'],
    'PurchaseLogs': ['id','supplier','invoiceNo','items','total','paid','due','date','user','branch','notes','created'],
    'Expenses': ['id','category','subcategory','description','amount','paidTo','date','user','branch','notes','created'],
    'Sales': ['id','type','user','shopName','location','items','total','discountType','discountValue','discountAmount','grandTotal','status','branch','created'],
    'Trash': ['id','_from','_deletedAt','data'],
    'Settings': ['storeName','contact','location','thankYou','footer','custom','theme','lang','branches']
  };
  if (headers[name]) sh.appendRow(headers[name]);
}

/*****************************************************************************
 * SECTION 3: DATA OPERATIONS
 *****************************************************************************/
function sheetToObjects(name) {
  try {
    const sh = getSheet(name);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    return data.slice(1).map(r => {
      const obj = {};
      headers.forEach((h,i) => {
        if ((h === 'items' || h === 'productType') && typeof r[i] === 'string' && r[i]) {
          try { obj[h] = JSON.parse(r[i]); } catch(e) { obj[h] = r[i]; }
        } else { obj[h] = r[i]; }
      });
      return obj;
    });
  } catch(e) { Logger.log('sheetToObjects error: ' + name); return []; }
}

function objectsToSheet(name, arr) {
  try {
    const sh = getSheet(name);
    if (!sh) return;
    sh.clear();
    if (!arr.length) return;
    const headers = Object.keys(arr[0]);
    sh.appendRow(headers);
    const rows = arr.map(r => headers.map(h => {
      const val = r[h];
      if ((h === 'items' || h === 'productType') && typeof val === 'object' && val !== null) return JSON.stringify(val);
      return val !== undefined && val !== null ? val : '';
    }));
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  } catch(e) { Logger.log('objectsToSheet error: ' + name); }
}

function appendRow(name, obj) {
  try {
    const sh = getSheet(name);
    if (!sh) return;
    if (sh.getLastRow() === 0) sh.appendRow(Object.keys(obj));
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => {
      const val = obj[h];
      if ((h === 'items' || h === 'productType') && typeof val === 'object' && val !== null) return JSON.stringify(val);
      return val !== undefined && val !== null ? val : '';
    });
    sh.appendRow(row);
  } catch(e) { Logger.log('appendRow error: ' + name); }
}

/*****************************************************************************
 * SECTION 4: UTILITIES & HELPERS
 *****************************************************************************/
function uid() { return 'ID' + new Date().getTime() + Math.floor(Math.random()*1000); }
function ts() { return new Date().toISOString(); }
function safeParseJSON(str, fallback) {
  try { return typeof str === 'string' ? JSON.parse(str) : str; } 
  catch(e) { return fallback !== undefined ? fallback : str; }
}
function normalizeString(str) { return (str || '').toString().trim().toLowerCase(); }
function sanitizeInput(str) { return (str || '').toString().trim().replace(/[<>]/g, ''); }

/*****************************************************************************
 * SECTION 5: AUTHENTICATION
 *****************************************************************************/
function login(username, password) {
  try {
    const validation = validateLoginCredentials(username, password);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    const users = sheetToObjects('Users');
    const u = users.find(x => x.username === username && x.password === password);
    if (u) return { ok: true, user: u };
    if (username === CONFIG.ADMIN_USER && password === CONFIG.ADMIN_PASS) {
      const admin = { id: uid(), username: CONFIG.ADMIN_USER, password: CONFIG.ADMIN_PASS, role: 'admin', name: 'Administrator', branch: 'Main', created: ts() };
      appendRow('Users', admin);
      return { ok: true, user: admin };
    }
    return { ok: false, msg: 'Invalid username or password' };
  } catch(e) { return { ok: false, msg: 'Server error: ' + e.message }; }
}

/*****************************************************************************
 * SECTION 6: AUTHORIZATION
 *****************************************************************************/
const ROLE_PERMISSIONS = {
  admin: ['all'],
  manager: ['view_inventory','view_purchaselogs','view_allsales','view_salesdetails','view_expenses','view_audit','view_lowstock','view_suppliers'],
  counter: ['create_sale','view_sale'],
  purchase: ['create_purchase','view_purchase','view_purchaselogs'],
  accountant: ['create_expense','view_expense','view_audit']
};
function checkPermission(userRole, action) {
  const perms = ROLE_PERMISSIONS[userRole] || [];
  return perms.includes('all') || perms.includes(action);
}

/*****************************************************************************
 * SECTION 7: VALIDATION
 *****************************************************************************/
function validateLoginCredentials(username, password) {
  if (!username || !password) return { valid: false, msg: 'Username and password required' };
  if (username.length < 3) return { valid: false, msg: 'Username too short' };
  if (password.length < 4) return { valid: false, msg: 'Password too short' };
  return { valid: true };
}

function validateProduct(data) {
  if (!data.name || !data.name.trim()) return { valid: false, msg: 'Product name required' };
  if (data.name.length > 100) return { valid: false, msg: 'Name too long' };
  if (data.quantity === undefined || data.quantity < 0) return { valid: false, msg: 'Invalid quantity' };
  if (data.productType === 'packet') {
    if (!data.packetPrice || data.packetPrice <= 0) return { valid: false, msg: 'Packet price required' };
    if (!data.itemsPerPacket || data.itemsPerPacket <= 0) return { valid: false, msg: 'Items per packet required' };
  } else if (data.productType === 'liquid') {
    if (!data.packetPrice || data.packetPrice <= 0) return { valid: false, msg: 'Bottle price required' };
    if (!data.itemsPerPacket || data.itemsPerPacket <= 0) return { valid: false, msg: 'ML per bottle required' };
  } else if (data.productType === 'single') {
    if (!data.salePrice || data.salePrice < 0) return { valid: false, msg: 'Sale price required' };
  }
  return { valid: true };
}

function validateUser(data) {
  if (!data.name || !data.name.trim()) return { valid: false, msg: 'Name required' };
  if (!data.username || !data.username.trim()) return { valid: false, msg: 'Username required' };
  if (!data.password || data.password.length < 4) return { valid: false, msg: 'Password min 4 chars' };
  if (!data.role) return { valid: false, msg: 'Role required' };
  return { valid: true };
}

function validateSale(data) {
  if (!data.items || !data.items.length) return { valid: false, msg: 'Cart empty' };
  if (!data.total || data.total <= 0) return { valid: false, msg: 'Invalid total' };
  if (data.discountType && data.discountType !== 'none') {
    if (!data.discountValue || data.discountValue < 0) return { valid: false, msg: 'Invalid discount value' };
    if (data.discountType === 'percent' && data.discountValue > 100) return { valid: false, msg: 'Discount cannot exceed 100%' };
  }
  return { valid: true };
}

function validatePurchase(data) {
  if (!data.supplier || !data.supplier.trim()) return { valid: false, msg: 'Supplier required' };
  if (!data.items || !data.items.length) return { valid: false, msg: 'Add items' };
  if (!data.total || data.total <= 0) return { valid: false, msg: 'Invalid total' };
  return { valid: true };
}

function validateExpense(data) {
  if (!data.category || !data.category.trim()) return { valid: false, msg: 'Category required' };
  if (!data.amount || data.amount <= 0) return { valid: false, msg: 'Amount must be positive' };
  if (!data.date) return { valid: false, msg: 'Date required' };
  return { valid: true };
}

function validateSupplier(data) {
  if (!data.name || !data.name.trim()) return { valid: false, msg: 'Supplier name required' };
  if (!data.contact || !data.contact.trim()) return { valid: false, msg: 'Contact number required' };
  return { valid: true };
}

/*****************************************************************************
 * SECTION 8: PRODUCTS MODULE
 *****************************************************************************/
function getProducts() { return sheetToObjects('Products'); }

function addProduct(p) {
  try {
    const validation = validateProduct(p);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    const all = getProducts();
    if (all.some(x => normalizeString(x.name) === normalizeString(p.name)))
      return { ok: false, msg: 'Duplicate product name' };
    p.id = p.id || uid();
    p.name = sanitizeInput(p.name);
    p.supplier = sanitizeInput(p.supplier || '');
    p.company = sanitizeInput(p.company || '');
    p.branch = sanitizeInput(p.branch || 'Main');
    p.created = ts();
    
    // Set salePrice and purchasePrice
    if (p.productType === 'packet' || p.productType === 'liquid') {
      p.price = calculatePerItemPrice(p.packetPrice, p.itemsPerPacket);
      if (!p.salePrice) p.salePrice = p.price;
    } else {
      if (!p.salePrice) p.salePrice = p.price || 0;
    }
    if (!p.purchasePrice) p.purchasePrice = 0;
    
    appendRow('Products', p);
    return { ok: true, id: p.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function updateProduct(id, data) {
  try {
    if (!id) return { ok: false, msg: 'Product ID required' };
    if (!data || Object.keys(data).length === 0) return { ok: false, msg: 'No data to update' };
    const all = getProducts();
    const i = all.findIndex(x => x.id === id);
    if (i < 0) return { ok: false, msg: 'Product not found' };
    const allowedFields = ['name','barcode','productType','packetPrice','itemsPerPacket','price','purchasePrice','salePrice','quantity','alertQty','unit','supplier','company','branch'];
    allowedFields.forEach(function(field) {
      if (data[field] !== undefined && data[field] !== null) {
        if (['name','barcode','unit','productType','supplier','company','branch'].indexOf(field) >= 0) {
          all[i][field] = String(data[field]).trim();
        } else {
          const num = Number(data[field]);
          if (!isNaN(num) && isFinite(num)) all[i][field] = num;
        }
      }
    });
    if (data.packetPrice !== undefined || data.itemsPerPacket !== undefined) {
      const pp = Number(all[i].packetPrice) || 0;
      const ip = Number(all[i].itemsPerPacket) || 0;
      if (pp > 0 && ip > 0) {
        all[i].price = calculatePerItemPrice(pp, ip);
        if (!all[i].salePrice) all[i].salePrice = all[i].price;
      }
    }
    objectsToSheet('Products', all);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function deleteEntity(table, id) {
  try {
    const all = sheetToObjects(table);
    const i = all.findIndex(x => x.id === id);
    if (i < 0) return { ok: false };
    const removed = all.splice(i,1)[0];
    objectsToSheet(table, all);
    appendRow('Trash', { id: uid(), _from: table, _deletedAt: ts(), data: JSON.stringify(removed) });
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 9: MEDICINE PRICING MODULE
 *****************************************************************************/
function calculatePerItemPrice(packetPrice, itemsPerPacket) {
  if (!packetPrice || !itemsPerPacket || itemsPerPacket <= 0) return 0;
  return Math.round((packetPrice / itemsPerPacket) * 100) / 100;
}

function calculateSalePrice(product, sellMode, quantity) {
  if (!product || !quantity || quantity <= 0) return 0;
  if (product.productType === 'single') return (Number(product.salePrice) || Number(product.price) || 0) * quantity;
  if (sellMode === 'packet') return (Number(product.packetPrice) || 0) * quantity;
  return calculatePerItemPrice(product.packetPrice, product.itemsPerPacket) * quantity;
}

/*****************************************************************************
 * SECTION 10: SALES MODULE
 *****************************************************************************/
function addSale(sale) {
  try {
    const validation = validateSale(sale);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    sale.id = uid();
    sale.type = 'counter';
    sale.created = ts();
    sale.branch = sale.branch || 'Main';
    
    if (!sale.discountType || sale.discountType === 'none') {
      sale.discountType = 'none';
      sale.discountValue = 0;
      sale.discountAmount = 0;
      sale.grandTotal = sale.total;
    } else {
      if (sale.discountType === 'percent') {
        sale.discountAmount = Math.round((sale.total * sale.discountValue / 100) * 100) / 100;
      } else if (sale.discountType === 'fixed') {
        sale.discountAmount = Math.min(Number(sale.discountValue) || 0, sale.total);
      }
      sale.grandTotal = Math.max(0, sale.total - sale.discountAmount);
    }
    
    appendRow('Sales', sale);
    
    const items = safeParseJSON(sale.items, []);
    (items || []).forEach(it => {
      const prods = getProducts();
      const p = prods.find(x => x.id === it.id);
      if (p) {
        let qtyToDeduct = 0;
        if (p.productType === 'packet' || p.productType === 'liquid') {
          if (it.sellMode === 'packet') qtyToDeduct = Number(it.qty) || 0;
          else qtyToDeduct = (Number(it.qty) || 0) / (Number(p.itemsPerPacket) || 1);
        } else {
          qtyToDeduct = Number(it.qty) || 0;
        }
        p.quantity = Math.max(0, (Number(p.quantity) || 0) - qtyToDeduct);
        updateProduct(p.id, { quantity: p.quantity });
      }
    });
    return { ok: true, id: sale.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function getAllSales() { return sheetToObjects('Sales'); }

/*****************************************************************************
 * SECTION 11: PURCHASE MODULE (NEW)
 *****************************************************************************/
function addPurchase(p) {
  try {
    const validation = validatePurchase(p);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    p.id = uid();
    p.created = ts();
    p.branch = p.branch || 'Main';
    p.paid = Number(p.paid) || 0;
    p.due = Math.max(0, (Number(p.total) || 0) - p.paid);
    
    appendRow('Purchases', p);
    appendRow('PurchaseLogs', p);
    
    // Update product stock
    const items = safeParseJSON(p.items, []);
    (items || []).forEach(it => {
      const prods = getProducts();
      const prod = prods.find(x => x.id === it.id);
      if (prod) {
        let qtyToAdd = Number(it.qty) || 0;
        if (prod.productType === 'packet' || prod.productType === 'liquid') {
          if (it.sellMode === 'individual') {
            qtyToAdd = qtyToAdd / (Number(prod.itemsPerPacket) || 1);
          }
        }
        prod.quantity = (Number(prod.quantity) || 0) + qtyToAdd;
        // Update purchase price if provided
        if (it.purchasePrice && it.purchasePrice > 0) {
          if (prod.productType === 'packet' || prod.productType === 'liquid') {
            if (it.sellMode === 'packet') {
              prod.purchasePrice = Number(it.purchasePrice);
            } else {
              prod.purchasePrice = Number(it.purchasePrice) * (Number(prod.itemsPerPacket) || 1);
            }
          } else {
            prod.purchasePrice = Number(it.purchasePrice);
          }
        }
        updateProduct(prod.id, { quantity: prod.quantity, purchasePrice: prod.purchasePrice });
      }
    });
    return { ok: true, id: p.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function getPurchases() { return sheetToObjects('Purchases'); }
function getPurchaseLogs() { return sheetToObjects('PurchaseLogs'); }

/*****************************************************************************
 * SECTION 12: EXPENSES MODULE (NEW)
 *****************************************************************************/
function addExpense(e) {
  try {
    const validation = validateExpense(e);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    e.id = uid();
    e.created = ts();
    e.branch = e.branch || 'Main';
    e.amount = Number(e.amount) || 0;
    e.category = sanitizeInput(e.category);
    e.subcategory = sanitizeInput(e.subcategory || '');
    e.description = sanitizeInput(e.description || '');
    e.paidTo = sanitizeInput(e.paidTo || '');
    e.notes = sanitizeInput(e.notes || '');
    appendRow('Expenses', e);
    return { ok: true, id: e.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function getExpenses() { return sheetToObjects('Expenses'); }

/*****************************************************************************
 * SECTION 13: AUDIT MODULE (NEW)
 *****************************************************************************/
function getAuditData(filters) {
  try {
    filters = filters || {};
    const from = filters.from ? new Date(filters.from) : null;
    const to = filters.to ? new Date(filters.to + 'T23:59:59') : null;
    const branch = filters.branch || '';
    
    const sales = getAllSales().filter(s => {
      const d = new Date(s.created);
      if (from && d < from) return false;
      if (to && d > to) return false;
      if (branch && s.branch !== branch) return false;
      return true;
    });
    
    const purchases = getPurchaseLogs().filter(p => {
      const d = new Date(p.date || p.created);
      if (from && d < from) return false;
      if (to && d > to) return false;
      if (branch && p.branch !== branch) return false;
      return true;
    });
    
    const expenses = getExpenses().filter(e => {
      const d = new Date(e.date || e.created);
      if (from && d < from) return false;
      if (to && d > to) return false;
      if (branch && e.branch !== branch) return false;
      return true;
    });
    
    const totalSales = sales.reduce((s, x) => s + (Number(x.grandTotal) || 0), 0);
    const totalDiscount = sales.reduce((s, x) => s + (Number(x.discountAmount) || 0), 0);
    const totalPurchase = purchases.reduce((s, x) => s + (Number(x.total) || 0), 0);
    const totalPaid = purchases.reduce((s, x) => s + (Number(x.paid) || 0), 0);
    const totalDue = purchases.reduce((s, x) => s + (Number(x.due) || 0), 0);
    const totalExpenses = expenses.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    const netProfit = totalSales - totalPurchase - totalExpenses;
    
    // Category-wise expenses
    const expByCategory = {};
    expenses.forEach(e => {
      const cat = e.category || 'Other';
      if (!expByCategory[cat]) expByCategory[cat] = 0;
      expByCategory[cat] += Number(e.amount) || 0;
    });
    
    return {
      ok: true,
      totalSales: totalSales,
      totalDiscount: totalDiscount,
      totalPurchase: totalPurchase,
      totalPaid: totalPaid,
      totalDue: totalDue,
      totalExpenses: totalExpenses,
      netProfit: netProfit,
      salesCount: sales.length,
      purchaseCount: purchases.length,
      expenseCount: expenses.length,
      expByCategory: expByCategory,
      sales: sales,
      purchases: purchases,
      expenses: expenses
    };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/*****************************************************************************
 * SECTION 14: USERS MODULE
 *****************************************************************************/
function getUsers() { return sheetToObjects('Users'); }
function addUser(u) {
  try {
    const validation = validateUser(u);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    const all = getUsers();
    if (all.some(x => normalizeString(x.username) === normalizeString(u.username)))
      return { ok: false, msg: 'Username already exists' };
    u.id = u.id || uid();
    u.name = sanitizeInput(u.name);
    u.username = sanitizeInput(u.username);
    u.branch = sanitizeInput(u.branch || 'Main');
    u.created = ts();
    appendRow('Users', u);
    return { ok: true, id: u.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 15: SUPPLIERS MODULE
 *****************************************************************************/
function getSuppliers() { return sheetToObjects('Suppliers'); }
function addSupplier(s) {
  try {
    const validation = validateSupplier(s);
    if (!validation.valid) return { ok: false, msg: validation.msg };
    const all = getSuppliers();
    if (all.some(x => normalizeString(x.name) === normalizeString(s.name)))
      return { ok: false, msg: 'Supplier name already exists' };
    s.id = s.id || uid();
    s.name = sanitizeInput(s.name);
    s.contact = sanitizeInput(s.contact || '');
    s.email = sanitizeInput(s.email || '');
    s.address = sanitizeInput(s.address || '');
    s.city = sanitizeInput(s.city || '');
    s.phone = sanitizeInput(s.phone || '');
    s.ntn = sanitizeInput(s.ntn || '');
    s.custom1 = sanitizeInput(s.custom1 || '');
    s.custom2 = sanitizeInput(s.custom2 || '');
    s.custom3 = sanitizeInput(s.custom3 || '');
    s.created = ts();
    appendRow('Suppliers', s);
    return { ok: true, id: s.id };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 16: TRASH MODULE
 *****************************************************************************/
function getTrash() { return sheetToObjects('Trash'); }
function restoreFromTrash(id) {
  try {
    const trash = getTrash();
    const i = trash.findIndex(x => x.id === id);
    if (i < 0) return { ok: false };
    const item = trash[i];
    const target = item._from || 'Products';
    const data = JSON.parse(item.data);
    appendRow(target, data);
    trash.splice(i,1);
    objectsToSheet('Trash', trash);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}
function permanentDelete(id) {
  try {
    const trash = getTrash();
    const i = trash.findIndex(x => x.id === id);
    if (i < 0) return { ok: false };
    trash.splice(i,1);
    objectsToSheet('Trash', trash);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 17: SETTINGS MODULE
 *****************************************************************************/
function getSettings() {
  try {
    const s = sheetToObjects('Settings');
    return s[0] || { 
      storeName: CONFIG.COMPANY_NAME, contact:'0300-1234567', location:'Main Branch',
      thankYou:'Get well soon! Thank you for choosing us.', footer:'A PROJECT BY HNK TECH & STUDIO',
      theme:'light', lang:'ur', branches:'Main'
    };
  } catch(e) { return { storeName: CONFIG.COMPANY_NAME, footer:'A PROJECT BY HNK TECH & STUDIO', theme:'light', lang:'ur', branches:'Main' }; }
}
function saveSettings(s) {
  try { objectsToSheet('Settings', [s]); return { ok: true }; } 
  catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 18: DATABASE MODULE
 *****************************************************************************/
function getFullDB() {
  return {
    Products: getProducts(), Suppliers: getSuppliers(), Users: getUsers(),
    Purchases: getPurchases(), PurchaseLogs: getPurchaseLogs(),
    Expenses: getExpenses(), Sales: getAllSales(), 
    Trash: getTrash(), Settings: [getSettings()]
  };
}
function uploadDB(data) {
  try {
    Object.keys(data).forEach(k => {
      if (data[k] && data[k].length) objectsToSheet(k, data[k]);
    });
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/*****************************************************************************
 * SECTION 18B: RESET ALL DATA (NEW)
 *****************************************************************************/
function resetAllData() {
  try {
    const sheetsToReset = ['Products', 'Suppliers', 'Users', 'Purchases', 'PurchaseLogs', 'Expenses', 'Sales', 'Trash'];
    
    sheetsToReset.forEach(function(sheetName) {
      const sh = getSheet(sheetName);
      if (sh) {
        sh.clear();
      }
    });
    
    // Re-create admin user
    const adminUser = {
      id: uid(),
      username: CONFIG.ADMIN_USER,
      password: CONFIG.ADMIN_PASS,
      role: 'admin',
      name: 'Administrator',
      branch: 'Main',
      created: ts()
    };
    appendRow('Users', adminUser);
    
    // Re-seed default data
    seedData();
    
    return { ok: true, msg: 'All data reset successfully. Admin user recreated.' };
  } catch(e) {
    return { ok: false, msg: 'Reset failed: ' + e.message };
  }
}

/*****************************************************************************
 * SECTION 19: SEED DATA
 *****************************************************************************/
function seedData() {
  try {
    if (getUsers().length === 0) {
      addUser({ username:'admin', password:'admin123', role:'admin', name:'Administrator', branch:'Main' });
      addUser({ username:'manager1', password:'m123', role:'manager', name:'Dr. Ali Manager', branch:'Main' });
      addUser({ username:'counter1', password:'c123', role:'counter', name:'Bilal Pharmacist', branch:'Main' });
      addUser({ username:'purchase1', password:'p123', role:'purchase', name:'Usman Purchase Manager', branch:'Main' });
      addUser({ username:'accountant1', password:'a123', role:'accountant', name:'Hamza Accountant', branch:'Main' });
    }
    if (getSuppliers().length === 0) {
      addSupplier({ name:'GSK Pakistan', contact:'0300-1111111', email:'info@gsk.com.pk', address:'Industrial Area', city:'Karachi', phone:'021-1111111', ntn:'1234567-8', custom1:'Multinational' });
      addSupplier({ name:'Getz Pharma', contact:'0300-2222222', email:'info@getzpharma.com', address:'Korangi Industrial', city:'Karachi', phone:'021-2222222', ntn:'2345678-9', custom1:'Local' });
      addSupplier({ name:'Searle Company', contact:'0300-3333333', email:'info@searle.com.pk', address:'S.I.T.E Area', city:'Karachi', phone:'021-3333333', ntn:'3456789-0', custom1:'Local' });
    }
    if (getProducts().length === 0) {
      addProduct({ name:'Panadol 500mg (Packet of 30)', productType:'packet', packetPrice:150, itemsPerPacket:30, purchasePrice:120, salePrice:150, quantity:100, alertQty:20, barcode:'MED001', unit:'tablet', supplier:'GSK Pakistan', company:'GSK', branch:'Main' });
      addProduct({ name:'Augmentin 625mg (Box of 20)', productType:'packet', packetPrice:850, itemsPerPacket:20, purchasePrice:700, salePrice:850, quantity:50, alertQty:10, barcode:'MED002', unit:'tablet', supplier:'GSK Pakistan', company:'GSK', branch:'Main' });
      addProduct({ name:'Brufen 400mg (Strip of 10)', productType:'packet', packetPrice:120, itemsPerPacket:10, purchasePrice:90, salePrice:120, quantity:200, alertQty:30, barcode:'MED003', unit:'tablet', supplier:'Getz Pharma', company:'Getz', branch:'Main' });
      addProduct({ name:'Brufen Syrup 60ml', productType:'liquid', packetPrice:250, itemsPerPacket:60, purchasePrice:200, salePrice:250, quantity:40, alertQty:10, barcode:'MED004', unit:'ml', supplier:'Getz Pharma', company:'Getz', branch:'Main' });
      addProduct({ name:'Surgical Mask (Single)', productType:'single', purchasePrice:10, salePrice:15, quantity:500, alertQty:50, barcode:'MED006', unit:'piece', branch:'Main' });
    }
    const s = getSettings();
    if (!s.storeName || s.storeName === 'My Store') {
      saveSettings({ storeName: CONFIG.COMPANY_NAME, contact:'0300-1234567', location:'Main Branch', thankYou:'Get well soon!', footer:'A PROJECT BY HNK TECH & STUDIO', theme:'light', lang:'ur', branches:'Main' });
    }
    return { ok: true };
  } catch(e) { Logger.log('seedData error: ' + e); return { ok: true }; }
}

/*****************************************************************************
 * SECTION 20: WEB ENTRY
 *****************************************************************************/
function doGet() { 
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.COMPANY_NAME + ' - Distribution System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

