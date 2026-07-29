/**
 * ZIPDAM Backend - Google Apps Script Web App
 *
 * Canonical customer identity:
 *   customerId === lineUserId (LINE user ID beginning with "U")
 * Same LINE ID is always one customer, even if displayName changes.
 *
 * Endpoints:
 *   GET  ?action=catalog
 *   GET  ?action=health
 *   POST {action:"order", idToken, lineUserId, displayName, store, area, address, phone, cart:[...]}
 *   POST {action:"me", idToken, lineUserId?}
 *   POST {action:"customer_profile|customer_profile_set", ...}
 *   POST {action:"customer_summary", idToken, lineUserId?}
 *   POST {action:"favorites_get|favorites_add|favorites_remove", ...}
 *   POST {action:"templates_get|templates_add|templates_delete", ...}
 *   POST {action:"frequent_get", idToken, lineUserId?, limit}
 */

const ZIPDAM_SCHEMA = Object.freeze({
  Product: ['SKU', 'Brand', 'Size', 'Name', 'mm', 'pack', 'price', 'promo_price', 'image_key', 'active', 'Cost'],
  Customer: ['lineUserId', 'customerId', 'name', 'displayName', 'type', 'store', 'storeId', 'area', 'phone', 'defaultAddress', 'createdAt', 'lastSeenAt', 'note', 'status', 'linkedAt', 'loyaltyNote'],
  Orders: ['OrderID', 'CreatedAt', 'lineUserId', 'displayName', 'customerId', 'store', 'itemsTotal', 'shippingFee', 'grandTotal', 'status', 'address', 'phone', 'note', 'createdByLineUserId', 'orderMode', 'loyaltyStatus', 'pointsEarned', 'rewardApplied'],
  OrderItems: ['OrderID', 'SKU', 'Brand', 'Size', 'Name', 'qty', 'unitPrice', 'lineTotal', 'Profit', 'Cost'],
  Favorites: ['lineUserId', 'SKU', 'Brand', 'Size', 'Name', 'createdAt', 'updatedAt'],
  Templates: ['templateId', 'lineUserId', 'templateName', 'itemsJson', 'createdAt', 'updatedAt', 'lastUsedAt', 'note'],
  Rewards: ['RewardID', 'RewardName', 'RequiredSpend', 'RequiredPoints', 'RewardType', 'RewardValue', 'Active', 'StartDate', 'EndDate', 'Note']
});

function getConfig_() {
  const props = PropertiesService.getScriptProperties();

  return {
    SHEET_ID: props.getProperty('SHEET_ID') || '11sUClcToFNjQXafrZUgRaLGZW3NO2-cAfSIxe2IgsoE',
    LINE_LOGIN_CHANNEL_ID: props.getProperty('LINE_LOGIN_CHANNEL_ID') || '2008727011',
    LINE_MESSAGING_TOKEN: props.getProperty('LINE_MESSAGING_TOKEN') || '',
    N8N_WEBHOOK_URL: props.getProperty('N8N_WEBHOOK_URL') || '',
    ADMIN_LINE_USER_ID: props.getProperty('ADMIN_LINE_USER_ID') || '',
    FIXED_SHIPPING: numberProperty_(props, 'FIXED_SHIPPING', 20),
    LOW_ORDER_SHIPPING: numberProperty_(props, 'LOW_ORDER_SHIPPING', 30),
    SHIPPING_THRESHOLD: numberProperty_(props, 'SHIPPING_THRESHOLD', 200),
    ALLOW_GUEST_ORDERS: boolProperty_(props, 'ALLOW_GUEST_ORDERS', false),
    ALLOW_LINE_ID_MISMATCH: boolProperty_(props, 'ALLOW_LINE_ID_MISMATCH', false)
  };
}

function numberProperty_(props, key, fallback) {
  const raw = props.getProperty(key);
  if (raw === null || raw === '') return fallback;
  const value = Number(raw);
  return Number.isFinite(value) ? value : fallback;
}

function boolProperty_(props, key, fallback) {
  const raw = props.getProperty(key);
  if (raw === null || raw === '') return fallback;
  return String(raw).toLowerCase() === 'true';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = String(e?.parameter?.action || '');
    if (action === 'catalog') return json_({ ok: true, catalog: getCatalog_() });
    if (action === 'health') return json_(getHealth_());
    return json_({ ok: false, error: 'Unknown GET action. Use catalog or health.' });
  } catch (err) {
    return json_({ ok: false, error: errorMessage_(err) });
  }
}

function doPost(e) {
  try {
    const body = parseJsonBody_(e);
    const action = String(body.action || '');

    switch (action) {
      case 'order': return json_(handleOrder_(body));
      case 'me': return json_(handleMe_(body));
      case 'customer_profile': return json_(handleCustomerProfileGet_(body));
      case 'customer_profile_set': return json_(handleCustomerProfileSet_(body));
      case 'customer_summary': return json_(handleCustomerSummary_(body));
      case 'favorites_get': return json_(handleFavoritesGet_(body));
      case 'favorites_add': return json_(handleFavoritesAdd_(body));
      case 'favorites_remove': return json_(handleFavoritesRemove_(body));
      case 'templates_get': return json_(handleTemplatesGet_(body));
      case 'templates_add': return json_(handleTemplatesAdd_(body));
      case 'templates_delete': return json_(handleTemplatesDelete_(body));
      case 'frequent_get': return json_(handleFrequentGet_(body));
      default: return json_({ ok: false, error: 'Unknown POST action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: errorMessage_(err) });
  }
}

function parseJsonBody_(e) {
  if (!e?.postData?.contents) return {};
  try {
    return JSON.parse(e.postData.contents);
  } catch (_) {
    throw new Error('Invalid JSON body');
  }
}

function errorMessage_(err) {
  return String(err?.message || err || 'Unknown error');
}

function assert_(condition, message) {
  if (!condition) throw new Error(message);
}

function text_(value) {
  return String(value ?? '').trim();
}

function toNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function nowIso_() {
  return new Date().toISOString();
}

function isRealLineUserId_(value) {
  return /^U[0-9a-f]{32}$/i.test(text_(value));
}

function formatTHB_(value) {
  return '฿' + Number(value || 0).toLocaleString('th-TH');
}

/* =========================
 * LINE IDENTITY
 * ========================= */

function verifyLineIdToken_(idToken) {
  const { LINE_LOGIN_CHANNEL_ID } = getConfig_();
  assert_(text_(idToken), 'Missing idToken');

  const response = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      id_token: idToken,
      client_id: LINE_LOGIN_CHANNEL_ID
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const raw = response.getContentText();
  let data = {};
  try { data = JSON.parse(raw); } catch (_) {}

  if (status !== 200) {
    throw new Error('LINE token verify failed: ' + (data.error_description || data.error || raw));
  }

  assert_(data.sub, 'LINE token is valid but sub is missing');
  return data;
}

function resolveIdentity_(body, options) {
  const settings = Object.assign({
    requireLine: false,
    allowGuest: false
  }, options || {});

  const config = getConfig_();
  const providedLineUserId = text_(body.lineUserId);
  const providedDisplayName = text_(body.displayName);

  const suppliedIdToken = text_(body.idToken);
  const verified = suppliedIdToken ? verifyLineIdToken_(suppliedIdToken) : null;
  const tokenVerified = Boolean(verified);

  const verifiedLineUserId = isRealLineUserId_(verified?.sub) ? text_(verified.sub) : '';
  const bodyLineUserId = isRealLineUserId_(providedLineUserId) ? providedLineUserId : '';

  if (
    verifiedLineUserId &&
    bodyLineUserId &&
    verifiedLineUserId !== bodyLineUserId &&
    !config.ALLOW_LINE_ID_MISMATCH
  ) {
    throw new Error('LINE identity mismatch');
  }

  if (providedLineUserId && !bodyLineUserId) {
    throw new Error('Invalid LINE user ID');
  }

  if (verifiedLineUserId) {
    const resolvedLineUserId =
      bodyLineUserId &&
      bodyLineUserId !== verifiedLineUserId &&
      config.ALLOW_LINE_ID_MISMATCH
        ? bodyLineUserId
        : verifiedLineUserId;
    return {
      lineUserId: resolvedLineUserId,
      customerId: resolvedLineUserId,
      displayName: text_(verified?.name) || providedDisplayName || 'LINE customer',
      tokenVerified,
      tokenError: '',
      isGuest: false
    };
  }

  if (bodyLineUserId) {
    throw new Error('LINE authentication required. Please reopen this page inside LINE and log in again');
  }

  if (settings.requireLine || !settings.allowGuest || !config.ALLOW_GUEST_ORDERS) {
    throw new Error('Please open this page inside LINE and log in again');
  }

  const guestId = 'GUEST-' + Date.now();
  return {
    lineUserId: guestId,
    customerId: '',
    displayName: providedDisplayName || 'Guest',
    tokenVerified,
    tokenError: '',
    isGuest: true
  };
}

/* =========================
 * SHEET HELPERS
 * ========================= */

function openDb_() {
  return SpreadsheetApp.openById(getConfig_().SHEET_ID);
}

function getSheet_(name) {
  const sheet = openDb_().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function getHeaderInfo_(sheet) {
  const lastColumn = sheet.getLastColumn();
  assert_(lastColumn > 0, 'Sheet has no columns: ' + sheet.getName());

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  const duplicates = [];

  headers.forEach((header, index) => {
    const key = text_(header);
    if (!key) return;
    if (map[key]) {
      if (!duplicates.includes(key)) duplicates.push(key);
      return;
    }
    map[key] = index + 1;
  });

  return { map, duplicates };
}

function getHeaderMap_(sheet) {
  return getHeaderInfo_(sheet).map;
}

function assertHeaders_(sheetName, requiredHeaders) {
  const sheet = getSheet_(sheetName);
  const { map, duplicates } = getHeaderInfo_(sheet);
  const missing = requiredHeaders.filter(header => !map[header]);
  if (missing.length) {
    throw new Error(`${sheetName} missing columns: ${missing.join(', ')}`);
  }
  const duplicateRequired = duplicates.filter(header => requiredHeaders.includes(header));
  if (duplicateRequired.length) {
    throw new Error(`${sheetName} has duplicate required columns: ${duplicateRequired.join(', ')}`);
  }
  return { sheet, map };
}

function objectToRow_(sheet, headerMap, data) {
  const row = new Array(sheet.getLastColumn()).fill('');
  Object.keys(data).forEach(key => {
    const column = headerMap[key];
    if (column) row[column - 1] = data[key];
  });
  return row;
}

function appendObjectRow_(sheetName, data, requiredHeaders) {
  const { sheet, map } = assertHeaders_(sheetName, requiredHeaders || []);
  const row = objectToRow_(sheet, map, data);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function appendObjectRows_(sheetName, dataRows, requiredHeaders) {
  if (!dataRows.length) return;
  const { sheet, map } = assertHeaders_(sheetName, requiredHeaders || []);
  const rows = dataRows.map(data => objectToRow_(sheet, map, data));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, sheet.getLastColumn()).setValues(rows);
}

function findRowByValue_(sheet, column, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const finder = sheet
    .getRange(2, column, lastRow - 1, 1)
    .createTextFinder(String(value))
    .matchEntireCell(true)
    .findNext();

  return finder ? finder.getRow() : 0;
}

function updateObjectRow_(sheet, rowNumber, headerMap, data) {
  Object.keys(data).forEach(key => {
    const column = headerMap[key];
    if (!column) return;
    sheet.getRange(rowNumber, column).setValue(data[key]);
  });
}

function getHealth_() {
  const checks = {};
  Object.keys(ZIPDAM_SCHEMA).forEach(sheetName => {
    try {
      const sheet = getSheet_(sheetName);
      const { map, duplicates } = getHeaderInfo_(sheet);
      const missing = ZIPDAM_SCHEMA[sheetName].filter(header => !map[header]);
      const duplicateRequired = duplicates.filter(header => ZIPDAM_SCHEMA[sheetName].includes(header));
      checks[sheetName] = {
        ok: missing.length === 0 && duplicateRequired.length === 0,
        missing,
        duplicateRequired,
        warnings: duplicates
          .filter(header => !duplicateRequired.includes(header))
          .map(header => `Duplicate non-required header: ${header}`),
        headers: Object.keys(map)
      };
    } catch (err) {
      checks[sheetName] = { ok: false, error: errorMessage_(err) };
    }
  });

  return {
    ok: Object.values(checks).every(check => check.ok),
    customerKey: 'LINE userId',
    checks
  };
}

/* =========================
 * CATALOG
 * ========================= */

function getCatalog_() {
  const { sheet, map } = assertHeaders_('Product', ['SKU', 'Brand', 'Size', 'Name', 'price']);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return rows
    .map(row => {
      const value = header => map[header] ? row[map[header] - 1] : '';
      const active = map.active ? value('active') : true;
      if (active === false || text_(active).toLowerCase() === 'false') return null;

      const regularPrice = toNumber_(value('price'));
      const promoPrice = toNumber_(value('promo_price'));

      return {
        SKU: text_(value('SKU')),
        Brand: text_(value('Brand')),
        Feature: [],
        Size: text_(value('Size')),
        Name: text_(value('Name')),
        mm: value('mm'),
        pack: value('pack'),
        price: regularPrice,
        promo_price: promoPrice > 0 ? promoPrice : null,
        final_price: promoPrice > 0 ? promoPrice : regularPrice,
        image_key: text_(value('image_key'))
      };
    })
    .filter(product => product && product.Brand && product.Name);
}

function buildProductIndex_() {
  const { sheet, map } = assertHeaders_('Product', ['SKU', 'Brand', 'Size', 'Name', 'price']);
  const lastRow = sheet.getLastRow();
  const bySku = new Map();
  const byKey = new Map();

  if (lastRow < 2) return { bySku, byKey };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  rows.forEach(row => {
    const value = header => map[header] ? row[map[header] - 1] : '';
    const active = map.active ? value('active') : true;
    if (active === false || text_(active).toLowerCase() === 'false') return;

    const regularPrice = toNumber_(value('price'));
    const promoPrice = toNumber_(value('promo_price'));
    const product = {
      SKU: text_(value('SKU')),
      Brand: text_(value('Brand')),
      Size: text_(value('Size')),
      Name: text_(value('Name')),
      pack: value('pack'),
      finalPrice: promoPrice > 0 ? promoPrice : regularPrice,
      unitCost: toNumber_(value('Cost'))
    };

    if (!product.Brand || !product.Name) return;
    if (product.SKU) bySku.set(product.SKU, product);
    byKey.set(productKey_(product.Brand, product.Size, product.Name), product);
  });

  return { bySku, byKey };
}

function productKey_(brand, size, name) {
  return [brand, size, name].map(value => text_(value).toLowerCase()).join('||');
}

function resolveProduct_(item, index) {
  const sku = text_(item.SKU || item.id);
  const brand = text_(item.Brand || item.brand);
  const size = text_(item.Size || item.size);
  const name = text_(item.Name || item.name);

  let product = sku ? index.bySku.get(sku) : null;
  if (!product) product = index.byKey.get(productKey_(brand, size, name));

  if (!product && brand && name) {
    for (const candidate of index.bySku.values()) {
      if (
        candidate.Brand.toLowerCase() === brand.toLowerCase() &&
        candidate.Name.toLowerCase() === name.toLowerCase()
      ) {
        product = candidate;
        break;
      }
    }
  }

  assert_(product, `Product not found: ${sku || brand} / ${size} / ${name}`);
  return product;
}

function getShippingFee_(itemsTotal) {
  const config = getConfig_();
  if (itemsTotal <= 0) return 0;
  return itemsTotal < config.SHIPPING_THRESHOLD
    ? config.LOW_ORDER_SHIPPING
    : config.FIXED_SHIPPING;
}

/* =========================
 * ORDERS
 * ========================= */

function handleOrder_(body) {
  const cart = Array.isArray(body.cart) ? body.cart : [];
  assert_(cart.length > 0, 'Cart is empty');

  const identity = resolveIdentity_(body, { allowGuest: true });
  const profile = identity.isGuest ? emptyProfile_() : getCustomerProfile_(identity.lineUserId);

  const store = text_(body.store || body.storeName) || profile.store;
  const area = text_(body.area || body.soi) || profile.area;
  const address = text_(body.address || body.defaultAddress) || profile.defaultAddress;
  const phone = normalisePhone_(body.phone) || profile.phone;

  if (!identity.isGuest) {
    upsertCustomer_({
      lineUserId: identity.lineUserId,
      displayName: identity.displayName,
      store,
      area,
      phone,
      defaultAddress: address
    });
  }

  const productIndex = buildProductIndex_();
  const items = cart.map(item => {
    const quantity = Math.floor(toNumber_(item.qty || item.quantity));
    assert_(quantity > 0, 'qty must be > 0');

    const product = resolveProduct_(item, productIndex);
    const unitPrice = product.finalPrice;
    assert_(unitPrice > 0, `Product price is unavailable: ${product.SKU || product.Name}`);
    const lineTotal = Math.round(unitPrice * quantity);
    const totalCost = Math.round(product.unitCost * quantity * 100) / 100;
    const profit = Math.round((lineTotal - totalCost) * 100) / 100;

    return {
      SKU: product.SKU,
      Brand: product.Brand,
      Size: product.Size,
      Name: product.Name,
      pack: product.pack,
      qty: quantity,
      unitPrice,
      lineTotal,
      Cost: totalCost,
      Profit: profit
    };
  });

  const itemsTotal = Math.round(items.reduce((sum, item) => sum + item.lineTotal, 0));
  const shippingFee = getShippingFee_(itemsTotal);
  const grandTotal = itemsTotal + shippingFee;
  const orderId = getNextOrderId_();
  const createdAt = nowIso_();
  const orderMode = ['SELF', 'ADMIN', 'LEGACY'].includes(text_(body.orderMode).toUpperCase())
    ? text_(body.orderMode).toUpperCase()
    : 'SELF';

  appendOrder_({
    OrderID: orderId,
    CreatedAt: createdAt,
    lineUserId: identity.lineUserId,
    displayName: identity.displayName,
    customerId: identity.customerId,
    store,
    itemsTotal,
    shippingFee,
    grandTotal,
    status: 'CONFIRMED',
    address,
    phone,
    note: text_(body.note),
    createdByLineUserId: identity.lineUserId,
    orderMode,
    loyaltyStatus: identity.isGuest ? 'EXCLUDED' : 'PENDING',
    pointsEarned: '',
    rewardApplied: ''
  });

  appendOrderItems_(orderId, items);

  const config = getConfig_();
  const linePush = safePushOrderConfirmation_(
    config.LINE_MESSAGING_TOKEN,
    identity.lineUserId,
    orderId,
    items,
    itemsTotal,
    shippingFee,
    grandTotal,
    identity.displayName,
    address,
    phone
  );

  const adminPush = config.ADMIN_LINE_USER_ID && config.ADMIN_LINE_USER_ID !== identity.lineUserId
    ? safePushOrderConfirmation_(
        config.LINE_MESSAGING_TOKEN,
        config.ADMIN_LINE_USER_ID,
        orderId,
        items,
        itemsTotal,
        shippingFee,
        grandTotal,
        identity.displayName,
        address,
        phone
      )
    : { attempted: false, ok: false };

  const n8nPush = sendN8NWebhook_(config.N8N_WEBHOOK_URL, {
    orderId,
    createdAt,
    lineUserId: identity.lineUserId,
    customerId: identity.customerId,
    displayName: identity.displayName,
    store,
    area,
    address,
    phone,
    itemsTotal,
    shippingFee,
    grandTotal,
    orderMode,
    cart: items,
    tokenVerified: identity.tokenVerified,
    tokenError: identity.tokenError
  });

  return {
    ok: true,
    orderId,
    lineUserId: identity.lineUserId,
    customerId: identity.customerId,
    displayName: identity.displayName,
    store,
    area,
    address,
    phone,
    itemsTotal,
    shippingFee,
    grandTotal,
    tokenVerified: identity.tokenVerified,
    tokenError: identity.tokenError,
    linePush,
    adminPush,
    n8nPush
  };
}

function appendOrder_(order) {
  appendObjectRow_('Orders', order, ZIPDAM_SCHEMA.Orders);
}

function appendOrderItems_(orderId, items) {
  const rows = items.map(item => ({
    OrderID: orderId,
    SKU: item.SKU,
    Brand: item.Brand,
    Size: item.Size,
    Name: item.Name,
    qty: item.qty,
    unitPrice: item.unitPrice,
    lineTotal: item.lineTotal,
    Profit: item.Profit,
    Cost: item.Cost
  }));

  appendObjectRows_('OrderItems', rows, ZIPDAM_SCHEMA.OrderItems);
}

function getNextOrderId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    let last = toNumber_(props.getProperty('LAST_ORDER_NO'));
    if (last <= 0) last = scanMaxOrderNo_();

    const next = last + 1;
    props.setProperty('LAST_ORDER_NO', String(next));
    return 'OD' + String(next).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

function scanMaxOrderNo_() {
  const { sheet, map } = assertHeaders_('Orders', ['OrderID']);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, map.OrderID, lastRow - 1, 1).getDisplayValues().flat();
  return values.reduce((max, value) => {
    const match = text_(value).match(/^OD0*(\d+)$/i);
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);
}

/* =========================
 * CUSTOMER
 * ========================= */

function upsertCustomer_(profile) {
  assert_(isRealLineUserId_(profile.lineUserId), 'Invalid LINE user ID');

  const { sheet, map } = assertHeaders_('Customer', ['lineUserId', 'customerId']);
  const rowNumber = findRowByValue_(sheet, map.lineUserId, profile.lineUserId);
  const now = nowIso_();

  const data = {
    lineUserId: profile.lineUserId,
    customerId: profile.lineUserId,
    name: profile.displayName,
    displayName: profile.displayName,
    lastSeenAt: now,
    status: 'ACTIVE'
  };

  if (rowNumber) {
    const contactUpdates = {
      store: text_(profile.store),
      area: text_(profile.area),
      phone: normalisePhone_(profile.phone),
      defaultAddress: text_(profile.defaultAddress)
    };
    Object.keys(contactUpdates).forEach(key => {
      if (contactUpdates[key]) data[key] = contactUpdates[key];
    });
    updateObjectRow_(sheet, rowNumber, map, data);
    return rowNumber;
  }

  data.store = text_(profile.store);
  data.area = text_(profile.area);
  data.phone = normalisePhone_(profile.phone);
  data.defaultAddress = text_(profile.defaultAddress);
  data.createdAt = now;
  data.linkedAt = now;
  const row = objectToRow_(sheet, map, data);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return sheet.getLastRow();
}

function getCustomerProfile_(lineUserId) {
  if (!isRealLineUserId_(lineUserId)) return emptyProfile_();

  const { sheet, map } = assertHeaders_('Customer', ['lineUserId']);
  const rowNumber = findRowByValue_(sheet, map.lineUserId, lineUserId);
  if (!rowNumber) return emptyProfile_();

  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const value = header => map[header] ? row[map[header] - 1] : '';

  return {
    lineUserId,
    customerId: text_(value('customerId')) || lineUserId,
    displayName: text_(value('displayName') || value('name')),
    store: text_(value('store')),
    area: text_(value('area')),
    phone: normalisePhone_(value('phone')),
    defaultAddress: text_(value('defaultAddress')),
    address: text_(value('defaultAddress'))
  };
}

function emptyProfile_() {
  return {
    lineUserId: '',
    customerId: '',
    displayName: '',
    store: '',
    area: '',
    phone: '',
    defaultAddress: '',
    address: ''
  };
}

function normalisePhone_(value) {
  const raw = text_(value);
  if (!raw) return '';

  const digits = raw.replace(/\D/g, '');
  if (digits.length === 9 && !digits.startsWith('0')) return '0' + digits;
  return digits || raw;
}

function handleMe_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });

  upsertCustomer_({
    lineUserId: identity.lineUserId,
    displayName: identity.displayName,
    store: '',
    area: '',
    phone: '',
    defaultAddress: ''
  });

  return {
    ok: true,
    lineUserId: identity.lineUserId,
    customerId: identity.lineUserId,
    displayName: identity.displayName,
    profile: getCustomerProfile_(identity.lineUserId)
  };
}

function handleCustomerProfileGet_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  return {
    ok: true,
    lineUserId: identity.lineUserId,
    profile: getCustomerProfile_(identity.lineUserId)
  };
}

function handleCustomerProfileSet_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });

  upsertCustomer_({
    lineUserId: identity.lineUserId,
    displayName: text_(body.displayName) || identity.displayName,
    store: text_(body.store || body.storeName),
    area: text_(body.area || body.soi),
    phone: normalisePhone_(body.phone),
    defaultAddress: text_(body.address || body.defaultAddress)
  });

  return {
    ok: true,
    lineUserId: identity.lineUserId,
    customerId: identity.lineUserId,
    profile: getCustomerProfile_(identity.lineUserId)
  };
}

function handleCustomerSummary_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const summary = calculateCustomerSummary_(identity.lineUserId);
  return { ok: true, summary };
}

function calculateCustomerSummary_(lineUserId) {
  const { sheet, map } = assertHeaders_('Orders', [
    'OrderID',
    'lineUserId',
    'itemsTotal',
    'shippingFee',
    'grandTotal',
    'status'
  ]);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      customerId: lineUserId,
      orderCount: 0,
      productTotal: 0,
      shippingTotal: 0,
      lifetimeSpend: 0,
      rewards: []
    };
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let orderCount = 0;
  let productTotal = 0;
  let shippingTotal = 0;
  let lifetimeSpend = 0;

  rows.forEach(row => {
    if (text_(row[map.lineUserId - 1]) !== lineUserId) return;

    const status = text_(row[map.status - 1]).toUpperCase();
    if (['CANCELLED', 'CANCELED', 'VOID'].includes(status)) return;

    orderCount += 1;
    productTotal += toNumber_(row[map.itemsTotal - 1]);
    shippingTotal += toNumber_(row[map.shippingFee - 1]);
    lifetimeSpend += toNumber_(row[map.grandTotal - 1]);
  });

  return {
    customerId: lineUserId,
    orderCount,
    productTotal,
    shippingTotal,
    lifetimeSpend,
    rewards: getEligibleRewards_(lifetimeSpend)
  };
}

function getEligibleRewards_(lifetimeSpend) {
  const sheet = openDb_().getSheetByName('Rewards');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const map = getHeaderMap_(sheet);
  if (!map.RewardID || !map.RewardName) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const now = new Date();

  return rows
    .map(row => {
      const value = header => map[header] ? row[map[header] - 1] : '';
      const active = value('Active');
      const startDate = value('StartDate');
      const endDate = value('EndDate');
      const requiredSpend = toNumber_(value('RequiredSpend'));

      if (!(active === true || text_(active).toLowerCase() === 'true')) return null;
      if (startDate && new Date(startDate) > now) return null;
      if (endDate && new Date(endDate) < now) return null;
      if (requiredSpend > lifetimeSpend) return null;

      return {
        rewardId: text_(value('RewardID')),
        rewardName: text_(value('RewardName')),
        requiredSpend,
        rewardType: text_(value('RewardType')),
        rewardValue: value('RewardValue'),
        note: text_(value('Note'))
      };
    })
    .filter(Boolean);
}

/* =========================
 * FAVORITES
 * ========================= */

function handleFavoritesGet_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const { sheet, map } = assertHeaders_('Favorites', ['lineUserId', 'Brand', 'Name']);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, favorites: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const favorites = rows
    .filter(row => text_(row[map.lineUserId - 1]) === identity.lineUserId)
    .map(row => ({
      SKU: map.SKU ? text_(row[map.SKU - 1]) : '',
      Brand: text_(row[map.Brand - 1]),
      Size: map.Size ? text_(row[map.Size - 1]) : '',
      Name: text_(row[map.Name - 1])
    }));

  return { ok: true, favorites };
}

function handleFavoritesAdd_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const item = body.item || {};
  const brand = text_(item.Brand || item.brand);
  const size = text_(item.Size || item.size);
  const name = text_(item.Name || item.name);
  assert_(brand && name, 'Invalid favorite item');

  const { sheet, map } = assertHeaders_('Favorites', ['lineUserId', 'Brand', 'Name']);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      if (
        text_(row[map.lineUserId - 1]) === identity.lineUserId &&
        text_(row[map.Brand - 1]) === brand &&
        (!map.Size || text_(row[map.Size - 1]) === size) &&
        text_(row[map.Name - 1]) === name
      ) {
        if (map.updatedAt) sheet.getRange(index + 2, map.updatedAt).setValue(nowIso_());
        return { ok: true, duplicate: true };
      }
    }
  }

  const now = nowIso_();
  appendObjectRow_('Favorites', {
    lineUserId: identity.lineUserId,
    SKU: text_(item.SKU || item.id),
    Brand: brand,
    Size: size,
    Name: name,
    createdAt: now,
    updatedAt: now
  }, ZIPDAM_SCHEMA.Favorites);

  return { ok: true };
}

function handleFavoritesRemove_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const item = body.item || {};
  const brand = text_(item.Brand || item.brand);
  const size = text_(item.Size || item.size);
  const name = text_(item.Name || item.name);
  assert_(brand && name, 'Invalid favorite item');

  const { sheet, map } = assertHeaders_('Favorites', ['lineUserId', 'Brand', 'Name']);
  if (sheet.getLastRow() < 2) return { ok: true, removed: 0 };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let removed = 0;

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index];
    if (
      text_(row[map.lineUserId - 1]) === identity.lineUserId &&
      text_(row[map.Brand - 1]) === brand &&
      (!map.Size || text_(row[map.Size - 1]) === size) &&
      text_(row[map.Name - 1]) === name
    ) {
      sheet.deleteRow(index + 2);
      removed += 1;
    }
  }

  return { ok: true, removed };
}

/* =========================
 * TEMPLATES
 * ========================= */

function handleTemplatesGet_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const { sheet, map } = assertHeaders_('Templates', ['templateId', 'lineUserId', 'templateName', 'itemsJson']);
  if (sheet.getLastRow() < 2) return { ok: true, templates: [] };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const templates = rows
    .filter(row => text_(row[map.lineUserId - 1]) === identity.lineUserId)
    .map(row => {
      let items = [];
      try { items = JSON.parse(text_(row[map.itemsJson - 1]) || '[]'); } catch (_) {}

      return {
        templateId: text_(row[map.templateId - 1]),
        templateName: text_(row[map.templateName - 1]),
        items,
        lastUsedAt: map.lastUsedAt ? row[map.lastUsedAt - 1] : ''
      };
    });

  return { ok: true, templates };
}

function handleTemplatesAdd_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const templateName = text_(body.templateName);
  const items = Array.isArray(body.items) ? body.items : [];
  assert_(templateName, 'templateName required');
  assert_(items.length > 0, 'items required');

  const now = nowIso_();
  const templateId = generateSequentialId_('LAST_TEMPLATE_NO', 'TMP', 4);

  appendObjectRow_('Templates', {
    templateId,
    lineUserId: identity.lineUserId,
    templateName,
    itemsJson: JSON.stringify(items),
    createdAt: now,
    updatedAt: now,
    lastUsedAt: '',
    note: text_(body.note)
  }, ZIPDAM_SCHEMA.Templates);

  return { ok: true, templateId };
}

function handleTemplatesDelete_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const templateId = text_(body.templateId);
  assert_(templateId, 'templateId required');

  const { sheet, map } = assertHeaders_('Templates', ['templateId', 'lineUserId']);
  if (sheet.getLastRow() < 2) return { ok: true, removed: 0 };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let removed = 0;

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index];
    if (
      text_(row[map.lineUserId - 1]) === identity.lineUserId &&
      text_(row[map.templateId - 1]) === templateId
    ) {
      sheet.deleteRow(index + 2);
      removed += 1;
    }
  }

  return { ok: true, removed };
}

function generateSequentialId_(propertyKey, prefix, width) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const next = toNumber_(props.getProperty(propertyKey)) + 1;
    props.setProperty(propertyKey, String(next));
    return prefix + String(next).padStart(width, '0');
  } finally {
    lock.releaseLock();
  }
}

/* =========================
 * FREQUENT ITEMS
 * ========================= */

function handleFrequentGet_(body) {
  const identity = resolveIdentity_(body, { requireLine: true });
  const limit = Math.max(1, Math.min(20, Math.floor(toNumber_(body.limit) || 6)));

  const { sheet: ordersSheet, map: ordersMap } = assertHeaders_('Orders', ['OrderID', 'lineUserId']);
  if (ordersSheet.getLastRow() < 2) return { ok: true, frequent: [] };

  const orderRows = ordersSheet
    .getRange(2, 1, ordersSheet.getLastRow() - 1, ordersSheet.getLastColumn())
    .getValues();

  const orderIds = new Set(
    orderRows
      .filter(row => text_(row[ordersMap.lineUserId - 1]) === identity.lineUserId)
      .map(row => text_(row[ordersMap.OrderID - 1]))
      .filter(Boolean)
  );

  if (!orderIds.size) return { ok: true, frequent: [] };

  const { sheet: itemsSheet, map: itemsMap } = assertHeaders_(
    'OrderItems',
    ['OrderID', 'Brand', 'Name', 'qty']
  );

  if (itemsSheet.getLastRow() < 2) return { ok: true, frequent: [] };

  const rows = itemsSheet
    .getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn())
    .getValues();

  const counter = new Map();

  rows.forEach(row => {
    const orderId = text_(row[itemsMap.OrderID - 1]);
    if (!orderIds.has(orderId)) return;

    const brand = text_(row[itemsMap.Brand - 1]);
    const size = itemsMap.Size ? text_(row[itemsMap.Size - 1]) : '';
    const name = text_(row[itemsMap.Name - 1]);
    const quantity = toNumber_(row[itemsMap.qty - 1]);
    const key = productKey_(brand, size, name);

    const current = counter.get(key) || { Brand: brand, Size: size, Name: name, count: 0 };
    current.count += quantity;
    counter.set(key, current);
  });

  const frequent = Array.from(counter.values())
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);

  return { ok: true, frequent };
}

/* =========================
 * NOTIFICATIONS
 * ========================= */

function safePushOrderConfirmation_(
  token,
  lineUserId,
  orderId,
  items,
  itemsTotal,
  shippingFee,
  grandTotal,
  displayName,
  address,
  phone
) {
  if (!token) return { attempted: false, ok: false, error: 'LINE_MESSAGING_TOKEN is not configured' };
  if (!isRealLineUserId_(lineUserId)) return { attempted: false, ok: false, error: 'Invalid LINE user ID' };

  try {
    const itemText = items.map((item, index) => [
      `${index + 1}) ${item.Name}${item.Size ? ' ' + item.Size : ''}${item.pack ? ` (${item.pack}ชิ้น)` : ''}`,
      `   จำนวน ${item.qty} กล่อง`,
      `   ราคากล่องละ ${formatTHB_(item.unitPrice)}`,
      `   รวม ${formatTHB_(item.lineTotal)}`
    ].join('\n')).join('\n\n');

    const message = [
      '🧾 ยืนยันคำสั่งซื้อ ZIPDAM',
      `Order: ${orderId}`,
      `ลูกค้า: ${displayName || '-'}`,
      `ที่อยู่: ${address || '-'}`,
      `โทร: ${phone || '-'}`,
      '────────────',
      itemText,
      '────────────',
      `ค่าสินค้า: ${formatTHB_(itemsTotal)}`,
      `ค่าส่ง: ${formatTHB_(shippingFee)}`,
      `ยอดรวมสุทธิ: ${formatTHB_(grandTotal)}`,
      '🙏 ขอบคุณที่สั่งซื้อกับเรา'
    ].join('\n');

    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        to: lineUserId,
        messages: [{ type: 'text', text: message }]
      }),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    return {
      attempted: true,
      ok: status >= 200 && status < 300,
      status,
      body: response.getContentText()
    };
  } catch (err) {
    return { attempted: true, ok: false, error: errorMessage_(err) };
  }
}

function sendN8NWebhook_(url, payload) {
  if (!url) return { attempted: false, ok: false, error: 'N8N_WEBHOOK_URL is not configured' };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    return {
      attempted: true,
      ok: status >= 200 && status < 300,
      status,
      body: response.getContentText()
    };
  } catch (err) {
    return { attempted: true, ok: false, error: errorMessage_(err) };
  }
}
