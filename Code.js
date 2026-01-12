/**
 * ZIPDAM Backend - Google Apps Script Web App
 * Endpoints:
 *  - GET  ?action=catalog
 *  - POST {action:"order", idToken, cart:[{Brand,Size,Name,qty}]}
 *  - POST {action:"me", idToken}
 *  - POST {action:"favorites_get|favorites_add|favorites_remove", idToken, item?}
 *  - POST {action:"templates_get|templates_add|templates_delete", idToken, ...}
 *  - POST {action:"frequent_get", idToken, limit}
 */

/** ====== CONFIG ====== */
function getConfig_() {
  // Hard-coded config so Script Properties are no longer required.
  // TODO: update LINE_MESSAGING_TOKEN with your real channel access token.
  const SHEET_ID = '11sUClcToFNjQXafrZUgRaLGZW3NO2-cAfSIxe2IgsoE';
  const LINE_LOGIN_CHANNEL_ID = '2008727011';
  const LINE_MESSAGING_TOKEN = '6W4tvkhwl5GbmBKHLr4D3P5wLhjUaO1Ak5WJYArUZXTehnwHmodl+KJG3GWc6bMfLfzXUVkTWTdrE6IHeQ7Id10/z+/tpN4hLLjyPuh5e2efRC/ADXgjAljHuFwinVGbXNiBylywsemWI3Ikm/YXDQdB04t89/1O/w1cDnyilFU=';
  const N8N_WEBHOOK_URL = 'https://n8n.srv1112305.hstgr.cloud/webhook-test/zipdam';
  const FIXED_SHIPPING = 20;
  const ADMIN_LINE_USER_ID = 'U1d0318233f66c6ddc1fd998e49c5dcef'; // notify admin on new order

  return { SHEET_ID, LINE_LOGIN_CHANNEL_ID, LINE_MESSAGING_TOKEN, N8N_WEBHOOK_URL, FIXED_SHIPPING, ADMIN_LINE_USER_ID };
}

function json_(obj, code) {
  const out = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  // Apps Script Web App ไม่รองรับ setStatusCode โดยตรงใน ContentService
  // แต่เราจะใส่ ok/error ใน JSON แทน
  return out;
}

/** ====== ENTRYPOINTS ====== */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';
    if (action === 'catalog') {
      const catalog = getCatalog_();
      return json_({ ok: true, catalog });
    }
    return json_({ ok: false, error: 'Unknown action (GET). Use ?action=catalog' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    const body = parseJsonBody_(e);
    const action = String(body.action || '');

    switch (action) {
      case 'me': return json_(handleMe_(body));
      case 'order': return json_(handleOrder_(body));

      case 'favorites_get': return json_(handleFavoritesGet_(body));
      case 'favorites_add': return json_(handleFavoritesAdd_(body));
      case 'favorites_remove': return json_(handleFavoritesRemove_(body));

      case 'templates_get': return json_(handleTemplatesGet_(body));
      case 'templates_add': return json_(handleTemplatesAdd_(body));
      case 'templates_delete': return json_(handleTemplatesDelete_(body));

      case 'frequent_get': return json_(handleFrequentGet_(body));
      case 'customer_profile': return json_(handleCustomerProfileGet_(body));

      default:
        return json_({ ok: false, error: 'Unknown action (POST)' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/** ====== PARSING / VALIDATION ====== */
function parseJsonBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const txt = e.postData.contents;
  try { return JSON.parse(txt); } catch (_) { throw new Error('Invalid JSON body'); }
}

function assert_(cond, msg) {
  if (!cond) throw new Error(msg);
}

function toNumber_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function formatTHB_(n) {
  return '฿' + Number(n || 0).toLocaleString('th-TH');
}

function nowIso_() {
  return new Date().toISOString();
}

/** ====== LINE ID TOKEN VERIFY ======
 * Uses LINE verify endpoint:
 * POST https://api.line.me/oauth2/v2.1/verify (x-www-form-urlencoded)
 * params: id_token, client_id
 */
function verifyLineIdToken_(idToken) {
  const { LINE_LOGIN_CHANNEL_ID } = getConfig_();
  assert_(idToken && typeof idToken === 'string', 'Missing idToken');

  const url = 'https://api.line.me/oauth2/v2.1/verify';
  const payload = {
    id_token: idToken,
    client_id: LINE_LOGIN_CHANNEL_ID
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload,
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const text = res.getContentText();
  let data = {};
  try { data = JSON.parse(text); } catch (_) {}

  // สำเร็จจะได้ sub (user id), name, picture, email(ถ้ามี) ฯลฯ
  if (code !== 200) {
    throw new Error('LINE token verify failed: ' + (data.error_description || data.error || text));
  }
  assert_(data.sub, 'LINE verify ok but missing sub');
  return data; // {sub,name,picture, ...}
}

/** ====== SHEET HELPERS ====== */
function openDb_() {
  const { SHEET_ID } = getConfig_();
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet_(name) {
  const ss = openDb_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, idx) => {
    const key = String(h || '').trim();
    if (key) map[key] = idx + 1; // 1-based col index
  });
  return map;
}

function getLastRowValues_(sheet, n) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - n + 1);
  const num = last - start + 1;
  return sheet.getRange(start, 1, num, sheet.getLastColumn()).getValues();
}

/** ====== CATALOG ====== */
function getCatalog_() {
  const sh = getSheet_('Product');
  const map = getHeaderMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

  // รองรับทั้ง “ไฟล์ใหม่” และ “ไฟล์เดิม” โดยใช้ชื่อหัวคอลัมน์
  // ใหม่ (ที่แนะนำ): SKU, Brand, Type, Feature, Size, Name, mm, pack, price, promo_price, image_key, active
  // เดิม: Brand, Size, Name, ขนาด (มม.), บรรจุ (ชิ้น/กล่อง), ราคา/กล่อง (บาท)
  const col = (name, fallbackName) => map[name] || (fallbackName ? map[fallbackName] : null);

  const cSKU = col('SKU');
  const cBrand = col('Brand');
  const cType = col('Type');
  const cFeature = col('Feature');
  const cSize = col('Size');
  const cName = col('Name');
  const cMM = col('mm', 'ขนาด (มม.)');
  const cPack = col('pack', 'บรรจุ (ชิ้น/กล่อง)');
  const cPrice = col('price', 'ราคา/กล่อง (บาท)');
  const cPromo = col('promo_price', 'Promo Price');
  const cImgKey = col('image_key', 'Image Key');
  const cActive = col('active');

  const out = [];

  for (const r of values) {
    const brand = cBrand ? r[cBrand - 1] : null;
    const name = cName ? r[cName - 1] : null;
    if (!brand || !name) continue;

    const active = cActive ? r[cActive - 1] : true;
    if (active === false || String(active).toLowerCase() === 'false') continue;

    const price = toNumber_(cPrice ? r[cPrice - 1] : 0);
    const promo = cPromo ? toNumber_(r[cPromo - 1]) : 0;
    const finalPrice = (promo && promo > 0) ? promo : price;

    const featureRaw = cFeature ? String(r[cFeature - 1] || '').trim() : '';
    const featureTags = featureRaw ? featureRaw.split('|').map(s => s.trim()).filter(Boolean) : [];

    out.push({
      SKU: cSKU ? (r[cSKU - 1] || '') : '',
      Brand: brand,
      Type: cType ? (r[cType - 1] || '') : '',
      Feature: featureTags,
      Size: cSize ? (r[cSize - 1] || '') : '',
      Name: name,
      mm: cMM ? (r[cMM - 1] || '') : '',
      pack: cPack ? (r[cPack - 1] || '') : '',
      price: price,
      promo_price: (promo && promo > 0) ? promo : null,
      final_price: finalPrice,
      image_key: cImgKey ? (r[cImgKey - 1] || '') : ''
    });
  }
  return out;
}

function buildProductIndex_() {
  const catalog = getCatalog_();
  const idx = new Map();
  const skuIdx = new Map();
  for (const p of catalog) {
    const key = `${p.Brand}||${p.Size}||${p.Name}`;
    idx.set(key, p);
    if (p.SKU) skuIdx.set(String(p.SKU), p);
  }
  return { idx, skuIdx };
}

/** ====== ORDER HANDLER ====== */
function handleOrder_(body) {
  const { FIXED_SHIPPING, LINE_MESSAGING_TOKEN, N8N_WEBHOOK_URL, ADMIN_LINE_USER_ID } = getConfig_();
  const idToken = body.idToken || '';
  const cart = body.cart;

  assert_(Array.isArray(cart) && cart.length > 0, 'Cart is empty');
  let tokenVerified = false;
  let tokenError = '';
  let verified;

  if (idToken) {
    try {
      verified = verifyLineIdToken_(idToken);
      tokenVerified = true;
    } catch (err) {
      tokenError = String(err && err.message ? err.message : err);
      // Allow order to proceed even if token expired/invalid (no auth-based features).
      verified = {
        sub: body.lineUserId || `GUEST-${Date.now()}`,
        name: body.displayName || 'Guest',
      };
    }
  } else {
    verified = {
      sub: body.lineUserId || `GUEST-${Date.now()}`,
      name: body.displayName || 'Guest',
    };
  }

  const loginSub = verified.sub;
  const displayName = verified.name || '';

  const bodyLineUserId = String(body.lineUserId || '').trim();
  // Prefer the Messaging API userId if provided; fallback to loginSub/guest id.
  const lineUserId = (bodyLineUserId && bodyLineUserId.startsWith('U')) ? bodyLineUserId : loginSub;

  // ensure customer row exists/update lastSeen (keyed by login sub/guest)
  const existingProfile = getCustomerProfile_(loginSub);
  const providedStore = String(body.store || body.storeName || '').trim();
  const providedArea = String(body.soi || body.area || '').trim();
  const providedAddress = String(body.address || body.defaultAddress || '').trim();
  const providedPhone = String(body.phone || '').trim();

  const store = providedStore || existingProfile.store || '';
  const area = providedArea || existingProfile.area || '';
  const address = providedAddress || existingProfile.address || existingProfile.defaultAddress || '';
  const phone = providedPhone || existingProfile.phone || '';

  upsertCustomer_(loginSub, displayName, address, phone, store, area);

  const { idx: productIndex, skuIdx } = buildProductIndex_();

  let itemsTotal = 0;
  const itemsToWrite = [];

  for (const item of cart) {
    const rawSku = String(item.SKU || item.id || '').trim();
    const Brand = String(item.Brand || item.brand || '').trim();
    const Size = String(item.Size || item.size || '').trim();
    const Name = String(item.Name || item.name || '').trim();
    const qty = Math.floor(Number(item.qty || item.quantity || 0));

    assert_(qty > 0, 'qty must be > 0');

    let p = rawSku ? skuIdx.get(rawSku) : null;
    if (!p && (Brand || Name)) {
      const key = `${Brand}||${Size}||${Name}`;
      p = productIndex.get(key);
      if (!p && Brand && Name) {
        // attempt match ignoring Size mismatch
        for (const candidate of productIndex.values()) {
          if (candidate.Brand === Brand && candidate.Name === Name) { p = candidate; break; }
        }
      }
    }

    assert_(p, `Product not found: ${Brand || rawSku} / ${Size} / ${Name}`);

    const unitPrice = toNumber_(p.final_price);
    const lineTotal = unitPrice * qty;

    itemsTotal += lineTotal;

    itemsToWrite.push({
      SKU: p.SKU || rawSku || '',
      Brand: Brand || p.Brand || '',
      Size: Size || p.Size || '',
      Name: Name || p.Name || '',
      pack: p.pack || '',
      qty,
      unitPrice,
      lineTotal
    });
  }

  itemsTotal = Math.round(itemsTotal);
  const shippingFee = FIXED_SHIPPING;
  const grandTotal = itemsTotal + shippingFee;

  const orderId = getNextOrderId_();

  // write to sheets
  appendOrder_(orderId, lineUserId, displayName, itemsTotal, shippingFee, grandTotal, loginSub, address, phone, store, area);
  appendOrderItems_(orderId, itemsToWrite);
  let linePush = { attempted: false, ok: false };
  let adminPush = { attempted: false, ok: false };
  try {
    linePush = pushLineOrderConfirm_(LINE_MESSAGING_TOKEN, lineUserId, orderId, itemsToWrite, itemsTotal, shippingFee, grandTotal, displayName, address) || linePush;
  } catch (err) {
    console.error('pushLineOrderConfirm_ call failed', err);
    linePush = { attempted: true, ok: false, error: String(err && err.message ? err.message : err) };
  }
  try {
    if (ADMIN_LINE_USER_ID) {
      // Reuse the same messaging token/settings as customer confirmation, but deliver to admin userId.
      adminPush = pushLineOrderConfirm_(LINE_MESSAGING_TOKEN, ADMIN_LINE_USER_ID, orderId, itemsToWrite, itemsTotal, shippingFee, grandTotal, displayName, address) || adminPush;
    }
  } catch (err) {
    console.error('admin push failed', err);
    adminPush = { attempted: true, ok: false, error: String(err && err.message ? err.message : err) };
  }
  console.log('linePush', JSON.stringify(linePush));
  console.log('adminPush', JSON.stringify(adminPush));

  const n8nPush = sendN8NWebhook_(N8N_WEBHOOK_URL, {
    orderId,
    lineUserId,
    loginSub,
    displayName,
    store,
    area,
    address,
    phone,
    itemsTotal,
    shippingFee,
    grandTotal,
    cart,
    linePush,
    tokenVerified,
    tokenError,
    createdAt: nowIso_(),
  });
  console.log('n8nPush', JSON.stringify(n8nPush));

  return {
    ok: true,
    orderId,
    lineUserId,
    loginSub,
    displayName,
    store,
    area,
    address,
    phone,
    tokenVerified,
    tokenError,
    itemsTotal,
    shippingFee,
    grandTotal,
    linePush,
    adminPush,
    n8nPush
  };
}

function appendOrder_(orderId, lineUserId, displayName, itemsTotal, shippingFee, grandTotal, customerId, address, phone, store, area) {
  const sh = getSheet_('Orders');
  const map = getHeaderMap_(sh);

  // รองรับหัวคอลัมน์ตามไฟล์ที่ผมทำให้:
  // OrderID, CreatedAt, lineUserId, displayName, customerId, store, itemsTotal, shippingFee, grandTotal, status, paymentStatus, paidAt, address, phone, note
  const row = new Array(sh.getLastColumn()).fill('');

  function set(colName, val) {
    const c = map[colName];
    if (c) row[c - 1] = val;
  }

  set('OrderID', orderId);
  set('CreatedAt', nowIso_());
  set('lineUserId', lineUserId);
  set('displayName', displayName);
  set('customerId', customerId || '');
  set('store', store || '');
  set('area', area || '');
  set('itemsTotal', itemsTotal);
  set('shippingFee', shippingFee);
  set('grandTotal', grandTotal);
  set('address', address || '');
  set('phone', phone || '');
  set('status', 'CONFIRMED');        // หรือ NEW ก็ได้
  set('paymentStatus', 'UNPAID');    // เริ่มต้น

  sh.appendRow(row);
}

function appendOrderItems_(orderId, items) {
  const sh = getSheet_('OrderItems');
  const map = getHeaderMap_(sh);

  // รองรับหัวคอลัมน์ตามไฟล์ที่ผมทำให้:
  // OrderID, SKU, Brand, Size, Name, qty, unitPrice, lineTotal
  const rows = items.map(it => {
    const row = new Array(sh.getLastColumn()).fill('');
    function set(colName, val) {
      const c = map[colName];
      if (c) row[c - 1] = val;
    }
    set('OrderID', orderId);
    set('SKU', it.SKU || '');
    set('Brand', it.Brand);
    set('Size', it.Size);
    set('Name', it.Name);
    set('qty', it.qty);
    set('unitPrice', it.unitPrice);
    set('lineTotal', it.lineTotal);
    return row;
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, sh.getLastColumn()).setValues(rows);
  }
}

function pushLineOrderConfirm_(token, lineUserId, orderId, items, itemsTotal, shippingFee, grandTotal, displayName, address) {
  if (!token) return { attempted: false, ok: false, error: 'Missing LINE_MESSAGING_TOKEN' };
  if (!lineUserId || !String(lineUserId).startsWith('U')) return { attempted: false, ok: false, error: 'Missing/invalid lineUserId' };

  try {
    const lines = Array.isArray(items) ? items : [];
    const itemLines = lines.map((it, idx) => {
      const namePart = `${it.Brand ? it.Brand + ' ' : ''}${it.Name || ''}`.trim();
      const packLabel = it.pack ? ` (${it.pack} ชิ้น)` : '';
      const unitLabel = `ราคากล่องละ ${formatTHB_(it.unitPrice)}`;
      const qtyLabel = `จำนวน ${it.qty} กล่อง`;
      const lineTotalLabel = `รวม ${formatTHB_(it.lineTotal)}`;
      return [
        `${idx + 1}) ${namePart}${packLabel}`,
        `   📦 ${qtyLabel}`,
        `   💰 ${unitLabel}`,
        `   🔸 ${lineTotalLabel}`
      ].join('\n');
    });

    const text = [
      '🧾 ยืนยันคำสั่งซื้อ ZIPDAM',
      `Order: ${orderId}`,
      `ลูกค้า: ${displayName || '-'}`,
      `ที่อยู่: ${address || '-'}`,
      '────────────',
      'รายการ:',
      itemLines.join('\n\n'),
      '────────────',
      `ค่าสินค้า: ${formatTHB_(itemsTotal)}`,
      `ค่าส่ง: ${formatTHB_(shippingFee)}`,
      `ยอดรวมสุทธิ: ${formatTHB_(grandTotal)}`,
      '🙏 ขอบคุณที่สั่งซื้อกับเรา'
    ].filter(Boolean).join('\n');

    const payload = {
      to: lineUserId,
      messages: [{
        type: 'text',
        text
      }]
    };

    const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = resp.getResponseCode();
    const body = resp.getContentText();

    return {
      attempted: true,
      ok: status >= 200 && status < 300,
      status,
      body
    };
  } catch (err) {
    console.error('pushLineOrderConfirm_ failed', err);
    return { attempted: true, ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function sendN8NWebhook_(url, payload) {
  if (!url) return { attempted: false, ok: false, error: 'Missing N8N_WEBHOOK_URL' };
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const status = resp.getResponseCode();
    const body = resp.getContentText();
    return { attempted: true, ok: status >= 200 && status < 300, status, body };
  } catch (err) {
    return { attempted: true, ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function pushAdminOrderAlert_(token, adminUserId, details) {
  if (!token || !adminUserId) return { attempted: false, ok: false, error: 'Missing admin notify config' };
  try {
    const lines = Array.isArray(details.items) ? details.items : [];
    const itemLines = lines.slice(0, 3).map((it, idx) => {
      const namePart = `${it.Brand ? it.Brand + ' ' : ''}${it.Name || ''}`.trim();
      return `${idx + 1}) ${namePart} x${it.qty} = ${formatTHB_(it.lineTotal)}`;
    });
    if (lines.length > 3) {
      itemLines.push(`...อีก ${lines.length - 3} รายการ`);
    }
    const text = [
      '[ADMIN] คำสั่งซื้อใหม่',
      `Order: ${details.orderId}`,
      `ลูกค้า: ${details.displayName || '-'} (${details.lineUserId || '-'})`,
      ...itemLines,
      `ค่าสินค้า: ${formatTHB_(details.itemsTotal)}`,
      `ค่าส่ง: ${formatTHB_(details.shippingFee)}`,
      `รวมสุทธิ: ${formatTHB_(details.grandTotal)}`
    ].join('\n');

    const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        to: adminUserId,
        messages: [{ type: 'text', text }]
      }),
      muteHttpExceptions: true
    });

    const status = resp.getResponseCode();
    const body = resp.getContentText();
    return {
      attempted: true,
      ok: status >= 200 && status < 300,
      status,
      body
    };
  } catch (err) {
    return { attempted: true, ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/** ====== ORDER ID GENERATION ====== */
function getNextOrderId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    let last = Number(props.getProperty('LAST_ORDER_NO') || 0);

    // ถ้ายังไม่เคยตั้ง ให้สแกนใน Orders หาเลขมากสุด
    if (!last || last <= 0) {
      last = scanMaxOrderNo_();
    }

    const next = last + 1;
    props.setProperty('LAST_ORDER_NO', String(next));
    return 'OD' + String(next).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

function scanMaxOrderNo_() {
  const sh = getSheet_('Orders');
  const map = getHeaderMap_(sh);
  const cOrderId = map['OrderID'];
  if (!cOrderId) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const vals = sh.getRange(2, cOrderId, lastRow - 1, 1).getValues().flat();
  let maxNo = 0;
  for (const v of vals) {
    const s = String(v || '');
    const m = s.match(/^OD(\d+)$/);
    if (m) {
      const n = Number(m[1]);
      if (n > maxNo) maxNo = n;
    }
  }
  return maxNo;
}

/** ====== CUSTOMER (optional but recommended) ====== */
function handleMe_(body) {
  const verified = verifyLineIdToken_(body.idToken);
  const lineUserId = verified.sub;
  const displayName = verified.name || '';
  upsertCustomer_(lineUserId, displayName);
  return { ok: true, lineUserId, displayName };
}

function upsertCustomer_(lineUserId, displayName, address, phone, store, area) {
  const sh = getSheet_('Customer');
  const map = getHeaderMap_(sh);
  const cUser = map['lineUserId'];
  if (!cUser) return;

  const last = sh.getLastRow();
  const now = nowIso_();

  // scan existing
  if (last >= 2) {
    const vals = sh.getRange(2, cUser, last - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '') === String(lineUserId)) {
        // update displayName + lastSeenAt
        if (map['displayName']) sh.getRange(i + 2, map['displayName']).setValue(displayName);
        if (address && map['address']) sh.getRange(i + 2, map['address']).setValue(address);
        if (address && map['defaultAddress']) sh.getRange(i + 2, map['defaultAddress']).setValue(address);
        if (phone && map['phone']) sh.getRange(i + 2, map['phone']).setValue(phone);
        if (store && map['store']) sh.getRange(i + 2, map['store']).setValue(store);
        if (area && map['area']) sh.getRange(i + 2, map['area']).setValue(area);
        if (map['lastSeenAt']) sh.getRange(i + 2, map['lastSeenAt']).setValue(now);
        return;
      }
    }
  }

  // insert new
  const row = new Array(sh.getLastColumn()).fill('');
  if (map['lineUserId']) row[map['lineUserId'] - 1] = lineUserId;
  if (map['displayName']) row[map['displayName'] - 1] = displayName;
  if (map['address']) row[map['address'] - 1] = address || '';
  if (map['defaultAddress']) row[map['defaultAddress'] - 1] = address || '';
  if (map['phone']) row[map['phone'] - 1] = phone || '';
  if (map['store']) row[map['store'] - 1] = store || '';
  if (map['area']) row[map['area'] - 1] = area || '';
  if (map['createdAt']) row[map['createdAt'] - 1] = now;
  if (map['lastSeenAt']) row[map['lastSeenAt'] - 1] = now;
  sh.appendRow(row);
}

/** ====== FAVORITES ====== */
function resolveFavoritesUser_(body) {
  const providedId = String(body.lineUserId || '').trim();
  if (body.idToken) {
    try {
      const verified = verifyLineIdToken_(body.idToken);
      // Prefer explicit provided lineUserId (from LIFF profile) if present; otherwise use verified sub.
      if (providedId) return providedId;
      return verified.sub;
    } catch (err) {
      if (providedId) return providedId; // fallback if verify fails
      throw err;
    }
  }
  assert_(providedId, 'Missing user identity for favorites');
  return providedId;
}

function handleFavoritesGet_(body) {
  const lineUserId = resolveFavoritesUser_(body);

  const sh = getSheet_('Favorites');
  const map = getHeaderMap_(sh);

  const cUser = map['lineUserId'];
  assert_(cUser, 'Favorites sheet missing lineUserId');

  const last = sh.getLastRow();
  if (last < 2) return { ok: true, favorites: [] };

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (const r of values) {
    if (String(r[cUser - 1] || '') !== String(lineUserId)) continue;
    out.push({
      SKU: map['SKU'] ? (r[map['SKU'] - 1] || '') : '',
      Brand: map['Brand'] ? (r[map['Brand'] - 1] || '') : '',
      Size: map['Size'] ? (r[map['Size'] - 1] || '') : '',
      Name: map['Name'] ? (r[map['Name'] - 1] || '') : ''
    });
  }
  return { ok: true, favorites: out };
}

function handleFavoritesAdd_(body) {
  const lineUserId = resolveFavoritesUser_(body);

  const item = body.item || {};
  const Brand = String(item.Brand || '').trim();
  const Size = String(item.Size || '').trim();
  const Name = String(item.Name || '').trim();
  assert_(Brand && Name, 'Invalid favorite item');

  const sh = getSheet_('Favorites');
  const map = getHeaderMap_(sh);
  const now = nowIso_();

  // prevent duplicates
  const last = sh.getLastRow();
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      if (String(r[map['lineUserId'] - 1] || '') === String(lineUserId) &&
          String(r[map['Brand'] - 1] || '') === Brand &&
          String(r[map['Size'] - 1] || '') === Size &&
          String(r[map['Name'] - 1] || '') === Name) {
        // update updatedAt
        if (map['updatedAt']) sh.getRange(i + 2, map['updatedAt']).setValue(now);
        return { ok: true };
      }
    }
  }

  const row = new Array(sh.getLastColumn()).fill('');
  if (map['lineUserId']) row[map['lineUserId'] - 1] = lineUserId;
  if (map['SKU']) row[map['SKU'] - 1] = String(item.SKU || '');
  if (map['Brand']) row[map['Brand'] - 1] = Brand;
  if (map['Size']) row[map['Size'] - 1] = Size;
  if (map['Name']) row[map['Name'] - 1] = Name;
  if (map['createdAt']) row[map['createdAt'] - 1] = now;
  if (map['updatedAt']) row[map['updatedAt'] - 1] = now;
  sh.appendRow(row);

  return { ok: true };
}

function handleFavoritesRemove_(body) {
  const lineUserId = resolveFavoritesUser_(body);

  const item = body.item || {};
  const Brand = String(item.Brand || '').trim();
  const Size = String(item.Size || '').trim();
  const Name = String(item.Name || '').trim();
  assert_(Brand && Name, 'Invalid favorite item');

  const sh = getSheet_('Favorites');
  const map = getHeaderMap_(sh);

  const last = sh.getLastRow();
  if (last < 2) return { ok: true };

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (String(r[map['lineUserId'] - 1] || '') === String(lineUserId) &&
        String(r[map['Brand'] - 1] || '') === Brand &&
        String(r[map['Size'] - 1] || '') === Size &&
        String(r[map['Name'] - 1] || '') === Name) {
      sh.deleteRow(i + 2);
    }
  }
  return { ok: true };
}

/** ====== TEMPLATES ====== */
function handleTemplatesGet_(body) {
  const verified = verifyLineIdToken_(body.idToken);
  const lineUserId = verified.sub;

  const sh = getSheet_('Templates');
  const map = getHeaderMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, templates: [] };

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (const r of values) {
    if (String(r[map['lineUserId'] - 1] || '') !== String(lineUserId)) continue;
    const itemsJson = map['itemsJson'] ? String(r[map['itemsJson'] - 1] || '[]') : '[]';
    let items = [];
    try { items = JSON.parse(itemsJson); } catch (_) {}
    out.push({
      templateId: map['templateId'] ? (r[map['templateId'] - 1] || '') : '',
      templateName: map['templateName'] ? (r[map['templateName'] - 1] || '') : '',
      items,
      lastUsedAt: map['lastUsedAt'] ? (r[map['lastUsedAt'] - 1] || '') : ''
    });
  }
  return { ok: true, templates: out };
}

function handleTemplatesAdd_(body) {
  const verified = verifyLineIdToken_(body.idToken);
  const lineUserId = verified.sub;

  const templateName = String(body.templateName || '').trim();
  const items = body.items;
  assert_(templateName, 'templateName required');
  assert_(Array.isArray(items) && items.length > 0, 'items required');

  const sh = getSheet_('Templates');
  const map = getHeaderMap_(sh);
  const now = nowIso_();

  const templateId = generateTemplateId_();

  const row = new Array(sh.getLastColumn()).fill('');
  if (map['templateId']) row[map['templateId'] - 1] = templateId;
  if (map['lineUserId']) row[map['lineUserId'] - 1] = lineUserId;
  if (map['templateName']) row[map['templateName'] - 1] = templateName;
  if (map['itemsJson']) row[map['itemsJson'] - 1] = JSON.stringify(items);
  if (map['createdAt']) row[map['createdAt'] - 1] = now;
  if (map['updatedAt']) row[map['updatedAt'] - 1] = now;
  if (map['lastUsedAt']) row[map['lastUsedAt'] - 1] = '';
  sh.appendRow(row);

  return { ok: true, templateId };
}

function handleTemplatesDelete_(body) {
  const verified = verifyLineIdToken_(body.idToken);
  const lineUserId = verified.sub;

  const templateId = String(body.templateId || '').trim();
  assert_(templateId, 'templateId required');

  const sh = getSheet_('Templates');
  const map = getHeaderMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true };

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (String(r[map['lineUserId'] - 1] || '') === String(lineUserId) &&
        String(r[map['templateId'] - 1] || '') === templateId) {
      sh.deleteRow(i + 2);
    }
  }
  return { ok: true };
}

function generateTemplateId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    let last = Number(props.getProperty('LAST_TEMPLATE_NO') || 0);
    const next = last + 1;
    props.setProperty('LAST_TEMPLATE_NO', String(next));
    return 'TMP' + String(next).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

/** ====== FREQUENT (compute from OrderItems) ====== */
function handleFrequentGet_(body) {
  const verified = verifyLineIdToken_(body.idToken);
  const lineUserId = verified.sub;
  const limit = Math.max(1, Math.min(20, Number(body.limit || 6)));

  // ต้องรู้ว่าออเดอร์ไหนเป็นของ user นี้: จาก Orders (lineUserId) -> list OrderID
  const ordersSh = getSheet_('Orders');
  const oMap = getHeaderMap_(ordersSh);
  const cUser = oMap['lineUserId'];
  const cOrder = oMap['OrderID'];
  assert_(cUser && cOrder, 'Orders sheet missing lineUserId/OrderID');

  const lastO = ordersSh.getLastRow();
  if (lastO < 2) return { ok: true, frequent: [] };

  const oValues = ordersSh.getRange(2, 1, lastO - 1, ordersSh.getLastColumn()).getValues();
  const orderIds = new Set();
  for (const r of oValues) {
    if (String(r[cUser - 1] || '') === String(lineUserId)) {
      orderIds.add(String(r[cOrder - 1] || ''));
    }
  }
  if (orderIds.size === 0) return { ok: true, frequent: [] };

  const itemSh = getSheet_('OrderItems');
  const iMap = getHeaderMap_(itemSh);
  const iOrder = iMap['OrderID'];
  const iBrand = iMap['Brand'];
  const iSize = iMap['Size'];
  const iName = iMap['Name'];
  const iQty = iMap['qty'];
  assert_(iOrder && iBrand && iName && iQty, 'OrderItems missing columns');

  const lastI = itemSh.getLastRow();
  if (lastI < 2) return { ok: true, frequent: [] };

  const iValues = itemSh.getRange(2, 1, lastI - 1, itemSh.getLastColumn()).getValues();
  const counter = new Map(); // key -> {Brand,Size,Name,count}
  for (const r of iValues) {
    const oid = String(r[iOrder - 1] || '');
    if (!orderIds.has(oid)) continue;

    const Brand = String(r[iBrand - 1] || '');
    const Size = iSize ? String(r[iSize - 1] || '') : '';
    const Name = String(r[iName - 1] || '');
    const qty = Number(r[iQty - 1] || 0);

    const key = `${Brand}||${Size}||${Name}`;
    const cur = counter.get(key) || { Brand, Size, Name, count: 0 };
    cur.count += qty;
    counter.set(key, cur);
  }

  const sorted = Array.from(counter.values()).sort((a, b) => b.count - a.count).slice(0, limit);
  return { ok: true, frequent: sorted };
}

function handleCustomerProfileGet_(body) {
  // similar to favorites: allow provided lineUserId fallback
  const lineUserId = resolveFavoritesUser_(body);
  const profile = getCustomerProfile_(lineUserId);
  return { ok: true, profile };
}

function getCustomerProfile_(lineUserId) {
  const sh = getSheet_('Customer');
  const map = getHeaderMap_(sh);
  const cUser = map['lineUserId'];
  if (!cUser) return { address: '', defaultAddress: '', phone: '', store: '', area: '' };

  const last = sh.getLastRow();
  if (last < 2) return { address: '', defaultAddress: '', phone: '', store: '', area: '' };

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    if (String(r[cUser - 1] || '') === String(lineUserId)) {
      return {
        address: map['address'] ? String(r[map['address'] - 1] || '') : '',
        defaultAddress: map['defaultAddress'] ? String(r[map['defaultAddress'] - 1] || '') : '',
        phone: map['phone'] ? String(r[map['phone'] - 1] || '') : '',
        store: map['store'] ? String(r[map['store'] - 1] || '') : '',
        area: map['area'] ? String(r[map['area'] - 1] || '') : ''
      };
    }
  }
  return { address: '', defaultAddress: '', phone: '', store: '', area: '' };
}
