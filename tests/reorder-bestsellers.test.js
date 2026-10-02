const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "Code.js"), "utf8");
const context = vm.createContext({ console });
vm.runInContext(source, context, { filename: "Code.js" });

const CUSTOMER = `U${"a".repeat(32)}`;
const OTHER = `U${"b".repeat(32)}`;

function fakeSheet(headers, rows) {
  const map = {};
  headers.forEach((header, i) => { map[header] = i + 1; });
  return {
    map,
    sheet: {
      getLastRow: () => rows.length + 1,
      getLastColumn: () => headers.length,
      getRange: () => ({ getValues: () => rows.map(row => headers.map(h => row[h] ?? "")) }),
    },
  };
}

function useSheets(sheets) {
  context.assertHeaders_ = name => {
    assert.ok(sheets[name], `unexpected sheet ${name}`);
    return sheets[name];
  };
}

const orders = fakeSheet(
  ["OrderID", "CreatedAt", "lineUserId", "status"],
  [
    { OrderID: "OD001", CreatedAt: "2025-06-14 0:00:00", lineUserId: "", status: "LEGACY" },
    { OrderID: "OD100", CreatedAt: "2026-09-10T03:00:00.000Z", lineUserId: CUSTOMER, status: "CONFIRMED" },
    { OrderID: "OD101", CreatedAt: "2026-09-28T03:00:00.000Z", lineUserId: CUSTOMER, status: "CONFIRMED" },
    { OrderID: "OD102", CreatedAt: "2026-09-30T03:00:00.000Z", lineUserId: CUSTOMER, status: "CANCELLED" },
    { OrderID: "OD103", CreatedAt: "2026-10-01T03:00:00.000Z", lineUserId: OTHER, status: "CONFIRMED" },
  ]
);

const items = fakeSheet(
  ["OrderID", "SKU", "Brand", "Size", "Name", "qty"],
  [
    { OrderID: "OD001", SKU: "", Brand: "Durex", Size: "Big", Name: "ดูเร็กซ์ แอรี่", qty: 50 },
    { OrderID: "OD100", SKU: "ONE-0029", Brand: "Onetouch", Size: "Big", Name: "Onetouch แฮปปี้", qty: 4 },
    { OrderID: "OD101", SKU: "ONE-0029", Brand: "Onetouch", Size: "Big", Name: "Onetouch แฮปปี้", qty: 5 },
    { OrderID: "OD101", SKU: "ONE-0037", Brand: "Onetouch", Size: "Gel", Name: "Onetouch เจลหล่อลื่น กลิ่นธรรมชาติ", qty: 2 },
    { OrderID: "OD102", SKU: "DUR-0058", Brand: "Durex", Size: "Small", Name: "ดูเร็กซ์ แอรี่ (กล่องเล็ก)", qty: 99 },
    { OrderID: "OD103", SKU: "", Brand: "Durex", Size: "Small", Name: "ดูเร็กซ์ แอรี่ (กล่องเล็ก)", qty: 3 },
  ]
);

const products = fakeSheet(
  ["SKU", "Brand", "Size", "Name", "price", "active"],
  [
    { SKU: "ONE-0029", Brand: "Onetouch", Size: "Big", Name: "Onetouch แฮปปี้", price: 70, active: true },
    { SKU: "ONE-0037", Brand: "Onetouch", Size: "Gel", Name: "Onetouch เจลหล่อลื่น กลิ่นธรรมชาติ", price: 65, active: true },
    { SKU: "DUR-0058", Brand: "Durex", Size: "Small", Name: "ดูเร็กซ์ แอรี่ (กล่องเล็ก)", price: 58, active: true },
    { SKU: "DUR-0001", Brand: "Durex", Size: "Big", Name: "ดูเร็กซ์ แอรี่", price: 200, active: false },
  ]
);

useSheets({ Orders: orders, OrderItems: items, Product: products });
context.resolveIdentity_ = () => ({ lineUserId: CUSTOMER });

// frequent_get: SKU travels with each frequent row, and the newest
// non-cancelled order comes back whole with its quantities.
const frequent = context.handleFrequentGet_({ limit: 6 });
assert.equal(frequent.ok, true);
assert.equal(frequent.frequent[0].SKU, "ONE-0029");
assert.equal(frequent.frequent[0].count, 9);
assert.equal(frequent.lastOrder.orderId, "OD101");
assert.equal(frequent.lastOrder.createdAt, "2026-09-28T03:00:00.000Z");
assert.deepEqual(
  JSON.parse(JSON.stringify(frequent.lastOrder.items)).map(i => [i.SKU, i.qty]),
  [["ONE-0029", 5], ["ONE-0037", 2]]
);

// A customer with no orders gets an explicit null rather than a missing key.
context.resolveIdentity_ = () => ({ lineUserId: `U${"c".repeat(32)}` });
const none = context.handleFrequentGet_({});
assert.equal(none.lastOrder, null);
assert.equal(none.frequent.length, 0);

// Unreadable dates fall back to sheet order instead of dropping the order.
const undated = fakeSheet(["OrderID", "lineUserId"], [
  { OrderID: "OD1", lineUserId: CUSTOMER },
  { OrderID: "OD2", lineUserId: CUSTOMER },
]);
assert.equal(context.findLastOrderRow_(
  undated.sheet.getRange().getValues(), undated.map
)[0], "OD2");

// Best sellers: last 30 days only, cancelled orders ignored, legacy rows
// without a SKU resolved by brand + name, inactive products dropped.
const now = new Date("2026-10-02T00:00:00.000Z");
assert.deepEqual(
  Array.from(context.calculateBestSellers_(now, 30, 5)),
  ["ONE-0029", "DUR-0058", "ONE-0037"]
);
assert.deepEqual(Array.from(context.calculateBestSellers_(now, 30, 1)), ["ONE-0029"]);

// The cache answers repeat calls, and a broken sheet never breaks the catalog.
const store = {};
context.CacheService = {
  getScriptCache: () => ({
    get: key => store[key] ?? null,
    put: (key, value) => { store[key] = value; },
  }),
};
context.Date = class extends Date {
  constructor(...args) { super(...(args.length ? args : ["2026-10-02T00:00:00.000Z"])); }
};
assert.deepEqual(Array.from(context.safeBestSellers_()), ["ONE-0029", "DUR-0058", "ONE-0037"]);
context.assertHeaders_ = () => { throw new Error("sheet missing"); };
assert.deepEqual(Array.from(context.safeBestSellers_()), ["ONE-0029", "DUR-0058", "ONE-0037"]);
delete store["bestSellers:v1"];
assert.deepEqual(Array.from(context.safeBestSellers_()), []);

console.log("reorder and best seller tests passed");
