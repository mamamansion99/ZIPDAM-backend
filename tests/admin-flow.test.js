const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "Code.js"), "utf8");
const context = vm.createContext({ console });
vm.runInContext(source, context, { filename: "Code.js" });

const ADMIN_ID = `U${"a".repeat(32)}`;
const CUSTOMER_ID = `U${"b".repeat(32)}`;
const OTHER_ID = `U${"c".repeat(32)}`;

context.getConfig_ = () => ({
  ADMIN_LINE_USER_ID: ADMIN_ID,
  ADMIN_LINE_USER_IDS: [ADMIN_ID],
  ALLOW_LINE_ID_MISMATCH: false,
  ALLOW_GUEST_ORDERS: false,
  FIXED_SHIPPING: 20,
  LOW_ORDER_SHIPPING: 30,
  SHIPPING_THRESHOLD: 200,
  LINE_MESSAGING_TOKEN: "",
  N8N_WEBHOOK_URL: "",
});
context.verifyLineIdToken_ = (token) => {
  if (token === "admin-token") return { sub: ADMIN_ID, name: "Admin" };
  if (token === "other-token") return { sub: OTHER_ID, name: "Other" };
  throw new Error("bad token");
};

assert.equal(
  context.handleAdminStatus_({
    idToken: "admin-token",
    lineUserId: ADMIN_ID,
  }).isAdmin,
  true,
);
assert.equal(
  context.handleAdminStatus_({
    idToken: "other-token",
    lineUserId: OTHER_ID,
  }).isAdmin,
  false,
);
assert.throws(
  () =>
    context.requireAdmin_({
      idToken: "admin-token",
      lineUserId: OTHER_ID,
    }),
  /LINE identity mismatch/,
);
assert.throws(
  () =>
    context.requireAdmin_({
      idToken: "other-token",
      lineUserId: OTHER_ID,
    }),
  /Admin access required/,
);

const originalCreateOrder = context.createOrder_;
let capturedCreate = null;
context.createOrder_ = (body, identity, options) => {
  capturedCreate = { body, identity, options };
  return { ok: true, orderId: "ODTEST" };
};
context.getCustomerProfile_ = (lineUserId) =>
  lineUserId === CUSTOMER_ID
    ? {
        lineUserId: CUSTOMER_ID,
        customerId: CUSTOMER_ID,
        displayName: "Customer",
        store: "Store",
        area: "Area",
        phone: "0812345678",
        defaultAddress: "Address",
      }
    : {
        lineUserId: "",
      };

context.handleAdminOrder_({
  idToken: "admin-token",
  lineUserId: ADMIN_ID,
  selectedCustomerId: CUSTOMER_ID,
  cart: [{ SKU: "SKU-1", qty: 1 }],
});
assert.equal(capturedCreate.identity.lineUserId, CUSTOMER_ID);
assert.equal(capturedCreate.identity.customerId, CUSTOMER_ID);
assert.equal(capturedCreate.options.createdByLineUserId, ADMIN_ID);
assert.equal(capturedCreate.options.orderMode, "ADMIN");

context.resolveIdentity_ = () => ({
  lineUserId: OTHER_ID,
  customerId: OTHER_ID,
  displayName: "Self customer",
  tokenVerified: true,
  tokenError: "",
  isGuest: false,
});
context.handleOrder_({ orderMode: "ADMIN", cart: [{ SKU: "SKU-1", qty: 1 }] });
assert.equal(capturedCreate.options.orderMode, "SELF");
assert.equal(capturedCreate.options.createdByLineUserId, OTHER_ID);

context.createOrder_ = originalCreateOrder;
let writtenOrder = null;
let writtenItems = null;
context.getCustomerProfile_ = () => ({
  lineUserId: CUSTOMER_ID,
  customerId: CUSTOMER_ID,
  displayName: "Customer",
  store: "Store",
  area: "Area",
  phone: "0812345678",
  defaultAddress: "Address",
});
context.upsertCustomer_ = () => 2;
context.buildProductIndex_ = () => ({});
context.resolveProduct_ = () => ({
  SKU: "SKU-1",
  Brand: "Brand",
  Size: "Size",
  Name: "Server product",
  pack: 10,
  finalPrice: 200,
  unitCost: 120,
});
context.getNextOrderId_ = () => "OD0001";
context.nowIso_ = () => "2026-07-29T00:00:00.000Z";
context.appendOrder_ = (order) => {
  writtenOrder = order;
};
context.appendOrderItems_ = (_orderId, items) => {
  writtenItems = items;
};
context.safePushOrderConfirmation_ = () => ({
  attempted: false,
  ok: false,
});
context.sendN8NWebhook_ = () => ({ attempted: false, ok: false });

const result = vm.runInContext(
  `createOrder_(
    {
      store: "Override store",
      cart: [{ SKU: "SKU-1", qty: 2, price: 1 }]
    },
    {
      lineUserId: "${CUSTOMER_ID}",
      customerId: "${CUSTOMER_ID}",
      displayName: "Customer",
      tokenVerified: true,
      tokenError: "",
      isGuest: false
    },
    {
      createdByLineUserId: "${ADMIN_ID}",
      orderMode: "ADMIN",
      adminNotificationLineUserId: "${ADMIN_ID}"
    }
  )`,
  context,
);

assert.equal(writtenOrder.lineUserId, CUSTOMER_ID);
assert.equal(writtenOrder.customerId, CUSTOMER_ID);
assert.equal(writtenOrder.createdByLineUserId, ADMIN_ID);
assert.equal(writtenOrder.orderMode, "ADMIN");
assert.equal(writtenOrder.itemsTotal, 400);
assert.equal(writtenItems[0].unitPrice, 200);
assert.equal(writtenItems[0].Cost, 240);
assert.equal(writtenItems[0].Profit, 160);
assert.equal(result.orderMode, "ADMIN");

console.log("admin flow tests passed");
