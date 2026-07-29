const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "Code.js"), "utf8");
const context = vm.createContext({ console });
vm.runInContext(source, context, { filename: "Code.js" });

const loyalty = {
  eligible: true,
  rewardId: "GIFT-10000",
  rewardName: "ของแถมยอดสะสม 10,000",
  rewardValue: "ของแถม 1 รายการ",
  totalSpend: 7650,
  targetSpend: 10000,
  cycleSpend: 7650,
  remainingSpend: 2350,
  progressPercent: 76,
  rewardEarned: false,
  earnedCycles: [],
  nextCycle: 1,
};
const order = {
  orderId: "OD0054",
  displayName: "ZIPDAM Customer",
  address: "Bangkok",
  phone: "0812345678",
  items: [
    {
      Name: "Durex Airy",
      Size: "52 mm",
      qty: 2,
      unitPrice: 200,
      lineTotal: 400,
    },
    {
      Name: "Okamoto 003",
      Size: "52 mm",
      qty: 1,
      unitPrice: 250,
      lineTotal: 250,
    },
  ],
  itemsTotal: 650,
  shippingFee: 20,
  grandTotal: 670,
  loyalty,
};

const flex = context.buildOrderConfirmationFlex_(order);
assert.equal(flex.type, "flex");
assert.match(flex.altText, /OD0054/);
assert.equal(flex.contents.type, "bubble");
assert.equal(flex.contents.styles.body.backgroundColor, "#FFFFFF");

const serialized = JSON.stringify(flex);
assert.match(serialized, /ยอดสั่งซื้อสะสม/);
assert.match(serialized, /อีก ฿2,350/);
assert.match(serialized, /Durex Airy/);
assert.match(serialized, /"width":"76%"/);

const earnedFlex = context.buildOrderConfirmationFlex_({
  ...order,
  loyalty: {
    ...loyalty,
    totalSpend: 10000,
    cycleSpend: 10000,
    remainingSpend: 0,
    progressPercent: 100,
    rewardEarned: true,
    earnedCycles: [1],
  },
});
assert.match(JSON.stringify(earnedFlex), /ครบยอดแล้ว/);

let pushedPayload = null;
context.UrlFetchApp = {
  fetch(_url, options) {
    pushedPayload = JSON.parse(options.payload);
    return {
      getResponseCode: () => 200,
      getContentText: () => "{}",
    };
  },
};
const pushResult = context.safePushOrderConfirmation_(
  "token",
  `U${"a".repeat(32)}`,
  order.orderId,
  order.items,
  order.itemsTotal,
  order.shippingFee,
  order.grandTotal,
  order.displayName,
  order.address,
  order.phone,
  loyalty,
);
assert.equal(pushResult.ok, true);
assert.equal(pushedPayload.messages[0].type, "flex");
assert.equal(pushedPayload.messages[0].contents.type, "bubble");

context.getActiveSpendReward_ = () => ({
  rewardId: "GIFT-10000",
  rewardName: "ของแถมยอดสะสม 10,000",
  rewardValue: "ของแถม 1 รายการ",
  requiredSpend: 10000,
});
context.calculateEligibleProductSpend_ = () => 12450;
context.ensureLoyaltyRewardCycles_ = () => [1];
let rewardApplied = "";
context.updateOrderRewardApplied_ = (_orderId, value) => {
  rewardApplied = value;
};
const applied = context.safeApplyLoyaltyAfterOrder_(
  `U${"b".repeat(32)}`,
  "OD0055",
  `U${"a".repeat(32)}`,
);
assert.equal(applied.totalSpend, 12450);
assert.equal(applied.rewardEarned, true);
assert.equal(applied.progressPercent, 100);
assert.equal(rewardApplied, "GIFT-10000#1");

console.log("flex message tests passed");
