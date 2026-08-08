const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "..", "Code.js"), "utf8");
const context = vm.createContext({ console });
vm.runInContext(source, context, { filename: "Code.js" });

const reward = {
  rewardId: "GIFT-10000",
  rewardName: "ของแถมยอดสะสม 10,000",
  rewardValue: "ของแถม 1 รายการ",
  requiredSpend: 10000,
};

context.getActiveSpendReward_ = () => reward;
context.calculateEligibleProductSpend_ = () => 127038;

const progress = context.safeLoyaltyProgress_(`U${"a".repeat(32)}`);
assert.equal(progress.eligible, true);
assert.equal(progress.totalSpend, 127038);
assert.equal(progress.targetSpend, 10000);
assert.equal(progress.cycleSpend, 7038);
assert.equal(progress.remainingSpend, 2962);
assert.equal(progress.progressPercent, 70);
assert.equal(progress.achievedCycles, 12);
assert.equal(progress.nextCycle, 13);

// A customer who has never ordered starts at cycle 1 with an empty bar.
context.calculateEligibleProductSpend_ = () => 0;
const fresh = context.safeLoyaltyProgress_(`U${"b".repeat(32)}`);
assert.equal(fresh.cycleSpend, 0);
assert.equal(fresh.remainingSpend, 10000);
assert.equal(fresh.progressPercent, 0);
assert.equal(fresh.nextCycle, 1);

// Reading progress must never grant a cycle or write to the ledger.
let wrote = false;
context.ensureLoyaltyRewardCycles_ = () => { wrote = true; return [1]; };
context.appendObjectRow_ = () => { wrote = true; };
context.calculateEligibleProductSpend_ = () => 20000;
const exact = context.safeLoyaltyProgress_(`U${"c".repeat(32)}`);
assert.equal(wrote, false);
assert.equal(exact.cycleSpend, 0);
assert.equal(exact.achievedCycles, 2);
assert.equal(exact.nextCycle, 3);

// No configured reward degrades to "not eligible" instead of throwing.
context.getActiveSpendReward_ = () => null;
assert.equal(context.safeLoyaltyProgress_(`U${"d".repeat(32)}`).eligible, false);

// A broken Rewards sheet must not take the whole summary down.
context.getActiveSpendReward_ = () => { throw new Error("Rewards sheet missing"); };
const failed = context.safeLoyaltyProgress_(`U${"e".repeat(32)}`);
assert.equal(failed.eligible, false);
assert.match(failed.error, /Rewards sheet missing/);

console.log("loyalty progress tests passed");
