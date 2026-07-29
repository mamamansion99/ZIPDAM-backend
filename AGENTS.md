# ZIPDAM Backend Codex Instructions

The connected frontend repository is `mamamansion99/ZIPDAM`.
The full system and spreadsheet contract is documented in `mamamansion99/ZIPDAM/docs/ZIPDAM_CODEX_CONTEXT.md`.

## Required behaviour

- `customerId` is the real LINE `lineUserId`.
- Same LINE ID always means the same customer, even if the display name changes.
- Do not count `GUEST-*` identities in loyalty or lifetime-spend totals.
- Write to Google Sheets by exact header name; never use fixed column positions.
- Current Orders schema intentionally has no `paymentStatus` or `paidAt`.
- Do not delete, rename or overwrite sales-report sheets.
- Keep secrets in Apps Script Properties. Never hard-code LINE tokens or webhook secrets.
- Keep order writes compatible with Customer, Orders, OrderItems, Favorites, Templates and Rewards.
- Use the `health` endpoint after changing sheet contracts.

Before modifying `Code.js`, verify the current spreadsheet headers and preserve backward compatibility with historical Orders and OrderItems.