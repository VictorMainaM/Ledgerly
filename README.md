# Ledgerly

Ledgerly is a local-first business recordkeeping app for transactions, inventory, production, transport, account balances, and reports.

## Run the app

Open `business-ledger.html` in Microsoft Edge or Chrome. Your records are stored in the browser on this device. Use **Documents & reports → Export data** regularly to create a backup.

## Separate companies

Use the **Company** switcher in the sidebar to add a company or switch between existing companies. Existing Ledgerly data is kept as the first company when this version is opened. Each company has a separate business profile, transactions, inventory, production batches, transport records, opening balances, and Excel sync history. Adding a company starts with an empty inventory and ledger; it does not change any other company's records.

Company profiles are stored locally in the same browser profile. They are separate records, not separate logins or password-protected accounts. Back up each company's records separately using **Export data** while that company is selected.

## Excel synchronization (Windows)

Synchronization is manual and requires desktop Microsoft Excel. Start `start-ledgerly-sync.cmd`, then open **Sync Excel** for the selected company. Assign each company a different full Windows path ending in `.xlsx`. Save workbook edits before syncing and resolve any conflicts shown by Ledgerly. You can have Ledgerly create a new workbook at the chosen path if it does not exist. The connector preserves existing workbook tabs.

## Included files

- `business-ledger.html` — the application
- `ledgerly-sync.ps1` and `start-ledgerly-sync.cmd` — the local Excel connector
- `ledgerly-icon.ico` and `ledgerly-icon.svg` — app icons
- `README.doc` — detailed usage and installation guide
- `Documentation/User Terms and Conditions.doc` — user terms
