# Ledgerly

Ledgerly is a local-first business recordkeeping app for sales, purchases, expenses, inventory, production, transport, account balances, and reports. It runs in a browser and does not require an online account.

## Run the app

1. Download or extract the Ledgerly files to a folder on a Windows computer.
2. Open `business-ledger.html` in a browser.
3. Enter your business details, then record transactions and inventory.
4. Use **Export data** regularly to save a backup. Records are stored in the browser on that device; clearing browser data or switching browsers can remove or hide them.

## Opening balances

In **Documents & reports**, choose **Set opening balances**. Set the date the balances apply from, then add cash/M-Pesa, bank, customer debts, supplier debts, loans, or capital. The report shows balances brought down (b/d), movements, and balances carried down (c/d). These opening balances and account movements do not count as profit or loss. Use consistent customer, supplier, loan, and capital account names in later transactions.

## Excel synchronization (Windows)

Synchronization is manual and requires desktop Microsoft Excel:

1. Open the workbook in Excel and save a separate backup.
2. Run `start-ledgerly-sync.cmd` from the Ledgerly folder and keep its window open.
3. In Ledgerly, select **Sync**, enter the one-time token shown by the connector, and start the sync.
4. Save Excel edits before syncing. Review any conflicts and choose which version to keep.

Ledgerly uses editable `LEDGERLY TRANSACTIONS`, `LEDGERLY INVENTORY`, `LEDGERLY TRANSPORT`, `LEDGERLY PRODUCTION`, and `LEDGERLY OPENING BALANCES` sheets. Existing workbook sheets are preserved. The app and workbook are not continuously synchronized; run the sync when you want to exchange changes.

## Included files

- `business-ledger.html` — the application
- `ledgerly-sync.ps1` and `start-ledgerly-sync.cmd` — the local Excel connector
- `ledgerly-icon.ico` and `ledgerly-icon.svg` — app icons
- `README.doc` — detailed usage and installation guide
- `Documentation/User Terms and Conditions.doc` — user terms
