# TaxTech ERP - Enterprise Business Management Suite

A modern, offline-first desktop ERP and accounting application built with **Electron**, **ExcelJS**, and **PDFKit**. It brings robust double-entry principles, 28 standard core accounting groups, and final accounts together with a fast, responsive user interface and automated GST compliance.

---

## 🌟 Key Features

### 1. Items & Services Master (Goods vs. Services)
- **Item Categorization**:
  - **Goods**: Track physical inventory with HSN Code, custom units of measurement (PCS, BOX, KGS, MTR, NOS), purchase cost, sale price, GST rate, and **Opening Stock**.
  - **Services**: Service-specific provision with SAC Code, non-inventory units (OTH, HRS, SAC), and zero stock overhead.
- **Visual Status & Type Badges**: Clean badges distinguishing Goods from Services across all search tables and dropdowns.
- **Filtering**: Tabbed views to filter between All Items, Goods Only, and Services Only.

### 2. Sales Invoicing & GST Tax Invoices
- **Automated Tax Calculation**:
  - **Intrastate Transactions**: Automatic 50/50 split between **CGST** and **SGST**.
  - **Interstate Transactions**: Automatic **IGST** levy based on the chosen Place of Supply vs. Company State.
- **Line Items**: Add Goods or Services with live calculation of taxable value, item discounts, tax breakdowns, and totals.
- **Discounts & Round-Off**: Built-in support for cash/trade discounts and automatic or manual round-off adjustments.
- **PDF Generation**: Generates professional, GST-compliant A4 Tax Invoices with:
  - Company header, GSTIN, PAN, and address.
  - Bill-To and Ship-To buyer details.
  - Itemized table with HSN/SAC, rates, and tax breakdown.
  - Amount in words (Indian Rupees).
  - Bank account details (Account No., IFSC, Bank Name, Branch).
  - Terms & conditions and authorized signatory block.

### 3. Purchases & Input Tax Credit (ITC)
- **Inward Purchase Bills**: Record purchases from registered or unregistered suppliers.
- **Input Tax Tracking**: Separate tracking for CGST, SGST, and IGST Input Tax Credits.
- **Supplier Invoice Mapping**: Store supplier bill numbers and reference dates alongside internal voucher numbers (`PUR/2024-25/0001`).
- **Inventory Integration**: Inward quantities automatically feed the Stock Register.

### 4. Stock Register & Inventory Valuation
- **Real-Time Stock Computation**:
  $$\text{Closing Stock} = \text{Opening Stock} + \text{Inward (Purchases)} - \text{Outward (Sales)}$$
- **Inventory Valuation**:
  $$\text{Closing Valuation} = \max(0, \text{Closing Stock}) \times \text{Cost Rate}$$
- **Stock Alert Badges**:
  - `In Stock` (Green)
  - `Low Stock` (Yellow - less than 5 units)
  - `Out of Stock` (Red - 0 or negative balance)
  - `Service` (Gray - non-inventory service items)
- **One-Click Export**: Export the live stock register to Excel.

### 5. Parties Master (Debtors & Creditors)
- **Party Roles**: Customer (Sundry Debtor), Supplier (Sundry Creditor), or Both.
- **Financial Controls**: Set Opening Balances (Dr/Cr), Credit Limits (₹), and Payment Terms (days).
- **GST Details**: Validation for 15-character GSTIN, PAN, State Code auto-fill, and full billing address.

### 6. Accounting Vouchers (Double-Entry Engine)
- **Receipt Voucher (`RCP/`)**: Money received into Cash or Bank accounts from Customers, Capital, or Other Incomes.
- **Payment Voucher (`PAY/`)**: Money paid out from Cash or Bank accounts to Suppliers, Rent, Salaries, or Utilities.
- **Contra Voucher (`CNT/`)**: Internal funds transfers between Cash and Bank (Cash Deposits / Bank Withdrawals).
- **Journal Voucher (`JRN/`)**: General journal adjustments between ledger accounts.

### 7. Cash Book & Bank Book
- **Dual Register View**: Instant switching between Cash Book and Bank Book.
- **Running Ledger**: Displays Date, Voucher No., Type, Particulars, Receipts (Debit), Payments (Credit), and live Running Balance.
- **Metrics Summary**: Real-time display of Opening Balance, Total Receipts, Total Payments, and Current Closing Balance.

### 8. Accounts & Ledgers (28 Standard Core Accounting Groups)
- **Pre-Seeded Core Groups**:
  - **15 Primary Groups**: *Capital Account, Loans (Liability), Current Liabilities, Fixed Assets, Current Assets, Investments, Misc. Expenses (Asset), Suspense Account, Branch / Divisions, Sales Accounts, Purchase Accounts, Direct Incomes, Indirect Incomes, Direct Expenses, Indirect Expenses*.
  - **13 Sub-Groups**: *Reserves & Surplus, Bank OD A/c, Secured Loans, Unsecured Loans, Duties & Taxes, Provisions, Sundry Creditors, Stock-in-hand, Deposits (Asset), Loans & Advances (Asset), Sundry Debtors, Bank Accounts, Cash-in-hand*.
- **Ledger Creation & Alteration**: Create custom ledgers categorized under any of the 28 groups with opening balances and Dr/Cr indicators.
- **System Default Ledgers**: Pre-configured defaults for Cash, Bank Account, Sales Account, Purchase Account, GST Inputs/Outputs, Capital Account, and typical office expenses.

### 9. Final Financial Statements
- **Trial Balance**:
  - Complete list of closing Debit and Credit balances across all ledgers.
  - Automatic debit/credit balance verification with a `Balanced (Dr = Cr)` indicator.
- **Trading Account (2-Column)**:
  - **Debit**: Opening Stock, Purchases, Direct Expenses, Gross Profit $c/d$.
  - **Credit**: Sales Turnover, Closing Stock, Direct Incomes, Gross Loss $c/d$.
- **Profit & Loss Account (2-Column)**:
  - **Debit**: Gross Loss $b/d$, Indirect Expenses (Rent, Salaries, Utilities), Net Profit $c/d$.
  - **Credit**: Gross Profit $b/d$, Indirect Incomes (Discounts, Interest), Net Loss $c/d$.
- **Balance Sheet (Standard Financial Schedule)**:
  - **Liabilities & Capital**: Capital Account ($+ \text{Net Profit}$ / $- \text{Net Loss}$), Loans, Sundry Creditors, Net GST Payable ($\text{Output} - \text{Input}$).
  - **Assets**: Fixed Assets, Closing Stock, Sundry Debtors, Bank Accounts, Cash in Hand, Net GST Credit.
  - Real-time `Balanced` status verification.

### 10. Local Excel Database, Storage & Backup
- **Zero Server Setup**: All application data is stored in a structured multi-sheet Excel file (`gst_billing_data.xlsx`).
- **Safety & Non-Destructive Migrations**: App launch auto-checks and creates missing sheets or columns without touching existing data.
- **One-Click Backup & Restore**: Create timestamped backup files (`gst_billing_backup_YYYY-MM-DD.xlsx`) or restore from any previous backup file.
- **Direct Excel Access**: Quick-open buttons to reveal the storage directory in Windows Explorer or open the workbook directly in Microsoft Excel.

---

## 📁 Database Schema (`gst_billing_data.xlsx`)

| Sheet Name | Description | Key Fields |
| :--- | :--- | :--- |
| `Products` | Goods and Services catalog | `id`, `name`, `type`, `hsn`, `unit`, `gstRate`, `salePrice`, `purchasePrice`, `openingStock` |
| `Clients` | Debtors & Creditors (Parties) | `id`, `name`, `partyType`, `gstin`, `pan`, `city`, `state`, `openingBalance`, `balanceType` |
| `Invoices` | Sales invoice headers | `id`, `invoiceNumber`, `date`, `clientId`, `subtotal`, `totalTax`, `grandTotal`, `status` |
| `InvoiceItems`| Sales invoice line items | `id`, `invoiceId`, `productId`, `quantity`, `rate`, `discountPercent`, `taxableValue`, `totalAmount` |
| `Purchases` | Inward purchase bills | `id`, `purchaseNumber`, `date`, `supplierId`, `supplierBillNo`, `subtotal`, `totalTax`, `grandTotal` |
| `PurchaseItems`| Purchase bill line items | `id`, `purchaseId`, `productId`, `quantity`, `rate`, `taxableValue`, `cgstAmount`, `totalAmount` |
| `LedgerGroups`| 28 standard core groups | `id`, `name`, `parentGroup`, `nature` (Assets, Liabilities, Income, Expenses), `isDefault` |
| `Ledgers` | Master accounting ledgers | `id`, `name`, `groupName`, `openingBalance`, `balanceType` (Dr/Cr), `isDefault` |
| `Vouchers` | Cash/Bank & Journal vouchers | `id`, `voucherNumber`, `date`, `type` (Receipt/Payment/Contra/Journal), `mode`, `debitLedger`, `creditLedger`, `amount` |
| `CompanySettings`| Firm & Bank configuration | `key`, `value` (Company Name, GSTIN, Bank details, Invoice Prefix, Series) |

---

## 🛠️ Tech Stack & Architecture

- **Runtime**: [Electron](https://www.electronjs.org/) (Chromium + Node.js)
- **Workbook Engine**: [ExcelJS](https://github.com/exceljs/exceljs) (xlsx generation, reading, writing, styling)
- **Document Engine**: [PDFKit](https://pdfkit.org/) (Vector-accurate A4 Tax Invoices)
- **Identifiers**: [UUID v4](https://github.com/uuidjs/uuid)
- **UI Architecture**: Vanilla ES6+ JavaScript, Responsive CSS3 Grid/Flexbox, contextBridge isolation with secure IPC communication.

---

## 🚀 Getting Started

### Prerequisites
- [Node.js](https://nodejs.org/) (v16.0.0 or higher)
- Windows, macOS, or Linux

### Installation

1. Clone or copy the project into your local directory:
   ```bash
   cd "C:\Users\babee\Desktop\Free Claude Code\GSTBillingApp"
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Launch the application:
   ```bash
   npm start
   ```

4. Launch with Developer Tools enabled:
   ```bash
   npm run dev
   ```

---

## 📂 Project Structure

```
GSTBillingApp/
├── index.html         # Main application UI layout, views, and modal dialogs
├── main.js            # Electron main process, Excel storage engine, calculations & IPC
├── preload.js         # Context bridge exposing secure APIs to renderer
├── renderer.js        # Renderer logic, state management, event listeners, calculations
├── styles.css         # Styling, themes, responsive tables, modal overlays & highlights
├── package.json       # Project manifest and dependencies
└── README.md          # Documentation
```

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
| :--- | :--- |
| `Ctrl + N` | Open New Sales Invoice modal |
| `Escape` | Close any open modal dialog |

---

## 🔒 Data Security & Offline Privacy

All database files and PDF invoices remain **100% on your local computer**:
- **Default Database Location**: `%APPDATA%\gst-billing-app\gst-billing-data\gst_billing_data.xlsx`
- **Zero Cloud Dependence**: The application functions completely offline without requiring internet access or third-party servers.

---

## 📄 License

ISC License. Free for commercial and personal business use.
