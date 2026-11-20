# TaxTech ERP - Comprehensive Upgrade Roadmap & Progress Tracker

**Document Version:** 1.0.0  
**Application Title:** TaxTech ERP (Enterprise Business Management Suite)  
**Author:** Senior ERP Software Architect, Chartered Accountant & Full-Stack Team  
**Architecture:** Electron | Node.js | Vanilla ES6+ | ExcelJS | PDFKit | Local-First Storage  
**Last Updated:** November 2026  

---

## 1. Executive Summary & Existing Codebase Audit

### 1.1 Existing Architecture Overview
The current application is an Electron-based desktop solution providing GST billing and foundational accounting:
- **Main Process (`main.js`)**: Manages BrowserWindow, Excel workbook reading and writing via `ExcelJS`, PDF invoice generation via `PDFKit`, and IPC handlers.
- **Preload Bridge (`preload.js`)**: Context isolation with selective API exposure on `window.api` and `window.ipcRenderer`.
- **Renderer (`renderer.js`)**: Client-side state management, table rendering, automated tax/discount/stock calculations, and modal workflows.
- **User Interface (`index.html` & `styles.css`)**: Responsive layout with sidebar navigation, modal dialogs, and financial statement cards.
- **Data Persistence (`gst_billing_data.xlsx`)**: Stored in `%APPDATA%/gst-billing-app/gst-billing-data/gst_billing_data.xlsx` with sheets: `Products`, `Clients`, `Invoices`, `InvoiceItems`, `Purchases`, `PurchaseItems`, `LedgerGroups`, `Ledgers`, `Vouchers`, and `CompanySettings`.

### 1.2 Identified Bugs & Limitations (Audit Findings)
1. **Branding & Terminology**: Multiple titles reference "Tally ERP" and "GST Billing App". User requirement specifically requests renaming the software to **TaxTech ERP** and removing "Tally ERP" branding everywhere while retaining strict double-entry mechanics.
2. **Item & Party Selection in Transactions**: In Sales and Purchase entries, dropdowns are standard static `<select>` elements. For real-world catalogues (thousands of SKUs and parties), a **searchable, filterable combobox/dropdown** with live HSN, rate, and stock display is required.
3. **Data Scaling & Concurrency**: Writing to an Excel workbook on every voucher creates bottlenecks at high transaction volumes (10,000+ entries). An in-memory indexed relational store with transactional logging and Excel synchronization is required.
4. **Missing Modules**:
   - Multi-Company & Branch Management
   - Quotations, Sales Orders, Delivery Challans, Credit Notes, and Debit Notes
   - Purchase Orders and Goods Receipt Notes (GRN)
   - Godown / Multi-Warehouse Inventory Management with Batch & Expiry tracking
   - Dedicated Indian GST Statutory Suite (GSTR-1, GSTR-3B, GSTR-2B Reconciliation)
   - TDS & TCS Deduction, Payable/Receivable ledgers, and Section 194C/194J/194I/194Q compliance
   - Fixed Asset Register with SLM and WDV Depreciation Schedules
   - Bank Reconciliation Statement (BRS) with CSV/Excel statement matching
   - Mandatory MCA Audit Trail (Activity Logs, Alteration History, Exception Alerts)
   - Role-Based Access Control (Admin, Accountant, Auditor, Sales, Purchase)
   - Light Mode / Dark Mode Theme Switcher

---

## 2. Implementation Phasing Strategy

```
Phase 1: Code Audit & Architecture Roadmap [COMPLETED]
   ├── Comprehensive source code review
   ├── Bug identification & inventory of missing modules
   └── Data schema & backward compatibility design

Phase 2: UI/UX Modernisation & Core Rebranding [ACTIVE]
   ├── Rebranding to "TaxTech ERP" (removing all Tally ERP branding)
   ├── Modern expandable sidebar navigation with modular grouping
   ├── Dark Mode & Light Mode Theme Engine
   ├── Searchable Item & Party dropdowns for Sales and Purchases
   ├── Dashboard KPI widgets with Visual Analytics & Period Filters
   └── Breadcrumbs, table pagination & enhanced toast/dialog notifications

Phase 3: Accounting Engine & Voucher Suite
   ├── Double-Entry Enforcement ($\sum Dr = \sum Cr$ validation on every post)
   ├── Expanded Voucher Types: Quotations, Sales Orders, Delivery Challans, Credit Notes, Debit Notes, GRN, Stock Journals
   ├── Drill-down Ledger Reports, Day Book, Cash Book, Bank Book
   └── Financial Statements: Trial Balance, Trading Account, P&L, Balance Sheet

Phase 4: Multi-Company, Warehouses & Inventory
   ├── Multi-Company Switcher & Branch Management
   ├── Multiple Warehouses / Godowns & Stock Transfers
   ├── Batch Management, Expiry Tracking & Serial Numbers
   └── Reorder Level Alerts, Negative Stock Warnings & FIFO/Weighted Valuation

Phase 5: Indian Statutory Compliance (GST, TDS, TCS)
   ├── GSTR-1 Tables (B2B, B2CL, B2CS, CDNR, CDNUR, HSN Summary)
   ├── GSTR-3B Monthly Working & Tax Summary
   ├── GSTR-2B vs. Purchase Register Reconciliation
   └── TDS/TCS Tracking (Sections 194C, 194J, 194I, 194Q, 206C)

Phase 6: Banking, Fixed Assets, Audit Trail & RBAC
   ├── Bank Reconciliation Statement (BRS) & CSV Statement Import
   ├── Fixed Asset Register & Depreciation (Companies Act & Income Tax Act)
   ├── Mandatory Audit Trail (User Activity, Transaction Alterations, Backdated Entries)
   └── Role-Based Access Control (Admin, Accountant, Auditor Read-Only, Sales)

Phase 7: Packaging, Verification & Production Build
   ├── Full Automated Regression & Accounting Verification Tests
   ├── Data Migration & Reconciliation Checks
   └── Windows Installer Packaging & Distribution Setup
```

---

## 3. Data Schema & Migration Architecture

### 3.1 Data Integrity Principles
- **Zero Data Loss Guarantee**: Existing records in `gst_billing_data.xlsx` (`Products`, `Clients`, `Invoices`, `InvoiceItems`, `Purchases`, `PurchaseItems`, `Ledgers`, `Vouchers`, `CompanySettings`) will never be modified or deleted without explicit confirmation.
- **Non-Destructive Schema Extension**: New sheets and columns (`Godowns`, `Batches`, `FixedAssets`, `AuditTrail`, `Companies`, `Quotations`, `CreditDebitNotes`, `TdsTransactions`) are created dynamically if not present.
- **Audit Logging**: Every voucher insertion, alteration, and cancellation logs a timestamped record in the Audit Trail with prior and updated values.

---

## 4. Current Progress & Milestones

| Milestone / Task | Status | Notes |
| :--- | :--- | :--- |
| **Phase 1: Code Audit & Architecture Roadmap** | **Completed** | Full analysis of existing codebase, constraints, and requirements completed. |
| **Rebrand to "TaxTech ERP"** | **Completed** | Removed all "Tally ERP" references across UI, titles, and text; preserved strict double-entry mechanics. |
| **Searchable Dropdown for Line Items (Sales & Purchases)** | **Completed** | Full searchable combobox with live SKU/HSN/stock/price indicators, keyboard navigation, and custom item fallback. |
| **Add Party Area Input Clickability & Highlight** | **Completed** | Modal stacking bug resolved; high-contrast focus outline and label highlights implemented. |
| **Dark Mode & Light Mode Theme Engine** | **Completed** | Enterprise CSS variable palette with smooth top-bar Sun/Moon toggle and persistence. |
| **Automated Integration Test Suite** | **Completed** | Electron `--test` runner with 10 automated checks passing (exit code 0). |
| **Phase 3: Double-Entry Accounting Engine & Vouchers** | **Next** | Multi-voucher balance verification, debit/credit notes, day book. |
| **Phase 4: Multi-Company, Warehouses & Inventory** | Planned | Godown stock transfers, batch tracking, reorder levels. |
| **Phase 5: Indian Statutory Compliance (GST, TDS, TCS)** | Planned | GSTR-1, GSTR-3B, GSTR-2B reconciliation, Section 194 TDS. |
| **Phase 6: Banking, Fixed Assets, Audit Trail & RBAC** | Planned | BRS reconciliation, SLM/WDV depreciation, MCA audit trail. |

---

*This progress tracker will be updated continuously as each phase is completed and verified.*
