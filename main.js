const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { v4: uuidv4 } = require('uuid');

let DATA_DIR;
let EXCEL_FILE;

const SHEET_COLUMNS = {
  Products: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Name', key: 'name', width: 30 },
    { header: 'Type', key: 'type', width: 12 },
    { header: 'HSN/SAC Code', key: 'hsn', width: 15 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'GST Rate (%)', key: 'gstRate', width: 15 },
    { header: 'CGST (%)', key: 'cgst', width: 12 },
    { header: 'SGST (%)', key: 'sgst', width: 12 },
    { header: 'IGST (%)', key: 'igst', width: 12 },
    { header: 'Sale Price', key: 'salePrice', width: 15 },
    { header: 'Purchase Price', key: 'purchasePrice', width: 18 },
    { header: 'Opening Stock', key: 'openingStock', width: 15 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Created At', key: 'createdAt', width: 20 },
    { header: 'Updated At', key: 'updatedAt', width: 20 }
  ],
  Clients: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Name', key: 'name', width: 30 },
    { header: 'Party Type', key: 'partyType', width: 15 },
    { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'PAN', key: 'pan', width: 15 },
    { header: 'Address Line 1', key: 'address1', width: 30 },
    { header: 'Address Line 2', key: 'address2', width: 30 },
    { header: 'City', key: 'city', width: 20 },
    { header: 'State', key: 'state', width: 20 },
    { header: 'State Code', key: 'stateCode', width: 15 },
    { header: 'Pincode', key: 'pincode', width: 12 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Email', key: 'email', width: 25 },
    { header: 'Contact Person', key: 'contactPerson', width: 20 },
    { header: 'Credit Limit', key: 'creditLimit', width: 15 },
    { header: 'Payment Terms', key: 'paymentTerms', width: 15 },
    { header: 'Opening Balance', key: 'openingBalance', width: 18 },
    { header: 'Balance Type', key: 'balanceType', width: 12 },
    { header: 'Created At', key: 'createdAt', width: 20 },
    { header: 'Updated At', key: 'updatedAt', width: 20 }
  ],
  Invoices: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Invoice Number', key: 'invoiceNumber', width: 20 },
    { header: 'Date', key: 'date', width: 15 },
    { header: 'Client ID', key: 'clientId', width: 38 },
    { header: 'Client Name', key: 'clientName', width: 30 },
    { header: 'Client GSTIN', key: 'clientGstin', width: 20 },
    { header: 'Place of Supply', key: 'placeOfSupply', width: 20 },
    { header: 'Reverse Charge', key: 'reverseCharge', width: 15 },
    { header: 'Subtotal', key: 'subtotal', width: 15 },
    { header: 'Total CGST', key: 'totalCgst', width: 15 },
    { header: 'Total SGST', key: 'totalSgst', width: 15 },
    { header: 'Total IGST', key: 'totalIgst', width: 15 },
    { header: 'Total Tax', key: 'totalTax', width: 15 },
    { header: 'Discount', key: 'discount', width: 12 },
    { header: 'Round Off', key: 'roundOff', width: 12 },
    { header: 'Grand Total', key: 'grandTotal', width: 15 },
    { header: 'Status', key: 'status', width: 15 },
    { header: 'Notes', key: 'notes', width: 30 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ],
  InvoiceItems: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Invoice ID', key: 'invoiceId', width: 38 },
    { header: 'Product ID', key: 'productId', width: 38 },
    { header: 'Product Name', key: 'productName', width: 30 },
    { header: 'Type', key: 'type', width: 12 },
    { header: 'HSN/SAC Code', key: 'hsn', width: 15 },
    { header: 'Quantity', key: 'quantity', width: 12 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'Rate', key: 'rate', width: 12 },
    { header: 'Discount %', key: 'discountPercent', width: 12 },
    { header: 'Discount Amt', key: 'discountAmount', width: 15 },
    { header: 'Taxable Value', key: 'taxableValue', width: 15 },
    { header: 'GST Rate', key: 'gstRate', width: 12 },
    { header: 'CGST %', key: 'cgstPercent', width: 10 },
    { header: 'CGST Amt', key: 'cgstAmount', width: 12 },
    { header: 'SGST %', key: 'sgstPercent', width: 10 },
    { header: 'SGST Amt', key: 'sgstAmount', width: 12 },
    { header: 'IGST %', key: 'igstPercent', width: 10 },
    { header: 'IGST Amt', key: 'igstAmount', width: 12 },
    { header: 'Total Amount', key: 'totalAmount', width: 15 }
  ],
  Purchases: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Purchase Number', key: 'purchaseNumber', width: 20 },
    { header: 'Date', key: 'date', width: 15 },
    { header: 'Supplier ID', key: 'supplierId', width: 38 },
    { header: 'Supplier Name', key: 'supplierName', width: 30 },
    { header: 'Supplier GSTIN', key: 'supplierGstin', width: 20 },
    { header: 'Supplier Bill No', key: 'supplierBillNo', width: 20 },
    { header: 'Place of Supply', key: 'placeOfSupply', width: 20 },
    { header: 'Subtotal', key: 'subtotal', width: 15 },
    { header: 'Total CGST', key: 'totalCgst', width: 15 },
    { header: 'Total SGST', key: 'totalSgst', width: 15 },
    { header: 'Total IGST', key: 'totalIgst', width: 15 },
    { header: 'Total Tax', key: 'totalTax', width: 15 },
    { header: 'Discount', key: 'discount', width: 12 },
    { header: 'Round Off', key: 'roundOff', width: 12 },
    { header: 'Grand Total', key: 'grandTotal', width: 15 },
    { header: 'Status', key: 'status', width: 15 },
    { header: 'Notes', key: 'notes', width: 30 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ],
  PurchaseItems: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Purchase ID', key: 'purchaseId', width: 38 },
    { header: 'Product ID', key: 'productId', width: 38 },
    { header: 'Product Name', key: 'productName', width: 30 },
    { header: 'Type', key: 'type', width: 12 },
    { header: 'HSN/SAC Code', key: 'hsn', width: 15 },
    { header: 'Quantity', key: 'quantity', width: 12 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'Rate', key: 'rate', width: 12 },
    { header: 'Discount %', key: 'discountPercent', width: 12 },
    { header: 'Discount Amt', key: 'discountAmount', width: 15 },
    { header: 'Taxable Value', key: 'taxableValue', width: 15 },
    { header: 'GST Rate', key: 'gstRate', width: 12 },
    { header: 'CGST %', key: 'cgstPercent', width: 10 },
    { header: 'CGST Amt', key: 'cgstAmount', width: 12 },
    { header: 'SGST %', key: 'sgstPercent', width: 10 },
    { header: 'SGST Amt', key: 'sgstAmount', width: 12 },
    { header: 'IGST %', key: 'igstPercent', width: 10 },
    { header: 'IGST Amt', key: 'igstAmount', width: 12 },
    { header: 'Total Amount', key: 'totalAmount', width: 15 }
  ],
  LedgerGroups: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Group Name', key: 'name', width: 30 },
    { header: 'Parent Group', key: 'parentGroup', width: 30 },
    { header: 'Nature', key: 'nature', width: 15 },
    { header: 'Is Default', key: 'isDefault', width: 12 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ],
  Ledgers: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Ledger Name', key: 'name', width: 30 },
    { header: 'Group Name', key: 'groupName', width: 30 },
    { header: 'Opening Balance', key: 'openingBalance', width: 18 },
    { header: 'Balance Type', key: 'balanceType', width: 12 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Is Default', key: 'isDefault', width: 12 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ],
  Vouchers: [
    { header: 'ID', key: 'id', width: 38 },
    { header: 'Voucher Number', key: 'voucherNumber', width: 20 },
    { header: 'Date', key: 'date', width: 15 },
    { header: 'Voucher Type', key: 'type', width: 15 },
    { header: 'Payment Mode', key: 'mode', width: 12 },
    { header: 'Debit Ledger', key: 'debitLedger', width: 30 },
    { header: 'Credit Ledger', key: 'creditLedger', width: 30 },
    { header: 'Amount', key: 'amount', width: 15 },
    { header: 'Reference No', key: 'referenceNo', width: 20 },
    { header: 'Narration', key: 'narration', width: 35 },
    { header: 'Created At', key: 'createdAt', width: 20 }
  ],
  CompanySettings: [
    { header: 'Key', key: 'key', width: 30 },
    { header: 'Value', key: 'value', width: 60 }
  ]
};

// 28 Standard Core Accounting Groups
const STANDARD_GROUPS = [
  // 15 Primary Groups
  { id: 'grp-1', name: 'Capital Account', parentGroup: '', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-2', name: 'Loans (Liability)', parentGroup: '', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-3', name: 'Current Liabilities', parentGroup: '', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-4', name: 'Fixed Assets', parentGroup: '', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-5', name: 'Current Assets', parentGroup: '', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-6', name: 'Investments', parentGroup: '', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-7', name: 'Misc. Expenses (Asset)', parentGroup: '', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-8', name: 'Suspense Account', parentGroup: '', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-9', name: 'Branch / Divisions', parentGroup: '', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-10', name: 'Sales Accounts', parentGroup: '', nature: 'Income', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-11', name: 'Purchase Accounts', parentGroup: '', nature: 'Expenses', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-12', name: 'Direct Incomes', parentGroup: '', nature: 'Income', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-13', name: 'Indirect Incomes', parentGroup: '', nature: 'Income', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-14', name: 'Direct Expenses', parentGroup: '', nature: 'Expenses', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-15', name: 'Indirect Expenses', parentGroup: '', nature: 'Expenses', isDefault: 'Yes', createdAt: new Date().toISOString() },

  // 13 Secondary / Sub-Groups
  { id: 'grp-16', name: 'Reserves & Surplus', parentGroup: 'Capital Account', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-17', name: 'Bank OD A/c', parentGroup: 'Loans (Liability)', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-18', name: 'Secured Loans', parentGroup: 'Loans (Liability)', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-19', name: 'Unsecured Loans', parentGroup: 'Loans (Liability)', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-20', name: 'Duties & Taxes', parentGroup: 'Current Liabilities', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-21', name: 'Provisions', parentGroup: 'Current Liabilities', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-22', name: 'Sundry Creditors', parentGroup: 'Current Liabilities', nature: 'Liabilities', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-23', name: 'Stock-in-hand', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-24', name: 'Deposits (Asset)', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-25', name: 'Loans & Advances (Asset)', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-26', name: 'Sundry Debtors', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-27', name: 'Bank Accounts', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'grp-28', name: 'Cash-in-hand', parentGroup: 'Current Assets', nature: 'Assets', isDefault: 'Yes', createdAt: new Date().toISOString() }
];

// Standard Seed Ledgers
const STANDARD_LEDGERS = [
  { id: 'led-1', name: 'Cash', groupName: 'Cash-in-hand', openingBalance: 10000, balanceType: 'Dr', description: 'Primary Cash in Hand', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-2', name: 'Bank Account', groupName: 'Bank Accounts', openingBalance: 50000, balanceType: 'Dr', description: 'Primary Bank Current/Savings Account', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-3', name: 'Sales Account', groupName: 'Sales Accounts', openingBalance: 0, balanceType: 'Cr', description: 'Domestic Sales Account', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-4', name: 'Purchase Account', groupName: 'Purchase Accounts', openingBalance: 0, balanceType: 'Dr', description: 'Domestic Purchases Account', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-5', name: 'CGST Output', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Cr', description: 'CGST Output on Sales', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-6', name: 'SGST Output', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Cr', description: 'SGST Output on Sales', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-7', name: 'IGST Output', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Cr', description: 'IGST Output on Sales', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-8', name: 'CGST Input', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Dr', description: 'CGST Input Tax Credit on Purchases', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-9', name: 'SGST Input', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Dr', description: 'SGST Input Tax Credit on Purchases', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-10', name: 'IGST Input', groupName: 'Duties & Taxes', openingBalance: 0, balanceType: 'Dr', description: 'IGST Input Tax Credit on Purchases', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-11', name: 'Capital Account', groupName: 'Capital Account', openingBalance: 60000, balanceType: 'Cr', description: 'Owner / Partner Capital', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-12', name: 'Office Rent', groupName: 'Indirect Expenses', openingBalance: 0, balanceType: 'Dr', description: 'Office & Premises Rent Expense', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-13', name: 'Salary & Wages', groupName: 'Indirect Expenses', openingBalance: 0, balanceType: 'Dr', description: 'Employee Salaries & Staff Wages', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-14', name: 'Electricity Expenses', groupName: 'Indirect Expenses', openingBalance: 0, balanceType: 'Dr', description: 'Office Electricity & Utility Charges', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-15', name: 'Telephone & Internet', groupName: 'Indirect Expenses', openingBalance: 0, balanceType: 'Dr', description: 'Communication & Internet Charges', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-16', name: 'Discount Received', groupName: 'Indirect Incomes', openingBalance: 0, balanceType: 'Cr', description: 'Discounts Received from Suppliers', isDefault: 'Yes', createdAt: new Date().toISOString() },
  { id: 'led-17', name: 'Interest Received', groupName: 'Indirect Incomes', openingBalance: 0, balanceType: 'Cr', description: 'Bank & Investment Interest', isDefault: 'Yes', createdAt: new Date().toISOString() }
];

const DEFAULT_SETTINGS = [
  { key: 'companyName', value: 'My Company Pvt Ltd' },
  { key: 'companyGstin', value: '29ABCDE1234F1Z5' },
  { key: 'companyPan', value: 'ABCDE1234F' },
  { key: 'companyAddress1', value: '123 Business Street' },
  { key: 'companyAddress2', value: 'Commercial Complex' },
  { key: 'companyCity', value: 'Bangalore' },
  { key: 'companyState', value: 'Karnataka' },
  { key: 'companyStateCode', value: '29' },
  { key: 'companyPincode', value: '560001' },
  { key: 'companyPhone', value: '+91-9876543210' },
  { key: 'companyEmail', value: 'info@mycompany.com' },
  { key: 'companyWebsite', value: 'www.mycompany.com' },
  { key: 'bankName', value: 'State Bank of India' },
  { key: 'bankAccount', value: '1234567890' },
  { key: 'bankIfsc', value: 'SBIN0001234' },
  { key: 'bankBranch', value: 'Main Branch' },
  { key: 'invoicePrefix', value: 'INV' },
  { key: 'invoiceStartNumber', value: '1' },
  { key: 'invoiceSeries', value: '2024-25' },
  { key: 'termsConditions', value: 'Payment due within 30 days. Goods once sold will not be taken back. Subject to Bangalore jurisdiction.' }
];

function ensureDataDir() {
  if (!DATA_DIR) {
    DATA_DIR = path.join(app.getPath('userData'), 'gst-billing-data');
    EXCEL_FILE = path.join(DATA_DIR, 'gst_billing_data.xlsx');
  }
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

async function initializeExcel() {
  ensureDataDir();
  const workbook = new ExcelJS.Workbook();
  const fileExisted = fs.existsSync(EXCEL_FILE);
  if (fileExisted) {
    try {
      await workbook.xlsx.readFile(EXCEL_FILE);
    } catch (e) {
      console.warn('Could not read existing Excel file, creating fresh:', e);
    }
  }

  let modified = false;

  for (const [sheetName, columns] of Object.entries(SHEET_COLUMNS)) {
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      sheet = workbook.addWorksheet(sheetName);
      sheet.columns = columns;
      modified = true;

      if (sheetName === 'CompanySettings') {
        sheet.addRows(DEFAULT_SETTINGS);
      } else if (sheetName === 'LedgerGroups') {
        sheet.addRows(STANDARD_GROUPS);
      } else if (sheetName === 'Ledgers') {
        sheet.addRows(STANDARD_LEDGERS);
      }
    } else {
      // Ensure columns/headers are up-to-date
      if (!fileExisted || sheet.rowCount <= 1) {
        if (sheetName === 'LedgerGroups' && sheet.rowCount <= 1) {
          sheet.columns = columns;
          sheet.addRows(STANDARD_GROUPS);
          modified = true;
        } else if (sheetName === 'Ledgers' && sheet.rowCount <= 1) {
          sheet.columns = columns;
          sheet.addRows(STANDARD_LEDGERS);
          modified = true;
        }
      }
    }
  }

  if (!fileExisted || modified) {
    await workbook.xlsx.writeFile(EXCEL_FILE);
  }
}

// Aliases to handle backwards compatibility with earlier sheet formats
const HEADER_ALIASES = {
  'hsn code': 'hsn',
  'hsn/sac code': 'hsn',
  'item type': 'type',
  'type': 'type',
  'opening stock': 'openingStock',
  'party type': 'partyType',
  'opening balance': 'openingBalance',
  'balance type': 'balanceType',
  'voucher type': 'type',
  'payment mode': 'mode',
  'debit ledger': 'debitLedger',
  'credit ledger': 'creditLedger',
  'reference no': 'referenceNo',
  'supplier bill no': 'supplierBillNo',
  'supplier id': 'supplierId',
  'supplier name': 'supplierName',
  'supplier gstin': 'supplierGstin',
  'group name': 'groupName',
  'parent group': 'parentGroup',
  'is default': 'isDefault'
};

async function readSheet(sheetName) {
  ensureDataDir();
  const workbook = new ExcelJS.Workbook();
  if (!fs.existsSync(EXCEL_FILE)) return [];
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) return [];

  const cols = SHEET_COLUMNS[sheetName] || [];
  const headerToKey = {};
  cols.forEach(col => {
    headerToKey[col.header] = col.key;
  });

  const colKeyMap = {};
  const firstRow = sheet.getRow(1);
  firstRow.eachCell((cell, colNumber) => {
    const val = cell.value ? String(cell.value).trim() : '';
    const lower = val.toLowerCase();
    colKeyMap[colNumber] = headerToKey[val] || HEADER_ALIASES[lower] || lower.replace(/[^a-z0-9]/g, '');
  });

  const rows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const obj = {};
      let hasData = false;
      firstRow.eachCell((cell, colNumber) => {
        const key = colKeyMap[colNumber];
        if (key) {
          let cellVal = row.getCell(colNumber).value;
          if (cellVal !== null && cellVal !== undefined) {
            hasData = true;
            if (typeof cellVal === 'object' && cellVal.text) {
              cellVal = cellVal.text;
            }
          }
          obj[key] = cellVal;
        }
      });
      if (hasData && (obj.id || obj.key || obj.name || obj.invoiceNumber || obj.purchaseNumber || obj.voucherNumber)) {
        rows.push(obj);
      }
    }
  });
  return rows;
}

async function writeSheet(sheetName, data, columns) {
  ensureDataDir();
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }
  let sheet = workbook.getWorksheet(sheetName);
  const cols = columns || SHEET_COLUMNS[sheetName];
  if (!sheet) {
    sheet = workbook.addWorksheet(sheetName);
  } else {
    sheet.spliceRows(2, Math.max(1, sheet.rowCount));
  }
  sheet.columns = cols;
  sheet.addRows(data);
  await workbook.xlsx.writeFile(EXCEL_FILE);
}

async function appendRow(sheetName, rowData, columns) {
  ensureDataDir();
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }
  let sheet = workbook.getWorksheet(sheetName);
  const cols = columns || SHEET_COLUMNS[sheetName];
  if (!sheet) {
    sheet = workbook.addWorksheet(sheetName);
  }
  sheet.columns = cols;
  const newRow = sheet.addRow(rowData);
  await workbook.xlsx.writeFile(EXCEL_FILE);
  return newRow;
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1200,
    minHeight: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    titleBarStyle: 'default',
    show: false,
  });

  win.loadFile('index.html');

  win.webContents.on('console-message', (event, level, message, line, sourceId) => {
    console.log(`[Renderer Console Level ${level}] ${message} (${sourceId}:${line})`);
  });

  if (process.argv.includes('--dev')) {
    win.webContents.openDevTools();
  }

  win.once('ready-to-show', async () => {
    if (process.argv.includes('--test')) {
      console.log('[TEST] Running TaxTech ERP automated integration tests...');
      try {
        await new Promise(r => setTimeout(r, 2000));
        const testResults = await win.webContents.executeJavaScript(`
          (async () => {
            const checks = [];

            // 1. App brand
            const brandH1 = document.querySelector('.sidebar-header h1')?.textContent;
            checks.push({ name: 'Brand Title is TaxTech ERP', passed: brandH1 === 'TaxTech ERP', actual: brandH1 });

            // 2. Theme Switcher
            const themeBtn = document.getElementById('themeToggleBtn');
            checks.push({ name: 'Theme Toggle Button exists', passed: !!themeBtn });
            if (themeBtn) {
              themeBtn.click();
              checks.push({ name: 'Dark Mode activation', passed: document.body.dataset.theme === 'dark' });
              themeBtn.click();
              checks.push({ name: 'Light Mode reactivation', passed: !document.body.dataset.theme });
            }

            // 3. Sales Invoice Searchable Combobox
            if (typeof openInvoiceModal === 'function') {
              await openInvoiceModal();
              const combobox = document.querySelector('#invoiceItemsTable .item-combobox');
              checks.push({ name: 'Invoice Searchable Combobox exists', passed: !!combobox });
              const searchInput = combobox?.querySelector('.item-search-input');
              checks.push({ name: 'Invoice Search Input exists', passed: !!searchInput });

              if (searchInput) {
                searchInput.focus();
                searchInput.value = 'Laptop';
                searchInput.dispatchEvent(new Event('input', { bubbles: true }));
                const hasEmptyMsg = combobox.querySelector('.combobox-empty')?.textContent.includes('No matching items found');
                checks.push({ name: 'Search input filtering and empty fallback works', passed: !!hasEmptyMsg });
              }
              closeInvoiceModal();
            }

            // 4. Purchase Bill Searchable Combobox
            if (typeof openPurchaseModal === 'function') {
              await openPurchaseModal();
              const pCombobox = document.querySelector('#purchaseItemsTable .pitem-combobox');
              checks.push({ name: 'Purchase Searchable Combobox exists', passed: !!pCombobox });
              const pSearchInput = pCombobox?.querySelector('.pitem-search-input');
              checks.push({ name: 'Purchase Search Input exists', passed: !!pSearchInput });
              closePurchaseModal();
            }

            // 5. Zero Tally references in UI text
            const allText = document.body.innerText;
            const hasTally = /tallyprime|tally erp/i.test(allText);
            checks.push({ name: 'Zero Tally references in UI text', passed: !hasTally });

            // 6. Invoice Format & Layout Settings Form
            const formatForm = document.getElementById('invoiceFormatSettingsForm');
            checks.push({ name: 'Invoice Format Settings Form exists', passed: !!formatForm });

            // 7. Logo Upload & Scale Controls
            const chooseLogoBtn = document.getElementById('chooseLogoBtn');
            const logoWidthSlider = document.getElementById('invoiceLogoWidth');
            checks.push({ name: 'Logo Upload Button and Size Slider exist', passed: !!chooseLogoBtn && !!logoWidthSlider });

            // 8. Logo & Header Alignment Controls
            const logoAlignBtns = document.querySelectorAll('#logoAlignGroup .pill-seg-btn');
            const headerAlignBtns = document.querySelectorAll('#headerAlignGroup .pill-seg-btn');
            checks.push({ name: 'Logo Alignment segmented buttons exist (3)', passed: logoAlignBtns.length === 3 });
            checks.push({ name: 'Header Alignment segmented buttons exist (3)', passed: headerAlignBtns.length === 3 });

            // 9. Theme Palette Swatches & Custom Color Picker
            const colorSwatches = document.querySelectorAll('#colorSwatches .color-swatch');
            const colorPicker = document.getElementById('invoiceThemeColor');
            checks.push({ name: 'Theme Color Swatches exist (6)', passed: colorSwatches.length === 6 });
            checks.push({ name: 'Custom Theme Color Picker exists', passed: !!colorPicker });

            // 10. Live Invoice Preview Modal
            if (typeof openInvoicePreviewModal === 'function' && typeof getSampleInvoiceData === 'function') {
              const { sampleInvoice, sampleItems, sampleClient } = getSampleInvoiceData({ invoiceTitle: 'TEST TAX INVOICE' });
              openInvoicePreviewModal(sampleInvoice, sampleItems, sampleClient, { invoiceThemeColor: '#059669', invoiceTitle: 'TEST TAX INVOICE' });
              const previewModal = document.getElementById('invoicePreviewModal');
              const paperContent = document.getElementById('invoicePaperPreview')?.innerHTML;
              const modalActive = previewModal?.classList.contains('active');
              const hasTestTitle = paperContent?.includes('TEST TAX INVOICE');
              checks.push({ name: 'Live Invoice Preview Modal opens with rendered paper content', passed: !!modalActive && !!hasTestTitle });
              closeInvoicePreviewModal();
              checks.push({ name: 'Live Invoice Preview Modal closes cleanly', passed: !previewModal?.classList.contains('active') });
            }

            // 11. Logo Upload & State Setting Functionality
            if (typeof setLogoImage === 'function') {
              const testLogoData = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
              setLogoImage(testLogoData);
              const logoImg = document.getElementById('logoPreviewImg');
              const compInput = document.getElementById('companyLogoData');
              const removeBtn = document.getElementById('removeLogoBtn');
              const logoSetPassed = logoImg?.style.display === 'block' && compInput?.value === testLogoData && removeBtn?.style.display !== 'none';
              checks.push({ name: 'Logo display & state update works on upload', passed: !!logoSetPassed });

              removeBtn?.click();
              const logoCleared = logoImg?.style.display === 'none' && !compInput?.value && removeBtn?.style.display === 'none';
              checks.push({ name: 'Logo removal resets state & UI cleanly', passed: !!logoCleared });
            }

            // 12. Universal Searchable Select / Dropdown Engine
            const allSelects = document.querySelectorAll('select');
            const allWraps = document.querySelectorAll('.searchable-select-wrap');
            checks.push({
              name: 'Searchable wrappers initialized for all dropdowns',
              passed: allWraps.length >= 10 && allWraps.length >= allSelects.length,
              actual: 'Wraps: ' + allWraps.length + ', Selects: ' + allSelects.length
            });

            // Test searching in State / Place of Supply dropdown
            const posSelect = document.getElementById('placeOfSupply');
            const posWrap = document.querySelector('.searchable-select-wrap[data-for-select="placeOfSupply"]');
            if (posSelect && posWrap) {
              const trigger = posWrap.querySelector('.searchable-select-trigger');
              trigger.click();
              const searchInput = posWrap.querySelector('.searchable-search-input');
              const isMenuOpen = posWrap.classList.contains('is-open');
              checks.push({ name: 'Dropdown menu opens on trigger click', passed: isMenuOpen });

              if (searchInput) {
                searchInput.value = 'Maha';
                searchInput.dispatchEvent(new Event('input', { bubbles: true }));
                const items = posWrap.querySelectorAll('.searchable-select-item');
                const hasMaharashtra = Array.from(items).some(item => item.textContent.includes('Maharashtra'));
                checks.push({ name: 'Search filter works in dropdown (typing "Maha")', passed: hasMaharashtra && items.length < 5 });

                const firstItem = items[0];
                if (firstItem) {
                  firstItem.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
                  checks.push({ name: 'Selecting filtered item updates native select value', passed: posSelect.value === 'Maharashtra' });
                }
              }

              // Verify dropdown closed after selection
              checks.push({ name: 'Dropdown closes automatically after option selection', passed: !posWrap.classList.contains('is-open') });
            }

            // 13. Excel Database & FY 2025-26 Sample Data Buttons
            const loadExcelBtn = document.getElementById('loadExcelDbBtn');
            const loadSampleBtn = document.getElementById('loadSampleFy2526Btn');
            const exportSampleBtn = document.getElementById('exportSampleFy2526Btn');
            checks.push({
              name: 'Excel Database Loading and Sample buttons exist in UI',
              passed: !!loadExcelBtn && !!loadSampleBtn && !!exportSampleBtn
            });

            // 14. API Exposure for Excel loading
            const hasLoadExcelApi = typeof window.api?.loadExcelDb === 'function';
            const hasLoadSampleApi = typeof window.api?.loadSampleFy2526 === 'function';
            const hasExportSampleApi = typeof window.api?.exportSampleFy2526 === 'function';
            checks.push({
              name: 'Excel database loading & sample APIs exposed in window.api',
              passed: hasLoadExcelApi && hasLoadSampleApi && hasExportSampleApi
            });

            return checks;
          })()
        `);
        console.log('[TEST RESULTS]:', JSON.stringify(testResults, null, 2));
        const failed = testResults.filter(c => !c.passed);
        if (failed.length > 0) {
          console.error('[TEST FAILED]', failed);
          app.exit(1);
        } else {
          console.log('[TEST PASSED] All TaxTech ERP features verified successfully!');
          app.exit(0);
        }
      } catch (err) {
        console.error('[TEST ERROR]', err);
        app.exit(1);
      }
      return;
    }
    win.show();
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(async () => {
  await initializeExcel();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ==========================================
// IPC HANDLERS: PRODUCTS & SERVICES
// ==========================================
ipcMain.handle('get-products', async () => readSheet('Products'));

ipcMain.handle('add-product', async (_, product) => {
  const newProduct = {
    ...product,
    id: uuidv4(),
    type: product.type || 'Goods',
    openingStock: product.type === 'Service' ? 0 : (Number(product.openingStock) || 0),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  await appendRow('Products', newProduct, SHEET_COLUMNS.Products);
  return newProduct;
});

ipcMain.handle('update-product', async (_, product) => {
  const products = await readSheet('Products');
  const index = products.findIndex(p => p.id === product.id);
  if (index !== -1) {
    products[index] = {
      ...products[index],
      ...product,
      type: product.type || products[index].type || 'Goods',
      openingStock: product.type === 'Service' ? 0 : (Number(product.openingStock) || 0),
      updatedAt: new Date().toISOString()
    };
    await writeSheet('Products', products, SHEET_COLUMNS.Products);
    return products[index];
  }
  return null;
});

ipcMain.handle('delete-product', async (_, id) => {
  const products = await readSheet('Products');
  const filtered = products.filter(p => p.id !== id);
  await writeSheet('Products', filtered, SHEET_COLUMNS.Products);
  return true;
});

// ==========================================
// IPC HANDLERS: CLIENTS & SUPPLIERS (PARTIES)
// ==========================================
ipcMain.handle('get-clients', async () => readSheet('Clients'));

ipcMain.handle('add-client', async (_, client) => {
  const newClient = {
    ...client,
    id: uuidv4(),
    partyType: client.partyType || 'Customer',
    openingBalance: Number(client.openingBalance) || 0,
    balanceType: client.balanceType || (client.partyType === 'Supplier' ? 'Cr' : 'Dr'),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  await appendRow('Clients', newClient, SHEET_COLUMNS.Clients);
  return newClient;
});

ipcMain.handle('update-client', async (_, client) => {
  const clients = await readSheet('Clients');
  const index = clients.findIndex(c => c.id === client.id);
  if (index !== -1) {
    clients[index] = {
      ...clients[index],
      ...client,
      partyType: client.partyType || clients[index].partyType || 'Customer',
      openingBalance: Number(client.openingBalance) || 0,
      balanceType: client.balanceType || clients[index].balanceType || 'Dr',
      updatedAt: new Date().toISOString()
    };
    await writeSheet('Clients', clients, SHEET_COLUMNS.Clients);
    return clients[index];
  }
  return null;
});

ipcMain.handle('delete-client', async (_, id) => {
  const clients = await readSheet('Clients');
  const filtered = clients.filter(c => c.id !== id);
  await writeSheet('Clients', filtered, SHEET_COLUMNS.Clients);
  return true;
});

// ==========================================
// IPC HANDLERS: SALES INVOICES
// ==========================================
ipcMain.handle('get-invoices', async () => readSheet('Invoices'));
ipcMain.handle('get-invoice-items', async () => readSheet('InvoiceItems'));

ipcMain.handle('save-invoice', async (_, { invoice, items }) => {
  const newInvoice = { ...invoice, id: uuidv4(), createdAt: new Date().toISOString() };
  await appendRow('Invoices', newInvoice, SHEET_COLUMNS.Invoices);

  for (const item of items) {
    const newItem = {
      ...item,
      id: uuidv4(),
      invoiceId: newInvoice.id,
      type: item.type || 'Goods'
    };
    await appendRow('InvoiceItems', newItem, SHEET_COLUMNS.InvoiceItems);
  }
  return newInvoice;
});

ipcMain.handle('update-invoice', async (_, { invoice, items }) => {
  const invoices = await readSheet('Invoices');
  const index = invoices.findIndex(i => i.id === invoice.id);
  if (index !== -1) {
    invoices[index] = invoice;
    await writeSheet('Invoices', invoices, SHEET_COLUMNS.Invoices);
  }

  // Delete old items and add new ones
  const allItems = await readSheet('InvoiceItems');
  const otherItems = allItems.filter(i => i.invoiceId !== invoice.id);
  for (const item of items) {
    const newItem = {
      ...item,
      id: uuidv4(),
      invoiceId: invoice.id,
      type: item.type || 'Goods'
    };
    otherItems.push(newItem);
  }
  await writeSheet('InvoiceItems', otherItems, SHEET_COLUMNS.InvoiceItems);
  return invoice;
});

ipcMain.handle('delete-invoice', async (_, id) => {
  const invoices = await readSheet('Invoices');
  const filtered = invoices.filter(i => i.id !== id);
  await writeSheet('Invoices', filtered, SHEET_COLUMNS.Invoices);

  const allItems = await readSheet('InvoiceItems');
  const filteredItems = allItems.filter(i => i.invoiceId !== id);
  await writeSheet('InvoiceItems', filteredItems, SHEET_COLUMNS.InvoiceItems);
  return true;
});

ipcMain.handle('get-next-invoice-number', async (_, settings) => {
  const prefix = settings.invoicePrefix || 'INV';
  const series = settings.invoiceSeries || '2024-25';
  const startNum = parseInt(settings.invoiceStartNumber) || 1;

  const invoices = await readSheet('Invoices');
  const seriesInvoices = invoices.filter(inv => inv.invoiceNumber && inv.invoiceNumber.startsWith(`${prefix}/${series}/`));
  let maxNum = startNum - 1;
  seriesInvoices.forEach(inv => {
    const parts = inv.invoiceNumber.split('/');
    const num = parseInt(parts[parts.length - 1]);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  });
  return `${prefix}/${series}/${String(maxNum + 1).padStart(4, '0')}`;
});

// ==========================================
// IPC HANDLERS: PURCHASES
// ==========================================
ipcMain.handle('get-purchases', async () => readSheet('Purchases'));
ipcMain.handle('get-purchase-items', async () => readSheet('PurchaseItems'));

ipcMain.handle('save-purchase', async (_, { purchase, items }) => {
  const newPurchase = { ...purchase, id: uuidv4(), createdAt: new Date().toISOString() };
  await appendRow('Purchases', newPurchase, SHEET_COLUMNS.Purchases);

  for (const item of items) {
    const newItem = {
      ...item,
      id: uuidv4(),
      purchaseId: newPurchase.id,
      type: item.type || 'Goods'
    };
    await appendRow('PurchaseItems', newItem, SHEET_COLUMNS.PurchaseItems);
  }
  return newPurchase;
});

ipcMain.handle('update-purchase', async (_, { purchase, items }) => {
  const purchases = await readSheet('Purchases');
  const index = purchases.findIndex(p => p.id === purchase.id);
  if (index !== -1) {
    purchases[index] = purchase;
    await writeSheet('Purchases', purchases, SHEET_COLUMNS.Purchases);
  }

  const allItems = await readSheet('PurchaseItems');
  const otherItems = allItems.filter(i => i.purchaseId !== purchase.id);
  for (const item of items) {
    const newItem = {
      ...item,
      id: uuidv4(),
      purchaseId: purchase.id,
      type: item.type || 'Goods'
    };
    otherItems.push(newItem);
  }
  await writeSheet('PurchaseItems', otherItems, SHEET_COLUMNS.PurchaseItems);
  return purchase;
});

ipcMain.handle('delete-purchase', async (_, id) => {
  const purchases = await readSheet('Purchases');
  const filtered = purchases.filter(p => p.id !== id);
  await writeSheet('Purchases', filtered, SHEET_COLUMNS.Purchases);

  const allItems = await readSheet('PurchaseItems');
  const filteredItems = allItems.filter(i => i.purchaseId !== id);
  await writeSheet('PurchaseItems', filteredItems, SHEET_COLUMNS.PurchaseItems);
  return true;
});

ipcMain.handle('get-next-purchase-number', async (_, settings = {}) => {
  const prefix = 'PUR';
  const series = settings.invoiceSeries || '2024-25';
  const purchases = await readSheet('Purchases');
  const seriesPurchases = purchases.filter(p => p.purchaseNumber && p.purchaseNumber.startsWith(`${prefix}/${series}/`));
  let maxNum = 0;
  seriesPurchases.forEach(p => {
    const parts = p.purchaseNumber.split('/');
    const num = parseInt(parts[parts.length - 1]);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  });
  return `${prefix}/${series}/${String(maxNum + 1).padStart(4, '0')}`;
});

// ==========================================
// IPC HANDLERS: LEDGER GROUPS (STANDARD CORE)
// ==========================================
ipcMain.handle('get-ledger-groups', async () => {
  const groups = await readSheet('LedgerGroups');
  if (!groups || groups.length === 0) {
    return STANDARD_GROUPS;
  }
  return groups;
});

ipcMain.handle('add-ledger-group', async (_, group) => {
  const newGroup = {
    ...group,
    id: uuidv4(),
    isDefault: 'No',
    createdAt: new Date().toISOString()
  };
  await appendRow('LedgerGroups', newGroup, SHEET_COLUMNS.LedgerGroups);
  return newGroup;
});

ipcMain.handle('update-ledger-group', async (_, group) => {
  const groups = await readSheet('LedgerGroups');
  const index = groups.findIndex(g => g.id === group.id);
  if (index !== -1) {
    groups[index] = { ...groups[index], ...group };
    await writeSheet('LedgerGroups', groups, SHEET_COLUMNS.LedgerGroups);
    return groups[index];
  }
  return null;
});

ipcMain.handle('delete-ledger-group', async (_, id) => {
  const groups = await readSheet('LedgerGroups');
  const target = groups.find(g => g.id === id);
  if (target && target.isDefault === 'Yes') {
    throw new Error('Standard core accounting groups cannot be deleted.');
  }
  const filtered = groups.filter(g => g.id !== id);
  await writeSheet('LedgerGroups', filtered, SHEET_COLUMNS.LedgerGroups);
  return true;
});

// ==========================================
// IPC HANDLERS: LEDGERS
// ==========================================
ipcMain.handle('get-ledgers', async () => {
  const ledgers = await readSheet('Ledgers');
  if (!ledgers || ledgers.length === 0) {
    return STANDARD_LEDGERS;
  }
  return ledgers;
});

ipcMain.handle('add-ledger', async (_, ledger) => {
  const newLedger = {
    ...ledger,
    id: uuidv4(),
    openingBalance: Number(ledger.openingBalance) || 0,
    balanceType: ledger.balanceType || 'Dr',
    isDefault: 'No',
    createdAt: new Date().toISOString()
  };
  await appendRow('Ledgers', newLedger, SHEET_COLUMNS.Ledgers);
  return newLedger;
});

ipcMain.handle('update-ledger', async (_, ledger) => {
  const ledgers = await readSheet('Ledgers');
  const index = ledgers.findIndex(l => l.id === ledger.id);
  if (index !== -1) {
    ledgers[index] = {
      ...ledgers[index],
      ...ledger,
      openingBalance: Number(ledger.openingBalance) || 0,
      balanceType: ledger.balanceType || 'Dr'
    };
    await writeSheet('Ledgers', ledgers, SHEET_COLUMNS.Ledgers);
    return ledgers[index];
  }
  return null;
});

ipcMain.handle('delete-ledger', async (_, id) => {
  const ledgers = await readSheet('Ledgers');
  const target = ledgers.find(l => l.id === id);
  if (target && target.isDefault === 'Yes') {
    throw new Error('Default system ledgers cannot be deleted.');
  }
  const filtered = ledgers.filter(l => l.id !== id);
  await writeSheet('Ledgers', filtered, SHEET_COLUMNS.Ledgers);
  return true;
});

// ==========================================
// IPC HANDLERS: VOUCHERS (RECEIPT, PAYMENT, CONTRA, JOURNAL)
// ==========================================
ipcMain.handle('get-vouchers', async () => readSheet('Vouchers'));

ipcMain.handle('save-voucher', async (_, voucher) => {
  let newVoucher;
  if (voucher.id) {
    const vouchers = await readSheet('Vouchers');
    const index = vouchers.findIndex(v => v.id === voucher.id);
    if (index !== -1) {
      vouchers[index] = { ...voucher, amount: Number(voucher.amount) || 0 };
      await writeSheet('Vouchers', vouchers, SHEET_COLUMNS.Vouchers);
      return vouchers[index];
    }
  }
  newVoucher = {
    ...voucher,
    id: uuidv4(),
    amount: Number(voucher.amount) || 0,
    createdAt: new Date().toISOString()
  };
  await appendRow('Vouchers', newVoucher, SHEET_COLUMNS.Vouchers);
  return newVoucher;
});

ipcMain.handle('delete-voucher', async (_, id) => {
  const vouchers = await readSheet('Vouchers');
  const filtered = vouchers.filter(v => v.id !== id);
  await writeSheet('Vouchers', filtered, SHEET_COLUMNS.Vouchers);
  return true;
});

ipcMain.handle('get-next-voucher-number', async (_, type = 'Payment') => {
  const prefixMap = {
    Payment: 'PAY',
    Receipt: 'RCP',
    Contra: 'CNT',
    Journal: 'JRN'
  };
  const prefix = prefixMap[type] || 'VCH';
  const vouchers = await readSheet('Vouchers');
  const matching = vouchers.filter(v => v.voucherNumber && v.voucherNumber.startsWith(`${prefix}/`));
  let maxNum = 0;
  matching.forEach(v => {
    const parts = v.voucherNumber.split('/');
    const num = parseInt(parts[parts.length - 1]);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  });
  return `${prefix}/${String(maxNum + 1).padStart(4, '0')}`;
});

// ==========================================
// IPC HANDLERS: STOCK REGISTER CALCULATION
// ==========================================
ipcMain.handle('get-stock-register', async () => {
  return calculateStockRegister();
});

async function calculateStockRegister() {
  const [products, invoiceItems, purchaseItems] = await Promise.all([
    readSheet('Products'),
    readSheet('InvoiceItems'),
    readSheet('PurchaseItems')
  ]);

  return products.map(product => {
    const isService = (product.type || '').toLowerCase().includes('service');
    if (isService) {
      return {
        id: product.id,
        name: product.name,
        type: 'Service',
        hsn: product.hsn || '',
        unit: product.unit || 'OTH',
        openingStock: 0,
        inwardQty: 0,
        outwardQty: 0,
        closingStock: 0,
        purchaseRate: Number(product.purchasePrice) || 0,
        saleRate: Number(product.salePrice) || 0,
        closingValue: 0,
        status: 'Service (No Stock)'
      };
    }

    const openingStock = Number(product.openingStock) || 0;
    
    // Inward from Purchases
    const inwardQty = purchaseItems
      .filter(item => item.productId === product.id || item.productName === product.name)
      .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);

    // Outward from Sales Invoices
    const outwardQty = invoiceItems
      .filter(item => item.productId === product.id || item.productName === product.name)
      .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);

    const closingStock = openingStock + inwardQty - outwardQty;
    const purchaseRate = Number(product.purchasePrice) || Number(product.salePrice) || 0;
    const closingValue = Math.max(0, closingStock) * purchaseRate;

    let status = 'In Stock';
    if (closingStock <= 0) {
      status = 'Out of Stock';
    } else if (closingStock < 5) {
      status = 'Low Stock';
    }

    return {
      id: product.id,
      name: product.name,
      type: 'Goods',
      hsn: product.hsn || '',
      unit: product.unit || 'PCS',
      openingStock,
      inwardQty,
      outwardQty,
      closingStock,
      purchaseRate,
      saleRate: Number(product.salePrice) || 0,
      closingValue,
      status
    };
  });
}

// ==========================================
// IPC HANDLERS: FINANCIAL REPORTS (DOUBLE-ENTRY SUITE)
// Trial Balance, Trading Account, P&L, Balance Sheet, Cash Book, Bank Book
// ==========================================
ipcMain.handle('get-financial-reports', async () => {
  return calculateFinancialReports();
});

async function calculateFinancialReports() {
  const [
    products,
    clients,
    invoices,
    purchases,
    vouchers,
    groups,
    customLedgers
  ] = await Promise.all([
    readSheet('Products'),
    readSheet('Clients'),
    readSheet('Invoices'),
    readSheet('Purchases'),
    readSheet('Vouchers'),
    readSheet('LedgerGroups').then(g => (g && g.length > 0 ? g : STANDARD_GROUPS)),
    readSheet('Ledgers').then(l => (l && l.length > 0 ? l : STANDARD_LEDGERS))
  ]);

  const stockRegister = await calculateStockRegister();
  const closingStockValuation = stockRegister
    .filter(i => i.type === 'Goods')
    .reduce((sum, i) => sum + (i.closingValue || 0), 0);

  const openingStockValuation = products
    .filter(p => (p.type || 'Goods') === 'Goods')
    .reduce((sum, p) => sum + ((Number(p.openingStock) || 0) * (Number(p.purchasePrice) || Number(p.salePrice) || 0)), 0);

  // Group Map for fast lookup
  const groupMap = {};
  groups.forEach(g => {
    groupMap[g.name] = g;
  });

  // Ledger Balances dictionary: ledgerName -> { debit: number, credit: number, groupName, nature }
  const ledgerAccounts = {};

  function ensureLedger(name, defaultGroup = 'Suspense Account', defaultNature = 'Liabilities') {
    if (!ledgerAccounts[name]) {
      const foundMaster = customLedgers.find(l => l.name === name);
      const grp = foundMaster ? foundMaster.groupName : defaultGroup;
      const nature = groupMap[grp] ? groupMap[grp].nature : defaultNature;
      ledgerAccounts[name] = {
        name,
        groupName: grp,
        nature,
        debit: 0,
        credit: 0
      };
    }
    return ledgerAccounts[name];
  }

  // 1. Initialize master ledgers with opening balances
  customLedgers.forEach(l => {
    const acc = ensureLedger(l.name, l.groupName);
    const ob = Number(l.openingBalance) || 0;
    if (l.balanceType === 'Dr') {
      acc.debit += ob;
    } else {
      acc.credit += ob;
    }
  });

  // 2. Add Party (Customer/Supplier) Opening Balances
  clients.forEach(c => {
    const isSupplier = c.partyType === 'Supplier';
    const defaultGrp = isSupplier ? 'Sundry Creditors' : 'Sundry Debtors';
    const acc = ensureLedger(c.name, defaultGrp, isSupplier ? 'Liabilities' : 'Assets');
    const ob = Number(c.openingBalance) || 0;
    if (c.balanceType === 'Cr') {
      acc.credit += ob;
    } else if (c.balanceType === 'Dr' || ob > 0) {
      acc.debit += ob;
    }
  });

  // 3. Process Sales Invoices
  // Client Dr (Grand Total), Sales Account Cr (Subtotal), Duties & Taxes Output Cr
  const salesAcc = ensureLedger('Sales Account', 'Sales Accounts', 'Income');
  const cgstOutAcc = ensureLedger('CGST Output', 'Duties & Taxes', 'Liabilities');
  const sgstOutAcc = ensureLedger('SGST Output', 'Duties & Taxes', 'Liabilities');
  const igstOutAcc = ensureLedger('IGST Output', 'Duties & Taxes', 'Liabilities');

  invoices.forEach(inv => {
    const clientName = inv.clientName || 'General Customer';
    const clientAcc = ensureLedger(clientName, 'Sundry Debtors', 'Assets');

    const grand = Number(inv.grandTotal) || 0;
    const subtotal = Number(inv.subtotal) || 0;
    const cgst = Number(inv.totalCgst) || 0;
    const sgst = Number(inv.totalSgst) || 0;
    const igst = Number(inv.totalIgst) || 0;

    clientAcc.debit += grand;
    salesAcc.credit += subtotal;
    cgstOutAcc.credit += cgst;
    sgstOutAcc.credit += sgst;
    igstOutAcc.credit += igst;
  });

  // 4. Process Purchases
  // Purchase Account Dr (Subtotal), Duties & Taxes Input Dr, Supplier Cr (Grand Total)
  const purchaseAcc = ensureLedger('Purchase Account', 'Purchase Accounts', 'Expenses');
  const cgstInAcc = ensureLedger('CGST Input', 'Duties & Taxes', 'Liabilities');
  const sgstInAcc = ensureLedger('SGST Input', 'Duties & Taxes', 'Liabilities');
  const igstInAcc = ensureLedger('IGST Input', 'Duties & Taxes', 'Liabilities');

  purchases.forEach(pur => {
    const supplierName = pur.supplierName || 'General Supplier';
    const supplierAcc = ensureLedger(supplierName, 'Sundry Creditors', 'Liabilities');

    const grand = Number(pur.grandTotal) || 0;
    const subtotal = Number(pur.subtotal) || 0;
    const cgst = Number(pur.totalCgst) || 0;
    const sgst = Number(pur.totalSgst) || 0;
    const igst = Number(pur.totalIgst) || 0;

    supplierAcc.credit += grand;
    purchaseAcc.debit += subtotal;
    cgstInAcc.debit += cgst;
    sgstInAcc.debit += sgst;
    igstInAcc.debit += igst;
  });

  // 5. Process Vouchers (Receipt, Payment, Contra, Journal)
  vouchers.forEach(v => {
    const amount = Number(v.amount) || 0;
    if (amount <= 0) return;

    if (v.type === 'Receipt') {
      const modeLedger = v.mode === 'Cash' ? 'Cash' : 'Bank Account';
      const dr = ensureLedger(modeLedger, v.mode === 'Cash' ? 'Cash-in-hand' : 'Bank Accounts', 'Assets');
      const cr = ensureLedger(v.creditLedger || 'Suspense Account', 'Sundry Debtors', 'Assets');
      dr.debit += amount;
      cr.credit += amount;
    } else if (v.type === 'Payment') {
      const modeLedger = v.mode === 'Cash' ? 'Cash' : 'Bank Account';
      const cr = ensureLedger(modeLedger, v.mode === 'Cash' ? 'Cash-in-hand' : 'Bank Accounts', 'Assets');
      const dr = ensureLedger(v.debitLedger || 'Suspense Account', 'Sundry Creditors', 'Liabilities');
      dr.debit += amount;
      cr.credit += amount;
    } else if (v.type === 'Contra') {
      const dr = ensureLedger(v.debitLedger || 'Bank Account', 'Bank Accounts', 'Assets');
      const cr = ensureLedger(v.creditLedger || 'Cash', 'Cash-in-hand', 'Assets');
      dr.debit += amount;
      cr.credit += amount;
    } else if (v.type === 'Journal') {
      const dr = ensureLedger(v.debitLedger || 'Suspense Account', 'Suspense Account', 'Liabilities');
      const cr = ensureLedger(v.creditLedger || 'Suspense Account', 'Suspense Account', 'Liabilities');
      dr.debit += amount;
      cr.credit += amount;
    }
  });

  // ==========================================
  // BUILD TRIAL BALANCE
  // ==========================================
  const trialBalanceRows = [];
  let tbTotalDebit = 0;
  let tbTotalCredit = 0;

  for (const acc of Object.values(ledgerAccounts)) {
    let closingDebit = 0;
    let closingCredit = 0;

    if (acc.debit >= acc.credit) {
      closingDebit = acc.debit - acc.credit;
    } else {
      closingCredit = acc.credit - acc.debit;
    }

    if (closingDebit > 0 || closingCredit > 0) {
      trialBalanceRows.push({
        ledgerName: acc.name,
        groupName: acc.groupName,
        nature: acc.nature,
        debit: closingDebit,
        credit: closingCredit
      });
      tbTotalDebit += closingDebit;
      tbTotalCredit += closingCredit;
    }
  }

  trialBalanceRows.sort((a, b) => a.ledgerName.localeCompare(b.ledgerName));

  // ==========================================
  // BUILD TRADING ACCOUNT
  // ==========================================
  const totalSales = ledgerAccounts['Sales Account'] ? (ledgerAccounts['Sales Account'].credit - ledgerAccounts['Sales Account'].debit) : 0;
  const totalPurchases = ledgerAccounts['Purchase Account'] ? (ledgerAccounts['Purchase Account'].debit - ledgerAccounts['Purchase Account'].credit) : 0;

  // Direct Expenses and Incomes
  let totalDirectExpenses = 0;
  const directExpenseRows = [];
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Direct Expenses') {
      const bal = acc.debit - acc.credit;
      if (bal > 0) {
        directExpenseRows.push({ name: acc.name, amount: bal });
        totalDirectExpenses += bal;
      }
    }
  }

  let totalDirectIncomes = 0;
  const directIncomeRows = [];
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Direct Incomes') {
      const bal = acc.credit - acc.debit;
      if (bal > 0) {
        directIncomeRows.push({ name: acc.name, amount: bal });
        totalDirectIncomes += bal;
      }
    }
  }

  const tradingDebitSide = openingStockValuation + totalPurchases + totalDirectExpenses;
  const tradingCreditSide = totalSales + totalDirectIncomes + closingStockValuation;
  const grossProfit = tradingCreditSide - tradingDebitSide;

  const tradingAccount = {
    openingStock: openingStockValuation,
    purchases: totalPurchases,
    directExpenses: directExpenseRows,
    totalDirectExpenses,
    sales: totalSales,
    directIncomes: directIncomeRows,
    totalDirectIncomes,
    closingStock: closingStockValuation,
    grossProfit: grossProfit >= 0 ? grossProfit : 0,
    grossLoss: grossProfit < 0 ? Math.abs(grossProfit) : 0,
    totalDebit: Math.max(tradingDebitSide, tradingCreditSide),
    totalCredit: Math.max(tradingDebitSide, tradingCreditSide)
  };

  // ==========================================
  // BUILD PROFIT AND LOSS ACCOUNT
  // ==========================================
  let totalIndirectExpenses = 0;
  const indirectExpenseRows = [];
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Indirect Expenses') {
      const bal = acc.debit - acc.credit;
      if (bal > 0) {
        indirectExpenseRows.push({ name: acc.name, amount: bal });
        totalIndirectExpenses += bal;
      }
    }
  }

  let totalIndirectIncomes = 0;
  const indirectIncomeRows = [];
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Indirect Incomes') {
      const bal = acc.credit - acc.debit;
      if (bal > 0) {
        indirectIncomeRows.push({ name: acc.name, amount: bal });
        totalIndirectIncomes += bal;
      }
    }
  }

  const plIncomeSide = (grossProfit >= 0 ? grossProfit : 0) + totalIndirectIncomes;
  const plExpenseSide = (grossProfit < 0 ? Math.abs(grossProfit) : 0) + totalIndirectExpenses;
  const netProfit = plIncomeSide - plExpenseSide;

  const profitLossAccount = {
    grossProfit: grossProfit >= 0 ? grossProfit : 0,
    grossLoss: grossProfit < 0 ? Math.abs(grossProfit) : 0,
    indirectExpenses: indirectExpenseRows,
    totalIndirectExpenses,
    indirectIncomes: indirectIncomeRows,
    totalIndirectIncomes,
    netProfit: netProfit >= 0 ? netProfit : 0,
    netLoss: netProfit < 0 ? Math.abs(netProfit) : 0,
    totalDebit: Math.max(plIncomeSide, plExpenseSide),
    totalCredit: Math.max(plIncomeSide, plExpenseSide)
  };

  // ==========================================
  // BUILD BALANCE SHEET
  // ==========================================
  const liabilitiesList = [];
  const assetsList = [];

  // 1. Capital Account + Net Profit / - Net Loss
  const capitalAcc = ledgerAccounts['Capital Account'];
  const baseCapital = capitalAcc ? (capitalAcc.credit - capitalAcc.debit) : 0;
  const effectiveCapital = baseCapital + netProfit;
  liabilitiesList.push({
    group: 'Capital Account',
    particulars: 'Capital Account (incl. Net Profit/Loss)',
    amount: effectiveCapital
  });

  // 2. Loans (Liability) & Subgroups
  let totalLoans = 0;
  for (const acc of Object.values(ledgerAccounts)) {
    if (['Loans (Liability)', 'Secured Loans', 'Unsecured Loans', 'Bank OD A/c'].includes(acc.groupName)) {
      const bal = acc.credit - acc.debit;
      if (bal > 0) {
        liabilitiesList.push({ group: acc.groupName, particulars: acc.name, amount: bal });
        totalLoans += bal;
      }
    }
  }

  // 3. Current Liabilities - Sundry Creditors
  let totalCreditors = 0;
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Sundry Creditors') {
      const bal = acc.credit - acc.debit;
      if (bal > 0) {
        totalCreditors += bal;
      }
    }
  }
  if (totalCreditors > 0) {
    liabilitiesList.push({ group: 'Current Liabilities', particulars: 'Sundry Creditors (Suppliers Payable)', amount: totalCreditors });
  }

  // 4. Duties & Taxes (Net GST output vs input)
  const netCgst = (ledgerAccounts['CGST Output']?.credit || 0) - (ledgerAccounts['CGST Input']?.debit || 0);
  const netSgst = (ledgerAccounts['SGST Output']?.credit || 0) - (ledgerAccounts['SGST Input']?.debit || 0);
  const netIgst = (ledgerAccounts['IGST Output']?.credit || 0) - (ledgerAccounts['IGST Input']?.debit || 0);
  const netGstPayable = netCgst + netSgst + netIgst;

  if (netGstPayable > 0) {
    liabilitiesList.push({ group: 'Current Liabilities', particulars: 'Duties & Taxes (Net GST Payable)', amount: netGstPayable });
  }

  // ASSETS SIDE
  // 1. Fixed Assets
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Fixed Assets') {
      const bal = acc.debit - acc.credit;
      if (bal > 0) {
        assetsList.push({ group: 'Fixed Assets', particulars: acc.name, amount: bal });
      }
    }
  }

  // 2. Current Assets: Closing Stock
  assetsList.push({ group: 'Current Assets', particulars: 'Closing Stock (from Stock Register)', amount: closingStockValuation });

  // 3. Sundry Debtors
  let totalDebtors = 0;
  for (const acc of Object.values(ledgerAccounts)) {
    if (acc.groupName === 'Sundry Debtors') {
      const bal = acc.debit - acc.credit;
      if (bal > 0) {
        totalDebtors += bal;
      }
    }
  }
  if (totalDebtors > 0) {
    assetsList.push({ group: 'Current Assets', particulars: 'Sundry Debtors (Receivables)', amount: totalDebtors });
  }

  // 4. Bank Accounts
  const bankBal = ledgerAccounts['Bank Account'] ? (ledgerAccounts['Bank Account'].debit - ledgerAccounts['Bank Account'].credit) : 0;
  assetsList.push({ group: 'Current Assets', particulars: 'Bank Accounts (Closing Balance)', amount: bankBal });

  // 5. Cash in Hand
  const cashBal = ledgerAccounts['Cash'] ? (ledgerAccounts['Cash'].debit - ledgerAccounts['Cash'].credit) : 0;
  assetsList.push({ group: 'Current Assets', particulars: 'Cash-in-hand (Closing Balance)', amount: cashBal });

  // 6. GST Input Tax Credit (if input > output)
  if (netGstPayable < 0) {
    assetsList.push({ group: 'Current Assets', particulars: 'GST Input Tax Credit (Net Credit)', amount: Math.abs(netGstPayable) });
  }

  // Other current assets / deposits / investments
  for (const acc of Object.values(ledgerAccounts)) {
    if (['Investments', 'Deposits (Asset)', 'Loans & Advances (Asset)'].includes(acc.groupName)) {
      const bal = acc.debit - acc.credit;
      if (bal > 0) {
        assetsList.push({ group: acc.groupName, particulars: acc.name, amount: bal });
      }
    }
  }

  const totalLiabilities = liabilitiesList.reduce((sum, item) => sum + item.amount, 0);
  const totalAssets = assetsList.reduce((sum, item) => sum + item.amount, 0);

  const balanceSheet = {
    liabilities: liabilitiesList,
    assets: assetsList,
    totalLiabilities,
    totalAssets,
    difference: Math.abs(totalLiabilities - totalAssets)
  };

  // ==========================================
  // BUILD CASH BOOK & BANK BOOK
  // ==========================================
  const cashMaster = customLedgers.find(l => l.name === 'Cash');
  const bankMaster = customLedgers.find(l => l.name === 'Bank Account');
  const initialCash = cashMaster ? (cashMaster.balanceType === 'Dr' ? Number(cashMaster.openingBalance) : -Number(cashMaster.openingBalance)) : 10000;
  const initialBank = bankMaster ? (bankMaster.balanceType === 'Dr' ? Number(bankMaster.openingBalance) : -Number(bankMaster.openingBalance)) : 50000;

  let runningCash = initialCash;
  let runningBank = initialBank;

  const cashTransactions = [];
  const bankTransactions = [];

  let totalCashReceipts = 0;
  let totalCashPayments = 0;
  let totalBankReceipts = 0;
  let totalBankPayments = 0;

  vouchers.sort((a, b) => new Date(a.date || 0) - new Date(b.date || 0)).forEach(v => {
    const amt = Number(v.amount) || 0;
    if (amt <= 0) return;

    if (v.type === 'Receipt') {
      if (v.mode === 'Cash') {
        runningCash += amt;
        totalCashReceipts += amt;
        cashTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Receipt',
          particulars: `Received from ${v.creditLedger || 'Party'}`,
          debit: amt,
          credit: 0,
          balance: runningCash,
          narration: v.narration || ''
        });
      } else {
        runningBank += amt;
        totalBankReceipts += amt;
        bankTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Receipt',
          particulars: `Received from ${v.creditLedger || 'Party'}`,
          debit: amt,
          credit: 0,
          balance: runningBank,
          narration: v.narration || ''
        });
      }
    } else if (v.type === 'Payment') {
      if (v.mode === 'Cash') {
        runningCash -= amt;
        totalCashPayments += amt;
        cashTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Payment',
          particulars: `Paid to ${v.debitLedger || 'Party'}`,
          debit: 0,
          credit: amt,
          balance: runningCash,
          narration: v.narration || ''
        });
      } else {
        runningBank -= amt;
        totalBankPayments += amt;
        bankTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Payment',
          particulars: `Paid to ${v.debitLedger || 'Party'}`,
          debit: 0,
          credit: amt,
          balance: runningBank,
          narration: v.narration || ''
        });
      }
    } else if (v.type === 'Contra') {
      // Contra: Cash to Bank or Bank to Cash
      if (v.debitLedger === 'Bank Account' && v.creditLedger === 'Cash') {
        // Cash deposited to Bank
        runningCash -= amt;
        totalCashPayments += amt;
        cashTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Contra',
          particulars: 'Deposited into Bank',
          debit: 0,
          credit: amt,
          balance: runningCash,
          narration: v.narration || ''
        });

        runningBank += amt;
        totalBankReceipts += amt;
        bankTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Contra',
          particulars: 'Cash Deposit received',
          debit: amt,
          credit: 0,
          balance: runningBank,
          narration: v.narration || ''
        });
      } else if (v.debitLedger === 'Cash' && v.creditLedger === 'Bank Account') {
        // Cash withdrawn from Bank
        runningCash += amt;
        totalCashReceipts += amt;
        cashTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Contra',
          particulars: 'Cash Withdrawn from Bank',
          debit: amt,
          credit: 0,
          balance: runningCash,
          narration: v.narration || ''
        });

        runningBank -= amt;
        totalBankPayments += amt;
        bankTransactions.push({
          date: v.date,
          voucherNumber: v.voucherNumber,
          type: 'Contra',
          particulars: 'Withdrawal to Cash',
          debit: 0,
          credit: amt,
          balance: runningBank,
          narration: v.narration || ''
        });
      }
    }
  });

  const cashBook = {
    openingBalance: initialCash,
    closingBalance: runningCash,
    totalReceipts: totalCashReceipts,
    totalPayments: totalCashPayments,
    transactions: cashTransactions
  };

  const bankBook = {
    openingBalance: initialBank,
    closingBalance: runningBank,
    totalReceipts: totalBankReceipts,
    totalPayments: totalBankPayments,
    transactions: bankTransactions
  };

  return {
    trialBalance: {
      rows: trialBalanceRows,
      totalDebit: tbTotalDebit,
      totalCredit: tbTotalCredit,
      difference: Math.abs(tbTotalDebit - tbTotalCredit)
    },
    tradingAccount,
    profitLossAccount,
    balanceSheet,
    cashBook,
    bankBook
  };
}

// ==========================================
// IPC HANDLERS: SETTINGS & UTILS
// ==========================================
ipcMain.handle('get-settings', async () => {
  ensureDataDir();
  const settings = await readSheet('CompanySettings');
  const obj = {};
  settings.forEach(s => {
    const k = s.key !== undefined ? s.key : s.Key;
    const v = s.value !== undefined ? s.value : s.Value;
    if (k) obj[k] = v;
  });

  // If companyLogo is stored as a file path, convert to Data URL so it renders reliably in the DOM
  const logo = obj.companyLogo || obj.invoiceLogo;
  if (logo && !logo.startsWith('data:image') && fs.existsSync(logo)) {
    try {
      const ext = path.extname(logo).toLowerCase().replace('.', '');
      const mimeType = ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg';
      const fileBuffer = fs.readFileSync(logo);
      obj.companyLogoDataUrl = `data:${mimeType};base64,${fileBuffer.toString('base64')}`;
    } catch (e) {
      console.error('Error generating logo Data URL in get-settings:', e);
    }
  }

  return obj;
});

ipcMain.handle('save-settings', async (_, settings) => {
  const rows = Object.entries(settings).map(([key, value]) => ({ key, value }));
  await writeSheet('CompanySettings', rows, SHEET_COLUMNS.CompanySettings);
  return true;
});

ipcMain.handle('select-logo-file', async () => {
  ensureDataDir();
  const focusedWin = BrowserWindow.getFocusedWindow();
  const options = {
    title: 'Select Company Logo Image',
    filters: [
      { name: 'Image Files', extensions: ['png', 'jpg', 'jpeg', 'webp'] }
    ],
    properties: ['openFile']
  };

  const result = focusedWin
    ? await dialog.showOpenDialog(focusedWin, options)
    : await dialog.showOpenDialog(options);

  if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
    return { success: false, cancelled: true };
  }

  try {
    const srcPath = result.filePaths[0];
    const ext = path.extname(srcPath).toLowerCase().replace('.', '');
    const mimeType = ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg';
    const fileBuffer = fs.readFileSync(srcPath);

    // Save a copy in DATA_DIR
    const savedLogoPath = path.join(DATA_DIR, `company_logo.${ext}`);
    fs.writeFileSync(savedLogoPath, fileBuffer);

    const base64Str = fileBuffer.toString('base64');
    const dataUrl = `data:${mimeType};base64,${base64Str}`;

    return {
      success: true,
      dataUrl,
      base64: dataUrl,
      filePath: savedLogoPath,
      savedPath: savedLogoPath,
      originalPath: srcPath
    };
  } catch (err) {
    console.error('Error reading selected logo file:', err);
    return { success: false, error: err.message };
  }
});

ipcMain.handle('generate-invoice-pdf', async (_, { invoice, items, settings, client }) => {
  const { filePath } = await dialog.showSaveDialog({
    defaultPath: `Invoice_${(invoice.invoiceNumber || 'INV').replace(/[\/\\]/g, '_')}.pdf`,
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
  });
  if (!filePath) return { success: false, cancelled: true };

  return new Promise((resolve) => {
    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    // Font configurations
    const fontChoice = settings.invoiceFont || 'Helvetica';
    const fontRegular = fontChoice === 'Times' ? 'Times-Roman' : fontChoice === 'Courier' ? 'Courier' : 'Helvetica';
    const fontBold = fontChoice === 'Times' ? 'Times-Bold' : fontChoice === 'Courier' ? 'Courier-Bold' : 'Helvetica-Bold';
    const fontItalic = fontChoice === 'Times' ? 'Times-Italic' : fontChoice === 'Courier' ? 'Courier-Oblique' : 'Helvetica-Oblique';

    // Theme Color & Alignment configurations
    const primaryColor = settings.invoiceThemeColor || '#1a1a2e';
    const logoAlign = settings.invoiceLogoAlign || 'left';
    const headerAlign = settings.invoiceHeaderAlign || 'left';
    const logoWidth = Math.min(160, Math.max(50, parseInt(settings.invoiceLogoWidth, 10) || 100));
    const showLogo = settings.showLogo !== 'false' && settings.showLogo !== false;

    // Decode logo if present
    let logoBuffer = null;
    const rawLogo = settings.companyLogo || settings.invoiceLogo;
    if (showLogo && rawLogo) {
      try {
        if (rawLogo.startsWith('data:image')) {
          const base64Data = rawLogo.replace(/^data:image\/\w+;base64,/, '');
          logoBuffer = Buffer.from(base64Data, 'base64');
        } else if (fs.existsSync(rawLogo)) {
          logoBuffer = fs.readFileSync(rawLogo);
        }
      } catch (e) {
        console.error('Error loading logo for PDF:', e);
      }
    }

    // --- 1. HEADER & LOGO ---
    const startY = 40;
    let headerBottomY = startY;

    if (logoBuffer) {
      if (logoAlign === 'left') {
        doc.image(logoBuffer, 40, startY, { width: logoWidth });
        const textX = 40 + logoWidth + 15;
        const textW = 515 - logoWidth - 15;
        doc.y = startY;
        doc.fontSize(18).font(fontBold).fillColor(primaryColor).text(settings.companyName || 'My Company', textX, doc.y, { width: textW, align: headerAlign === 'right' ? 'right' : 'left' });
        doc.fontSize(8.5).font(fontRegular).fillColor('#334155');
        doc.text(`${settings.companyAddress1 || ''}${settings.companyAddress2 ? ', ' + settings.companyAddress2 : ''}`, textX, doc.y + 2, { width: textW, align: headerAlign === 'right' ? 'right' : 'left' });
        doc.text(`${settings.companyCity || ''}${settings.companyPincode ? ' - ' + settings.companyPincode : ''}, ${settings.companyState || ''}${settings.companyStateCode ? ' (' + settings.companyStateCode + ')' : ''}`, textX, doc.y, { width: textW, align: headerAlign === 'right' ? 'right' : 'left' });
        doc.text(`Phone: ${settings.companyPhone || '-'} | Email: ${settings.companyEmail || '-'}`, textX, doc.y, { width: textW, align: headerAlign === 'right' ? 'right' : 'left' });
        doc.font(fontBold).text(`GSTIN: ${settings.companyGstin || '-'} | PAN: ${settings.companyPan || '-'}`, textX, doc.y, { width: textW, align: headerAlign === 'right' ? 'right' : 'left' });
        headerBottomY = Math.max(doc.y, startY + (logoWidth * 0.75));
      } else if (logoAlign === 'right') {
        const textW = 515 - logoWidth - 15;
        doc.image(logoBuffer, 555 - logoWidth, startY, { width: logoWidth });
        doc.y = startY;
        doc.fontSize(18).font(fontBold).fillColor(primaryColor).text(settings.companyName || 'My Company', 40, doc.y, { width: textW, align: 'left' });
        doc.fontSize(8.5).font(fontRegular).fillColor('#334155');
        doc.text(`${settings.companyAddress1 || ''}${settings.companyAddress2 ? ', ' + settings.companyAddress2 : ''}`, 40, doc.y + 2, { width: textW, align: 'left' });
        doc.text(`${settings.companyCity || ''}${settings.companyPincode ? ' - ' + settings.companyPincode : ''}, ${settings.companyState || ''}${settings.companyStateCode ? ' (' + settings.companyStateCode + ')' : ''}`, 40, doc.y, { width: textW, align: 'left' });
        doc.text(`Phone: ${settings.companyPhone || '-'} | Email: ${settings.companyEmail || '-'}`, 40, doc.y, { width: textW, align: 'left' });
        doc.font(fontBold).text(`GSTIN: ${settings.companyGstin || '-'} | PAN: ${settings.companyPan || '-'}`, 40, doc.y, { width: textW, align: 'left' });
        headerBottomY = Math.max(doc.y, startY + (logoWidth * 0.75));
      } else {
        // Center
        const logoX = 40 + (515 - logoWidth) / 2;
        doc.image(logoBuffer, logoX, startY, { width: logoWidth });
        doc.y = startY + (logoWidth * 0.6) + 6;
        doc.fontSize(18).font(fontBold).fillColor(primaryColor).text(settings.companyName || 'My Company', 40, doc.y, { width: 515, align: 'center' });
        doc.fontSize(8.5).font(fontRegular).fillColor('#334155');
        doc.text(`${settings.companyAddress1 || ''}${settings.companyAddress2 ? ', ' + settings.companyAddress2 : ''}`, 40, doc.y + 2, { width: 515, align: 'center' });
        doc.text(`${settings.companyCity || ''}${settings.companyPincode ? ' - ' + settings.companyPincode : ''}, ${settings.companyState || ''}${settings.companyStateCode ? ' (' + settings.companyStateCode + ')' : ''}`, 40, doc.y, { width: 515, align: 'center' });
        doc.text(`Phone: ${settings.companyPhone || '-'} | Email: ${settings.companyEmail || '-'}`, 40, doc.y, { width: 515, align: 'center' });
        doc.font(fontBold).text(`GSTIN: ${settings.companyGstin || '-'} | PAN: ${settings.companyPan || '-'}`, 40, doc.y, { width: 515, align: 'center' });
        headerBottomY = doc.y;
      }
    } else {
      // No Logo
      doc.y = startY;
      doc.fontSize(20).font(fontBold).fillColor(primaryColor).text(settings.companyName || 'My Company', 40, doc.y, { width: 515, align: headerAlign });
      doc.fontSize(8.5).font(fontRegular).fillColor('#334155');
      doc.text(`${settings.companyAddress1 || ''}${settings.companyAddress2 ? ', ' + settings.companyAddress2 : ''}`, 40, doc.y + 2, { width: 515, align: headerAlign });
      doc.text(`${settings.companyCity || ''}${settings.companyPincode ? ' - ' + settings.companyPincode : ''}, ${settings.companyState || ''}${settings.companyStateCode ? ' (' + settings.companyStateCode + ')' : ''}`, 40, doc.y, { width: 515, align: headerAlign });
      doc.text(`Phone: ${settings.companyPhone || '-'} | Email: ${settings.companyEmail || '-'}`, 40, doc.y, { width: 515, align: headerAlign });
      doc.font(fontBold).text(`GSTIN: ${settings.companyGstin || '-'} | PAN: ${settings.companyPan || '-'}`, 40, doc.y, { width: 515, align: headerAlign });
      headerBottomY = doc.y;
    }

    doc.y = headerBottomY + 8;

    // Decorative Accent Bar
    doc.strokeColor(primaryColor).lineWidth(1.8).moveTo(40, doc.y).lineTo(555, doc.y).stroke();
    doc.moveDown(0.4);

    // Document Title & Subtitle
    const invoiceTitle = settings.invoiceTitle || 'TAX INVOICE';
    doc.fontSize(15).font(fontBold).fillColor(primaryColor).text(invoiceTitle, { align: 'center' });
    if (settings.invoiceSubtitle) {
      doc.fontSize(8).font(fontItalic).fillColor('#64748b').text(settings.invoiceSubtitle, { align: 'center' });
    }
    doc.moveDown(0.4);

    // --- 2. INVOICE META & BILL TO / SHIP TO ---
    const metaBoxY = doc.y;
    const halfW = 250;

    // Bill To Box (Left)
    doc.roundedRect(40, metaBoxY, halfW, 76, 4).fillAndStroke('#f8fafc', '#e2e8f0');
    doc.fillColor('#0f172a');
    doc.fontSize(8.5).font(fontBold).fillColor(primaryColor).text('BILL TO (BUYER):', 48, metaBoxY + 6);
    doc.fontSize(9).font(fontBold).fillColor('#0f172a').text(client.name || 'Walk-in Customer', 48, metaBoxY + 18, { width: halfW - 16 });
    doc.fontSize(8).font(fontRegular).fillColor('#475569');
    doc.text(`${client.address1 || ''}${client.city ? ', ' + client.city : ''}${client.state ? ', ' + client.state : ''}`, 48, metaBoxY + 30, { width: halfW - 16 });
    if (client.gstin) {
      doc.font(fontBold).fillColor('#0f172a').text(`GSTIN: ${client.gstin}`, 48, metaBoxY + 44);
    }
    doc.font(fontRegular).fillColor('#475569').text(`State Code: ${client.stateCode || '-'} | Place of Supply: ${invoice.placeOfSupply || client.state || '-'}`, 48, metaBoxY + 56);

    // Invoice Details Box (Right)
    const rightBoxX = 305;
    doc.roundedRect(rightBoxX, metaBoxY, halfW, 76, 4).fillAndStroke('#f8fafc', '#e2e8f0');
    doc.fontSize(8.5).font(fontBold).fillColor(primaryColor).text('INVOICE DETAILS:', rightBoxX + 8, metaBoxY + 6);
    doc.fontSize(8).font(fontBold).fillColor('#0f172a');
    doc.text('Invoice No:', rightBoxX + 8, metaBoxY + 18);
    doc.font(fontRegular).text(invoice.invoiceNumber || '-', rightBoxX + 80, metaBoxY + 18);

    doc.font(fontBold).text('Date:', rightBoxX + 8, metaBoxY + 30);
    doc.font(fontRegular).text(invoice.date ? new Date(invoice.date).toLocaleDateString('en-IN') : '-', rightBoxX + 80, metaBoxY + 30);

    doc.font(fontBold).text('Reverse Charge:', rightBoxX + 8, metaBoxY + 42);
    doc.font(fontRegular).text(invoice.reverseCharge ? 'Yes' : 'No', rightBoxX + 80, metaBoxY + 42);

    doc.font(fontBold).text('Series:', rightBoxX + 8, metaBoxY + 54);
    doc.font(fontRegular).text(settings.invoiceSeries || 'Default', rightBoxX + 80, metaBoxY + 54);

    doc.y = metaBoxY + 84;

    // --- 3. DYNAMIC LINE ITEMS TABLE ---
    const showHsn = settings.showHsn !== 'false' && settings.showHsn !== false;
    const showQty = settings.showQty !== 'false' && settings.showQty !== false;
    const showRate = settings.showRate !== 'false' && settings.showRate !== false;
    const showDiscount = settings.showDiscount !== 'false' && settings.showDiscount !== false;
    const showTaxable = settings.showTaxable !== 'false' && settings.showTaxable !== false;
    const showGstRate = settings.showGstRate !== 'false' && settings.showGstRate !== false;

    // Define table columns dynamically
    const cols = [{ key: 'num', title: '#', width: 25, align: 'center' }];
    if (showHsn) cols.push({ key: 'hsn', title: 'HSN/SAC', width: 55, align: 'center' });
    if (showQty) {
      cols.push({ key: 'qty', title: 'Qty', width: 40, align: 'right' });
      cols.push({ key: 'unit', title: 'Unit', width: 38, align: 'center' });
    }
    if (showRate) cols.push({ key: 'rate', title: 'Rate (₹)', width: 52, align: 'right' });
    if (showDiscount) cols.push({ key: 'disc', title: 'Disc%', width: 40, align: 'right' });
    if (showTaxable) cols.push({ key: 'taxable', title: 'Taxable (₹)', width: 62, align: 'right' });
    if (showGstRate) cols.push({ key: 'gst', title: 'GST%', width: 40, align: 'right' });
    cols.push({ key: 'total', title: 'Total (₹)', width: 63, align: 'right' });

    // Compute description width to fit exactly 515pt
    const sumFixed = cols.reduce((sum, c) => sum + c.width, 0);
    const descWidth = Math.max(100, 515 - sumFixed);
    // Insert description right after '#'
    cols.splice(1, 0, { key: 'desc', title: 'Item / Service Description', width: descWidth, align: 'left' });

    // Header Row
    let yPos = doc.y;
    doc.rect(40, yPos, 515, 20).fill(primaryColor);
    doc.fillColor('white').fontSize(8).font(fontBold);
    let curX = 40;
    cols.forEach(c => {
      doc.text(c.title, curX + 2, yPos + 5, { width: c.width - 4, align: c.align });
      curX += c.width;
    });
    yPos += 20;

    // Line Items Rows
    doc.fontSize(7.5).font(fontRegular);
    items.forEach((item, idx) => {
      if (yPos > 680) {
        doc.addPage();
        yPos = 40;
      }
      const taxable = item.taxableValue !== undefined ? Number(item.taxableValue) : (item.quantity * item.rate - (item.discountAmount || 0));
      const lineTotal = item.totalAmount !== undefined ? Number(item.totalAmount) : (taxable + (item.cgstAmount || 0) + (item.sgstAmount || 0) + (item.igstAmount || 0));

      const rowMap = {
        num: idx + 1,
        desc: item.productName || 'Item',
        hsn: item.hsn || '-',
        qty: Number(item.quantity || 1).toFixed(item.unit === 'PCS' || item.unit === 'NOS' ? 0 : 2),
        unit: item.unit || '',
        rate: Number(item.rate || 0).toFixed(2),
        disc: Number(item.discountPercent || 0).toFixed(1),
        taxable: Number(taxable || 0).toFixed(2),
        gst: Number(item.gstRate || 0).toFixed(1),
        total: Number(lineTotal || 0).toFixed(2)
      };

      const fill = idx % 2 === 0 ? '#f8fafc' : '#ffffff';
      doc.rect(40, yPos, 515, 18).fillAndStroke(fill, '#e2e8f0');
      doc.fillColor('#0f172a');

      let rowX = 40;
      cols.forEach(c => {
        doc.text(String(rowMap[c.key] || ''), rowX + 2, yPos + 4, { width: c.width - 4, align: c.align });
        rowX += c.width;
      });
      yPos += 18;
    });

    doc.y = yPos + 10;

    // --- 4. TOTALS & SUMMARY ---
    const summaryStartY = doc.y;
    const leftBlockW = 280;
    const rightBlockX = 330;
    const rightBlockW = 225;

    // Left Block: Amount in Words, Bank Details & UPI
    doc.y = summaryStartY;
    if (settings.showAmountInWords !== 'false') {
      doc.fontSize(8.5).font(fontBold).fillColor(primaryColor).text('Amount in Words:', 40, doc.y);
      doc.fontSize(8).font(fontItalic).fillColor('#1e293b').text(numberToWords(invoice.grandTotal) + ' Rupees Only', 40, doc.y + 1);
      doc.moveDown(0.6);
    }

    if (settings.showBankDetails !== 'false' && (settings.bankName || settings.bankAccount)) {
      doc.roundedRect(40, doc.y, leftBlockW, 58, 4).fillAndStroke('#f8fafc', '#e2e8f0');
      const bY = doc.y + 4;
      doc.fontSize(8).font(fontBold).fillColor(primaryColor).text('BANK PAYMENT DETAILS:', 46, bY);
      doc.fontSize(7.5).font(fontRegular).fillColor('#334155');
      doc.text(`Bank: ${settings.bankName || '-'} | A/c No: ${settings.bankAccount || '-'}`, 46, bY + 12);
      doc.text(`IFSC: ${settings.bankIfsc || '-'} | Branch: ${settings.bankBranch || '-'}`, 46, bY + 24);
      if (settings.companyUpiId && settings.showUpi !== 'false') {
        doc.font(fontBold).fillColor('#059669').text(`UPI ID for Instant Payment: ${settings.companyUpiId}`, 46, bY + 36);
      }
      doc.y += 64;
    }

    // Right Block: Tax Summary Box
    doc.y = summaryStartY;
    const taxRows = [
      ['Taxable Subtotal', Number(invoice.subtotal || 0).toFixed(2)]
    ];
    if (Number(invoice.totalCgst || 0) > 0) taxRows.push(['CGST Output', Number(invoice.totalCgst).toFixed(2)]);
    if (Number(invoice.totalSgst || 0) > 0) taxRows.push(['SGST Output', Number(invoice.totalSgst).toFixed(2)]);
    if (Number(invoice.totalIgst || 0) > 0) taxRows.push(['IGST Output', Number(invoice.totalIgst).toFixed(2)]);
    if (Number(invoice.discount || 0) > 0) taxRows.push(['Invoice Discount', '-' + Number(invoice.discount).toFixed(2)]);
    if (Number(invoice.roundOff || 0) !== 0) taxRows.push(['Round Off', Number(invoice.roundOff).toFixed(2)]);
    taxRows.push(['Grand Total (INR)', Number(invoice.grandTotal || 0).toFixed(2)]);

    taxRows.forEach((row, i) => {
      const isGrand = i === taxRows.length - 1;
      const curY = summaryStartY + (i * 18);
      const fill = isGrand ? primaryColor : (i % 2 === 0 ? '#f8fafc' : '#ffffff');
      doc.rect(rightBlockX, curY, rightBlockW, 18).fillAndStroke(fill, '#cbd5e1');
      doc.fillColor(isGrand ? '#ffffff' : '#0f172a');
      doc.font(isGrand ? fontBold : fontRegular).fontSize(isGrand ? 9 : 8);
      doc.text(row[0], rightBlockX + 6, curY + 4, { width: 120 });
      doc.text('₹ ' + row[1], rightBlockX + 130, curY + 4, { width: 85, align: 'right' });
    });

    const afterSummaryY = Math.max(doc.y, summaryStartY + (taxRows.length * 18) + 12);
    doc.y = afterSummaryY;

    // --- 5. TERMS, DECLARATION & SIGNATURE ---
    if (doc.y > 690) {
      doc.addPage();
      doc.y = 40;
    }

    const footerY = doc.y;

    // Terms & Declaration (Left)
    const termsW = 320;
    doc.y = footerY;
    if (settings.showTerms !== 'false' && settings.termsConditions) {
      doc.fontSize(8).font(fontBold).fillColor(primaryColor).text('Terms & Conditions:', 40, doc.y);
      doc.fontSize(7).font(fontRegular).fillColor('#475569');
      const termsList = settings.termsConditions.split('\n');
      termsList.forEach(t => {
        if (t.trim()) doc.text(`• ${t.trim()}`, 40, doc.y + 1, { width: termsW });
      });
      doc.moveDown(0.4);
    }

    if (settings.invoiceDeclaration) {
      doc.fontSize(7).font(fontItalic).fillColor('#64748b').text(settings.invoiceDeclaration, 40, doc.y + 2, { width: termsW });
    }

    // Signature Block (Right)
    if (settings.showSignature !== 'false') {
      const sigX = 390;
      const sigW = 160;
      doc.y = footerY;
      const compSignPrefix = settings.signatoryCompany || (`For ${settings.companyName || 'TaxTech ERP'}`);
      doc.fontSize(8).font(fontBold).fillColor('#0f172a').text(compSignPrefix, sigX, doc.y, { width: sigW, align: 'center' });
      doc.moveDown(2.5);
      doc.strokeColor('#94a3b8').lineWidth(1).moveTo(sigX + 10, doc.y).lineTo(sigX + sigW - 10, doc.y).stroke();
      doc.fontSize(8).font(fontRegular).fillColor('#475569').text(settings.signatoryLabel || 'Authorised Signatory', sigX, doc.y + 4, { width: sigW, align: 'center' });
    }

    doc.end();

    stream.on('finish', () => {
      resolve({ success: true, path: filePath });
    });
    stream.on('error', (err) => {
      resolve({ success: false, error: err.message });
    });
  });
});

ipcMain.handle('open-file', async (_, filePath) => {
  shell.openPath(filePath);
});

// Database & Backup Functions
ipcMain.handle('get-db-path', async () => {
  ensureDataDir();
  return EXCEL_FILE;
});

ipcMain.handle('open-db-folder', async () => {
  ensureDataDir();
  shell.showItemInFolder(EXCEL_FILE);
  return true;
});

ipcMain.handle('open-db-file', async () => {
  ensureDataDir();
  await shell.openPath(EXCEL_FILE);
  return true;
});

ipcMain.handle('backup-data', async () => {
  ensureDataDir();
  const dateStr = new Date().toISOString().split('T')[0];
  const { filePath } = await dialog.showSaveDialog({
    title: 'Backup GST Billing Database',
    defaultPath: `gst_billing_backup_${dateStr}.xlsx`,
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });
  if (!filePath) return { success: false, cancelled: true };

  try {
    fs.copyFileSync(EXCEL_FILE, filePath);
    return { success: true, path: filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('restore-data', async () => {
  ensureDataDir();
  const { filePaths } = await dialog.showOpenDialog({
    title: 'Select Backup File to Restore',
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    properties: ['openFile']
  });
  if (!filePaths || filePaths.length === 0) return { success: false, cancelled: true };

  const backupPath = filePaths[0];
  try {
    const testWorkbook = new ExcelJS.Workbook();
    await testWorkbook.xlsx.readFile(backupPath);
    fs.copyFileSync(backupPath, EXCEL_FILE);
    return { success: true, path: backupPath };
  } catch (err) {
    return { success: false, error: 'Invalid or corrupt Excel backup file: ' + err.message };
  }
});

ipcMain.handle('export-excel', async (_, { sheetName } = {}) => {
  ensureDataDir();
  const dateStr = new Date().toISOString().split('T')[0];
  const defaultName = sheetName ? `GST_${sheetName}_${dateStr}.xlsx` : `GST_Billing_Database_${dateStr}.xlsx`;
  const { filePath } = await dialog.showSaveDialog({
    title: sheetName ? `Export ${sheetName} to Excel` : 'Export Database to Excel',
    defaultPath: defaultName,
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });
  if (!filePath) return { success: false, cancelled: true };

  try {
    if (!sheetName) {
      fs.copyFileSync(EXCEL_FILE, filePath);
    } else {
      const srcWb = new ExcelJS.Workbook();
      await srcWb.xlsx.readFile(EXCEL_FILE);
      const srcSheet = srcWb.getWorksheet(sheetName);
      if (!srcSheet) {
        return { success: false, error: `Sheet ${sheetName} not found` };
      }
      const newWb = new ExcelJS.Workbook();
      const newSheet = newWb.addWorksheet(sheetName);
      const cols = SHEET_COLUMNS[sheetName];
      if (cols) newSheet.columns = cols;
      srcSheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          newSheet.addRow(row.values.slice(1));
        }
      });
      await newWb.xlsx.writeFile(filePath);
    }
    return { success: true, path: filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('reset-data', async () => {
  ensureDataDir();
  try {
    if (fs.existsSync(EXCEL_FILE)) {
      fs.unlinkSync(EXCEL_FILE);
    }
    await initializeExcel();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// Load external Excel file as active database
ipcMain.handle('load-excel-db', async () => {
  ensureDataDir();
  const { filePaths } = await dialog.showOpenDialog({
    title: 'Select Excel Database File to Load',
    filters: [{ name: 'Excel Files (*.xlsx)', extensions: ['xlsx'] }],
    properties: ['openFile']
  });
  if (!filePaths || filePaths.length === 0) return { success: false, cancelled: true };

  const selectedPath = filePaths[0];
  try {
    const testWb = new ExcelJS.Workbook();
    await testWb.xlsx.readFile(selectedPath);

    const sheetNames = testWb.worksheets.map(w => w.name);
    if (!sheetNames || sheetNames.length === 0) {
      return { success: false, error: 'The selected Excel file contains no worksheets.' };
    }

    // Safety backup of current database before overwriting
    if (fs.existsSync(EXCEL_FILE)) {
      const backupDir = path.join(DATA_DIR, 'backups');
      if (!fs.existsSync(backupDir)) fs.mkdirSync(backupDir, { recursive: true });
      const stamp = new Date().toISOString().replace(/[:.]/g, '-');
      fs.copyFileSync(EXCEL_FILE, path.join(backupDir, `gst_backup_pre_load_${stamp}.xlsx`));
    }

    // Copy selected file to active DB
    fs.copyFileSync(selectedPath, EXCEL_FILE);

    // Initialize to verify standard structure and missing sheets
    await initializeExcel();

    return { success: true, path: selectedPath, sheetCount: sheetNames.length, sheets: sheetNames };
  } catch (err) {
    return { success: false, error: 'Could not load Excel file: ' + err.message };
  }
});

// Instantly load comprehensive FY 2025-26 dummy database
ipcMain.handle('load-sample-fy2526', async () => {
  ensureDataDir();
  try {
    if (fs.existsSync(EXCEL_FILE)) {
      const backupDir = path.join(DATA_DIR, 'backups');
      if (!fs.existsSync(backupDir)) fs.mkdirSync(backupDir, { recursive: true });
      const stamp = new Date().toISOString().replace(/[:.]/g, '-');
      fs.copyFileSync(EXCEL_FILE, path.join(backupDir, `gst_backup_pre_sample_${stamp}.xlsx`));
    }

    const { createFY2526Workbook } = require('./sampleDataFy2526');
    await createFY2526Workbook(EXCEL_FILE);
    return { success: true, path: EXCEL_FILE };
  } catch (err) {
    return { success: false, error: 'Failed to load FY 2025-26 sample data: ' + err.message };
  }
});

// Export FY 2025-26 standalone dummy database to user-chosen path
ipcMain.handle('export-sample-fy2526', async () => {
  const { filePath } = await dialog.showSaveDialog({
    title: 'Save FY 2025-26 Sample Excel Database',
    defaultPath: 'gst_billing_dummy_fy2025_26.xlsx',
    filters: [{ name: 'Excel Files (*.xlsx)', extensions: ['xlsx'] }]
  });
  if (!filePath) return { success: false, cancelled: true };

  try {
    const { createFY2526Workbook } = require('./sampleDataFy2526');
    await createFY2526Workbook(filePath);
    return { success: true, path: filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

function numberToWords(num) {
  const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function convertHundreds(n) {
    if (n < 20) return a[n];
    if (n < 100) return b[Math.floor(n / 10)] + (n % 10 ? ' ' + a[n % 10] : '');
    return a[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' and ' + convertHundreds(n % 100) : '');
  }

  if (num === 0) return 'Zero';
  const crore = Math.floor(num / 10000000);
  const lakh = Math.floor((num % 10000000) / 100000);
  const thousand = Math.floor((num % 100000) / 1000);
  const hundred = Math.floor((num % 1000) / 100);
  const rest = Math.floor(num % 100);
  const decimal = Math.round((num - Math.floor(num)) * 100);

  let words = '';
  if (crore) words += convertHundreds(crore) + ' Crore ';
  if (lakh) words += convertHundreds(lakh) + ' Lakh ';
  if (thousand) words += convertHundreds(thousand) + ' Thousand ';
  if (hundred) words += convertHundreds(hundred) + ' Hundred ';
  if (rest) words += (words ? 'and ' : '') + convertHundreds(rest);
  if (decimal) words += ' and ' + convertHundreds(decimal) + ' Paise';
  return words.trim();
}
