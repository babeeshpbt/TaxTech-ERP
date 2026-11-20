const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('ipcRenderer', {
  invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args),
  send: (channel, ...args) => ipcRenderer.send(channel, ...args),
  on: (channel, listener) => {
    const subscription = (event, ...args) => listener(event, ...args);
    ipcRenderer.on(channel, subscription);
    return () => ipcRenderer.removeListener(channel, subscription);
  }
});

contextBridge.exposeInMainWorld('api', {
  // Products & Services
  getProducts: () => ipcRenderer.invoke('get-products'),
  addProduct: (product) => ipcRenderer.invoke('add-product', product),
  updateProduct: (product) => ipcRenderer.invoke('update-product', product),
  deleteProduct: (id) => ipcRenderer.invoke('delete-product', id),

  // Clients & Suppliers (Parties)
  getClients: () => ipcRenderer.invoke('get-clients'),
  addClient: (client) => ipcRenderer.invoke('add-client', client),
  updateClient: (client) => ipcRenderer.invoke('update-client', client),
  deleteClient: (id) => ipcRenderer.invoke('delete-client', id),

  // Invoices (Sales)
  getInvoices: () => ipcRenderer.invoke('get-invoices'),
  getInvoiceItems: () => ipcRenderer.invoke('get-invoice-items'),
  saveInvoice: (data) => ipcRenderer.invoke('save-invoice', data),
  updateInvoice: (data) => ipcRenderer.invoke('update-invoice', data),
  deleteInvoice: (id) => ipcRenderer.invoke('delete-invoice', id),
  getNextInvoiceNumber: (settings) => ipcRenderer.invoke('get-next-invoice-number', settings),

  // Purchases
  getPurchases: () => ipcRenderer.invoke('get-purchases'),
  getPurchaseItems: () => ipcRenderer.invoke('get-purchase-items'),
  savePurchase: (data) => ipcRenderer.invoke('save-purchase', data),
  updatePurchase: (data) => ipcRenderer.invoke('update-purchase', data),
  deletePurchase: (id) => ipcRenderer.invoke('delete-purchase', id),
  getNextPurchaseNumber: (settings) => ipcRenderer.invoke('get-next-purchase-number', settings),

  // Stock Register
  getStockRegister: () => ipcRenderer.invoke('get-stock-register'),

  // Ledger Groups (Standard Core)
  getLedgerGroups: () => ipcRenderer.invoke('get-ledger-groups'),
  addLedgerGroup: (group) => ipcRenderer.invoke('add-ledger-group', group),
  updateLedgerGroup: (group) => ipcRenderer.invoke('update-ledger-group', group),
  deleteLedgerGroup: (id) => ipcRenderer.invoke('delete-ledger-group', id),

  // Ledgers
  getLedgers: () => ipcRenderer.invoke('get-ledgers'),
  addLedger: (ledger) => ipcRenderer.invoke('add-ledger', ledger),
  updateLedger: (ledger) => ipcRenderer.invoke('update-ledger', ledger),
  deleteLedger: (id) => ipcRenderer.invoke('delete-ledger', id),

  // Vouchers (Receipt, Payment, Contra, Journal)
  getVouchers: () => ipcRenderer.invoke('get-vouchers'),
  saveVoucher: (voucher) => ipcRenderer.invoke('save-voucher', voucher),
  deleteVoucher: (id) => ipcRenderer.invoke('delete-voucher', id),
  getNextVoucherNumber: (type) => ipcRenderer.invoke('get-next-voucher-number', type),

  // Financial Reports (Trial Balance, Trading A/c, P&L, Balance Sheet, Cash & Bank Books)
  getFinancialReports: () => ipcRenderer.invoke('get-financial-reports'),

  // Settings
  getSettings: () => ipcRenderer.invoke('get-settings'),
  saveSettings: (settings) => ipcRenderer.invoke('save-settings', settings),

  // PDF Generation
  generateInvoicePDF: (data) => ipcRenderer.invoke('generate-invoice-pdf', data),
  openFile: (path) => ipcRenderer.invoke('open-file', path),

  // Database, Backup & Local Excel Storage
  getDbPath: () => ipcRenderer.invoke('get-db-path'),
  openDbFolder: () => ipcRenderer.invoke('open-db-folder'),
  openDbFile: () => ipcRenderer.invoke('open-db-file'),
  backupData: () => ipcRenderer.invoke('backup-data'),
  restoreData: () => ipcRenderer.invoke('restore-data'),
  exportExcel: (options) => ipcRenderer.invoke('export-excel', options),
  resetData: () => ipcRenderer.invoke('reset-data')
});
