// Renderer Process for TaxTech ERP Suite

// State
let products = [];
let clients = [];
let invoices = [];
let invoiceItems = [];
let purchases = [];
let purchaseItems = [];
let vouchers = [];
let ledgerGroups = [];
let ledgers = [];
let stockRegister = [];
let financialReports = null;
let settings = {};

let editingInvoiceId = null;
let editingPurchaseId = null;
let editingProductId = null;
let editingClientId = null;
let editingVoucherId = null;
let editingLedgerId = null;

let itemCounter = 0;
let purchaseItemCounter = 0;

let currentProductFilter = 'all';
let currentPartyFilter = 'all';
let currentVoucherFilter = 'all';
let currentCashBankBook = 'cash';
let currentLedgerMasterTab = 'ledgers';
let currentFinancialSheet = 'trialBalance';

// Indian States for Place of Supply
const indianStates = [
    { code: '01', name: 'Jammu & Kashmir' },
    { code: '02', name: 'Himachal Pradesh' },
    { code: '03', name: 'Punjab' },
    { code: '04', name: 'Chandigarh' },
    { code: '05', name: 'Uttarakhand' },
    { code: '06', name: 'Haryana' },
    { code: '07', name: 'Delhi' },
    { code: '08', name: 'Rajasthan' },
    { code: '09', name: 'Uttar Pradesh' },
    { code: '10', name: 'Bihar' },
    { code: '11', name: 'Sikkim' },
    { code: '12', name: 'Arunachal Pradesh' },
    { code: '13', name: 'Nagaland' },
    { code: '14', name: 'Manipur' },
    { code: '15', name: 'Mizoram' },
    { code: '16', name: 'Tripura' },
    { code: '17', name: 'Meghalaya' },
    { code: '18', name: 'Assam' },
    { code: '19', name: 'West Bengal' },
    { code: '20', name: 'Jharkhand' },
    { code: '21', name: 'Odisha' },
    { code: '22', name: 'Chhattisgarh' },
    { code: '23', name: 'Madhya Pradesh' },
    { code: '24', name: 'Gujarat' },
    { code: '25', name: 'Daman & Diu' },
    { code: '26', name: 'Dadra & Nagar Haveli' },
    { code: '27', name: 'Maharashtra' },
    { code: '28', name: 'Andhra Pradesh' },
    { code: '29', name: 'Karnataka' },
    { code: '30', name: 'Goa' },
    { code: '31', name: 'Lakshadweep' },
    { code: '32', name: 'Kerala' },
    { code: '33', name: 'Tamil Nadu' },
    { code: '34', name: 'Puducherry' },
    { code: '35', name: 'Andaman & Nicobar Islands' },
    { code: '36', name: 'Telangana' },
    { code: '37', name: 'Andhra Pradesh (New)' },
    { code: '38', name: 'Ladakh' },
    { code: '97', name: 'Other Territory' },
    { code: '99', name: 'Center Jurisdiction' }
];

// Initialize App
function initApp() {
    setupEventListeners();
    populateStateDropdowns();
    loadAllData().then(() => {
        updateDashboard();
        renderAllTables();
    }).catch(err => {
        console.error('Error during initial load:', err);
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}

// Load all data from main process
async function loadAllData() {
    try {
        const [
            p, c, inv, invItems, pur, purItems,
            vch, groups, ldg, stock, fin, set, dbPath
        ] = await Promise.all([
            ipcRenderer.invoke('get-products').catch(() => []),
            ipcRenderer.invoke('get-clients').catch(() => []),
            ipcRenderer.invoke('get-invoices').catch(() => []),
            ipcRenderer.invoke('get-invoice-items').catch(() => []),
            ipcRenderer.invoke('get-purchases').catch(() => []),
            ipcRenderer.invoke('get-purchase-items').catch(() => []),
            ipcRenderer.invoke('get-vouchers').catch(() => []),
            ipcRenderer.invoke('get-ledger-groups').catch(() => []),
            ipcRenderer.invoke('get-ledgers').catch(() => []),
            ipcRenderer.invoke('get-stock-register').catch(() => []),
            ipcRenderer.invoke('get-financial-reports').catch(() => null),
            ipcRenderer.invoke('get-settings').catch(() => ({})),
            ipcRenderer.invoke('get-db-path').catch(() => '')
        ]);

        products = p || [];
        clients = c || [];
        invoices = inv || [];
        invoiceItems = invItems || [];
        purchases = pur || [];
        purchaseItems = purItems || [];
        vouchers = vch || [];
        ledgerGroups = groups || [];
        ledgers = ldg || [];
        stockRegister = stock || [];
        financialReports = fin;
        settings = set || {};

        populateSettingsForm(dbPath);
        populateLedgerGroupDropdown();
        console.log('All ERP data loaded successfully.');
    } catch (error) {
        console.error('Error loading data:', error);
        showToast('Error loading accounting data', 'error');
    }
}

// Setup Event Listeners
function setupEventListeners() {
    // Navigation
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => switchView(item.dataset.view));
    });

    // Mobile Menu toggle
    document.getElementById('menuToggle').addEventListener('click', () => {
        document.getElementById('sidebar').classList.toggle('open');
    });

    // Close modal on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
        overlay.addEventListener('click', closeAllModals);
    });

    // Prevent clicks inside modal dialogs from bubbling to the overlay
    document.querySelectorAll('.modal-content').forEach(content => {
        content.addEventListener('click', (e) => {
            e.stopPropagation();
        });
    });

    // Top Bar Buttons
    document.getElementById('topVoucherBtn')?.addEventListener('click', () => openVoucherModal());
    document.getElementById('topPurchaseBtn')?.addEventListener('click', () => openPurchaseModal());
    document.getElementById('globalNewInvoiceBtn')?.addEventListener('click', () => openInvoiceModal());
    initTheme();

    // Global click listener to close combobox menus on outside click
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.combobox-container')) {
            closeAllComboboxes();
        }
    });

    // Dashboard Quick Actions
    document.querySelector('[data-action="newInvoice"]')?.addEventListener('click', () => openInvoiceModal());
    document.querySelector('[data-action="newPurchase"]')?.addEventListener('click', () => openPurchaseModal());
    document.querySelector('[data-action="newVoucher"]')?.addEventListener('click', () => openVoucherModal());
    document.querySelector('[data-action="newProduct"]')?.addEventListener('click', () => openProductModal());
    document.querySelector('[data-action="stockRegister"]')?.addEventListener('click', () => switchView('stockRegister'));
    document.querySelector('[data-action="financials"]')?.addEventListener('click', () => switchView('financials'));
    document.getElementById('viewAllInvoicesBtn')?.addEventListener('click', () => switchView('invoices'));

    // View Action Buttons
    document.getElementById('newInvoiceBtn')?.addEventListener('click', () => openInvoiceModal());
    document.getElementById('newPurchaseBtn')?.addEventListener('click', () => openPurchaseModal());
    document.getElementById('newProductBtn')?.addEventListener('click', () => openProductModal());
    document.getElementById('newClientBtn')?.addEventListener('click', () => openClientModal());
    document.getElementById('newVoucherBtn')?.addEventListener('click', () => openVoucherModal());
    document.getElementById('newLedgerBtn')?.addEventListener('click', () => openLedgerModal());

    // Search Inputs
    document.getElementById('invoiceSearch')?.addEventListener('input', (e) => filterTable('invoicesTable', e.target.value));
    document.getElementById('purchaseSearch')?.addEventListener('input', (e) => filterTable('purchasesTable', e.target.value));
    document.getElementById('productSearch')?.addEventListener('input', (e) => filterTable('productsTable', e.target.value));
    document.getElementById('stockSearch')?.addEventListener('input', (e) => filterTable('stockRegisterTable', e.target.value));
    document.getElementById('clientSearch')?.addEventListener('input', (e) => filterTable('clientsTable', e.target.value));
    document.getElementById('voucherSearch')?.addEventListener('input', (e) => filterTable('vouchersTable', e.target.value));

    // Filter Pills
    // 1. Products Filter (All / Goods / Service)
    document.querySelectorAll('#productFilterTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#productFilterTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentProductFilter = btn.dataset.filter;
            renderProductsTable();
        });
    });

    // 2. Party Filter (All / Customer / Supplier)
    document.querySelectorAll('#partyFilterTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#partyFilterTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentPartyFilter = btn.dataset.filter;
            renderClientsTable();
        });
    });

    // 3. Voucher Filter (All / Receipt / Payment / Contra / Journal)
    document.querySelectorAll('#voucherFilterTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#voucherFilterTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentVoucherFilter = btn.dataset.filter;
            renderVouchersTable();
        });
    });

    // 4. Cash / Bank Book Tabs
    document.querySelectorAll('#cashBankTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#cashBankTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentCashBankBook = btn.dataset.book;
            renderCashBankRegister();
        });
    });

    // 5. Ledger Master Tabs (Ledgers / 28 Groups)
    document.querySelectorAll('#ledgerMasterTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#ledgerMasterTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentLedgerMasterTab = btn.dataset.tab;
            if (currentLedgerMasterTab === 'ledgers') {
                document.getElementById('ledgersSubView').style.display = 'block';
                document.getElementById('groupsSubView').style.display = 'none';
            } else {
                document.getElementById('ledgersSubView').style.display = 'none';
                document.getElementById('groupsSubView').style.display = 'block';
            }
        });
    });

    // 6. Financial Statements Tabs
    document.querySelectorAll('#financialTabs .pill-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#financialTabs .pill-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentFinancialSheet = btn.dataset.sheet;
            showFinancialSheet(currentFinancialSheet);
        });
    });

    document.getElementById('refreshFinancialsBtn')?.addEventListener('click', async () => {
        await loadAllData();
        renderFinancialStatements();
        showToast('Financial reports updated', 'success');
    });

    // Export to Excel Buttons
    document.getElementById('exportInvoicesExcelBtn')?.addEventListener('click', () => exportExcel('Invoices'));
    document.getElementById('exportPurchasesExcelBtn')?.addEventListener('click', () => exportExcel('Purchases'));
    document.getElementById('exportProductsExcelBtn')?.addEventListener('click', () => exportExcel('Products'));
    document.getElementById('exportStockExcelBtn')?.addEventListener('click', () => exportStockToExcel());
    document.getElementById('exportClientsExcelBtn')?.addEventListener('click', () => exportExcel('Clients'));
    document.getElementById('exportVouchersExcelBtn')?.addEventListener('click', () => exportExcel('Vouchers'));
    document.getElementById('exportCashBankExcelBtn')?.addEventListener('click', () => exportExcel('Vouchers'));
    document.getElementById('exportExcelBtn')?.addEventListener('click', () => exportExcel());

    // Invoice Modal
    document.getElementById('closeInvoiceModal')?.addEventListener('click', closeInvoiceModal);
    document.getElementById('cancelInvoiceBtn')?.addEventListener('click', closeInvoiceModal);
    document.getElementById('invoiceForm')?.addEventListener('submit', (e) => { e.preventDefault(); saveInvoice(); });
    document.getElementById('addItemBtn')?.addEventListener('click', addInvoiceItem);
    document.getElementById('saveInvoiceBtn')?.addEventListener('click', saveInvoice);
    document.getElementById('generatePdfBtn')?.addEventListener('click', generateInvoicePDF);
    document.getElementById('invoiceDiscount')?.addEventListener('input', calculateInvoiceTotals);
    document.getElementById('invoiceRoundOff')?.addEventListener('input', calculateInvoiceTotals);
    document.getElementById('invoiceClient')?.addEventListener('change', onClientChange);

    // Purchase Modal
    document.getElementById('closePurchaseModal')?.addEventListener('click', closePurchaseModal);
    document.getElementById('cancelPurchaseBtn')?.addEventListener('click', closePurchaseModal);
    document.getElementById('purchaseForm')?.addEventListener('submit', (e) => { e.preventDefault(); savePurchase(); });
    document.getElementById('addPurchaseItemBtn')?.addEventListener('click', addPurchaseItem);
    document.getElementById('savePurchaseBtn')?.addEventListener('click', savePurchase);
    document.getElementById('purchaseDiscount')?.addEventListener('input', calculatePurchaseTotals);
    document.getElementById('purchaseRoundOff')?.addEventListener('input', calculatePurchaseTotals);
    document.getElementById('purchaseSupplier')?.addEventListener('change', onSupplierChange);

    // Product Modal & Type Radio Toggle
    document.getElementById('closeProductModal')?.addEventListener('click', closeProductModal);
    document.getElementById('cancelProductBtn')?.addEventListener('click', closeProductModal);
    document.getElementById('productForm')?.addEventListener('submit', saveProduct);
    document.querySelectorAll('input[name="type"]').forEach(radio => {
        radio.addEventListener('change', (e) => handleProductTypeToggle(e.target.value));
    });

    // Client Modal
    document.getElementById('closeClientModal')?.addEventListener('click', closeClientModal);
    document.getElementById('cancelClientBtn')?.addEventListener('click', closeClientModal);
    document.getElementById('clientForm')?.addEventListener('submit', saveClient);

    // Voucher Modal
    document.getElementById('closeVoucherModal')?.addEventListener('click', closeVoucherModal);
    document.getElementById('cancelVoucherBtn')?.addEventListener('click', closeVoucherModal);
    document.getElementById('voucherForm')?.addEventListener('submit', saveVoucher);
    document.getElementById('voucherType')?.addEventListener('change', onVoucherTypeChange);

    // Ledger Modal
    document.getElementById('closeLedgerModal')?.addEventListener('click', closeLedgerModal);
    document.getElementById('cancelLedgerBtn')?.addEventListener('click', closeLedgerModal);
    document.getElementById('ledgerForm')?.addEventListener('submit', saveLedger);

    // Settings & Database
    document.getElementById('saveSettingsBtn')?.addEventListener('click', saveSettings);
    document.getElementById('backupDataBtn')?.addEventListener('click', backupData);
    document.getElementById('restoreDataBtn')?.addEventListener('click', restoreData);
    document.getElementById('openDbFolderBtn')?.addEventListener('click', openDbFolder);
    document.getElementById('openDbFileBtn')?.addEventListener('click', openDbFile);
    document.getElementById('resetDataBtn')?.addEventListener('click', resetData);

    // Invoice Output Format, Alignment & Logo Controls
    initInvoiceFormatControls();

    // Keyboard Shortcuts
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAllModals();
        if (e.ctrlKey && e.key.toLowerCase() === 'n') {
            e.preventDefault();
            openInvoiceModal();
        }
    });
}

// Navigation
function switchView(viewName) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

    const targetView = document.getElementById(viewName + 'View');
    const targetNav = document.querySelector(`[data-view="${viewName}"]`);
    if (targetView) targetView.classList.add('active');
    if (targetNav) targetNav.classList.add('active');

    const titleMap = {
        dashboard: 'Dashboard',
        invoices: 'Sales Invoices',
        purchases: 'Purchases (Inward Bills)',
        products: 'Items & Services Master',
        stockRegister: 'Stock Register & Inventory Valuation',
        clients: 'Parties Master (Debtors & Creditors)',
        vouchers: 'Receipt & Payment Vouchers',
        cashBank: 'Cash Book & Bank Book',
        ledgers: 'Accounts & Ledgers',
        financials: 'Financial Statements (Trial Balance, P&L, Balance Sheet)',
        settings: 'Settings & Database Backup'
    };
    document.getElementById('viewTitle').textContent = titleMap[viewName] || 'TaxTech ERP';
    document.getElementById('sidebar').classList.remove('open');

    // Auto-refresh financial reports or stock register when navigating
    if (viewName === 'financials') {
        renderFinancialStatements();
    } else if (viewName === 'stockRegister') {
        renderStockRegisterTable();
    } else if (viewName === 'cashBank') {
        renderCashBankRegister();
    }
}

// Populate State Dropdowns
function populateStateDropdowns() {
    const selects = [
        document.getElementById('placeOfSupply'),
        document.getElementById('purchasePlaceOfSupply'),
        document.getElementById('clientState')
    ];

    selects.forEach(sel => {
        if (!sel || sel.tagName !== 'SELECT') return;
        const currentVal = sel.value;
        sel.innerHTML = '<option value="">Select State</option>';
        indianStates.forEach(st => {
            const opt = new Option(`${st.name} (${st.code})`, st.name);
            sel.add(opt);
        });
        if (currentVal) sel.value = currentVal;
    });

    const clientState = document.getElementById('clientState');
    const clientStateCode = document.getElementById('clientStateCode');
    if (clientState) {
        clientState.addEventListener('change', () => {
            const found = indianStates.find(s => s.name === clientState.value);
            if (found && clientStateCode) clientStateCode.value = found.code;
        });
    }
}

// Populate Ledger Groups in Modal
function populateLedgerGroupDropdown() {
    const sel = document.getElementById('ledgerGroupName');
    if (!sel) return;
    sel.innerHTML = '<option value="">Select Accounting Group</option>';

    const natureOrder = ['Liabilities', 'Assets', 'Income', 'Expenses'];
    natureOrder.forEach(nat => {
        const optgroup = document.createElement('optgroup');
        optgroup.label = `--- ${nat.toUpperCase()} ---`;
        const matched = ledgerGroups.filter(g => g.nature === nat);
        matched.forEach(g => {
            const opt = new Option(`${g.name}${g.parentGroup ? ` (Sub of ${g.parentGroup})` : ''}`, g.name);
            optgroup.appendChild(opt);
        });
        sel.appendChild(optgroup);
    });
}

// Render All Tables
function renderAllTables() {
    renderInvoicesTable();
    renderPurchasesTable();
    renderProductsTable();
    renderStockRegisterTable();
    renderClientsTable();
    renderVouchersTable();
    renderCashBankRegister();
    renderLedgersTable();
    renderGroupsTable();
    renderFinancialStatements();
    renderRecentInvoicesTable();
    updateDashboard();
}

// Update Dashboard
function updateDashboard() {
    document.getElementById('totalInvoices').textContent = invoices.length;
    const rev = invoices.reduce((sum, inv) => sum + (Number(inv.grandTotal) || 0), 0);
    document.getElementById('totalRevenue').textContent = '₹' + formatNumber(rev);

    const purTotal = purchases.reduce((sum, p) => sum + (Number(p.grandTotal) || 0), 0);
    document.getElementById('totalPurchasesVal').textContent = '₹' + formatNumber(purTotal);

    const stockTotal = stockRegister
        .filter(i => i.type === 'Goods')
        .reduce((sum, i) => sum + (Number(i.closingValue) || 0), 0);
    document.getElementById('totalStockVal').textContent = '₹' + formatNumber(stockTotal);
}

// Filter Table helper
function filterTable(tableId, query) {
    const rows = document.querySelectorAll(`#${tableId} tbody tr`);
    const q = query.toLowerCase();
    rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(q) ? '' : 'none';
    });
}

// ==========================================================
// 1. INVOICES (SALES)
// ==========================================================
function renderInvoicesTable() {
    const tbody = document.querySelector('#invoicesTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    invoices.forEach(inv => {
        const statusClass = inv.status?.toLowerCase() || 'draft';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${inv.invoiceNumber || '-'}</strong></td>
            <td>${formatDate(inv.date)}</td>
            <td>${inv.clientName || '-'}</td>
            <td>${inv.clientGstin || '-'}</td>
            <td>${inv.placeOfSupply || '-'}</td>
            <td>₹${formatNumber(inv.subtotal)}</td>
            <td>₹${formatNumber(inv.totalTax)}</td>
            <td><strong>₹${formatNumber(inv.grandTotal)}</strong></td>
            <td><span class="status-badge ${statusClass}">${inv.status || 'Draft'}</span></td>
            <td>
                <button class="btn-icon" onclick="editInvoice('${inv.id}')" title="Edit">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
                <button class="btn-icon" onclick="viewInvoice('${inv.id}')" title="PDF Invoice">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                </button>
                <button class="btn-icon delete" onclick="deleteInvoice('${inv.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderRecentInvoicesTable() {
    const tbody = document.querySelector('#recentInvoicesTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const recent = [...invoices].sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0)).slice(0, 5);
    recent.forEach(inv => {
        const statusClass = inv.status?.toLowerCase() || 'draft';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${inv.invoiceNumber || '-'}</strong></td>
            <td>${formatDate(inv.date)}</td>
            <td>${inv.clientName || '-'}</td>
            <td>₹${formatNumber(inv.grandTotal)}</td>
            <td><span class="status-badge ${statusClass}">${inv.status || 'Draft'}</span></td>
            <td>
                <button class="btn-icon" onclick="editInvoice('${inv.id}')"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
                <button class="btn-icon" onclick="viewInvoice('${inv.id}')"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

async function openInvoiceModal(invoice = null) {
    editingInvoiceId = invoice?.id || null;
    document.getElementById('invoiceModalTitle').textContent = invoice ? 'Edit Sales Invoice' : 'Create Sales Invoice';
    document.getElementById('generatePdfBtn').style.display = invoice ? 'inline-flex' : 'none';
    document.getElementById('invoiceForm').reset();
    document.getElementById('invoiceItemsTable').querySelector('tbody').innerHTML = '';
    itemCounter = 0;

    // Populate client dropdown
    const clientSelect = document.getElementById('invoiceClient');
    clientSelect.innerHTML = '<option value="">Select Customer</option>';
    const customers = clients.filter(c => c.partyType !== 'Supplier');
    (customers.length > 0 ? customers : clients).forEach(c => {
        const opt = new Option(`${c.name} (${c.gstin || 'No GSTIN'})`, c.id);
        opt.dataset.client = JSON.stringify(c);
        clientSelect.add(opt);
    });

    if (invoice) {
        document.getElementById('invoiceNumber').value = invoice.invoiceNumber || '';
        document.getElementById('invoiceDate').value = invoice.date ? invoice.date.split('T')[0] : new Date().toISOString().split('T')[0];
        document.getElementById('invoiceClient').value = invoice.clientId || '';
        document.getElementById('placeOfSupply').value = invoice.placeOfSupply || '';
        document.getElementById('reverseCharge').checked = !!invoice.reverseCharge;
        document.getElementById('invoiceDiscount').value = invoice.discount || 0;
        document.getElementById('invoiceRoundOff').value = invoice.roundOff || 0;
        document.getElementById('invoiceNotes').value = invoice.notes || '';

        const items = invoiceItems.filter(i => i.invoiceId === invoice.id);
        items.forEach(item => addInvoiceItemRow(item));
    } else {
        const nextNum = await ipcRenderer.invoke('get-next-invoice-number', settings);
        document.getElementById('invoiceNumber').value = nextNum;
        document.getElementById('invoiceDate').value = new Date().toISOString().split('T')[0];
        addInvoiceItemRow();
    }

    calculateInvoiceTotals();
    document.getElementById('invoiceModal').classList.add('active');
}

function closeInvoiceModal() {
    document.getElementById('invoiceModal').classList.remove('active');
    editingInvoiceId = null;
}

function onClientChange() {
    const select = document.getElementById('invoiceClient');
    const opt = select.options[select.selectedIndex];
    if (opt && opt.dataset.client) {
        const client = JSON.parse(opt.dataset.client);
        if (client.state) {
            document.getElementById('placeOfSupply').value = client.state;
            calculateInvoiceTotals();
        }
    }
}

// Helper: Escape HTML strings
function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// ==========================================================
// THEME MANAGEMENT (LIGHT / DARK ENTERPRISE MODE)
// ==========================================================
function initTheme() {
    const savedTheme = localStorage.getItem('taxtech_theme') || 'light';
    applyTheme(savedTheme);

    const toggleBtn = document.getElementById('themeToggleBtn');
    if (toggleBtn) {
        toggleBtn.addEventListener('click', () => {
            const current = document.body.dataset.theme === 'dark' ? 'dark' : 'light';
            const next = current === 'dark' ? 'light' : 'dark';
            applyTheme(next);
            localStorage.setItem('taxtech_theme', next);
        });
    }
}

function applyTheme(theme) {
    const sunIcon = document.querySelector('#themeToggleBtn .sun-icon');
    const moonIcon = document.querySelector('#themeToggleBtn .moon-icon');
    if (theme === 'dark') {
        document.body.dataset.theme = 'dark';
        if (sunIcon) sunIcon.style.display = 'inline-block';
        if (moonIcon) moonIcon.style.display = 'none';
    } else {
        delete document.body.dataset.theme;
        if (sunIcon) sunIcon.style.display = 'none';
        if (moonIcon) moonIcon.style.display = 'inline-block';
    }
}

// ==========================================================
// SEARCHABLE COMBOBOX COMPONENT FOR INVOICES & PURCHASES
// ==========================================================
function closeAllComboboxes() {
    document.querySelectorAll('.combobox-menu').forEach(menu => {
        menu.style.display = 'none';
    });
}

function setupProductCombobox(container, mode, onSelectCallback) {
    if (!container) return;
    const input = container.querySelector('input[type="text"]');
    const hidden = container.querySelector('input[type="hidden"]');
    const menu = container.querySelector('.combobox-menu');
    const arrowBtn = container.querySelector('.combobox-arrow-btn');
    const tr = container.closest('tr');

    // Pre-populate if hidden value exists or item already assigned
    if (hidden && hidden.value) {
        const existing = products.find(p => p.id === hidden.value);
        if (existing) {
            const badge = (existing.type === 'Service' || (existing.type || '').toLowerCase().includes('service')) ? '[Service]' : '[Goods]';
            input.value = `${badge} ${existing.name}`;
        }
    }

    function renderMenu(query = '') {
        const q = (query || '').trim().toLowerCase();
        // Clean out any existing [Goods] or [Service] prefixes from search query
        const cleanQ = q.replace(/^\[(goods|service)\]\s*/i, '');
        const filtered = products.filter(p => {
            if (!cleanQ) return true;
            return (p.name && p.name.toLowerCase().includes(cleanQ)) ||
                   (p.hsn && p.hsn.toLowerCase().includes(cleanQ)) ||
                   (p.category && p.category.toLowerCase().includes(cleanQ)) ||
                   (p.type && p.type.toLowerCase().includes(cleanQ));
        });

        if (filtered.length === 0) {
            menu.innerHTML = `<div class="combobox-empty">No matching items found.<br><small style="color:var(--text-muted);">Press Enter or Tab to use custom description.</small></div>`;
            menu.style.display = 'block';
            return;
        }

        const slice = filtered.slice(0, 50);
        menu.innerHTML = slice.map((p, idx) => {
            const isService = (p.type === 'Service' || (p.type || '').toLowerCase().includes('service'));
            const badgeClass = isService ? 'badge-service' : 'badge-goods';
            const badgeText = isService ? 'Service' : 'Goods';
            const price = mode === 'purchase' ? (p.purchasePrice || p.salePrice || 0) : (p.salePrice || 0);
            const stockDisplay = isService ? 'Non-Stock Service' : `Stock: ${formatNumber(p.stock || p.openingStock || 0)} ${p.unit || 'PCS'}`;
            return `
                <div class="combobox-item ${idx === 0 ? 'active' : ''}" data-id="${p.id}" data-idx="${idx}">
                    <div class="combobox-item-top">
                        <span class="combobox-badge ${badgeClass}">${badgeText}</span>
                        <span class="combobox-item-name">${escapeHtml(p.name)}</span>
                    </div>
                    <div class="combobox-item-bottom">
                        <span class="combobox-meta-hsn">HSN/SAC: ${escapeHtml(p.hsn || '-')}</span>
                        <span class="combobox-meta-stock">${stockDisplay}</span>
                        <span class="combobox-meta-price">₹${formatNumber(price)}</span>
                    </div>
                </div>
            `;
        }).join('');
        menu.style.display = 'block';
    }

    function selectProduct(prod) {
        if (!prod) return;
        hidden.value = prod.id;
        const badge = (prod.type === 'Service' || (prod.type || '').toLowerCase().includes('service')) ? '[Service]' : '[Goods]';
        input.value = `${badge} ${prod.name}`;
        menu.style.display = 'none';

        if (typeof onSelectCallback === 'function') {
            onSelectCallback(prod, tr);
        }

        // Auto-advance focus to quantity input
        const qtyInput = tr.querySelector(mode === 'purchase' ? '.pitem-qty' : '.item-qty');
        if (qtyInput) {
            qtyInput.focus();
            qtyInput.select();
        }
    }

    input.addEventListener('focus', () => {
        closeAllComboboxes();
        renderMenu(input.value);
    });

    input.addEventListener('input', () => {
        renderMenu(input.value);
    });

    if (arrowBtn) {
        arrowBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (menu.style.display === 'block') {
                menu.style.display = 'none';
            } else {
                closeAllComboboxes();
                renderMenu('');
                input.focus();
            }
        });
    }

    menu.addEventListener('mousedown', (e) => {
        const itemEl = e.target.closest('.combobox-item');
        if (itemEl) {
            e.preventDefault();
            const prodId = itemEl.dataset.id;
            const prod = products.find(p => p.id === prodId);
            if (prod) {
                selectProduct(prod);
            }
        }
    });

    input.addEventListener('keydown', (e) => {
        if (menu.style.display !== 'block') {
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                closeAllComboboxes();
                renderMenu(input.value);
                e.preventDefault();
                return;
            }
        }

        const items = Array.from(menu.querySelectorAll('.combobox-item'));
        const activeIdx = items.findIndex(el => el.classList.contains('active'));

        if (e.key === 'ArrowDown') {
            e.preventDefault();
            if (items.length > 0) {
                const nextIdx = activeIdx < items.length - 1 ? activeIdx + 1 : 0;
                items.forEach(el => el.classList.remove('active'));
                items[nextIdx].classList.add('active');
                items[nextIdx].scrollIntoView({ block: 'nearest' });
            }
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            if (items.length > 0) {
                const prevIdx = activeIdx > 0 ? activeIdx - 1 : items.length - 1;
                items.forEach(el => el.classList.remove('active'));
                items[prevIdx].classList.add('active');
                items[prevIdx].scrollIntoView({ block: 'nearest' });
            }
        } else if (e.key === 'Enter') {
            if (menu.style.display === 'block' && items.length > 0) {
                e.preventDefault();
                e.stopPropagation();
                const chosenEl = activeIdx >= 0 ? items[activeIdx] : items[0];
                const prod = products.find(p => p.id === chosenEl.dataset.id);
                if (prod) {
                    selectProduct(prod);
                } else {
                    menu.style.display = 'none';
                }
            }
        } else if (e.key === 'Escape') {
            menu.style.display = 'none';
        } else if (e.key === 'Tab') {
            menu.style.display = 'none';
        }
    });
}

function addInvoiceItemRow(item = null) {
    itemCounter++;
    const tbody = document.querySelector('#invoiceItemsTable tbody');
    const tr = document.createElement('tr');
    tr.dataset.itemId = item?.id || `new_${itemCounter}`;

    const qty = item?.quantity || 1;
    const rate = item?.rate || 0;
    const disc = item?.discountPercent || 0;
    const defaultSearchVal = item?.productName ? item.productName : '';

    tr.innerHTML = `
        <td>${tbody.children.length + 1}</td>
        <td>
            <div class="combobox-container item-combobox">
                <div class="combobox-input-wrap">
                    <input type="text" class="item-search-input" placeholder="🔍 Search item / service..." autocomplete="off" value="${escapeHtml(defaultSearchVal)}">
                    <button type="button" class="combobox-arrow-btn" tabindex="-1">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
                    </button>
                </div>
                <input type="hidden" class="item-product" value="${item?.productId || ''}">
                <div class="combobox-menu" style="display: none;"></div>
            </div>
        </td>
        <td><input type="text" class="item-hsn" value="${item?.hsn || ''}" placeholder="HSN/SAC"></td>
        <td><input type="number" class="item-qty" value="${qty}" min="0.01" step="any" oninput="calculateItemRow(this)"></td>
        <td><input type="text" class="item-unit" value="${item?.unit || 'PCS'}" style="width: 60px;"></td>
        <td><input type="number" class="item-rate" value="${rate}" min="0" step="0.01" oninput="calculateItemRow(this)"></td>
        <td><input type="number" class="item-disc" value="${disc}" min="0" max="100" step="0.01" oninput="calculateItemRow(this)"></td>
        <td><input type="text" class="item-taxable" value="0.00" readonly></td>
        <td><input type="number" class="item-gstrate" value="${item?.gstRate || 18}" min="0" max="100" step="0.01" oninput="calculateItemRow(this)"></td>
        <td><input type="text" class="item-cgst" value="0.00" readonly></td>
        <td><input type="text" class="item-sgst" value="0.00" readonly></td>
        <td><input type="text" class="item-igst" value="0.00" readonly></td>
        <td><input type="text" class="item-total" value="0.00" readonly></td>
        <td><button type="button" class="btn-icon delete" onclick="removeInvoiceItem(this)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></td>
    `;

    tbody.appendChild(tr);

    setupProductCombobox(tr.querySelector('.item-combobox'), 'sales', onProductSelect);

    calculateItemRow(tr.querySelector('.item-qty'));
    updateInvoiceItemNumbers();
}

function addInvoiceItem() {
    addInvoiceItemRow();
}

function removeInvoiceItem(btn) {
    btn.closest('tr').remove();
    updateInvoiceItemNumbers();
    calculateInvoiceTotals();
}

function updateInvoiceItemNumbers() {
    document.querySelectorAll('#invoiceItemsTable tbody tr').forEach((tr, idx) => {
        tr.cells[0].textContent = idx + 1;
    });
}

function onProductSelect(prodOrElem, trContext) {
    let prod = prodOrElem;
    let tr = trContext;
    if (prodOrElem && prodOrElem.tagName) {
        tr = prodOrElem.closest('tr');
        if (prodOrElem.options) {
            const opt = prodOrElem.options[prodOrElem.selectedIndex];
            if (opt && opt.dataset.product) prod = JSON.parse(opt.dataset.product);
        }
    }
    if (!prod || !tr) return;
    tr.querySelector('.item-hsn').value = prod.hsn || '';
    tr.querySelector('.item-unit').value = prod.unit || (prod.type === 'Service' ? 'OTH' : 'PCS');
    tr.querySelector('.item-rate').value = Number(prod.salePrice || 0);
    tr.querySelector('.item-gstrate').value = Number(prod.gstRate || 18);
    calculateItemRow(tr.querySelector('.item-qty'));
}

function calculateItemRow(input) {
    const tr = input.closest('tr');
    const qty = parseFloat(tr.querySelector('.item-qty').value) || 0;
    const rate = parseFloat(tr.querySelector('.item-rate').value) || 0;
    const discPercent = parseFloat(tr.querySelector('.item-disc').value) || 0;
    const gstRate = parseFloat(tr.querySelector('.item-gstrate').value) || 0;

    const discountAmt = (qty * rate * discPercent) / 100;
    const taxable = Math.max(0, qty * rate - discountAmt);

    const supplyState = document.getElementById('placeOfSupply').value || '';
    const companyState = settings.companyState || '';
    const isInterstate = supplyState && companyState && supplyState.toLowerCase() !== companyState.toLowerCase();

    let cgstAmt = 0;
    let sgstAmt = 0;
    let igstAmt = 0;

    if (isInterstate) {
        igstAmt = (taxable * gstRate) / 100;
    } else {
        cgstAmt = (taxable * (gstRate / 2)) / 100;
        sgstAmt = (taxable * (gstRate / 2)) / 100;
    }
    const total = taxable + cgstAmt + sgstAmt + igstAmt;

    tr.querySelector('.item-taxable').value = taxable.toFixed(2);
    tr.querySelector('.item-cgst').value = cgstAmt.toFixed(2);
    tr.querySelector('.item-sgst').value = sgstAmt.toFixed(2);
    tr.querySelector('.item-igst').value = igstAmt.toFixed(2);
    tr.querySelector('.item-total').value = total.toFixed(2);

    calculateInvoiceTotals();
}

function calculateInvoiceTotals() {
    let subtotal = 0;
    let totalCgst = 0;
    let totalSgst = 0;
    let totalIgst = 0;

    document.querySelectorAll('#invoiceItemsTable tbody tr').forEach(tr => {
        subtotal += parseFloat(tr.querySelector('.item-taxable').value) || 0;
        totalCgst += parseFloat(tr.querySelector('.item-cgst').value) || 0;
        totalSgst += parseFloat(tr.querySelector('.item-sgst').value) || 0;
        totalIgst += parseFloat(tr.querySelector('.item-igst').value) || 0;
    });

    const discount = parseFloat(document.getElementById('invoiceDiscount').value) || 0;
    const roundOff = parseFloat(document.getElementById('invoiceRoundOff').value) || 0;
    const totalTax = totalCgst + totalSgst + totalIgst;
    const grandTotal = Math.max(0, subtotal + totalTax - discount + roundOff);

    document.getElementById('summarySubtotal').textContent = '₹' + formatNumber(subtotal);
    document.getElementById('summaryCgst').textContent = '₹' + formatNumber(totalCgst);
    document.getElementById('summarySgst').textContent = '₹' + formatNumber(totalSgst);
    document.getElementById('summaryIgst').textContent = '₹' + formatNumber(totalIgst);
    document.getElementById('summaryTotalTax').textContent = '₹' + formatNumber(totalTax);
    document.getElementById('summaryDiscount').textContent = '₹' + formatNumber(discount);
    document.getElementById('summaryRoundOff').textContent = '₹' + formatNumber(roundOff);
    document.getElementById('summaryGrandTotal').textContent = '₹' + formatNumber(grandTotal);
}

async function saveInvoice() {
    const invoiceNumber = document.getElementById('invoiceNumber').value;
    const date = document.getElementById('invoiceDate').value;
    const clientId = document.getElementById('invoiceClient').value;
    const placeOfSupply = document.getElementById('placeOfSupply').value;
    const reverseCharge = document.getElementById('reverseCharge').checked;
    const discount = parseFloat(document.getElementById('invoiceDiscount').value) || 0;
    const roundOff = parseFloat(document.getElementById('invoiceRoundOff').value) || 0;
    const notes = document.getElementById('invoiceNotes').value;

    if (!invoiceNumber || !date || !clientId || !placeOfSupply) {
        showToast('Please fill all required fields', 'error');
        return;
    }

    const client = clients.find(c => c.id === clientId);
    if (!client) {
        showToast('Please select a valid customer', 'error');
        return;
    }

    const items = [];
    document.querySelectorAll('#invoiceItemsTable tbody tr').forEach(tr => {
        const prodSelect = tr.querySelector('.item-product');
        const prodId = prodSelect ? prodSelect.value : '';
        let pData = products.find(p => p.id === prodId) || null;
        const searchInput = tr.querySelector('.item-search-input');
        const rawSearchText = searchInput ? searchInput.value.replace(/^\[(Goods|Service)\]\s*/i, '').trim() : '';

        const qty = parseFloat(tr.querySelector('.item-qty').value) || 0;
        const rate = parseFloat(tr.querySelector('.item-rate').value) || 0;
        const disc = parseFloat(tr.querySelector('.item-disc').value) || 0;
        const taxable = parseFloat(tr.querySelector('.item-taxable').value) || 0;
        const gstRate = parseFloat(tr.querySelector('.item-gstrate').value) || 0;
        const cgstAmt = parseFloat(tr.querySelector('.item-cgst').value) || 0;
        const sgstAmt = parseFloat(tr.querySelector('.item-sgst').value) || 0;
        const igstAmt = parseFloat(tr.querySelector('.item-igst').value) || 0;
        const total = parseFloat(tr.querySelector('.item-total').value) || 0;

        if (taxable > 0 || qty > 0) {
            items.push({
                productId: prodId || '',
                productName: pData ? pData.name : (rawSearchText || 'Item'),
                type: pData ? (pData.type || 'Goods') : (searchInput && searchInput.value.includes('[Service]') ? 'Service' : 'Goods'),
                hsn: tr.querySelector('.item-hsn').value || '',
                quantity: qty,
                unit: tr.querySelector('.item-unit').value || 'PCS',
                rate,
                discountPercent: disc,
                discountAmount: (qty * rate * disc) / 100,
                taxableValue: taxable,
                gstRate,
                cgstPercent: cgstAmt > 0 ? gstRate / 2 : 0,
                cgstAmount: cgstAmt,
                sgstPercent: sgstAmt > 0 ? gstRate / 2 : 0,
                sgstAmount: sgstAmt,
                igstPercent: igstAmt > 0 ? gstRate : 0,
                igstAmount: igstAmt,
                totalAmount: total
            });
        }
    });

    if (items.length === 0) {
        showToast('Please add at least one line item', 'error');
        return;
    }

    const subtotal = items.reduce((sum, i) => sum + i.taxableValue, 0);
    const totalCgst = items.reduce((sum, i) => sum + i.cgstAmount, 0);
    const totalSgst = items.reduce((sum, i) => sum + i.sgstAmount, 0);
    const totalIgst = items.reduce((sum, i) => sum + i.igstAmount, 0);
    const totalTax = totalCgst + totalSgst + totalIgst;
    const grandTotal = Math.max(0, subtotal + totalTax - discount + roundOff);

    const invoiceData = {
        invoiceNumber,
        date,
        clientId,
        clientName: client.name,
        clientGstin: client.gstin || '',
        placeOfSupply,
        reverseCharge: reverseCharge ? 'Yes' : 'No',
        subtotal,
        totalCgst,
        totalSgst,
        totalIgst,
        totalTax,
        discount,
        roundOff,
        grandTotal,
        status: 'Confirmed',
        notes
    };

    try {
        if (editingInvoiceId) {
            await ipcRenderer.invoke('update-invoice', { invoice: { ...invoiceData, id: editingInvoiceId }, items });
            showToast('Sales Invoice updated successfully', 'success');
        } else {
            await ipcRenderer.invoke('save-invoice', { invoice: invoiceData, items });
            showToast('Sales Invoice created successfully', 'success');
        }
        await loadAllData();
        renderAllTables();
        closeInvoiceModal();
    } catch (err) {
        showToast('Error saving invoice: ' + err.message, 'error');
    }
}

async function editInvoice(id) {
    const inv = invoices.find(i => i.id === id);
    if (inv) openInvoiceModal(inv);
}

async function viewInvoice(id) {
    const invoice = invoices.find(i => i.id === id);
    if (!invoice) return;

    const items = invoiceItems.filter(i => i.invoiceId === id);
    const client = clients.find(c => c.id === invoice.clientId) || { name: invoice.clientName };

    openInvoicePreviewModal(invoice, items, client, settings);
}

async function generateInvoicePDF() {
    if (editingInvoiceId) viewInvoice(editingInvoiceId);
}

async function deleteInvoice(id) {
    if (!confirm('Are you sure you want to delete this invoice? Stock and accounts will be updated.')) return;
    try {
        await ipcRenderer.invoke('delete-invoice', id);
        await loadAllData();
        renderAllTables();
        showToast('Invoice deleted', 'success');
    } catch (err) {
        showToast('Error deleting invoice: ' + err.message, 'error');
    }
}

// ==========================================================
// 2. PURCHASES (INWARD BILLS)
// ==========================================================
function renderPurchasesTable() {
    const tbody = document.querySelector('#purchasesTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    purchases.forEach(p => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${p.purchaseNumber || '-'}</strong></td>
            <td>${formatDate(p.date)}</td>
            <td>${p.supplierName || '-'}</td>
            <td>${p.supplierBillNo || '-'}</td>
            <td>${p.supplierGstin || '-'}</td>
            <td>₹${formatNumber(p.subtotal)}</td>
            <td>₹${formatNumber(p.totalTax)}</td>
            <td><strong>₹${formatNumber(p.grandTotal)}</strong></td>
            <td>
                <button class="btn-icon" onclick="editPurchase('${p.id}')" title="Edit">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
                <button class="btn-icon delete" onclick="deletePurchase('${p.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

async function openPurchaseModal(purchase = null) {
    editingPurchaseId = purchase?.id || null;
    document.getElementById('purchaseModalTitle').textContent = purchase ? 'Edit Purchase Bill' : 'Record Purchase Bill (Inward)';
    document.getElementById('purchaseForm').reset();
    document.getElementById('purchaseItemsTable').querySelector('tbody').innerHTML = '';
    purchaseItemCounter = 0;

    // Populate suppliers dropdown
    const suppSelect = document.getElementById('purchaseSupplier');
    suppSelect.innerHTML = '<option value="">Select Supplier</option>';
    const suppliers = clients.filter(c => c.partyType !== 'Customer');
    (suppliers.length > 0 ? suppliers : clients).forEach(c => {
        const opt = new Option(`${c.name} (${c.gstin || 'Unregistered'})`, c.id);
        opt.dataset.supplier = JSON.stringify(c);
        suppSelect.add(opt);
    });

    if (purchase) {
        document.getElementById('purchaseNumber').value = purchase.purchaseNumber || '';
        document.getElementById('purchaseDate').value = purchase.date ? purchase.date.split('T')[0] : new Date().toISOString().split('T')[0];
        document.getElementById('purchaseSupplier').value = purchase.supplierId || '';
        document.getElementById('supplierBillNo').value = purchase.supplierBillNo || '';
        document.getElementById('purchasePlaceOfSupply').value = purchase.placeOfSupply || '';
        document.getElementById('purchaseDiscount').value = purchase.discount || 0;
        document.getElementById('purchaseRoundOff').value = purchase.roundOff || 0;
        document.getElementById('purchaseNotes').value = purchase.notes || '';

        const items = purchaseItems.filter(i => i.purchaseId === purchase.id);
        items.forEach(item => addPurchaseItemRow(item));
    } else {
        const nextNum = await ipcRenderer.invoke('get-next-purchase-number', settings);
        document.getElementById('purchaseNumber').value = nextNum;
        document.getElementById('purchaseDate').value = new Date().toISOString().split('T')[0];
        addPurchaseItemRow();
    }

    calculatePurchaseTotals();
    document.getElementById('purchaseModal').classList.add('active');
}

function closePurchaseModal() {
    document.getElementById('purchaseModal').classList.remove('active');
    editingPurchaseId = null;
}

function onSupplierChange() {
    const sel = document.getElementById('purchaseSupplier');
    const opt = sel.options[sel.selectedIndex];
    if (opt && opt.dataset.supplier) {
        const supp = JSON.parse(opt.dataset.supplier);
        if (supp.state) {
            document.getElementById('purchasePlaceOfSupply').value = supp.state;
            calculatePurchaseTotals();
        }
    }
}

function addPurchaseItemRow(item = null) {
    purchaseItemCounter++;
    const tbody = document.querySelector('#purchaseItemsTable tbody');
    const tr = document.createElement('tr');
    tr.dataset.itemId = item?.id || `new_${purchaseItemCounter}`;

    const qty = item?.quantity || 1;
    const rate = item?.rate || 0;
    const disc = item?.discountPercent || 0;
    const defaultSearchVal = item?.productName ? item.productName : '';

    tr.innerHTML = `
        <td>${tbody.children.length + 1}</td>
        <td>
            <div class="combobox-container pitem-combobox">
                <div class="combobox-input-wrap">
                    <input type="text" class="pitem-search-input" placeholder="🔍 Search item / service..." autocomplete="off" value="${escapeHtml(defaultSearchVal)}">
                    <button type="button" class="combobox-arrow-btn" tabindex="-1">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
                    </button>
                </div>
                <input type="hidden" class="pitem-product" value="${item?.productId || ''}">
                <div class="combobox-menu" style="display: none;"></div>
            </div>
        </td>
        <td><input type="text" class="pitem-hsn" value="${item?.hsn || ''}" placeholder="HSN/SAC"></td>
        <td><input type="number" class="pitem-qty" value="${qty}" min="0.01" step="any" oninput="calculatePurchaseItemRow(this)"></td>
        <td><input type="text" class="pitem-unit" value="${item?.unit || 'PCS'}" style="width: 60px;"></td>
        <td><input type="number" class="pitem-rate" value="${rate}" min="0" step="0.01" oninput="calculatePurchaseItemRow(this)"></td>
        <td><input type="number" class="pitem-disc" value="${disc}" min="0" max="100" step="0.01" oninput="calculatePurchaseItemRow(this)"></td>
        <td><input type="text" class="pitem-taxable" value="0.00" readonly></td>
        <td><input type="number" class="pitem-gstrate" value="${item?.gstRate || 18}" min="0" max="100" step="0.01" oninput="calculatePurchaseItemRow(this)"></td>
        <td><input type="text" class="pitem-cgst" value="0.00" readonly></td>
        <td><input type="text" class="pitem-sgst" value="0.00" readonly></td>
        <td><input type="text" class="pitem-igst" value="0.00" readonly></td>
        <td><input type="text" class="pitem-total" value="0.00" readonly></td>
        <td><button type="button" class="btn-icon delete" onclick="removePurchaseItem(this)"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></td>
    `;

    tbody.appendChild(tr);

    setupProductCombobox(tr.querySelector('.pitem-combobox'), 'purchase', onPurchaseProductSelect);

    calculatePurchaseItemRow(tr.querySelector('.pitem-qty'));
    updatePurchaseItemNumbers();
}

function addPurchaseItem() {
    addPurchaseItemRow();
}

function removePurchaseItem(btn) {
    btn.closest('tr').remove();
    updatePurchaseItemNumbers();
    calculatePurchaseTotals();
}

function updatePurchaseItemNumbers() {
    document.querySelectorAll('#purchaseItemsTable tbody tr').forEach((tr, idx) => {
        tr.cells[0].textContent = idx + 1;
    });
}

function onPurchaseProductSelect(prodOrElem, trContext) {
    let prod = prodOrElem;
    let tr = trContext;
    if (prodOrElem && prodOrElem.tagName) {
        tr = prodOrElem.closest('tr');
        if (prodOrElem.options) {
            const opt = prodOrElem.options[prodOrElem.selectedIndex];
            if (opt && opt.dataset.product) prod = JSON.parse(opt.dataset.product);
        }
    }
    if (!prod || !tr) return;
    tr.querySelector('.pitem-hsn').value = prod.hsn || '';
    tr.querySelector('.pitem-unit').value = prod.unit || (prod.type === 'Service' ? 'OTH' : 'PCS');
    tr.querySelector('.pitem-rate').value = Number(prod.purchasePrice || prod.salePrice || 0);
    tr.querySelector('.pitem-gstrate').value = Number(prod.gstRate || 18);
    calculatePurchaseItemRow(tr.querySelector('.pitem-qty'));
}

function calculatePurchaseItemRow(input) {
    const tr = input.closest('tr');
    const qty = parseFloat(tr.querySelector('.pitem-qty').value) || 0;
    const rate = parseFloat(tr.querySelector('.pitem-rate').value) || 0;
    const discPercent = parseFloat(tr.querySelector('.pitem-disc').value) || 0;
    const gstRate = parseFloat(tr.querySelector('.pitem-gstrate').value) || 0;

    const discountAmt = (qty * rate * discPercent) / 100;
    const taxable = Math.max(0, qty * rate - discountAmt);

    const supplyState = document.getElementById('purchasePlaceOfSupply').value || '';
    const companyState = settings.companyState || '';
    const isInterstate = supplyState && companyState && supplyState.toLowerCase() !== companyState.toLowerCase();

    let cgstAmt = 0;
    let sgstAmt = 0;
    let igstAmt = 0;

    if (isInterstate) {
        igstAmt = (taxable * gstRate) / 100;
    } else {
        cgstAmt = (taxable * (gstRate / 2)) / 100;
        sgstAmt = (taxable * (gstRate / 2)) / 100;
    }
    const total = taxable + cgstAmt + sgstAmt + igstAmt;

    tr.querySelector('.pitem-taxable').value = taxable.toFixed(2);
    tr.querySelector('.pitem-cgst').value = cgstAmt.toFixed(2);
    tr.querySelector('.pitem-sgst').value = sgstAmt.toFixed(2);
    tr.querySelector('.pitem-igst').value = igstAmt.toFixed(2);
    tr.querySelector('.pitem-total').value = total.toFixed(2);

    calculatePurchaseTotals();
}

function calculatePurchaseTotals() {
    let subtotal = 0;
    let totalCgst = 0;
    let totalSgst = 0;
    let totalIgst = 0;

    document.querySelectorAll('#purchaseItemsTable tbody tr').forEach(tr => {
        subtotal += parseFloat(tr.querySelector('.pitem-taxable').value) || 0;
        totalCgst += parseFloat(tr.querySelector('.pitem-cgst').value) || 0;
        totalSgst += parseFloat(tr.querySelector('.pitem-sgst').value) || 0;
        totalIgst += parseFloat(tr.querySelector('.pitem-igst').value) || 0;
    });

    const discount = parseFloat(document.getElementById('purchaseDiscount').value) || 0;
    const roundOff = parseFloat(document.getElementById('purchaseRoundOff').value) || 0;
    const totalTax = totalCgst + totalSgst + totalIgst;
    const grandTotal = Math.max(0, subtotal + totalTax - discount + roundOff);

    document.getElementById('purchaseSummarySubtotal').textContent = '₹' + formatNumber(subtotal);
    document.getElementById('purchaseSummaryCgst').textContent = '₹' + formatNumber(totalCgst);
    document.getElementById('purchaseSummarySgst').textContent = '₹' + formatNumber(totalSgst);
    document.getElementById('purchaseSummaryIgst').textContent = '₹' + formatNumber(totalIgst);
    document.getElementById('purchaseSummaryTotalTax').textContent = '₹' + formatNumber(totalTax);
    document.getElementById('purchaseSummaryDiscount').textContent = '₹' + formatNumber(discount);
    document.getElementById('purchaseSummaryRoundOff').textContent = '₹' + formatNumber(roundOff);
    document.getElementById('purchaseSummaryGrandTotal').textContent = '₹' + formatNumber(grandTotal);
}

async function savePurchase() {
    const purchaseNumber = document.getElementById('purchaseNumber').value;
    const date = document.getElementById('purchaseDate').value;
    const supplierId = document.getElementById('purchaseSupplier').value;
    const supplierBillNo = document.getElementById('supplierBillNo').value;
    const placeOfSupply = document.getElementById('purchasePlaceOfSupply').value;
    const discount = parseFloat(document.getElementById('purchaseDiscount').value) || 0;
    const roundOff = parseFloat(document.getElementById('purchaseRoundOff').value) || 0;
    const notes = document.getElementById('purchaseNotes').value;

    if (!purchaseNumber || !date || !supplierId || !placeOfSupply) {
        showToast('Please fill all required fields', 'error');
        return;
    }

    const supplier = clients.find(c => c.id === supplierId);
    if (!supplier) {
        showToast('Please select a valid supplier', 'error');
        return;
    }

    const items = [];
    document.querySelectorAll('#purchaseItemsTable tbody tr').forEach(tr => {
        const prodSelect = tr.querySelector('.pitem-product');
        const prodId = prodSelect ? prodSelect.value : '';
        let pData = products.find(p => p.id === prodId) || null;
        const searchInput = tr.querySelector('.pitem-search-input');
        const rawSearchText = searchInput ? searchInput.value.replace(/^\[(Goods|Service)\]\s*/i, '').trim() : '';

        const qty = parseFloat(tr.querySelector('.pitem-qty').value) || 0;
        const rate = parseFloat(tr.querySelector('.pitem-rate').value) || 0;
        const disc = parseFloat(tr.querySelector('.pitem-disc').value) || 0;
        const taxable = parseFloat(tr.querySelector('.pitem-taxable').value) || 0;
        const gstRate = parseFloat(tr.querySelector('.pitem-gstrate').value) || 0;
        const cgstAmt = parseFloat(tr.querySelector('.pitem-cgst').value) || 0;
        const sgstAmt = parseFloat(tr.querySelector('.pitem-sgst').value) || 0;
        const igstAmt = parseFloat(tr.querySelector('.pitem-igst').value) || 0;
        const total = parseFloat(tr.querySelector('.pitem-total').value) || 0;

        if (taxable > 0 || qty > 0) {
            items.push({
                productId: prodId || '',
                productName: pData ? pData.name : (rawSearchText || 'Item'),
                type: pData ? (pData.type || 'Goods') : (searchInput && searchInput.value.includes('[Service]') ? 'Service' : 'Goods'),
                hsn: tr.querySelector('.pitem-hsn').value || '',
                quantity: qty,
                unit: tr.querySelector('.pitem-unit').value || 'PCS',
                rate,
                discountPercent: disc,
                discountAmount: (qty * rate * disc) / 100,
                taxableValue: taxable,
                gstRate,
                cgstPercent: cgstAmt > 0 ? gstRate / 2 : 0,
                cgstAmount: cgstAmt,
                sgstPercent: sgstAmt > 0 ? gstRate / 2 : 0,
                sgstAmount: sgstAmt,
                igstPercent: igstAmt > 0 ? gstRate : 0,
                igstAmount: igstAmt,
                totalAmount: total
            });
        }
    });

    if (items.length === 0) {
        showToast('Please add at least one line item', 'error');
        return;
    }

    const subtotal = items.reduce((sum, i) => sum + i.taxableValue, 0);
    const totalCgst = items.reduce((sum, i) => sum + i.cgstAmount, 0);
    const totalSgst = items.reduce((sum, i) => sum + i.sgstAmount, 0);
    const totalIgst = items.reduce((sum, i) => sum + i.igstAmount, 0);
    const totalTax = totalCgst + totalSgst + totalIgst;
    const grandTotal = Math.max(0, subtotal + totalTax - discount + roundOff);

    const purchaseData = {
        purchaseNumber,
        date,
        supplierId,
        supplierName: supplier.name,
        supplierGstin: supplier.gstin || '',
        supplierBillNo,
        placeOfSupply,
        subtotal,
        totalCgst,
        totalSgst,
        totalIgst,
        totalTax,
        discount,
        roundOff,
        grandTotal,
        status: 'Received',
        notes
    };

    try {
        if (editingPurchaseId) {
            await ipcRenderer.invoke('update-purchase', { purchase: { ...purchaseData, id: editingPurchaseId }, items });
            showToast('Purchase bill updated successfully', 'success');
        } else {
            await ipcRenderer.invoke('save-purchase', { purchase: purchaseData, items });
            showToast('Purchase bill recorded successfully', 'success');
        }
        await loadAllData();
        renderAllTables();
        closePurchaseModal();
    } catch (err) {
        showToast('Error saving purchase: ' + err.message, 'error');
    }
}

async function editPurchase(id) {
    const p = purchases.find(item => item.id === id);
    if (p) openPurchaseModal(p);
}

async function deletePurchase(id) {
    if (!confirm('Are you sure you want to delete this purchase bill? Stock and accounts will be updated.')) return;
    try {
        await ipcRenderer.invoke('delete-purchase', id);
        await loadAllData();
        renderAllTables();
        showToast('Purchase bill deleted', 'success');
    } catch (err) {
        showToast('Error deleting purchase: ' + err.message, 'error');
    }
}

// ==========================================================
// 3. PRODUCTS & SERVICES MASTER (GOODS VS SERVICES)
// ==========================================================
function renderProductsTable() {
    const tbody = document.querySelector('#productsTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let list = products;
    if (currentProductFilter === 'Goods') {
        list = products.filter(p => (p.type || 'Goods') === 'Goods');
    } else if (currentProductFilter === 'Service') {
        list = products.filter(p => (p.type || '').toLowerCase().includes('service'));
    }

    list.forEach(p => {
        const isServ = (p.type || '').toLowerCase().includes('service');
        const typeBadge = isServ ? '<span class="type-badge service">Service</span>' : '<span class="type-badge goods">Goods</span>';
        const openStockDisplay = isServ ? 'N/A' : (p.openingStock || 0);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${typeBadge}</td>
            <td><strong>${p.name || '-'}</strong></td>
            <td>${p.hsn || '-'}</td>
            <td>${p.unit || '-'}</td>
            <td>${p.gstRate || 0}%</td>
            <td>₹${formatNumber(p.salePrice)}</td>
            <td>₹${formatNumber(p.purchasePrice || 0)}</td>
            <td>${openStockDisplay}</td>
            <td>${p.description || '-'}</td>
            <td>
                <button class="btn-icon" onclick="editProduct('${p.id}')" title="Edit / Alter">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
                <button class="btn-icon delete" onclick="deleteProduct('${p.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function handleProductTypeToggle(type) {
    const isService = type === 'Service';
    document.getElementById('typeGoodsLabel').classList.toggle('active', !isService);
    document.getElementById('typeServiceLabel').classList.toggle('active', isService);

    document.getElementById('productNameLabel').textContent = isService ? 'Service Name *' : 'Item Name *';
    document.getElementById('productHsnLabel').textContent = isService ? 'SAC Code *' : 'HSN Code *';
    document.getElementById('productHsn').placeholder = isService ? 'e.g. 9983' : 'e.g. 8471';

    const stockGroup = document.getElementById('openingStockGroup');
    if (isService) {
        stockGroup.style.display = 'none';
        document.getElementById('productUnit').value = 'OTH';
    } else {
        stockGroup.style.display = 'block';
        if (document.getElementById('productUnit').value === 'OTH') {
            document.getElementById('productUnit').value = 'PCS';
        }
    }
}

function openProductModal(product = null) {
    editingProductId = product?.id || null;
    document.getElementById('productModalTitle').textContent = product ? 'Edit / Alter Item' : 'Add Item / Service';
    document.getElementById('productForm').reset();
    document.getElementById('productId').value = product?.id || '';

    const pType = product?.type || 'Goods';
    const typeRadios = document.querySelectorAll('input[name="type"]');
    typeRadios.forEach(r => {
        r.checked = (r.value === pType);
    });
    handleProductTypeToggle(pType);

    if (product) {
        document.getElementById('productName').value = product.name || '';
        document.getElementById('productHsn').value = product.hsn || '';
        document.getElementById('productUnit').value = product.unit || (pType === 'Service' ? 'OTH' : 'PCS');
        document.getElementById('productGstRate').value = product.gstRate || 0;
        document.getElementById('productSalePrice').value = product.salePrice || 0;
        document.getElementById('productPurchasePrice').value = product.purchasePrice || 0;
        document.getElementById('productOpeningStock').value = product.openingStock || 0;
        document.getElementById('productDescription').value = product.description || '';
    } else {
        document.getElementById('productUnit').value = 'PCS';
        document.getElementById('productGstRate').value = 18;
        document.getElementById('productOpeningStock').value = 0;
    }

    document.getElementById('productModal').classList.add('active');
    setTimeout(() => {
        const nameInput = document.getElementById('productName');
        if (nameInput) {
            nameInput.focus();
            nameInput.select();
        }
    }, 60);
}

function closeProductModal() {
    document.getElementById('productModal').classList.remove('active');
    editingProductId = null;
}

async function saveProduct(e) {
    e.preventDefault();
    const form = e.target;
    const formData = new FormData(form);
    const product = Object.fromEntries(formData);

    product.type = product.type || 'Goods';
    product.gstRate = parseFloat(product.gstRate) || 0;
    product.salePrice = parseFloat(product.salePrice) || 0;
    product.purchasePrice = parseFloat(product.purchasePrice) || 0;
    product.openingStock = product.type === 'Service' ? 0 : (parseFloat(product.openingStock) || 0);
    product.cgst = product.gstRate / 2;
    product.sgst = product.gstRate / 2;
    product.igst = product.gstRate;

    try {
        if (editingProductId) {
            await ipcRenderer.invoke('update-product', { ...product, id: editingProductId });
            showToast('Item updated successfully', 'success');
        } else {
            await ipcRenderer.invoke('add-product', product);
            showToast('Item added successfully', 'success');
        }
        await loadAllData();
        renderAllTables();
        closeProductModal();
    } catch (err) {
        showToast('Error saving item: ' + err.message, 'error');
    }
}

async function editProduct(id) {
    const p = products.find(prod => prod.id === id);
    if (p) openProductModal(p);
}

async function deleteProduct(id) {
    if (!confirm('Are you sure you want to delete this item?')) return;
    try {
        await ipcRenderer.invoke('delete-product', id);
        await loadAllData();
        renderAllTables();
        showToast('Item deleted', 'success');
    } catch (err) {
        showToast('Error deleting item: ' + err.message, 'error');
    }
}

// ==========================================================
// 4. STOCK REGISTER (INVENTORY VALUATION)
// ==========================================================
function renderStockRegisterTable() {
    const tbody = document.querySelector('#stockRegisterTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const goodsItems = stockRegister.filter(i => i.type === 'Goods');
    const inStock = goodsItems.filter(i => i.closingStock > 0);
    const lowStock = goodsItems.filter(i => i.closingStock <= 0 || i.closingStock < 5);
    const totalVal = goodsItems.reduce((sum, i) => sum + (Number(i.closingValue) || 0), 0);

    document.getElementById('stockTotalItems').textContent = goodsItems.length;
    document.getElementById('stockInStockItems').textContent = inStock.length;
    document.getElementById('stockLowStockItems').textContent = lowStock.length;
    document.getElementById('stockTotalValue').textContent = '₹' + formatNumber(totalVal);

    stockRegister.forEach(item => {
        let badgeClass = 'in-stock';
        if (item.type === 'Service') {
            badgeClass = 'service';
        } else if (item.closingStock <= 0) {
            badgeClass = 'out-of-stock';
        } else if (item.closingStock < 5) {
            badgeClass = 'low-stock';
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${item.name}</strong> ${item.type === 'Service' ? '<span class="type-badge service">Service</span>' : ''}</td>
            <td>${item.hsn || '-'}</td>
            <td>${item.unit || '-'}</td>
            <td>${item.type === 'Service' ? 'N/A' : item.openingStock}</td>
            <td>${item.type === 'Service' ? 'N/A' : item.inwardQty}</td>
            <td>${item.type === 'Service' ? 'N/A' : item.outwardQty}</td>
            <td><strong>${item.type === 'Service' ? 'N/A' : item.closingStock}</strong></td>
            <td>₹${formatNumber(item.purchaseRate)}</td>
            <td><strong>₹${formatNumber(item.closingValue)}</strong></td>
            <td><span class="stock-status ${badgeClass}">${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

async function exportStockToExcel() {
    exportExcel('Products');
}

// ==========================================================
// 5. PARTIES (CUSTOMERS & SUPPLIERS)
// ==========================================================
function renderClientsTable() {
    const tbody = document.querySelector('#clientsTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let list = clients;
    if (currentPartyFilter === 'Customer') {
        list = clients.filter(c => c.partyType !== 'Supplier');
    } else if (currentPartyFilter === 'Supplier') {
        list = clients.filter(c => c.partyType !== 'Customer');
    }

    list.forEach(c => {
        const tr = document.createElement('tr');
        const openBal = Number(c.openingBalance) || 0;
        const balBadge = openBal > 0 ? `₹${formatNumber(openBal)} ${c.balanceType || 'Dr'}` : '₹0.00';

        tr.innerHTML = `
            <td><strong>${c.name || '-'}</strong></td>
            <td><span class="badge ${c.partyType === 'Supplier' ? 'orange' : 'blue'}">${c.partyType || 'Customer'}</span></td>
            <td>${c.gstin || '-'}</td>
            <td>${c.pan || '-'}</td>
            <td>${c.city || '-'}</td>
            <td>${c.state || '-'}</td>
            <td>${c.phone || '-'}</td>
            <td>${balBadge}</td>
            <td>₹${formatNumber(c.creditLimit || 0)}</td>
            <td>
                <button class="btn-icon" onclick="editClient('${c.id}')" title="Edit / Alter">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
                <button class="btn-icon delete" onclick="deleteClient('${c.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function openClientModal(client = null) {
    editingClientId = client?.id || null;
    document.getElementById('clientModalTitle').textContent = client ? 'Edit / Alter Party' : 'Add Party (Customer / Supplier)';
    document.getElementById('clientForm').reset();
    document.getElementById('clientId').value = client?.id || '';

    if (client) {
        document.getElementById('clientName').value = client.name || '';
        document.getElementById('clientPartyType').value = client.partyType || 'Customer';
        document.getElementById('clientOpeningBalance').value = client.openingBalance || 0;
        document.getElementById('clientBalanceType').value = client.balanceType || (client.partyType === 'Supplier' ? 'Cr' : 'Dr');
        document.getElementById('clientGstin').value = client.gstin || '';
        document.getElementById('clientPan').value = client.pan || '';
        document.getElementById('clientAddress1').value = client.address1 || '';
        document.getElementById('clientAddress2').value = client.address2 || '';
        document.getElementById('clientCity').value = client.city || '';
        document.getElementById('clientState').value = client.state || '';
        document.getElementById('clientStateCode').value = client.stateCode || '';
        document.getElementById('clientPincode').value = client.pincode || '';
        document.getElementById('clientPhone').value = client.phone || '';
        document.getElementById('clientEmail').value = client.email || '';
        document.getElementById('clientContactPerson').value = client.contactPerson || '';
        document.getElementById('clientCreditLimit').value = client.creditLimit || 0;
    } else {
        document.getElementById('clientPartyType').value = 'Customer';
        document.getElementById('clientBalanceType').value = 'Dr';
        document.getElementById('clientOpeningBalance').value = 0;
        document.getElementById('clientCreditLimit').value = 0;
    }

    document.getElementById('clientModal').classList.add('active');
    setTimeout(() => {
        const nameInput = document.getElementById('clientName');
        if (nameInput) {
            nameInput.focus();
            nameInput.select();
        }
    }, 60);
}

function closeClientModal() {
    document.getElementById('clientModal').classList.remove('active');
    editingClientId = null;
}

async function saveClient(e) {
    e.preventDefault();
    const form = e.target;
    const formData = new FormData(form);
    const client = Object.fromEntries(formData);

    client.partyType = client.partyType || 'Customer';
    client.openingBalance = parseFloat(client.openingBalance) || 0;
    client.creditLimit = parseFloat(client.creditLimit) || 0;
    client.paymentTerms = parseInt(client.paymentTerms) || 30;

    try {
        if (editingClientId) {
            await ipcRenderer.invoke('update-client', { ...client, id: editingClientId });
            showToast('Party updated successfully', 'success');
        } else {
            await ipcRenderer.invoke('add-client', client);
            showToast('Party created successfully', 'success');
        }
        await loadAllData();
        renderAllTables();
        closeClientModal();
    } catch (err) {
        showToast('Error saving party: ' + err.message, 'error');
    }
}

async function editClient(id) {
    const c = clients.find(client => client.id === id);
    if (c) openClientModal(c);
}

async function deleteClient(id) {
    if (!confirm('Are you sure you want to delete this party?')) return;
    try {
        await ipcRenderer.invoke('delete-client', id);
        await loadAllData();
        renderAllTables();
        showToast('Party deleted', 'success');
    } catch (err) {
        showToast('Error deleting party: ' + err.message, 'error');
    }
}

// ==========================================================
// 6. VOUCHERS (RECEIPT, PAYMENT, CONTRA, JOURNAL)
// ==========================================================
function renderVouchersTable() {
    const tbody = document.querySelector('#vouchersTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let list = vouchers;
    if (currentVoucherFilter !== 'all') {
        list = vouchers.filter(v => (v.type || '').toLowerCase() === currentVoucherFilter.toLowerCase());
    }

    list.forEach(v => {
        const typeClass = (v.type || 'payment').toLowerCase();
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${v.voucherNumber || '-'}</strong></td>
            <td>${formatDate(v.date)}</td>
            <td><span class="voucher-badge ${typeClass}">${v.type || 'Payment'}</span></td>
            <td>${v.mode || 'N/A'}</td>
            <td>${v.debitLedger || '-'}</td>
            <td>${v.creditLedger || '-'}</td>
            <td><strong>₹${formatNumber(v.amount)}</strong></td>
            <td>${v.referenceNo || '-'}</td>
            <td>${v.narration || '-'}</td>
            <td>
                <button class="btn-icon delete" onclick="deleteVoucher('${v.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

async function openVoucherModal() {
    editingVoucherId = null;
    document.getElementById('voucherForm').reset();
    document.getElementById('voucherDate').value = new Date().toISOString().split('T')[0];
    document.getElementById('voucherType').value = 'Payment';

    populateVoucherLedgerDropdowns();
    onVoucherTypeChange();

    document.getElementById('voucherModal').classList.add('active');
    setTimeout(() => {
        const amtInput = document.getElementById('voucherAmount');
        if (amtInput) {
            amtInput.focus();
            amtInput.select();
        }
    }, 60);
}

function closeVoucherModal() {
    document.getElementById('voucherModal').classList.remove('active');
    editingVoucherId = null;
}

function populateVoucherLedgerDropdowns() {
    const partyLedgerSelect = document.getElementById('voucherPartyLedger');
    const journalDr = document.getElementById('journalDebit');
    const journalCr = document.getElementById('journalCredit');

    // Combine all ledgers + parties
    const allLedgerNames = new Set();
    ledgers.forEach(l => allLedgerNames.add(l.name));
    clients.forEach(c => allLedgerNames.add(c.name));

    const sortedNames = Array.from(allLedgerNames).sort((a, b) => a.localeCompare(b));

    const buildOptions = () => sortedNames.map(name => `<option value="${name}">${name}</option>`).join('');

    if (partyLedgerSelect) {
        partyLedgerSelect.innerHTML = '<option value="">Select Ledger / Party</option>' + buildOptions();
    }
    if (journalDr) {
        journalDr.innerHTML = '<option value="">Select Debit Account</option>' + buildOptions();
    }
    if (journalCr) {
        journalCr.innerHTML = '<option value="">Select Credit Account</option>' + buildOptions();
    }
}

async function onVoucherTypeChange() {
    const type = document.getElementById('voucherType').value;
    const modeRow = document.getElementById('voucherModeRow');
    const contraRow = document.getElementById('contraRow');
    const journalRow = document.getElementById('journalRow');

    const nextNum = await ipcRenderer.invoke('get-next-voucher-number', type);
    document.getElementById('voucherNumber').value = nextNum;

    if (type === 'Contra') {
        modeRow.style.display = 'none';
        contraRow.style.display = 'grid';
        journalRow.style.display = 'none';
        document.getElementById('voucherPartyLedger').required = false;
    } else if (type === 'Journal') {
        modeRow.style.display = 'none';
        contraRow.style.display = 'none';
        journalRow.style.display = 'grid';
        document.getElementById('voucherPartyLedger').required = false;
    } else {
        // Payment or Receipt
        modeRow.style.display = 'grid';
        contraRow.style.display = 'none';
        journalRow.style.display = 'none';
        document.getElementById('voucherPartyLedger').required = true;

        if (type === 'Payment') {
            document.getElementById('voucherModeLabel').textContent = 'Paid Via (Cash / Bank) *';
            document.getElementById('voucherLedgerLabel').textContent = 'Paid To (Party / Expense Ledger) *';
        } else {
            document.getElementById('voucherModeLabel').textContent = 'Received In (Cash / Bank) *';
            document.getElementById('voucherLedgerLabel').textContent = 'Received From (Customer / Income) *';
        }
    }
}

async function saveVoucher(e) {
    e.preventDefault();
    const type = document.getElementById('voucherType').value;
    const voucherNumber = document.getElementById('voucherNumber').value;
    const date = document.getElementById('voucherDate').value;
    const amount = parseFloat(document.getElementById('voucherAmount').value) || 0;
    const referenceNo = document.getElementById('voucherRefNo').value || '';
    const narration = document.getElementById('voucherNarration').value || '';

    if (!voucherNumber || !date || amount <= 0) {
        showToast('Please enter a valid date and amount', 'error');
        return;
    }

    let mode = '';
    let debitLedger = '';
    let creditLedger = '';

    if (type === 'Payment') {
        mode = document.getElementById('voucherMode').value;
        const target = document.getElementById('voucherPartyLedger').value;
        if (!target) { showToast('Please select the party or ledger', 'error'); return; }
        debitLedger = target;
        creditLedger = mode === 'Cash' ? 'Cash' : 'Bank Account';
    } else if (type === 'Receipt') {
        mode = document.getElementById('voucherMode').value;
        const target = document.getElementById('voucherPartyLedger').value;
        if (!target) { showToast('Please select the party or ledger', 'error'); return; }
        debitLedger = mode === 'Cash' ? 'Cash' : 'Bank Account';
        creditLedger = target;
    } else if (type === 'Contra') {
        debitLedger = document.getElementById('contraDebit').value;
        creditLedger = document.getElementById('contraCredit').value;
        mode = debitLedger === 'Bank Account' ? 'Bank' : 'Cash';
        if (debitLedger === creditLedger) {
            showToast('Contra requires different accounts for debit and credit', 'error');
            return;
        }
    } else if (type === 'Journal') {
        debitLedger = document.getElementById('journalDebit').value;
        creditLedger = document.getElementById('journalCredit').value;
        mode = 'Journal';
        if (!debitLedger || !creditLedger) {
            showToast('Please select both debit and credit ledgers', 'error');
            return;
        }
    }

    const voucherData = {
        voucherNumber,
        date,
        type,
        mode,
        debitLedger,
        creditLedger,
        amount,
        referenceNo,
        narration
    };

    try {
        await ipcRenderer.invoke('save-voucher', voucherData);
        showToast(`${type} voucher recorded successfully`, 'success');
        await loadAllData();
        renderAllTables();
        closeVoucherModal();
    } catch (err) {
        showToast('Error saving voucher: ' + err.message, 'error');
    }
}

async function deleteVoucher(id) {
    if (!confirm('Are you sure you want to delete this voucher? Accounts and bank balances will be recalculated.')) return;
    try {
        await ipcRenderer.invoke('delete-voucher', id);
        await loadAllData();
        renderAllTables();
        showToast('Voucher deleted', 'success');
    } catch (err) {
        showToast('Error deleting voucher: ' + err.message, 'error');
    }
}

// ==========================================================
// 7. CASH BOOK & BANK BOOK
// ==========================================================
function renderCashBankRegister() {
    const tbody = document.querySelector('#cashBankTable tbody');
    if (!tbody || !financialReports) return;
    tbody.innerHTML = '';

    const isCash = currentCashBankBook === 'cash';
    const book = isCash ? financialReports.cashBook : financialReports.bankBook;
    if (!book) return;

    document.getElementById('cbOpeningBalance').textContent = '₹' + formatNumber(book.openingBalance);
    document.getElementById('cbTotalReceipts').textContent = '₹' + formatNumber(book.totalReceipts);
    document.getElementById('cbTotalPayments').textContent = '₹' + formatNumber(book.totalPayments);
    document.getElementById('cbClosingBalance').textContent = '₹' + formatNumber(book.closingBalance);

    book.transactions.forEach(tx => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${formatDate(tx.date)}</td>
            <td><strong>${tx.voucherNumber || '-'}</strong></td>
            <td><span class="voucher-badge ${tx.type.toLowerCase()}">${tx.type}</span></td>
            <td>${tx.particulars}</td>
            <td style="color: #15803d; font-weight: 600;">${tx.debit > 0 ? `₹${formatNumber(tx.debit)}` : '-'}</td>
            <td style="color: #b91c1c; font-weight: 600;">${tx.credit > 0 ? `₹${formatNumber(tx.credit)}` : '-'}</td>
            <td><strong>₹${formatNumber(tx.balance)}</strong></td>
            <td>${tx.narration || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

// ==========================================================
// 8. ACCOUNTS & LEDGERS (STANDARD 28 CORE GROUPS)
// ==========================================================
function renderLedgersTable() {
    const tbody = document.querySelector('#ledgersTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    ledgers.forEach(l => {
        const grp = ledgerGroups.find(g => g.name === l.groupName);
        const nature = grp ? grp.nature : 'General';
        const isDefault = l.isDefault === 'Yes';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${l.name}</strong></td>
            <td>${l.groupName}</td>
            <td><span class="badge ${nature === 'Assets' ? 'blue' : nature === 'Liabilities' ? 'purple' : nature === 'Income' ? 'green' : 'orange'}">${nature}</span></td>
            <td>₹${formatNumber(l.openingBalance || 0)}</td>
            <td><strong>${l.balanceType || 'Dr'}</strong></td>
            <td>${l.description || '-'}</td>
            <td>
                <button class="btn-icon" onclick="editLedger('${l.id}')" title="Edit / Alter">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
                ${!isDefault ? `
                <button class="btn-icon delete" onclick="deleteLedger('${l.id}')" title="Delete">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                </button>` : ''}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderGroupsTable() {
    const tbody = document.querySelector('#groupsTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    ledgerGroups.forEach(g => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${g.name}</strong></td>
            <td>${g.parentGroup ? `↳ ${g.parentGroup}` : '<em style="color:#64748b;">Primary Group</em>'}</td>
            <td><span class="badge ${g.nature === 'Assets' ? 'blue' : g.nature === 'Liabilities' ? 'purple' : g.nature === 'Income' ? 'green' : 'orange'}">${g.nature}</span></td>
            <td>${g.isDefault === 'Yes' ? 'Standard Core' : 'User Defined'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function openLedgerModal(ledger = null) {
    editingLedgerId = ledger?.id || null;
    document.getElementById('ledgerModalTitle').textContent = ledger ? 'Alter Ledger Master' : 'Create Ledger Master';
    document.getElementById('ledgerForm').reset();
    document.getElementById('ledgerId').value = ledger?.id || '';

    populateLedgerGroupDropdown();

    if (ledger) {
        document.getElementById('ledgerName').value = ledger.name || '';
        document.getElementById('ledgerGroupName').value = ledger.groupName || '';
        document.getElementById('ledgerOpeningBalance').value = ledger.openingBalance || 0;
        document.getElementById('ledgerBalanceType').value = ledger.balanceType || 'Dr';
        document.getElementById('ledgerDescription').value = ledger.description || '';
    } else {
        document.getElementById('ledgerBalanceType').value = 'Dr';
        document.getElementById('ledgerOpeningBalance').value = 0;
    }

    document.getElementById('ledgerModal').classList.add('active');
    setTimeout(() => {
        const nameInput = document.getElementById('ledgerName');
        if (nameInput) {
            nameInput.focus();
            nameInput.select();
        }
    }, 60);
}

function closeLedgerModal() {
    document.getElementById('ledgerModal').classList.remove('active');
    editingLedgerId = null;
}

async function saveLedger(e) {
    e.preventDefault();
    const form = e.target;
    const formData = new FormData(form);
    const ledger = Object.fromEntries(formData);

    ledger.openingBalance = parseFloat(ledger.openingBalance) || 0;
    ledger.balanceType = ledger.balanceType || 'Dr';

    try {
        if (editingLedgerId) {
            await ipcRenderer.invoke('update-ledger', { ...ledger, id: editingLedgerId });
            showToast('Ledger updated successfully', 'success');
        } else {
            await ipcRenderer.invoke('add-ledger', ledger);
            showToast('Ledger created successfully', 'success');
        }
        await loadAllData();
        renderAllTables();
        closeLedgerModal();
    } catch (err) {
        showToast('Error saving ledger: ' + err.message, 'error');
    }
}

async function editLedger(id) {
    const l = ledgers.find(item => item.id === id);
    if (l) openLedgerModal(l);
}

async function deleteLedger(id) {
    if (!confirm('Are you sure you want to delete this ledger?')) return;
    try {
        await ipcRenderer.invoke('delete-ledger', id);
        await loadAllData();
        renderAllTables();
        showToast('Ledger deleted', 'success');
    } catch (err) {
        showToast('Error deleting ledger: ' + err.message, 'error');
    }
}

// ==========================================================
// 9. FINANCIAL STATEMENTS (FINAL ACCOUNTS)
// Trial Balance, Trading Account, Profit & Loss Account, Balance Sheet
// ==========================================================
function showFinancialSheet(sheetName) {
    const views = ['trialBalance', 'tradingAccount', 'profitLoss', 'balanceSheet'];
    views.forEach(v => {
        const el = document.getElementById(v + 'View');
        if (el) el.style.display = (v === sheetName) ? 'block' : 'none';
    });
}

function renderFinancialStatements() {
    if (!financialReports) return;

    // 1. Trial Balance
    const tb = financialReports.trialBalance;
    const tbTbody = document.querySelector('#trialBalanceTable tbody');
    if (tbTbody && tb) {
        tbTbody.innerHTML = '';
        tb.rows.forEach(r => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${r.ledgerName}</strong></td>
                <td>${r.groupName}</td>
                <td><span class="badge ${r.nature === 'Assets' ? 'blue' : r.nature === 'Liabilities' ? 'purple' : r.nature === 'Income' ? 'green' : 'orange'}">${r.nature}</span></td>
                <td style="text-align: right;">${r.debit > 0 ? `₹${formatNumber(r.debit)}` : '-'}</td>
                <td style="text-align: right;">${r.credit > 0 ? `₹${formatNumber(r.credit)}` : '-'}</td>
            `;
            tbTbody.appendChild(tr);
        });

        document.getElementById('tbTotalDr').textContent = '₹' + formatNumber(tb.totalDebit);
        document.getElementById('tbTotalCr').textContent = '₹' + formatNumber(tb.totalCredit);

        const badge = document.getElementById('tbStatusBadge');
        if (badge) {
            if (Math.abs(tb.totalDebit - tb.totalCredit) < 0.05) {
                badge.className = 'badge green';
                badge.textContent = 'Balanced (Dr = Cr)';
            } else {
                badge.className = 'badge red';
                badge.textContent = `Diff: ₹${formatNumber(tb.difference)}`;
            }
        }
    }

    // 2. Trading Account
    const tr = financialReports.tradingAccount;
    if (tr) {
        document.getElementById('trOpeningStock').textContent = '₹' + formatNumber(tr.openingStock);
        document.getElementById('trPurchases').textContent = '₹' + formatNumber(tr.purchases);
        document.getElementById('trSales').textContent = '₹' + formatNumber(tr.sales);
        document.getElementById('trClosingStock').textContent = '₹' + formatNumber(tr.closingStock);

        const directExpList = document.getElementById('trDirectExpensesList');
        directExpList.innerHTML = '';
        tr.directExpenses.forEach(e => {
            directExpList.innerHTML += `<div class="statement-row"><span>${e.name}</span><strong>₹${formatNumber(e.amount)}</strong></div>`;
        });

        const directIncList = document.getElementById('trDirectIncomesList');
        directIncList.innerHTML = '';
        tr.directIncomes.forEach(i => {
            directIncList.innerHTML += `<div class="statement-row"><span>${i.name}</span><strong>₹${formatNumber(i.amount)}</strong></div>`;
        });

        if (tr.grossProfit > 0) {
            document.getElementById('trGrossProfitRow').style.display = 'flex';
            document.getElementById('trGrossProfit').textContent = '₹' + formatNumber(tr.grossProfit);
            document.getElementById('trGrossLossRow').style.display = 'none';
            document.getElementById('tradingProfitBadge').className = 'badge green';
            document.getElementById('tradingProfitBadge').textContent = `Gross Profit: ₹${formatNumber(tr.grossProfit)}`;
        } else {
            document.getElementById('trGrossProfitRow').style.display = 'none';
            document.getElementById('trGrossLossRow').style.display = 'flex';
            document.getElementById('trGrossLoss').textContent = '₹' + formatNumber(tr.grossLoss);
            document.getElementById('tradingProfitBadge').className = 'badge red';
            document.getElementById('tradingProfitBadge').textContent = `Gross Loss: ₹${formatNumber(tr.grossLoss)}`;
        }

        document.getElementById('trTotalDebit').textContent = '₹' + formatNumber(tr.totalDebit);
        document.getElementById('trTotalCredit').textContent = '₹' + formatNumber(tr.totalCredit);
    }

    // 3. Profit & Loss Account
    const pl = financialReports.profitLossAccount;
    if (pl) {
        if (pl.grossProfit > 0) {
            document.getElementById('plGrossProfitRow').style.display = 'flex';
            document.getElementById('plGrossProfit').textContent = '₹' + formatNumber(pl.grossProfit);
            document.getElementById('plGrossLossRow').style.display = 'none';
        } else {
            document.getElementById('plGrossProfitRow').style.display = 'none';
            document.getElementById('plGrossLossRow').style.display = 'flex';
            document.getElementById('plGrossLoss').textContent = '₹' + formatNumber(pl.grossLoss);
        }

        const indExpList = document.getElementById('plIndirectExpensesList');
        indExpList.innerHTML = '';
        pl.indirectExpenses.forEach(e => {
            indExpList.innerHTML += `<div class="statement-row"><span>${e.name}</span><strong>₹${formatNumber(e.amount)}</strong></div>`;
        });

        const indIncList = document.getElementById('plIndirectIncomesList');
        indIncList.innerHTML = '';
        pl.indirectIncomes.forEach(i => {
            indIncList.innerHTML += `<div class="statement-row"><span>${i.name}</span><strong>₹${formatNumber(i.amount)}</strong></div>`;
        });

        if (pl.netProfit >= 0) {
            document.getElementById('plNetProfitRow').style.display = 'flex';
            document.getElementById('plNetProfit').textContent = '₹' + formatNumber(pl.netProfit);
            document.getElementById('plNetLossRow').style.display = 'none';
            document.getElementById('plNetProfitBadge').className = 'badge green';
            document.getElementById('plNetProfitBadge').textContent = `Net Profit: ₹${formatNumber(pl.netProfit)}`;
        } else {
            document.getElementById('plNetProfitRow').style.display = 'none';
            document.getElementById('plNetLossRow').style.display = 'flex';
            document.getElementById('plNetLoss').textContent = '₹' + formatNumber(pl.netLoss);
            document.getElementById('plNetProfitBadge').className = 'badge red';
            document.getElementById('plNetProfitBadge').textContent = `Net Loss: ₹${formatNumber(pl.netLoss)}`;
        }

        document.getElementById('plTotalDebit').textContent = '₹' + formatNumber(pl.totalDebit);
        document.getElementById('plTotalCredit').textContent = '₹' + formatNumber(pl.totalCredit);
    }

    // 4. Balance Sheet
    const bs = financialReports.balanceSheet;
    if (bs) {
        const liabList = document.getElementById('bsLiabilitiesList');
        liabList.innerHTML = '';
        bs.liabilities.forEach(item => {
            liabList.innerHTML += `<div class="statement-row"><span>${item.particulars} <small style="color:#64748b;">(${item.group})</small></span><strong>₹${formatNumber(item.amount)}</strong></div>`;
        });
        document.getElementById('bsTotalLiabilities').textContent = '₹' + formatNumber(bs.totalLiabilities);

        const assList = document.getElementById('bsAssetsList');
        assList.innerHTML = '';
        bs.assets.forEach(item => {
            assList.innerHTML += `<div class="statement-row"><span>${item.particulars} <small style="color:#64748b;">(${item.group})</small></span><strong>₹${formatNumber(item.amount)}</strong></div>`;
        });
        document.getElementById('bsTotalAssets').textContent = '₹' + formatNumber(bs.totalAssets);

        const bsBadge = document.getElementById('bsBalancedBadge');
        if (bsBadge) {
            if (bs.difference < 0.05) {
                bsBadge.className = 'badge green';
                bsBadge.textContent = 'Balanced (Assets = Liabilities)';
            } else {
                bsBadge.className = 'badge red';
                bsBadge.textContent = `Difference: ₹${formatNumber(bs.difference)}`;
            }
        }
    }
}

// ==========================================================
// 10. SETTINGS, DATABASE & BACKUP
// ==========================================================
function populateSettingsForm(dbPath = '') {
    if (dbPath) {
        const dbInput = document.getElementById('dbPathDisplay');
        if (dbInput) dbInput.value = dbPath;
    }
    if (!settings) return;

    // Populate all text inputs, textareas, selects and checkboxes
    Object.entries(settings).forEach(([key, value]) => {
        const input = document.querySelector(`[name="${key}"]`);
        if (input && value !== undefined && value !== null) {
            if (input.type === 'checkbox') {
                input.checked = (value !== 'false' && value !== false);
            } else {
                input.value = value;
            }
        }
    });

    // Populate logo preview and storage fields
    const rawLogo = settings.companyLogo || settings.invoiceLogo;
    const logoDataUrl = settings.companyLogoDataUrl || (rawLogo && rawLogo.startsWith('data:image') ? rawLogo : '');
    const logoImg = document.getElementById('logoPreviewImg');
    const logoPlaceholder = document.getElementById('logoPlaceholderText');
    const removeLogoBtn = document.getElementById('removeLogoBtn');
    const compLogoInput = document.getElementById('companyLogoData');
    const invLogoInput = document.getElementById('invoiceLogoData');

    if (rawLogo) {
        if (logoImg) {
            logoImg.src = logoDataUrl || rawLogo;
            logoImg.style.display = 'block';
        }
        if (logoPlaceholder) logoPlaceholder.style.display = 'none';
        if (removeLogoBtn) removeLogoBtn.style.display = 'inline-flex';
        if (compLogoInput) compLogoInput.value = rawLogo;
        if (invLogoInput) invLogoInput.value = rawLogo;
    } else {
        if (logoImg) {
            logoImg.src = '';
            logoImg.style.display = 'none';
        }
        if (logoPlaceholder) logoPlaceholder.style.display = 'block';
        if (removeLogoBtn) removeLogoBtn.style.display = 'none';
        if (compLogoInput) compLogoInput.value = '';
        if (invLogoInput) invLogoInput.value = '';
    }

    // Populate logo width slider
    const logoWidth = parseInt(settings.invoiceLogoWidth, 10) || 100;
    const widthSlider = document.getElementById('invoiceLogoWidth');
    const widthDisplay = document.getElementById('logoWidthDisplay');
    if (widthSlider) widthSlider.value = logoWidth;
    if (widthDisplay) widthDisplay.textContent = logoWidth + ' px';

    // Populate logo align buttons
    const logoAlign = settings.invoiceLogoAlign || 'left';
    document.querySelectorAll('#logoAlignGroup .pill-seg-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === logoAlign);
    });
    const logoAlignInput = document.getElementById('invoiceLogoAlign');
    if (logoAlignInput) logoAlignInput.value = logoAlign;

    // Populate header text align buttons
    const headerAlign = settings.invoiceHeaderAlign || 'left';
    document.querySelectorAll('#headerAlignGroup .pill-seg-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === headerAlign);
    });
    const headerAlignInput = document.getElementById('invoiceHeaderAlign');
    if (headerAlignInput) headerAlignInput.value = headerAlign;

    // Populate theme color & swatches
    const themeColor = settings.invoiceThemeColor || '#1a1a2e';
    const colorPicker = document.getElementById('invoiceThemeColor');
    const colorHex = document.getElementById('themeColorHex');
    if (colorPicker) colorPicker.value = themeColor;
    if (colorHex) colorHex.textContent = themeColor;
    document.querySelectorAll('#colorSwatches .color-swatch').forEach(swatch => {
        swatch.classList.toggle('active', (swatch.dataset.color || '').toLowerCase() === themeColor.toLowerCase());
    });

    // Populate font and template dropdowns
    if (settings.invoiceFont) {
        const fontSel = document.getElementById('invoiceFont');
        if (fontSel) fontSel.value = settings.invoiceFont;
    }
    if (settings.invoiceTemplate) {
        const tplSel = document.getElementById('invoiceTemplate');
        if (tplSel) tplSel.value = settings.invoiceTemplate;
    }
}

async function saveSettings() {
    const newSettings = {};

    // 1. Company Info
    document.querySelectorAll('#companySettingsForm input, #companySettingsForm textarea').forEach(input => {
        if (input.type === 'checkbox') newSettings[input.name] = input.checked ? 'true' : 'false';
        else newSettings[input.name] = input.value;
    });

    // 2. Bank Info
    document.querySelectorAll('#bankSettingsForm input').forEach(input => {
        if (input.type === 'checkbox') newSettings[input.name] = input.checked ? 'true' : 'false';
        else newSettings[input.name] = input.value;
    });

    // 3. Invoice Series
    document.querySelectorAll('#invoiceSettingsForm input, #invoiceSettingsForm textarea').forEach(input => {
        if (input.type === 'checkbox') newSettings[input.name] = input.checked ? 'true' : 'false';
        else newSettings[input.name] = input.value;
    });

    // 4. Invoice Design & Output Settings
    document.querySelectorAll('#invoiceFormatSettingsForm input, #invoiceFormatSettingsForm select, #invoiceFormatSettingsForm textarea').forEach(input => {
        if (input.type === 'checkbox') {
            newSettings[input.name] = input.checked ? 'true' : 'false';
        } else {
            newSettings[input.name] = input.value;
        }
    });

    // Sync companyLogo & invoiceLogo
    const logo = newSettings.invoiceLogo || newSettings.companyLogo || '';
    newSettings.companyLogo = logo;
    newSettings.invoiceLogo = logo;

    try {
        await ipcRenderer.invoke('save-settings', newSettings);
        settings = newSettings;
        showToast('Settings & Invoice Design saved successfully', 'success');
    } catch (error) {
        showToast('Error saving settings: ' + error.message, 'error');
    }
}

async function openDbFolder() {
    try {
        await ipcRenderer.invoke('open-db-folder');
    } catch (err) {
        showToast('Error opening folder: ' + err.message, 'error');
    }
}

async function openDbFile() {
    try {
        await ipcRenderer.invoke('open-db-file');
    } catch (err) {
        showToast('Error opening in Excel: ' + err.message, 'error');
    }
}

async function backupData() {
    try {
        const result = await ipcRenderer.invoke('backup-data');
        if (result.success) {
            showToast('Backup saved successfully!', 'success');
        } else if (!result.cancelled) {
            showToast('Backup failed: ' + (result.error || 'Unknown error'), 'error');
        }
    } catch (err) {
        showToast('Backup error: ' + err.message, 'error');
    }
}

async function restoreData() {
    if (!confirm('Restoring a backup will overwrite the current database with the selected backup file. Continue?')) return;
    try {
        const result = await ipcRenderer.invoke('restore-data');
        if (result.success) {
            showToast('Database restored successfully!', 'success');
            await loadAllData();
            renderAllTables();
        } else if (!result.cancelled) {
            showToast('Restore failed: ' + (result.error || 'Unknown error'), 'error');
        }
    } catch (err) {
        showToast('Restore error: ' + err.message, 'error');
    }
}

async function exportExcel(sheetName = null) {
    try {
        const result = await ipcRenderer.invoke('export-excel', { sheetName });
        if (result.success) {
            showToast((sheetName ? sheetName : 'Database') + ' exported to Excel successfully!', 'success');
        } else if (!result.cancelled) {
            showToast('Export failed: ' + (result.error || 'Unknown error'), 'error');
        }
    } catch (err) {
        showToast('Export error: ' + err.message, 'error');
    }
}

async function resetData() {
    if (!confirm('This will delete all invoices, purchases, vouchers, and custom ledgers and reset to TaxTech ERP defaults. Continue?')) return;
    try {
        const result = await ipcRenderer.invoke('reset-data');
        if (result.success) {
            showToast('Database has been reset to defaults', 'success');
            await loadAllData();
            renderAllTables();
        } else {
            showToast('Reset failed: ' + (result.error || 'Unknown error'), 'error');
        }
    } catch (err) {
        showToast('Reset error: ' + err.message, 'error');
    }
}

function closeAllModals() {
    document.querySelectorAll('.modal').forEach(m => m.classList.remove('active'));
    editingInvoiceId = null;
    editingPurchaseId = null;
    editingProductId = null;
    editingClientId = null;
    editingVoucherId = null;
    editingLedgerId = null;
}

// Utility Functions
function formatNumber(num) {
    return Number(num || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDate(dateStr) {
    if (!dateStr) return '-';
    return new Date(dateStr).toLocaleDateString('en-IN');
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icons = {
        success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
        error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
        warning: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
    };
    toast.innerHTML = `
        <div class="toast-icon">${icons[type]}</div>
        <div class="toast-message">${message}</div>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
}

// ==========================================================
// 11. INVOICE FORMAT, LOGO, ALIGNMENT & LIVE PREVIEW ENGINE
// ==========================================================
let currentPreviewData = null;

function setLogoImage(dataUrl, filePath = '') {
    const logoImg = document.getElementById('logoPreviewImg');
    const logoPlaceholder = document.getElementById('logoPlaceholderText');
    const removeLogoBtn = document.getElementById('removeLogoBtn');
    const compLogoInput = document.getElementById('companyLogoData');
    const invLogoInput = document.getElementById('invoiceLogoData');

    if (dataUrl) {
        if (logoImg) {
            logoImg.src = dataUrl;
            logoImg.style.display = 'block';
        }
        if (logoPlaceholder) logoPlaceholder.style.display = 'none';
        if (removeLogoBtn) removeLogoBtn.style.display = 'inline-flex';

        // If dataUrl is too large (> 30000 chars) for Excel cell and filePath is provided, store filePath
        const storeVal = (dataUrl.length > 30000 && filePath) ? filePath : dataUrl;
        if (compLogoInput) compLogoInput.value = storeVal;
        if (invLogoInput) invLogoInput.value = storeVal;

        if (settings) {
            settings.companyLogo = storeVal;
            settings.invoiceLogo = storeVal;
            settings.companyLogoDataUrl = dataUrl;
        }
    } else {
        if (logoImg) {
            logoImg.src = '';
            logoImg.style.display = 'none';
        }
        if (logoPlaceholder) logoPlaceholder.style.display = 'block';
        if (removeLogoBtn) removeLogoBtn.style.display = 'none';
        if (compLogoInput) compLogoInput.value = '';
        if (invLogoInput) invLogoInput.value = '';
        const fileInput = document.getElementById('logoFileInput');
        if (fileInput) fileInput.value = '';

        if (settings) {
            settings.companyLogo = '';
            settings.invoiceLogo = '';
            settings.companyLogoDataUrl = '';
        }
    }
}

function initInvoiceFormatControls() {
    // 1. Logo Upload & Remove
    document.getElementById('chooseLogoBtn')?.addEventListener('click', async (e) => {
        e.preventDefault();
        try {
            const res = await ipcRenderer.invoke('select-logo-file');
            if (res && res.cancelled) {
                return; // User clicked Cancel in file dialog
            }
            if (res && res.success) {
                const logo = res.dataUrl || res.base64;
                if (logo) {
                    setLogoImage(logo, res.savedPath || res.filePath);
                    showToast('Company logo uploaded successfully!', 'success');
                    return;
                }
            } else if (res && res.error) {
                showToast('Failed to load logo: ' + res.error, 'error');
                return;
            }
        } catch (err) {
            console.warn('select-logo-file IPC error, using DOM file input:', err);
            document.getElementById('logoFileInput')?.click();
        }
    });

    // Also clicking the logo box triggers file selection
    document.getElementById('logoPreviewBox')?.addEventListener('click', (e) => {
        if (e.target.id !== 'removeLogoBtn') {
            document.getElementById('chooseLogoBtn')?.click();
        }
    });

    document.getElementById('logoFileInput')?.addEventListener('change', (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            const dataUrl = evt.target.result;
            if (dataUrl) {
                setLogoImage(dataUrl);
                showToast('Company logo uploaded successfully!', 'success');
            }
        };
        reader.onerror = (err) => {
            showToast('Error reading selected image: ' + err.message, 'error');
        };
        reader.readAsDataURL(file);
    });

    document.getElementById('removeLogoBtn')?.addEventListener('click', (e) => {
        e.stopPropagation();
        setLogoImage('');
        showToast('Logo removed', 'info');
    });

    // 2. Logo Width Range Slider
    document.getElementById('invoiceLogoWidth')?.addEventListener('input', (e) => {
        const val = e.target.value;
        const disp = document.getElementById('logoWidthDisplay');
        if (disp) disp.textContent = val + ' px';
    });

    // 3. Logo Alignment Segmented Buttons
    document.querySelectorAll('#logoAlignGroup .pill-seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#logoAlignGroup .pill-seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const input = document.getElementById('invoiceLogoAlign');
            if (input) input.value = btn.dataset.align;
        });
    });

    // 4. Header Alignment Segmented Buttons
    document.querySelectorAll('#headerAlignGroup .pill-seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#headerAlignGroup .pill-seg-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const input = document.getElementById('invoiceHeaderAlign');
            if (input) input.value = btn.dataset.align;
        });
    });

    // 5. Preset Color Swatches
    document.querySelectorAll('#colorSwatches .color-swatch').forEach(swatch => {
        swatch.addEventListener('click', () => {
            document.querySelectorAll('#colorSwatches .color-swatch').forEach(s => s.classList.remove('active'));
            swatch.classList.add('active');
            const color = swatch.dataset.color;
            const colorInput = document.getElementById('invoiceThemeColor');
            const hexDisplay = document.getElementById('themeColorHex');
            if (colorInput) colorInput.value = color;
            if (hexDisplay) hexDisplay.textContent = color;
        });
    });

    // 6. Custom Color Input
    document.getElementById('invoiceThemeColor')?.addEventListener('input', (e) => {
        const color = e.target.value;
        const hexDisplay = document.getElementById('themeColorHex');
        if (hexDisplay) hexDisplay.textContent = color;
        document.querySelectorAll('#colorSwatches .color-swatch').forEach(s => {
            s.classList.toggle('active', (s.dataset.color || '').toLowerCase() === color.toLowerCase());
        });
    });

    // 7. Preview Sample Invoice Trigger
    document.getElementById('previewSampleInvoiceBtn')?.addEventListener('click', () => {
        const liveSettings = getLiveFormatSettings();
        const { sampleInvoice, sampleItems, sampleClient } = getSampleInvoiceData(liveSettings);
        openInvoicePreviewModal(sampleInvoice, sampleItems, sampleClient, liveSettings, 'Sample Invoice Preview');
    });

    // 8. Preview Draft Invoice Trigger from Invoice Modal
    document.getElementById('previewInvoiceDraftBtn')?.addEventListener('click', () => {
        const draft = getDraftInvoiceData();
        if (!draft) return;
        const liveSettings = getLiveFormatSettings();
        openInvoicePreviewModal(draft.invoice, draft.items, draft.client, liveSettings, 'Draft Invoice Preview');
    });

    // 9. Close Preview Modal
    document.getElementById('closeInvoicePreviewModal')?.addEventListener('click', closeInvoicePreviewModal);
    document.getElementById('invoicePreviewModalOverlay')?.addEventListener('click', closeInvoicePreviewModal);

    // 10. Print Preview Invoice
    document.getElementById('printPreviewInvoiceBtn')?.addEventListener('click', () => {
        window.print();
    });

    // 11. Export PDF from Preview Modal
    document.getElementById('exportPdfPreviewBtn')?.addEventListener('click', async () => {
        if (!currentPreviewData) return;
        try {
            const result = await ipcRenderer.invoke('generate-invoice-pdf', {
                invoice: currentPreviewData.invoice,
                items: currentPreviewData.items,
                settings: currentPreviewData.settings,
                client: currentPreviewData.client
            });
            if (result && result.success) {
                showToast('Invoice PDF generated successfully', 'success');
                await ipcRenderer.invoke('open-file', result.path);
            } else if (result && !result.cancelled) {
                showToast('PDF error: ' + (result.error || 'Failed to generate PDF'), 'error');
            }
        } catch (err) {
            showToast('Error generating PDF: ' + err.message, 'error');
        }
    });
}

function getLiveFormatSettings() {
    const live = { ...settings };
    // Read live fields from company & invoice format form so unsaved adjustments can be previewed
    document.querySelectorAll('#companySettingsForm input, #companySettingsForm textarea').forEach(inp => {
        if (inp.name) live[inp.name] = inp.value;
    });
    document.querySelectorAll('#bankSettingsForm input').forEach(inp => {
        if (inp.name) live[inp.name] = inp.value;
    });
    document.querySelectorAll('#invoiceSettingsForm input, #invoiceSettingsForm textarea').forEach(inp => {
        if (inp.name) live[inp.name] = inp.value;
    });
    document.querySelectorAll('#invoiceFormatSettingsForm input, #invoiceFormatSettingsForm select, #invoiceFormatSettingsForm textarea').forEach(inp => {
        if (inp.type === 'checkbox') {
            live[inp.name] = inp.checked ? 'true' : 'false';
        } else if (inp.name) {
            live[inp.name] = inp.value;
        }
    });

    const logo = live.invoiceLogo || live.companyLogo || '';
    live.companyLogo = logo;
    live.invoiceLogo = logo;
    return live;
}

function getDraftInvoiceData() {
    const invNum = document.getElementById('invoiceNumber')?.value || 'INV-DRAFT';
    const invDate = document.getElementById('invoiceDate')?.value || new Date().toISOString().split('T')[0];
    const clientId = document.getElementById('invoiceClient')?.value || '';
    const pos = document.getElementById('placeOfSupply')?.value || '';
    const revCharge = document.getElementById('reverseCharge')?.checked ? 'Yes' : 'No';
    const disc = parseFloat(document.getElementById('invoiceDiscount')?.value) || 0;
    const round = parseFloat(document.getElementById('invoiceRoundOff')?.value) || 0;

    let client = clients.find(c => c.id === clientId);
    if (!client) {
        client = {
            name: 'Valued Customer',
            address1: 'Customer Address Line 1',
            city: 'Commercial Hub',
            state: pos || settings.companyState || 'Delhi',
            stateCode: '07',
            gstin: '07AAAAA0000A1Z5'
        };
    }

    const items = [];
    document.querySelectorAll('#invoiceItemsTable tbody tr').forEach((tr, idx) => {
        const prodSelect = tr.querySelector('.item-product');
        const prodId = prodSelect ? prodSelect.value : '';
        const pData = products.find(p => p.id === prodId);
        const searchInput = tr.querySelector('.item-search-input');
        const rawSearchText = searchInput ? searchInput.value.replace(/^\[(Goods|Service)\]\s*/i, '').trim() : '';

        const qty = parseFloat(tr.querySelector('.item-qty')?.value) || 1;
        const rate = parseFloat(tr.querySelector('.item-rate')?.value) || 0;
        const itemDisc = parseFloat(tr.querySelector('.item-disc')?.value) || 0;
        const taxable = parseFloat(tr.querySelector('.item-taxable')?.value) || (qty * rate);
        const gstRate = parseFloat(tr.querySelector('.item-gstrate')?.value) || 18;
        const cgstAmt = parseFloat(tr.querySelector('.item-cgst')?.value) || 0;
        const sgstAmt = parseFloat(tr.querySelector('.item-sgst')?.value) || 0;
        const igstAmt = parseFloat(tr.querySelector('.item-igst')?.value) || 0;
        const total = parseFloat(tr.querySelector('.item-total')?.value) || (taxable + cgstAmt + sgstAmt + igstAmt);

        items.push({
            productName: pData ? pData.name : (rawSearchText || `Item ${idx + 1}`),
            hsn: tr.querySelector('.item-hsn')?.value || (pData ? pData.hsn : '998313'),
            type: pData ? (pData.type || 'Goods') : 'Goods',
            quantity: qty,
            unit: tr.querySelector('.item-unit')?.value || 'PCS',
            rate,
            discountPercent: itemDisc,
            discountAmount: (qty * rate * itemDisc) / 100,
            taxableValue: taxable,
            gstRate,
            cgstPercent: cgstAmt > 0 ? gstRate / 2 : 0,
            cgstAmount: cgstAmt,
            sgstPercent: sgstAmt > 0 ? gstRate / 2 : 0,
            sgstAmount: sgstAmt,
            igstPercent: igstAmt > 0 ? gstRate : 0,
            igstAmount: igstAmt,
            totalAmount: total
        });
    });

    if (items.length === 0) {
        items.push({
            productName: 'Professional Technical & Software Services',
            hsn: '998313',
            type: 'Service',
            quantity: 1,
            unit: 'NOS',
            rate: 15000,
            discountPercent: 0,
            discountAmount: 0,
            taxableValue: 15000,
            gstRate: 18,
            cgstPercent: 9,
            cgstAmount: 1350,
            sgstPercent: 9,
            sgstAmount: 1350,
            igstPercent: 0,
            igstAmount: 0,
            totalAmount: 17700
        });
    }

    const subtotal = items.reduce((s, i) => s + i.taxableValue, 0);
    const totalCgst = items.reduce((s, i) => s + i.cgstAmount, 0);
    const totalSgst = items.reduce((s, i) => s + i.sgstAmount, 0);
    const totalIgst = items.reduce((s, i) => s + i.igstAmount, 0);
    const totalTax = totalCgst + totalSgst + totalIgst;
    const grandTotal = Math.max(0, subtotal + totalTax - disc + round);

    const invoice = {
        invoiceNumber: invNum,
        date: invDate,
        clientName: client.name,
        clientGstin: client.gstin || '',
        placeOfSupply: pos || client.state || 'Local',
        reverseCharge: revCharge,
        subtotal,
        totalCgst,
        totalSgst,
        totalIgst,
        totalTax,
        discount: disc,
        roundOff: round,
        grandTotal,
        status: 'Draft'
    };

    return { invoice, items, client };
}

function getSampleInvoiceData(st = {}) {
    const sampleClient = {
        name: 'Apex Infotech & Engineering Pvt Ltd',
        address1: 'Plot 42, Tech Park, Cyber Gateway',
        city: 'Hyderabad',
        state: 'Telangana',
        stateCode: '36',
        gstin: '36AABCU9603R1ZM'
    };

    const sampleItems = [
        {
            productName: 'TaxTech ERP Enterprise Software License (Annual)',
            hsn: '998313',
            type: 'Service',
            quantity: 1,
            unit: 'NOS',
            rate: 28000,
            discountPercent: 10,
            discountAmount: 2800,
            taxableValue: 25200,
            gstRate: 18,
            cgstPercent: 9,
            cgstAmount: 2268,
            sgstPercent: 9,
            sgstAmount: 2268,
            igstPercent: 0,
            igstAmount: 0,
            totalAmount: 29736
        },
        {
            productName: 'High-Speed Thermal Barcode Scanner & Stand',
            hsn: '847160',
            type: 'Goods',
            quantity: 2,
            unit: 'PCS',
            rate: 3500,
            discountPercent: 0,
            discountAmount: 0,
            taxableValue: 7000,
            gstRate: 18,
            cgstPercent: 9,
            cgstAmount: 630,
            sgstPercent: 9,
            sgstAmount: 630,
            igstPercent: 0,
            igstAmount: 0,
            totalAmount: 8260
        }
    ];

    const subtotal = 32200;
    const totalCgst = 2898;
    const totalSgst = 2898;
    const totalIgst = 0;
    const totalTax = 5796;
    const discount = 0;
    const roundOff = 0;
    const grandTotal = 37996;

    const sampleInvoice = {
        invoiceNumber: (st.invoicePrefix || 'INV') + '-2024-001',
        date: new Date().toISOString().split('T')[0],
        clientName: sampleClient.name,
        clientGstin: sampleClient.gstin,
        placeOfSupply: 'Telangana',
        reverseCharge: 'No',
        subtotal,
        totalCgst,
        totalSgst,
        totalIgst,
        totalTax,
        discount,
        roundOff,
        grandTotal,
        status: 'Sample'
    };

    return { sampleInvoice, sampleItems, sampleClient };
}

function openInvoicePreviewModal(invoice, items, client, customSettings = null, badgeText = 'TaxTech ERP Preview') {
    const st = customSettings || settings || {};
    currentPreviewData = { invoice, items, client, settings: st };

    const badge = document.getElementById('previewTemplateBadge');
    if (badge) badge.textContent = badgeText;

    const container = document.getElementById('invoicePaperPreview');
    if (container) {
        container.innerHTML = renderInvoicePaperHTML(invoice, items, client, st);
    }

    const modal = document.getElementById('invoicePreviewModal');
    if (modal) modal.classList.add('active');
}

function closeInvoicePreviewModal() {
    const modal = document.getElementById('invoicePreviewModal');
    if (modal) modal.classList.remove('active');
}

function renderInvoicePaperHTML(invoice, items, client, st) {
    const primaryColor = st.invoiceThemeColor || '#1a1a2e';
    const logoAlign = st.invoiceLogoAlign || 'left';
    const headerAlign = st.invoiceHeaderAlign || 'left';
    const logoWidth = Math.min(180, Math.max(50, parseInt(st.invoiceLogoWidth, 10) || 100));
    const showLogo = st.showLogo !== 'false' && st.showLogo !== false;
    const logoSrc = st.companyLogo || st.invoiceLogo || '';
    const template = st.invoiceTemplate || 'modern';

    let fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif';
    if (st.invoiceFont === 'Times-Roman') fontFamily = '"Times New Roman", Times, Georgia, serif';
    else if (st.invoiceFont === 'Courier') fontFamily = '"Courier New", Courier, monospace';

    const docTitle = st.invoiceTitle || 'TAX INVOICE';
    const docSubtitle = st.invoiceSubtitle || 'Original for Recipient';

    const showHsn = st.showHsn !== 'false' && st.showHsn !== false;
    const showQty = st.showQty !== 'false' && st.showQty !== false;
    const showRate = st.showRate !== 'false' && st.showRate !== false;
    const showDiscount = st.showDiscount !== 'false' && st.showDiscount !== false;
    const showTaxable = st.showTaxable !== 'false' && st.showTaxable !== false;
    const showGstRate = st.showGstRate !== 'false' && st.showGstRate !== false;
    const showAmountInWords = st.showAmountInWords !== 'false' && st.showAmountInWords !== false;
    const showBankDetails = st.showBankDetails !== 'false' && st.showBankDetails !== false;
    const showUpi = st.showUpi !== 'false' && st.showUpi !== false;
    const showTerms = st.showTerms !== 'false' && st.showTerms !== false;
    const showSignature = st.showSignature !== 'false' && st.showSignature !== false;

    // Header layout computation
    let headerHTML = '';
    const companyInfoHTML = `
        <div class="inv-company-block" style="text-align: ${headerAlign};">
            <h1 class="inv-company-title" style="color: ${primaryColor};">${st.companyName || 'My Company'}</h1>
            <p class="inv-company-text">
                ${st.companyAddress1 || ''}${st.companyAddress2 ? ', ' + st.companyAddress2 : ''}<br>
                ${st.companyCity || ''}${st.companyPincode ? ' - ' + st.companyPincode : ''}, ${st.companyState || ''}${st.companyStateCode ? ' (' + st.companyStateCode + ')' : ''}<br>
                Phone: ${st.companyPhone || '-'} | Email: ${st.companyEmail || '-'}<br>
                <strong>GSTIN:</strong> ${st.companyGstin || '-'} | <strong>PAN:</strong> ${st.companyPan || '-'}
            </p>
        </div>
    `;

    const logoHTML = (showLogo && logoSrc) ? `
        <div class="inv-logo-block" style="width: ${logoWidth}px; text-align: ${logoAlign};">
            <img src="${logoSrc}" alt="Company Logo" style="max-width: 100%; max-height: 80px; object-fit: contain;">
        </div>
    ` : '';

    if (logoAlign === 'center' && logoHTML) {
        headerHTML = `
            <div class="inv-header-center-wrap" style="text-align: center; margin-bottom: 12px;">
                <div style="margin-bottom: 6px;">${logoHTML}</div>
                ${companyInfoHTML}
            </div>
        `;
    } else if (logoAlign === 'right' && logoHTML) {
        headerHTML = `
            <div class="inv-header-flex-wrap" style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;">
                <div style="flex: 1;">${companyInfoHTML}</div>
                <div style="margin-left: 20px;">${logoHTML}</div>
            </div>
        `;
    } else if (logoHTML) {
        // Left
        headerHTML = `
            <div class="inv-header-flex-wrap" style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px;">
                <div style="margin-right: 20px;">${logoHTML}</div>
                <div style="flex: 1;">${companyInfoHTML}</div>
            </div>
        `;
    } else {
        headerHTML = `
            <div class="inv-header-plain-wrap" style="margin-bottom: 12px;">
                ${companyInfoHTML}
            </div>
        `;
    }

    // Line items table rows
    let tableRowsHTML = '';
    items.forEach((item, idx) => {
        const taxable = item.taxableValue !== undefined ? Number(item.taxableValue) : (item.quantity * item.rate - (item.discountAmount || 0));
        const lineTotal = item.totalAmount !== undefined ? Number(item.totalAmount) : (taxable + (item.cgstAmount || 0) + (item.sgstAmount || 0) + (item.igstAmount || 0));
        const bg = idx % 2 === 0 ? '#f8fafc' : '#ffffff';

        tableRowsHTML += `
            <tr style="background: ${bg};">
                <td style="text-align: center; width: 30px;">${idx + 1}</td>
                <td style="text-align: left; font-weight: 500;">
                    ${item.productName || 'Item'}
                    ${item.type ? `<span class="inv-type-pill">${item.type}</span>` : ''}
                </td>
                ${showHsn ? `<td style="text-align: center;">${item.hsn || '-'}</td>` : ''}
                ${showQty ? `
                    <td style="text-align: right;">${Number(item.quantity || 1).toFixed(item.unit === 'PCS' || item.unit === 'NOS' ? 0 : 2)}</td>
                    <td style="text-align: center;">${item.unit || ''}</td>
                ` : ''}
                ${showRate ? `<td style="text-align: right;">₹${Number(item.rate || 0).toFixed(2)}</td>` : ''}
                ${showDiscount ? `<td style="text-align: right;">${Number(item.discountPercent || 0).toFixed(1)}%</td>` : ''}
                ${showTaxable ? `<td style="text-align: right;">₹${Number(taxable || 0).toFixed(2)}</td>` : ''}
                ${showGstRate ? `<td style="text-align: right;">${Number(item.gstRate || 0).toFixed(1)}%</td>` : ''}
                <td style="text-align: right; font-weight: 600;">₹${Number(lineTotal || 0).toFixed(2)}</td>
            </tr>
        `;
    });

    return `
        <div class="invoice-paper-root template-${template}" style="font-family: ${fontFamily}; --inv-primary: ${primaryColor};">
            <!-- 1. Header & Company Info -->
            ${headerHTML}

            <!-- 2. Accent Bar & Title Banner -->
            <div class="inv-accent-bar" style="height: 3px; background: ${primaryColor}; margin: 8px 0 12px 0; border-radius: 2px;"></div>
            <div class="inv-title-container" style="text-align: center; margin-bottom: 14px;">
                <h2 class="inv-title-heading" style="color: ${primaryColor}; margin: 0; font-size: 1.35rem; font-weight: 700; letter-spacing: 1px;">${docTitle}</h2>
                ${docSubtitle ? `<div class="inv-title-sub" style="font-size: 0.78rem; color: #64748b; font-style: italic; margin-top: 2px;">${docSubtitle}</div>` : ''}
            </div>

            <!-- 3. Meta Grid: Buyer Box & Invoice Box -->
            <div class="inv-meta-grid" style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px;">
                <div class="inv-box" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 14px;">
                    <div style="font-size: 0.75rem; font-weight: 700; color: ${primaryColor}; text-transform: uppercase; margin-bottom: 4px; letter-spacing: 0.5px;">Bill To (Buyer)</div>
                    <div style="font-size: 0.88rem; font-weight: 700; color: #0f172a; margin-bottom: 2px;">${client.name || 'Walk-in Customer'}</div>
                    <div style="font-size: 0.78rem; color: #475569; line-height: 1.4;">
                        ${client.address1 || ''}${client.city ? ', ' + client.city : ''}${client.state ? ', ' + client.state : ''}<br>
                        ${client.gstin ? `<strong>GSTIN:</strong> ${client.gstin}<br>` : ''}
                        State Code: ${client.stateCode || '-'} | Place of Supply: ${invoice.placeOfSupply || client.state || '-'}
                    </div>
                </div>

                <div class="inv-box" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 14px;">
                    <div style="font-size: 0.75rem; font-weight: 700; color: ${primaryColor}; text-transform: uppercase; margin-bottom: 6px; letter-spacing: 0.5px;">Invoice Details</div>
                    <div style="font-size: 0.80rem; line-height: 1.5; color: #334155;">
                        <div style="display:flex; justify-content:space-between;"><span style="color:#64748b;">Invoice No:</span> <strong style="color:#0f172a;">${invoice.invoiceNumber || '-'}</strong></div>
                        <div style="display:flex; justify-content:space-between;"><span style="color:#64748b;">Invoice Date:</span> <span>${formatDate(invoice.date)}</span></div>
                        <div style="display:flex; justify-content:space-between;"><span style="color:#64748b;">Reverse Charge:</span> <span>${invoice.reverseCharge || 'No'}</span></div>
                        <div style="display:flex; justify-content:space-between;"><span style="color:#64748b;">Invoice Series:</span> <span>${st.invoiceSeries || 'Default'}</span></div>
                    </div>
                </div>
            </div>

            <!-- 4. Dynamic Line Items Table -->
            <table class="inv-render-table" style="width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 0.78rem;">
                <thead>
                    <tr style="background: ${primaryColor}; color: #ffffff;">
                        <th style="padding: 6px 8px; text-align: center; border: 1px solid ${primaryColor}; width: 30px;">#</th>
                        <th style="padding: 6px 8px; text-align: left; border: 1px solid ${primaryColor};">Item / Service Description</th>
                        ${showHsn ? `<th style="padding: 6px 8px; text-align: center; border: 1px solid ${primaryColor};">HSN/SAC</th>` : ''}
                        ${showQty ? `
                            <th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">Qty</th>
                            <th style="padding: 6px 8px; text-align: center; border: 1px solid ${primaryColor};">Unit</th>
                        ` : ''}
                        ${showRate ? `<th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">Rate (₹)</th>` : ''}
                        ${showDiscount ? `<th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">Disc%</th>` : ''}
                        ${showTaxable ? `<th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">Taxable (₹)</th>` : ''}
                        ${showGstRate ? `<th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">GST%</th>` : ''}
                        <th style="padding: 6px 8px; text-align: right; border: 1px solid ${primaryColor};">Total (₹)</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRowsHTML}
                </tbody>
            </table>

            <!-- 5. Totals & Payment Section -->
            <div class="inv-summary-section" style="display: grid; grid-template-columns: 1.2fr 1fr; gap: 16px; margin-bottom: 14px;">
                <!-- Left: Words, Bank & UPI -->
                <div>
                    ${showAmountInWords ? `
                        <div style="margin-bottom: 8px;">
                            <span style="font-size: 0.75rem; font-weight: 700; color: ${primaryColor};">Amount in Words:</span><br>
                            <span style="font-size: 0.80rem; font-style: italic; color: #1e293b; font-weight: 600;">${numberToWords(invoice.grandTotal)} Rupees Only</span>
                        </div>
                    ` : ''}

                    ${showBankDetails && (st.bankName || st.bankAccount) ? `
                        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 12px; margin-top: 6px; font-size: 0.75rem;">
                            <div style="font-weight: 700; color: ${primaryColor}; margin-bottom: 3px; font-size: 0.72rem; text-transform: uppercase;">Bank Payment Details</div>
                            <div style="color: #334155;">Bank: <strong>${st.bankName || '-'}</strong> | A/c: <strong>${st.bankAccount || '-'}</strong></div>
                            <div style="color: #334155;">IFSC: <strong>${st.bankIfsc || '-'}</strong> | Branch: ${st.bankBranch || '-'}</div>
                            ${showUpi && st.companyUpiId ? `
                                <div style="color: #059669; font-weight: 700; margin-top: 4px;">UPI ID for Instant Payment: ${st.companyUpiId}</div>
                            ` : ''}
                        </div>
                    ` : ''}
                </div>

                <!-- Right: Summary Table -->
                <div>
                    <table style="width: 100%; border-collapse: collapse; font-size: 0.80rem;">
                        <tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#475569;">Taxable Subtotal</td><td style="padding: 4px 6px; text-align: right; font-weight: 600;">₹${formatNumber(invoice.subtotal)}</td></tr>
                        ${Number(invoice.totalCgst || 0) > 0 ? `<tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#475569;">CGST Output</td><td style="padding: 4px 6px; text-align: right;">₹${formatNumber(invoice.totalCgst)}</td></tr>` : ''}
                        ${Number(invoice.totalSgst || 0) > 0 ? `<tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#475569;">SGST Output</td><td style="padding: 4px 6px; text-align: right;">₹${formatNumber(invoice.totalSgst)}</td></tr>` : ''}
                        ${Number(invoice.totalIgst || 0) > 0 ? `<tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#475569;">IGST Output</td><td style="padding: 4px 6px; text-align: right;">₹${formatNumber(invoice.totalIgst)}</td></tr>` : ''}
                        ${Number(invoice.discount || 0) > 0 ? `<tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#dc2626;">Discount</td><td style="padding: 4px 6px; text-align: right; color:#dc2626;">-₹${formatNumber(invoice.discount)}</td></tr>` : ''}
                        ${Number(invoice.roundOff || 0) !== 0 ? `<tr style="border-bottom: 1px solid #e2e8f0;"><td style="padding: 4px 6px; color:#475569;">Round Off</td><td style="padding: 4px 6px; text-align: right;">₹${formatNumber(invoice.roundOff)}</td></tr>` : ''}
                        <tr style="background: ${primaryColor}; color: #ffffff; font-weight: 700; font-size: 0.88rem;">
                            <td style="padding: 6px 8px;">Grand Total (INR)</td>
                            <td style="padding: 6px 8px; text-align: right;">₹${formatNumber(invoice.grandTotal)}</td>
                        </tr>
                    </table>
                </div>
            </div>

            <!-- 6. Terms, Declaration & Signature Footer -->
            <div class="inv-footer-section" style="display: grid; grid-template-columns: 1.4fr 1fr; gap: 16px; margin-top: 14px; border-top: 1px dashed #cbd5e1; padding-top: 10px;">
                <div>
                    ${showTerms && st.termsConditions ? `
                        <div style="margin-bottom: 6px;">
                            <span style="font-size: 0.72rem; font-weight: 700; color: ${primaryColor};">Terms & Conditions:</span>
                            <div style="font-size: 0.70rem; color: #475569; line-height: 1.35;">${st.termsConditions.replace(/\n/g, '<br>')}</div>
                        </div>
                    ` : ''}
                    ${st.invoiceDeclaration ? `
                        <div style="font-size: 0.68rem; color: #64748b; font-style: italic;">${st.invoiceDeclaration}</div>
                    ` : ''}
                </div>

                ${showSignature ? `
                    <div style="text-align: right; display: flex; flex-direction: column; justify-content: flex-end; align-items: flex-end;">
                        <div style="font-size: 0.75rem; font-weight: 600; color: #334155;">For <strong>${st.signatoryCompany || st.companyName || 'TaxTech ERP Solutions'}</strong></div>
                        <div style="height: 40px;"></div>
                        <div style="font-size: 0.72rem; font-weight: 600; color: #64748b; border-top: 1px solid #94a3b8; padding-top: 4px; min-width: 140px; text-align: center;">${st.signatoryLabel || 'Authorized Signatory'}</div>
                    </div>
                ` : ''}
            </div>
        </div>
    `;
}

function numberToWords(num) {
    const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    function convertHundreds(n) {
        if (n < 20) return a[n];
        if (n < 100) return b[Math.floor(n / 10)] + (n % 10 ? ' ' + a[n % 10] : '');
        return a[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' and ' + convertHundreds(n % 100) : '');
    }

    const n = Math.abs(Number(num) || 0);
    if (n === 0) return 'Zero';

    const intPart = Math.floor(n);
    const decimal = Math.round((n - intPart) * 100);

    const crore = Math.floor(intPart / 10000000);
    const lakh = Math.floor((intPart % 10000000) / 100000);
    const thousand = Math.floor((intPart % 100000) / 1000);
    const hundred = Math.floor((intPart % 1000) / 100);
    const rest = Math.floor(intPart % 100);

    let words = '';
    if (crore) words += convertHundreds(crore) + ' Crore ';
    if (lakh) words += convertHundreds(lakh) + ' Lakh ';
    if (thousand) words += convertHundreds(thousand) + ' Thousand ';
    if (hundred) words += convertHundreds(hundred) + ' Hundred ';
    if (rest) words += (words ? 'and ' : '') + convertHundreds(rest);
    if (decimal) words += ' and ' + convertHundreds(decimal) + ' Paise';

    return words.trim();
}

// Global scope exposure for inline onclick handlers
window.editInvoice = editInvoice;
window.viewInvoice = viewInvoice;
window.deleteInvoice = deleteInvoice;
window.editPurchase = editPurchase;
window.deletePurchase = deletePurchase;
window.editProduct = editProduct;
window.deleteProduct = deleteProduct;
window.editClient = editClient;
window.deleteClient = deleteClient;
window.deleteVoucher = deleteVoucher;
window.editLedger = editLedger;
window.deleteLedger = deleteLedger;
window.onProductSelect = onProductSelect;
window.onPurchaseProductSelect = onPurchaseProductSelect;
window.calculateItemRow = calculateItemRow;
window.calculatePurchaseItemRow = calculatePurchaseItemRow;
window.removeInvoiceItem = removeInvoiceItem;
window.removePurchaseItem = removePurchaseItem;
window.closeAllComboboxes = closeAllComboboxes;
window.setupProductCombobox = setupProductCombobox;
window.openInvoicePreviewModal = openInvoicePreviewModal;
window.closeInvoicePreviewModal = closeInvoicePreviewModal;
window.setLogoImage = setLogoImage;
window.getLiveFormatSettings = getLiveFormatSettings;
window.numberToWords = numberToWords;
