// ============================================
// FoodX POS PRO - Multiple Client Rows System
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    // ==========================================
    // DATA MIGRATION (EXCEL IMPORT)
            window.downloadSalesTemplate = function() {
        if (typeof XLSX === 'undefined') {
            showNotification('Error: Libreria Excel no cargada', 'error');
            return;
        }
        const data = [
            { "FECHA (DD/MM/AAAA)": "15/01/2026", "TOTAL": 15000 },
            { "FECHA (DD/MM/AAAA)": "16/01/2026", "TOTAL": 25500 }
        ];
        const ws = XLSX.utils.json_to_sheet(data);
        
        ws['!cols'] = [ { wch: 25 }, { wch: 15 } ];
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Plantilla_Ventas");
        XLSX.writeFile(wb, "Plantilla_Importar_Ventas_Minesof.xlsx");
        showNotification('Plantilla de Ventas descargada');
    };

    window.downloadExpensesTemplate = function() {
        if (typeof XLSX === 'undefined') {
            showNotification('Error: Libreria Excel no cargada', 'error');
            return;
        }
        const data = [
            { "CATEGORIA": "Servicios", "DESCRIPCION": "Pago de Luz", "VALOR": 80000, "FECHA": "2026-01-15" },
            { "CATEGORIA": "Suministros", "DESCRIPCION": "Compra de bolsas", "VALOR": 12000, "FECHA": "2026-01-16" }
        ];
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Plantilla_Gastos");
        XLSX.writeFile(wb, "Plantilla_Importar_Gastos_Minesof.xlsx");
        showNotification('Plantilla de Gastos descargada');
    };
    // ==========================================
    const importSalesFile = document.getElementById('importSalesFile');
    if (importSalesFile) {
        importSalesFile.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            handleImportSales(file);
            e.target.value = '';
        });
    }

    const importExpensesFile = document.getElementById('importExpensesFile');
    if (importExpensesFile) {
        importExpensesFile.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            handleImportExpenses(file);
            e.target.value = '';
        });
    }

    function handleImportSales(file) {
        if (typeof XLSX === 'undefined') {
            showNotification('Error: Libreria Excel no cargada', 'error');
            return;
        }
        showNotification('Procesando ventas...', 'info');
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array', cellDates: true});
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
                const jsonDataRaw = XLSX.utils.sheet_to_json(worksheet, {raw: true, defval: null});
                let jsonData = jsonDataRaw;
                
                const hasValidKeys = (row) => {
                    const keys = Object.keys(row).map(k => k.toLowerCase());
                    return keys.some(k => k.includes('total') || k.includes('valor') || k.includes('fecha'));
                };
                
                if (jsonData.length > 0 && !hasValidKeys(jsonData[0])) {
                    const aoa = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                    let headerRowIndex = 0;
                    for (let i = 0; i < aoa.length; i++) {
                        if (aoa[i] && aoa[i].some(cell => typeof cell === 'string' && (cell.toLowerCase().includes('total') || cell.toLowerCase().includes('valor') || cell.toLowerCase().includes('fecha')))) {
                            headerRowIndex = i;
                            break;
                        }
                    }
                    if (headerRowIndex > 0) {
                        const range = XLSX.utils.decode_range(worksheet['!ref']);
                        range.s.r = headerRowIndex;
                        worksheet['!ref'] = XLSX.utils.encode_range(range);
                        jsonData = XLSX.utils.sheet_to_json(worksheet, {raw: true, defval: null});
                    }
                }
                
                if (jsonData.length === 0) {
                    showNotification('El Excel está vacío', 'error');
                    return;
                }
                
                const orders = StorageManager.getOrders();
                let importedCount = 0;
                
                jsonData.forEach(row => {
                    const getCol = (keyParts) => {
                        const key = Object.keys(row).find(k => keyParts.some(p => k.toLowerCase().includes(p)));
                        return key ? row[key] : null;
                    };
                    
                    const rawDate = getCol(['fecha', 'date', 'dia']);
                    const rawTotal = getCol(['total', 'valor', 'monto', 'precio']);
                    const rawNumber = getCol(['numero', 'num', 'pedido', 'order', 'factura']);
                    
                    if (rawTotal === null || rawTotal === undefined) return;
                    if (typeof rawDate === 'string' && rawDate.toLowerCase().includes('total')) return;
                    
                                        let dateStr = new Date().toISOString();
                    if (rawDate) {
                        if (typeof rawDate === 'string') {
                            let str = rawDate.trim();
                            // Attempt to parse DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD
                            const parts = str.split(/[/-]/);
                            if (parts.length >= 3) {
                                let year = 0, month = 0, day = 0;
                                if (parts[0].length === 4) { // YYYY-MM-DD
                                    year = parseInt(parts[0], 10);
                                    month = parseInt(parts[1], 10) - 1;
                                    day = parseInt(parts[2], 10);
                                } else { // DD-MM-YYYY
                                    day = parseInt(parts[0], 10);
                                    month = parseInt(parts[1], 10) - 1;
                                    year = parseInt(parts[2], 10);
                                    if (year < 100) year += 2000;
                                }
                                const pd = new Date(year, month, day);
                                if (!isNaN(pd)) dateStr = pd.toISOString();
                            } else {
                                const pd = new Date(str);
                                if (!isNaN(pd)) dateStr = pd.toISOString();
                            }
                        } else {
                            const pd = new Date(rawDate);
                            if (!isNaN(pd)) dateStr = pd.toISOString();
                        }
                    }
                    
                    let price = 0;
                    if (typeof rawTotal === 'number') {
                        price = rawTotal;
                    } else {
                        let s = rawTotal.toString().toLowerCase().replace(/[^0-9,-]/g, '');
                        s = s.replace(',', '.');
                        price = parseFloat(s);
                    }
                    
                    if (isNaN(price)) return;
                    
                    const orderNum = rawNumber || ('MIG-' + Math.floor(Math.random() * 10000));
                    
                    const newOrder = {
                        id: 'mig_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
                        orderNumber: orderNum.toString(),
                        createdAt: dateStr,
                        paid: true,
                        totalPrice: price,
                        paymentMethod: 'Efectivo',
                        checkoutPrinted: true,
                        createdBy: 'Migracion',
                        items: [{ name: 'Venta HistÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³rica Migrada', price: price, quantity: 1, total: price }]
                    };
                    orders.push(newOrder);
                    if (typeof StorageManager.syncOrderToCloud === 'function') {
                        StorageManager.syncOrderToCloud(newOrder);
                    }
                    importedCount++;
                });
                
                StorageManager.saveOrders(orders);
                showNotification('Exito: ' + importedCount + ' ventas importadas');
                if (state.currentPage === 'history') renderHistoryPage();
            } catch (err) {
                console.error(err);
                showNotification('Error procesando Excel', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleImportExpenses(file) {
        if (typeof XLSX === 'undefined') return;
        showNotification('Procesando gastos...', 'info');
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array', cellDates: true});
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
                const jsonDataRaw = XLSX.utils.sheet_to_json(worksheet, {raw: true, defval: null});
                let jsonData = jsonDataRaw;
                
                const hasValidKeys = (row) => {
                    const keys = Object.keys(row).map(k => k.toLowerCase());
                    return keys.some(k => k.includes('valor') || k.includes('monto') || k.includes('fecha'));
                };
                
                if (jsonData.length > 0 && !hasValidKeys(jsonData[0])) {
                    const aoa = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                    let headerRowIndex = 0;
                    for (let i = 0; i < aoa.length; i++) {
                        if (aoa[i] && aoa[i].some(cell => typeof cell === 'string' && (cell.toLowerCase().includes('valor') || cell.toLowerCase().includes('monto') || cell.toLowerCase().includes('fecha')))) {
                            headerRowIndex = i;
                            break;
                        }
                    }
                    if (headerRowIndex > 0) {
                        const range = XLSX.utils.decode_range(worksheet['!ref']);
                        range.s.r = headerRowIndex;
                        worksheet['!ref'] = XLSX.utils.encode_range(range);
                        jsonData = XLSX.utils.sheet_to_json(worksheet, {raw: true, defval: null});
                    }
                }
                
                if (jsonData.length === 0) return;
                
                const expenses = StorageManager.getExpenses();
                let importedCount = 0;
                
                jsonData.forEach(row => {
                    const getCol = (keyParts) => {
                        const key = Object.keys(row).find(k => keyParts.some(p => k.toLowerCase().includes(p)));
                        return key ? row[key] : null;
                    };
                    
                    const rawDate = getCol(['fecha', 'date']);
                    const rawAmount = getCol(['valor', 'monto', 'total', 'precio']);
                    const rawDesc = getCol(['desc', 'concepto', 'detalle']);
                    const rawCat = getCol(['cat', 'rubro']);
                    
                    if (rawAmount === null || rawAmount === undefined) return;
                    if (typeof rawDate === 'string' && rawDate.toLowerCase().includes('total')) return;
                    
                                        let dateStr = new Date().toISOString();
                    if (rawDate) {
                        if (typeof rawDate === 'string') {
                            let str = rawDate.trim();
                            // Attempt to parse DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD
                            const parts = str.split(/[/-]/);
                            if (parts.length >= 3) {
                                let year = 0, month = 0, day = 0;
                                if (parts[0].length === 4) { // YYYY-MM-DD
                                    year = parseInt(parts[0], 10);
                                    month = parseInt(parts[1], 10) - 1;
                                    day = parseInt(parts[2], 10);
                                } else { // DD-MM-YYYY
                                    day = parseInt(parts[0], 10);
                                    month = parseInt(parts[1], 10) - 1;
                                    year = parseInt(parts[2], 10);
                                    if (year < 100) year += 2000;
                                }
                                const pd = new Date(year, month, day);
                                if (!isNaN(pd)) dateStr = pd.toISOString();
                            } else {
                                const pd = new Date(str);
                                if (!isNaN(pd)) dateStr = pd.toISOString();
                            }
                        } else {
                            const pd = new Date(rawDate);
                            if (!isNaN(pd)) dateStr = pd.toISOString();
                        }
                    }
                    
                    let amount = 0;
                    if (typeof rawAmount === 'number') {
                        amount = rawAmount;
                    } else {
                        let s = rawAmount.toString().toLowerCase().replace(/[^0-9,-]/g, '');
                        s = s.replace(',', '.');
                        amount = parseFloat(s);
                    }
                    
                    if (isNaN(amount)) return;
                    
                    let matchedCatId = 'gastos_generales';
                    if (rawCat) {
                        const searchName = rawCat.toString().toLowerCase().trim();
                        const allCats = StorageManager.getExpenseCategories();
                        const found = allCats.find(c => c.label.toLowerCase().trim() === searchName);
                        matchedCatId = found ? found.id : rawCat.toString().trim();
                    }
                    
                    const newExpense = {
                        id: 'mig_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
                        description: rawDesc ? rawDesc.toString().trim() : 'Gasto Migrado',
                        amount: amount,
                        category: matchedCatId,
                        date: dateStr.split('T')[0],
                        createdAt: dateStr
                    };
                    expenses.push(newExpense);
                    if (typeof StorageManager.syncExpenseToCloud === 'function') {
                        StorageManager.syncExpenseToCloud(newExpense);
                    }
                    importedCount++;
                });
                
                StorageManager.saveExpenses(expenses);
                showNotification('Exito: ' + importedCount + ' gastos importados');
                if (state.currentPage === 'expenses') renderExpensesPage();
            } catch (err) {
                console.error(err);
                showNotification('Error procesando Excel', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }
        // --- Account (Login) Password Change Logic ---
    const saveAccountPasswordBtn = document.getElementById('saveAccountPasswordBtn');
    const newAccountPassword = document.getElementById('newAccountPassword');
    
    if (saveAccountPasswordBtn && newAccountPassword) {
        saveAccountPasswordBtn.addEventListener('click', () => {
            const newPass = newAccountPassword.value;
            if (newPass.length < 6) {
                showNotification('La contrase\u00f1a debe tener al menos 6 caracteres', 'error');
                return;
            }
            const user = window.auth ? window.auth.currentUser : null;
            if (user) {
                user.updatePassword(newPass).then(() => {
                    showNotification('ContraseÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±a de ingreso actualizada exitosamente. Reiniciando sistema...');
                    newAccountPassword.value = '';
                    setTimeout(() => window.location.reload(), 1500);
                }).catch((error) => {
                    if (error.code === 'auth/requires-recent-login') {
                        showNotification('Por seguridad, debes cerrar sesi\u00f3n y volver a entrar antes de cambiar tu contrase\u00f1a.', 'error');
                    } else {
                        showNotification('Error al cambiar contrase\u00f1a: ' + error.message, 'error');
                    }
                });
            } else {
                showNotification('No hay un usuario autenticado', 'error');
            }
        });
    }

    
    let checkoutMode = 'to-print';
    
    const businessNameInput = document.getElementById('businessNameInput');
    const billingSystemInput = document.getElementById('billingSystemInput');
    const deviceUserInput = document.getElementById('deviceUserInput');
    const saveBusinessBrandBtn = document.getElementById('saveBusinessBrandBtn');
    
    function updateAppBranding() {
        const config = StorageManager.getConfig();
        const bName = config.businessName || 'Minesof';
        const billingSys = config.billingSystem || 'standard';
        const dUser = localStorage.getItem('minesof_deviceUser') || 'Caja';
        
        const headerNames = document.querySelectorAll('.dynamic-business-name');
        headerNames.forEach(el => el.textContent = bName);

        const deviceUserNames = document.querySelectorAll('.dynamic-device-user');
        deviceUserNames.forEach(el => el.textContent = dUser);
        
        if (businessNameInput) businessNameInput.value = bName;
        if (billingSystemInput) {
            billingSystemInput.value = billingSys;
            const btn = document.getElementById('adminBillingSystemBtn');
            const btnText = document.getElementById('adminBillingSystemBtnText');
            if (btn && btnText) {
                if (billingSys === 'standard') {
                    btn.style.backgroundColor = '#f97316';
                    btn.style.boxShadow = '0 4px 10px rgba(249, 115, 22, 0.2)';
                    btnText.innerHTML = 'Cobro Din&aacute;mico (Mesas-Pedidos)';
                } else {
                    btn.style.backgroundColor = '#3b82f6';
                    btn.style.boxShadow = '0 4px 10px rgba(59, 130, 246, 0.2)';
                    btnText.innerHTML = 'Cobro Directo (Caja R&aacute;pida)';
                }
            }
        }
        if (deviceUserInput) deviceUserInput.value = dUser === 'Caja' ? '' : dUser;
        
        // Hide/Show tabs based on billing system
        const tabs = document.querySelectorAll('.checkout-tab');
        let needsTabSwitch = false;
        tabs.forEach(tab => {
            if (billingSys === 'direct') {
                if (tab.dataset.tab === 'to-print' || tab.dataset.tab === 'pending') {
                    tab.style.display = 'none';
                    if (tab.classList.contains('active')) {
                        tab.classList.remove('active');
                        needsTabSwitch = true;
                    }
                } else {
                    tab.style.display = 'flex';
                }
            } else {
                tab.style.display = 'flex';
            }
        });
        
        if (needsTabSwitch) {
            const paidTab = document.querySelector('.checkout-tab[data-tab="paid"]');
            if (paidTab) {
                paidTab.classList.add('active');
            }
            checkoutMode = 'paid';
        }
    }
    
    if (saveBusinessBrandBtn) {
        saveBusinessBrandBtn.addEventListener('click', () => {
            const config = StorageManager.getConfig();
            config.businessName = businessNameInput.value.trim();
            // Removed billingSystemInput from here
            StorageManager.saveConfig(config);
            
            if (deviceUserInput) {
                localStorage.setItem('minesof_deviceUser', deviceUserInput.value.trim());
            }
            
            updateAppBranding();
            showNotification('Ajustes guardados. Actualizando sistema...');
            setTimeout(() => window.location.reload(), 1500);
        });
    }

    const saveBillingSystemBtn = document.getElementById('saveBillingSystemBtn');
    if (saveBillingSystemBtn) {
        saveBillingSystemBtn.addEventListener('click', () => {
            window.minesofConfirm("ATENCION: Cambiar el sistema de cobro modificara el flujo de tu negocio.\n\nEstas seguro de querer guardar este cambio?", () => {
                window.minesofPrompt("Escribe CAMBIAR en mayusculas para confirmar:", (confirmWord) => {
                    if (confirmWord !== "CAMBIAR") {
                        showNotification("Cambio cancelado.", "error");
                        return;
                    }
                    
                    const config = StorageManager.getConfig();
                    if (billingSystemInput) config.billingSystem = billingSystemInput.value;
                    StorageManager.saveConfig(config);
                    
                    updateAppBranding();
                    showNotification('Sistema de cobro actualizado...');
                    setTimeout(() => window.location.reload(), 1500);
                });
            });
        });
    }

    

    // Escuchar cambios de configuracion para actualizar logo
    window.addEventListener('configLoadedFromCloud', updateAppBranding);
    
    // Inicializar branding al cargar
    updateAppBranding();
    // Initialize Lucide icons
    if (typeof lucide !== 'undefined') lucide.createIcons();

    // Custom Prompt logic
    window.minesofPrompt = function(message, onAccept, defaultValue = '') {
        const modal = document.getElementById('customPromptModal');
        const input = document.getElementById('customPromptInput');
        const msgEl = document.getElementById('customPromptMessage');
        const cancelBtn = document.getElementById('customPromptCancelBtn');
        const acceptBtn = document.getElementById('customPromptAcceptBtn');

        if (!modal) {
            // Fallback to native
            const val = prompt(message, defaultValue);
            if (typeof onAccept === 'function') onAccept(val);
            return;
        }

        // Remove old listeners
        const newCancelBtn = cancelBtn.cloneNode(true);
        const newAcceptBtn = acceptBtn.cloneNode(true);
        cancelBtn.parentNode.replaceChild(newCancelBtn, cancelBtn);
        acceptBtn.parentNode.replaceChild(newAcceptBtn, acceptBtn);

        msgEl.textContent = message;
        input.value = defaultValue;

        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        input.focus();

        const close = () => modal.classList.remove('open');

        newCancelBtn.addEventListener('click', () => {
            close();
            if (typeof onAccept === 'function') onAccept(null);
        });

        newAcceptBtn.addEventListener('click', () => {
            const val = document.getElementById('customPromptInput').value;
            close();
            if (typeof onAccept === 'function') onAccept(val);
        });

        const newInput = document.getElementById('customPromptInput');
        newInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                newAcceptBtn.click();
            }
        };
    };

    // Custom Confirm logic
    window.minesofConfirm = function(message, onAccept) {
        const modal = document.getElementById('customConfirmModal');
        const overlay = document.getElementById('customConfirmOverlay');
        const msgEl = document.getElementById('customConfirmMessage');
        const cancelBtn = document.getElementById('customConfirmCancelBtn');
        const acceptBtn = document.getElementById('customConfirmAcceptBtn');

        if (!modal) {
            // Fallback to native if not found
            if (confirm('Minesof\n\n' + message) && typeof onAccept === 'function') {
                onAccept();
            }
            return;
        }

        // Remove old listeners by cloning
        const newCancelBtn = cancelBtn.cloneNode(true);
        const newAcceptBtn = acceptBtn.cloneNode(true);
        cancelBtn.parentNode.replaceChild(newCancelBtn, cancelBtn);
        acceptBtn.parentNode.replaceChild(newAcceptBtn, acceptBtn);

        // Update text
        msgEl.textContent = message.replace('Minesof\n\n', '').replace('Minesof\n', '');

        // Show modal
        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();

        // Handlers
        const close = () => modal.classList.remove('open');

        newCancelBtn.addEventListener('click', close);
        overlay.addEventListener('click', close);
        newAcceptBtn.addEventListener('click', () => {
            close();
            if (typeof onAccept === 'function') onAccept();
        });
    };

    // App State
    window.appState = {
        currentPage: 'new-order',
        serviceType: 'salon',
        selectedCategory: 'all',
        searchQuery: '',
        cart: [],
        clients: ['P1'],
        activeClient: 'P1',
        isEditingClient: false,
        orderTotal: 0,
        categoryData: {},
        rowCounter: 0,
        isAdminAuthenticated: false,
        pendingAdminAction: null,
        pendingAdminPage: null,
        appendingOrderId: null,
        selectedPaymentMethod: 'efectivo',
        editingNoteItemId: null
    };
    const state = window.appState;

    // Listen for config loaded from cloud
    window.addEventListener('configLoadedFromCloud', () => {
        console.log('ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â°ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚ÂÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¾ Config loaded from cloud, refreshing UI...');
        renderPosCategories();
        renderPosProducts();
        renderPosCart();
        if (typeof lucide !== 'undefined') lucide.createIcons();
    });

    function initializeCategories() {
        renderPosCategories();
        renderPosProducts();
        renderPosCart();
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // DOM Elements
    const elements = {
        menuBtn: document.getElementById('menuBtn'),
        navDrawer: document.getElementById('navDrawer'),
        drawerOverlay: document.getElementById('drawerOverlay'),
        closeDrawer: document.getElementById('closeDrawer'),
        drawerItems: document.querySelectorAll('.drawer-item'),
        pages: document.querySelectorAll('.page'),
        serviceTabs: document.querySelectorAll('.service-tab'),
        categorySections: document.querySelectorAll('.category-section'),
        totalAmount: document.getElementById('totalAmount'),
        sendToKitchenBtn: document.getElementById('sendToKitchenBtn'),
        btnAddClient: document.getElementById('btnAddClient'),
        posClientsTabs: document.getElementById('posClientsTabs'),
        posProductSearch: document.getElementById('posProductSearch'),
        posClearSearch: document.getElementById('posClearSearch'),
        posCategoriesBar: document.getElementById('posCategoriesBar'),
        posProductsGrid: document.getElementById('posProductsGrid'),
        posClearCartBtn: document.getElementById('posClearCartBtn'),
        posSubmitOrderBtn: document.getElementById('posSubmitOrderBtn'),
        itemNoteModal: document.getElementById('itemNoteModal'),
        closeItemNoteModal: document.getElementById('closeItemNoteModal'),
        closeItemNoteOverlay: document.getElementById('closeItemNoteOverlay'),
        cancelItemNoteModal: document.getElementById('cancelItemNoteModal'),
        saveItemNoteModal: document.getElementById('saveItemNoteModal'),
        // Checkout / Payment
        toPrintCount: document.getElementById('toPrintCount'),
        pendingPaymentCount: document.getElementById('pendingPaymentCount'),
        paidOrdersCount: document.getElementById('paidOrdersCount'),
        toPrintList: document.getElementById('toPrintList'),
        pendingPaymentList: document.getElementById('pendingPaymentList'),
        paidOrdersList: document.getElementById('paidOrdersList'),
        paymentModal: document.getElementById('paymentModal'),
        paymentModalOverlay: document.getElementById('paymentModalOverlay'),
        paymentOrderNum: document.getElementById('paymentOrderNum'),
        paymentTicketContent: document.getElementById('paymentTicketContent'),
        paymentTotal: document.getElementById('paymentTotal'),
        printPaymentTicket: document.getElementById('printPaymentTicket'),
        markReadyBtn: document.getElementById('markReadyBtn'),
        cancelPayment: document.getElementById('cancelPayment'),
        confirmPayment: document.getElementById('confirmPayment'),
        invoicePaymentTicket: document.getElementById('invoicePaymentTicket'),
        deleteOrderBtn: document.getElementById('deleteOrderBtn'),
        // Ticket Modal
        ticketModal: document.getElementById('ticketModal'),
        ticketContent: document.getElementById('ticketContent'),
        cancelTicket: document.getElementById('cancelTicket'),
        cancelTicketFooter: document.getElementById('cancelTicketFooter'),
        closeTicketModal: document.getElementById('closeTicketModal'),
        printTicket: document.getElementById('printTicket'),
        confirmTicket: document.getElementById('confirmTicket'),
        // Admin
        adminTabs: document.querySelectorAll('.admin-tab'),
        adminPanels: document.querySelectorAll('.admin-panel'),
        adminCategoriesList: document.getElementById('adminCategoriesList'),
        adminCategorySelectFlavors: document.getElementById('adminCategorySelectFlavors'),
        adminCategorySelectExtras: document.getElementById('adminCategorySelectExtras'),
        adminCategorySelectObs: document.getElementById('adminCategorySelectObs'),
        adminFlavorsList: document.getElementById('adminFlavorsList'),
        adminExtrasList: document.getElementById('adminExtrasList'),
        adminObsList: document.getElementById('adminObsList'),
        adminModal: document.getElementById('adminModal'),
        adminModalTitle: document.getElementById('adminModalTitle'),
        adminModalBody: document.getElementById('adminModalBody'),
        closeAdminModal: document.getElementById('closeAdminModal'),
            cancelAdminModal: document.getElementById('cancelAdminModal'),
        confirmAdminModal: document.getElementById('confirmAdminModal'),
        addCategoryBtn: document.getElementById('addCategoryBtn'),
        addFlavorBtn: document.getElementById('addFlavorBtn'),
        addExtraBtn: document.getElementById('addExtraBtn'),
        addObsBtn: document.getElementById('addObsBtn'),
        // History
        historyTabs: document.querySelectorAll('.history-tab'),
        datePickerContainer: document.getElementById('datePickerContainer'),
        historyDatePicker: document.getElementById('historyDatePicker'),
        searchDateBtn: document.getElementById('searchDateBtn'),
        historyOrdersList: document.getElementById('historyOrdersList'),
        historyOrderModal: document.getElementById('historyOrderModal'),
        historyModalOverlay: document.getElementById('historyModalOverlay'),
        historyOrderDetail: document.getElementById('historyOrderDetail'),
        historyTicketContent: document.getElementById('historyTicketContent'),
        backToHistoryBtn: document.getElementById('backToHistoryBtn'),
        reprintOrderBtn: document.getElementById('reprintOrderBtn'),
        invoiceOrderBtn: document.getElementById('invoiceOrderBtn'),
        deleteOrderBtnHistory: document.getElementById('deleteOrderBtnHistory'),
        // Reports
        reportDatePicker: document.getElementById('reportDatePicker'),
        reportPeriodSelect: document.getElementById('reportPeriodSelect'),
        reportDatePickerGroup: document.getElementById('reportDatePickerGroup'),
        reportMonthPicker: document.getElementById('reportMonthPicker'),
        reportMonthPickerGroup: document.getElementById('reportMonthPickerGroup'),
        searchReportBtn: document.getElementById('searchReportBtn'),
        reportDailySales: document.getElementById('reportDailySales'),
        reportFoodSales: document.getElementById('reportFoodSales'),
        reportBebidasSales: document.getElementById('reportBebidasSales'),
        reportDesechablesSales: document.getElementById('reportDesechablesSales'),
        reportEfectivoSales: document.getElementById('reportEfectivoSales'),
        reportNequiSales: document.getElementById('reportNequiSales'),
        reportDaviplataSales: document.getElementById('reportDaviplataSales'),
        // Admin Security
        adminLoginModal: document.getElementById('adminLoginModal'),
        adminPasswordInput: document.getElementById('adminPasswordInput'),
        confirmAdminLogin: document.getElementById('confirmAdminLogin'),
        closeAdminLoginModal: document.getElementById('closeAdminLoginModal'),
        newAdminPassword: document.getElementById('newAdminPassword'),
        confirmAdminPassword: document.getElementById('confirmAdminPassword'),
        saveAdminPasswordBtn: document.getElementById('saveAdminPasswordBtn'),
        // History Summary
        historyTotalSales: document.getElementById('historyTotalSales'),
        historyTotalEfectivo: document.getElementById('historyTotalEfectivo'),
        historyTotalNequi: document.getElementById('historyTotalNequi'),
        historyTotalDaviplata: document.getElementById('historyTotalDaviplata'),
        historyTotalFood: document.getElementById('historyTotalFood'),
        historyTotalBebidas: document.getElementById('historyTotalBebidas'),
        historyTotalDesechables: document.getElementById('historyTotalDesechables'),
        // Detailed Reports
        categorySalesList: document.getElementById('categorySalesList'),
        extrasSalesList: document.getElementById('extrasSalesList'),
        flavorSalesList: document.getElementById('flavorSalesList'),
        sizeSalesList: document.getElementById('sizeSalesList'),
        categoryQtyList: document.getElementById('categoryQtyList'),
        // Report Detail Modal
        reportDetailModal: document.getElementById('reportDetailModal'),
        reportDetailList: document.getElementById('reportDetailList'),
        reportDetailTitle: document.getElementById('reportDetailTitle'),
        closeReportDetailModal: document.getElementById('closeReportDetailModal'),
        closeReportDetailModalOverlay: document.getElementById('closeReportDetailModalOverlay'),
        downloadReportBtn: document.getElementById('downloadReportBtn'),
        expenseSearchInput: document.getElementById('expenseSearchInput'),
    };

    let currentReportOrders = [];
    let lastSalesBreakdown = {};

    // ============================================
    // Navigation Drawer
    // ============================================

    if (elements.menuBtn) {
        elements.menuBtn.addEventListener('click', () => {
            elements.navDrawer.classList.add('open');
        });
    }

    if (elements.closeDrawer) {
        elements.closeDrawer.addEventListener('click', () => {
            elements.navDrawer.classList.remove('open');
        });
    }

    if (elements.drawerOverlay) {
        elements.drawerOverlay.addEventListener('click', () => {
            elements.navDrawer.classList.remove('open');
        });
    }

    elements.drawerItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const page = item.dataset.page;

            // Protection for Admin, History, and Reports pages
            const protectedPages = ['admin', 'history', 'reports', 'expenses', 'balance'];
            if (protectedPages.includes(page) && !state.isAdminAuthenticated) {
                if (elements.adminLoginModal) {
                    elements.adminLoginModal.classList.add('open');
                    elements.adminPasswordInput.value = '';
                    elements.adminPasswordInput.focus();
                    elements.navDrawer.classList.remove('open');
                    // Store the intended page for after login
                    state.pendingAdminPage = page;
                }
                return;
            }

            elements.drawerItems.forEach(i => i.classList.remove('active'));
            item.classList.add('active');

            elements.pages.forEach(p => p.classList.remove('active'));
            const targetPage = document.getElementById(`page-${page}`);
            if (targetPage) targetPage.classList.add('active');
            const mainContent = document.querySelector('.main-content');
            const mc = document.querySelector('.main-content'); if(mc) mc.scrollTop = 0;

            state.currentPage = page;
            elements.navDrawer.classList.remove('open');

            // Show/Hide footer based on page
            const appFooter = document.getElementById('appFooter');
            if (appFooter) {
                appFooter.style.display = (page === 'new-order') ? 'flex' : 'none';
            }

            if (page === 'kitchen') {
                renderKitchenPage();
            } else if (page === 'checkout') {
                renderCheckoutPage();
            } else if (page === 'orders') {
                renderOrdersPage();
            } else if (page === 'history') {
                renderHistoryPage();
            } else if (page === 'balance') {
                renderBalancePage();
            } else if (page === 'reports') {
                // Initialize report date to today
                if (elements.reportDatePicker && !elements.reportDatePicker.value) {
                    elements.reportDatePicker.value = new Date().toISOString().split('T')[0];
                }
                renderReportsPage();
            } else if (page === 'expenses') {
                renderExpensesPage();
            } else if (page === 'admin') {
                renderAdminPage();
            } else if (page === 'new-order') {
                state.appendingOrderId = null; // Clear if navigating manually to new order
                initializeCategories();
                refreshOrderPageUI();
            }

            // Remove flag if already redirected
            if (page === 'admin') state.pendingAdminRedirect = false;

            if (typeof lucide !== 'undefined') lucide.createIcons();
        });
    });

    function refreshOrderPageUI() {
        const locIn = document.getElementById('posLocationInput');
        const typeIn = document.getElementById('posOrderTypeInput');
        const config = StorageManager.getConfig();
        if (config.billingSystem === 'direct') {
            if (typeIn) typeIn.style.display = 'none';
            if (locIn) locIn.style.width = '100%';
        } else {
            if (typeIn) typeIn.style.display = '';
            if (locIn) locIn.style.width = '65%';
        }
        
        // Only re-enable inputs if we are NOT appending to an existing order
        if (!state.appendingOrderId) {
            if (locIn) {
                locIn.disabled = false;
                locIn.style.opacity = '1';
            }
            if (typeIn) {
                typeIn.disabled = false;
                typeIn.style.opacity = '1';
            }
        }

        const headerBtn = document.getElementById('posCancelAppendHeaderBtn');
        const sendBtnSpan = document.querySelector('#sendToKitchenBtn span');
        const confirmBtn = document.getElementById('confirmTicket');

        if (state.appendingOrderId) {
            if (headerBtn) headerBtn.style.display = 'flex';
            if (sendBtnSpan) sendBtnSpan.textContent = 'ADICIONAR';
            if (confirmBtn) confirmBtn.innerHTML = '<i data-lucide="send"></i> ADICIONAR';
        } else {
            if (headerBtn) headerBtn.style.display = 'none';
            if (sendBtnSpan) sendBtnSpan.textContent = 'GENERAR';
            if (confirmBtn) confirmBtn.innerHTML = '<i data-lucide="send"></i> GENERAR';
        }

        renderPosCategories();
        renderPosProducts();
        renderPosCart();
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ============================================
    // Multi-Sector POS Order Taking Engine
    // ============================================

    // Service Tabs Setup
    function setupServiceTabs() {
        document.querySelectorAll('.service-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.service-tab').forEach(t => t.classList.remove('active'));
                const service = tab.dataset.service;
                state.serviceType = service;
                document.querySelectorAll(`.service-tab[data-service="${service}"]`).forEach(t => t.classList.add('active'));
            });
        });
    }
    setupServiceTabs();

    // Search Input Setup
    if (elements.posProductSearch) {
        elements.posProductSearch.addEventListener('input', (e) => {
            state.searchQuery = e.target.value.trim();
            if (elements.posClearSearch) {
                elements.posClearSearch.classList.toggle('hidden', state.searchQuery === '');
            }
            renderPosProducts();
        });
    }

    if (elements.posClearSearch) {
        elements.posClearSearch.addEventListener('click', () => {
            if (elements.posProductSearch) elements.posProductSearch.value = '';
            state.searchQuery = '';
            elements.posClearSearch.classList.add('hidden');
            renderPosProducts();
            if (elements.posProductSearch) elements.posProductSearch.focus();
        });
    }

    // Get Active Products from Config or Fallback
    function getActiveProductsList(config) {
        if (config.products && config.products.length > 0) {
            return config.products.filter(p => p.active !== false);
        }
        const list = [];
        Object.keys(config.flavors || {}).forEach(catId => {
            (config.flavors[catId] || []).forEach(f => {
                if (f.active !== false) {
                    list.push({
                        id: f.id,
                        name: f.name,
                        price: f.price || 0,
                        category: catId,
                        icon: 'ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢',
                        active: true,
                        prodType: f.prodType || 'fixed'
                    });
                }
            });
        });
        return list;
    }

    // Render Categories Filter Chips
    function renderPosCategories() {
        const container = elements.posCategoriesBar || document.getElementById('posCategoriesBar');
        if (!container) return;

        const config = StorageManager.getConfig();
        const allProducts = getActiveProductsList(config);
        const totalCount = allProducts.length;

        let html = `
            <button type="button" class="pos-cat-chip ${state.selectedCategory === 'all' ? 'active' : ''}" data-cat="all">
                <span>Todos</span>
                <span class="pos-cat-chip-count">${totalCount}</span>
            </button>
        `;

        config.categories.forEach(cat => {
            const count = allProducts.filter(p => p.category === cat.id).length;
            const isActive = state.selectedCategory === cat.id ? 'active' : '';
            html += `
                <button type="button" class="pos-cat-chip ${isActive}" data-cat="${cat.id}">
                    <span>${cat.name}</span>
                    <span class="pos-cat-chip-count">${count}</span>
                </button>
            `;
        });

        container.innerHTML = html;

        container.querySelectorAll('.pos-cat-chip').forEach(btn => {
            btn.addEventListener('click', () => {
                state.selectedCategory = btn.dataset.cat;
                renderPosCategories();
                renderPosProducts();
            });
        });
    }

    // Render Products Grid
                    function renderPosProducts() { renderSplitUI(); }
    function renderPosCart() {
        const totalAmountFooter = document.getElementById('totalAmount');
        let grandTotal = 0;
        state.cart.forEach(item => {
            grandTotal += item.subtotal;
        });
        state.orderTotal = grandTotal;
        if (totalAmountFooter) totalAmountFooter.textContent = formatPrice(grandTotal);
    }
    // --- SPLIT UI LOGIC ---
function renderPosClientTabs() {
    const container = document.getElementById('posPeopleTabsContainer');
    const el = document.getElementById('posClientsTabs');
    if (!el) return;

    const config = StorageManager.getConfig();
    if (config.billingSystem === 'direct') {
        if (container) container.style.display = 'none';
        // Enforce only one client in direct mode
        if (state.clients.length > 1) {
            state.clients = ['P1'];
            state.activeClient = 'P1';
            state.cart = state.cart.filter(item => item.clientName === 'P1');
        }
    } else {
        if (container) container.style.display = 'flex';
    }

    el.innerHTML = state.clients.map(client => `
        <button type="button" class="pos-client-pill ${client === state.activeClient ? 'active' : ''}" style="display: flex; align-items: center; gap: 8px; padding-right: 8px;" onclick="window.switchClient('${client}')">
            <span>${client}</span>
            ${state.clients.length > 1 ? `<div onclick="window.removeClient(event, '${client}')" style="display: flex; align-items: center; justify-content: center; width: 20px; height: 20px; border-radius: 50%; background: rgba(0,0,0,0.1);"><i data-lucide="x" style="width: 12px; height: 12px;"></i></div>` : ''}
        </button>
    `).join('');
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

window.removeClient = function(e, client) {
    e.stopPropagation();
    if (state.clients.length <= 1) return;
    
    window.minesofConfirm(`Eliminar al cliente ${client} y todos sus productos seleccionados?`, () => {
        state.clients = state.clients.filter(c => c !== client);
        state.cart = state.cart.filter(item => item.clientName !== client);
        
        if (state.activeClient === client) {
            window.switchClient(state.clients[0]);
        } else {
            renderPosClientTabs();
            renderSplitUI();
            if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
        }
    });
};
window.switchClient = function(client) {
    state.activeClient = client;
    renderPosClientTabs();
    renderSplitUI();
};
    const btn = document.getElementById('btnAddClient');
    if (btn) {
        btn.addEventListener('click', () => {
                        let maxNum = 0;
            state.clients.forEach(c => {
                const num = parseInt(c.replace('P', '')) || 0;
                if (num > maxNum) maxNum = num;
            });
            const newClient = 'P' + (maxNum + 1);
            state.clients.push(newClient);
            window.switchClient(newClient);
        });
    }
    renderPosClientTabs();

    window.updateProductQty = function(productId, delta) {
        const clientId = state.activeClient;
        if (!clientId) return;
        const config = StorageManager.getConfig();
        const products = getActiveProductsList(config);
        const product = products.find(p => p.id === productId);
        if (!product) return;

        const existingIndex = state.cart.findIndex(item => item.productId === product.id && item.clientName === clientId);
        
        if (existingIndex !== -1) {
            let item = state.cart[existingIndex];
            item.qty += delta;
            if (item.qty <= 0) {
                state.cart.splice(existingIndex, 1);
            } else {
                item.subtotal = item.qty * item.unitPrice;
            }
        } else if (delta > 0) {
            state.cart.push({
                id: 'cart_' + Date.now(),
                productId: product.id,
                name: product.name,
                unitPrice: product.price || 0,
                qty: 1,
                subtotal: product.price || 0,
                clientName: clientId,
                categoryId: product.category,
                notes: ''
            });
        }
        renderSplitUI();
        if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
    };

    let activeQuantityProductId = null;
        let activeFixedObsProductId = null;
    let activeFixedObsEditIndex = -1;

    window.openFixedWithObsForProduct = function(product, editIndex = -1) {
        if (!product) return;
        activeFixedObsProductId = product.id;
        activeFixedObsEditIndex = editIndex;
        
        const modal = document.getElementById('fixedObsModal');
        const nameEl = document.getElementById('fixedObsModalProductName');
        const priceEl = document.getElementById('fixedObsModalProductPrice');
        const input = document.getElementById('fixedObsModalInput');
        const confirmBtn = document.getElementById('fixedObsConfirmBtn');
        const removeBtn = document.getElementById('fixedObsModalRemoveBtn');
        const titleEl = document.getElementById('fixedObsModalTitle');

        if (!modal) return;

        nameEl.textContent = product.name;
        priceEl.textContent = '$' + (product.price || 0).toLocaleString('es-CO');

        if (editIndex !== -1) {
            titleEl.innerHTML = '<i data-lucide="edit" style="width: 18px; height: 18px; color: #3b82f6;"></i> Editar Observaci&oacute;n';
            confirmBtn.innerHTML = '<i data-lucide="save" style="width: 16px; height: 16px;"></i> Guardar';
            input.value = state.cart[editIndex].notes || '';
            removeBtn.style.display = 'block';
        } else {
            titleEl.innerHTML = '<i data-lucide="plus-circle" style="width: 18px; height: 18px; color: #10b981;"></i> Agregar Producto';
            confirmBtn.innerHTML = '<i data-lucide="plus" style="width: 16px; height: 16px;"></i> Agregar';
            input.value = '';
            removeBtn.style.display = 'none';
        }

        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => input.focus(), 100);
    };

    const fixedObsConfirmBtn = document.getElementById('fixedObsConfirmBtn');
    if (fixedObsConfirmBtn) {
        fixedObsConfirmBtn.addEventListener('click', () => {
            const config = StorageManager.getConfig();
            const product = getActiveProductsList(config).find(p => p.id === activeFixedObsProductId);
            if (!product) return;

            const input = document.getElementById('fixedObsModalInput');
            const notes = input.value.trim();
            const clientId = state.activeClient;

            if (activeFixedObsEditIndex !== -1) {
                state.cart[activeFixedObsEditIndex].notes = notes;
            } else {
                state.cart.push({
                    id: 'cart_' + Date.now(),
                    productId: product.id,
                    name: product.name,
                    price: product.price || 0,
                    unitPrice: product.price || 0,
                    qty: 1,
                    subtotal: product.price || 0,
                    clientName: clientId,
                    categoryId: product.category,
                    notes: notes
                });
            }

            document.getElementById('fixedObsModal').classList.remove('open');
            renderSplitUI();
            if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
        });
    }

    const fixedObsModalRemoveBtn = document.getElementById('fixedObsModalRemoveBtn');
    if (fixedObsModalRemoveBtn) {
        fixedObsModalRemoveBtn.addEventListener('click', () => {
            if (activeFixedObsEditIndex !== -1) {
                state.cart.splice(activeFixedObsEditIndex, 1);
            }
            document.getElementById('fixedObsModal').classList.remove('open');
            renderSplitUI();
            if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
        });
    }

    window.openQuantityForProduct = function(product, editIndex = -1) {
        if (!product) return;
        activeQuantityProductId = product.id;
        const input = document.getElementById('quantityInputValue');
        const modal = document.getElementById('quantityModal');
        const label = document.getElementById('quantityModalProductLabel');
        
        if (!input || !modal) {
            if(typeof showNotification === 'function') showNotification('Por favor, cierra sesiÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n y recarga la pÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡gina para actualizar', 'error');
            return;
        }

        if (label) {
            label.textContent = product.name.toUpperCase();
        }
        
        if (editIndex !== -1) {
            input.value = state.cart[editIndex].qty || 1;
        } else {
            input.value = '1';
        }
        
        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => input.select(), 100);
    };

    let activeOpenPriceProductId = null;
    window.openPriceForProduct = function(product, editIndex = -1) {
        if (!product) return;
        activeOpenPriceProductId = product.id;
        const input = document.getElementById('openPriceInput');
        const modal = document.getElementById('openPriceModal');
        const label = document.getElementById('openPriceModalProductLabel');
        
        if (!input || !modal) {
            if(typeof showNotification === 'function') showNotification('Por favor, cierra sesiÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n y recarga la pÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡gina para actualizar', 'error');
            return;
        }

        if (label) {
            label.textContent = product.name.toUpperCase();
        }
        
        if (editIndex !== -1) {
            input.value = state.cart[editIndex].unitPrice.toLocaleString('es-CO');
        } else {
            input.value = '';
        }
        
        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => input.focus(), 100);
    };

    let activeTextProductId = null;
    window.openTextForProduct = function(product, editIndex = -1) {
        if (!product) return;
        activeTextProductId = product.id;
        const input = document.getElementById('textInputValue');
        const modal = document.getElementById('textInputModal');
        const label = document.getElementById('textInputModalProductLabel');

        if (!input || !modal) {
            if(typeof showNotification === 'function') showNotification('Por favor, cierra sesiÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n y recarga la pÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡gina para actualizar', 'error');
            return;
        }

        if (label) {
            label.textContent = product.name.toUpperCase();
        }
        
        if (editIndex !== -1) {
            input.value = state.cart[editIndex].notes || '';
        } else {
            input.value = '';
        }
        
        modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => input.focus(), 100);
    };

    window.triggerToggleProduct = function(productId) {
        const config = StorageManager.getConfig();
        const products = getActiveProductsList(config);
        const product = products.find(p => p.id === productId);
        if (!product) return;

        const prodType = product.prodType || 'fixed';
        const clientId = state.activeClient;

        // If it's already in the cart, edit it or toggle it off
        const existingIndex = state.cart.findIndex(item => item.productId === product.id && item.clientName === clientId);
        if (existingIndex !== -1) {
            if (prodType === 'open_price') {
                window.openPriceForProduct(product, existingIndex);
                return;
            } else if (prodType === 'text') {
                window.openTextForProduct(product, existingIndex);
                return;
            } else if (prodType === 'quantity') {
                window.openQuantityForProduct(product, existingIndex);
                return;
            } else if (prodType === 'fixed_with_obs') {
                window.openFixedWithObsForProduct(product, existingIndex);
                return;
            }
            // For fixed products, simply remove from cart
            state.cart.splice(existingIndex, 1);
            renderSplitUI();
            if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
            return;
        }

        if (prodType === 'open_price') {
            window.openPriceForProduct(product);
        } else if (prodType === 'text') {
            window.openTextForProduct(product);
        } else if (prodType === 'quantity') {
            window.openQuantityForProduct(product);
        } else if (prodType === 'fixed_with_obs') {
            window.openFixedWithObsForProduct(product);
        } else {
            toggleProduct(product);
        }
    };
    
    function toggleProduct(product) {
        const clientId = state.activeClient;
        if (!clientId) return;

        const existingIndex = state.cart.findIndex(item => item.productId === product.id && item.clientName === clientId);
        
        let isActiveNow = false;
        if (existingIndex !== -1) {
            state.cart.splice(existingIndex, 1);
        } else {
            state.cart.push({
                id: 'cart_' + Date.now(),
                productId: product.id,
                name: product.name,
                unitPrice: product.price || 0,
                qty: 1,
                subtotal: product.price || 0,
                clientName: clientId,
                categoryId: product.category,
                notes: ''
            });
            isActiveNow = true;
        }
        
        // Optimistic UI update instead of full render
        const cardEls = document.querySelectorAll(`.split-card[data-id="${product.id}"]`);
        cardEls.forEach(el => {
            if (isActiveNow) {
                el.classList.add('active');
            } else {
                el.classList.remove('active');
            }
            
        });
        
        renderPosCart();
    }

    window.clearCategorySearch = function(catId) {
        const inputEl = document.getElementById('search-input-' + catId);
        if (inputEl) {
            inputEl.value = '';
            window.filterCategory(inputEl, catId);
        }
    };

    window.filterCategory = function(input, catId) {
        const term = input.value.toLowerCase().trim();
        const clearBtn = document.getElementById('clear-search-' + catId);
        if (clearBtn) clearBtn.style.display = term ? 'flex' : 'none';
        
        const container = document.getElementById('col-content-' + catId);
        if (!container) return;
        
        const cards = container.querySelectorAll('.split-card');
        cards.forEach(card => {
            const name = card.dataset.name || '';
            if (name.includes(term)) {
                card.style.display = 'flex';
            } else {
                card.style.display = 'none';
            }
        });
    };

function renderSplitUI() {
    const container = document.getElementById('dynamicCategoriesContainer');
    if (!container) return;

    const config = StorageManager.getConfig();
    const products = getActiveProductsList(config);
    
    const groups = {};
    products.forEach(p => {
        const catId = p.category || 'otros';
        if(!groups[catId]) groups[catId] = [];
        groups[catId].push(p);
    });
    
    let html = '';
    
    const palette = [
        { main: '#2563eb', border: '#1e3a8a', bg: 'rgba(37,99,235,0.08)' },   // Blue
        { main: '#ea580c', border: '#7c2d12', bg: 'rgba(234,88,12,0.08)' },   // Orange
        { main: '#16a34a', border: '#14532d', bg: 'rgba(22,163,74,0.08)' },   // Green
        { main: '#9333ea', border: '#581c87', bg: 'rgba(147,51,234,0.08)' },  // Purple
        { main: '#eab308', border: '#713f12', bg: 'rgba(234,179,8,0.08)' }    // Yellow
    ];
    let colorIndex = 0;
    
    const renderColumn = (catId, catName) => {
        const items = groups[catId] || [];
        
        const colors = palette[colorIndex % palette.length];
        colorIndex++;
        
        const nameLower = catName.toLowerCase();

        let colHtml = `<div class="category-col-wrapper" style="height: 100%; display: flex; flex-direction: column; flex: 1; min-width: 17ch;"><div class="category-col" style="width: 100%; flex: 0 1 auto; display: flex; flex-direction: column; gap: 12px; max-height: 100%; min-height: 0; 
            background: ${colors.bg}; border: 1px solid ${colors.border}; border-radius: 16px; padding: 10px;
            --active-check: ${colors.main}; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
            
            <div style="background: ${colors.main}; border-radius: 8px; padding: 4px; text-align: center; font-weight: 800; color: white;  font-size: 0.9rem; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.2);">
                ${catName}
            </div>
            
            <div style="display: flex; align-items: center; background: var(--bg-card); border: 1px solid var(--border-subtle); border-radius: 8px; padding: 6px 12px; min-height: 40px; position: relative;">
                <i data-lucide="search" style="width: 16px; height: 16px; color: #94a3b8;"></i>
                <input type="text" autocomplete="off" id="search-input-${catId}" placeholder="Buscar..." oninput="window.filterCategory(this, '${catId}')" style="border: none; outline: none; width: 100%; padding: 4px; font-size: 0.85rem; margin-left: 8px; padding-right: 24px; background: transparent; color: var(--text-primary);">
                <div id="clear-search-${catId}" onclick="window.clearCategorySearch('${catId}')" style="display: none; position: absolute; right: 8px; cursor: pointer; padding: 4px; border-radius: 50%; background: rgba(0,0,0,0.1); align-items: center; justify-content: center;">
                    <i data-lucide="x" style="width: 14px; height: 14px; color: var(--text-primary);"></i>
                </div>
            </div>
            
            <div class="category-col-content" id="col-content-${catId}" style="display: flex; flex-direction: column; gap: 8px; overflow-y: auto; flex: 1; min-height: 0; padding-bottom: 20px; padding-right: 4px; scrollbar-width: none;">`;
        
        if (items.length === 0) {
            colHtml += `<div style="text-align: center; padding: 20px; color: #94a3b8; font-size: 0.85rem;">No hay productos</div>`;
        } else {
            colHtml += items.map(p => {
                const prodType = p.prodType || 'fixed';
                const isActive = state.cart.some(item => item.productId === p.id && item.clientName === state.activeClient);
                
                if (prodType === 'quantity' || prodType === 'open_price' || prodType === 'text') {
                    const cartItem = state.cart.find(item => item.productId === p.id && item.clientName === state.activeClient);
                    const isItemActive = !!cartItem;
                    let valueDisplay = '';
                    if (cartItem) {
                        if (prodType === 'open_price') valueDisplay = `$${cartItem.unitPrice.toLocaleString('es-CO')}`;
                        else if (prodType === 'text') valueDisplay = cartItem.notes;
                        else if (prodType === 'quantity') valueDisplay = `${cartItem.qty} unds`;
                    }
                    
                    if (isItemActive) {
                        return `<div class="split-card dynamic-card active" data-id="${p.id}" data-name="${p.name.toLowerCase()}" onclick="window.triggerToggleProduct('${p.id}')"
                            style="border: none; border-radius: 8px; padding: 2px 4px; font-size: 0.9rem; font-weight: 600; display: flex; flex-direction: column; gap: 6px; flex-shrink: 0;">
                                <div style="display: flex; justify-content: center; align-items: center; width: 100%;">
                                    <span style="text-align: center;">${p.name}</span>
                                </div>
                                <div style="background: white; border-radius: 6px; padding: 4px 8px; font-size: 0.85rem; color: #334155; text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; box-shadow: inset 0 1px 3px rgba(0,0,0,0.1);">
                                    ${valueDisplay}
                                </div>
                            </div>`;
                    } else {
                        return `<div class="split-card dynamic-card" data-id="${p.id}" data-name="${p.name.toLowerCase()}" onclick="window.triggerToggleProduct('${p.id}')"
                            style="border: none; border-radius: 8px; padding: 2px 4px; font-size: 0.9rem; font-weight: 600; display: flex; justify-content: center; align-items: center; flex-shrink: 0;">
                               <span style="text-align: center;">${p.name}</span>
                            </div>`;
                    }
                } else {
                    return `<div class="split-card dynamic-card ${isActive ? 'active' : ''}" data-id="${p.id}" data-name="${p.name.toLowerCase()}" onclick="window.triggerToggleProduct('${p.id}')"
                        style="border: none; border-radius: 8px; padding: 2px 4px; font-size: 0.9rem; font-weight: 600; display: flex; justify-content: center; text-align: center; align-items: center; flex-shrink: 0;">
                           <span>${p.name}</span>
                           </div>`;
                }
            }).join('');
        }
        
        colHtml += `</div></div></div>`;
        return colHtml;
    };

    config.categories.forEach(cat => {
        html += renderColumn(cat.id, cat.name);
        delete groups[cat.id];
    });
    
    const remaining = Object.keys(groups);
    if (remaining.length > 0) {
        html += renderColumn('otros', 'Otros');
    }

    if (html === '') {
        html = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; width: 100%; height: 100%; color: #94a3b8; text-align: center; padding: 40px;">
            <i data-lucide="package-open" style="width: 64px; height: 64px; margin-bottom: 16px; opacity: 0.5;"></i>
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--text-primary); margin-bottom: 8px;">Todo listo para empezar!</h3>
            <p style="font-size: 0.95rem; max-width: 300px;">Aun no tienes el menu configurado. Crea tus categorias y productos en la opcion <b style="color: var(--accent-primary);">Administrador</b>.</p>
        </div>`;
    }

    container.innerHTML = html;
    if (typeof lucide !== 'undefined') lucide.createIcons();
}


    // Expose Cart methods
    // Expose Cart methods to window for inline onclick handlers
    function clearPosCart(confirmClear = false) {
        const performClear = () => {
            state.cart = [];
            state.orderTotal = 0;
            renderSplitUI();
        };
        if (confirmClear) {
            window.minesofConfirm('Estas seguro de vaciar todo el pedido actual?', performClear);
        } else {
            performClear();
        }
    }
    // Remove the invalid window assignments that throw ReferenceError

    // Item Note Modal Logic
    window.openItemNoteModal = function (cartItemId) {
        const item = state.cart.find(i => i.id === cartItemId);
        if (!item) return;
        state.editingNoteItemId = cartItemId;

        const modal = elements.itemNoteModal || document.getElementById('itemNoteModal');
        const prodNameEl = document.getElementById('itemNoteModalProductName');
        const inputEl = document.getElementById('itemNoteModalInput');
        const tagsContainer = document.getElementById('itemNoteQuickTags');

        if (prodNameEl) prodNameEl.textContent = `${item.qty}x ${item.name}`;
        if (inputEl) inputEl.value = item.notes || '';

        if (tagsContainer) {
            const config = StorageManager.getConfig();
            const obs = (config.observations && config.observations[item.category]) || [];
            tagsContainer.innerHTML = obs.map(o => `
                <span class="quick-obs-chip" onclick="window.appendQuickTag('${o.name}')">${o.name}</span>
            `).join('');
        }

        if (modal) modal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
        if (inputEl) inputEl.focus();
    };

    window.appendQuickTag = function (tagName) {
        const inputEl = document.getElementById('itemNoteModalInput');
        if (!inputEl) return;
        if (inputEl.value.trim() === '') {
            inputEl.value = tagName;
        } else {
            inputEl.value += ', ' + tagName;
        }
    };

    function saveItemNoteModal() {
        if (!state.editingNoteItemId) return;
        const item = state.cart.find(i => i.id === state.editingNoteItemId);
        const inputEl = document.getElementById('itemNoteModalInput');
        if (item && inputEl) {
            item.notes = inputEl.value.trim();
        }
        const modal = elements.itemNoteModal || document.getElementById('itemNoteModal');
        if (modal) modal.classList.remove('open');
        state.editingNoteItemId = null;
        renderPosCart();
    }

    function closeItemNoteModalFunc() {
        const modal = elements.itemNoteModal || document.getElementById('itemNoteModal');
        if (modal) modal.classList.remove('open');
        state.editingNoteItemId = null;
    }

    if (elements.closeItemNoteModal) elements.closeItemNoteModal.addEventListener('click', closeItemNoteModalFunc);
    if (elements.closeItemNoteOverlay) elements.closeItemNoteOverlay.addEventListener('click', closeItemNoteModalFunc);
    if (elements.cancelItemNoteModal) elements.cancelItemNoteModal.addEventListener('click', closeItemNoteModalFunc);
    if (elements.saveItemNoteModal) elements.saveItemNoteModal.addEventListener('click', saveItemNoteModal);

    // Wire Clear Cart Button
    if (elements.posClearCartBtn) {
        elements.posClearCartBtn.addEventListener('click', () => clearPosCart(true));
    }

    // Submit Order (Send to Kitchen & Create Ticket)
    let pendingOrder = null;

    async function submitOrder() {
        if (state.cart.length === 0) {
            showNotification('Atenci&oacute;n: Agrega productos al pedido antes de enviar', 'error');
            return;
        }

        const typeInput = document.getElementById('posOrderTypeInput');
        const locationInput = document.getElementById('posLocationInput');
        const config = StorageManager.getConfig();
        const typeText = (typeInput && config.billingSystem !== 'direct') ? typeInput.value.toUpperCase() : '';
        const rawLocText = locationInput ? locationInput.value.trim().toUpperCase() : '';
        const locationText = (typeText && rawLocText) ? typeText + " - " + rawLocText : (rawLocText || typeText);
        const customerText = ''; // Removed clients suffix as requested

        if (!state.serviceType) state.serviceType = 'salon';

        if (!state.appendingOrderId && state.clients.length === 0 && !locationText) {
            showNotification('Atenci&oacute;n: Ingresa al menos un cliente en la orden', 'error');
            return;
        }

        const submitBtn = elements.posSubmitOrderBtn || elements.sendToKitchenBtn || document.getElementById('posSubmitOrderBtn');
        const origText = submitBtn ? submitBtn.innerHTML : '';
        if (submitBtn) {
            submitBtn.innerHTML = '<i data-lucide="loader-2" class="animate-spin"></i> Generando...';
            submitBtn.disabled = true;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        try {
            const config = StorageManager.getConfig();
            const items = state.cart.map(i => {
                const catInfo = config.categories.find(c => c.id === i.category);
                return {
                    id: i.id,
                    productId: i.productId,
                    name: i.name,
                    category: i.category,
                    categoryName: catInfo ? catInfo.name : i.category,
                    unitPrice: i.unitPrice,
                    qty: i.qty,
                    notes: i.notes,
                    observations: i.notes,
                    extras: i.extras || [],
                    price: i.subtotal,
                    clientName: i.clientName
                };
            });

            if (state.appendingOrderId) {
                const originalOrder = StorageManager.getOrders().find(o => o.id == state.appendingOrderId);
                if (originalOrder) {
                    const updatedItems = [...originalOrder.items, ...items];
                    const updatedTotalPrice = updatedItems.reduce((sum, item) => sum + item.price, 0);
                    StorageManager.updateOrder(originalOrder.id, {
                        items: updatedItems,
                        totalPrice: updatedTotalPrice
                    });

                    // Create partial order for printing
                    const partialOrder = {
                        orderNumber: `${originalOrder.orderNumber} (ADI)`,
                        sequenceNumber: originalOrder.sequenceNumber,
                        serviceType: state.serviceType,
                        customerInfo: originalOrder.customerInfo,
                        createdAt: new Date().toISOString(),
                        items: items,
                        totalPrice: items.reduce((s, i) => s + i.price, 0),
                        isAppending: true,
                        isPartial: true,
                        checkoutPrinted: false,
                        paid: false
                    };
                    if (config.billingSystem === 'direct') {
                        partialOrder.checkoutPrinted = true;
                    }
                    StorageManager.addOrder(partialOrder);
                    
                    if (config.billingSystem === 'direct') {
                        showNotification(`Adición agregada y lista para cobro`);
                        setTimeout(() => window.openPaymentModal(partialOrder.id), 150);
                    } else {
                        showNotification(`Adición agregada al pedido ${originalOrder.orderNumber}`);
                    }
                }
                state.appendingOrderId = null;
                        } else {
                const seqNum = await generateOrderNumber();
                
                let orderIdentifier = seqNum;
                if (locationText && customerText) {
                    orderIdentifier = `${locationText} | ${customerText}`;
                } else if (locationText) {
                    orderIdentifier = locationText;
                } else if (customerText) {
                    orderIdentifier = customerText;
                } else {
                    orderIdentifier = 'CLIENTE';
                }

                const newOrder = {
                    orderNumber: orderIdentifier,
                    sequenceNumber: seqNum,
                    serviceType: state.serviceType,
                    customerInfo: orderIdentifier,
                    customerName: orderIdentifier,
                    items: items,
                    status: 'pending',
                    totalPrice: items.reduce((sum, item) => sum + item.price, 0),
                    createdBy: localStorage.getItem('minesof_deviceUser') || 'Cajero 1',
                    needsPrint: true,
                    printed: false,
                    checkoutPrinted: config.billingSystem === 'direct' ? true : false,
                    isAppending: false,
                    createdAt: new Date().toISOString()
                };

                StorageManager.addOrder(newOrder);
                
                if (config.billingSystem === 'direct') {
                    showNotification('Pedido ' + newOrder.orderNumber + ' listo para cobro');
                    setTimeout(() => {
                        window.openPaymentModal(newOrder.id);
                    }, 150);
                } else {
                    showNotification('Pedido ' + newOrder.orderNumber + ' generado');
                }
            }

            clearPosCart(false);
            if (locationInput) {
                locationInput.value = '';
                locationInput.disabled = false;
                locationInput.style.opacity = '1';
            }
            if (typeInput) {
                typeInput.value = 'Local';
                typeInput.disabled = false;
                typeInput.style.opacity = '1';
            }
            state.clients = ['P1'];
            state.activeClient = 'P1';
            renderPosClientTabs();
            refreshOrderPageUI();
        } catch (err) {
            console.error('Error submitting order:', err);
            showNotification('ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¯ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â Error al procesar pedido', 'error');
        } finally {
            if (submitBtn) {
                submitBtn.disabled = false;
                // ensure correct text based on append state
                if (state.appendingOrderId) {
                    submitBtn.innerHTML = '<i data-lucide="send"></i> <span>ADICIONAR</span>';
                } else {
                    submitBtn.innerHTML = '<i data-lucide="send"></i> <span>GENERAR</span>';
                }
                document.getElementById('posCancelAppendHeaderBtn').style.display = 'none';
                if (typeof lucide !== 'undefined') lucide.createIcons();
            }
        }
    }

    if (elements.posSubmitOrderBtn) elements.posSubmitOrderBtn.addEventListener('click', submitOrder);
    
    const cancelAppendBtn = document.getElementById('posCancelAppendHeaderBtn');
    if (cancelAppendBtn) {
        cancelAppendBtn.addEventListener('click', () => {
            state.appendingOrderId = null;
            clearPosCart(false);
            state.clients = ['P1'];
            state.activeClient = 'P1';
            if (typeof renderPosClientTabs === 'function') renderPosClientTabs();
            const locIn = document.getElementById('posLocationInput');
            const typeIn = document.getElementById('posOrderTypeInput');
            if (locIn) {
                locIn.value = '';
                locIn.disabled = false;
                locIn.style.opacity = '1';
            }
            if (typeIn) {
                typeIn.value = 'Local';
                typeIn.disabled = false;
                typeIn.style.opacity = '1';
            }
            refreshOrderPageUI();
        });
    }
    if (elements.sendToKitchenBtn) elements.sendToKitchenBtn.addEventListener('click', submitOrder);

    // Initial render of POS workspace
    refreshOrderPageUI();
    renderPosCategories();
    renderPosProducts();
    renderPosCart();

    function showTicketModal(order) {
        if (!elements.ticketModal) return;
        elements.ticketContent.innerHTML = generateTicketText(order);
        elements.ticketModal.classList.add('open');
    }

    function closeTicketModal() {
        if (elements.ticketModal) {
            elements.ticketModal.classList.remove('open');
            pendingOrder = null;
        }
    }

    if (elements.cancelTicket) elements.cancelTicket.addEventListener('click', closeTicketModal);
    if (elements.cancelTicketFooter) elements.cancelTicketFooter.addEventListener('click', closeTicketModal);
    if (elements.closeTicketModal) elements.closeTicketModal.addEventListener('click', closeTicketModal);

    if (elements.printTicket) {
        elements.printTicket.addEventListener('click', () => {
            if (pendingOrder) {
                if (pendingOrder.isAppending) {
                    const originalOrder = StorageManager.getOrders().find(o => o.id == pendingOrder.id);
                    if (originalOrder) {
                        const updatedItems = [...originalOrder.items, ...pendingOrder.newItems];
                        const updatedTotalPrice = updatedItems.reduce((sum, item) => sum + item.price, 0);
                        StorageManager.updateOrder(pendingOrder.id, {
                            items: updatedItems,
                            totalPrice: updatedTotalPrice,
                            checkoutPrinted: false
                        });
                    }
                    state.appendingOrderId = null;
                } else {
                    pendingOrder.printed = true;
                    StorageManager.addOrder(pendingOrder);
                }
                window.print();
                showNotification(`Pedido ${pendingOrder.orderNumber} impreso y enviado`);
                closeTicketModal();
                resetAllCategories();
                clearPosCart(false);
                refreshOrderPageUI();
            }
        });
    }

    if (elements.confirmTicket) {
        elements.confirmTicket.addEventListener('click', () => {
            if (pendingOrder) {
                if (pendingOrder.isAppending) {
                    const originalOrder = StorageManager.getOrders().find(o => o.id == pendingOrder.id);
                    if (originalOrder) {
                        const updatedItems = [...originalOrder.items, ...pendingOrder.newItems];
                        const updatedTotalPrice = updatedItems.reduce((sum, item) => sum + item.price, 0);
                        StorageManager.updateOrder(pendingOrder.id, {
                            items: updatedItems,
                            totalPrice: updatedTotalPrice,
                            checkoutPrinted: false
                        });
                    }
                    state.appendingOrderId = null;
                } else {
                    StorageManager.addOrder(pendingOrder);
                }
                showNotification(`Pedido ${pendingOrder.orderNumber} enviado a cocina`);
                closeTicketModal();
                resetAllCategories();
                clearPosCart(false);
                refreshOrderPageUI();
            }
        });
    }

    function resetAllCategories() {
        Object.keys(state.categoryData).forEach(category => {
            state.categoryData[category].rows = [];
            const section = document.querySelector(`.category-section[data-category="${category}"]`);
            if (section) {
                const container = section.querySelector('.category-rows-container');
                container.innerHTML = '';

                // Reset category total to $0
                const priceEl = section.querySelector('.category-total-price');
                if (priceEl) {
                    priceEl.textContent = '$0';
                    priceEl.dataset.value = '0';
                }
            }
        });

        // Reset Service Type
        state.serviceType = 'salon';
        elements.serviceTabs.forEach(tab => {
            tab.classList.remove('active');
            if (tab.dataset.service === 'salon') tab.classList.add('active');
        });

        // Reset order total
        state.orderTotal = 0;
        if(typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
    }
    



    // ============================================
    // Checkout / Payment
    // ============================================

    let selectedPaymentOrder = null;

    function renderCheckoutPage() {
        const orders = StorageManager.getOrders();
        const config = StorageManager.getConfig();
        if (config.billingSystem === 'direct') {
            checkoutMode = 'paid';
            document.querySelectorAll('.checkout-tab').forEach(t => t.classList.remove('active'));
            const paidTab = document.querySelector('.checkout-tab[data-tab="paid"]');
            if (paidTab) paidTab.classList.add('active');
        }

        // Filter logic:
        // to-print: Not paid AND NOT printed for checkout
        // pending: Not paid AND printed for checkout
        // paid: Paid
        const today = new Date().toDateString();
        const toPrint = orders.filter(o => o.paid !== true && o.checkoutPrinted !== true);
        const pending = orders.filter(o => o.paid !== true && o.checkoutPrinted === true);
        const paid = orders.filter(o => o.paid === true && new Date(o.createdAt).toDateString() === today);

        if (elements.toPrintCount) elements.toPrintCount.textContent = toPrint.length;
        if (elements.pendingPaymentCount) elements.pendingPaymentCount.textContent = pending.length;
        if (elements.paidOrdersCount) elements.paidOrdersCount.textContent = paid.length;

        const emptyStateHTML = (msg, submsg) => `
            <div style="grid-column: 1 / -1; width: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 60px 20px; color: #94a3b8; text-align: center;">
                <i data-lucide="inbox" style="width: 64px; height: 64px; margin-bottom: 16px; opacity: 0.5;"></i>
                <h3 style="font-size: 1.1rem; font-weight: 600; color: var(--text-primary); margin-bottom: 8px;">${msg}</h3>
                <p style="font-size: 0.9rem; max-width: 250px; margin: 0;">${submsg}</p>
            </div>`;

        if (elements.toPrintList) {
            elements.toPrintList.innerHTML = toPrint.length > 0 
                ? toPrint.reverse().map(o => safeCreateCheckoutCard(o)).join('') 
                : emptyStateHTML('No hay pedidos pendientes', 'Los pedidos activos apareceran aqui.');
        }
        
        if (elements.pendingPaymentList) {
            elements.pendingPaymentList.innerHTML = pending.length > 0 
                ? pending.reverse().map(o => safeCreateCheckoutCard(o)).join('') 
                : emptyStateHTML('No hay pedidos por cobrar', 'No hay pedidos esperando por cobrar en caja.');
        }
        
        if (elements.paidOrdersList) {
            elements.paidOrdersList.innerHTML = paid.length > 0 
                ? paid.reverse().map(o => safeCreateCheckoutCard(o)).join('') 
                : emptyStateHTML('No hay ventas cobradas hoy', 'Los pedidos pagados apareceran aqui.');
        }
        
        if (typeof lucide !== 'undefined') lucide.createIcons();

        // Visibility toggle
        const lists = {
            'to-print': elements.toPrintList,
            'pending': elements.pendingPaymentList,
            'paid': elements.paidOrdersList
        };

        Object.keys(lists).forEach(mode => {
            if (lists[mode]) {
                if (mode === checkoutMode) {
                    lists[mode].classList.remove('hidden');
                } else {
                    lists[mode].classList.add('hidden');
                }
            }
        });

        document.querySelectorAll('.order-list-card[data-order-id]').forEach(card => {
            card.addEventListener('click', () => {
                window.openPaymentModal(card.dataset.orderId);
            });
        });

        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ============================================
    // Kitchen (KDS)
    // ============================================

    function renderKitchenPage() {
        const orders = StorageManager.getActiveOrders();

        const pending = orders.filter(o => o.status === 'pending');
        const preparing = orders.filter(o => o.status === 'preparing');
        const ready = orders.filter(o => o.status === 'ready');

        const updateColumn = (listId, countId, items, action) => {
            const list = document.getElementById(listId);
            const count = document.getElementById(countId);
            if (count) count.textContent = items.length;
            if (list) {
                list.innerHTML = items.map(o => `
                    <div class="kitchen-card">
                        <div class="kitchen-card-header">
                            <span class="kitchen-order-number">${o.orderNumber} ${o.sequenceNumber && o.sequenceNumber !== o.orderNumber ? `(${o.sequenceNumber})` : ''}</span>
                            <span class="kitchen-time">${o.customerInfo}</span>
                        </div>
                        <div class="kitchen-items">
                            ${(() => {
                                const kItemsByClient = {};
                                o.items.forEach(item => {
                                    const cName = item.clientName || 'CLIENTE';
                                    if (!kItemsByClient[cName]) kItemsByClient[cName] = [];
                                    kItemsByClient[cName].push(item);
                                });
                                return Object.entries(kItemsByClient).map(([cName, cItems]) => `
                                    <div style="font-size: 0.85rem; font-weight: bold; color: var(--accent-gold); margin: 6px 0 2px 0;">
                                        Cliente: ${cName}
                                    </div>
                                    ${cItems.map(item => `
                                        <div class="k-item">
                                            <strong>${item.qty}x</strong> ${item.name || item.categoryName} ${item.size ? item.size : ''}
                                            ${item.notes && item.notes !== item.name ? `<div style="font-size:0.8rem; color:#f0c040; margin-left:14px;">* ${item.notes}</div>` : ''}
                                            ${item.extras && item.extras.length > 0 ? `<div style="font-size:0.8rem; color:#4ecdc4; margin-left:14px;">+ ${(Array.isArray(item.extras) ? item.extras.map(e => typeof e === 'object' ? e.name : e).join(', ') : item.extras)}</div>` : ''}
                                        </div>
                                    `).join('')}
                                `).join('');
                            })()}
                        </div>
                        <button class="k-action-btn" onclick="window.advanceOrder('${o.id}')">${action}</button>
                    </div>
                `).join('');
            }
        };

        updateColumn('listPending', 'countPending', pending, 'EMPEZAR');
        updateColumn('listPreparing', 'countPreparing', preparing, 'LISTO');
        updateColumn('listReady', 'countReady', ready, 'ENTREGAR');

        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    window.advanceOrder = function (id) {
        const orders = StorageManager.getOrders();
        const order = orders.find(o => o.id === id);
        if (!order) return;

        const nextStatus = {
            'pending': 'preparing',
            'preparing': 'ready',
            'ready': 'delivered'
        };

        const newStatus = nextStatus[order.status];
        if (newStatus) {
            StorageManager.updateOrder(id, { status: newStatus });
            renderKitchenPage();
            showNotification(`Orden ${order.orderNumber} movida a ${newStatus}`);
        }
    };

    function safeCreateCheckoutCard(o) {
    try {
        return createCheckoutCard(o);
    } catch(e) {
        console.error("Error creating card for order", o, e);
        return '<div style="color:red; padding:10px; border:1px solid red;">Error renderizando orden ' + (o ? o.orderNumber : 'null') + ': ' + e.message + '</div>';
    }
}
function createCheckoutCard(order) {
        const labels = { pending: 'Pendiente', preparing: 'Preparando', ready: 'Listo', delivered: 'Entregado' };
        const titleName = order.createdBy ? order.createdBy.toUpperCase() : 'CAJA';
        const titleSeq = order.sequenceNumber ? `(${order.sequenceNumber})` : `(${order.orderNumber})`;
        
        return `
            <div class="order-list-card ${order.paid ? 'paid' : ''}" data-order-id="${order.id}">
                <div class="order-card-header">
                    <span class="order-number">${titleName} ${titleSeq}</span>
                    <span class="order-status-badge">${order.paid ? 'Pagado' : labels[order.status]}</span>
                </div>
                <div class="order-customer-info">
                    <span class="order-time">${new Date(order.createdAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                    <span> - ${order.customerInfo}</span>
                </div>
                <div class="order-items-preview">
                    ${[...(order.items || [])].sort((a,b) => (a.clientName || "Z").localeCompare(b.clientName || "Z")).map((item, i, arr) => `
                        ${i > 0 && item.clientName !== arr[i-1].clientName ? '<div style="height: 12px; width: 100%;"></div>' : ''}
                        <div class="preview-item">
                            <div class="item-main">
                                <span class="preview-qty">${item.clientName || item.qty}</span>
                                <span class="preview-name" style="display: flex; flex-direction: column; color: var(--text-primary); font-weight: 600; ">
                                    <span>${item.name || item.categoryName || ''}</span>
                                    ${item.notes && item.notes !== item.name ? `<span style="font-size: 0.85rem; margin-top: 0px;"><span style="position: relative; top: 3px;">*</span> ${item.notes}</span>` : ''}
                                    ${item.extras && item.extras.length > 0 ? `<span style="font-size: 0.85rem; margin-top: 0px;">+ ${(Array.isArray(item.extras) ? item.extras.map(e => typeof e === 'object' ? e.name : e).join(', ') : item.extras)}</span>` : ''}
                                </span>
                            </div>
                            <span class="item-price">${formatPrice(item.price || (item.unitPrice * item.qty))}</span>
                        </div>
                    `).join('')}
                </div>
                <div class="order-card-footer">
                    <div style="display: flex; gap: var(--space-sm); align-items: center;">
                        <span class="order-total">${formatPrice(order.totalPrice)}</span>
                        ${!order.paid ? `
                            <button class="btn-append-items" onclick="event.stopPropagation(); window.appendToOrder('${order.id}')" 
                                style="background: var(--accent-primary); color: white; border: none; padding: 4px 12px; border-radius: var(--radius-sm); font-size: 0.85rem; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                                <i data-lucide="plus" style="width: 14px; height: 14px;"></i> ADICIONAR
                            </button>
                        ` : ''}
                    </div>
                    ${order.paid ? `
                        <div style="display: flex; gap: 8px; align-items: center;">
                            <span class="status-indicator" style="background: var(--bg-tertiary); color: var(--text-primary); border: 1px solid var(--border-color); font-size: 0.75rem; padding: 2px 8px; border-radius: 999px;">
                                ${(order.paymentMethod || 'EFECTIVO').toUpperCase()}
                            </span>
                            <span class="status-indicator paid-chip">PAGADO</span>
                        </div>
                    ` : ''}
                </div>
            </div>
        `;
    }

    window.appendToOrder = function (orderId) {
        const order = StorageManager.getOrders().find(o => o.id == orderId);
        if (!order) return;

        state.appendingOrderId = orderId;
        state.serviceType = order.serviceType;

        // Reset UI to "New Order" page
        state.currentPage = 'new-order';
        elements.pages.forEach(p => p.classList.remove('active'));
        const targetPage = document.getElementById('page-new-order');
        if (targetPage) targetPage.classList.add('active');

        // Update drawer state
        elements.drawerItems.forEach(i => i.classList.remove('active'));
        const newOrderTab = Array.from(elements.drawerItems).find(i => i.dataset.page === 'new-order');
        if (newOrderTab) newOrderTab.classList.add('active');

        // Initialize/Clear category rows
                resetAllCategories(); // Ensure we start with a clean UI
        initializeCategories();
        
        // Populate clients based on existing order items
        if (order.items && order.items.length > 0) {
            const uniqueClients = [...new Set(order.items.map(i => i.clientName || 'P1'))];
            // AUTOMATICALLY ADD ONLY THE NEW CLIENT TAB WHEN APPENDING
            let maxNum = 0;
            uniqueClients.forEach(c => {
                const num = parseInt(c.replace('P', '')) || 0;
                if (num > maxNum) maxNum = num;
            });
            const newClient = 'P' + (maxNum + 1);
            state.clients = [newClient]; // Only show the new person's tab
            state.activeClient = newClient;
            if (typeof renderPosClientTabs === 'function') renderPosClientTabs();
        }

        state.serviceType = order.serviceType;

        // Update Service Tabs to match order
        elements.serviceTabs.forEach(tab => {
            tab.classList.toggle('active', tab.dataset.service === order.serviceType);
        });

        const locIn = document.getElementById('posLocationInput');
        const typeIn = document.getElementById('posOrderTypeInput');
        if (locIn) {
            locIn.value = order.customerInfo || '';
            locIn.disabled = true;
            locIn.style.opacity = '0.6';
        }
        if (typeIn) {
            typeIn.disabled = true;
            typeIn.style.opacity = '0.6';
        }

        // Show Footer
        const appFooter = document.getElementById('appFooter');
        if (appFooter) appFooter.style.display = 'flex';



        
showNotification(`Agregando productos a la Orden ${order.orderNumber}`);
        refreshOrderPageUI();
    };


    window.openPaymentModal = function(orderId) {
        const order = StorageManager.getOrders().find(o => o.id == orderId);
        if (!order || !elements.paymentModal) return;

        selectedPaymentOrder = order;
        elements.paymentOrderNum.textContent = order.orderNumber;
        elements.paymentTotal.textContent = formatPrice(order.totalPrice);
        elements.paymentTicketContent.innerHTML = generateTicketText(order);

        const modalTitle = elements.paymentModal.querySelector('h3');

        // Button visibility logic
        if (order.paid) {
            modalTitle.textContent = 'Pedido Pagado - Detalle';
            elements.confirmPayment.style.display = 'none';
            elements.printPaymentTicket.style.display = 'flex'; // Allow re-print
            if (elements.markReadyBtn) elements.markReadyBtn.style.display = 'none';
            if (elements.deleteOrderBtn) elements.deleteOrderBtn.style.display = 'flex';
        } else if (!order.checkoutPrinted) {
            modalTitle.textContent = 'Pedido Pendiente';
            elements.confirmPayment.style.display = 'none';
            elements.printPaymentTicket.style.display = 'flex';
            if (elements.markReadyBtn) elements.markReadyBtn.style.display = 'flex';
            if (elements.deleteOrderBtn) elements.deleteOrderBtn.style.display = 'flex';
        } else {
            modalTitle.textContent = 'Cobrar Pedido';
            elements.confirmPayment.style.display = 'flex';
            elements.printPaymentTicket.style.display = 'flex'; // Allow re-print even if in pending
            if (elements.markReadyBtn) elements.markReadyBtn.style.display = 'none';
            if (elements.deleteOrderBtn) elements.deleteOrderBtn.style.display = 'flex';
        }

        // Factura button visibility
        if (elements.invoicePaymentTicket) {
            // Only show Factura in "Por Cobrar" (printed) or "Pagadas" (paid)
            if (order.paid || order.checkoutPrinted) {
                elements.invoicePaymentTicket.style.display = 'flex';
            } else {
                elements.invoicePaymentTicket.style.display = 'none';
            }
        }



        // Reset Payment Method Logic (Radio Buttons)
        const radios = document.querySelectorAll('input[name="paymentMethod"]');
        radios.forEach(r => r.checked = false);
        state.selectedPaymentMethod = null; // Clear state just in case, though we read DOM now.

        // Show/Hide method selector based on payment status
        const methodContainer = document.querySelector('.payment-methods-container');
        if (methodContainer) {
            methodContainer.style.display = elements.confirmPayment.style.display === 'none' ? 'none' : 'block';
        }

        // Hide combined payment panel
        const combinedPanel = document.getElementById('combinedPaymentPanel');
        if (combinedPanel) {
            combinedPanel.style.display = 'none';
            document.getElementById('combinedEfectivo').value = '';
            document.getElementById('combinedNequi').value = '';
            document.getElementById('combinedDaviplata').value = '';
        }

        if (typeof lucide !== 'undefined') lucide.createIcons();
        elements.paymentModal.classList.remove('hidden');
    }



    document.querySelectorAll('.checkout-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.checkout-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            checkoutMode = tab.dataset.tab;
            renderCheckoutPage();
        });
    });

    function handleDirectCancel() {
        if (selectedPaymentOrder && !selectedPaymentOrder.paid) {
            const config = StorageManager.getConfig();
            if (config.billingSystem === 'direct') {
                // In direct mode, unpaid orders are discarded if payment is cancelled, but items are returned to cart
                state.cart = JSON.parse(JSON.stringify(selectedPaymentOrder.items));
                const uniqueClients = [...new Set(state.cart.map(i => i.clientName))];
                state.clients = uniqueClients.length > 0 ? uniqueClients : ['P1'];
                state.activeClient = state.clients[0];
                
                StorageManager.deleteOrder(selectedPaymentOrder.id);
                showNotification('Pedido devuelto para corrección.');
                
                // Return to new-order
                const newOrderDrawerItem = document.querySelector('.drawer-item[data-page="new-order"]');
                if (newOrderDrawerItem) newOrderDrawerItem.click();
                
                if(typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                renderPosClientTabs();
            }
        }
        elements.paymentModal.classList.add('hidden');
        selectedPaymentOrder = null;
    }

    if (elements.cancelPayment) {
        elements.cancelPayment.addEventListener('click', handleDirectCancel);
    }

    if (elements.paymentModalOverlay) {
        elements.paymentModalOverlay.addEventListener('click', handleDirectCancel);
    }

    // Floating Dropdown Logic (Mimics native select)
    function openFloatingDropdown(trigger, title, options, currentIds, onUpdate) {
        // Close any existing
        closeFloatingDropdown();

        const rect = trigger.getBoundingClientRect();

        const dropdown = document.createElement('div');
        dropdown.className = 'floating-dropdown';
        dropdown.style.position = 'fixed';
        dropdown.style.top = `${rect.bottom + 2}px`;
        dropdown.style.left = `${rect.left}px`;
        dropdown.style.width = `${rect.width}px`;
        dropdown.style.minWidth = '200px';
        dropdown.style.maxHeight = '300px';
        dropdown.style.overflowY = 'auto';
        dropdown.style.background = 'var(--bg-card)';
        dropdown.style.border = '1px solid var(--border-default)';
        dropdown.style.borderRadius = 'var(--radius-sm)';
        dropdown.style.boxShadow = '0 4px 12px rgba(0,0,0,0.3)';
        dropdown.style.zIndex = '9999';
        dropdown.style.padding = '4px';
        dropdown.id = 'activeFloatingDropdown';

        // Header for context?
        // Optional: Add a title row
        // const header = document.createElement('div');
        // header.textContent = title; ...

        let tempSelected = [...currentIds];

        options.filter(o => o.active !== false).forEach(opt => {
            const row = document.createElement('div');
            row.className = 'selection-option';
            row.style.padding = '8px';
            row.style.display = 'flex';
            row.style.alignItems = 'center';
            row.style.cursor = 'pointer';
            row.style.gap = '8px';

            if (tempSelected.includes(opt.id)) {
                row.style.background = 'var(--bg-secondary)';
                row.style.color = 'var(--accent-primary)';
            }

            row.innerHTML = `
                <input type="checkbox" ${tempSelected.includes(opt.id) ? 'checked' : ''} style="pointer-events:none;">
                <span>${opt.name}</span>
            `;

            row.addEventListener('click', (e) => {
                e.stopPropagation();
                const input = row.querySelector('input');
                // Toggle
                const isChecked = !input.checked;
                input.checked = isChecked;

                if (isChecked) {
                    if (!tempSelected.includes(opt.id)) tempSelected.push(opt.id);
                    row.style.background = 'var(--bg-secondary)';
                    row.style.color = 'var(--accent-primary)';
                } else {
                    const idx = tempSelected.indexOf(opt.id);
                    if (idx > -1) tempSelected.splice(idx, 1);
                    row.style.background = 'transparent';
                    row.style.color = 'var(--text-primary)';
                }
                onUpdate(tempSelected);
            });
            dropdown.appendChild(row);
        });

        // Adjust position if offscreen
        document.body.appendChild(dropdown);
        const dropRect = dropdown.getBoundingClientRect();
        if (dropRect.bottom > window.innerHeight) {
            dropdown.style.top = `${rect.top - dropRect.height - 2}px`;
        }
        if (dropRect.right > window.innerWidth) {
            dropdown.style.left = `${window.innerWidth - dropRect.width - 10}px`;
        }

        // Click outside closes
        setTimeout(() => {
            document.addEventListener('click', closeOnOutsideClick);
        }, 0);
    }

    function closeFloatingDropdown() {
        const existing = document.getElementById('activeFloatingDropdown');
        if (existing) existing.remove();
        document.removeEventListener('click', closeOnOutsideClick);
    }

    function closeOnOutsideClick(e) {
        if (!e.target.closest('#activeFloatingDropdown')) {
            closeFloatingDropdown();
        }
    }

    if (elements.confirmPayment) {
        elements.confirmPayment.addEventListener('click', async () => {
            if (selectedPaymentOrder) {
                // Get selected radio
                const selectedRadio = document.querySelector('input[name="paymentMethod"]:checked');

                if (!selectedRadio) {
                    showNotification('Atenci&oacute;n: Selecciona un medio de pago', 'error');
                    return;
                }

                const method = selectedRadio.value;
                let paymentDetails = null;

                // Handle combined payment
                if (method === 'combinado') {
                    const efectivo = parseFloat(document.getElementById('combinedEfectivo').value) || 0;
                    const nequi = parseFloat(document.getElementById('combinedNequi').value) || 0;
                    const daviplata = parseFloat(document.getElementById('combinedDaviplata').value) || 0;
                    const total = efectivo + nequi + daviplata;

                    
if (total < selectedPaymentOrder.totalPrice) {
                        showNotification(`Atención: Faltan $$${formatPrice(selectedPaymentOrder.totalPrice - total).replace('$', '')} para completar el pago`, 'error');
                        return;

                    }

                    paymentDetails = {
                        efectivo: efectivo,
                        nequi: nequi,
                        daviplata: daviplata,
                        total: total
                    };
                }

                // Open cash drawer if printer is connected
                if (window.openCashDrawer) {
                    await window.openCashDrawer();
                }

                StorageManager.updateOrder(selectedPaymentOrder.id, {
                    paid: true,
                    status: 'delivered',
                    paymentMethod: method,
                    paymentDetails: paymentDetails
                });
                showNotification(`Pedido ${selectedPaymentOrder.orderNumber} pagado`);
                elements.paymentModal.classList.add('hidden');
                renderCheckoutPage();
            }
        });
    }

    // Combined payment panel toggle
    document.querySelectorAll('input[name="paymentMethod"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            const panel = document.getElementById('combinedPaymentPanel');
            if (panel) {
                panel.style.display = e.target.value === 'combinado' ? 'block' : 'none';
                if (e.target.value === 'combinado') {
                    // Pre-fill with total in efectivo
                    const totalEl = document.getElementById('paymentTotal');
                    if (totalEl && selectedPaymentOrder) {
                        document.getElementById('combinedEfectivo').value = selectedPaymentOrder.totalPrice;
                        document.getElementById('combinedNequi').value = '';
                        document.getElementById('combinedDaviplata').value = '';
                        updateCombinedTotal();
                    }
                }
            }
        });
    });

    // Update combined total display
    function updateCombinedTotal() {
        const efectivo = parseFloat(document.getElementById('combinedEfectivo').value) || 0;
        const nequi = parseFloat(document.getElementById('combinedNequi').value) || 0;
        const daviplata = parseFloat(document.getElementById('combinedDaviplata').value) || 0;
        const total = efectivo + nequi + daviplata;
        const totalEl = document.getElementById('combinedTotal');
        if (totalEl && selectedPaymentOrder) {
            const diff = total - selectedPaymentOrder.totalPrice;
            if (diff >= 0) {
                  totalEl.innerHTML = "Cambio Total: " + formatPrice(diff);
                  totalEl.style.color = '#059669';
            } else {
                totalEl.innerHTML = `Faltan: ${formatPrice(Math.abs(diff))}`;
                totalEl.style.color = '#dc2626';
            }
        }
    }

    // Bind combined payment inputs
    ['combinedEfectivo', 'combinedNequi', 'combinedDaviplata'].forEach(id => {
        const input = document.getElementById(id);
        if (input) {
            input.addEventListener('input', updateCombinedTotal);
        }
    });

    // Print button in payment modal
    if (elements.printPaymentTicket) {
        elements.printPaymentTicket.addEventListener('click', () => {
            if (selectedPaymentOrder) {
                // Show Printing Dialog
                window.print();

                if (selectedPaymentOrder.isPartial) {
                    // If it's a partial order (addition), delete it after printing
                    StorageManager.deleteOrder(selectedPaymentOrder.id);
                    showNotification(`Ticket de adición impreso`);
                } else {
                    // Normal order: Set as printed for checkout
                    StorageManager.updateOrder(selectedPaymentOrder.id, { checkoutPrinted: true });
                    showNotification(`Pedido ${selectedPaymentOrder.orderNumber} enviado a cobrar`);
                }

                                // Refresh and close
                elements.paymentModal.classList.add('hidden');
                renderCheckoutPage();
            }
        });
    }

    if (elements.markReadyBtn) {
        elements.markReadyBtn.addEventListener('click', () => {
            if (selectedPaymentOrder) {
                if (selectedPaymentOrder.isPartial) {
                    StorageManager.deleteOrder(selectedPaymentOrder.id);
                    showNotification('Adición procesada');
                } else {
                    StorageManager.updateOrder(selectedPaymentOrder.id, { checkoutPrinted: true });
                    showNotification('Pedido ' + selectedPaymentOrder.orderNumber + ' marcado como listo');
                }
                elements.paymentModal.classList.add('hidden');
                renderCheckoutPage();
            }
        });
    }

    if (elements.invoicePaymentTicket) {
        elements.invoicePaymentTicket.addEventListener('click', () => {
            if (selectedPaymentOrder) {
                const originalContent = elements.paymentTicketContent.innerHTML;
                elements.paymentTicketContent.innerHTML = generateInvoiceText(selectedPaymentOrder);
                window.print();
                elements.paymentTicketContent.innerHTML = originalContent; // Revert to comanda
                showNotification(`Factura de pedido ${selectedPaymentOrder.orderNumber} generada`);
            }
        });
    }

    if (elements.deleteOrderBtn) {
        elements.deleteOrderBtn.addEventListener('click', () => {
            if (!selectedPaymentOrder) return;

            const performDelete = () => {
                window.minesofConfirm(`Estas seguro de que deseas eliminar permanentemente el pedido ${selectedPaymentOrder.orderNumber}?`, async () => {
                    await StorageManager.deleteOrder(selectedPaymentOrder.id);
                    showNotification(`Pedido ${selectedPaymentOrder.orderNumber} eliminado`);
                    elements.paymentModal.classList.add('hidden');
                    renderCheckoutPage();
                });
            };

            if (state.isAdminAuthenticated) {
                performDelete();
            } else {
                state.pendingAdminAction = performDelete;
                elements.adminLoginModal.classList.add('open');
                elements.adminPasswordInput.value = '';
                elements.adminPasswordInput.focus();
            }
        });
    }

    // ============================================
    // Orders / Kitchen
    // ============================================

    function renderOrdersPage() {
        const orders = StorageManager.getActiveOrders().reverse();
        const container = document.getElementById('ordersList');
        if (!container) return;

        if (orders.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <i data-lucide="clipboard-list"></i>
                    <h3>No hay pedidos activos</h3>
                </div>
            `;
        } else {
            container.innerHTML = orders.map(order => createOrderListCard(order)).join('');
        }
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    function createOrderListCard(order) {
        const labels = { pending: 'Pendiente', preparing: 'Preparando', ready: 'Listo', delivered: 'Entregado' };
        const titleName = order.createdBy ? order.createdBy.toUpperCase() : 'CAJA';
        const titleSeq = order.sequenceNumber ? `(${order.sequenceNumber})` : `(${order.orderNumber})`;
        return `
            <div class="order-list-card">
                <div class="order-card-header">
                    <span class="order-number">${titleName} ${titleSeq}</span>
                    <span class="order-status-badge">${labels[order.status]}</span>
                </div>
                <div class="order-customer-info"><span>${order.customerInfo}</span></div>
                <div class="order-items-preview">
                    ${[...(order.items || [])].sort((a,b) => (a.clientName || "Z").localeCompare(b.clientName || "Z")).map((item, i, arr) => `
                        ${i > 0 && item.clientName !== arr[i-1].clientName ? '<div style="height: 12px; width: 100%;"></div>' : ''}
                        <div class="preview-item">
                            <div class="item-main">
                                <span class="preview-qty">${item.clientName || item.qty}</span>
                                <span class="preview-name" style="display: flex; flex-direction: column; color: var(--text-primary); font-weight: 600; ">
                                    <span>${item.name || item.categoryName || ''}</span>
                                    ${item.notes && item.notes !== item.name ? `<span style="font-size: 0.85rem; margin-top: 0px;"><span style="position: relative; top: 3px;">*</span> ${item.notes}</span>` : ''}
                                    ${item.extras && item.extras.length > 0 ? `<span style="font-size: 0.85rem; margin-top: 0px;">+ ${(Array.isArray(item.extras) ? item.extras.map(e => typeof e === 'object' ? e.name : e).join(', ') : item.extras)}</span>` : ''}
                                </span>
                            </div>
                            <span class="item-price">${formatPrice(item.price / item.qty)}</span>
                        </div>
                    `).join('')}
                </div>
                <div class="order-card-footer">
                    <span class="order-total">${formatPrice(order.totalPrice)}</span>
                </div>
            </div>
        `;
    }

    // ============================================
    // Reports
    // ============================================

    function renderReportsPage() {
        const period = elements.reportPeriodSelect?.value || 'today';
        let orders = [];

        switch (period) {
            case 'today':
                orders = StorageManager.getTodayOrders();
                break;
            case 'date':
                const filterDate = elements.reportDatePicker?.value;
                orders = filterDate ? StorageManager.getOrdersByDate(filterDate) : StorageManager.getTodayOrders();
                break;
            case 'month':
                orders = StorageManager.getCurrentMonthOrders();
                break;
            case 'specific-month':
                const filterMonth = elements.reportMonthPicker?.value;
                orders = filterMonth ? StorageManager.getOrdersByMonth(filterMonth) : StorageManager.getCurrentMonthOrders();
                break;
            case 'total':
                orders = StorageManager.getOrders();
                break;
            default:
                orders = StorageManager.getTodayOrders();
        }

        // Filter out partial/temporary orders from reports
        orders = orders.filter(o => !o.isPartial);

        const paidOrders = orders.filter(o => o.paid);

        const totalSales = paidOrders.reduce((sum, o) => sum + o.totalPrice, 0);
        
        let totalEfectivo = 0;
        let totalNequi = 0;
        let totalDaviplata = 0;

        

        // Metrics Maps
        const categorySales = {};
        const categoryQtyStats = {}; // { cat: { total: 0, sizes: {} } }
        const flavorStats = { all: {} };
        const sizeCounts = { 'XS': 0, 'XM': 0, 'XL': 0, 'X': 0, 'HB': 0, 'PE': 0, 'SA': 0 };
        const extrasSales = {};
        const paymentList = { efectivo: [], nequi: [], daviplata: [] };
        const salesBreakdownByDay = {}; // Grouped by YYYY-MM-DD

        paidOrders.forEach(order => {
            // Process and categorize all items in the order
            let orderFood = 0;
            let orderDrinks = 0;
            let orderDisposables = 0;
            let orderOthers = 0;

            order.items.forEach(item => {
                const catId = (item.category || '').toLowerCase();
                const catName = (item.categoryName || '').toLowerCase();

                                (item.flavors || []).forEach(f => {
                    if(f) flavorStats.all[f] = (flavorStats.all[f] || 0) + item.qty;
                });
                // Category Sales breakdown
                const displayCatName = item.categoryName || 'Otros';
                categorySales[displayCatName] = (categorySales[displayCatName] || 0) + item.price;

                // Category Quantity breakdown
                if (!categoryQtyStats[displayCatName]) {
                    categoryQtyStats[displayCatName] = { total: 0, sizes: {} };
                }
                categoryQtyStats[displayCatName].total += item.qty;
                if (item.size) {
                    categoryQtyStats[displayCatName].sizes[item.size] = (categoryQtyStats[displayCatName].sizes[item.size] || 0) + item.qty;
                }

                // Size Counts
                if (item.size && sizeCounts.hasOwnProperty(item.size)) {
                    sizeCounts[item.size] += item.qty;
                }

                // Extras Sales
                (item.extras || []).forEach(extraName => {
                    if (extraName) extrasSales[extraName] = (extrasSales[extraName] || 0) + item.qty;
                });
            });

            // Use Local Date for grouping
            const d = new Date(order.createdAt);
            const dateKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

            if (!salesBreakdownByDay[dateKey]) {
                salesBreakdownByDay[dateKey] = { food: 0, drinks: 0, desechables: 0, otros: 0, total: 0 };
            }
            salesBreakdownByDay[dateKey].food += orderFood;
            salesBreakdownByDay[dateKey].drinks += orderDrinks;
            salesBreakdownByDay[dateKey].desechables += orderDisposables;
            salesBreakdownByDay[dateKey].otros += orderOthers;
            salesBreakdownByDay[dateKey].total += order.totalPrice;

            // Payment Method Breakdown (including combined payments)
            const method = order.paymentMethod || 'efectivo';
            if (method === 'combinado' && order.paymentDetails) {
                // Split combined payments to their respective methods
                const ef = order.paymentDetails.efectivo || 0;
                const nq = order.paymentDetails.nequi || 0;
                const dv = order.paymentDetails.daviplata || 0;

                totalEfectivo += ef;
                totalNequi += nq;
                totalDaviplata += dv;

                if (ef > 0) paymentList.efectivo.push({ date: order.createdAt, amount: ef });
                if (nq > 0) paymentList.nequi.push({ date: order.createdAt, amount: nq });
                if (dv > 0) paymentList.daviplata.push({ date: order.createdAt, amount: dv });
            } else if (method === 'nequi') {
                totalNequi += order.totalPrice;
                paymentList.nequi.push({ date: order.createdAt, amount: order.totalPrice });
            } else if (method === 'daviplata') {
                totalDaviplata += order.totalPrice;
                paymentList.daviplata.push({ date: order.createdAt, amount: order.totalPrice });
            } else {
                totalEfectivo += order.totalPrice;
                paymentList.efectivo.push({ date: order.createdAt, amount: order.totalPrice });
            }
        });

        // Store for details
        currentReportOrders = paidOrders;
        lastSalesBreakdown = salesBreakdownByDay;

        // Update top cards
        if (elements.reportDailySales) elements.reportDailySales.textContent = formatPrice(totalSales);
        
        
        
        if (elements.reportEfectivoSales) elements.reportEfectivoSales.textContent = formatPrice(totalEfectivo);
        if (elements.reportNequiSales) elements.reportNequiSales.textContent = formatPrice(totalNequi);
        if (elements.reportDaviplataSales) elements.reportDaviplataSales.textContent = formatPrice(totalDaviplata);

        // Re-attach listeners for detailed view
        document.querySelectorAll('.report-clickable').forEach(card => {
            card.addEventListener('click', () => {
                const method = card.getAttribute('data-report-filter');
                if (method) showReportPaymentDetail(method);
            });
        });

        // 1. Render Categories (Sorted by Price)
        if (elements.categorySalesList) {
            const sortedCats = Object.entries(categorySales).sort((a, b) => b[1] - a[1]);
            const maxSales = sortedCats.length > 0 ? sortedCats[0][1] : 1;
            elements.categorySalesList.innerHTML = sortedCats.map(([name, amount]) => {
                const percentage = (amount / maxSales) * 100;
                return `
                    <div class="category-sales-item">
                        <span class="cat-sales-name">${name}</span>
                        <div class="cat-sales-bar-bg"><div class="cat-sales-bar-fill" style="width: ${percentage}%"></div></div>
                        <span class="cat-sales-amount">${formatPrice(amount)}</span>
                    </div>`;
            }).join('') || '<div class="empty-state">Sin ventas</div>';
        }

        // 1.5. Render Category Quantities
        if (elements.categoryQtyList) {
            const sortedCats = Object.entries(categoryQtyStats).sort((a, b) => b[1].total - a[1].total);
            elements.categoryQtyList.innerHTML = sortedCats.map(([name, stat]) => {
                const sizesHtml = Object.entries(stat.sizes)
                    .map(([size, qty]) => `<span class="qty-pill">${size}: ${qty}</span>`)
                    .join(' ');

                return `
                    <div class="category-qty-item">
                        <div class="qty-item-header">
                            <span class="cat-sales-name">${name}</span>
                            <span class="cat-qty-total">${stat.total} ud.</span>
                        </div>
                        <div class="qty-item-details">
                            ${sizesHtml}
                        </div>
                    </div>`;
            }).join('') || '<div class="empty-state">Sin datos</div>';
        }

        // 2. Render Top Flavors (Grouped and Sorted)
        if (elements.flavorSalesList) {
            let flavorsHtml = '';

            const entries = Object.entries(flavorStats.all || {}).sort((a, b) => b[1] - a[1]);
            if (entries.length > 0) {
                flavorsHtml += entries.map(([name, count]) => `
                    <div class="stats-row">
                        <span class="stats-label">${name}</span>
                        <span class="stats-value">${count} ud.</span>
                    </div>
                `).join('');
            }

            elements.flavorSalesList.innerHTML = flavorsHtml || '<div class="empty-state">Sin datos</div>';
        }

        // 3. Render Sizes
        if (elements.sizeSalesList) {
            const sizesToShow = ['XS', 'XM', 'XL', 'X'];
            if (sizeCounts['HB'] > 0 || sizeCounts['PE'] > 0 || sizeCounts['SA'] > 0) {
                sizesToShow.push('HB', 'PE', 'SA');
            }
            elements.sizeSalesList.innerHTML = sizesToShow.map(size => `
                <div class="size-stat-box">
                    <span class="size-name">${size}</span>
                    <span class="size-count">${sizeCounts[size] || 0}</span>
                </div>
            `).join('');
        }

        // 4. Render Extras (Sorted by Popularity/Count)
        if (elements.extrasSalesList) {
            const sortedExtras = Object.entries(extrasSales).sort((a, b) => b[1] - a[1]);
            const maxExtras = sortedExtras.length > 0 ? sortedExtras[0][1] : 1;
            elements.extrasSalesList.innerHTML = sortedExtras.map(([name, count]) => {
                const percentage = (count / maxExtras) * 100;
                return `
                    <div class="category-sales-item">
                        <span class="cat-sales-name">${name}</span>
                        <div class="cat-sales-bar-bg"><div class="cat-sales-bar-fill" style="background: var(--accent-gold); width: ${percentage}%"></div></div>
                        <span class="cat-sales-amount">${count} ud.</span>
                    </div>`;
            }).join('') || '<div class="empty-state">Sin adicionales</div>';
        }

        // 5. Render Sales Breakdown Table (Grouped by Day)
        if (true) {
            const container = document.getElementById('salesBreakdownTable');
            if (container) {
                const days = Object.keys(salesBreakdownByDay).sort((a, b) => new Date(b) - new Date(a));
                if (days.length === 0) {
                    container.innerHTML = '<div style="padding: 1rem; text-align: center; color: var(--text-muted); font-size: 0.9rem;">Sin datos</div>';
                } else {
                    let html = `
                        <div style="overflow-x: auto; border-radius: var(--radius-md); border: 1px solid var(--border-subtle);">
                        <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
                            <thead>
                                <tr style="background: var(--bg-tertiary);">
                                    <th style="padding: 10px 12px; text-align: left; color: var(--text-muted); font-weight: 600;  font-size: 0.65rem;">Fecha</th>
                                    
                                    
                                    
                                    
                                    <th style="padding: 10px 12px; text-align: right; color: var(--text-muted); font-weight: 600;  font-size: 0.65rem;">Total</th>
                                </tr>
                            </thead>
                            <tbody>
                    `;

                    days.forEach(dateStr => {
                        const s = salesBreakdownByDay[dateStr];
                        const dateObj = new Date(dateStr + 'T12:00:00'); // Midday to avoid TZ issues
                        const displayDate = dateObj.toLocaleDateString([], { day: '2-digit', month: '2-digit', year: 'numeric' });

                        html += `
                            <tr style="border-top: 1px solid var(--border-subtle); background: var(--bg-secondary);">
                                <td style="padding: 10px 12px; color: var(--text-primary); font-weight: 600;">${displayDate}</td>
                                
                                
                                
                                
                                <td style="padding: 10px 12px; text-align: right; font-weight: 800; color: var(--accent-primary);">${formatPrice(s.total)}</td>
                            </tr>
                        `;
                    });

                    html += `</tbody></table></div>`;
                    container.innerHTML = html;
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }
            }
        }

        // Update total label in header
        const salesTotalEl = document.getElementById('reportSalesTotal');
        if (salesTotalEl) salesTotalEl.textContent = formatPrice(totalSales);

        // 6. Render Payment Detail Tables
        const renderPaymentTable = (elId, list, color) => {
            const container = document.getElementById(elId);
            if (!container) return;

            if (list.length === 0) {
                container.innerHTML = '<div style="padding: 1rem; text-align: center; color: var(--text-muted); font-size: 0.9rem;">Sin transacciones</div>';
                return;
            }

            // Sort by date descending (most recent first)
            list.sort((a, b) => new Date(b.date) - new Date(a.date));

            let html = `
                <div style="overflow-y: auto; max-height: 250px; border-radius: var(--radius-md); border: 1px solid var(--border-subtle);">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
                    <thead>
                        <tr style="background: var(--bg-tertiary);">
                            <th style="padding: 8px 12px; text-align: left; color: var(--text-muted); font-weight: 600;  font-size: 0.65rem; position: sticky; top: 0; background: var(--bg-tertiary); z-index: 10;">Fecha</th>
                            <th style="padding: 8px 12px; text-align: right; color: var(--text-muted); font-weight: 600;  font-size: 0.65rem; position: sticky; top: 0; background: var(--bg-tertiary); z-index: 10;">Monto</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            list.forEach(p => {
                const dateObj = new Date(p.date);
                const dateStr = dateObj.toLocaleDateString([], { day: '2-digit', month: '2-digit' });
                const timeStr = dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

                html += `
                    <tr style="border-top: 1px solid var(--border-subtle);">
                        <td style="padding: 8px 12px; color: var(--text-secondary);">
                            ${dateStr} <span style="font-size: 0.7rem; opacity: 0.6;">${timeStr}</span>
                        </td>
                        <td style="padding: 8px 12px; text-align: right; font-weight: 700; color: ${color};">
                            ${formatPrice(p.amount)}
                        </td>
                    </tr>
                `;
            });

            html += `</tbody></table></div>`;
            container.innerHTML = html;
        if (typeof lucide !== 'undefined') lucide.createIcons();
    };

        renderPaymentTable('paymentTableEfectivo', paymentList.efectivo, '#16a34a');
        renderPaymentTable('paymentTableNequi', paymentList.nequi, '#60a5fa');
        renderPaymentTable('paymentTableDaviplata', paymentList.daviplata, '#fb923c');

        // Update total labels in headers
        const efTotalEl = document.getElementById('reportEfectivoSalesTotal');
        const nqTotalEl = document.getElementById('reportNequiSalesTotal');
        const dvTotalEl = document.getElementById('reportDaviplataSalesTotal');
        if (efTotalEl) efTotalEl.textContent = formatPrice(totalEfectivo);
        if (nqTotalEl) nqTotalEl.textContent = formatPrice(totalNequi);
        if (dvTotalEl) dvTotalEl.textContent = formatPrice(totalDaviplata);

        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    if (elements.searchReportBtn) {
        elements.searchReportBtn.addEventListener('click', () => {
            renderReportsPage();
        });
    }

    if (elements.downloadReportBtn) {
        elements.downloadReportBtn.addEventListener('click', () => {
            exportSalesToExcel();
        });
    }

    function exportSalesToExcel() {
        if (!lastSalesBreakdown || Object.keys(lastSalesBreakdown).length === 0) {
            showNotification('No hay datos para exportar', 'error');
            return;
        }

        const days = Object.keys(lastSalesBreakdown).sort((a, b) => new Date(a) - new Date(b));

        // Prepare data for SheetJS
        const data = days.map(date => {
            const s = lastSalesBreakdown[date];
            return {
                "Fecha": date,
                
                "Total": s.total
            };
        });

        // Add a Footer row with totals
        const totals = {
            "Fecha": "TOTALES",
            
            "Total": days.reduce((sum, d) => sum + lastSalesBreakdown[d].total, 0)
        };
        data.push(totals);

        try {
            const worksheet = XLSX.utils.json_to_sheet(data);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Ventas Diarias");

            // Adjust column widths
            worksheet['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }];

            XLSX.writeFile(workbook, `Reporte_Ventas_${new Date().toISOString().split('T')[0]}.xlsx`);
            showNotification('Reporte Excel generado');
        } catch (e) {
            console.error("SheetJS Error:", e);
            showNotification('Error al generar Excel', 'error');
        }
    }

    function showReportPaymentDetail(method) {
        if (!elements.reportDetailModal || !elements.reportDetailList) return;

        // Filter orders and get amounts per method (including combined)
        const filtered = [];
        currentReportOrders.forEach(o => {
            const m = o.paymentMethod || 'efectivo';

            if (m === 'combinado' && o.paymentDetails) {
                // For combined payments, check if this method has an amount
                const amount = o.paymentDetails[method] || 0;
                if (amount > 0) {
                    filtered.push({ ...o, displayAmount: amount, isCombined: true });
                }
            } else if (m === method) {
                filtered.push({ ...o, displayAmount: o.totalPrice, isCombined: false });
            }
        });

        // Sort descending (most recent first)
        filtered.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

        elements.reportDetailTitle.textContent = `Detalle: ${method.toUpperCase()}`;

        let html = `
            <table class="report-detail-table">
                <thead>
                    <tr>
                        <th>Fecha/Hora</th>
                        <th>Venta</th>
                        <th>Cliente</th>
                        <th style="text-align: right;">Monto</th>
                    </tr>
                </thead>
                <tbody>
        `;

        if (filtered.length === 0) {
            html += `<tr><td colspan="4" style="text-align:center; padding: 2rem;">No hay transacciones registradas</td></tr>`;
        } else {
            filtered.forEach(o => {
                const dateObj = new Date(o.createdAt);
                const timeStr = dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                const dateStr = dateObj.toLocaleDateString([], { day: '2-digit', month: '2-digit' });

                html += `
                    <tr>
                        <td>
                            <div class="detail-date">${dateStr}</div>
                            <div style="font-size: 0.7rem; color: var(--text-muted);">${timeStr}</div>
                        </td>
                        <td class="detail-order-num">
                            ${o.orderNumber}
                            ${o.isCombined ? '<span style="font-size: 0.65rem; color: var(--text-muted); display: block;">(Combinado)</span>' : ''}
                        </td>
                        <td style="font-size: 0.9rem; max-width: 120px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                            ${o.customerInfo}
                        </td>
                        <td class="detail-amount">${formatPrice(o.displayAmount)}</td>
                    </tr>
                `;
            });
        }

        html += `</tbody></table>`;
        elements.reportDetailList.innerHTML = html;
        elements.reportDetailModal.classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // Modal close handlers for report detail
    if (elements.closeReportDetailModal) {
        elements.closeReportDetailModal.addEventListener('click', () => {
            elements.reportDetailModal.classList.remove('open');
        });
    }
    if (elements.closeReportDetailModalOverlay) {
        elements.closeReportDetailModalOverlay.addEventListener('click', () => {
            elements.reportDetailModal.classList.remove('open');
        });
    }

    if (elements.reportPeriodSelect) {
        elements.reportPeriodSelect.addEventListener('change', (e) => {
            const val = e.target.value;
            // Handle date picker visibility
            if (val === 'date') {
                elements.reportDatePickerGroup?.classList.remove('hidden');
            } else {
                elements.reportDatePickerGroup?.classList.add('hidden');
            }

            // Handle month picker visibility
            if (val === 'specific-month') {
                elements.reportMonthPickerGroup?.classList.remove('hidden');
            } else {
                elements.reportMonthPickerGroup?.classList.add('hidden');
            }

            // Auto-refresh for static options
            if (val !== 'date' && val !== 'specific-month') {
                renderReportsPage();
            }
        });
    }

    // ============================================
    // History
    // ============================================

    let historyMode = 'today';
    let historyFilter = 'all';

    function renderHistoryPage() {
        const ordersRaw = historyMode === 'today'
            ? StorageManager.getTodayOrders().reverse()
            : StorageManager.getOrdersByDate(elements.historyDatePicker.value).reverse();

        // Filter out partial orders from history and show only PAID
        let baseOrders = ordersRaw.filter(o => !o.isPartial && o.paid);
        
        // Populate user dropdown
        const historyUserFilterSelect = document.getElementById('historyUserFilter');
        if (historyUserFilterSelect) {
            const currentVal = historyUserFilterSelect.value || 'all';
            const users = new Set();
            baseOrders.forEach(o => { if (o.createdBy) users.add(o.createdBy); });
            const sortedUsers = Array.from(users).sort();
            
            // Only update DOM if the list changed, to prevent loss of focus/selection state
            const currentOptions = Array.from(historyUserFilterSelect.options).map(o => o.value).filter(v => v !== 'all');
            if (JSON.stringify(currentOptions) !== JSON.stringify(sortedUsers)) {
                let html = '<option value="all">Todos los Usuarios</option>';
                sortedUsers.forEach(u => html += `<option value="${u}">${u}</option>`);
                historyUserFilterSelect.innerHTML = html;
                historyUserFilterSelect.value = sortedUsers.includes(currentVal) ? currentVal : 'all';
            }
        }
        
        let orders = baseOrders;

        // Apply User Filter
        if (historyUserFilterSelect && historyUserFilterSelect.value !== 'all') {
            orders = orders.filter(o => o.createdBy === historyUserFilterSelect.value);
        }

        // Apply Payment Method or Category Filter
        if (historyFilter !== 'all') {
            if (['efectivo', 'nequi', 'daviplata'].includes(historyFilter)) {
                orders = orders.filter(o => {
                    const m = o.paymentMethod || 'efectivo';
                    if (m === historyFilter) return true;
                    if (m === 'combinado' && o.paymentDetails && (o.paymentDetails[historyFilter] || 0) > 0) return true;
                    return false;
                });
            }
        }

        renderHistoryOrdersList(orders);
        
        // Calculate total summary taking into account the USER filter, but NOT the payment method filter
        const summaryOrders = historyUserFilterSelect && historyUserFilterSelect.value !== 'all' 
            ? baseOrders.filter(o => o.createdBy === historyUserFilterSelect.value) 
            : baseOrders;
            
        calculateHistorySummary(summaryOrders); // Total summary shows all payment methods for the selected user

        // Ensure modal is hidden
        if (elements.historyOrderModal) {
            elements.historyOrderModal.classList.add('hidden');
            elements.historyOrderModal.style.display = 'none';
        }
        elements.historyOrdersList.classList.remove('hidden');

        // Show/Hide date picker container
        if (historyMode === 'date') {
            elements.datePickerContainer.classList.remove('hidden');
        } else {
            elements.datePickerContainer.classList.add('hidden');
        }
    }

    // Initialize Clickable Filter Cards in Summary
    document.querySelectorAll('.filter-card').forEach(card => {
        card.addEventListener('click', () => {
            document.querySelectorAll('.filter-card').forEach(c => c.classList.remove('active'));
            card.classList.add('active');
            historyFilter = card.dataset.filter;
            renderHistoryPage();
        });
    });

    elements.historyTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            elements.historyTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            historyMode = tab.dataset.tab;
            renderHistoryPage();
        });
    });

    if (elements.searchDateBtn) {
        elements.searchDateBtn.addEventListener('click', () => {
            if (!elements.historyDatePicker.value) {
                showNotification('Selecciona una fecha');
                return;
            }
            renderHistoryPage();
        });
    }

    const historyUserFilterSelect = document.getElementById('historyUserFilter');
    if (historyUserFilterSelect) {
        historyUserFilterSelect.addEventListener('change', () => {
            renderHistoryPage();
        });
    }

    if (elements.historyDatePicker) {
        // Set default to today
        const today = new Date().toISOString().split('T')[0];
        elements.historyDatePicker.value = today;

        // Auto-search on change
        elements.historyDatePicker.addEventListener('change', () => {
            if (historyMode === 'date') {
                renderHistoryPage();
            }
        });
    }

    if (elements.backToHistoryBtn) {
        elements.backToHistoryBtn.addEventListener('click', () => {
            elements.historyOrderModal.classList.add('hidden');
            elements.historyOrderModal.style.display = 'none';
        });
    }

    if (elements.historyModalOverlay) {
        elements.historyModalOverlay.addEventListener('click', () => {
            elements.historyOrderModal.classList.add('hidden');
            elements.historyOrderModal.style.display = 'none';
        });
    }

    let selectedHistoryOrder = null;

    function showOrderDetail(orderId) {
        const order = StorageManager.getOrders().find(o => o.id == orderId);
        if (!order) return;

        selectedHistoryOrder = order;
        elements.historyTicketContent.innerHTML = generateTicketText(order);

        elements.historyOrderModal.classList.remove('hidden');
        elements.historyOrderModal.style.display = 'flex';
    }

    if (elements.reprintOrderBtn) {
        elements.reprintOrderBtn.addEventListener('click', () => {
            if (selectedHistoryOrder) {
                // In a real app, this would send to a printer
                // For now we use browser print
                showTicketModal(selectedHistoryOrder);
                window.print();
            }
        });
    }

    if (elements.invoiceOrderBtn) {
        elements.invoiceOrderBtn.addEventListener('click', () => {
            if (selectedHistoryOrder) {
                const originalContent = elements.historyTicketContent.innerHTML;
                elements.historyTicketContent.innerHTML = generateInvoiceText(selectedHistoryOrder);
                window.print();
                elements.historyTicketContent.innerHTML = originalContent;
                showNotification(`Factura de pedido ${selectedHistoryOrder.orderNumber} generada`);
            }
        });
    }

    if (elements.deleteOrderBtnHistory) {
        elements.deleteOrderBtnHistory.addEventListener('click', () => {
            if (!selectedHistoryOrder) return;

            const performDelete = () => {
                window.minesofConfirm(`Estas seguro de que deseas eliminar permanentemente el pedido ${selectedHistoryOrder.orderNumber}?`, async () => {
                    await StorageManager.deleteOrder(selectedHistoryOrder.id);
                    showNotification(`Pedido ${selectedHistoryOrder.orderNumber} eliminado`);
                    elements.historyOrderModal.classList.add('hidden');
                    renderHistoryPage();
                });
            };

            if (state.isAdminAuthenticated) {
                performDelete();
            } else {
                state.pendingAdminAction = performDelete;
                elements.adminLoginModal.classList.add('open');
                elements.adminPasswordInput.value = '';
                elements.adminPasswordInput.focus();
            }
        });
    }

    function renderHistoryOrdersList(orders) {
        const container = document.getElementById('historyOrdersList');
        if (!container) return;
        const labels = { pending: 'Pendiente', preparing: 'Preparando', ready: 'Listo', delivered: 'Entregado' };

        if (orders.length === 0) {
            container.innerHTML = `
            <div style="grid-column: 1 / -1; width: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 60px 20px; color: #94a3b8; text-align: center;">
                <i data-lucide="receipt" style="width: 64px; height: 64px; margin-bottom: 16px; opacity: 0.5;"></i>
                <h3 style="font-size: 1.1rem; font-weight: 600; color: var(--text-primary); margin-bottom: 8px;">Aun no hay ventas</h3>
                <p style="font-size: 0.9rem; max-width: 250px;">Las ventas pagadas apareceran aqui para que lleves tu historial.</p>
            </div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
            return;
        }

        container.innerHTML = orders.map(order => {
            const titleName = order.createdBy ? order.createdBy.toUpperCase() : 'CAJA';
            const titleSeq = order.sequenceNumber ? `(${order.sequenceNumber})` : `(${order.orderNumber})`;
            return `
            <div class="order-list-card history-order-card" data-order-id="${order.id}">
                <div class="order-card-header">
                    <span class="order-number">${titleName} ${titleSeq}</span>
                    <span class="order-status-badge">${order.paid ? 'Pagado' : labels[order.status]}</span>
                </div>
                <div class="order-customer-info">
                    <span class="order-time">${new Date(order.createdAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                    <span> - ${order.customerInfo}</span>
                </div>
                <div class="order-items-preview">
                    ${[...(order.items || [])].sort((a,b) => (a.clientName || "Z").localeCompare(b.clientName || "Z")).map((item, i, arr) => `
                        ${i > 0 && item.clientName !== arr[i-1].clientName ? '<div style="height: 12px; width: 100%;"></div>' : ''}
                        <div class="preview-item">
                            <div class="item-main">
                                <span class="preview-qty">${item.clientName || item.qty}</span>
                                <span class="preview-name" style="display: flex; flex-direction: column; color: var(--text-primary); font-weight: 600; ">
                                    <span>${item.name || item.categoryName || ''}</span>
                                    ${item.notes && item.notes !== item.name ? `<span style="font-size: 0.85rem; margin-top: 0px;"><span style="position: relative; top: 3px;">*</span> ${item.notes}</span>` : ''}
                                    ${item.extras && item.extras.length > 0 ? `<span style="font-size: 0.85rem; margin-top: 0px;">+ ${(Array.isArray(item.extras) ? item.extras.map(e => typeof e === 'object' ? e.name : e).join(', ') : item.extras)}</span>` : ''}
                                </span>
                            </div>
                            <span class="item-price">${formatPrice(item.price / item.qty)}</span>
                        </div>
                    `).join('')}
                </div>
                <div class="order-card-footer">
                    <div class="footer-left">
                        <span class="order-total">${formatPrice(order.totalPrice)}</span>
                        <span class="payment-method-tag">${(order.paymentMethod || 'efectivo').toUpperCase()}</span>
                    </div>
                    ${order.paid ? '<span class="status-indicator paid-chip">PAGADO</span>' : ''}
                </div>
            </div>
        `;
        }).join('');

        // Re-attach click listeners for history detail
        document.querySelectorAll('.history-order-card').forEach(card => {
            card.addEventListener('click', () => {
                showOrderDetail(card.dataset.orderId);
            });
        });
    }

    function calculateHistorySummary(orders) {
        const paidOrders = orders.filter(o => o.paid);
        let totalSales = 0;
        let totalEfectivo = 0;
        let totalNequi = 0;
        let totalDaviplata = 0;
        

        

        paidOrders.forEach(order => {
            totalSales += order.totalPrice;
            const method = order.paymentMethod || 'efectivo';

            if (method === 'combinado' && order.paymentDetails) {
                totalEfectivo += order.paymentDetails.efectivo || 0;
                totalNequi += order.paymentDetails.nequi || 0;
                totalDaviplata += order.paymentDetails.daviplata || 0;
            } else if (method === 'nequi') {
                totalNequi += order.totalPrice;
            } else if (method === 'daviplata') {
                totalDaviplata += order.totalPrice;
            } else {
                totalEfectivo += order.totalPrice;
            }

            });

        if (elements.historyTotalSales) elements.historyTotalSales.textContent = formatPrice(totalSales);
        if (elements.historyTotalEfectivo) elements.historyTotalEfectivo.textContent = formatPrice(totalEfectivo);
        if (elements.historyTotalNequi) elements.historyTotalNequi.textContent = formatPrice(totalNequi);
        if (elements.historyTotalDaviplata) elements.historyTotalDaviplata.textContent = formatPrice(totalDaviplata);
        
        
        
    }

        function generateTicketText(order) {
        if (!order || !order.items) return 'Error: Pedido sin productos';
        const W = 32;
        const labels = { salon: 'SALON', llevar: 'LLEVAR', domicilio: 'DOMICILIO' };
        const now = new Date(order.createdAt || Date.now());
        const dateStr = now.toLocaleDateString('es-CO');
        const timeStr = now.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });

        const center = (str) => {
            str = String(str).toUpperCase();
            if (str.length >= W) return str.substring(0, W);
            const left = Math.floor((W - str.length) / 2);
            return ' '.repeat(left) + str;
        };

        const justify = (l, r) => {
            l = String(l); r = String(r);
            const sp = W - l.length - r.length;
            return sp < 1 ? l + ' ' + r : l + ' '.repeat(sp) + r;
        };

        const line = '-'.repeat(W);
        const doubleLine = '='.repeat(W);

        let t = '';
        t += doubleLine + '\n';
        t += center('Minesof') + '\n';
        if (FOODX_DATA.businessName) {
            t += center(FOODX_DATA.businessName.toUpperCase()) + '\n';
        }
        t += doubleLine + '\n';

        if (order.isAppending) {
            t += center('*** ADICION ***') + '\n';
        }

        t += center('PEDIDO ' + (order.sequenceNumber || order.orderNumber || '---')) + '\n';
        t += line + '\n';
        t += justify(dateStr, timeStr) + '\n';
        if (order.createdBy) {
            t += justify('USUARIO:', order.createdBy.toUpperCase().substring(0, 20)) + '\n';
        }

        if (order.customerInfo) {
            let cInfo = order.customerInfo.toUpperCase();
            if (cInfo === order.sequenceNumber || cInfo === order.orderNumber) {
                cInfo = 'CLIENTE';
            }
            t += justify('CLIENTE:', cInfo.substring(0, 23)) + '\n';
        }
        t += line + '\n';

        // Group items by client
        const itemsByClient = {};
        order.items.forEach(item => {
            const cName = item.clientName || 'GENERAL';
            if (!itemsByClient[cName]) itemsByClient[cName] = [];
            itemsByClient[cName].push(item);
        });

        const clientKeys = Object.keys(itemsByClient);

        let firstClient = true;
        for (const [clientName, cItems] of Object.entries(itemsByClient)) {
            if (!firstClient && clientKeys.length > 1) {
                t += '\n';
            }
            firstClient = false;
            cItems.forEach(item => {
                const qty = item.qty || 1;
                let name = (item.name || 'ITEM').toUpperCase();
                const prefix = (item.clientName || qty + 'x') + ' ';
                
                t += prefix + name + '\n';

                if (item.notes && item.notes.trim() !== '' && item.notes !== item.name) {
                    t += '  * ' + item.notes.toUpperCase() + '\n';
                }
                if (item.extras && item.extras.length > 0) {
                    const ext = Array.isArray(item.extras) ? item.extras.map(e => typeof e === 'object' ? e.name : e).join(', ') : item.extras;
                    t += '  + ' + ext.toUpperCase() + '\n';
                }
            });
        }

        t += doubleLine + '\n';
        
        if (order.paymentMethod) {
            t += justify('METODO DE PAGO:', order.paymentMethod.toUpperCase()) + '\n';
        }
        t += doubleLine + '\n';

        t += '\n';
        t += center('GRACIAS POR SU VISITA!') + '\n';
        t += center((FOODX_DATA.businessName || 'Minesof').toUpperCase()) + '\n';
        t += '\n\n.';

        return t;
    }

        function generateInvoiceText(order) {
        if (!order || !order.items) return 'Error: Pedido sin productos';
        const W = 32;
        const labels = { salon: 'SALON', llevar: 'LLEVAR', domicilio: 'DOMICILIO' };
        const now = new Date();
        const dateStr = now.toLocaleDateString('es-CO');
        const timeStr = now.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });

        const center = (str) => {
            str = String(str).toUpperCase();
            if (str.length >= W) return str.substring(0, W);
            const left = Math.floor((W - str.length) / 2);
            return ' '.repeat(left) + str;
        };

        const justify = (l, r) => {
            l = String(l); r = String(r);
            const sp = W - l.length - r.length;
            return sp < 1 ? l + ' ' + r : l + ' '.repeat(sp) + r;
        };

        const line = '-'.repeat(W);
        const doubleLine = '='.repeat(W);

        let t = '';
        t += doubleLine + '\n';
        t += center('FACTURA DE VENTA') + '\n';
        t += center('Minesof') + '\n';
        if (FOODX_DATA.businessName) {
            t += center(FOODX_DATA.businessName.toUpperCase()) + '\n';
        }
        t += doubleLine + '\n';

        t += center('PEDIDO ' + (order.sequenceNumber || order.orderNumber || '---')) + '\n';
        t += line + '\n';
        t += justify('FECHA PAGO:', dateStr) + '\n';
        t += justify('HORA PAGO:', timeStr) + '\n';
        if (order.createdBy) {
            t += justify('USUARIO:', order.createdBy.toUpperCase().substring(0, 20)) + '\n';
        }

        if (order.customerInfo) {
            let cInfo = order.customerInfo.toUpperCase();
            if (cInfo === order.sequenceNumber || cInfo === order.orderNumber) {
                cInfo = 'CLIENTE';
            }
            t += justify('CLIENTE:', cInfo.substring(0, 23)) + '\n';
        }
        t += line + '\n';

        // Group items by client
        const itemsByClient = {};
        order.items.forEach(item => {
            const cName = item.clientName || 'GENERAL';
            if (!itemsByClient[cName]) itemsByClient[cName] = [];
            itemsByClient[cName].push(item);
        });

        const clientKeys = Object.keys(itemsByClient);
        let grandTotal = 0;

        let firstClient = true;
        for (const [clientName, cItems] of Object.entries(itemsByClient)) {
            if (!firstClient && clientKeys.length > 1) {
                t += '\n';
            }
            firstClient = false;
            let clientSubtotal = 0;
            cItems.forEach(item => {
                const qty = item.qty || 1;
                const unitP = item.unitPrice || item.price || 0;
                const totalP = item.price || (unitP * qty);
                clientSubtotal += totalP;

                let name = (item.name || 'ITEM').toUpperCase();
                const priceStr = formatPrice(totalP);
                const prefix = (item.clientName || qty + 'x') + ' ';
                
                // Allow exactly 1 space between name and price
                const maxNameLen = W - priceStr.length - prefix.length - 1;
                if (name.length > maxNameLen) {
                    name = name.substring(0, maxNameLen);
                }
                
                t += justify(prefix + name, priceStr) + '\n';
            });

            if (clientKeys.length > 1) {
                t += justify('  SUBTOTAL ' + clientName.toUpperCase() + ':', formatPrice(clientSubtotal)) + '\n';
            }
            grandTotal += clientSubtotal;
        }

        t += doubleLine + '\n';
        t += justify('TOTAL PAGADO:', formatPrice(order.totalPrice || grandTotal)) + '\n';
        
        if (order.paymentMethod) {
            t += justify('METODO DE PAGO:', order.paymentMethod.toUpperCase()) + '\n';
        }
        t += doubleLine + '\n';

        t += '\n';
        t += center('GRACIAS POR SU VISITA!') + '\n';
        t += center((FOODX_DATA.businessName || 'Minesof').toUpperCase()) + '\n';
        t += '\n\n.';

        return t;
    }

    // ============================================
    // Expenses (Gastos)
    // ============================================

    // Helper to get categories as a map { id: { label, emoji } }
    function getExpenseCatMap() {
        const cats = StorageManager.getExpenseCategories();
        const map = {};
        cats.forEach(c => { map[c.id] = { label: c.label, emoji: c.emoji }; });
        return map;
    }

    // Color palette for category cards
    const expenseCatColors = ['#f59e0b', '#ef4444', '#8b5cf6', '#3b82f6', '#06b6d4', '#10b981', '#f97316', '#6b7280', '#e11d48', '#84cc16', '#14b8a6', '#a855f7'];

    function renderExpensesPage() {
        const period = document.getElementById('expensePeriodSelect')?.value || 'today';
        const searchField = document.getElementById('expenseSearchInput');
        const query = (searchField?.value || '').toLowerCase();
        let income = 0;

        switch (period) {
            case 'today':
                expenses = StorageManager.getTodayExpenses();
                income = StorageManager.getTodaySales();
                break;
            case 'date':
                const filterDate = document.getElementById('expenseDatePicker')?.value;
                expenses = filterDate ? StorageManager.getExpensesByDate(filterDate) : StorageManager.getTodayExpenses();
                income = filterDate ? StorageManager.getSalesByDate(filterDate) : StorageManager.getTodaySales();
                break;
            case 'month':
                expenses = StorageManager.getCurrentMonthExpenses();
                income = StorageManager.getCurrentMonthSales();
                break;
            case 'total':
                expenses = StorageManager.getExpenses();
                income = StorageManager.getTotalSales();
                break;
            case 'specific-month':
                const filterMonth = document.getElementById('expenseMonthPicker')?.value;
                expenses = filterMonth ? StorageManager.getExpensesByMonth(filterMonth) : StorageManager.getCurrentMonthExpenses();
                income = filterMonth ? StorageManager.getSalesByMonth(filterMonth) : StorageManager.getCurrentMonthSales();
                break;
            default:
                expenses = StorageManager.getTodayExpenses();
                income = StorageManager.getTodaySales();
        }

        // Filter by Search Query
        if (query) {
            expenses = expenses.filter(e => (e.description || '').toLowerCase().includes(query));
        }

        // Sort by date descending
        expenses.sort((a, b) => new Date(b.date || b.createdAt) - new Date(a.date || a.createdAt));

        const CATS = getExpenseCatMap();

        // Populate category select dynamically
        const catSelect = document.getElementById('expenseCategory');
        if (catSelect) {
            const currentVal = catSelect.value;
            const allCats = StorageManager.getExpenseCategories();
            catSelect.innerHTML = allCats.map(c => `<option value="${c.id}">${c.label}</option>`).join('');
            if (currentVal && allCats.find(c => c.id === currentVal)) catSelect.value = currentVal;
        }

        // Calculate totals
        const totalExpenses = expenses.reduce((sum, e) => sum + (e.amount || 0), 0);
        const netBalance = income - totalExpenses;

        const categoryTotals = {};
        expenses.forEach(e => {
            const cat = e.category || 'otros';
            categoryTotals[cat] = (categoryTotals[cat] || 0) + (e.amount || 0);
        });

        // Update cards
        const incomeEl = document.getElementById('expenseTotalIncome');
        const totalEl = document.getElementById('expenseTotalAmount');
        const netEl = document.getElementById('expenseNetBalance');
        const statusEl = document.getElementById('balanceStatus');

        if (incomeEl) incomeEl.textContent = formatPrice(income);
        if (totalEl) totalEl.textContent = formatPrice(totalExpenses);
        if (netEl) netEl.textContent = formatPrice(netBalance);

        if (statusEl) {
            if (netBalance > 0) {
                statusEl.textContent = 'Excedente';
                statusEl.style.background = '#dcfce7';
                statusEl.style.color = '#16a34a';
            } else if (netBalance < 0) {
                statusEl.textContent = 'DÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â©ficit';
                statusEl.style.background = '#fee2e2';
                statusEl.style.color = '#dc2626';
            } else {
                statusEl.textContent = 'Equilibrio';
                statusEl.style.background = '#f3f4f6';
                statusEl.style.color = '#6b7280';
            }
        }

        // Render category summary
        const summaryEl = document.getElementById('expenseCategorySummary');
        if (summaryEl) {
            const allCatsForColors = StorageManager.getExpenseCategories();
            summaryEl.innerHTML = Object.entries(categoryTotals)
                .sort((a, b) => b[1] - a[1])
                .map(([catId, amount]) => {
                    const cat = CATS[catId] || { label: catId, emoji: '📦' };
                    const idx = allCatsForColors.findIndex(c => c.id === catId);
                    const color = expenseCatColors[idx % expenseCatColors.length] || '#6b7280';
                    return `
                        <div style="background: var(--bg-secondary); border-radius: var(--radius-md); padding: var(--space-sm) var(--space-md); border-left: 3px solid ${color};">
                            <div style="font-size: 0.7rem; color: var(--text-muted);">${cat.label}</div>
                            <div style="font-size: 1rem; font-weight: 700; color: var(--text-primary);">${formatPrice(amount)}</div>
                        </div>
                    `;
                }).join('');
        }

        // Render expense list as table
        const listEl = document.getElementById('expensesList');
        if (listEl) {
            if (expenses.length === 0) {
                listEl.innerHTML = `
                    <div class="empty-state" style="padding: 2rem; text-align: center;">
                        <i data-lucide="wallet" style="width: 40px; height: 40px; color: var(--text-muted); margin-bottom: 8px;"></i>
                        <p style="color: var(--text-muted);">No hay Gastos registrados</p>
                    </div>
                `;
            } else {
                let tableHtml = `
                    <div style="overflow-x: auto; border-radius: var(--radius-md); border: 1px solid var(--border-subtle);">
                    <table style="width: 100%; border-collapse: collapse; font-size: 0.82rem;">
                        <thead>
                            <tr style="background: var(--bg-tertiary);">
                                <th style="padding: 10px 12px; text-align: left; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Fecha</th>
                                <th style="padding: 10px 12px; text-align: left; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Categoria</th>
                                <th style="padding: 10px 12px; text-align: left; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Descripcion</th>
                                <th style="padding: 10px 12px; text-align: center; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Cant.</th>
                                <th style="padding: 10px 12px; text-align: right; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Unit.</th>
                                <th style="padding: 10px 12px; text-align: right; color: var(--text-muted); font-weight: 600; font-size: 0.7rem;  letter-spacing: 0.5px;">Total</th>
                                <th style="padding: 10px 6px; width: 30px;"></th>
                            </tr>
                        </thead>
                        <tbody>
                `;

                expenses.forEach(expense => {
                    const cat = CATS[expense.category] || { label: 'Otros', emoji: '📦' };
                    const dateObj = new Date(expense.date || expense.createdAt);
                    const dateStr = dateObj.toLocaleDateString([], { day: '2-digit', month: '2-digit', year: '2-digit' });
                    const qty = expense.qty || 1;
                    const unitCost = expense.amount / qty;

                    tableHtml += `
                        <tr style="border-top: 1px solid var(--border-subtle); background: var(--bg-secondary);">
                            <td style="padding: 10px 12px; color: var(--text-secondary); white-space: nowrap; font-size: 0.9rem;">${dateStr}</td>
                            <td style="padding: 10px 12px; white-space: nowrap;">
                                <span style="font-size: 0.9rem;">${cat.label}</span>
                            </td>
                            <td style="padding: 10px 12px; color: var(--text-primary); max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 0.9rem;">
                                ${expense.description || '-'}
                            </td>
                            <td style="padding: 10px 12px; text-align: center; color: var(--text-primary); font-size: 0.9rem;">
                                ${qty}
                            </td>
                            <td style="padding: 10px 12px; text-align: right; color: var(--text-muted); font-size: 0.75rem;">
                                ${formatPrice(unitCost)}
                            </td>
                            <td style="padding: 10px 12px; text-align: right; font-weight: 700; color: #ef4444; white-space: nowrap;">
                                -${formatPrice(expense.amount)}
                            </td>
                            <td style="padding: 10px 6px; text-align: center;">
                                <button onclick="window.deleteExpense('${expense.id}')"
                                    style="background: none; border: none; color: var(--text-muted); cursor: pointer; padding: 2px; opacity: 0.5;"
                                    title="Eliminar">
                                    <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i>
                                </button>
                            </td>
                        </tr>
                    `;
                });

                tableHtml += `</tbody></table></div>`;
                listEl.innerHTML = tableHtml;
            }
        }

        // Set default date to today (local time) if empty
        const dateInput = document.getElementById('expenseDate');
        if (dateInput && !dateInput.value) {
            const now = new Date();
            const yyyy = now.getFullYear();
            const mm = String(now.getMonth() + 1).padStart(2, '0');
            const dd = String(now.getDate()).padStart(2, '0');
            dateInput.value = `${yyyy}-${mm}-${dd}`;
        }

        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // ---- Expense Category Manager ----
    function renderExpenseCategoriesManager() {
        const container = document.getElementById('expenseCatManager');
        if (!container) return;

        const cats = StorageManager.getExpenseCategories();
        let html = `
            <div style="overflow-x: auto; border-radius: var(--radius-md); border: 1px solid var(--border-subtle); margin-bottom: var(--space-sm);">
            <table style="width: 100%; border-collapse: collapse; font-size: 0.82rem;">
                <thead>
                    <tr style="background: var(--bg-tertiary);">
                        <th style="padding: 8px 12px; text-align: left; color: var(--text-muted); font-weight: 600; font-size: 0.7rem; ">Nombre de categoria</th>
                        <th style="padding: 8px 6px; width: 60px; text-align: center; color: var(--text-muted); font-weight: 600; font-size: 0.7rem; ">Acciones</th>
                    </tr>
                </thead>
                <tbody>
        `;

        cats.forEach(cat => {
            html += `
                <tr style="border-top: 1px solid var(--border-subtle); background: var(--bg-secondary);">
                    <td style="padding: 8px 12px; color: var(--text-primary); font-size: 0.85rem;">${cat.label}</td>
                    <td style="padding: 8px 6px; text-align: center; white-space: nowrap;">
                        <button onclick="window.moveExpenseCategory('${cat.id}', -1)" style="background: none; border: none; color: #64748b; cursor: pointer; padding: 4px;" title="Subir">
                            <i data-lucide="chevron-up" style="width: 16px; height: 16px;"></i>
                        </button>
                        <button onclick="window.moveExpenseCategory('${cat.id}', 1)" style="background: none; border: none; color: #64748b; cursor: pointer; padding: 4px; margin-right: 4px;" title="Bajar">
                            <i data-lucide="chevron-down" style="width: 16px; height: 16px;"></i>
                        </button>
                        <button onclick="window.editExpenseCategory('${cat.id}')"
                            style="background: none; border: none; color: var(--accent-primary); cursor: pointer; padding: 4px;" title="Editar">
                            <i data-lucide="edit-2" style="width: 16px; height: 16px;"></i>
                        </button>
                        <button onclick="window.deleteExpenseCategory('${cat.id}')"
                            style="background: none; border: none; color: #ef4444; cursor: pointer; padding: 4px; margin-left: 2px;" title="Eliminar">
                            <i data-lucide="trash-2" style="width: 16px; height: 16px;"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        html += `</tbody></table></div>`;

        // Add new category form
        html += `
            <div style="display: flex; gap: var(--space-xs); align-items: center;">
                <input type="text" autocomplete="off" id="newExpenseCatLabel" placeholder="Nombre de categoria"
                    style="flex: 1; padding: 8px 12px; background: var(--bg-tertiary); border: 1px solid var(--border-subtle); border-radius: var(--radius-md); color: var(--text-primary); font-size: 0.85rem; box-sizing: border-box;">
                <button onclick="window.addExpenseCategory()"
                    style="padding: 8px 14px; background: var(--accent-primary); color: var(--bg-primary); border: none; border-radius: var(--radius-md); font-weight: 700; font-size: 0.9rem; cursor: pointer; white-space: nowrap;">
                    + Agregar
                </button>
            </div>
        `;

        container.innerHTML = html;
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // Global handlers for category management
    window.moveExpenseCategory = function(catId, direction) {
        const cats = StorageManager.getExpenseCategories();
        const index = cats.findIndex(c => c.id === catId);
        if (index === -1) return;

        let targetIndex = -1;
        if (direction === -1 && index > 0) targetIndex = index - 1;
        if (direction === 1 && index < cats.length - 1) targetIndex = index + 1;

        if (targetIndex !== -1) {
            const temp = cats[index];
            cats[index] = cats[targetIndex];
            cats[targetIndex] = temp;

            StorageManager.saveExpenseCategories(cats);
            renderExpenseCategoriesManager();
        }
    };

    window.addExpenseCategory = function () {
        const label = document.getElementById('newExpenseCatLabel')?.value.trim();

        if (!label) {
            showNotification('ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¯ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â Ingresa un nombre para la Categoria', 'error');
            return;
        }

        const cats = StorageManager.getExpenseCategories();
        const id = label.toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');

        if (cats.find(c => c.id === id)) {
            showNotification('ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¯ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â Ya existe una Categoria con ese nombre', 'error');
            return;
        }

        cats.push({ id, label, emoji: '📦' });
        StorageManager.saveExpenseCategories(cats);
        showNotification(`Categoria "${label}" creada`);
        renderExpenseCategoriesManager();
    };

    window.editExpenseCategory = function (catId) {
        const cats = StorageManager.getExpenseCategories();
        const cat = cats.find(c => c.id === catId);
        if (!cat) return;

        window.minesofPrompt('Nombre de la categoria:', (newLabel) => {
            if (newLabel === null) return;
            cat.label = newLabel.trim() || cat.label;
            StorageManager.saveExpenseCategories(cats);
            showNotification(`Categoria actualizada: ${cat.label}`);
            renderExpenseCategoriesManager();
        }, cat.label);
    };

    window.deleteExpenseCategory = function (catId) {
        const performDelete = () => {
            window.minesofConfirm('Eliminar esta categoria de Gasto?', () => {
                const cats = StorageManager.getExpenseCategories().filter(c => c.id !== catId);
                StorageManager.saveExpenseCategories(cats);
                showNotification('Categoria eliminada');
                renderExpenseCategoriesManager();
            });
        };

        if (state.isAdminAuthenticated) {
            performDelete();
        } else {
            state.pendingAdminAction = performDelete;
            elements.adminLoginModal.classList.add('open');
            elements.adminPasswordInput.value = '';
            elements.adminPasswordInput.focus();
        }
    };

    // Add expense handler
    const addExpenseBtn = document.getElementById('addExpenseBtn');
    if (addExpenseBtn) {
        addExpenseBtn.addEventListener('click', () => {
            const category = document.getElementById('expenseCategory').value;
            const description = document.getElementById('expenseDescription').value.trim();
            const qty = parseFloat(document.getElementById('expenseQty').value) || 1;
            const amount = parseFloat(document.getElementById('expenseAmount').value);
            const date = document.getElementById('expenseDate').value;

            if (!amount || amount <= 0) {
                showNotification('ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¯ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¸ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â Ingresa un monto vÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡lido', 'error');
                return;
            }

            if (!date) {
                showNotification('Selecciona una fecha', 'error');
                return;
            }

            const CATS = getExpenseCatMap();
            const cat = CATS[category] || { label: 'Otros', emoji: '📦' };

            StorageManager.addExpense({
                category: category,
                categoryLabel: cat.label,
                description: description || cat.label,
                qty: qty,
                amount: amount,
                date: date + 'T12:00:00',
                createdBy: localStorage.getItem('minesof_deviceUser') || 'Cajero 1'
            });

            showNotification(`Gasto registrado: ${formatPrice(amount)}`);

            // Clear form
            document.getElementById('expenseDescription').value = '';
            document.getElementById('expenseQty').value = '1';
            document.getElementById('expenseAmount').value = '';
            if (document.getElementById('expenseAmountPreview')) {
                document.getElementById('expenseAmountPreview').textContent = '$0';
            }

            renderExpensesPage();
        });
    }

    const expenseAmountInput = document.getElementById('expenseAmount');
    if (expenseAmountInput) {
        expenseAmountInput.addEventListener('input', (e) => {
            const val = parseFloat(e.target.value) || 0;
            const preview = document.getElementById('expenseAmountPreview');
            if (preview) preview.textContent = formatPrice(val);
        });
    }

    // Period filter change
    const expensePeriodSelect = document.getElementById('expensePeriodSelect');
    if (expensePeriodSelect) {
        expensePeriodSelect.addEventListener('change', (e) => {
            const val = e.target.value;
            const monthPicker = document.getElementById('expenseMonthPickerWrapper') || document.getElementById('expenseMonthPicker');
            const datePicker = document.getElementById('expenseDatePickerWrapper') || document.getElementById('expenseDatePicker');
            
            if (monthPicker) monthPicker.classList.add('hidden');
            if (datePicker) datePicker.classList.add('hidden');

            if (val === 'specific-month') {
                monthPicker?.classList.remove('hidden');
            } else if (val === 'date') {
                datePicker?.classList.remove('hidden');
            } else {
                renderExpensesPage();
            }
        });
    }

    const expenseMonthPicker = document.getElementById('expenseMonthPicker');
    if (expenseMonthPicker) {
        expenseMonthPicker.addEventListener('change', () => {
            renderExpensesPage();
        });
    }

    const expenseDatePicker = document.getElementById('expenseDatePicker');
    if (expenseDatePicker) {
        expenseDatePicker.addEventListener('change', () => {
            renderExpensesPage();
        });
    }

    const expenseSearchInput = document.getElementById('expenseSearchInput');
    if (expenseSearchInput) {
        expenseSearchInput.addEventListener('input', () => {
            renderExpensesPage();
        });
    }

    // Delete expense (global handler)
    window.deleteExpense = function (expenseId) {
        const performDelete = () => {
            window.minesofConfirm('Eliminar este Gasto?', async () => {
                await StorageManager.deleteExpense(expenseId);
                showNotification('Gasto eliminado');
                renderExpensesPage();
            });
        };

        if (state.isAdminAuthenticated) {
            performDelete();
        } else {
            state.pendingAdminAction = performDelete;
            elements.adminLoginModal.classList.add('open');
            elements.adminPasswordInput.value = '';
            elements.adminPasswordInput.focus();
        }
    };

    // Download Expenses as Excel (.xlsx)
    window.downloadExpensesExcel = function () {
        const period = document.getElementById('expensePeriodSelect')?.value || 'today';
        const query = (document.getElementById('expenseSearchInput')?.value || '').toLowerCase();
        let expenses = [];
        let periodLabel = '';

        switch (period) {
            case 'today':
                expenses = StorageManager.getTodayExpenses();
                periodLabel = 'Hoy';
                break;
            case 'date':
                const exFilterDate = document.getElementById('expenseDatePicker')?.value;
                expenses = exFilterDate ? StorageManager.getExpensesByDate(exFilterDate) : StorageManager.getTodayExpenses();
                periodLabel = exFilterDate || 'Dia_Especifico';
                break;
            case 'month':
                expenses = StorageManager.getCurrentMonthExpenses();
                periodLabel = 'Este_Mes';
                break;
            case 'specific-month':
                const filterMonth = document.getElementById('expenseMonthPicker')?.value;
                expenses = filterMonth ? StorageManager.getExpensesByMonth(filterMonth) : StorageManager.getCurrentMonthExpenses();
                periodLabel = filterMonth || 'Mes_Especifico';
                break;
            case 'total':
                expenses = StorageManager.getExpenses();
                periodLabel = 'Todo';
                break;
            default:
                expenses = StorageManager.getTodayExpenses();
                periodLabel = 'Hoy';
        }

        if (expenses.length === 0) {
            showNotification('No hay Gastos para descargar', 'error');
            return;
        }

        // Filter by Search Query
        if (query) {
            expenses = expenses.filter(e => (e.description || '').toLowerCase().includes(query));
        }

        // Sort by date ascending
        expenses.sort((a, b) => new Date(a.date || a.createdAt) - new Date(b.date || b.createdAt));

        const CATS = getExpenseCatMap();

        // Build data rows matching the table: Fecha | Categoria | Descripcion | Cant. | V. Unit. | Total
        const rows = [['Fecha', 'Categoria', 'Descripcion', 'Cant.', 'V. Unit.', 'Total']];

        let total = 0;
        expenses.forEach(expense => {
            const cat = CATS[expense.category] || { label: 'Otros', emoji: '📦' };
            const dateObj = new Date(expense.date || expense.createdAt);
            const dateStr = dateObj.toLocaleDateString('es-CO', { day: '2-digit', month: '2-digit', year: 'numeric' });
            const qty = expense.qty || 1;
            const amount = expense.amount || 0;
            const unit = amount / qty;
            total += amount;

            rows.push([
                dateStr,
                cat.label,
                expense.description || cat.label,
                qty,
                unit,
                amount
            ]);
        });

        // Total row
        rows.push(['', '', '', '', 'TOTAL', total]);

        // Create workbook
        const ws = XLSX.utils.aoa_to_sheet(rows);

        // Set column widths
        ws['!cols'] = [
            { wch: 12 },  // Fecha
            { wch: 20 },  // Categoria
            { wch: 30 },  // Descripcion
            { wch: 8 },   // Cant.
            { wch: 12 },  // V. Unit.
            { wch: 12 }   // Total
        ];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Gastos');

        // Generate filename
        const now = new Date();
        const dateFile = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

        XLSX.writeFile(wb, `Gastos_${periodLabel}_${dateFile}.xlsx`);

        showNotification('Excel descargado');
    };

    // ============================================
    // Administration Logic
    // ============================================

    let currentAdminTab = 'categories';
    let adminEditContext = null;

    function renderAdminPage() {
        renderAdminPanel(currentAdminTab);
    }

    function renderAdminPanel(tab) {
        const config = StorageManager.getConfig();
        const categories = config.categories;

        if (tab === 'categories') renderCategoriesList(categories);
        else if (tab === 'flavors') {
            populateAdminCategorySelect(elements.adminCategorySelectFlavors, categories);
            renderFlavorsList(config.flavors, elements.adminCategorySelectFlavors.value);
        } else if (tab === 'extras') {
            populateAdminCategorySelect(elements.adminCategorySelectExtras, categories);
            renderExtrasList(config.extras, elements.adminCategorySelectExtras.value);
        } else if (tab === 'observations') {
            populateAdminCategorySelect(elements.adminCategorySelectObs, categories);
            renderObsList(config.observations, elements.adminCategorySelectObs.value);
        } else if (tab === 'expense-cats') {
            renderExpenseCategoriesManager();
        }
    }

    elements.adminTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            elements.adminTabs.forEach(t => t.classList.remove('active'));
            elements.adminPanels.forEach(p => p.classList.remove('active'));
            tab.classList.add('active');
            currentAdminTab = tab.dataset.tab;
            const target = document.getElementById(`panel-${currentAdminTab}`);
            if (target) target.classList.add('active');
            const mc = document.querySelector('.main-content'); if(mc) mc.scrollTop = 0;
            renderAdminPage();
        });
    });

    window.openAddProductModal = function(catId) {
        adminEditContext = { type: 'flavor', id: null, parentId: catId };
        elements.adminModalTitle.innerHTML = '<i data="lucide" name="plus-circle" style="width: 18px; height: 18px; color: #10b981;"></i> Nuevo Producto';
        elements.adminModalBody.innerHTML = `
            <div class="form-group"><label>Nombre del Producto</label><input type="text" autocomplete="off" id="editName" placeholder="Ej: Hamburguesa, Gaseosa, Promo"></div>
            <div class="form-group">
                <label>Tipo de Producto</label>
                <select id="editProdType" style="width: 100%; padding: 8px; border-radius: 6px; border: 1px solid #ccc; font-family: inherit;">
                    <option value="fixed">Precio Fijo (Normal)</option>
                    <option value="open_price">Precio Abierto (Ingresar al cobrar)</option>
                    <option value="quantity">Selector de Cantidad (+ / -)</option>
                    <option value="fixed_with_obs">Precio Fijo + Observacion</option>
                </select>
            </div>
            <div class="form-group" id="editPriceGroup"><label>Precio Unitario ($)</label><input type="number" autocomplete="off" id="editPrice" placeholder="4500" value="0"></div>
        `;
        const typeSelect = document.getElementById('editProdType');
        const priceGroup = document.getElementById('editPriceGroup');
        typeSelect.addEventListener('change', (e) => {
            if (e.target.value === 'open_price' || e.target.value === 'text') {
                priceGroup.style.display = 'none';
            } else {
                priceGroup.style.display = 'block';
            }
        });
        elements.adminModal.classList.add('open');
    };

window.moveAdminItem = function(type, id, direction) {
    const config = StorageManager.getConfig();
    let arr = null;

    if (type === 'category') {
        arr = config.categories;
    } else if (type === 'flavor') {
        arr = config.products;
    }
    
    if (!arr) return;

    const index = arr.findIndex(item => item.id === id);
    if (index === -1) return;

    let targetIndex = -1;

    if (type === 'flavor') {
        const item = arr[index];
        const catId = item.category;
        const catProducts = arr.filter(p => p.category === catId);
        const catIndex = catProducts.findIndex(p => p.id === id);

        if (direction === -1 && catIndex > 0) {
            const targetId = catProducts[catIndex - 1].id;
            targetIndex = arr.findIndex(p => p.id === targetId);
        } else if (direction === 1 && catIndex < catProducts.length - 1) {
            const targetId = catProducts[catIndex + 1].id;
            targetIndex = arr.findIndex(p => p.id === targetId);
        }
    } else {
        if (direction === -1 && index > 0) targetIndex = index - 1;
        if (direction === 1 && index < arr.length - 1) targetIndex = index + 1;
    }

    if (targetIndex !== -1) {
        const temp = arr[index];
        arr[index] = arr[targetIndex];
        arr[targetIndex] = temp;

        StorageManager.saveConfig(config);
        if (typeof renderAdminPage === 'function') renderAdminPage();
    }
};

    function renderCategoriesList(categories) {
        if (!elements.adminCategoriesList) return;
        const config = StorageManager.getConfig();
        const allProducts = config.products || [];

        if (categories.length === 0) {
            elements.adminCategoriesList.innerHTML = `
            <div style="grid-column: 1 / -1; width: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 60px 20px; color: #94a3b8; text-align: center;">
                <i data-lucide="layers" style="width: 64px; height: 64px; margin-bottom: 16px; opacity: 0.5;"></i>
                <h3 style="font-size: 1.1rem; font-weight: 600; color: var(--text-primary); margin-bottom: 8px;">Aun no hay menu</h3>
                <p style="font-size: 0.9rem; max-width: 250px;">Comienza creando tu primera categoria con el boton "+ Nueva Categoria".</p>
            </div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
            return;
        }
        
        elements.adminCategoriesList.innerHTML = categories.map(cat => {
            const catProducts = allProducts.filter(p => p.category === cat.id);
            
            let productsHtml = '';
            if (catProducts.length === 0) {
                productsHtml = `<div style="padding: 20px; text-align: center; color: #94a3b8; font-size: 0.85rem;"><i data-lucide="package-x" style="width: 24px; height: 24px; margin-bottom: 8px; opacity: 0.6;"></i><br>No hay productos en esta categoria</div>`;
            } else {
                productsHtml = catProducts.map(f => {
                    let tLabel = "Precio Fijo";
                    if(f.prodType === "open_price") tLabel = "Precio Abierto";
                    else if(f.prodType === "quantity") tLabel = "Selector de Cantidad";
                    else if(f.prodType === "fixed_with_obs") tLabel = "Precio Fijo + Obs.";
                    
                    return `
                    <div class="admin-item" style="background: rgba(0,0,0,0.03); margin-bottom: 5px; border-radius: 4px; border-left: 3px solid var(--accent-royal); align-items: flex-start; padding: 10px;">
                        <div class="admin-item-info" style="display: flex; flex-direction: column; align-items: flex-start; gap: 2px;">
                            <span style="font-size: 0.95rem; font-weight: 600;">${f.name}</span>
                            <span style="font-size: 0.7rem; color: #64748b; font-weight: 700; text-transform: uppercase;">${tLabel}</span>
                            <span style="font-weight: 800; color: var(--accent-gold); font-size: 0.95rem; margin-top: 2px;">${formatPrice(f.price || 0)}</span>
                        </div>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px; flex-shrink: 0;">
                            <button class="btn-icon" onclick="window.editAdminItem('flavor', '${f.id}', '${cat.id}')" style="padding: 4px; min-width: 28px; height: 28px;" title="Editar">
                                <i data-lucide="edit-2" style="width: 14px; height: 14px;"></i>
                            </button>
                            <button class="btn-icon" onclick="window.moveAdminItem('flavor', '${f.id}', -1)" style="padding: 4px; min-width: 28px; height: 28px;" title="Subir">
                                <i data-lucide="chevron-up" style="width: 14px; height: 14px;"></i>
                            </button>
                            <button class="btn-icon delete-btn" onclick="window.deleteAdminItem('flavor', '${f.id}', '${cat.id}')" style="padding: 4px; min-width: 28px; height: 28px;" title="Eliminar">
                                <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i>
                            </button>
                            <button class="btn-icon" onclick="window.moveAdminItem('flavor', '${f.id}', 1)" style="padding: 4px; min-width: 28px; height: 28px;" title="Bajar">
                                <i data-lucide="chevron-down" style="width: 14px; height: 14px;"></i>
                            </button>
                        </div>
                    </div>
                `; }).join('');
            }

            return `
            <div class="admin-category-card" style="min-width: 320px; flex: 1; background: var(--bg-card); padding: 15px; border-radius: 8px; border: 1px solid var(--border-subtle); margin-bottom: 5px;">
                <div class="admin-item" style="border: none; padding: 0; background: transparent; margin-bottom: 10px;">
                    <div class="admin-item-info">
                        <span style="font-size: 1.1rem; font-weight: 800; color: var(--accent-primary);">${cat.name}</span>
                    </div>
                    <div class="admin-item-actions">
                        <button class="btn-icon" onclick="window.moveAdminItem('category', '${cat.id}', -1)" title="Mover Izquierda">
                            <i data-lucide="chevron-left" style="width: 18px; height: 18px;"></i>
                        </button>
                        <button class="btn-icon" onclick="window.moveAdminItem('category', '${cat.id}', 1)" title="Mover Derecha">
                            <i data-lucide="chevron-right" style="width: 18px; height: 18px;"></i>
                        </button>
                        <button class="btn-icon" onclick="window.editAdminItem('category', '${cat.id}')">
                            <i data-lucide="edit-2"></i>
                        </button>
                        <button class="btn-icon delete-btn" onclick="window.deleteAdminItem('category', '${cat.id}')">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </div>
                </div>
                
                <div class="admin-products-list" style="margin-top: 10px; display: flex; flex-direction: column; gap: 5px;">
                    ${productsHtml}
                    <button class="btn-add-inline" onclick="window.openAddProductModal('${cat.id}')" style="margin-top: 10px; padding: 4px 6px; font-size: 0.85rem; width: 100%; border: 1px dashed var(--accent-royal); color: var(--accent-royal); background: transparent; border-radius: 8px; cursor: pointer;">
                        <i data-lucide="plus" style="width: 16px; height: 16px;"></i> Agregar Producto a ${cat.name}
                    </button>
                </div>
            </div>
            `;
        }).join('');
        
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    function populateAdminCategorySelect(selectEl, categories) {
        if (!selectEl) return;
        const current = selectEl.value;
        selectEl.innerHTML = categories.map(cat => `<option value="${cat.id}" ${cat.id === current ? 'selected' : ''}>${cat.name}</option>`).join('');
    }

    if (elements.adminCategorySelectFlavors) {
        elements.adminCategorySelectFlavors.addEventListener('change', () => {
            renderFlavorsList(StorageManager.getConfig().flavors, elements.adminCategorySelectFlavors.value);
        });
    }

    if (elements.adminCategorySelectExtras) {
        elements.adminCategorySelectExtras.addEventListener('change', () => {
            renderExtrasList(StorageManager.getConfig().extras, elements.adminCategorySelectExtras.value);
        });
    }

    if (elements.adminCategorySelectObs) {
        elements.adminCategorySelectObs.addEventListener('change', () => {
            renderObsList(StorageManager.getConfig().observations, elements.adminCategorySelectObs.value);
        });
    }

    function renderFlavorsList(all, catId) {
        if (!elements.adminFlavorsList) return;
        const config = StorageManager.getConfig();
        const list = (config.products && config.products.length > 0)
            ? config.products.filter(p => p.category === catId)
            : (all[catId] || []);

        elements.adminFlavorsList.innerHTML = list.map(f => `
            <div class="admin-item">
                <div class="admin-item-info">
                    <span>${f.name}</span>
                    <span style="font-weight: 700; color: var(--accent-gold);">${formatPrice(f.price || 0)}</span>
                </div>
                <div class="admin-item-actions">
                    <button class="btn-icon" onclick="window.editAdminItem('flavor', '${f.id}', '${catId}')">
                        <i data-lucide="edit-2"></i>
                    </button>
                    <button class="btn-icon delete-btn" onclick="window.deleteAdminItem('flavor', '${f.id}', '${catId}')">
                        <i data-lucide="trash-2"></i>
                    </button>
                </div>
            </div>
        `).join('');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    function renderExtrasList(all, catId) {
        if (!elements.adminExtrasList) return;
        const list = all[catId] || [];
        elements.adminExtrasList.innerHTML = list.map(e => `
            <div class="admin-item">
                <div class="admin-item-info"><span>${e.name}</span><span>${formatPrice(e.price)}</span></div>
                <div class="admin-item-actions">
                    <button class="btn-icon" onclick="window.editAdminItem('extra', '${e.id}', '${catId}')">
                        <i data-lucide="edit-2"></i>
                    </button>
                    <button class="btn-icon delete-btn" onclick="window.deleteAdminItem('extra', '${e.id}', '${catId}')">
                        <i data-lucide="trash-2"></i>
                    </button>
                </div>
            </div>
        `).join('');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    function renderObsList(all, catId) {
        if (!elements.adminObsList) return;
        const list = all[catId] || [];
        elements.adminObsList.innerHTML = list.map(o => `
            <div class="admin-item">
                <div class="admin-item-info">
                    <span>${o.name}</span>
                    <span>${formatPrice(o.price || 0)}</span>
                </div>
                <div class="admin-item-actions">
                    <button class="btn-icon" onclick="window.editAdminItem('observation', '${o.id}', '${catId}')">
                        <i data-lucide="edit-2"></i>
                    </button>
                    <button class="btn-icon delete-btn" onclick="window.deleteAdminItem('observation', '${o.id}', '${catId}')">
                        <i data-lucide="trash-2"></i>
                    </button>
                </div>
            </div>
        `).join('');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    window.editAdminItem = function (type, id, parentId = null) {
        adminEditContext = { type, id, parentId };
        const config = StorageManager.getConfig();
        const displayType = type === 'flavor' ? 'Producto' : (type === 'category' ? 'Categoria' : (type === 'extra' ? 'Adicional' : 'Observacion'));
        elements.adminModalTitle.innerHTML = `<i data-lucide="edit" style="width: 18px; height: 18px; color: #3b82f6;"></i> Editar ${displayType}`;
        let html = '';
        if (type === 'category') {
            const item = config.categories.find(c => c.id === id);
            const itemType = (item && item.type) ? item.type : 'comida';
            html = '<div class="form-group"><label>Nombre de categoria</label><input type="text" autocomplete="off" id="editName" value="' + item.name + '"></div>';;
        } else if (type === 'flavor') {
            const allProds = getActiveProductsList(config);
            const item = allProds.find(p => p.id === id) || (config.flavors[parentId] && config.flavors[parentId].find(f => f.id === id)) || { name: '', price: 0, prodType: 'fixed' };
            const pt = item.prodType || 'fixed';
            html = `<div class="form-group"><label>Nombre del Producto</label><input type="text" autocomplete="off" id="editName" value="${item.name}"></div>
                    <div class="form-group">
                        <label>Tipo de Producto</label>
                        <select id="editProdType" style="width: 100%; padding: 8px; border-radius: 6px; border: 1px solid #ccc; font-family: inherit;">
                            <option value="fixed" ${pt === 'fixed' ? 'selected' : ''}>Precio Fijo (Normal)</option>
                            <option value="open_price" ${pt === 'open_price' ? 'selected' : ''}>Precio Abierto (Ingresar al cobrar)</option>
                            <option value="quantity" ${pt === 'quantity' ? 'selected' : ''}>Selector de Cantidad (+ / -)</option>
                            <option value="fixed_with_obs" ${pt === 'fixed_with_obs' ? 'selected' : ''}>Precio Fijo + Observacion</option>
                        </select>
                    </div>
                    <div class="form-group" id="editPriceGroup" style="display: ${pt === 'open_price' ? 'none' : 'block'};"><label>Precio Unitario ($)</label><input type="number" autocomplete="off" id="editPrice" value="${item.price || 0}"></div>`;
        } else if (type === 'extra') {
            const item = config.extras[parentId].find(e => e.id === id);
            html = `<div class="form-group"><label>Nombre</label><input type="text" autocomplete="off" id="editName" value="${item.name}"></div>
                    <div class="form-group"><label>Precio ($)</label><input type="number" autocomplete="off" id="editPrice" value="${item.price}"></div>`;
        } else if (type === 'observation') {
            const item = config.observations[parentId].find(o => o.id === id);
            html = `<div class="form-group"><label>Descripcion / Nota</label><input type="text" autocomplete="off" id="editName" value="${item.name}"></div>
                    <div class="form-group"><label>Precio Extra si aplica ($)</label><input type="number" autocomplete="off" id="editPrice" value="${item.price || 0}"></div>`;
        }
        elements.adminModalBody.innerHTML = html;
        if (type === 'flavor') {
            const typeSelect = document.getElementById('editProdType');
            const priceGroup = document.getElementById('editPriceGroup');
            if (typeSelect && priceGroup) {
                typeSelect.addEventListener('change', (e) => {
                    if (e.target.value === 'open_price' || e.target.value === 'text') {
                        priceGroup.style.display = 'none';
                    } else {
                        priceGroup.style.display = 'block';
                    }
                });
            }
        }
        elements.adminModal.classList.add('open');
    };

    window.deleteAdminItem = function (type, id, pId) {
        window.minesofConfirm('Seguro que quieres eliminar este elemento?', () => {
            const config = StorageManager.getConfig();
            if (type === 'category') {
                config.categories = config.categories.filter(c => c.id !== id);
                if (config.products) config.products = config.products.filter(p => p.category !== id);
                if (config.flavors) delete config.flavors[id];
                if (config.extras) delete config.extras[id];
                if (config.observations) delete config.observations[id];
            } else if (type === 'flavor') {
                if (config.products) config.products = config.products.filter(p => p.id !== id);
                if (config.flavors && config.flavors[pId]) {
                    config.flavors[pId] = config.flavors[pId].filter(f => f.id !== id);
                }
            } else if (type === 'extra') {
                if (config.extras && config.extras[pId]) config.extras[pId] = config.extras[pId].filter(e => e.id !== id);
            } else if (type === 'observation') {
                if (config.observations && config.observations[pId]) config.observations[pId] = config.observations[pId].filter(o => o.id !== id);
            }

            StorageManager.saveConfig(config);
            renderAdminPage();
            renderPosCategories();
            renderPosProducts();
            showNotification('Eliminado correctamente');
        });
    };

    if (elements.cancelAdminModal) elements.cancelAdminModal.onclick = () => elements.adminModal.classList.remove('open');
        if (elements.closeAdminModal) elements.closeAdminModal.onclick = () => elements.adminModal.classList.remove('open');
    if (elements.confirmAdminModal) {
        elements.confirmAdminModal.onclick = () => {
            const config = StorageManager.getConfig();
            const { type, id, parentId } = adminEditContext;
            const name = document.getElementById('editName').value.trim();
            if (!name) {
                showNotification('Ingresa un nombre vÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡lido', 'error');
                return;
            }

            if (type === 'category') {
                if (id) {
                    const cat = config.categories.find(c => c.id === id);
                    if (cat) {
                        cat.name = name;
                    }
                } else {
                    const newId = 'cat_' + Date.now();
                    config.categories.push({ id: newId, name, active: true });
                    if (!config.flavors) config.flavors = {};
                    config.flavors[newId] = [];
                    if (!config.extras) config.extras = {};
                    config.extras[newId] = [];
                    if (!config.observations) config.observations = {};
                    config.observations[newId] = [];
                }
            } else if (type === 'flavor') {
                const typeEl = document.getElementById('editProdType');
                const prodType = typeEl ? typeEl.value : 'fixed';
                const price = (prodType === 'open_price' || prodType === 'text') ? 0 : (+document.getElementById('editPrice').value || 0);

                if (!config.products) config.products = [];

                if (id) {
                    const prod = config.products.find(p => p.id === id);
                    if (prod) {
                        prod.name = name;
                        prod.price = price;
                        prod.prodType = prodType;
                    }
                    if (config.flavors && config.flavors[parentId]) {
                        const fl = config.flavors[parentId].find(f => f.id === id);
                        if (fl) {
                            fl.name = name;
                            fl.price = price;
                            fl.prodType = prodType;
                        }
                    }
                } else {
                    const newId = 'prod_' + Date.now();
                    const newProd = { id: newId, name, price, category: parentId, active: true, prodType };
                    config.products.push(newProd);

                    if (!config.flavors) config.flavors = {};
                    if (!config.flavors[parentId]) config.flavors[parentId] = [];
                    config.flavors[parentId].push({ id: newId, name, price, active: true, prodType });
                }
            } else if (type === 'extra') {
                if (!config.extras) config.extras = {};
                if (!config.extras[parentId]) config.extras[parentId] = [];
                const price = +document.getElementById('editPrice').value || 0;
                if (id) {
                    const e = config.extras[parentId].find(x => x.id === id);
                    if (e) {
                        e.name = name;
                        e.price = price;
                    }
                } else {
                    config.extras[parentId].push({ id: 'e_' + Date.now(), name, price, active: true });
                }
            } else if (type === 'observation') {
                if (!config.observations) config.observations = {};
                if (!config.observations[parentId]) config.observations[parentId] = [];
                const price = +document.getElementById('editPrice').value || 0;
                if (id) {
                    const o = config.observations[parentId].find(x => x.id === id);
                    if (o) {
                        o.name = name;
                        o.price = price;
                    }
                } else {
                    config.observations[parentId].push({ id: 'o_' + Date.now(), name, price, active: true });
                }
            }

            StorageManager.saveConfig(config);
            elements.adminModal.classList.remove('open');
            renderAdminPage();
            renderPosCategories();
            renderPosProducts();
            showNotification('Guardado correctamente');
        };
    }

    if (elements.addCategoryBtn) {
        elements.addCategoryBtn.onclick = () => {
            adminEditContext = { type: 'category', id: null };
            elements.adminModalTitle.innerHTML = '<i data="lucide" name="plus-circle" style="width: 18px; height: 18px; color: #10b981;"></i> Nueva Categoria';
            elements.adminModalBody.innerHTML = '<div class="form-group"><label>Nombre de categoria</label><input type="text" autocomplete="off" id="editName" placeholder="Ej: Comida, Bebidas, Combos"></div>';
            elements.adminModal.classList.add('open');
        };
    }
    
    if (elements.addFlavorBtn) {
        elements.addFlavorBtn.onclick = () => {
            const catId = elements.adminCategorySelectFlavors ? elements.adminCategorySelectFlavors.value : 'panaderia';
            adminEditContext = { type: 'flavor', id: null, parentId: catId };
            elements.adminModalTitle.innerHTML = '<i data="lucide" name="plus-circle" style="width: 18px; height: 18px; color: #10b981;"></i> Nuevo Producto';
            elements.adminModalBody.innerHTML = `
                <div class="form-group"><label>Nombre del Producto</label><input type="text" autocomplete="off" id="editName" placeholder="Ej: Hamburguesa, Gaseosa, Promo"></div>
                <div class="form-group"><label>Precio Unitario ($)</label><input type="number" autocomplete="off" id="editPrice" placeholder="4500" value="0"></div>
            `;
            elements.adminModal.classList.add('open');
        };
    }

    if (elements.addExtraBtn) {
        elements.addExtraBtn.onclick = () => {
            const catId = elements.adminCategorySelectExtras.value;
            adminEditContext = { type: 'extra', id: null, parentId: catId };
            elements.adminModalTitle.textContent = 'Nuevo Adicional';
            elements.adminModalBody.innerHTML = `<div class="form-group"><label>Nombre</label><input type="text" autocomplete="off" id="editName"></div>
                <div class="form-group"><label>Precio</label><input type="number" autocomplete="off" id="editPrice" value="0"></div>`;
            elements.adminModal.classList.add('open');
        };
    }

    if (elements.addObsBtn) {
        elements.addObsBtn.onclick = () => {
            const catId = elements.adminCategorySelectObs.value;
            adminEditContext = { type: 'observation', id: null, parentId: catId };
            elements.adminModalTitle.textContent = 'Nueva Observacion';
            elements.adminModalBody.innerHTML = `
                <div class="form-group"><label>Nombre</label><input type="text" autocomplete="off" id="editName"></div>
                <div class="form-group"><label>Precio</label><input type="number" autocomplete="off" id="editPrice" value="0"></div>
            `;
            elements.adminModal.classList.add('open');
        };
    }

    // ============================================
    // Security / Password Logic
    // ============================================

    if (elements.confirmAdminLogin) {
        const handleLogin = () => {
                        const config = StorageManager.getConfig();
            const input = elements.adminPasswordInput.value;

            if (input === config.adminPassword) {
                state.isAdminAuthenticated = true;
                elements.adminLoginModal.classList.remove('open');
                showNotification('Acceso concedido');

                // Trigger the pending action (like deletion)
                if (state.pendingAdminAction) {
                    state.pendingAdminAction();
                    state.pendingAdminAction = null;
                }

                // Trigger the pending navigation to protected page
                if (state.pendingAdminPage) {
                    const targetBtn = Array.from(elements.drawerItems).find(i => i.dataset.page === state.pendingAdminPage);
                    if (targetBtn) targetBtn.click();
                    state.pendingAdminPage = null;
                }
            } else {
                showNotification('Contrase\u00f1a incorrecta', 'error');
                elements.adminPasswordInput.value = '';
                elements.adminPasswordInput.focus();
            }
        };

        elements.confirmAdminLogin.addEventListener('click', handleLogin);
        console.log('Listener for confirmAdminLogin attached');
        elements.adminPasswordInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') handleLogin();
        });
    }

    if (elements.closeAdminLoginModal) {
        elements.closeAdminLoginModal.addEventListener('click', () => {
            elements.adminLoginModal.classList.remove('open');
            state.pendingAdminPage = null;
            state.pendingAdminAction = null;
        });
    }

    if (elements.saveAdminPasswordBtn) {
        elements.saveAdminPasswordBtn.addEventListener('click', () => {
            const newPass = elements.newAdminPassword.value;
            const confirmPass = elements.confirmAdminPassword.value;

            if (newPass.length < 4) {
                showNotification('La contraseÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±a debe tener al menos 4 caracteres', 'error');
                return;
            }

            if (newPass !== confirmPass) {
                showNotification('Las contraseÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±as no coinciden', 'error');
                return;
            }

            const config = StorageManager.getConfig();
            config.adminPassword = newPass;
            StorageManager.saveConfig(config);

            showNotification('ContraseÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â±a actualizada correctamente');
            elements.newAdminPassword.value = '';
            elements.confirmAdminPassword.value = '';
        });
    }

    // Modal Handlers for Dynamic Product Types
    
    // Quantity Modal Plus/Minus Buttons
    const qtyModalMinus = document.getElementById('qtyModalMinus');
    const qtyModalPlus = document.getElementById('qtyModalPlus');
    const quantityInputValue = document.getElementById('quantityInputValue');
    if (qtyModalMinus && qtyModalPlus && quantityInputValue) {
        qtyModalMinus.addEventListener('click', () => {
            let v = parseInt(quantityInputValue.value) || 0;
            if (v > 0) quantityInputValue.value = v - 1;
        });
        qtyModalPlus.addEventListener('click', () => {
            let v = parseInt(quantityInputValue.value) || 0;
            quantityInputValue.value = v + 1;
        });
    }

    const quantityConfirmBtn = document.getElementById('quantityConfirmBtn');
    if (quantityConfirmBtn) {
        quantityConfirmBtn.addEventListener('click', () => {
            const qty = parseInt(document.getElementById('quantityInputValue').value);
            
            // If 0 or empty, remove item
            if (isNaN(qty) || qty <= 0) {
                if (activeQuantityProductId) {
                    const clientId = state.activeClient;
                    const existingIndex = state.cart.findIndex(item => item.productId === activeQuantityProductId && item.clientName === clientId);
                    if (existingIndex !== -1) {
                        state.cart.splice(existingIndex, 1);
                        renderSplitUI();
                        if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                    }
                }
                document.getElementById('quantityModal').classList.remove('open');
                activeQuantityProductId = null;
                return;
            }

            if (!activeQuantityProductId) return;

            const config = StorageManager.getConfig();
            const product = getActiveProductsList(config).find(p => p.id === activeQuantityProductId);
            if (product) {
                const clientId = state.activeClient;
                const existingIndex = state.cart.findIndex(item => item.productId === activeQuantityProductId && item.clientName === clientId);
                
                if (existingIndex !== -1) {
                    state.cart[existingIndex].qty = qty;
                    state.cart[existingIndex].subtotal = (product.price || 0) * qty;
                } else {
                    state.cart.push({
                        id: 'cart_' + Date.now(),
                        productId: product.id,
                        name: product.name,
                        unitPrice: product.price || 0,
                        qty: qty,
                        subtotal: (product.price || 0) * qty,
                        clientName: clientId,
                        categoryId: product.category,
                        notes: ''
                    });
                }
                renderSplitUI();
                if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
            }
            document.getElementById('quantityModal').classList.remove('open');
            activeQuantityProductId = null;
        });
    }

    const openPriceInputEl = document.getElementById('openPriceInput');
    if (openPriceInputEl) {
        openPriceInputEl.addEventListener('input', function(e) {
            let value = e.target.value.replace(/\D/g, '');
            if (value !== '') {
                value = parseInt(value, 10).toLocaleString('es-CO');
            }
            e.target.value = value;
        });
    }

    const openPriceConfirmBtn = document.getElementById('openPriceConfirmBtn');
    if (openPriceConfirmBtn) {
        openPriceConfirmBtn.addEventListener('click', () => {
            const inputVal = document.getElementById('openPriceInput').value;
            const cleanVal = inputVal.replace(/\D/g, '');
            const price = parseFloat(cleanVal);
            
            // If empty or 0, remove the item
            if (inputVal === '' || price === 0) {
                if (activeOpenPriceProductId) {
                    const clientId = state.activeClient;
                    const existingIndex = state.cart.findIndex(item => item.productId === activeOpenPriceProductId && item.clientName === clientId);
                    if (existingIndex !== -1) {
                        state.cart.splice(existingIndex, 1);
                        renderSplitUI();
                        if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                    }
                }
                document.getElementById('openPriceModal').classList.remove('open');
                activeOpenPriceProductId = null;
                return;
            }

            if (isNaN(price) || price < 0) {
                showNotification('Ingresa un valor vÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡lido', 'error');
                return;
            }
            if (!activeOpenPriceProductId) return;

            const config = StorageManager.getConfig();
            const product = getActiveProductsList(config).find(p => p.id === activeOpenPriceProductId);
            if (product) {
                const clientId = state.activeClient;
                const existingIndex = state.cart.findIndex(item => item.productId === activeOpenPriceProductId && item.clientName === clientId);
                
                if (existingIndex !== -1) {
                    // Update existing
                    state.cart[existingIndex].unitPrice = price;
                    state.cart[existingIndex].subtotal = price;
                } else {
                    // Add new
                    state.cart.push({
                        id: 'cart_' + Date.now(),
                        productId: product.id,
                        name: product.name,
                        unitPrice: price,
                        qty: 1,
                        subtotal: price,
                        clientName: clientId,
                        categoryId: product.category,
                        notes: ''
                    });
                }
                renderSplitUI();
                if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
            }
            document.getElementById('openPriceModal').classList.remove('open');
            activeOpenPriceProductId = null;
        });
    }

    const textInputConfirmBtn = document.getElementById('textInputConfirmBtn');
    if (textInputConfirmBtn) {
        textInputConfirmBtn.addEventListener('click', () => {
            const text = document.getElementById('textInputValue').value.trim();
            
            // If empty text, remove the item
            if (!text) {
                if (activeTextProductId) {
                    const clientId = state.activeClient;
                    const existingIndex = state.cart.findIndex(item => item.productId === activeTextProductId && item.clientName === clientId);
                    if (existingIndex !== -1) {
                        state.cart.splice(existingIndex, 1);
                        renderSplitUI();
                        if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                    }
                }
                document.getElementById('textInputModal').classList.remove('open');
                activeTextProductId = null;
                return;
            }

            if (!activeTextProductId) return;

            const config = StorageManager.getConfig();
            const product = getActiveProductsList(config).find(p => p.id === activeTextProductId);
            if (product) {
                const clientId = state.activeClient;
                const existingIndex = state.cart.findIndex(item => item.productId === activeTextProductId && item.clientName === clientId);
                
                if (existingIndex !== -1) {
                    // Update existing
                    state.cart[existingIndex].notes = text;
                    state.cart[existingIndex].name = text;
                } else {
                    // Add new
                    state.cart.push({
                        id: 'cart_' + Date.now(),
                        productId: product.id,
                        name: text,
                        unitPrice: 0,
                        qty: 1,
                        subtotal: 0,
                        clientName: clientId,
                        categoryId: product.category,
                        notes: text
                    });
                }
                renderSplitUI();
                if (typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
            }
            document.getElementById('textInputModal').classList.remove('open');
            activeTextProductId = null;
        });
    }

    // ============================================
    // Order Counter Reset Functionality
    // ============================================
    const resetOrderCounterBtn = document.getElementById('resetOrderCounterBtn');
    const currentOrderCounterEl = document.getElementById('currentOrderCounter');

    // Load and display current counter
        async function loadCurrentOrderCounter() {
        const currentOrderCounterEl = document.getElementById('currentOrderCounter');
        if (!currentOrderCounterEl) return;
        
        if (typeof db !== 'undefined') {
            try {
                const doc = await getDbCollection('minesof_settings').doc('global_config').get();
                if (doc.exists) {
                    const data = doc.data();
                    const cloudCounter = data.orderCounter || 0;
                    currentOrderCounterEl.textContent = '#' + String(parseInt(cloudCounter)).padStart(3, '0');
                    return;
                }
            } catch(e) {}
        }

        const localCounter = localStorage.getItem('galeria_order_counter') || '0';
        currentOrderCounterEl.textContent = '#' + String(parseInt(localCounter)).padStart(3, '0');
    }

    // Reset counter to 0
    async function resetOrderCounter() {
        try {
            // Reset local storage
            localStorage.setItem('galeria_order_counter', '0');
            localStorage.setItem('galeria_last_order_date', new Date().toDateString());

            if (typeof db !== 'undefined') {
                try {
                    await getDbCollection(STORAGE_KEYS.SETTINGS).doc('global_config').set({
                        orderCounter: 0,
                        lastWipe: Date.now()
                    }, { merge: true });
                } catch(e) {}
            }

            showNotification('Contador reiniciado a #001. Actualizando sistema...');
            setTimeout(() => window.location.reload(), 1500);
        } catch (error) {
            console.error('Error resetting counter:', error);
            showNotification('Error al reiniciar: ' + error.message, 'error');
        }
    }

    if (resetOrderCounterBtn) {
        resetOrderCounterBtn.addEventListener('click', () => {
            window.minesofConfirm('Estas seguro que deseas reiniciar el contador de pedidos a #001?', async () => {
                await resetOrderCounter();
            });
        });
    }

    // Load counter when navigating to admin
    const originalRenderAdminPage = typeof renderAdminPage === 'function' ? renderAdminPage : null;
    if (originalRenderAdminPage) {
        const extendedRenderAdmin = () => {
            originalRenderAdminPage();
            loadCurrentOrderCounter();
        };
        // Override if needed - for now just call on page load
    }

    // Also load on DOMContentLoaded for admin page
    loadCurrentOrderCounter();

    function showNotification(message, type = 'success') {
        const notification = document.createElement('div');
        const isError = type === 'error';
        const bgColor = isError ? 'rgba(220, 38, 38, 0.95)' : 'rgba(16, 185, 129, 0.95)';
        const icon = isError ? 'alert-circle' : 'check-circle';

        // Vibration and sound for errors
        if (isError) {
            // Vibrate (mobile devices)
            if (navigator.vibrate) {
                navigator.vibrate([100, 50, 100]); // vibrate-pause-vibrate pattern
            }

            // Play error sound using Web Audio API
            try {
                const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                const oscillator = audioCtx.createOscillator();
                const gainNode = audioCtx.createGain();

                oscillator.connect(gainNode);
                gainNode.connect(audioCtx.destination);

                oscillator.frequency.value = 400; // Low frequency buzz
                oscillator.type = 'square';
                gainNode.gain.value = 0.3;

                oscillator.start();
                gainNode.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.2);
                oscillator.stop(audioCtx.currentTime + 0.2);
            } catch (e) {
                console.log('Audio not supported');
            }
        }

        notification.style.cssText = `
            position: fixed;
            top: 20px;
            left: 50%;
            transform: translateX(-50%);
            background: ${bgColor};
            backdrop-filter: blur(10px);
            color: white;
            padding: 12px 24px;
            border-radius: 12px;
            font-weight: 600;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2), 0 8px 10px -6px rgba(0, 0, 0, 0.15);
            z-index: 9999;
            animation: slideInDown 0.3s ease-out;
            display: flex;
            align-items: center;
            gap: 10px;
        `;
        notification.innerHTML = `<i data-lucide="${icon}" style="width: 18px; height: 18px;"></i> ${message}`;
        document.body.appendChild(notification);
        if (typeof lucide !== 'undefined') lucide.createIcons();

        setTimeout(() => {
            notification.style.animation = 'fadeOutUp 0.3s ease-in forwards';
            setTimeout(() => notification.remove(), 300);
        }, 2500);
    }

    // Initialize Cloud Sync (wait for auth to guarantee correct tenant)
    if (window.auth) {
        window.auth.onAuthStateChanged((user) => {
            if (user && typeof StorageManager.initCloudSync === 'function') {
                StorageManager.initCloudSync(
                    // Orders & Expenses callback
                    () => {
                        console.log('Cloud sync: Orders or Expenses updated. Updating UI...');
                        try {
                            if (state.currentPage === 'checkout') renderCheckoutPage();
                        } catch (e) { console.error('Error in renderCheckoutPage during sync', e); }
                        
                        try {
                            if (state.currentPage === 'history') renderHistoryPage();
                        } catch (e) { console.error('Error in renderHistoryPage during sync', e); }
                        
                        try {
                            if (state.currentPage === 'new-order') {
                                if(typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                            }
                        } catch (e) { console.error('Error in new-order render during sync', e); }
                        
                        try {
                            if (state.currentPage === 'expenses') renderExpensesPage();
                        } catch (e) { console.error('Error in renderExpensesPage during sync', e); }
                    },
                    // Config callback (Admin changes from other devices)
                    () => {
                        if (state.currentPage === 'admin') renderAdminPage();
                        if (state.currentPage === 'expenses') renderExpensesPage();
                        if (state.currentPage === 'new-order') {
                            initializeCategories();
                            renderPosClientTabs();
                            if(typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
                        }
                        renderPosClientTabs();
                        renderPosCategories();
                        renderPosProducts();
                        updateAppBranding();
                        console.log('Config synced from cloud');
                        window.isCloudConfigSynced = true;
                        window.dispatchEvent(new CustomEvent('cloudConfigSynced'));
                    },
                    // Print callback (Remote print from other devices) - DISABLED
                    null
                );
            }
        });
    }

    // Initialize
    renderPosCategories();
    renderPosProducts();
    if(typeof updateOrderTotal === "function") updateOrderTotal(); else renderPosCart();
    
  });






































































    // ============================================
        // ============================================
    // System Data Management
    // ============================================
    window.clearSystemData = function() {
        const msg1 = "ADVERTENCIA CRITICA \n\nEstas seguro de querer BORRAR TODO el historial de pedidos y gastos?\n\n- Esta accion es irreversible.\n- Tu catalogo (productos, categorias) NO se borrara.\n- Tu contador de pedidos volvera a cero.";
        window.minesofConfirm(msg1, async () => {
            window.minesofPrompt("Escribe BORRAR en mayusculas para confirmar la eliminacion:", async (confirmWord) => {
                if (confirmWord !== "BORRAR") {
                    showNotification("Eliminacion cancelada.", "error");
                    return;
                }

            const overlay = document.createElement('div');
        overlay.style.position = 'fixed';
        overlay.style.top = '0'; overlay.style.left = '0'; overlay.style.width = '100vw'; overlay.style.height = '100vh';
        overlay.style.backgroundColor = 'rgba(0,0,0,0.85)';
        overlay.style.color = 'white';
        overlay.style.display = 'flex';
        overlay.style.flexDirection = 'column';
        overlay.style.justifyContent = 'center';
        overlay.style.alignItems = 'center';
        overlay.style.zIndex = '999999';
        overlay.style.fontFamily = 'system-ui, sans-serif';
        overlay.innerHTML = `
            <div style="border: 4px solid rgba(255,255,255,0.2); border-top: 4px solid white; border-radius: 50%; width: 50px; height: 50px; animation: spinLoader 1s linear infinite; margin-bottom: 20px;"></div>
            <h2 id="wipeProgressText" style="margin: 0; padding: 0;">Limpiando la nube...</h2>
            <p style="margin-top: 10px; color: #ccc;">No cierres la aplicaciÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â³n</p>
            <style>@keyframes spinLoader { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>
        `;
        document.body.appendChild(overlay);

        try {
            if (typeof db !== 'undefined') {
                const batchDelete = async (snapshot, name) => {
                    if (snapshot.empty) return;
                    document.getElementById('wipeProgressText').textContent = `Borrando ${name}...`;
                    const chunks = [];
                    for (let i = 0; i < snapshot.docs.length; i += 450) {
                        chunks.push(snapshot.docs.slice(i, i + 450));
                    }
                    for (let chunk of chunks) {
                        const batch = db.batch();
                        chunk.forEach(doc => batch.delete(doc.ref));
                        await batch.commit();
                    }
                };
                
                try {
                    document.getElementById('wipeProgressText').textContent = 'Buscando pedidos...';
                    const ordersSnapshot = await getDbCollection(STORAGE_KEYS.ORDERS).get();
                    await batchDelete(ordersSnapshot, 'pedidos');
                } catch(e) { console.error('Error deleting orders', e); }
                
                try {
                    document.getElementById('wipeProgressText').textContent = 'Buscando gastos...';
                    const expensesSnapshot = await getDbCollection(STORAGE_KEYS.EXPENSES).get();
                    await batchDelete(expensesSnapshot, 'gastos');
                } catch(e) { console.error('Error deleting expenses', e); }
            }

            localStorage.removeItem(STORAGE_KEYS.ORDERS);
            localStorage.removeItem(STORAGE_KEYS.EXPENSES);
            localStorage.setItem('galeria_order_counter', '0');
            
            if (typeof db !== 'undefined') {
                try {
                    await getDbCollection(STORAGE_KEYS.SETTINGS).doc('global_config').set({
                        orderCounter: 0,
                        lastWipe: Date.now()
                    }, { merge: true });
                } catch(e) {}
            }

            document.getElementById('wipeProgressText').textContent = 'ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡Listo!';
            showNotification("Todo el historial ha sido borrado.");
            setTimeout(() => {
                window.location.reload();
            }, 1000);
        } catch(error) {
            console.error(error);
            overlay.remove();
            showNotification("Error al limpiar historial", "error");
        }
        }); // Close minesofPrompt
        }); // Close minesofConfirm
    };












    // ============================================
    // Balance
    // ============================================
    
        function renderBalancePage() {
        const periodSelect = document.getElementById('balancePeriodSelect');
        const period = periodSelect ? periodSelect.value : 'today';
        const tbody = document.getElementById('balanceTableBody');
        if (!tbody) return;

        let orders = [];
        let expenses = [];

        switch (period) {
            case 'today':
                orders = StorageManager.getTodayOrders();
                expenses = StorageManager.getTodayExpenses();
                break;
            case 'date':
                const datePicker = document.getElementById('balanceDatePicker');
                const filterDate = datePicker ? datePicker.value : null;
                orders = filterDate ? StorageManager.getOrdersByDate(filterDate) : StorageManager.getTodayOrders();
                expenses = filterDate ? StorageManager.getExpensesByDate(filterDate) : StorageManager.getTodayExpenses();
                break;
            case 'month':
                orders = StorageManager.getCurrentMonthOrders();
                expenses = StorageManager.getCurrentMonthExpenses();
                break;
            case 'specific-month':
                const monthPicker = document.getElementById('balanceMonthPicker');
                const filterMonth = monthPicker ? monthPicker.value : null;
                orders = filterMonth ? StorageManager.getOrdersByMonth(filterMonth) : StorageManager.getCurrentMonthOrders();
                expenses = filterMonth ? StorageManager.getExpensesByMonth(filterMonth) : StorageManager.getCurrentMonthExpenses();
                break;
            case 'total':
                orders = StorageManager.getOrders();
                expenses = StorageManager.getExpenses();
                break;
            default:
                orders = StorageManager.getTodayOrders();
                expenses = StorageManager.getTodayExpenses();
        }

        orders = orders.filter(o => !o.isPartial && o.paid);

        const dailyData = {};

        // Aggregate Sales
        orders.forEach(o => {
            const dateObj = new Date(o.createdAt);
            const dateStr = dateObj.toLocaleDateString('es-CO', { year: 'numeric', month: '2-digit', day: '2-digit' });
            if (!dailyData[dateStr]) dailyData[dateStr] = { sales: 0, expenses: 0, timestamp: dateObj.getTime() };
            dailyData[dateStr].sales += o.totalPrice || 0;
        });

        // Aggregate Expenses
        expenses.forEach(e => {
            const dateObj = new Date(e.date || e.createdAt);
            const dateStr = dateObj.toLocaleDateString('es-CO', { year: 'numeric', month: '2-digit', day: '2-digit' });
            if (!dailyData[dateStr]) dailyData[dateStr] = { sales: 0, expenses: 0, timestamp: dateObj.getTime() };
            dailyData[dateStr].expenses += e.amount || 0;
        });

        const sortedDates = Object.keys(dailyData).sort((a, b) => dailyData[b].timestamp - dailyData[a].timestamp);

        if (sortedDates.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; padding: 2rem; color: var(--text-muted);">No hay datos en este periodo</td></tr>';
            const sumBlock = document.getElementById('balanceSummaryBlock');
            if (sumBlock) sumBlock.innerHTML = '';
            const tRow = document.getElementById('balanceTotalsRow');
            if (tRow) tRow.remove();
            const tSpacer = document.getElementById('balanceTotalsSpacer');
            if (tSpacer) tSpacer.remove();
            if (window.chartsMap) {
                Object.values(window.chartsMap).forEach(c => c.destroy());
                window.chartsMap = {};
            }
            ['chartTotalBalance', 'chartTotalVentas', 'chartTotalGastos'].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.innerHTML = '';
            });
            return;
        }

        let rowsHtml = '';
        let totalSales = 0;
        let totalExp = 0;

        sortedDates.forEach((dateStr, idx) => {
            const data = dailyData[dateStr];
            const balance = data.sales - data.expenses;
            
            totalSales += data.sales;
            totalExp += data.expenses;
            
            const balanceColor = balance >= 0 ? '#16a34a' : '#dc2626';
            const rowBg = idx % 2 === 0 ? 'var(--bg-primary, #ffffff)' : 'rgba(241, 245, 249, 0.4)';

            rowsHtml += `
                <tr style="border-top: 1px solid var(--border-subtle); background: ${rowBg};">
                    <td style="padding: 10px 12px; font-weight: 600; color: var(--text-primary);">${dateStr}</td>
                    <td style="padding: 10px 12px; text-align: right; color: var(--text-primary);">${formatPrice(data.sales)}</td>
                    <td style="padding: 10px 12px; text-align: right; color: var(--text-primary);">${formatPrice(data.expenses)}</td>
                    <td style="padding: 10px 12px; text-align: right; font-weight: 800; color: ${balanceColor};">${formatPrice(balance)}</td>
                    <td style="padding: 10px 12px; text-align: center;">
                        <button onclick="window.openBalanceDetails('${dateStr}')" style="background: linear-gradient(135deg, #3b82f6, #2563eb); border: none; border-radius: 20px; box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2), inset 0 -2px 0 rgba(0,0,0,0.1), inset 0 2px 4px rgba(255,255,255,0.3); color: white; width: 60px; height: 26px; font-size: 0.75rem; font-weight: 600; display: inline-flex; align-items: center; justify-content: center; cursor: pointer; transition: transform 0.1s ease, box-shadow 0.1s ease;">
                            Ver
                        </button>
                    </td>
                </tr>
            `;
        });

        const totalBalance = totalSales - totalExp;
        const totalColor = totalBalance >= 0 ? '#16a34a' : '#dc2626';

        const darkColor = totalBalance >= 0 ? '#4ade80' : '#f87171';
        
        const summaryBlock = document.getElementById("balanceSummaryBlock");
        if (summaryBlock) summaryBlock.innerHTML = "";
        
        const cBal = document.getElementById("chartTotalBalance");
        if (cBal) cBal.innerHTML = `<span style="display:block; color:${darkColor}; font-size:1.6rem; font-weight:800;">${formatPrice(totalBalance)}</span>`;
        
        const cVen = document.getElementById("chartTotalVentas");
        if (cVen) cVen.innerHTML = `<span style="display:block; color:#38bdf8; font-size:1.6rem; font-weight:800;">${formatPrice(totalSales)}</span>`;
        
        const cGas = document.getElementById("chartTotalGastos");
        if (cGas) cGas.innerHTML = `<span style="display:block; color:#f43f5e; font-size:1.6rem; font-weight:800;">${formatPrice(totalExp)}</span>`;
        
                                const chartDates = [...sortedDates].reverse();
                const daysOfWeek = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
                const chartLabels = chartDates.map(d => {
                    const parts = d.split('/');
                    if (parts.length === 3) {
                        const dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                        const shortYear = parts[2].slice(-2);
                        const dayName = daysOfWeek[dateObj.getDay()];
                        return `${parts[0]}/${parts[1]}/${shortYear} (${dayName})`;
                    }
                    return d;
                });
        const salesData = chartDates.map(d => dailyData[d].sales);
        const expData = chartDates.map(d => dailyData[d].expenses);
        const balData = chartDates.map(d => dailyData[d].sales - dailyData[d].expenses);

        if (window.chartsMap) {
            Object.values(window.chartsMap).forEach(c => c.destroy());
        }
        window.chartsMap = {};

        if (typeof Chart !== 'undefined' && chartDates.length > 0) {
            const dynamicWidth = (chartDates.length * 20) + 'px';
            const wBal = document.getElementById('wrapperBalance'); if (wBal) wBal.style.width = dynamicWidth;
            const wVen = document.getElementById('wrapperVentas'); if (wVen) wVen.style.width = dynamicWidth;
            const wGas = document.getElementById('wrapperGastos'); if (wGas) wGas.style.width = dynamicWidth;

            const commonOptions = {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        if (typeof window.openBalanceDetails === 'function') {
                            window.openBalanceDetails(chartDates[index]);
                        }
                    }
                },
                plugins: { legend: { display: false } },
                scales: {
                    x: {
                        ticks: {
                            callback: function(val, index) {
                                return chartDates[index].split('/')[0];
                            }
                        }
                    },
                    y: { beginAtZero: true }
                }
            };

            const ctxBal = document.getElementById('chartBalance');
            if (ctxBal) {
                const balOptions = JSON.parse(JSON.stringify(commonOptions));
                balOptions.plugins.title = { display: true, text: 'Comportamiento del Balance', color: '#94a3b8' };
                balOptions.scales.y.grid = {
                    color: (ctx) => ctx.tick.value === 0 ? '#ef4444' : 'rgba(0,0,0,0.1)',
                    lineWidth: (ctx) => ctx.tick.value === 0 ? 2 : 1
                };
                // We must copy functions because JSON stringify lost them
                balOptions.onClick = commonOptions.onClick;
                balOptions.scales.x.ticks.callback = commonOptions.scales.x.ticks.callback;

                const balPointColors = balData.map(v => v < 0 ? '#ef4444' : '#10b981');

                window.chartsMap.bal = new Chart(ctxBal, {
                    type: 'line',
                    data: {
                        labels: chartLabels,
                        datasets: [{
                            label: 'Balance Neto ($)',
                            data: balData,
                            borderColor: '#10b981',
                            backgroundColor: 'rgba(16, 185, 129, 0.1)',
                            borderWidth: 2,
                            pointBackgroundColor: balPointColors,
                            pointBorderColor: balPointColors,
                            pointRadius: 5,
                            pointHoverRadius: 7,
                            fill: true,
                            tension: 0.3
                        }]
                    },
                    options: balOptions
                });
            }

            const ctxSales = document.getElementById('chartVentas');
            if (ctxSales) {
                const salesOptions = JSON.parse(JSON.stringify(commonOptions));
                salesOptions.plugins.title = { display: true, text: 'Comportamiento de Ventas', color: '#94a3b8' };
                salesOptions.onClick = commonOptions.onClick;
                salesOptions.scales.x.ticks.callback = commonOptions.scales.x.ticks.callback;

                window.chartsMap.sales = new Chart(ctxSales, {
                    type: 'line',
                    data: {
                        labels: chartLabels,
                        datasets: [{
                            label: 'Ventas Totales ($)',
                            data: salesData,
                            borderColor: '#38bdf8',
                            backgroundColor: 'rgba(56, 189, 248, 0.1)',
                            borderWidth: 2,
                            fill: true,
                            tension: 0.3
                        }]
                    },
                    options: salesOptions
                });
            }

            const ctxExp = document.getElementById('chartGastos');
            if (ctxExp) {
                const expOptions = JSON.parse(JSON.stringify(commonOptions));
                expOptions.plugins.title = { display: true, text: 'Comportamiento de Gastos', color: '#94a3b8' };
                expOptions.onClick = commonOptions.onClick;
                expOptions.scales.x.ticks.callback = commonOptions.scales.x.ticks.callback;

                window.chartsMap.exp = new Chart(ctxExp, {
                    type: 'line',
                    data: {
                        labels: chartLabels,
                        datasets: [{
                            label: 'Gastos ($)',
                            data: expData,
                            borderColor: '#f43f5e',
                            backgroundColor: 'rgba(244, 63, 94, 0.1)',
                            borderWidth: 2,
                            fill: true,
                            tension: 0.3
                        }]
                    },
                    options: expOptions
                });
            }
        }

        const thead = document.querySelector("#balanceTable thead");
        if (thead) {
            let tRow = document.getElementById("balanceTotalsRow");
            if (!tRow) {
                tRow = document.createElement("tr");
                tRow.id = "balanceTotalsRow";
                tRow.style = "background: #0f172a; box-shadow: 0 4px 6px rgba(0,0,0,0.1);";
                thead.insertBefore(tRow, thead.firstChild);
            }
            let spacer = document.getElementById("balanceTotalsSpacer");
            if (!spacer) {
                spacer = document.createElement("tr");
                spacer.id = "balanceTotalsSpacer";
                spacer.style = "height: 12px; background: transparent; border: none;";
                spacer.innerHTML = "<th colspan='5' style='border: none;'></th>";
                thead.insertBefore(spacer, tRow.nextSibling);
            }
                        tRow.innerHTML = `
                <th style="width: 22%; padding: 15px 10px; text-align: left; box-sizing: border-box;">
                    <div style="display: flex; align-items: center; gap: 8px; color: #94a3b8; font-weight: 700;  letter-spacing: 1px;">
                        <i data-lucide="pie-chart" style="width: 16px; height: 16px; color: #38bdf8;"></i>
                        <span style="font-size: 0.85rem;">RESUMEN TOTAL</span>
                    </div>
                </th>
                <th style="width: 22%; padding: 15px 10px; text-align: right; box-sizing: border-box;">
                    <div style="display: flex; flex-direction: column; align-items: flex-end; overflow: hidden;">
                        <span style="color: #94a3b8; font-size: 0.7rem;  font-weight: 600; margin-bottom: 4px;">Ventas</span>
                        <span style="color: #f8fafc; font-size: 0.95rem; font-weight: 700; white-space: nowrap;">${formatPrice(totalSales)}</span>
                    </div>
                </th>
                <th style="width: 22%; padding: 15px 10px; text-align: right; box-sizing: border-box;">
                    <div style="display: flex; flex-direction: column; align-items: flex-end; overflow: hidden;">
                        <span style="color: #94a3b8; font-size: 0.7rem;  font-weight: 600; margin-bottom: 4px;">Gastos</span>
                        <span style="color: #f8fafc; font-size: 0.95rem; font-weight: 700; white-space: nowrap;">${formatPrice(totalExp)}</span>
                    </div>
                </th>
                <th colspan="2" style="width: 34%; padding: 15px 10px; text-align: right; box-sizing: border-box;">
                      <div style="display: flex; flex-direction: column; align-items: flex-end; padding-right: 5px; overflow: hidden;">
                          <span style="color: #94a3b8; font-size: 0.7rem;  font-weight: 600; margin-bottom: 4px;">Balance</span>
                          <span style="color: ${darkColor}; font-size: 1.15rem; font-weight: 900; text-shadow: 0 2px 4px rgba(0,0,0,0.3); white-space: nowrap;">${formatPrice(totalBalance)}</span>
                      </div>
                </th>
            `;
        }
        tbody.innerHTML = rowsHtml;
    }

        // Event listeners for Balance
    const balSelect = document.getElementById('balancePeriodSelect');
    const balDatePickerGroup = document.getElementById('balanceDatePickerGroup');
    const balMonthPickerGroup = document.getElementById('balanceMonthPickerGroup');
    const searchBalanceBtn = document.getElementById('searchBalanceBtn');

    const balanceViewSelect = document.getElementById('balanceViewSelect');
    const chartsContainer = document.getElementById('balanceChartsContainer');
    const tableContainer = document.getElementById('balanceTableContainer');
    
    if (balanceViewSelect && chartsContainer && tableContainer) {
        balanceViewSelect.addEventListener('change', (e) => {
            const sumBlock = document.getElementById('balanceSummaryBlock');
            if (e.target.value === 'charts') {
                chartsContainer.style.display = 'flex';
                tableContainer.style.display = 'none';
                if (sumBlock) sumBlock.style.display = 'block';
                if (window.chartsMap) Object.values(window.chartsMap).forEach(c => c.resize());
            } else {
                chartsContainer.style.display = 'none';
                tableContainer.style.display = 'block';
                if (sumBlock) sumBlock.style.display = 'none';
            }
        });
    }
    

    

    
    if (balSelect) {
        balSelect.addEventListener('change', (e) => {
            const val = e.target.value;
            
            if (balDatePickerGroup) balDatePickerGroup.classList.add('hidden');
            if (balMonthPickerGroup) balMonthPickerGroup.classList.add('hidden');
            if (searchBalanceBtn) searchBalanceBtn.classList.add('hidden');
            
            if (val === 'date') {
                if (balDatePickerGroup) balDatePickerGroup.classList.remove('hidden');
                if (searchBalanceBtn) searchBalanceBtn.classList.remove('hidden');
            } else if (val === 'specific-month') {
                if (balMonthPickerGroup) balMonthPickerGroup.classList.remove('hidden');
                if (searchBalanceBtn) searchBalanceBtn.classList.remove('hidden');
            } else {
                renderBalancePage();
            }
        });
        
        // Setup initial visibility
        if (balSelect.value !== 'date' && balSelect.value !== 'specific-month') {
            if (searchBalanceBtn) searchBalanceBtn.classList.add('hidden');
        }
    }
    
    if (searchBalanceBtn) {
        searchBalanceBtn.addEventListener('click', renderBalancePage);
    }


    window.openBalanceDetails = function(dateStr) {
        const parts = dateStr.split('/');
        let displayStr = dateStr;
        if (parts.length === 3) {
            const dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            const shortYear = parts[2].slice(-2);
            const daysOfWeek = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
            const dayName = daysOfWeek[dateObj.getDay()];
            displayStr = `${parts[0]}/${parts[1]}/${shortYear} (${dayName})`;
        }
        document.getElementById('balanceDetailTitle').textContent = 'Detalles - ' + displayStr;
        const sList = document.getElementById('balanceDetailSalesList');
        const eList = document.getElementById('balanceDetailExpensesList');
        
        let sHtml = '';
        let eHtml = '';
        let tSales = 0;
        let tExp = 0;

        const allOrders = StorageManager.getOrders().filter(o => o.paid && !o.isPartial);
        const dayOrders = allOrders.filter(o => {
            const d = new Date(o.createdAt);
            return d.toLocaleDateString('es-CO', { year: 'numeric', month: '2-digit', day: '2-digit' }) === dateStr;
        });

        if(dayOrders.length === 0) {
            sHtml = '<div style="color: var(--text-muted); text-align: center; padding: 10px;">Sin ventas</div>';
        } else {
            dayOrders.forEach(o => {
                tSales += o.totalPrice;
            });
            sHtml = ''; // The total is already shown in the header, no need for extra box
        }

        const allExpenses = StorageManager.getExpenses();
        const dayExpenses = allExpenses.filter(e => {
            const d = new Date(e.date || e.createdAt);
            return d.toLocaleDateString('es-CO', { year: 'numeric', month: '2-digit', day: '2-digit' }) === dateStr;
        });

        if(dayExpenses.length === 0) {
            eHtml = '<div style="color: var(--text-muted); text-align: center; padding: 10px;">Sin gastos</div>';
        } else {
            dayExpenses.forEach(e => {
                tExp += e.amount;
                eHtml += `
                    <div style="display: flex; justify-content: center; text-align: center; border-bottom: 1px solid rgba(0,0,0,0.05); padding-bottom: 4px;">
                        <span style="display: flex; flex-direction: column;">
                            <span>${e.description}</span>
                            <span style="color: var(--text-muted); font-size: 0.7rem;">${e.category}</span>
                        </span>
                        <span style="font-weight: 600; color: #dc2626;">${formatPrice(e.amount)}</span>
                    </div>
                `;
            });
        }

        sList.innerHTML = sHtml;
        eList.innerHTML = eHtml;
        document.getElementById('balanceDetailSalesTotal').textContent = formatPrice(tSales);
        document.getElementById('balanceDetailExpTotal').textContent = formatPrice(tExp);

        document.getElementById('balanceDetailModal').classList.add('open');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    };
