// ============================================
// FoodX POS - Storage Manager
// ============================================

// Unique prefix for this specific application to prevent crossover
const PREFIX = 'minesof_';

// Multi-tenant Firestore helper
function getDbCollection(key) {
    const tenantId = window.currentUserTenant || 'default_tenant';
    return db.collection('tenants').doc(tenantId).collection(key);
}
const STORAGE_KEYS = {
    ORDERS: PREFIX + 'orders',
    SETTINGS: PREFIX + 'settings',
    CATEGORIES: PREFIX + 'categories_v3',
    PRODUCTS: PREFIX + 'products_v3',
    FLAVORS: PREFIX + 'flavors_v3',
    EXTRAS: PREFIX + 'extras_v3',
    PRICES: PREFIX + 'prices_v3',
    EXPENSES: PREFIX + 'expenses',
    EXPENSE_CATEGORIES: PREFIX + 'expense_categories'
};

const StorageManager = {
    // Get all orders
    getOrders() {
        const data = localStorage.getItem(STORAGE_KEYS.ORDERS);
        return data ? JSON.parse(data) : [];
    },

    // Save all orders
    saveOrders(orders) {
        localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(orders));
    },

    async deleteOrder(orderId) {
        const orders = this.getOrders().filter(o => o.id !== orderId);
        this.saveOrders(orders);
        this.deleteOrderFromCloud(orderId);
    },

    // Get orders by status
    // --- Sales Calculation Methods ---
    getTodaySales() {
        return this.getTodayOrders().filter(o => o.paid && !o.isPartial).reduce((sum, o) => sum + (o.totalPrice || 0), 0);
    },

    getCurrentMonthSales() {
        return this.getCurrentMonthOrders().filter(o => o.paid && !o.isPartial).reduce((sum, o) => sum + (o.totalPrice || 0), 0);
    },

    getTotalSales() {
        return this.getOrders().filter(o => o.paid && !o.isPartial).reduce((sum, o) => sum + (o.totalPrice || 0), 0);
    },

    // Get active orders (not delivered)
    getActiveOrders() {
        return this.getOrders().filter(o => o.status !== 'delivered');
    },

    // Get order counts by status
    getOrderCounts() {
        const orders = this.getOrders();
        return {
            pending: orders.filter(o => o.status === 'pending').length,
            preparing: orders.filter(o => o.status === 'preparing').length,
            ready: orders.filter(o => o.status === 'ready').length
        };
    },

    // Configuration Management
    getConfig() {
        let categories = localStorage.getItem(STORAGE_KEYS.CATEGORIES);
        let products = localStorage.getItem(STORAGE_KEYS.PRODUCTS);
        let flavors = localStorage.getItem(STORAGE_KEYS.FLAVORS);
        let extras = localStorage.getItem(STORAGE_KEYS.EXTRAS);
        let observations = localStorage.getItem('galeria_observations');
        let prices = localStorage.getItem(STORAGE_KEYS.PRICES);

        // Auto-upgrade to multi-sector catalog if storage has old categories or no products
        if (!products && categories && categories.includes('hamburguesas')) {
            categories = null;
        }

        const parsedCategories = categories ? JSON.parse(categories) : FOODX_DATA.categories;
        const parsedProducts = products ? JSON.parse(products) : (FOODX_DATA.products || []);

        const config = {
            businessName: localStorage.getItem(PREFIX + 'businessName') || FOODX_DATA.businessName,
            businessLogo: localStorage.getItem(PREFIX + 'businessLogo') || FOODX_DATA.businessLogo,
            billingSystem: localStorage.getItem(PREFIX + 'billingSystem') || 'standard',
            categories: parsedCategories,
            products: parsedProducts,
            flavors: flavors ? JSON.parse(flavors) : FOODX_DATA.flavors,
            extras: extras ? JSON.parse(extras) : FOODX_DATA.extras,
            observations: observations ? JSON.parse(observations) : FOODX_DATA.observations,
            prices: prices ? JSON.parse(prices) : (FOODX_DATA.prices || {}),
            adminPassword: localStorage.getItem('galeria_admin_password') || '1234'
        };

        // Migration: If extras or observations are arrays, convert to objects keyed by category
        if (Array.isArray(config.extras)) {
            const migrated = {};
            config.categories.forEach(c => {
                migrated[c.id] = JSON.parse(JSON.stringify(config.extras));
            });
            config.extras = migrated;
        }

        if (Array.isArray(config.observations)) {
            const migrated = {};
            config.categories.forEach(c => {
                migrated[c.id] = JSON.parse(JSON.stringify(config.observations));
            });
            config.observations = migrated;
        }

        return config;
    },

    getProducts() {
        return this.getConfig().products || [];
    },

    saveProducts(products) {
        localStorage.setItem(STORAGE_KEYS.PRODUCTS, JSON.stringify(products));
        const cfg = this.getConfig();
        cfg.products = products;
        this.saveConfig(cfg);
    },

    saveConfig(config) {
        // Validation: Ensure extras and observations are objects, not arrays
        if (Array.isArray(config.extras)) {
            const migrated = {};
            config.categories.forEach(c => migrated[c.id] = [...config.extras]);
            config.extras = migrated;
        }
        if (Array.isArray(config.observations)) {
            const migrated = {};
            config.categories.forEach(c => migrated[c.id] = [...config.observations]);
            config.observations = migrated;
        }

        if (config.categories) localStorage.setItem(STORAGE_KEYS.CATEGORIES, JSON.stringify(config.categories));
        if (config.products) localStorage.setItem(STORAGE_KEYS.PRODUCTS, JSON.stringify(config.products));
        if (config.flavors) localStorage.setItem(STORAGE_KEYS.FLAVORS, JSON.stringify(config.flavors));
        if (config.extras) localStorage.setItem(STORAGE_KEYS.EXTRAS, JSON.stringify(config.extras));
        if (config.observations) localStorage.setItem('galeria_observations', JSON.stringify(config.observations));
        if (config.businessName !== undefined) localStorage.setItem(PREFIX + 'businessName', config.businessName);
        if (config.businessLogo !== undefined) localStorage.setItem(PREFIX + 'businessLogo', config.businessLogo);
        if (config.billingSystem !== undefined) localStorage.setItem(PREFIX + 'billingSystem', config.billingSystem);
        if (config.prices) localStorage.setItem(STORAGE_KEYS.PRICES, JSON.stringify(config.prices));
        if (config.adminPassword) localStorage.setItem('galeria_admin_password', config.adminPassword);

        // Update the global object too so the app uses latest
        Object.assign(FOODX_DATA, config);

        // Sync to cloud
        this.syncConfigToCloud(config);
    },

    // Today's orders
    getTodayOrders() {
        const today = new Date().toDateString();
        return this.getOrders().filter(o => new Date(o.createdAt).toDateString() === today);
    },

    // Current month orders
    getCurrentMonthOrders() {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();
        return this.getOrders().filter(o => {
            const date = new Date(o.createdAt);
            return date.getMonth() === currentMonth && date.getFullYear() === currentYear;
        });
    },

    // Get orders by specific date (YYYY-MM-DD)
    getOrdersByDate(dateStr) {
        if (!dateStr) return [];
        const searchDate = new Date(dateStr + 'T00:00:00').toDateString();
        return this.getOrders().filter(o => new Date(o.createdAt).toDateString() === searchDate);
    },

    // Get orders by specific month (YYYY-MM)
    getOrdersByMonth(monthStr) {
        if (!monthStr) return [];
        const [year, month] = monthStr.split('-').map(Number);
        return this.getOrders().filter(o => {
            const date = new Date(o.createdAt);
            return date.getFullYear() === year && (date.getMonth() + 1) === month;
        });
    },

    getSalesByMonth(monthStr) {
        return this.getOrdersByMonth(monthStr).filter(o => o.paid && !o.isPartial).reduce((sum, o) => sum + (o.totalPrice || 0), 0);
    },

        // --- Firebase Sync Methods ---

    async syncOrderToCloud(order) {
        if (typeof db === 'undefined') return;
        try {
            const cleanOrder = JSON.parse(JSON.stringify(order));
            await getDbCollection(STORAGE_KEYS.ORDERS).doc(order.id).set(cleanOrder, { merge: true });
        } catch (e) {
            console.error('Error syncing order:', e);
        }
    },

    async deleteOrderFromCloud(orderId) {
        if (typeof db === 'undefined') return;
        try {
            await getDbCollection(STORAGE_KEYS.ORDERS).doc(orderId).delete();
        } catch (e) {
            console.error('Error deleting order:', e);
        }
    },

    async syncConfigToCloud(config) {
        if (typeof db === 'undefined') return;
        try {
            const cleanConfig = JSON.parse(JSON.stringify(config));
            await getDbCollection(STORAGE_KEYS.SETTINGS).doc('global_config').set(cleanConfig);
        } catch (e) {
            console.error('Error syncing config:', e);
        }
    },

    async syncExpenseToCloud(expense) {
        if (typeof db === 'undefined') return;
        try {
            const cleanExpense = JSON.parse(JSON.stringify(expense));
            await getDbCollection(STORAGE_KEYS.EXPENSES).doc(expense.id).set(cleanExpense, { merge: true });
        } catch (e) {
            console.error('Error syncing expense:', e);
        }
    },
    
    async deleteExpenseFromCloud(expenseId) {
        if (typeof db === 'undefined') return;
        try {
            await getDbCollection(STORAGE_KEYS.EXPENSES).doc(expenseId).delete();
        } catch (e) {
            console.error('Error deleting expense:', e);
        }
    },

    unsubConfig: null, unsubOrders: null, unsubExpenses: null,

    initCloudSync(callback, configCallback, printCallback) {
        if (typeof db === 'undefined') {
            console.warn('Firebase db no detectado. Modo 100% offline.');
            return;
        }

        if (this.unsubConfig) this.unsubConfig();
        if (this.unsubOrders) this.unsubOrders();
        if (this.unsubExpenses) this.unsubExpenses();

        // 1. Escuchar Configuración Global
        this.unsubConfig = getDbCollection(STORAGE_KEYS.SETTINGS).doc('global_config').onSnapshot(doc => {
            if (doc.exists) {
                const data = doc.data();
                if (data.lastWipe) {
                    const localWipe = localStorage.getItem('galeria_last_wipe');
                    if (!localWipe || data.lastWipe > parseInt(localWipe)) {
                        localStorage.removeItem(STORAGE_KEYS.ORDERS);
                        localStorage.removeItem(STORAGE_KEYS.EXPENSES);
                        localStorage.setItem('galeria_last_wipe', data.lastWipe.toString());
                        setTimeout(() => window.location.reload(), 500);
                        return;
                    }
                }
                if (data.categories) localStorage.setItem(STORAGE_KEYS.CATEGORIES, JSON.stringify(data.categories));
                if (data.products) localStorage.setItem(STORAGE_KEYS.PRODUCTS, JSON.stringify(data.products));
                if (data.flavors) localStorage.setItem(STORAGE_KEYS.FLAVORS, JSON.stringify(data.flavors));
                if (data.extras) localStorage.setItem(STORAGE_KEYS.EXTRAS, JSON.stringify(data.extras));
                if (data.observations) localStorage.setItem('galeria_observations', JSON.stringify(data.observations));
                if (data.businessName !== undefined) localStorage.setItem(PREFIX + 'businessName', data.businessName);
                if (data.businessLogo !== undefined) localStorage.setItem(PREFIX + 'businessLogo', data.businessLogo);
                if (data.billingSystem !== undefined) localStorage.setItem(PREFIX + 'billingSystem', data.billingSystem);
                if (data.prices) localStorage.setItem(STORAGE_KEYS.PRICES, JSON.stringify(data.prices));
                if (data.adminPassword) localStorage.setItem('galeria_admin_password', data.adminPassword);

                Object.assign(FOODX_DATA, data);
                if (typeof configCallback === 'function') configCallback();
            }
        });

        // 2. Escuchar Pedidos (Para no descargar todo el historial, escuchamos los recientes)
        this.unsubOrders = getDbCollection(STORAGE_KEYS.ORDERS)
            .orderBy('createdAt', 'desc')
            .onSnapshot(snapshot => {
                let localOrders = this.getOrders();
                let changed = false;
                
                snapshot.docChanges().forEach(change => {
                    const order = change.doc.data();
                    if (change.type === 'added' || change.type === 'modified') {
                        const idx = localOrders.findIndex(o => o.id === order.id);
                        if (idx !== -1) {
                            localOrders[idx] = order;
                            changed = true;
                        } else {
                            localOrders.push(order);
                            changed = true;
                        }
                    } else if (change.type === 'removed') {
                        localOrders = localOrders.filter(o => o.id !== order.id);
                        changed = true;
                    }
                });

                if (changed) {
                    this.saveOrders(localOrders);
                    if (typeof callback === 'function') callback();
                }
            });

        // 3. Escuchar Egresos
        this.unsubExpenses = getDbCollection(STORAGE_KEYS.EXPENSES)
            .orderBy('createdAt', 'desc')
            .onSnapshot(snapshot => {
                let local = this.getExpenses();
                let changed = false;
                snapshot.docChanges().forEach(change => {
                    const item = change.doc.data();
                    if (change.type === 'added' || change.type === 'modified') {
                        const idx = local.findIndex(o => o.id === item.id);
                        if (idx !== -1) {
                            local[idx] = item;
                        } else {
                            local.push(item);
                        }
                        changed = true;
                    } else if (change.type === 'removed') {
                        local = local.filter(o => o.id !== item.id);
                        changed = true;
                    }
                });
                if (changed) {
                    this.saveExpenses(local);
                    if (typeof callback === 'function') callback();
                }
            });
    },

    // --- Original methods with cloud hooks ---

    async addOrder(order) {
        order.id = generateId();
        order.createdAt = new Date().toISOString();
        const orders = this.getOrders();
        orders.push(order);
        this.saveOrders(orders);
        
        try {
            await this.syncOrderToCloud(order);
        } catch (e) {
            console.error(e);
        }
        
        return order;
    },

    updateOrder(orderId, updates) {
        const orders = this.getOrders();
        const index = orders.findIndex(o => o.id === orderId);
        if (index !== -1) {
            orders[index] = { ...orders[index], ...updates, modifiedAt: new Date().toISOString() };
            this.saveOrders(orders);
            this.syncOrderToCloud(orders[index]); // Hook
            return orders[index];
        }
        return null;
    },

    async deleteOrder(orderId) {
        const orders = this.getOrders().filter(o => o.id !== orderId);
        this.saveOrders(orders);
        this.deleteOrderFromCloud(orderId);
    },

    // ============================================
    // Expense Categories (Categorías de Egresos)
    // ============================================

    getExpenseCategories() {
        const data = localStorage.getItem(STORAGE_KEYS.EXPENSE_CATEGORIES);
        if (data) return JSON.parse(data);
                return [
            { id: 'materia_prima', label: 'Materia Prima', emoji: '📦' },
            { id: 'servicios', label: 'Servicios', emoji: '💡' },
            { id: 'arrendamiento', label: 'Arrendamiento', emoji: '🏠' },
            { id: 'nomina', label: 'Nómina', emoji: '👥' }
        ];
    },

    saveExpenseCategories(categories) {
        localStorage.setItem(STORAGE_KEYS.EXPENSE_CATEGORIES, JSON.stringify(categories));
    },

    // ============================================
    // Expenses (Egresos)
    // ============================================

    getExpenses() {
        const data = localStorage.getItem(STORAGE_KEYS.EXPENSES);
        return data ? JSON.parse(data) : [];
    },

    saveExpenses(expenses) {
        localStorage.setItem(STORAGE_KEYS.EXPENSES, JSON.stringify(expenses));
    },

    addExpense(expense) {
        expense.id = generateId();
        expense.createdAt = new Date().toISOString();
        const expenses = this.getExpenses();
        expenses.push(expense);
        this.saveExpenses(expenses);
        this.syncExpenseToCloud(expense);
        return expense;
    },

    async deleteExpense(expenseId) {
        const expenses = this.getExpenses().filter(e => e.id !== expenseId);
        this.saveExpenses(expenses);
        this.deleteExpenseFromCloud(expenseId);
    },

    

    getTodayExpenses() {
        const today = new Date().toDateString();
        return this.getExpenses().filter(e => new Date(e.date || e.createdAt).toDateString() === today);
    },

    getCurrentMonthExpenses() {
        const now = new Date();
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();
        return this.getExpenses().filter(e => {
            const date = new Date(e.date || e.createdAt);
            return date.getMonth() === currentMonth && date.getFullYear() === currentYear;
        });
    },

    // Get expenses by specific month (YYYY-MM)
    getExpensesByMonth(monthStr) {
        if (!monthStr) return [];
        const [year, month] = monthStr.split('-').map(Number);
        return this.getExpenses().filter(e => {
            const date = new Date(e.date || e.createdAt);
            return date.getFullYear() === year && (date.getMonth() + 1) === month;
        });
    },

    // Clear all data (for testing)
    clearAll() {
        Object.values(STORAGE_KEYS).forEach(key => localStorage.removeItem(key));
        localStorage.removeItem(PREFIX + 'businessName');
        localStorage.removeItem(PREFIX + 'businessLogo');
        localStorage.removeItem(PREFIX + 'billingSystem');
        localStorage.removeItem('galeria_admin_password');
        localStorage.removeItem('galeria_observations');
    }
};

// Flag to track if config has been loaded from cloud
StorageManager.configLoaded = false;

// Initialize FOODX_DATA from storage
(function initConfig() {
    const config = StorageManager.getConfig();
    Object.assign(FOODX_DATA, config);
    StorageManager.configLoaded = true;
    window.dispatchEvent(new CustomEvent('configLoadedFromCloud')); // Kept name for compatibility
})();












