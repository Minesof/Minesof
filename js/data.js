// ============================================
// FoodX POS - Data Configuration
// ============================================

const FOODX_DATA = {
    businessName: 'Minesof',
    businessSubtitle: 'SISTEMA POS PRO',

    // Multi-sector Categories
    categories: [],

    // Products by unit
    products: [],

    // Backward-compatible flavors map grouped by category
    flavors: {
        panaderia: [],
        reposteria: [],
        cafeteria: [],
        bebidas: [],
        desayunos: []
    },

    // Extras/Additions by category
    extras: {
        panaderia: [
            { id: 'e1', name: 'Queso Extra', price: 1500, active: true },
            { id: 'e2', name: 'Arequipe Extra', price: 1500, active: true },
            { id: 'e3', name: 'Mantequilla', price: 1000, active: true }
        ],
        cafeteria: [
            { id: 'e4', name: 'Leche Deslactosada', price: 1000, active: true },
            { id: 'e5', name: 'Leche Almendras', price: 2000, active: true },
            { id: 'e6', name: 'Shot Espresso Extra', price: 2000, active: true },
            { id: 'e7', name: 'Sirope Vainilla', price: 1500, active: true }
        ],
        desayunos: [
            { id: 'e8', name: 'PorciÃ³n Tocineta', price: 3000, active: true },
            { id: 'e9', name: 'Queso Extra', price: 2000, active: true },
            { id: 'e10', name: 'Huevo Adicional', price: 2000, active: true }
        ],
        reposteria: [],
        bebidas: []
    },

    // Observations/Special requests by category
    observations: {
        panaderia: [
            { id: 'o1', name: 'Bien caliente', active: true },
            { id: 'o2', name: 'Tostado', active: true },
            { id: 'o3', name: 'Blando', active: true }
        ],
        cafeteria: [
            { id: 'o4', name: 'Sin azÃºcar', active: true },
            { id: 'o5', name: 'Poco dulce', active: true },
            { id: 'o6', name: 'Bien caliente', active: true },
            { id: 'o7', name: 'Tibio', active: true }
        ],
        bebidas: [
            { id: 'o8', name: 'Sin hielo', active: true },
            { id: 'o9', name: 'Con hielo', active: true },
            { id: 'o10', name: 'Sin azÃºcar', active: true }
        ],
        desayunos: [
            { id: 'o11', name: 'Huevos revueltos', active: true },
            { id: 'o12', name: 'Huevos fritos', active: true },
            { id: 'o13', name: 'Huevos pericos', active: true },
            { id: 'o14', name: 'Sin cebolla', active: true }
        ],
        reposteria: [
            { id: 'o15', name: 'Para llevar empacado', active: true }
        ]
    },

    // Base price placeholder for backward compatibility
    prices: {},

    adminPassword: '1234',

    // Service types
    serviceTypes: [
        { id: 'salon', name: 'SalÃ³n', label: 'Mesa' },
        { id: 'llevar', name: 'Para Llevar', label: 'Nombre' },
        { id: 'domicilio', name: 'Domicilio', label: 'DirecciÃ³n/Nombre' }
    ]
};

// Size calculation fallback (returns empty string since products have unit prices)
function calculateSize(blocksCount = 0, category = '', selectedFlavors = []) {
    return '';
}

// Format price to Colombian pesos
function formatPrice(price) {
    if (price === null || price === undefined || isNaN(price)) return '$0';
    return '$' + Number(price).toLocaleString('es-CO');
}
// Generate unique ID
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Generate order number with daily reset - SYNCHRONIZED via Firebase
let orderCounter = parseInt(localStorage.getItem('galeria_order_counter') || '0');
let lastOrderDate = localStorage.getItem('galeria_last_order_date') || '';

// Get local date key (YYYY-MM-DD in local timezone)
function getLocalDateKey() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

async function getNextOrderNumber() {
    return generateOrderNumberLocal();
}

// Local fallback function (original logic)
function generateOrderNumberLocal() {
    const today = new Date().toDateString();
    const lastDate = localStorage.getItem('galeria_last_order_date');

    if (lastDate !== today) {
        orderCounter = 0;
        localStorage.setItem('galeria_last_order_date', today);
    }

    orderCounter++;
    localStorage.setItem('galeria_order_counter', orderCounter.toString());
    return '#' + String(orderCounter).padStart(3, '0');
}

// Sync wrapper - returns a promise
function generateOrderNumber() {
    // Return the async result, but for backward compatibility also have sync fallback
    return getNextOrderNumber();
}


 + Number(price).toLocaleString('es-CO');
}

// Generate unique ID
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Generate order number with daily reset - SYNCHRONIZED via Firebase
let orderCounter = parseInt(localStorage.getItem('galeria_order_counter') || '0');
let lastOrderDate = localStorage.getItem('galeria_last_order_date') || '';

// Get local date key (YYYY-MM-DD in local timezone)
function getLocalDateKey() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

async function getNextOrderNumber() {
    return generateOrderNumberLocal();
}

// Local fallback function (original logic)
function generateOrderNumberLocal() {
    const today = new Date().toDateString();
    const lastDate = localStorage.getItem('galeria_last_order_date');

    if (lastDate !== today) {
        orderCounter = 0;
        localStorage.setItem('galeria_last_order_date', today);
    }

    orderCounter++;
    localStorage.setItem('galeria_order_counter', orderCounter.toString());
    return '#' + String(orderCounter).padStart(3, '0');
}

// Sync wrapper - returns a promise
function generateOrderNumber() {
    // Return the async result, but for backward compatibility also have sync fallback
    return getNextOrderNumber();
}


