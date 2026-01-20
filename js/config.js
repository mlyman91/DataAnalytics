/**
 * ============================================
 * PVM Bridge Tool - Configuration
 * ============================================
 * 
 * Central configuration for column mappings,
 * validation rules, and processing settings.
 * 
 * All thresholds and defaults are auditable
 * and documented here for transparency.
 * ============================================
 */

const CONFIG = {
    // ----------------------------------------
    // Version & Metadata
    // ----------------------------------------
    VERSION: '1.0.0',
    APP_NAME: 'PVM & Gross Margin Bridge Tool',
    
    // ----------------------------------------
    // Processing Settings
    // ----------------------------------------
    CHUNK_SIZE: 1024 * 1024, // 1MB chunks for CSV streaming
    PROGRESS_UPDATE_INTERVAL: 100, // Update UI every N rows
    MAX_PREVIEW_ROWS: 5, // Rows to show for date format preview
    
    // ----------------------------------------
    // Column Detection Patterns
    // ----------------------------------------
    // These patterns help auto-detect column mappings
    // Order matters: first match wins
    COLUMN_PATTERNS: {
        date: [
            /^date$/i,
            /^transaction.?date$/i,
            /^invoice.?date$/i,
            /^order.?date$/i,
            /^sale.?date$/i,
            /date/i
        ],
        sales: [
            /^sales$/i,
            /^revenue$/i,
            /^net.?sales$/i,
            /^total.?sales$/i,
            /^amount$/i,
            /^sales.?amount$/i,
            /sales/i,
            /revenue/i
        ],
        quantity: [
            /^quantity$/i,
            /^qty$/i,
            /^volume$/i,
            /^units$/i,
            /^count$/i,
            /quantity/i,
            /volume/i
        ],
        cost: [
            /^cost$/i,
            /^cogs$/i,
            /^cost.?of.?goods$/i,
            /^total.?cost$/i,
            /^unit.?cost$/i,
            /cost/i
        ]
    },
    
    // ----------------------------------------
    // Date Format Detection
    // ----------------------------------------
    // Patterns for auto-detecting date formats
    DATE_FORMATS: [
        {
            id: 'YYYY-MM-DD',
            pattern: /^\d{4}-\d{2}-\d{2}$/,
            parse: (s) => {
                const [y, m, d] = s.split('-').map(Number);
                return new Date(y, m - 1, d);
            },
            example: '2024-01-15'
        },
        {
            id: 'MM/DD/YYYY',
            pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/,
            parse: (s) => {
                const [m, d, y] = s.split('/').map(Number);
                return new Date(y, m - 1, d);
            },
            example: '01/15/2024'
        },
        {
            id: 'DD/MM/YYYY',
            pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/,
            parse: (s) => {
                const [d, m, y] = s.split('/').map(Number);
                return new Date(y, m - 1, d);
            },
            example: '15/01/2024'
        },
        {
            id: 'MM-DD-YYYY',
            pattern: /^\d{1,2}-\d{1,2}-\d{4}$/,
            parse: (s) => {
                const [m, d, y] = s.split('-').map(Number);
                return new Date(y, m - 1, d);
            },
            example: '01-15-2024'
        },
        {
            id: 'DD-MMM-YYYY',
            pattern: /^\d{1,2}-[A-Za-z]{3}-\d{4}$/,
            parse: (s) => {
                const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
                const parts = s.split('-');
                const d = parseInt(parts[0], 10);
                const m = months[parts[1].toLowerCase()];
                const y = parseInt(parts[2], 10);
                return new Date(y, m, d);
            },
            example: '15-Jan-2024'
        },
        {
            id: 'MMM DD, YYYY',
            pattern: /^[A-Za-z]{3}\s+\d{1,2},?\s+\d{4}$/,
            parse: (s) => {
                const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
                const match = s.match(/([A-Za-z]{3})\s+(\d{1,2}),?\s+(\d{4})/);
                if (match) {
                    const m = months[match[1].toLowerCase()];
                    const d = parseInt(match[2], 10);
                    const y = parseInt(match[3], 10);
                    return new Date(y, m, d);
                }
                return null;
            },
            example: 'Jan 15, 2024'
        },
        {
            id: 'YYYYMMDD',
            pattern: /^\d{8}$/,
            parse: (s) => {
                const y = parseInt(s.substring(0, 4), 10);
                const m = parseInt(s.substring(4, 6), 10);
                const d = parseInt(s.substring(6, 8), 10);
                return new Date(y, m - 1, d);
            },
            example: '20240115'
        },
        {
            id: 'M/D/YYYY',
            pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/,
            parse: (s) => {
                const [m, d, y] = s.split('/').map(Number);
                return new Date(y, m - 1, d);
            },
            example: '1/5/2024'
        }
    ],
    
    // ----------------------------------------
    // Month Names (for display)
    // ----------------------------------------
    MONTH_NAMES: [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ],
    
    MONTH_NAMES_SHORT: [
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
    ],
    
    // ----------------------------------------
    // Validation Rules
    // ----------------------------------------
    VALIDATION: {
        // Rows with sales <= 0 OR quantity <= 0 are excluded from PVM math
        // These are tracked separately as "negative values"
        MIN_SALES_FOR_PVM: 0,
        MIN_QUANTITY_FOR_PVM: 0,
        
        // Maximum dimensions allowed (soft limit, warning only)
        MAX_RECOMMENDED_DIMENSIONS: 4,
        
        // Warn if estimated combinations exceed this
        MAX_RECOMMENDED_COMBINATIONS: 100000
    },
    
    // ----------------------------------------
    // Default Values
    // ----------------------------------------
    DEFAULTS: {
        FISCAL_YEAR_END_MONTH: 12, // December
        UNKNOWN_DIMENSION_VALUE: 'Unknown',
        
        // GM mode price definition
        GM_PRICE_DEFINITION: 'margin-per-unit' // or 'sales-per-unit'
    },
    
    // ----------------------------------------
    // Number Formatting
    // ----------------------------------------
    FORMAT: {
        // Decimal places for different values
        DECIMALS_CURRENCY: 2,
        DECIMALS_QUANTITY: 0,
        DECIMALS_PRICE: 4,
        DECIMALS_PERCENT: 1,
        
        // Thousand separator
        THOUSAND_SEP: ',',
        DECIMAL_SEP: '.'
    },
    
    // ----------------------------------------
    // Local Storage Keys
    // ----------------------------------------
    STORAGE_KEYS: {
        LAST_CONFIG: 'pvm_bridge_last_config',
        PREFERENCES: 'pvm_bridge_preferences'
    },
    
    // ----------------------------------------
    // Excel Export Settings
    // ----------------------------------------
    EXCEL: {
        // Tab names for export
        TAB_NAMES: {
            SUMMARY: 'Summary Bridge',
            DETAIL: 'Detail by LOD',
            NEGATIVES: 'Negative Values',
            ASSUMPTIONS: 'Assumptions'
        },
        
        // Column widths (in characters)
        COL_WIDTH_DEFAULT: 15,
        COL_WIDTH_DIMENSION: 20,
        COL_WIDTH_NUMBER: 18
    }
};

// Freeze config to prevent accidental modification
Object.freeze(CONFIG);
Object.freeze(CONFIG.COLUMN_PATTERNS);
Object.freeze(CONFIG.VALIDATION);
Object.freeze(CONFIG.DEFAULTS);
Object.freeze(CONFIG.FORMAT);
Object.freeze(CONFIG.STORAGE_KEYS);
Object.freeze(CONFIG.EXCEL);
Object.freeze(CONFIG.EXCEL.TAB_NAMES);

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CONFIG;
}
