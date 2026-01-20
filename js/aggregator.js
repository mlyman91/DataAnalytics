/**
 * ============================================
 * PVM Bridge Tool - Aggregator
 * ============================================
 * 
 * Incremental aggregation engine.
 * Accumulates data by LOD combination without
 * storing individual rows.
 * 
 * Memory-efficient: only stores unique LOD keys
 * plus their aggregated values.
 * ============================================
 */

const Aggregator = {
    /**
     * Create a new aggregation context
     *
     * @param {Object} config - Aggregation configuration
     * @param {string[]} config.dimensions - Dimension columns for LOD
     * @param {string} config.dateColumn - Date column name
     * @param {string} config.salesColumn - Sales column name
     * @param {string} config.quantityColumn - Quantity column name
     * @param {string} config.costColumn - Cost column name (optional, for GM mode)
     * @param {string} config.dateFormat - Date format ID
     * @param {Object} config.pyRange - Prior year date range { start, end } (two-period mode)
     * @param {Object} config.cyRange - Current year/LTM date range { start, end } (two-period mode)
     * @param {Object[]} config.fiscalYears - Array of fiscal year configs (multi-year mode)
     * @param {number} config.fyEndMonth - Fiscal year end month (multi-year mode)
     * @returns {Object} - Aggregation context
     */
    createContext: function(config) {
        // Detect mode: multi-year or two-period
        const isMultiYear = config.fiscalYears && config.fiscalYears.length > 0;

        // Build period lookup for multi-year mode
        let yearLookup = null;
        if (isMultiYear) {
            yearLookup = config.fiscalYears.map(fy => ({
                fy: fy.fiscalYear,
                start: fy.start,
                end: fy.end,
                label: fy.label
            }));
        }

        return {
            config: config,
            isMultiYear: isMultiYear,
            yearLookup: yearLookup,

            // Main aggregation map: LOD key -> { years: {2016: {...}, 2017: {...}} } or { py: {...}, cy: {...} }
            data: new Map(),

            // Negative values tracking (by year for multi-year, by py/cy for two-period)
            negatives: isMultiYear ? {} : {
                py: { count: 0, sales: 0, quantity: 0, cost: 0 },
                cy: { count: 0, sales: 0, quantity: 0, cost: 0 }
            },

            // Statistics
            stats: {
                totalRows: 0,
                includedRows: 0,
                excludedRows: 0,
                pyRows: isMultiYear ? undefined : 0,
                cyRows: isMultiYear ? undefined : 0,
                yearRows: isMultiYear ? {} : undefined,
                outsidePeriodRows: 0,
                parseErrors: 0,
                uniqueLODKeys: 0
            },

            // Date range tracking
            dateRange: {
                min: null,
                max: null
            }
        };
    },

    /**
     * Process a single row and update aggregation
     * 
     * @param {Object} ctx - Aggregation context
     * @param {Object} row - Row data object
     * @returns {boolean} - True if row was included in aggregation
     */
    processRow: function(ctx, row) {
        ctx.stats.totalRows++;
        
        const config = ctx.config;
        
        // Parse date
        const dateStr = row[config.dateColumn];
        const date = PeriodUtils.parseDate(dateStr, config.dateFormat);
        
        if (!date) {
            ctx.stats.parseErrors++;
            ctx.stats.excludedRows++;
            return false;
        }
        
        // Track date range
        if (!ctx.dateRange.min || date < ctx.dateRange.min) {
            ctx.dateRange.min = new Date(date);
        }
        if (!ctx.dateRange.max || date > ctx.dateRange.max) {
            ctx.dateRange.max = new Date(date);
        }
        
        // Classify period (multi-year or two-period mode)
        let period = null;
        let fiscalYear = null;

        if (ctx.isMultiYear) {
            // Multi-year mode: find which fiscal year this date belongs to
            for (const yearInfo of ctx.yearLookup) {
                if (date >= yearInfo.start && date <= yearInfo.end) {
                    period = String(yearInfo.fy);
                    fiscalYear = yearInfo.fy;
                    break;
                }
            }
        } else {
            // Two-period mode: classify as PY or CY
            period = PeriodUtils.classifyPeriod(date, config.pyRange, config.cyRange);
        }

        if (!period) {
            ctx.stats.outsidePeriodRows++;
            ctx.stats.excludedRows++;
            return false;
        }

        // Parse numeric values
        const sales = CSVParser.parseNumber(row[config.salesColumn]);
        const quantity = CSVParser.parseNumber(row[config.quantityColumn]);
        const cost = config.costColumn ? CSVParser.parseNumber(row[config.costColumn]) : 0;

        // Check for parse errors
        if (isNaN(sales) || isNaN(quantity)) {
            ctx.stats.parseErrors++;
            ctx.stats.excludedRows++;
            return false;
        }

        // Check for negative/zero values (excluded from standard PVM)
        if (sales <= 0 || quantity <= 0) {
            if (ctx.isMultiYear) {
                if (!ctx.negatives[period]) {
                    ctx.negatives[period] = { count: 0, sales: 0, quantity: 0, cost: 0 };
                }
                const negBucket = ctx.negatives[period];
                negBucket.count++;
                negBucket.sales += isNaN(sales) ? 0 : sales;
                negBucket.quantity += isNaN(quantity) ? 0 : quantity;
                negBucket.cost += isNaN(cost) ? 0 : cost;
            } else {
                const negBucket = period === 'PY' ? ctx.negatives.py : ctx.negatives.cy;
                negBucket.count++;
                negBucket.sales += isNaN(sales) ? 0 : sales;
                negBucket.quantity += isNaN(quantity) ? 0 : quantity;
                negBucket.cost += isNaN(cost) ? 0 : cost;
            }

            ctx.stats.excludedRows++;
            return false;
        }

        // Build LOD key
        const lodKey = this._buildLODKey(row, config.dimensions);

        // Get or create aggregation bucket
        if (!ctx.data.has(lodKey)) {
            const newBucket = {
                dimensions: this._extractDimensions(row, config.dimensions)
            };

            if (ctx.isMultiYear) {
                // Multi-year mode: create years object
                newBucket.years = {};
                for (const yearInfo of ctx.yearLookup) {
                    newBucket.years[yearInfo.fy] = { sales: 0, quantity: 0, cost: 0, count: 0 };
                }
            } else {
                // Two-period mode: create py/cy
                newBucket.py = { sales: 0, quantity: 0, cost: 0, count: 0 };
                newBucket.cy = { sales: 0, quantity: 0, cost: 0, count: 0 };
            }

            ctx.data.set(lodKey, newBucket);
            ctx.stats.uniqueLODKeys++;
        }

        const bucket = ctx.data.get(lodKey);

        let periodBucket;
        if (ctx.isMultiYear) {
            periodBucket = bucket.years[fiscalYear];
        } else {
            periodBucket = period === 'PY' ? bucket.py : bucket.cy;
        }

        // Accumulate
        periodBucket.sales += sales;
        periodBucket.quantity += quantity;
        periodBucket.cost += isNaN(cost) ? 0 : cost;
        periodBucket.count++;

        // Update stats
        ctx.stats.includedRows++;
        if (ctx.isMultiYear) {
            if (!ctx.stats.yearRows[fiscalYear]) {
                ctx.stats.yearRows[fiscalYear] = 0;
            }
            ctx.stats.yearRows[fiscalYear]++;
        } else {
            if (period === 'PY') {
                ctx.stats.pyRows++;
            } else {
                ctx.stats.cyRows++;
            }
        }

        return true;
    },

    /**
     * Build LOD key from row and dimension columns
     * 
     * @private
     */
    _buildLODKey: function(row, dimensions) {
        if (!dimensions || dimensions.length === 0) {
            return '__TOTAL__';
        }
        
        return dimensions.map(dim => {
            const value = row[dim];
            return value && value.trim() ? value.trim() : CONFIG.DEFAULTS.UNKNOWN_DIMENSION_VALUE;
        }).join('|||');
    },

    /**
     * Extract dimension values from row
     * 
     * @private
     */
    _extractDimensions: function(row, dimensions) {
        const result = {};
        for (const dim of dimensions) {
            const value = row[dim];
            result[dim] = value && value.trim() ? value.trim() : CONFIG.DEFAULTS.UNKNOWN_DIMENSION_VALUE;
        }
        return result;
    },

    /**
     * Finalize aggregation and return results
     *
     * @param {Object} ctx - Aggregation context
     * @returns {Object} - Final aggregation results
     */
    finalize: function(ctx) {
        // Convert Map to array for easier processing
        const aggregatedData = [];

        for (const [key, bucket] of ctx.data) {
            const dataItem = {
                lodKey: key,
                dimensions: bucket.dimensions
            };

            if (ctx.isMultiYear) {
                // Multi-year mode: copy years object
                dataItem.years = {};
                for (const fy in bucket.years) {
                    dataItem.years[fy] = { ...bucket.years[fy] };
                }
            } else {
                // Two-period mode: copy py/cy
                dataItem.py = { ...bucket.py };
                dataItem.cy = { ...bucket.cy };
            }

            aggregatedData.push(dataItem);
        }

        const result = {
            data: aggregatedData,
            negatives: ctx.negatives,
            stats: ctx.stats,
            dateRange: ctx.dateRange,
            isMultiYear: ctx.isMultiYear,
            config: {
                dimensions: ctx.config.dimensions
            }
        };

        if (ctx.isMultiYear) {
            result.config.fiscalYears = ctx.config.fiscalYears;
            result.config.fyEndMonth = ctx.config.fyEndMonth;
        } else {
            result.config.pyRange = ctx.config.pyRange;
            result.config.cyRange = ctx.config.cyRange;
        }

        return result;
    },

    /**
     * Calculate totals across all LOD combinations
     *
     * @param {Object[]} aggregatedData - Array of aggregated buckets
     * @param {boolean} isMultiYear - Whether this is multi-year data
     * @returns {Object} - Totals { py: {...}, cy: {...} } or { years: {2016: {...}, 2017: {...}} }
     */
    calculateTotals: function(aggregatedData, isMultiYear = false) {
        if (isMultiYear) {
            // Multi-year mode: calculate totals for each year
            const totals = { years: {} };

            // First pass: discover all years
            for (const bucket of aggregatedData) {
                for (const fy in bucket.years) {
                    if (!totals.years[fy]) {
                        totals.years[fy] = { sales: 0, quantity: 0, cost: 0, count: 0 };
                    }
                }
            }

            // Second pass: sum up values
            for (const bucket of aggregatedData) {
                for (const fy in bucket.years) {
                    totals.years[fy].sales += bucket.years[fy].sales;
                    totals.years[fy].quantity += bucket.years[fy].quantity;
                    totals.years[fy].cost += bucket.years[fy].cost;
                    totals.years[fy].count += bucket.years[fy].count;
                }
            }

            return totals;
        } else {
            // Two-period mode: calculate PY/CY totals
            const totals = {
                py: { sales: 0, quantity: 0, cost: 0, count: 0 },
                cy: { sales: 0, quantity: 0, cost: 0, count: 0 }
            };

            for (const bucket of aggregatedData) {
                totals.py.sales += bucket.py.sales;
                totals.py.quantity += bucket.py.quantity;
                totals.py.cost += bucket.py.cost;
                totals.py.count += bucket.py.count;

                totals.cy.sales += bucket.cy.sales;
                totals.cy.quantity += bucket.cy.quantity;
                totals.cy.cost += bucket.cy.cost;
                totals.cy.count += bucket.cy.count;
            }

            return totals;
        }
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = Aggregator;
}
