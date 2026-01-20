/**
 * ============================================
 * PVM Bridge Tool - Bridge Calculator
 * ============================================
 * 
 * Core PVM and Gross Margin bridge calculations.
 * All formulas are documented for auditability.
 * 
 * PVM Bridge Formulas:
 * - Price Impact = (CY Price - PY Price) × PY Volume
 * - Volume Impact = (CY Volume - PY Volume) × PY Price
 * - Mix Impact = Total Change - Price Impact - Volume Impact
 * 
 * Special Cases:
 * - New items (PY = 0): All impact goes to Volume
 * - Discontinued (CY = 0): All impact goes to Volume
 * ============================================
 */

const BridgeCalculator = {
    /**
     * Calculate PVM bridge for aggregated data
     * 
     * @param {Object[]} aggregatedData - Array of LOD buckets
     * @param {Object} options - Calculation options
     * @param {string} options.mode - 'pvm' or 'gm'
     * @param {string} options.gmPriceDefinition - For GM mode: 'margin-per-unit' or 'sales-per-unit'
     * @returns {Object} - Bridge results
     */
    calculate: function(aggregatedData, options = {}) {
        const mode = options.mode || 'pvm';
        const gmPriceDef = options.gmPriceDefinition || 'margin-per-unit';
        
        // Calculate bridge for each LOD combination
        const detailResults = aggregatedData.map(bucket => {
            return this._calculateBucket(bucket, mode, gmPriceDef);
        });
        
        // Calculate summary totals
        const summary = this._calculateSummary(detailResults, mode);
        
        return {
            detail: detailResults,
            summary: summary,
            mode: mode,
            gmPriceDefinition: gmPriceDef
        };
    },

    /**
     * Calculate bridge for a single LOD bucket
     * 
     * @private
     */
    _calculateBucket: function(bucket, mode, gmPriceDef) {
        const py = bucket.py;
        const cy = bucket.cy;
        
        // Determine if this is new, discontinued, or continuing
        const isNew = py.sales <= 0 || py.quantity <= 0;
        const isDiscontinued = cy.sales <= 0 || cy.quantity <= 0;
        
        // Get the "value" based on mode
        let pyValue, cyValue;
        
        if (mode === 'gm') {
            if (gmPriceDef === 'margin-per-unit') {
                // Value = Gross Margin = Sales - Cost
                pyValue = py.sales - py.cost;
                cyValue = cy.sales - cy.cost;
            } else {
                // Value = Sales (cost handled separately)
                pyValue = py.sales;
                cyValue = cy.sales;
            }
        } else {
            // PVM mode: Value = Sales (Revenue)
            pyValue = py.sales;
            cyValue = cy.sales;
        }
        
        // Calculate prices (value per unit)
        const pyPrice = py.quantity > 0 ? pyValue / py.quantity : 0;
        const cyPrice = cy.quantity > 0 ? cyValue / cy.quantity : 0;
        
        // Volumes
        const pyVolume = py.quantity;
        const cyVolume = cy.quantity;
        
        // Total change
        const totalChange = cyValue - pyValue;
        
        // Initialize impacts
        let priceImpact = 0;
        let volumeImpact = 0;
        let mixImpact = 0;
        let costImpact = 0; // Only used in GM mode with sales-per-unit
        
        // Classification for audit trail
        let classification = 'continuing';
        
        if (isNew) {
            // New item: all impact goes to volume
            classification = 'new';
            priceImpact = 0;
            volumeImpact = totalChange;
            mixImpact = 0;
        } else if (isDiscontinued) {
            // Discontinued: all impact goes to volume
            classification = 'discontinued';
            priceImpact = 0;
            volumeImpact = totalChange;
            mixImpact = 0;
        } else {
            // Continuing item: calculate all three components
            classification = 'continuing';
            
            // Price Impact = (CY Price - PY Price) × PY Volume
            priceImpact = (cyPrice - pyPrice) * pyVolume;
            
            // Volume Impact = (CY Volume - PY Volume) × PY Price
            volumeImpact = (cyVolume - pyVolume) * pyPrice;
            
            // Mix Impact = Residual (ensures exact reconciliation)
            mixImpact = totalChange - priceImpact - volumeImpact;
        }
        
        // For GM mode with sales-per-unit, calculate cost impact
        if (mode === 'gm' && gmPriceDef === 'sales-per-unit') {
            costImpact = -(cy.cost - py.cost);
        }
        
        return {
            lodKey: bucket.lodKey,
            dimensions: bucket.dimensions,
            
            // Period values
            py: {
                value: pyValue,
                price: pyPrice,
                volume: pyVolume,
                sales: py.sales,
                cost: py.cost,
                count: py.count
            },
            cy: {
                value: cyValue,
                price: cyPrice,
                volume: cyVolume,
                sales: cy.sales,
                cost: cy.cost,
                count: cy.count
            },
            
            // Bridge components
            totalChange: totalChange,
            priceImpact: priceImpact,
            volumeImpact: volumeImpact,
            mixImpact: mixImpact,
            costImpact: costImpact,
            
            // Classification
            classification: classification,
            isNew: isNew,
            isDiscontinued: isDiscontinued
        };
    },

    /**
     * Calculate summary totals from detail results
     * 
     * @private
     */
    _calculateSummary: function(detailResults, mode) {
        const summary = {
            py: { value: 0, sales: 0, quantity: 0, cost: 0, count: 0 },
            cy: { value: 0, sales: 0, quantity: 0, cost: 0, count: 0 },
            
            totalChange: 0,
            priceImpact: 0,
            volumeImpact: 0,
            mixImpact: 0,
            costImpact: 0,
            
            counts: {
                total: detailResults.length,
                new: 0,
                discontinued: 0,
                continuing: 0
            }
        };
        
        for (const result of detailResults) {
            // Sum period values
            summary.py.value += result.py.value;
            summary.py.sales += result.py.sales;
            summary.py.quantity += result.py.volume;
            summary.py.cost += result.py.cost;
            summary.py.count += result.py.count;
            
            summary.cy.value += result.cy.value;
            summary.cy.sales += result.cy.sales;
            summary.cy.quantity += result.cy.volume;
            summary.cy.cost += result.cy.cost;
            summary.cy.count += result.cy.count;
            
            // Sum bridge components
            summary.totalChange += result.totalChange;
            summary.priceImpact += result.priceImpact;
            summary.volumeImpact += result.volumeImpact;
            summary.mixImpact += result.mixImpact;
            summary.costImpact += result.costImpact;
            
            // Count classifications
            summary.counts[result.classification]++;
        }
        
        // Calculate percentages
        const totalChange = summary.totalChange;
        summary.priceImpactPct = totalChange !== 0 ? (summary.priceImpact / Math.abs(totalChange)) * 100 : 0;
        summary.volumeImpactPct = totalChange !== 0 ? (summary.volumeImpact / Math.abs(totalChange)) * 100 : 0;
        summary.mixImpactPct = totalChange !== 0 ? (summary.mixImpact / Math.abs(totalChange)) * 100 : 0;
        summary.costImpactPct = totalChange !== 0 ? (summary.costImpact / Math.abs(totalChange)) * 100 : 0;
        
        // Calculate overall change percentage
        summary.changePct = summary.py.value !== 0 ? ((summary.cy.value - summary.py.value) / Math.abs(summary.py.value)) * 100 : 0;
        
        return summary;
    },

    /**
     * Sort detail results by specified criteria
     * 
     * @param {Object[]} detailResults - Array of detail results
     * @param {string} sortBy - Sort criteria
     * @returns {Object[]} - Sorted results
     */
    sortResults: function(detailResults, sortBy) {
        const sorted = [...detailResults];
        
        switch (sortBy) {
            case 'total-desc':
                sorted.sort((a, b) => Math.abs(b.totalChange) - Math.abs(a.totalChange));
                break;
            case 'total-asc':
                sorted.sort((a, b) => Math.abs(a.totalChange) - Math.abs(b.totalChange));
                break;
            case 'price-desc':
                sorted.sort((a, b) => Math.abs(b.priceImpact) - Math.abs(a.priceImpact));
                break;
            case 'volume-desc':
                sorted.sort((a, b) => Math.abs(b.volumeImpact) - Math.abs(a.volumeImpact));
                break;
            case 'mix-desc':
                sorted.sort((a, b) => Math.abs(b.mixImpact) - Math.abs(a.mixImpact));
                break;
            default:
                // No sorting
                break;
        }
        
        return sorted;
    },

    /**
     * Filter detail results by search term
     * 
     * @param {Object[]} detailResults - Array of detail results
     * @param {string} searchTerm - Search term
     * @returns {Object[]} - Filtered results
     */
    filterResults: function(detailResults, searchTerm) {
        if (!searchTerm || searchTerm.trim() === '') {
            return detailResults;
        }
        
        const term = searchTerm.toLowerCase().trim();
        
        return detailResults.filter(result => {
            // Search in dimension values
            for (const value of Object.values(result.dimensions)) {
                if (String(value).toLowerCase().includes(term)) {
                    return true;
                }
            }
            return false;
        });
    },

    /**
     * Get methodology description for assumptions tab
     */
    getMethodologyDescription: function(mode, gmPriceDef) {
        if (mode === 'gm') {
            if (gmPriceDef === 'margin-per-unit') {
                return {
                    title: 'Gross Margin Bridge (Margin per Unit)',
                    description: 'Analyzes drivers of gross margin change using margin per unit as the price metric.',
                    formulas: [
                        { name: 'Margin per Unit (Price)', formula: '(Sales − Cost) / Quantity' },
                        { name: 'Price Impact', formula: '(CY Margin/Unit − PY Margin/Unit) × PY Volume' },
                        { name: 'Volume Impact', formula: '(CY Volume − PY Volume) × PY Margin/Unit' },
                        { name: 'Mix Impact', formula: 'Total Change − Price Impact − Volume Impact' }
                    ],
                    notes: [
                        'New items (no PY data): Entire change attributed to Volume',
                        'Discontinued items (no CY data): Entire change attributed to Volume',
                        'Mix Impact captures the interaction effect and ensures exact reconciliation'
                    ]
                };
            } else {
                return {
                    title: 'Gross Margin Bridge (Sales per Unit)',
                    description: 'Analyzes drivers of gross margin change with cost impact shown separately.',
                    formulas: [
                        { name: 'Sales per Unit (Price)', formula: 'Sales / Quantity' },
                        { name: 'Price Impact', formula: '(CY Price − PY Price) × PY Volume' },
                        { name: 'Volume Impact', formula: '(CY Volume − PY Volume) × PY Price' },
                        { name: 'Mix Impact', formula: 'Sales Change − Price Impact − Volume Impact' },
                        { name: 'Cost Impact', formula: '−(CY Cost − PY Cost)' }
                    ],
                    notes: [
                        'New items (no PY data): Entire change attributed to Volume',
                        'Discontinued items (no CY data): Entire change attributed to Volume',
                        'Cost Impact is shown separately from the PVM decomposition'
                    ]
                };
            }
        } else {
            return {
                title: 'Sales PVM Bridge',
                description: 'Decomposes revenue change into Price, Volume, and Mix components.',
                formulas: [
                    { name: 'Average Price', formula: 'Sales / Quantity' },
                    { name: 'Price Impact', formula: '(CY Price − PY Price) × PY Volume' },
                    { name: 'Volume Impact', formula: '(CY Volume − PY Volume) × PY Price' },
                    { name: 'Mix Impact', formula: 'Total Change − Price Impact − Volume Impact' }
                ],
                notes: [
                    'New items (no PY data): Entire change attributed to Volume',
                    'Discontinued items (no CY data): Entire change attributed to Volume',
                    'Mix Impact captures both product mix shifts and the interaction between price and volume changes',
                    'The three components sum exactly to the total revenue change'
                ]
            };
        }
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = BridgeCalculator;
}
