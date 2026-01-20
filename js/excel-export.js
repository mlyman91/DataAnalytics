/**
 * ============================================
 * PVM Bridge Tool - Excel Export
 * ============================================
 * 
 * Generates XLSX files with multiple tabs:
 * 1. Summary Bridge
 * 2. Detail by LOD
 * 3. Negative Values
 * 4. Assumptions & Metadata
 * 
 * Uses SheetJS (xlsx.mini.js) for Excel generation.
 * All data is numeric/text for easy charting.
 * ============================================
 */

const ExcelExport = {
    /**
     * Export results to Excel file
     * 
     * @param {Object} bridgeResults - Results from BridgeCalculator
     * @param {Object} aggregationResults - Results from Aggregator
     * @param {Object} config - Analysis configuration
     */
    exportToExcel: function(bridgeResults, aggregationResults, config) {
        // Create workbook
        const wb = XLSX.utils.book_new();
        
        // Tab 1: Summary Bridge
        const summarySheet = this._createSummarySheet(bridgeResults.summary, config, aggregationResults.negatives);
        XLSX.utils.book_append_sheet(wb, summarySheet, CONFIG.EXCEL.TAB_NAMES.SUMMARY);
        
        // Tab 2: Detail by LOD
        const detailSheet = this._createDetailSheet(bridgeResults.detail, config);
        XLSX.utils.book_append_sheet(wb, detailSheet, CONFIG.EXCEL.TAB_NAMES.DETAIL);
        
        // Tab 3: Negative Values
        const negativesSheet = this._createNegativesSheet(aggregationResults.negatives, aggregationResults.stats);
        XLSX.utils.book_append_sheet(wb, negativesSheet, CONFIG.EXCEL.TAB_NAMES.NEGATIVES);
        
        // Tab 4: Assumptions
        const assumptionsSheet = this._createAssumptionsSheet(config, aggregationResults.stats, bridgeResults.mode);
        XLSX.utils.book_append_sheet(wb, assumptionsSheet, CONFIG.EXCEL.TAB_NAMES.ASSUMPTIONS);
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 10);
        const modeLabel = config.mode === 'gm' ? 'GM_Bridge' : 'PVM_Bridge';
        const filename = `${modeLabel}_${timestamp}.xlsx`;
        
        // Trigger download
        XLSX.writeFile(wb, filename);
    },

    /**
     * Create Summary Bridge sheet
     * @private
     */
    _createSummarySheet: function(summary, config, negatives) {
        const data = [];
        
        // Header section
        data.push([config.mode === 'gm' ? 'Gross Margin Bridge Summary' : 'Sales PVM Bridge Summary']);
        data.push([`Generated: ${new Date().toLocaleString()}`]);
        data.push([]);
        
        // Period summary
        data.push(['Period Summary']);
        data.push(['Period', 'Start Date', 'End Date', 'Value', 'Quantity', 'Transactions']);
        data.push([
            'Prior Year',
            PeriodUtils.formatDate(config.pyRange.start),
            PeriodUtils.formatDate(config.pyRange.end),
            summary.py.value,
            summary.py.quantity,
            summary.py.count
        ]);
        data.push([
            'Current Year / LTM',
            PeriodUtils.formatDate(config.cyRange.start),
            PeriodUtils.formatDate(config.cyRange.end),
            summary.cy.value,
            summary.cy.quantity,
            summary.cy.count
        ]);
        data.push([]);
        
        // Bridge components
        data.push(['Bridge Analysis']);
        data.push(['Component', 'Impact ($)', '% of Total Change']);
        data.push(['Prior Year Starting Value', summary.py.value, '']);
        data.push(['Price Impact', summary.priceImpact, summary.priceImpactPct / 100]);
        data.push(['Volume Impact', summary.volumeImpact, summary.volumeImpactPct / 100]);
        data.push(['Mix Impact', summary.mixImpact, summary.mixImpactPct / 100]);
        
        if (config.mode === 'gm' && summary.costImpact !== 0) {
            data.push(['Cost Impact', summary.costImpact, summary.costImpactPct / 100]);
        }
        
        // Negative values row
        const negTotal = negatives.cy.sales - negatives.py.sales;
        if (negTotal !== 0) {
            data.push(['Negative Values (excluded from above)', negTotal, '']);
        }
        
        data.push(['Current Year / LTM Ending Value', summary.cy.value, '']);
        data.push([]);
        
        // Total change summary
        data.push(['Total Change', summary.totalChange, summary.changePct / 100]);
        data.push([]);
        
        // LOD counts
        data.push(['LOD Analysis Summary']);
        data.push(['Category', 'Count']);
        data.push(['Total LOD Combinations', summary.counts.total]);
        data.push(['New Items (in CY only)', summary.counts.new]);
        data.push(['Discontinued Items (in PY only)', summary.counts.discontinued]);
        data.push(['Continuing Items', summary.counts.continuing]);
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths
        ws['!cols'] = [
            { wch: 30 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 },
            { wch: 15 },
            { wch: 15 }
        ];
        
        return ws;
    },

    /**
     * Create Detail sheet
     * @private
     */
    _createDetailSheet: function(detail, config) {
        const data = [];
        const dimensions = config.dimensions || [];
        
        // Header row
        const headerRow = [
            ...dimensions,
            'Status',
            'PY Value',
            'PY Quantity',
            'PY Price',
            'CY Value',
            'CY Quantity',
            'CY Price',
            'Total Change',
            'Price Impact',
            'Volume Impact',
            'Mix Impact'
        ];
        
        if (config.mode === 'gm') {
            headerRow.push('Cost Impact');
            headerRow.push('PY Cost');
            headerRow.push('CY Cost');
        }
        
        data.push(headerRow);
        
        // Data rows
        for (const row of detail) {
            const dataRow = [
                ...dimensions.map(d => row.dimensions[d] || 'Unknown'),
                row.classification,
                row.py.value,
                row.py.volume,
                row.py.price,
                row.cy.value,
                row.cy.volume,
                row.cy.price,
                row.totalChange,
                row.priceImpact,
                row.volumeImpact,
                row.mixImpact
            ];
            
            if (config.mode === 'gm') {
                dataRow.push(row.costImpact);
                dataRow.push(row.py.cost);
                dataRow.push(row.cy.cost);
            }
            
            data.push(dataRow);
        }
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths
        const cols = [];
        for (let i = 0; i < headerRow.length; i++) {
            if (i < dimensions.length) {
                cols.push({ wch: CONFIG.EXCEL.COL_WIDTH_DIMENSION });
            } else if (headerRow[i] === 'Status') {
                cols.push({ wch: 12 });
            } else {
                cols.push({ wch: CONFIG.EXCEL.COL_WIDTH_NUMBER });
            }
        }
        ws['!cols'] = cols;
        
        return ws;
    },

    /**
     * Create Negative Values sheet
     * @private
     */
    _createNegativesSheet: function(negatives, stats) {
        const data = [];
        
        data.push(['Negative Values Analysis']);
        data.push(['Rows with Sales ≤ 0 or Quantity ≤ 0 are excluded from PVM calculations.']);
        data.push([]);
        
        // Summary
        data.push(['Summary']);
        data.push(['Metric', 'Prior Year', 'Current Year / LTM', 'Total']);
        data.push([
            'Excluded Row Count',
            negatives.py.count,
            negatives.cy.count,
            negatives.py.count + negatives.cy.count
        ]);
        data.push([
            'Excluded Sales Total',
            negatives.py.sales,
            negatives.cy.sales,
            negatives.py.sales + negatives.cy.sales
        ]);
        data.push([
            'Excluded Quantity Total',
            negatives.py.quantity,
            negatives.cy.quantity,
            negatives.py.quantity + negatives.cy.quantity
        ]);
        data.push([
            'Excluded Cost Total',
            negatives.py.cost,
            negatives.cy.cost,
            negatives.py.cost + negatives.cy.cost
        ]);
        data.push([]);
        
        // Processing stats
        data.push(['Processing Statistics']);
        data.push(['Metric', 'Value']);
        data.push(['Total Rows Read', stats.totalRows]);
        data.push(['Rows Included in Analysis', stats.includedRows]);
        data.push(['Rows Excluded (negative/zero)', negatives.py.count + negatives.cy.count]);
        data.push(['Rows Outside Analysis Periods', stats.outsidePeriodRows]);
        data.push(['Rows with Parse Errors', stats.parseErrors]);
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        ws['!cols'] = [
            { wch: 25 },
            { wch: 18 },
            { wch: 18 },
            { wch: 18 }
        ];
        
        return ws;
    },

    /**
     * Create Assumptions sheet
     * @private
     */
    _createAssumptionsSheet: function(config, stats, mode) {
        const data = [];
        const methodology = BridgeCalculator.getMethodologyDescription(mode, config.gmPriceDefinition);
        
        data.push(['Analysis Assumptions & Methodology']);
        data.push([`Generated: ${new Date().toLocaleString()}`]);
        data.push([]);
        
        // Configuration
        data.push(['Analysis Configuration']);
        data.push(['Parameter', 'Value']);
        data.push(['Analysis Mode', mode === 'gm' ? 'Gross Margin Bridge' : 'Sales PVM Bridge']);
        
        if (mode === 'gm') {
            data.push(['GM Price Definition', 
                config.gmPriceDefinition === 'margin-per-unit' ? 
                    'Margin per Unit: (Sales - Cost) / Quantity' :
                    'Sales per Unit: Sales / Quantity'
            ]);
        }
        
        data.push(['Fiscal Year End Month', CONFIG.MONTH_NAMES[config.fyEndMonth - 1]]);
        data.push(['Prior Year Period', PeriodUtils.formatDateRange(config.pyRange.start, config.pyRange.end)]);
        data.push(['LTM Period', PeriodUtils.formatDateRange(config.cyRange.start, config.cyRange.end)]);
        data.push(['Level of Detail Dimensions', config.dimensions.length > 0 ? config.dimensions.join(', ') : 'Total Only']);
        data.push([]);
        
        // Column mappings
        data.push(['Column Mappings']);
        data.push(['Field', 'Source Column']);
        data.push(['Date', config.dateColumn]);
        data.push(['Sales', config.salesColumn]);
        data.push(['Quantity', config.quantityColumn]);
        if (config.costColumn) {
            data.push(['Cost', config.costColumn]);
        }
        data.push([]);
        
        // Methodology
        data.push(['Calculation Methodology']);
        data.push([methodology.title]);
        data.push([methodology.description]);
        data.push([]);
        
        data.push(['Formulas Used']);
        for (const f of methodology.formulas) {
            data.push([f.name, f.formula]);
        }
        data.push([]);
        
        // Rules
        data.push(['Business Rules Applied']);
        for (const note of methodology.notes) {
            data.push([note]);
        }
        data.push([]);
        
        // Data quality
        data.push(['Data Quality Notes']);
        data.push([`Total rows processed: ${stats.totalRows.toLocaleString()}`]);
        data.push([`Rows included in analysis: ${stats.includedRows.toLocaleString()}`]);
        data.push([`Rows excluded (negative/zero values): ${stats.excludedRows.toLocaleString()}`]);
        
        if (stats.outsidePeriodRows > 0) {
            data.push([`Rows outside analysis periods: ${stats.outsidePeriodRows.toLocaleString()}`]);
        }
        if (stats.parseErrors > 0) {
            data.push([`Rows with parse errors: ${stats.parseErrors.toLocaleString()}`]);
        }
        
        data.push([`Unique LOD combinations: ${stats.uniqueLODKeys.toLocaleString()}`]);
        data.push([]);
        
        // Reconciliation check
        data.push(['Reconciliation Check']);
        data.push(['The sum of Price Impact + Volume Impact + Mix Impact should equal Total Change.']);
        data.push(['Mix Impact is calculated as the residual to ensure exact reconciliation.']);
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        ws['!cols'] = [
            { wch: 40 },
            { wch: 50 }
        ];
        
        return ws;
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelExport;
}
