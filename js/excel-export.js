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
        const pyLabel = config.pyLabel || 'Prior Year';
        const cyLabel = config.cyLabel || 'Current Year';

        data.push(['Period Summary']);
        data.push(['Period', 'Start Date', 'End Date', 'Value', 'Quantity', 'Transactions']);
        data.push([
            pyLabel,
            PeriodUtils.formatDate(config.pyRange.start),
            PeriodUtils.formatDate(config.pyRange.end),
            summary.py.value,
            summary.py.quantity,
            summary.py.count
        ]);
        data.push([
            cyLabel,
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
        data.push([`${pyLabel} Starting Value`, summary.py.value, '']);
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

        data.push([`${cyLabel} Ending Value`, summary.cy.value, '']);
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

        // Check if multi-year mode
        const isMultiYear = config.fiscalYears && config.fiscalYears.length > 0;
        const fiscalYears = isMultiYear ? config.fiscalYears : [];

        let headerRow;

        if (isMultiYear) {
            // Multi-year mode: header with all years and YoY bridges
            headerRow = [...dimensions];

            for (let i = 0; i < fiscalYears.length; i++) {
                const fy = fiscalYears[i];
                const fyLabel = fy.label || `FY ${fy.fiscalYear}`;

                // Year columns
                headerRow.push(`${fyLabel} Sales`);
                headerRow.push(`${fyLabel} Volume`);
                headerRow.push(`${fyLabel} Avg Price`);

                // YoY bridge columns (between this year and next)
                if (i < fiscalYears.length - 1) {
                    const nextFY = fiscalYears[i + 1];
                    const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;
                    headerRow.push(`${bridgeLabel} Price Δ`);
                    headerRow.push(`${bridgeLabel} Vol Δ`);
                    headerRow.push(`${bridgeLabel} Mix Δ`);

                    if (config.mode === 'gm') {
                        headerRow.push(`${bridgeLabel} Cost Δ`);
                    }
                }
            }
        } else {
            // Two-period mode: original behavior
            const pyLabel = config.pyLabel || 'PY';
            const cyLabel = config.cyLabel || 'CY';

            headerRow = [
                ...dimensions,
                'Status',
                `${pyLabel} Sales`,
                `${pyLabel} Volume`,
                `${pyLabel} Avg Price`,
                `${cyLabel} Sales`,
                `${cyLabel} Volume`,
                `${cyLabel} Avg Price`,
                'Total Change',
                'Price Impact',
                'Volume Impact',
                'Mix Impact'
            ];

            if (config.mode === 'gm') {
                headerRow.push('Cost Impact');
            }
        }

        data.push(headerRow);

        // Data rows
        for (let i = 0; i < detail.length; i++) {
            const row = detail[i];

            if (isMultiYear) {
                // Multi-year mode: populate all years
                const dataRow = [...dimensions.map(d => row.dimensions[d] || 'Unknown')];

                for (let j = 0; j < fiscalYears.length; j++) {
                    const fy = fiscalYears[j].fiscalYear;
                    const yearData = row.years[fy] || { sales: 0, volume: 0, price: 0 };

                    // Year data
                    dataRow.push(yearData.sales);
                    dataRow.push(yearData.volume);
                    dataRow.push(yearData.price);

                    // YoY bridge values (will be replaced with formulas below)
                    if (j < fiscalYears.length - 1) {
                        dataRow.push(0); // Price Δ (formula)
                        dataRow.push(0); // Vol Δ (formula)
                        dataRow.push(0); // Mix Δ (formula)

                        if (config.mode === 'gm') {
                            dataRow.push(0); // Cost Δ (formula)
                        }
                    }
                }

                data.push(dataRow);
            } else {
                // Two-period mode: original behavior
                const dataRow = [
                    ...dimensions.map(d => row.dimensions[d] || 'Unknown'),
                    row.classification,
                    row.py.sales,
                    row.py.volume,
                    row.py.price,
                    row.cy.sales,
                    row.cy.volume,
                    row.cy.price
                ];

                data.push(dataRow);
            }
        }

        // Create worksheet from data
        const ws = XLSX.utils.aoa_to_sheet(data);

        // Column letter helper
        const colLetter = (col) => {
            let letter = '';
            let num = col;
            while (num >= 0) {
                letter = String.fromCharCode((num % 26) + 65) + letter;
                num = Math.floor(num / 26) - 1;
            }
            return letter;
        };

        // Now add formulas for bridge calculations
        const dimCount = dimensions.length;

        if (isMultiYear) {
            // Multi-year mode: add formulas for YoY bridges
            for (let i = 0; i < detail.length; i++) {
                const rowNum = i + 2; // Excel row number (1-indexed, +1 for header)
                let colIndex = dimCount;

                for (let j = 0; j < fiscalYears.length; j++) {
                    // Year columns: Sales, Volume, Price
                    const salesCol = colIndex++;
                    const volCol = colIndex++;
                    const priceCol = colIndex++;

                    // YoY bridge formulas (between this year and next)
                    if (j < fiscalYears.length - 1) {
                        // Get next year column indices
                        const nextSalesCol = colIndex + (config.mode === 'gm' ? 4 : 3);
                        const nextVolCol = nextSalesCol + 1;
                        const nextPriceCol = nextSalesCol + 2;

                        // Price Impact = (Next Price - This Price) × This Volume
                        const priceImpactCol = colIndex++;
                        ws[`${colLetter(priceImpactCol)}${rowNum}`] = {
                            f: `(${colLetter(nextPriceCol)}${rowNum}-${colLetter(priceCol)}${rowNum})*${colLetter(volCol)}${rowNum}`,
                            t: 'n'
                        };

                        // Volume Impact = (Next Volume - This Volume) × This Price
                        const volImpactCol = colIndex++;
                        ws[`${colLetter(volImpactCol)}${rowNum}`] = {
                            f: `(${colLetter(nextVolCol)}${rowNum}-${colLetter(volCol)}${rowNum})*${colLetter(priceCol)}${rowNum}`,
                            t: 'n'
                        };

                        // Mix Impact = (Next Sales - This Sales) - Price Impact - Volume Impact
                        const mixImpactCol = colIndex++;
                        ws[`${colLetter(mixImpactCol)}${rowNum}`] = {
                            f: `(${colLetter(nextSalesCol)}${rowNum}-${colLetter(salesCol)}${rowNum})-${colLetter(priceImpactCol)}${rowNum}-${colLetter(volImpactCol)}${rowNum}`,
                            t: 'n'
                        };

                        if (config.mode === 'gm') {
                            // Cost Impact (placeholder for now)
                            const costImpactCol = colIndex++;
                            ws[`${colLetter(costImpactCol)}${rowNum}`] = { v: 0, t: 'n' };
                        }
                    }
                }
            }
        } else {
            // Two-period mode: original formulas
            const statusCol = dimCount;
            const pySalesCol = dimCount + 1;
            const pyVolCol = dimCount + 2;
            const pyPriceCol = dimCount + 3;
            const cySalesCol = dimCount + 4;
            const cyVolCol = dimCount + 5;
            const cyPriceCol = dimCount + 6;
            const totalChangeCol = dimCount + 7;
            const priceImpactCol = dimCount + 8;
            const volumeImpactCol = dimCount + 9;
            const mixImpactCol = dimCount + 10;

            // Add formulas for each row
            for (let i = 0; i < detail.length; i++) {
                const rowNum = i + 2; // Excel row number (1-indexed, +1 for header)

                const pySales = `${colLetter(pySalesCol)}${rowNum}`;
                const pyVol = `${colLetter(pyVolCol)}${rowNum}`;
                const pyPrice = `${colLetter(pyPriceCol)}${rowNum}`;
                const cySales = `${colLetter(cySalesCol)}${rowNum}`;
                const cyVol = `${colLetter(cyVolCol)}${rowNum}`;
                const cyPrice = `${colLetter(cyPriceCol)}${rowNum}`;
                const totalChange = `${colLetter(totalChangeCol)}${rowNum}`;
                const priceImpact = `${colLetter(priceImpactCol)}${rowNum}`;
                const volumeImpact = `${colLetter(volumeImpactCol)}${rowNum}`;

                // Total Change = CY Sales - PY Sales
                ws[totalChange] = { f: `${cySales}-${pySales}`, t: 'n' };

                // Price Impact = (CY Price - PY Price) × PY Volume
                ws[priceImpact] = { f: `(${cyPrice}-${pyPrice})*${pyVol}`, t: 'n' };

                // Volume Impact = (CY Volume - PY Volume) × PY Price
                ws[volumeImpact] = { f: `(${cyVol}-${pyVol})*${pyPrice}`, t: 'n' };

                // Mix Impact = Total Change - Price Impact - Volume Impact
                ws[`${colLetter(mixImpactCol)}${rowNum}`] = { f: `${totalChange}-${priceImpact}-${volumeImpact}`, t: 'n' };

                if (config.mode === 'gm') {
                    // Add Cost Impact formula if needed
                    // For now, we'll skip this as it's more complex
                }
            }
        }
        
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
