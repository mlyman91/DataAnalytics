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
        // Create workbook with cell styles enabled
        const wb = XLSX.utils.book_new();
        wb.Workbook = wb.Workbook || {};
        wb.Workbook.Views = wb.Workbook.Views || [{}];
        wb.Workbook.Views[0] = wb.Workbook.Views[0] || {};
        
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

        if (summary.years) {
            // Multi-year mode
            const years = Object.keys(summary.years).sort((a, b) => Number(a) - Number(b));

            // Period summary
            data.push(['Period Summary']);
            data.push(['Fiscal Year', 'Value', 'Quantity', 'Transactions']);
            for (const year of years) {
                const yearData = summary.years[year];
                data.push([
                    `FY ${year}`,
                    yearData.value,
                    yearData.quantity,
                    yearData.count
                ]);
            }
            data.push([]);

            // Year-over-year bridge analysis
            data.push(['Year-over-Year Bridge Analysis']);
            data.push(['Component', 'Impact ($)', '% of Change']);

            // Add each year with bridge impacts to next year
            for (let i = 0; i < years.length; i++) {
                const year = years[i];
                const yearData = summary.years[year];

                // Year total row
                data.push([`FY ${year}`, yearData.value, '']);

                // If not the last year, show bridge to next year
                if (i < years.length - 1) {
                    const nextYear = years[i + 1];
                    const bridgeKey = `${year}-${nextYear}`;
                    const bridge = summary.bridges[bridgeKey];

                    if (bridge) {
                        // Calculate percentages based on starting year value
                        const startValue = yearData.value;
                        const pricePct = startValue !== 0 ? (bridge.priceImpact / Math.abs(startValue)) * 100 : 0;
                        const volumePct = startValue !== 0 ? (bridge.volumeImpact / Math.abs(startValue)) * 100 : 0;
                        const mixPct = startValue !== 0 ? (bridge.mixImpact / Math.abs(startValue)) * 100 : 0;

                        data.push([`${year}→${nextYear} Price Impact`, bridge.priceImpact, pricePct / 100]);
                        data.push([`${year}→${nextYear} Volume Impact`, bridge.volumeImpact, volumePct / 100]);
                        data.push([`${year}→${nextYear} Mix Impact`, bridge.mixImpact, mixPct / 100]);

                        if (config.mode === 'gm' && bridge.costImpact !== 0) {
                            const costPct = startValue !== 0 ? (bridge.costImpact / Math.abs(startValue)) * 100 : 0;
                            data.push([`${year}→${nextYear} Cost Impact`, bridge.costImpact, costPct / 100]);
                        }
                    }
                }
            }

            data.push([]);

            // LOD counts
            data.push(['LOD Analysis Summary']);
            data.push(['Category', 'Count']);
            data.push(['Total LOD Combinations', summary.counts.total]);

        } else {
            // Two-period mode: original behavior
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

            // Negative values row (only for two-period mode)
            if (negatives.py && negatives.cy) {
                const negTotal = negatives.cy.sales - negatives.py.sales;
                if (negTotal !== 0) {
                    data.push(['Negative Values (excluded from above)', negTotal, '']);
                }
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
        }
        
        const ws = XLSX.utils.aoa_to_sheet(data);

        // Apply number formatting to value columns
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = XLSX.utils.encode_cell({r: R, c: C});
                if (!ws[cell_address]) continue;

                // Format based on column position
                if (typeof ws[cell_address].v === 'number') {
                    if (C === 1) {
                        // Column 1: Currency format
                        ws[cell_address].z = '$#,##0';
                        ws[cell_address].t = 'n';
                    } else if (C === 2) {
                        // Column 2: Percentage format
                        ws[cell_address].z = '0.0%';
                        ws[cell_address].t = 'n';
                    } else if (C >= 3) {
                        // Other number columns: comma format
                        ws[cell_address].z = '#,##0';
                        ws[cell_address].t = 'n';
                    }
                }
            }
        }

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
            // Multi-year mode: header grouped by KPI with blank columns between groups for visual separation
            headerRow = [...dimensions];

            // Group 1: Annual Sales for all years
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length; i++) {
                const fyLabel = fiscalYears[i].label || `FY ${fiscalYears[i].fiscalYear}`;
                headerRow.push(`${fyLabel} Sales`);
            }

            // Group 2: Annual Volume for all years
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length; i++) {
                const fyLabel = fiscalYears[i].label || `FY ${fiscalYears[i].fiscalYear}`;
                headerRow.push(`${fyLabel} Volume`);
            }

            // Group 3: Annual Avg Price for all years
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length; i++) {
                const fyLabel = fiscalYears[i].label || `FY ${fiscalYears[i].fiscalYear}`;
                headerRow.push(`${fyLabel} Avg Price`);
            }

            // Group 4: Price Impacts (YoY)
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length - 1; i++) {
                const fy = fiscalYears[i];
                const nextFY = fiscalYears[i + 1];
                const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;
                headerRow.push(`${bridgeLabel} Price Impact`);
            }

            // Group 5: Volume Impacts (YoY)
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length - 1; i++) {
                const fy = fiscalYears[i];
                const nextFY = fiscalYears[i + 1];
                const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;
                headerRow.push(`${bridgeLabel} Volume Impact`);
            }

            // Group 6: Mix Impacts (YoY)
            headerRow.push(''); // Blank column for separation
            for (let i = 0; i < fiscalYears.length - 1; i++) {
                const fy = fiscalYears[i];
                const nextFY = fiscalYears[i + 1];
                const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;
                headerRow.push(`${bridgeLabel} Mix Impact`);
            }

            // Group 7: Cost Impacts (YoY) - if GM mode
            if (config.mode === 'gm') {
                headerRow.push(''); // Blank column for separation
                for (let i = 0; i < fiscalYears.length - 1; i++) {
                    const fy = fiscalYears[i];
                    const nextFY = fiscalYears[i + 1];
                    const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;
                    headerRow.push(`${bridgeLabel} Cost Impact`);
                }
            }
        } else{
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
                // Multi-year mode: populate grouped by KPI with blank columns between groups
                const dataRow = [...dimensions.map(d => row.dimensions[d] || 'Unknown')];

                // Group 1: Annual Sales for all years
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length; j++) {
                    const fy = fiscalYears[j].fiscalYear;
                    const yearData = row.years[fy] || { sales: 0, volume: 0, price: 0 };
                    dataRow.push(yearData.sales);
                }

                // Group 2: Annual Volume for all years
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length; j++) {
                    const fy = fiscalYears[j].fiscalYear;
                    const yearData = row.years[fy] || { sales: 0, volume: 0, price: 0 };
                    dataRow.push(yearData.volume);
                }

                // Group 3: Annual Avg Price for all years (will be replaced with formulas)
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length; j++) {
                    dataRow.push(0); // Placeholder for formula
                }

                // Group 4: Price Impacts (YoY) - placeholders for formulas
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length - 1; j++) {
                    dataRow.push(0);
                }

                // Group 5: Volume Impacts (YoY) - placeholders for formulas
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length - 1; j++) {
                    dataRow.push(0);
                }

                // Group 6: Mix Impacts (YoY) - placeholders for formulas
                dataRow.push(''); // Blank column for separation
                for (let j = 0; j < fiscalYears.length - 1; j++) {
                    dataRow.push(0);
                }

                // Group 7: Cost Impacts (YoY) - placeholders for formulas (if GM mode)
                if (config.mode === 'gm') {
                    dataRow.push(''); // Blank column for separation
                    for (let j = 0; j < fiscalYears.length - 1; j++) {
                        dataRow.push(0);
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
            // Multi-year mode: add formulas for grouped columns
            // Note: Each group has a blank separator column before it
            const numYears = fiscalYears.length;
            const numBridges = numYears - 1;
            const numCostImpactCols = config.mode === 'gm' ? 1 : 0;

            // Calculate column positions for each group (accounting for blank separator columns)
            const salesStartCol = dimCount + 1; // +1 for blank column before sales group
            const volumeStartCol = salesStartCol + numYears + 1; // +1 for blank column before volume group
            const priceStartCol = volumeStartCol + numYears + 1; // +1 for blank column before price group
            const priceImpactStartCol = priceStartCol + numYears + 1; // +1 for blank column before price impact group
            const volumeImpactStartCol = priceImpactStartCol + numBridges + 1; // +1 for blank column
            const mixImpactStartCol = volumeImpactStartCol + numBridges + 1; // +1 for blank column
            const costImpactStartCol = config.mode === 'gm' ? mixImpactStartCol + numBridges + 1 : -1;

            for (let i = 0; i < detail.length; i++) {
                const rowNum = i + 2; // Excel row number (1-indexed, +1 for header)

                // Add Avg Price formulas: Sales / Volume for each year
                for (let j = 0; j < numYears; j++) {
                    const salesCol = salesStartCol + j;
                    const volCol = volumeStartCol + j;
                    const priceCol = priceStartCol + j;

                    ws[`${colLetter(priceCol)}${rowNum}`] = {
                        f: `IF(${colLetter(volCol)}${rowNum}=0,0,${colLetter(salesCol)}${rowNum}/${colLetter(volCol)}${rowNum})`,
                        t: 'n'
                    };
                }

                // Add YoY bridge impact formulas
                for (let j = 0; j < numBridges; j++) {
                    const thisSalesCol = salesStartCol + j;
                    const thisVolCol = volumeStartCol + j;
                    const thisPriceCol = priceStartCol + j;
                    const nextSalesCol = salesStartCol + j + 1;
                    const nextVolCol = volumeStartCol + j + 1;
                    const nextPriceCol = priceStartCol + j + 1;

                    const priceImpactCol = priceImpactStartCol + j;
                    const volumeImpactCol = volumeImpactStartCol + j;
                    const mixImpactCol = mixImpactStartCol + j;

                    // Price Impact = (Next Price - This Price) × This Volume
                    ws[`${colLetter(priceImpactCol)}${rowNum}`] = {
                        f: `(${colLetter(nextPriceCol)}${rowNum}-${colLetter(thisPriceCol)}${rowNum})*${colLetter(thisVolCol)}${rowNum}`,
                        t: 'n'
                    };

                    // Volume Impact = (Next Volume - This Volume) × This Price
                    ws[`${colLetter(volumeImpactCol)}${rowNum}`] = {
                        f: `(${colLetter(nextVolCol)}${rowNum}-${colLetter(thisVolCol)}${rowNum})*${colLetter(thisPriceCol)}${rowNum}`,
                        t: 'n'
                    };

                    // Mix Impact = (Next Sales - This Sales) - Price Impact - Volume Impact
                    ws[`${colLetter(mixImpactCol)}${rowNum}`] = {
                        f: `(${colLetter(nextSalesCol)}${rowNum}-${colLetter(thisSalesCol)}${rowNum})-${colLetter(priceImpactCol)}${rowNum}-${colLetter(volumeImpactCol)}${rowNum}`,
                        t: 'n'
                    };

                    if (config.mode === 'gm') {
                        // Cost Impact (placeholder for now)
                        const costImpactCol = costImpactStartCol + j;
                        ws[`${colLetter(costImpactCol)}${rowNum}`] = { v: 0, t: 'n' };
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
            } else if (headerRow[i] === '') {
                // Narrow separator columns
                cols.push({ wch: 2 });
            } else {
                cols.push({ wch: CONFIG.EXCEL.COL_WIDTH_NUMBER });
            }
        }
        ws['!cols'] = cols;

        // Apply number formatting to data cells
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {  // Start from row 1 (skip header)
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = XLSX.utils.encode_cell({r: R, c: C});
                if (!ws[cell_address] || ws[cell_address].t !== 'n') continue;

                const header = headerRow[C];
                if (!header) continue;

                // Volume columns: comma format without currency
                if (header.includes('Vol') || header.includes('Volume')) {
                    ws[cell_address].z = '#,##0';
                }
                // Sales, Price, Impact, Change columns: currency format
                else if (header.includes('Sales') || header.includes('Price') ||
                         header.includes('Impact') || header.includes('Change') ||
                         header.includes('Δ')) {
                    ws[cell_address].z = '$#,##0';
                }
            }
        }

        // Note: Cell styling (background colors, borders) requires SheetJS Pro.
        // We use blank separator columns (width: 2) between KPI groups for visual separation.
        // Users can manually format the header row in Excel after opening the file.

        // Remove gridlines
        ws['!cols'] = ws['!cols'] || cols;
        ws['!views'] = [{ showGridLines: false }];

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

        // Check if multi-year mode
        const isMultiYear = !negatives.py && !negatives.cy;

        if (isMultiYear) {
            // Multi-year mode: negatives keyed by year
            const years = Object.keys(negatives).sort();
            const headerRow = ['Metric', ...years.map(y => `FY ${y}`), 'Total'];

            // Calculate totals
            let totalCount = 0;
            let totalSales = 0;
            let totalQty = 0;
            let totalCost = 0;

            for (const year of years) {
                totalCount += negatives[year].count;
                totalSales += negatives[year].sales;
                totalQty += negatives[year].quantity;
                totalCost += negatives[year].cost;
            }

            data.push(['Summary']);
            data.push(headerRow);

            data.push(['Excluded Row Count', ...years.map(y => negatives[y].count), totalCount]);
            data.push(['Excluded Sales Total', ...years.map(y => negatives[y].sales), totalSales]);
            data.push(['Excluded Quantity Total', ...years.map(y => negatives[y].quantity), totalQty]);
            data.push(['Excluded Cost Total', ...years.map(y => negatives[y].cost), totalCost]);
            data.push([]);

            // Processing stats
            data.push(['Processing Statistics']);
            data.push(['Metric', 'Value']);
            data.push(['Total Rows Read', stats.totalRows]);
            data.push(['Rows Included in Analysis', stats.includedRows]);
            data.push(['Rows Excluded (negative/zero)', totalCount]);
            data.push(['Rows Outside Analysis Periods', stats.outsidePeriodRows]);
            data.push(['Rows with Parse Errors', stats.parseErrors]);

        } else {
            // Two-period mode: original behavior
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
        }

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

        // Add period information based on mode
        if (config.fiscalYears && config.fiscalYears.length > 0) {
            // Multi-year mode
            const yearLabels = config.fiscalYears.map(fy => fy.label || `FY ${fy.fiscalYear}`).join(', ');
            const firstYear = config.fiscalYears[0];
            const lastYear = config.fiscalYears[config.fiscalYears.length - 1];
            data.push(['Analysis Type', 'Multi-Year Comparison']);
            data.push(['Fiscal Years', yearLabels]);
            data.push(['Date Range', `${PeriodUtils.formatDate(firstYear.start)} to ${PeriodUtils.formatDate(lastYear.end)}`]);
        } else if (config.pyRange && config.cyRange) {
            // Two-period mode
            data.push(['Prior Year Period', PeriodUtils.formatDateRange(config.pyRange.start, config.pyRange.end)]);
            data.push(['LTM Period', PeriodUtils.formatDateRange(config.cyRange.start, config.cyRange.end)]);
        }

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
