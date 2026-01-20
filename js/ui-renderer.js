/**
 * ============================================
 * PVM Bridge Tool - UI Renderer
 * ============================================
 * 
 * Pure DOM manipulation for rendering results.
 * No innerHTML for user data (security).
 * Accessible markup with ARIA labels.
 * ============================================
 */

const UIRenderer = {
    /**
     * Format a number for display
     * 
     * @param {number} value - Number to format
     * @param {string} type - 'currency', 'quantity', 'price', 'percent'
     * @returns {string} - Formatted string
     */
    formatNumber: function(value, type = 'currency') {
        if (value === null || value === undefined || isNaN(value)) {
            return '--';
        }
        
        let decimals;
        let prefix = '';
        let suffix = '';
        
        switch (type) {
            case 'currency':
                decimals = CONFIG.FORMAT.DECIMALS_CURRENCY;
                prefix = '$';
                break;
            case 'quantity':
                decimals = CONFIG.FORMAT.DECIMALS_QUANTITY;
                break;
            case 'price':
                decimals = CONFIG.FORMAT.DECIMALS_PRICE;
                prefix = '$';
                break;
            case 'percent':
                decimals = CONFIG.FORMAT.DECIMALS_PERCENT;
                suffix = '%';
                break;
            default:
                decimals = 2;
        }
        
        const absValue = Math.abs(value);
        const formatted = absValue.toLocaleString('en-US', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
        
        const sign = value < 0 ? '-' : '';
        return sign + prefix + formatted + suffix;
    },

    /**
     * Create a text node (safe from XSS)
     */
    text: function(content) {
        return document.createTextNode(content);
    },

    /**
     * Create an element with attributes
     */
    createElement: function(tag, attrs = {}, children = []) {
        const el = document.createElement(tag);

        for (const [key, value] of Object.entries(attrs)) {
            // Skip null/undefined values
            if (value === null || value === undefined) continue;

            if (key === 'className') {
                el.className = value;
            } else if (key === 'textContent') {
                el.textContent = value;
            } else if (key.startsWith('data')) {
                el.setAttribute(key.replace(/([A-Z])/g, '-$1').toLowerCase(), value);
            } else {
                el.setAttribute(key, value);
            }
        }
        
        for (const child of children) {
            if (typeof child === 'string') {
                el.appendChild(this.text(child));
            } else if (child) {
                el.appendChild(child);
            }
        }
        
        return el;
    },

    /**
     * Clear all children from an element
     */
    clearElement: function(element) {
        while (element.firstChild) {
            element.removeChild(element.firstChild);
        }
    },

    /**
     * Render summary cards
     */
    renderSummaryCards: function(summary, pyPeriodLabel, cyPeriodLabel) {
        const pyValueEl = document.getElementById('summary-py-value');
        const cyValueEl = document.getElementById('summary-cy-value');
        const changeValueEl = document.getElementById('summary-change-value');
        const pyPeriodEl = document.getElementById('summary-py-period');
        const cyPeriodEl = document.getElementById('summary-cy-period');
        const changePctEl = document.getElementById('summary-change-pct');

        // Multi-year mode vs two-period mode
        if (summary.years) {
            // Multi-year: show first and last year
            const years = Object.keys(summary.years).sort((a, b) => Number(a) - Number(b));
            const firstYear = years[0];
            const lastYear = years[years.length - 1];

            pyValueEl.textContent = this.formatNumber(summary.years[firstYear].value, 'currency');
            cyValueEl.textContent = this.formatNumber(summary.years[lastYear].value, 'currency');

            // Calculate total change from first to last year
            const totalChange = summary.years[lastYear].value - summary.years[firstYear].value;
            changeValueEl.textContent = this.formatNumber(totalChange, 'currency');

            pyPeriodEl.textContent = `FY ${firstYear}`;
            cyPeriodEl.textContent = `FY ${lastYear}`;

            const changePct = summary.years[firstYear].value !== 0 ?
                (totalChange / Math.abs(summary.years[firstYear].value)) * 100 : 0;
            const changePctFormatted = this.formatNumber(changePct, 'percent');
            changePctEl.textContent = changePct >= 0 ? '+' + changePctFormatted : changePctFormatted;
            changePctEl.className = 'card-pct ' + (changePct >= 0 ? 'positive' : 'negative');
        } else {
            // Two-period mode: original behavior
            pyValueEl.textContent = this.formatNumber(summary.py.value, 'currency');
            cyValueEl.textContent = this.formatNumber(summary.cy.value, 'currency');
            changeValueEl.textContent = this.formatNumber(summary.totalChange, 'currency');

            pyPeriodEl.textContent = pyPeriodLabel;
            cyPeriodEl.textContent = cyPeriodLabel;

            const changePct = this.formatNumber(summary.changePct, 'percent');
            changePctEl.textContent = summary.changePct >= 0 ? '+' + changePct : changePct;
            changePctEl.className = 'card-pct ' + (summary.changePct >= 0 ? 'positive' : 'negative');
        }
    },

    /**
     * Render bridge summary table
     */
    renderBridgeSummary: function(summary, mode, negatives) {
        const tbody = document.getElementById('bridge-summary-body');
        this.clearElement(tbody);

        let rows = [];

        if (summary.years) {
            // Multi-year mode: show all YoY bridges with year totals
            const years = Object.keys(summary.years).sort((a, b) => Number(a) - Number(b));
            const firstYear = years[0];
            const lastYear = years[years.length - 1];

            // Add year-by-year breakdown
            for (let i = 0; i < years.length; i++) {
                const year = years[i];
                const yearData = summary.years[year];

                // Starting year total
                rows.push({
                    label: `FY ${year}`,
                    value: yearData.value,
                    pct: null,
                    isTotal: true,
                    isStarting: i === 0
                });

                // If not the last year, show bridge to next year
                if (i < years.length - 1) {
                    const nextYear = years[i + 1];
                    const bridgeKey = `${year}-${nextYear}`;
                    const bridge = summary.bridges[bridgeKey];

                    if (bridge) {
                        // Calculate percentages based on starting year value
                        const startValue = yearData.value;
                        const pricePct = startValue !== 0 ? (bridge.priceImpact / Math.abs(startValue)) * 100 : 0;
                        const volPct = startValue !== 0 ? (bridge.volumeImpact / Math.abs(startValue)) * 100 : 0;
                        const mixPct = startValue !== 0 ? (bridge.mixImpact / Math.abs(startValue)) * 100 : 0;

                        rows.push({
                            label: `${year}→${nextYear} Price Impact`,
                            value: bridge.priceImpact,
                            pct: pricePct,
                            isTotal: false
                        });
                        rows.push({
                            label: `${year}→${nextYear} Volume Impact`,
                            value: bridge.volumeImpact,
                            pct: volPct,
                            isTotal: false
                        });
                        rows.push({
                            label: `${year}→${nextYear} Mix Impact`,
                            value: bridge.mixImpact,
                            pct: mixPct,
                            isTotal: false
                        });

                        if (mode === 'gm' && bridge.costImpact !== 0) {
                            const costPct = startValue !== 0 ? (bridge.costImpact / Math.abs(startValue)) * 100 : 0;
                            rows.push({
                                label: `${year}→${nextYear} Cost Impact`,
                                value: bridge.costImpact,
                                pct: costPct,
                                isTotal: false
                            });
                        }
                    }
                }
            }

            // Add negative values if present (multi-year)
            if (negatives && typeof negatives === 'object' && !negatives.py && !negatives.cy) {
                let negTotal = 0;
                for (const year of years) {
                    if (negatives[year]) {
                        negTotal += negatives[year].sales;
                    }
                }
                if (negTotal !== 0) {
                    rows.push({
                        label: 'Negative Values (excluded)',
                        value: negTotal,
                        pct: null,
                        isTotal: false,
                        isNegatives: true
                    });
                }
            }

        } else {
            // Two-period mode: original behavior
            rows = [
                {
                    label: 'Prior Year',
                    value: summary.py.value,
                    pct: null,
                    isTotal: false,
                    isStarting: true
                },
                {
                    label: 'Price Impact',
                    value: summary.priceImpact,
                    pct: summary.priceImpactPct,
                    isTotal: false
                },
                {
                    label: 'Volume Impact',
                    value: summary.volumeImpact,
                    pct: summary.volumeImpactPct,
                    isTotal: false
                },
                {
                    label: 'Mix Impact',
                    value: summary.mixImpact,
                    pct: summary.mixImpactPct,
                    isTotal: false
                }
            ];

            // Add cost impact for GM mode with sales-per-unit
            if (mode === 'gm' && summary.costImpact !== 0) {
                rows.push({
                    label: 'Cost Impact',
                    value: summary.costImpact,
                    pct: summary.costImpactPct,
                    isTotal: false
                });
            }

            // Add negative values if present
            if (negatives.py && negatives.cy) {
                const negTotal = (negatives.cy.sales - negatives.py.sales);
                if (negTotal !== 0) {
                    rows.push({
                        label: 'Negative Values (excluded)',
                        value: negTotal,
                        pct: null,
                        isTotal: false,
                        isNegatives: true
                    });
                }
            }

            // Add total row
            rows.push({
                label: 'Current Year / LTM',
                value: summary.cy.value,
                pct: null,
                isTotal: true
            });
        }
        
        for (const row of rows) {
            const tr = this.createElement('tr', {
                className: row.isTotal ? 'total-row' : ''
            });
            
            // Label cell
            const labelTd = this.createElement('td', {}, [row.label]);
            tr.appendChild(labelTd);
            
            // Value cell
            const valueTd = this.createElement('td', {
                className: 'num ' + (row.isStarting || row.isTotal ? '' : (row.value >= 0 ? 'positive' : 'negative'))
            });
            valueTd.textContent = this.formatNumber(row.value, 'currency');
            tr.appendChild(valueTd);
            
            // Percent cell
            const pctTd = this.createElement('td', { className: 'num' });
            if (row.pct !== null && row.pct !== undefined) {
                pctTd.textContent = this.formatNumber(row.pct, 'percent');
            } else {
                pctTd.textContent = '--';
            }
            tr.appendChild(pctTd);
            
            tbody.appendChild(tr);
        }
    },

    /**
     * Render detail table header
     */
    renderDetailTableHeader: function(dimensions, mode, pyLabel, cyLabel, isMultiYear = false, fiscalYears = []) {
        const thead = document.getElementById('detail-table-head');
        this.clearElement(thead);

        const tr = this.createElement('tr');

        // Dimension columns
        for (const dim of dimensions) {
            tr.appendChild(this.createElement('th', {}, [dim]));
        }

        if (isMultiYear && fiscalYears.length > 0) {
            // Multi-year mode: show all years with bridges between them
            for (let i = 0; i < fiscalYears.length; i++) {
                const fy = fiscalYears[i];
                const fyLabel = fy.label || `FY ${fy.fiscalYear}`;

                // Year metrics
                tr.appendChild(this.createElement('th', { className: 'num' }, [`${fyLabel} Sales`]));
                tr.appendChild(this.createElement('th', { className: 'num' }, [`${fyLabel} Vol`]));
                tr.appendChild(this.createElement('th', { className: 'num' }, [`${fyLabel} Price`]));

                // YoY bridge (except for the last year)
                if (i < fiscalYears.length - 1) {
                    const nextFY = fiscalYears[i + 1];
                    const bridgeLabel = `${fy.fiscalYear}→${nextFY.fiscalYear}`;

                    tr.appendChild(this.createElement('th', { className: 'num' }, [`${bridgeLabel} Price Δ`]));
                    tr.appendChild(this.createElement('th', { className: 'num' }, [`${bridgeLabel} Vol Δ`]));
                    tr.appendChild(this.createElement('th', { className: 'num' }, [`${bridgeLabel} Mix Δ`]));

                    if (mode === 'gm') {
                        tr.appendChild(this.createElement('th', { className: 'num' }, [`${bridgeLabel} Cost Δ`]));
                    }
                }
            }
        } else {
            // Two-period mode: original layout
            // Classification
            tr.appendChild(this.createElement('th', {}, ['Status']));

            // PY values by metric
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${pyLabel} Sales`]));
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${pyLabel} Volume`]));
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${pyLabel} Avg Price`]));

            // CY values by metric
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${cyLabel} Sales`]));
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${cyLabel} Volume`]));
            tr.appendChild(this.createElement('th', { className: 'num' }, [`${cyLabel} Avg Price`]));

            // Bridge components
            tr.appendChild(this.createElement('th', { className: 'num' }, ['Total Change']));
            tr.appendChild(this.createElement('th', { className: 'num' }, ['Price Impact']));
            tr.appendChild(this.createElement('th', { className: 'num' }, ['Volume Impact']));
            tr.appendChild(this.createElement('th', { className: 'num' }, ['Mix Impact']));

            if (mode === 'gm') {
                tr.appendChild(this.createElement('th', { className: 'num' }, ['Cost Impact']));
            }
        }

        thead.appendChild(tr);
    },

    /**
     * Render detail table body (paginated)
     */
    renderDetailTableBody: function(results, dimensions, mode, page = 1, pageSize = 50, isMultiYear = false, fiscalYears = []) {
        const tbody = document.getElementById('detail-table-body');
        this.clearElement(tbody);

        const start = (page - 1) * pageSize;
        const end = Math.min(start + pageSize, results.length);
        const pageResults = results.slice(start, end);

        for (const result of pageResults) {
            const tr = this.createElement('tr');

            // Dimension values
            for (const dim of dimensions) {
                tr.appendChild(this.createElement('td', {}, [result.dimensions[dim] || '--']));
            }

            if (isMultiYear && fiscalYears.length > 0) {
                // Multi-year mode: show all years and YoY bridges

                for (let i = 0; i < fiscalYears.length; i++) {
                    const fy = fiscalYears[i].fiscalYear;
                    const yearData = result.years[fy] || { sales: 0, volume: 0, price: 0 };

                    // Year metrics: Sales, Volume, Price
                    tr.appendChild(this.createElement('td', { className: 'num' }, [
                        this.formatNumber(yearData.sales, 'currency')
                    ]));
                    tr.appendChild(this.createElement('td', { className: 'num' }, [
                        this.formatNumber(yearData.volume, 'number')
                    ]));
                    tr.appendChild(this.createElement('td', { className: 'num' }, [
                        this.formatNumber(yearData.price, 'currency')
                    ]));

                    // YoY bridge columns (between this year and next)
                    if (i < fiscalYears.length - 1) {
                        const nextFY = fiscalYears[i + 1].fiscalYear;
                        const bridgeKey = `${fy}-${nextFY}`;
                        const bridge = result.bridges[bridgeKey] || {
                            priceImpact: 0,
                            volumeImpact: 0,
                            mixImpact: 0
                        };

                        tr.appendChild(this.createElement('td', {
                            className: 'num ' + (bridge.priceImpact >= 0 ? 'positive' : 'negative')
                        }, [this.formatNumber(bridge.priceImpact, 'currency')]));

                        tr.appendChild(this.createElement('td', {
                            className: 'num ' + (bridge.volumeImpact >= 0 ? 'positive' : 'negative')
                        }, [this.formatNumber(bridge.volumeImpact, 'currency')]));

                        tr.appendChild(this.createElement('td', {
                            className: 'num ' + (bridge.mixImpact >= 0 ? 'positive' : 'negative')
                        }, [this.formatNumber(bridge.mixImpact, 'currency')]));

                        if (mode === 'gm') {
                            tr.appendChild(this.createElement('td', {
                                className: 'num ' + (bridge.costImpact >= 0 ? 'positive' : 'negative')
                            }, [this.formatNumber(bridge.costImpact || 0, 'currency')]));
                        }
                    }
                }

            } else {
                // Two-period mode (original behavior)

                // Classification
                const statusClass = result.classification === 'new' ? 'positive' :
                                   result.classification === 'discontinued' ? 'negative' : '';
                tr.appendChild(this.createElement('td', { className: statusClass }, [
                    result.classification.charAt(0).toUpperCase() + result.classification.slice(1)
                ]));

                // PY values: Sales, Volume, Avg Price
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.py.sales, 'currency')
                ]));
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.py.volume, 'number')
                ]));
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.py.price, 'currency')
                ]));

                // CY values: Sales, Volume, Avg Price
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.cy.sales, 'currency')
                ]));
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.cy.volume, 'number')
                ]));
                tr.appendChild(this.createElement('td', { className: 'num' }, [
                    this.formatNumber(result.cy.price, 'currency')
                ]));

                // Bridge components
                const totalClass = 'num ' + (result.totalChange >= 0 ? 'positive' : 'negative');
                tr.appendChild(this.createElement('td', { className: totalClass }, [
                    this.formatNumber(result.totalChange, 'currency')
                ]));

                tr.appendChild(this.createElement('td', {
                    className: 'num ' + (result.priceImpact >= 0 ? 'positive' : 'negative')
                }, [this.formatNumber(result.priceImpact, 'currency')]));

                tr.appendChild(this.createElement('td', {
                    className: 'num ' + (result.volumeImpact >= 0 ? 'positive' : 'negative')
                }, [this.formatNumber(result.volumeImpact, 'currency')]));

                tr.appendChild(this.createElement('td', {
                    className: 'num ' + (result.mixImpact >= 0 ? 'positive' : 'negative')
                }, [this.formatNumber(result.mixImpact, 'currency')]));

                if (mode === 'gm') {
                    tr.appendChild(this.createElement('td', {
                        className: 'num ' + (result.costImpact >= 0 ? 'positive' : 'negative')
                    }, [this.formatNumber(result.costImpact, 'currency')]));
                }
            }

            tbody.appendChild(tr);
        }

        // Render pagination
        this.renderPagination(results.length, page, pageSize);
    },

    /**
     * Render pagination controls
     */
    renderPagination: function(totalItems, currentPage, pageSize) {
        const container = document.getElementById('detail-pagination');
        this.clearElement(container);
        
        const totalPages = Math.ceil(totalItems / pageSize);
        
        if (totalPages <= 1) return;
        
        // Previous button
        const prevBtn = this.createElement('button', {
            className: 'btn-secondary',
            disabled: currentPage === 1 ? 'disabled' : null
        }, ['← Previous']);
        prevBtn.dataset.page = currentPage - 1;
        container.appendChild(prevBtn);
        
        // Page info
        const pageInfo = this.createElement('span', { className: 'page-info' }, [
            `Page ${currentPage} of ${totalPages} (${totalItems} items)`
        ]);
        container.appendChild(pageInfo);
        
        // Next button
        const nextBtn = this.createElement('button', {
            className: 'btn-secondary',
            disabled: currentPage === totalPages ? 'disabled' : null
        }, ['Next →']);
        nextBtn.dataset.page = currentPage + 1;
        container.appendChild(nextBtn);
    },

    /**
     * Render negatives table
     */
    renderNegativesTable: function(negatives) {
        const tbody = document.getElementById('negatives-table-body');
        this.clearElement(tbody);

        const summary = document.getElementById('negatives-summary');
        this.clearElement(summary);

        // Check if multi-year mode or two-period mode
        const isMultiYear = !negatives.py && !negatives.cy;

        let totalNegCount = 0;
        let totalNegSales = 0;
        let rows = [];

        if (isMultiYear) {
            // Multi-year mode: negatives are keyed by year
            for (const [year, neg] of Object.entries(negatives)) {
                totalNegCount += neg.count;
                totalNegSales += neg.sales;
                if (neg.count > 0) {
                    rows.push({ period: `FY ${year}`, type: 'Negative/Zero Sales or Qty', ...neg });
                }
            }
        } else {
            // Two-period mode: negatives.py and negatives.cy
            totalNegCount = negatives.py.count + negatives.cy.count;
            totalNegSales = negatives.py.sales + negatives.cy.sales;
            rows = [
                { period: 'Prior Year', type: 'Negative/Zero Sales or Qty', ...negatives.py },
                { period: 'Current Year / LTM', type: 'Negative/Zero Sales or Qty', ...negatives.cy }
            ];
        }

        // Summary cards
        summary.appendChild(this.createElement('div', { className: 'summary-card' }, [
            this.createElement('span', { className: 'card-label' }, ['Total Excluded Rows']),
            this.createElement('span', { className: 'card-value' }, [totalNegCount.toLocaleString()])
        ]));

        summary.appendChild(this.createElement('div', { className: 'summary-card' }, [
            this.createElement('span', { className: 'card-label' }, ['Total Excluded Sales']),
            this.createElement('span', { className: 'card-value' }, [this.formatNumber(totalNegSales, 'currency')])
        ]));

        // Table rows
        for (const row of rows) {
            if (row.count === 0) continue;

            const tr = this.createElement('tr');
            tr.appendChild(this.createElement('td', {}, [row.period]));
            tr.appendChild(this.createElement('td', {}, [row.type]));
            tr.appendChild(this.createElement('td', { className: 'num' }, [row.count.toLocaleString()]));
            tr.appendChild(this.createElement('td', { className: 'num' }, [this.formatNumber(row.sales, 'currency')]));
            tr.appendChild(this.createElement('td', { className: 'num' }, [this.formatNumber(row.quantity, 'quantity')]));
            tbody.appendChild(tr);
        }

        if (tbody.children.length === 0) {
            const tr = this.createElement('tr');
            tr.appendChild(this.createElement('td', { colspan: '5', className: 'text-muted' }, [
                'No negative or zero values found'
            ]));
            tbody.appendChild(tr);
        }
    },

    /**
     * Render assumptions tab
     */
    renderAssumptions: function(config, stats, methodology) {
        const configBody = document.getElementById('assumptions-config');
        this.clearElement(configBody);

        const configRows = [
            ['Analysis Mode', config.mode === 'gm' ? 'Gross Margin Bridge' : 'Sales PVM Bridge'],
            ['Fiscal Year End', CONFIG.MONTH_NAMES[config.fyEndMonth - 1]]
        ];

        // Add period information based on mode
        if (config.fiscalYears && config.fiscalYears.length > 0) {
            // Multi-year mode
            const yearLabels = config.fiscalYears.map(fy => fy.label || `FY ${fy.fiscalYear}`).join(', ');
            const firstYear = config.fiscalYears[0];
            const lastYear = config.fiscalYears[config.fiscalYears.length - 1];
            configRows.push(['Analysis Type', 'Multi-Year Comparison']);
            configRows.push(['Fiscal Years', yearLabels]);
            configRows.push(['Date Range', `${PeriodUtils.formatDate(firstYear.start)} to ${PeriodUtils.formatDate(lastYear.end)}`]);
        } else if (config.pyRange && config.cyRange) {
            // Two-period mode
            configRows.push(['Prior Year Period', PeriodUtils.formatDateRange(config.pyRange.start, config.pyRange.end)]);
            configRows.push(['LTM Period', PeriodUtils.formatDateRange(config.cyRange.start, config.cyRange.end)]);
        }

        configRows.push(
            ['Level of Detail', config.dimensions.length > 0 ? config.dimensions.join(', ') : 'Total Only'],
            ['Total Rows Processed', stats.totalRows.toLocaleString()],
            ['Rows Included in Analysis', stats.includedRows.toLocaleString()],
            ['Rows Excluded', stats.excludedRows.toLocaleString()],
            ['Unique LOD Combinations', stats.uniqueLODKeys.toLocaleString()]
        );
        
        if (config.mode === 'gm') {
            configRows.splice(1, 0, ['GM Price Definition', 
                config.gmPriceDefinition === 'margin-per-unit' ? 
                    'Margin per Unit: (Sales − Cost) / Qty' : 
                    'Sales per Unit: Sales / Qty']);
        }
        
        for (const [label, value] of configRows) {
            const tr = this.createElement('tr');
            tr.appendChild(this.createElement('th', {}, [label]));
            tr.appendChild(this.createElement('td', {}, [value]));
            configBody.appendChild(tr);
        }
        
        // Methodology
        const methodologyEl = document.getElementById('methodology-content');
        this.clearElement(methodologyEl);
        
        methodologyEl.appendChild(this.createElement('p', {}, [methodology.description]));
        
        const formulaList = this.createElement('ul');
        for (const f of methodology.formulas) {
            const li = this.createElement('li');
            li.appendChild(this.createElement('strong', {}, [f.name + ': ']));
            li.appendChild(this.createElement('code', {}, [f.formula]));
            formulaList.appendChild(li);
        }
        methodologyEl.appendChild(formulaList);
        
        // Notes
        const dataQualityEl = document.getElementById('data-quality-content');
        this.clearElement(dataQualityEl);
        
        const notesList = this.createElement('ul');
        for (const note of methodology.notes) {
            notesList.appendChild(this.createElement('li', {}, [note]));
        }
        dataQualityEl.appendChild(notesList);
        
        // Add stats notes
        if (stats.parseErrors > 0) {
            dataQualityEl.appendChild(this.createElement('p', { className: 'text-muted' }, [
                `Note: ${stats.parseErrors.toLocaleString()} rows had parse errors and were excluded.`
            ]));
        }
        
        if (stats.outsidePeriodRows > 0) {
            dataQualityEl.appendChild(this.createElement('p', { className: 'text-muted' }, [
                `Note: ${stats.outsidePeriodRows.toLocaleString()} rows were outside both analysis periods.`
            ]));
        }
    },

    /**
     * Show error modal
     */
    showError: function(message) {
        const modal = document.getElementById('error-modal');
        const body = document.getElementById('error-modal-body');
        body.textContent = message;
        modal.classList.remove('hidden');
    },

    /**
     * Hide error modal
     */
    hideError: function() {
        const modal = document.getElementById('error-modal');
        modal.classList.add('hidden');
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = UIRenderer;
}
