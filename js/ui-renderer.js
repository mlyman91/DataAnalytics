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
        
        pyValueEl.textContent = this.formatNumber(summary.py.value, 'currency');
        cyValueEl.textContent = this.formatNumber(summary.cy.value, 'currency');
        changeValueEl.textContent = this.formatNumber(summary.totalChange, 'currency');
        
        pyPeriodEl.textContent = pyPeriodLabel;
        cyPeriodEl.textContent = cyPeriodLabel;
        
        const changePct = this.formatNumber(summary.changePct, 'percent');
        changePctEl.textContent = summary.changePct >= 0 ? '+' + changePct : changePct;
        changePctEl.className = 'card-pct ' + (summary.changePct >= 0 ? 'positive' : 'negative');
    },

    /**
     * Render bridge summary table
     */
    renderBridgeSummary: function(summary, mode, negatives) {
        const tbody = document.getElementById('bridge-summary-body');
        this.clearElement(tbody);
        
        const rows = [
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
        
        // Add total row
        rows.push({
            label: 'Current Year / LTM',
            value: summary.cy.value,
            pct: null,
            isTotal: true
        });
        
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
    renderDetailTableHeader: function(dimensions, mode) {
        const thead = document.getElementById('detail-table-head');
        this.clearElement(thead);
        
        const tr = this.createElement('tr');
        
        // Dimension columns
        for (const dim of dimensions) {
            tr.appendChild(this.createElement('th', {}, [dim]));
        }
        
        // Classification
        tr.appendChild(this.createElement('th', {}, ['Status']));
        
        // PY/CY values
        tr.appendChild(this.createElement('th', { className: 'num' }, ['PY Value']));
        tr.appendChild(this.createElement('th', { className: 'num' }, ['CY Value']));
        
        // Bridge components
        tr.appendChild(this.createElement('th', { className: 'num' }, ['Total Change']));
        tr.appendChild(this.createElement('th', { className: 'num' }, ['Price Impact']));
        tr.appendChild(this.createElement('th', { className: 'num' }, ['Volume Impact']));
        tr.appendChild(this.createElement('th', { className: 'num' }, ['Mix Impact']));
        
        if (mode === 'gm') {
            tr.appendChild(this.createElement('th', { className: 'num' }, ['Cost Impact']));
        }
        
        thead.appendChild(tr);
    },

    /**
     * Render detail table body (paginated)
     */
    renderDetailTableBody: function(results, dimensions, mode, page = 1, pageSize = 50) {
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
            
            // Classification
            const statusClass = result.classification === 'new' ? 'positive' : 
                               result.classification === 'discontinued' ? 'negative' : '';
            tr.appendChild(this.createElement('td', { className: statusClass }, [
                result.classification.charAt(0).toUpperCase() + result.classification.slice(1)
            ]));
            
            // PY/CY values
            tr.appendChild(this.createElement('td', { className: 'num' }, [
                this.formatNumber(result.py.value, 'currency')
            ]));
            tr.appendChild(this.createElement('td', { className: 'num' }, [
                this.formatNumber(result.cy.value, 'currency')
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
        
        // Summary cards
        const totalNegCount = negatives.py.count + negatives.cy.count;
        const totalNegSales = negatives.py.sales + negatives.cy.sales;
        
        summary.appendChild(this.createElement('div', { className: 'summary-card' }, [
            this.createElement('span', { className: 'card-label' }, ['Total Excluded Rows']),
            this.createElement('span', { className: 'card-value' }, [totalNegCount.toLocaleString()])
        ]));
        
        summary.appendChild(this.createElement('div', { className: 'summary-card' }, [
            this.createElement('span', { className: 'card-label' }, ['Total Excluded Sales']),
            this.createElement('span', { className: 'card-value' }, [this.formatNumber(totalNegSales, 'currency')])
        ]));
        
        // Table rows
        const rows = [
            { period: 'Prior Year', type: 'Negative/Zero Sales or Qty', ...negatives.py },
            { period: 'Current Year / LTM', type: 'Negative/Zero Sales or Qty', ...negatives.cy }
        ];
        
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
            ['Fiscal Year End', CONFIG.MONTH_NAMES[config.fyEndMonth - 1]],
            ['Prior Year Period', PeriodUtils.formatDateRange(config.pyRange.start, config.pyRange.end)],
            ['LTM Period', PeriodUtils.formatDateRange(config.cyRange.start, config.cyRange.end)],
            ['Level of Detail', config.dimensions.length > 0 ? config.dimensions.join(', ') : 'Total Only'],
            ['Total Rows Processed', stats.totalRows.toLocaleString()],
            ['Rows Included in Analysis', stats.includedRows.toLocaleString()],
            ['Rows Excluded', stats.excludedRows.toLocaleString()],
            ['Unique LOD Combinations', stats.uniqueLODKeys.toLocaleString()]
        ];
        
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
