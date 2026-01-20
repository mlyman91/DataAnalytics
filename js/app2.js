/**
 * ============================================
 * PVM Bridge Tool - Application Part 2
 * Processing, Results, and Export
 * ============================================
 */

// Extend App object with additional methods
Object.assign(App, {
    async goToPeriods() {
        this.updateDimensionMappings();
        this.showScreen('screen-periods');
        document.getElementById('fy-end-month').value = this.state.fyEndMonth;
        document.getElementById('date-range-info').innerHTML = '<p>Scanning file for date range...</p>';
        await this.scanDateRange();
    },

    async scanDateRange() {
        const dateColumn = this.state.columnMappings.date;
        const dateFormat = this.state.dateFormat;
        let minDate = null, maxDate = null, rowCount = 0;

        // Use appropriate parser based on file type
        const parser = this.state.isExcelFile ? ExcelParser : CSVParser;
        await parser.parseFile(this.state.file, {
            onRow: (row) => {
                rowCount++;
                const date = PeriodUtils.parseDate(row[dateColumn], dateFormat);
                if (date) {
                    if (!minDate || date < minDate) minDate = new Date(date);
                    if (!maxDate || date > maxDate) maxDate = new Date(date);
                }
            },
            onComplete: () => {
                this.state.dataDateRange = { min: minDate, max: maxDate };
                const info = document.getElementById('date-range-info');
                if (minDate && maxDate) {
                    // Detect fiscal years in the data
                    const fiscalYears = PeriodUtils.detectFiscalYears(minDate, maxDate, this.state.fyEndMonth);
                    this.state.detectedFiscalYears = fiscalYears;

                    // Check if multi-year mode should be enabled
                    const fullYears = fiscalYears.filter(y => y.fullyCovered);
                    this.state.hasMultipleYears = fullYears.length >= 2;

                    let dateRangeHTML = `<p><strong>Date Range:</strong> ${PeriodUtils.formatDate(minDate)} to ${PeriodUtils.formatDate(maxDate)}</p>
                        <p><strong>Total Rows:</strong> ${rowCount.toLocaleString()}</p>`;

                    // Show detected fiscal years
                    if (fiscalYears.length >= 2) {
                        const yearLabels = fiscalYears.map(y => y.label).join(', ');
                        const completeCount = fullYears.length;
                        const partialCount = fiscalYears.length - completeCount;

                        dateRangeHTML += `<p><strong>Detected Fiscal Years:</strong> ${yearLabels}</p>`;
                        if (completeCount >= 2) {
                            dateRangeHTML += `<p style="color: var(--color-success); font-weight: 600;">✓ Multi-year analysis available (${completeCount} complete, ${partialCount} partial)</p>`;
                        } else if (completeCount === 1) {
                            dateRangeHTML += `<p style="color: var(--color-info);">⚠ ${completeCount} complete year, ${partialCount} partial. You can still analyze partial years.</p>`;
                        } else {
                            dateRangeHTML += `<p style="color: var(--color-warning);">⚠ All years have partial data. Results may be incomplete.</p>`;
                        }

                        // Show multi-year section with ALL fiscal years (including partial)
                        this.showMultiYearSelection(fiscalYears);
                    } else if (fiscalYears.length === 1) {
                        dateRangeHTML += `<p><strong>Detected:</strong> ${fiscalYears[0].label}</p>`;

                        // Show two-period config and hide multi-year section
                        document.getElementById('multi-year-section').style.display = 'none';
                        document.getElementById('two-period-config').style.display = 'block';
                    } else {
                        // No fiscal years detected, show two-period config
                        document.getElementById('multi-year-section').style.display = 'none';
                        document.getElementById('two-period-config').style.display = 'block';
                    }

                    info.innerHTML = dateRangeHTML;

                    // Set default LTM end date
                    this.state.ltmEndDate = maxDate;
                    document.getElementById('ltm-end-date').value = PeriodUtils.toISODateString(maxDate);

                    // Set default current fiscal year based on max date
                    const maxYear = maxDate.getFullYear();
                    const maxMonth = maxDate.getMonth() + 1;
                    const fyEndMonth = this.state.fyEndMonth;
                    // If we're past the FY end month, we're in the next fiscal year
                    const currentFY = maxMonth > fyEndMonth ? maxYear + 1 : maxYear;
                    this.state.cyFiscalYear = currentFY;
                    document.getElementById('cy-fiscal-year').value = currentFY;

                    this.updatePeriodPreviews();
                } else {
                    info.innerHTML = '<p>Could not determine date range.</p>';
                }
            },
            onError: (error) => {
                document.getElementById('date-range-info').innerHTML = `<p>Error: ${error.message}</p>`;
            }
        });
    },

    showMultiYearSelection(fiscalYears) {
        // Show multi-year section, hide two-period config
        document.getElementById('multi-year-section').style.display = 'block';
        document.getElementById('two-period-config').style.display = 'none';

        // Populate fiscal year checkboxes
        const container = document.getElementById('fiscal-year-checkboxes');
        UIRenderer.clearElement(container);

        // Initialize selected years if not already set (default to all years)
        if (!this.state.selectedFiscalYears || this.state.selectedFiscalYears.length === 0) {
            this.state.selectedFiscalYears = fiscalYears.map(fy => fy.fiscalYear);
        }

        for (const fyConfig of fiscalYears) {
            const isChecked = this.state.selectedFiscalYears.includes(fyConfig.fiscalYear);

            const label = UIRenderer.createElement('label', { className: 'fiscal-year-option' });

            const checkbox = UIRenderer.createElement('input', {
                type: 'checkbox',
                value: fyConfig.fiscalYear,
                checked: isChecked ? 'checked' : null,
                'data-fy-start': PeriodUtils.toISODateString(fyConfig.start),
                'data-fy-end': PeriodUtils.toISODateString(fyConfig.end)
            });

            checkbox.addEventListener('change', () => this.updateFiscalYearSelection());

            const textContainer = UIRenderer.createElement('div');

            // Add warning icon if not fully covered
            const labelText = fyConfig.fullyCovered ? fyConfig.label : `${fyConfig.label} ⚠`;
            const fyLabel = UIRenderer.createElement('div', { className: 'fy-label' }, [labelText]);

            const dateRangeText = fyConfig.fullyCovered ?
                `${PeriodUtils.formatDate(fyConfig.start)} - ${PeriodUtils.formatDate(fyConfig.end)}` :
                `${PeriodUtils.formatDate(fyConfig.start)} - ${PeriodUtils.formatDate(fyConfig.end)} (partial)`;

            const dateRange = UIRenderer.createElement('div', { className: 'fy-date-range' }, [dateRangeText]);

            textContainer.appendChild(fyLabel);
            textContainer.appendChild(dateRange);
            label.appendChild(checkbox);
            label.appendChild(textContainer);
            container.appendChild(label);
        }

        this.updateFiscalYearSelection();
    },

    updateFiscalYearSelection() {
        const checkboxes = document.querySelectorAll('#fiscal-year-checkboxes input[type="checkbox"]');
        this.state.selectedFiscalYears = [];

        checkboxes.forEach(cb => {
            if (cb.checked) {
                this.state.selectedFiscalYears.push(Number(cb.value));
            }
        });

        // Sort selected years
        this.state.selectedFiscalYears.sort((a, b) => a - b);

        // Update summary and validation
        const summary = document.getElementById('period-summary');
        const btnContinue = document.getElementById('btn-to-lod');

        if (this.state.selectedFiscalYears.length < 2) {
            summary.innerHTML = '<p style="color: var(--color-warning);">⚠️ Please select at least 2 fiscal years for multi-year analysis.</p>';
            btnContinue.disabled = true;
        } else {
            const selectedYearObjs = this.state.detectedFiscalYears.filter(fy =>
                this.state.selectedFiscalYears.includes(fy.fiscalYear)
            );

            const yearLabels = selectedYearObjs.map(fy => fy.label).join(', ');
            const firstYear = selectedYearObjs[0];
            const lastYear = selectedYearObjs[selectedYearObjs.length - 1];

            summary.innerHTML = `<p><strong>Multi-Year Analysis:</strong></p>
                <p>Selected Years: ${yearLabels}</p>
                <p>Period: ${PeriodUtils.formatDate(firstYear.start)} to ${PeriodUtils.formatDate(lastYear.end)}</p>
                <p style="color: var(--color-success);">✓ ${this.state.selectedFiscalYears.length} fiscal years selected</p>`;

            btnContinue.disabled = false;

            // Update state to use multi-year mode
            this.state.useMultiYearMode = true;
        }
    },

    updatePeriodPreviews() {
        const { fyEndMonth, ltmEndDate, useLTM, cyFiscalYear } = this.state;

        // Skip if we're in multi-year mode
        if (this.state.useMultiYearMode) {
            return;
        }

        let cyRange, pyRange, currentFY, priorFY;

        if (useLTM) {
            // LTM mode
            if (!ltmEndDate) return;

            cyRange = PeriodUtils.getLTMRange(ltmEndDate);
            this.state.cyRange = cyRange;

            priorFY = PeriodUtils.getPriorFiscalYear(ltmEndDate, fyEndMonth);
            pyRange = PeriodUtils.getFiscalYearRange(priorFY, fyEndMonth);
            this.state.pyRange = pyRange;

            // Store fiscal year labels
            this.state.pyFiscalYear = priorFY;
            this.state.pyLabel = `FY ${priorFY}`;
            this.state.cyLabel = 'LTM';

            document.getElementById('ltm-range-preview').textContent = PeriodUtils.formatDateRange(cyRange.start, cyRange.end);
            document.getElementById('pfy-range').textContent = `FY ${priorFY}: ${PeriodUtils.formatDateRange(pyRange.start, pyRange.end)}`;

            const summary = document.getElementById('period-summary');
            summary.innerHTML = `<p><strong>Comparison:</strong></p>
                <p>Prior Year (FY ${priorFY}): ${PeriodUtils.formatDateRange(pyRange.start, pyRange.end)}</p>
                <p>vs. LTM: ${PeriodUtils.formatDateRange(cyRange.start, cyRange.end)}</p>`;

            const validation = PeriodUtils.validatePeriodConfig({ fyEndMonth, ltmEndDate, pyRange, cyRange });
            if (validation.warnings.length > 0) {
                summary.innerHTML += `<p style="color: var(--color-warning);">⚠️ ${validation.warnings.join(' ')}</p>`;
            }
            document.getElementById('btn-to-lod').disabled = !validation.valid;

        } else {
            // Fiscal Year mode
            if (!cyFiscalYear) return;

            currentFY = cyFiscalYear;
            priorFY = currentFY - 1;

            cyRange = PeriodUtils.getFiscalYearRange(currentFY, fyEndMonth);
            this.state.cyRange = cyRange;

            pyRange = PeriodUtils.getFiscalYearRange(priorFY, fyEndMonth);
            this.state.pyRange = pyRange;

            // Store fiscal year labels
            this.state.pyFiscalYear = priorFY;
            this.state.cyFiscalYear = currentFY;
            this.state.pyLabel = `FY ${priorFY}`;
            this.state.cyLabel = `FY ${currentFY}`;

            document.getElementById('cy-range-preview').textContent = `FY ${currentFY}: ${PeriodUtils.formatDateRange(cyRange.start, cyRange.end)}`;
            document.getElementById('pfy-range').textContent = `FY ${priorFY}: ${PeriodUtils.formatDateRange(pyRange.start, pyRange.end)}`;

            const summary = document.getElementById('period-summary');
            summary.innerHTML = `<p><strong>Comparison:</strong></p>
                <p>Prior Year (FY ${priorFY}): ${PeriodUtils.formatDateRange(pyRange.start, pyRange.end)}</p>
                <p>vs. Current Year (FY ${currentFY}): ${PeriodUtils.formatDateRange(cyRange.start, cyRange.end)}</p>`;

            // Validate (no warnings for fiscal year mode typically)
            document.getElementById('btn-to-lod').disabled = false;
        }
    },

    goToLOD() {
        this.showScreen('screen-lod');
        const container = document.getElementById('lod-selection');
        UIRenderer.clearElement(container);

        const dimensions = this.state.columnMappings.dimensions;
        if (dimensions.length === 0) {
            container.innerHTML = '<p class="text-muted">No dimensions available. Analysis at total level.</p>';
        } else {
            for (const dim of dimensions) {
                const label = UIRenderer.createElement('label', { className: 'checkbox-label' });
                const checkbox = UIRenderer.createElement('input', {
                    type: 'checkbox', value: dim,
                    checked: this.state.selectedDimensions.includes(dim) ? 'checked' : null
                });
                checkbox.addEventListener('change', () => this.updateLODSelection());
                label.appendChild(checkbox);
                label.appendChild(UIRenderer.text(dim));
                container.appendChild(label);
            }
        }
        this.updateLODSelection();
    },

    updateLODSelection() {
        const checkboxes = document.querySelectorAll('#lod-selection input[type="checkbox"]');
        this.state.selectedDimensions = [];
        checkboxes.forEach(cb => { if (cb.checked) this.state.selectedDimensions.push(cb.value); });

        const preview = document.getElementById('lod-preview-text');
        const warning = document.getElementById('lod-warning');

        if (this.state.selectedDimensions.length === 0) {
            preview.textContent = 'Analysis will be at total level (no dimensions).';
            warning.classList.add('hidden');
        } else {
            preview.textContent = `Analysis by: ${this.state.selectedDimensions.join(' → ')}`;
            warning.classList.toggle('hidden', this.state.selectedDimensions.length <= CONFIG.VALIDATION.MAX_RECOMMENDED_DIMENSIONS);
        }
    },

    async startProcessing() {
        this.state.isProcessing = true;
        this.state.shouldCancel = false;
        this.showScreen('screen-processing');

        // Reset progress UI
        document.getElementById('progress-fill').style.width = '0%';
        document.getElementById('progress-percent').textContent = '0%';
        document.getElementById('progress-rows').textContent = '0 rows';
        document.querySelectorAll('.stage').forEach(s => {
            s.classList.remove('active', 'complete');
            s.querySelector('.stage-icon').textContent = '⏳';
        });

        const setStage = (id, status) => {
            const stage = document.getElementById(id);
            stage.classList.remove('active', 'complete');
            if (status === 'active') stage.classList.add('active');
            if (status === 'complete') {
                stage.classList.add('complete');
                stage.querySelector('.stage-icon').textContent = '✓';
            }
        };

        try {
            // Stage 1: Parsing & Aggregating
            setStage('stage-parsing', 'active');

            // Determine if multi-year mode should be used
            const useMultiYear = this.state.useMultiYearMode && this.state.selectedFiscalYears && this.state.selectedFiscalYears.length >= 2;
            let selectedFiscalYearConfigs = [];

            if (useMultiYear) {
                // Get the fiscal year configs for selected years only
                selectedFiscalYearConfigs = this.state.detectedFiscalYears.filter(fy =>
                    this.state.selectedFiscalYears.includes(fy.fiscalYear)
                );
            }

            const aggregatorConfig = {
                dimensions: this.state.selectedDimensions,
                dateColumn: this.state.columnMappings.date,
                salesColumn: this.state.columnMappings.sales,
                quantityColumn: this.state.columnMappings.quantity,
                costColumn: this.state.columnMappings.cost,
                dateFormat: this.state.dateFormat,
                fyEndMonth: this.state.fyEndMonth
            };

            if (useMultiYear) {
                // Multi-year mode: pass selected fiscal years
                aggregatorConfig.fiscalYears = selectedFiscalYearConfigs;
            } else {
                // Two-period mode: pass pyRange and cyRange
                aggregatorConfig.pyRange = this.state.pyRange;
                aggregatorConfig.cyRange = this.state.cyRange;
            }

            const ctx = Aggregator.createContext(aggregatorConfig);

            // Use appropriate parser based on file type
            const parser = this.state.isExcelFile ? ExcelParser : CSVParser;
            await parser.parseFile(this.state.file, {
                onRow: (row, rowNum) => {
                    Aggregator.processRow(ctx, row);
                },
                onProgress: (bytes, total, rows) => {
                    const pct = Math.round((bytes / total) * 100);
                    document.getElementById('progress-fill').style.width = pct + '%';
                    document.getElementById('progress-percent').textContent = pct + '%';
                    document.getElementById('progress-rows').textContent = rows.toLocaleString() + ' rows';
                },
                shouldCancel: () => this.state.shouldCancel,
                onComplete: (result) => {
                    if (result.cancelled) {
                        this.showScreen('screen-lod');
                        return;
                    }
                    setStage('stage-parsing', 'complete');
                    setStage('stage-aggregating', 'complete');
                },
                onError: (error) => {
                    UIRenderer.showError('Processing error: ' + error.message);
                    this.showScreen('screen-lod');
                }
            });

            if (this.state.shouldCancel) return;

            // Stage 2: Calculate Bridge
            setStage('stage-calculating', 'active');
            
            const aggregationResults = Aggregator.finalize(ctx);
            this.state.aggregationResults = aggregationResults;

            const bridgeCalcOptions = {
                mode: this.state.mode,
                gmPriceDefinition: this.state.gmPriceDefinition,
                isMultiYear: aggregationResults.isMultiYear
            };

            if (aggregationResults.isMultiYear) {
                bridgeCalcOptions.fiscalYears = selectedFiscalYearConfigs;
            }

            const bridgeResults = BridgeCalculator.calculate(aggregationResults.data, bridgeCalcOptions);
            this.state.bridgeResults = bridgeResults;

            setStage('stage-calculating', 'complete');

            // Stage 3: Finalize
            setStage('stage-finalizing', 'active');
            
            // Small delay for UI
            await new Promise(r => setTimeout(r, 200));
            
            setStage('stage-finalizing', 'complete');

            // Show results
            this.showResults();

        } catch (error) {
            UIRenderer.showError('Error: ' + error.message);
            this.showScreen('screen-lod');
        } finally {
            this.state.isProcessing = false;
        }
    },

    cancelProcessing() {
        this.state.shouldCancel = true;
    },

    showResults() {
        this.showScreen('screen-results');

        const { bridgeResults, aggregationResults, pyRange, cyRange, mode, gmPriceDefinition, selectedDimensions } = this.state;
        const summary = bridgeResults.summary;

        // Check if multi-year mode
        const isMultiYear = aggregationResults.isMultiYear;
        const fiscalYears = isMultiYear && this.state.selectedFiscalYears ?
            this.state.detectedFiscalYears.filter(fy => this.state.selectedFiscalYears.includes(fy.fiscalYear)) :
            [];

        // Period labels
        let pyLabel, cyLabel;
        if (isMultiYear) {
            // In multi-year mode, labels are handled differently (first year, last year)
            pyLabel = '';
            cyLabel = '';
        } else {
            pyLabel = pyRange ? PeriodUtils.formatDateRange(pyRange.start, pyRange.end) : '';
            cyLabel = cyRange ? PeriodUtils.formatDateRange(cyRange.start, cyRange.end) : '';
        }

        // Summary cards
        UIRenderer.renderSummaryCards(summary, pyLabel, cyLabel);

        // Bridge summary table
        UIRenderer.renderBridgeSummary(summary, mode, aggregationResults.negatives);

        // Detail table
        UIRenderer.renderDetailTableHeader(selectedDimensions, mode, this.state.pyLabel, this.state.cyLabel, isMultiYear, fiscalYears);
        this.renderDetailTable();

        // Negatives table
        UIRenderer.renderNegativesTable(aggregationResults.negatives);

        // Assumptions
        const methodology = BridgeCalculator.getMethodologyDescription(mode, gmPriceDefinition);
        const assumptionsConfig = {
            mode,
            gmPriceDefinition,
            fyEndMonth: this.state.fyEndMonth,
            pyRange,
            cyRange,
            dimensions: selectedDimensions,
            dateColumn: this.state.columnMappings.date,
            salesColumn: this.state.columnMappings.sales,
            quantityColumn: this.state.columnMappings.quantity,
            costColumn: this.state.columnMappings.cost
        };

        // Add fiscal years if in multi-year mode
        if (isMultiYear) {
            assumptionsConfig.fiscalYears = fiscalYears;
        }

        UIRenderer.renderAssumptions(assumptionsConfig, aggregationResults.stats, methodology);

        this.switchTab('summary');
    },

    renderDetailTable() {
        const { bridgeResults, aggregationResults, selectedDimensions, mode, sortBy, searchTerm, currentPage } = this.state;

        let results = bridgeResults.detail;
        results = BridgeCalculator.filterResults(results, searchTerm);
        results = BridgeCalculator.sortResults(results, sortBy);

        // Check if multi-year mode
        const isMultiYear = aggregationResults.isMultiYear;
        const fiscalYears = isMultiYear ? this.state.detectedFiscalYears.filter(y => y.fullyCovered) : [];

        UIRenderer.renderDetailTableBody(results, selectedDimensions, mode, currentPage, 50, isMultiYear, fiscalYears);
    },

    switchTab(tabId) {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tabId));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === 'tab-' + tabId));
    },

    exportExcel() {
        const { aggregationResults } = this.state;
        const isMultiYear = aggregationResults.isMultiYear;

        const config = {
            mode: this.state.mode,
            gmPriceDefinition: this.state.gmPriceDefinition,
            fyEndMonth: this.state.fyEndMonth,
            pyRange: this.state.pyRange,
            cyRange: this.state.cyRange,
            pyLabel: this.state.pyLabel,
            cyLabel: this.state.cyLabel,
            dimensions: this.state.selectedDimensions,
            dateColumn: this.state.columnMappings.date,
            salesColumn: this.state.columnMappings.sales,
            quantityColumn: this.state.columnMappings.quantity,
            costColumn: this.state.columnMappings.cost
        };

        // Add fiscal years if in multi-year mode
        if (isMultiYear && this.state.selectedFiscalYears) {
            config.fiscalYears = this.state.detectedFiscalYears.filter(fy =>
                this.state.selectedFiscalYears.includes(fy.fiscalYear)
            );
        }

        ExcelExport.exportToExcel(this.state.bridgeResults, this.state.aggregationResults, config);
    }
});
