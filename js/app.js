/**
 * ============================================
 * PVM Bridge Tool - Main Application
 * ============================================
 * Orchestrates UI flow, state, and modules.
 * ============================================
 */

const App = {
    state: {
        mode: 'pvm',
        file: null,
        isExcelFile: false,
        headers: [],
        sampleRows: [],
        columnMappings: { date: null, sales: null, quantity: null, cost: null, dimensions: [] },
        dateFormat: null,
        detectedDateFormat: null,
        fyEndMonth: 12,
        useLTM: true,
        ltmEndDate: null,
        pyFiscalYear: null,
        cyFiscalYear: null,
        pyRange: null,
        cyRange: null,
        pyLabel: '',
        cyLabel: '',
        dataDateRange: { min: null, max: null },
        selectedDimensions: [],
        gmPriceDefinition: 'margin-per-unit',
        aggregationResults: null,
        bridgeResults: null,
        isProcessing: false,
        shouldCancel: false,
        currentPage: 1,
        sortBy: 'total-desc',
        searchTerm: '',
        // Multi-year support
        detectedFiscalYears: [],
        selectedFiscalYears: [],
        hasMultipleYears: false,
        useMultiYearMode: false
    },

    init: function() {
        this.loadPreferences();
        this.bindEvents();
        this.showScreen('screen-landing');
    },

    loadPreferences: function() {
        try {
            const saved = localStorage.getItem(CONFIG.STORAGE_KEYS.LAST_CONFIG);
            if (saved) {
                const prefs = JSON.parse(saved);
                if (prefs.fyEndMonth) this.state.fyEndMonth = prefs.fyEndMonth;
                if (prefs.mode) this.state.mode = prefs.mode;
            }
        } catch (e) {}
    },

    savePreferences: function() {
        try {
            localStorage.setItem(CONFIG.STORAGE_KEYS.LAST_CONFIG, JSON.stringify({
                fyEndMonth: this.state.fyEndMonth,
                mode: this.state.mode
            }));
        } catch (e) {}
    },

    formatFileSize: function(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
        return (bytes / (1024 * 1024 * 1024)).toFixed(2) + ' GB';
    },

    showScreen: function(screenId) {
        document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
        document.getElementById(screenId).classList.add('active');

        const modeLabel = this.state.mode === 'gm' ? 'Gross Margin Mode' : 'Sales PVM Mode';
        document.querySelectorAll('.mode-indicator').forEach(el => el.textContent = modeLabel);
        document.querySelectorAll('.gm-only').forEach(el => el.classList.toggle('hidden', this.state.mode !== 'gm'));
        document.getElementById('gm-price-definition').classList.toggle('hidden', this.state.mode !== 'gm');
    },

    togglePeriodSections: function() {
        const useLTM = this.state.useLTM;
        document.getElementById('ltm-config-section').style.display = useLTM ? 'block' : 'none';
        document.getElementById('cy-fiscal-year-section').style.display = useLTM ? 'none' : 'block';
    },

    bindEvents: function() {
        // Mode selection
        document.querySelectorAll('.mode-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const b = e.target.closest('.mode-btn');
                if (!b) return;
                document.querySelectorAll('.mode-btn').forEach(x => x.classList.remove('selected'));
                b.classList.add('selected');
                this.state.mode = b.dataset.mode;
                this.savePreferences();
            });
        });

        // Navigation
        document.getElementById('btn-start').addEventListener('click', () => this.showScreen('screen-upload'));
        document.getElementById('btn-to-periods').addEventListener('click', () => this.goToPeriods());
        document.getElementById('btn-to-lod').addEventListener('click', () => this.goToLOD());
        document.getElementById('btn-process').addEventListener('click', () => this.startProcessing());
        document.getElementById('btn-cancel').addEventListener('click', () => this.cancelProcessing());
        document.getElementById('btn-export-excel').addEventListener('click', () => this.exportExcel());

        document.querySelectorAll('.btn-back').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const target = e.target.dataset.target;
                if (target) this.showScreen(target);
            });
        });

        // File upload
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');
        
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('drag-over'); });
        uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('drag-over'));
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) this.handleFileSelect(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) this.handleFileSelect(e.target.files[0]);
        });

        document.getElementById('btn-remove-file').addEventListener('click', () => this.removeFile());

        document.querySelectorAll('.column-select').forEach(select => {
            select.addEventListener('change', (e) => this.handleColumnMappingChange(e));
        });

        document.getElementById('date-format-select').addEventListener('change', (e) => {
            this.state.dateFormat = e.target.value === 'auto' ? this.state.detectedDateFormat?.formatId : e.target.value;
            this.updateDateParsePreview();
            this.validateColumnMappings();
        });

        document.getElementById('fy-end-month').addEventListener('change', (e) => {
            this.state.fyEndMonth = parseInt(e.target.value, 10);

            // Re-detect fiscal years if we have date range data
            if (this.state.dataDateRange.min && this.state.dataDateRange.max) {
                const fiscalYears = PeriodUtils.detectFiscalYears(
                    this.state.dataDateRange.min,
                    this.state.dataDateRange.max,
                    this.state.fyEndMonth
                );
                this.state.detectedFiscalYears = fiscalYears;
                const fullYears = fiscalYears.filter(y => y.fullyCovered);
                this.state.hasMultipleYears = fullYears.length >= 2;

                // If in multi-year mode, update the selection UI
                if (this.state.hasMultipleYears && fullYears.length >= 2) {
                    // Reset selected years when FY end changes
                    this.state.selectedFiscalYears = [];
                    this.showMultiYearSelection(fullYears);
                } else {
                    // Switch to two-period mode
                    document.getElementById('multi-year-section').style.display = 'none';
                    document.getElementById('two-period-config').style.display = 'block';
                    this.state.useMultiYearMode = false;
                    this.updatePeriodPreviews();
                }
            } else {
                this.updatePeriodPreviews();
            }

            this.savePreferences();
        });

        document.getElementById('ltm-end-date').addEventListener('change', (e) => {
            this.state.ltmEndDate = PeriodUtils.fromISODateString(e.target.value);
            this.updatePeriodPreviews();
        });

        document.getElementById('use-ltm-period').addEventListener('change', (e) => {
            this.state.useLTM = e.target.checked;
            this.togglePeriodSections();
            this.updatePeriodPreviews();
        });

        document.getElementById('cy-fiscal-year').addEventListener('change', (e) => {
            this.state.cyFiscalYear = parseInt(e.target.value, 10);
            this.updatePeriodPreviews();
        });

        document.querySelectorAll('input[name="gm-price-def"]').forEach(radio => {
            radio.addEventListener('change', (e) => this.state.gmPriceDefinition = e.target.value);
        });

        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        document.getElementById('detail-search').addEventListener('input', (e) => {
            this.state.searchTerm = e.target.value;
            this.state.currentPage = 1;
            this.renderDetailTable();
        });

        document.getElementById('detail-sort').addEventListener('change', (e) => {
            this.state.sortBy = e.target.value;
            this.renderDetailTable();
        });

        document.getElementById('detail-pagination').addEventListener('click', (e) => {
            if (e.target.dataset.page) {
                this.state.currentPage = parseInt(e.target.dataset.page, 10);
                this.renderDetailTable();
            }
        });

        document.getElementById('error-modal-close').addEventListener('click', () => UIRenderer.hideError());
        document.getElementById('error-modal-ok').addEventListener('click', () => UIRenderer.hideError());
    },

    async handleFileSelect(file) {
        const fileName = file.name.toLowerCase();
        const isExcel = fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
        const isCsv = fileName.endsWith('.csv');

        if (!isExcel && !isCsv) {
            UIRenderer.showError('Please select a CSV or Excel file (.csv, .xlsx, .xls)');
            return;
        }

        this.state.file = file;
        this.state.isExcelFile = isExcel;
        document.getElementById('file-info').classList.remove('hidden');
        document.getElementById('file-name').textContent = file.name;
        document.getElementById('file-size').textContent = this.formatFileSize(file.size);

        try {
            // Use appropriate parser based on file type
            const parser = isExcel ? ExcelParser : CSVParser;
            const { headers, sampleRows } = await parser.scanFile(file, 100);
            this.state.headers = headers;
            this.state.sampleRows = sampleRows;

            const mappings = CSVParser.detectColumnMappings(headers, sampleRows);
            this.state.columnMappings = mappings;
            
            this.populateColumnMappings();
            document.getElementById('column-mapping').classList.remove('hidden');
            this.validateColumnMappings();
        } catch (error) {
            UIRenderer.showError('Error reading file: ' + error.message);
        }
    },

    removeFile: function() {
        this.state.file = null;
        this.state.headers = [];
        this.state.sampleRows = [];
        document.getElementById('file-info').classList.add('hidden');
        document.getElementById('column-mapping').classList.add('hidden');
        document.getElementById('file-input').value = '';
        document.getElementById('btn-to-periods').disabled = true;
    },

    populateColumnMappings: function() {
        const headers = this.state.headers;
        const mappings = this.state.columnMappings;
        
        for (const field of ['date', 'sales', 'quantity', 'cost']) {
            const select = document.getElementById(`map-${field}`);
            UIRenderer.clearElement(select);
            select.appendChild(UIRenderer.createElement('option', { value: '' }, ['-- Select --']));
            
            for (const header of headers) {
                const opt = UIRenderer.createElement('option', { value: header }, [header]);
                if (mappings[field] === header) opt.selected = true;
                select.appendChild(opt);
            }
        }
        
        const dimContainer = document.getElementById('dimension-checkboxes');
        UIRenderer.clearElement(dimContainer);
        
        for (const header of headers) {
            const isRequired = [mappings.date, mappings.sales, mappings.quantity, mappings.cost].includes(header);
            const label = UIRenderer.createElement('label', { className: 'checkbox-label' });
            const checkbox = UIRenderer.createElement('input', {
                type: 'checkbox',
                value: header,
                checked: mappings.dimensions.includes(header) ? 'checked' : null,
                disabled: isRequired ? 'disabled' : null
            });
            checkbox.addEventListener('change', () => this.updateDimensionMappings());
            label.appendChild(checkbox);
            label.appendChild(UIRenderer.text(header + (isRequired ? ' (mapped)' : '')));
            dimContainer.appendChild(label);
        }
        
        if (mappings.date) this.detectAndShowDateFormat(mappings.date);
    },

    handleColumnMappingChange: function(e) {
        const field = e.target.dataset.field;
        const value = e.target.value || null;
        this.state.columnMappings[field] = value;
        this.populateColumnMappings();
        if (field === 'date' && value) this.detectAndShowDateFormat(value);
        this.validateColumnMappings();
    },

    updateDimensionMappings: function() {
        const checkboxes = document.querySelectorAll('#dimension-checkboxes input[type="checkbox"]');
        this.state.columnMappings.dimensions = [];
        checkboxes.forEach(cb => {
            if (cb.checked && !cb.disabled) this.state.columnMappings.dimensions.push(cb.value);
        });
    },

    detectAndShowDateFormat: function(dateColumn) {
        // Use appropriate parser based on file type
        const parser = this.state.isExcelFile ? ExcelParser : CSVParser;
        const samples = parser.extractDateSamples(this.state.sampleRows, dateColumn);
        if (samples.length === 0) {
            document.getElementById('date-format-section').classList.add('hidden');
            return;
        }

        const detected = PeriodUtils.detectDateFormat(samples);
        this.state.detectedDateFormat = detected;
        this.state.dateFormat = detected.formatId;

        const sampleList = document.getElementById('date-samples');
        UIRenderer.clearElement(sampleList);
        for (const sample of samples.slice(0, 5)) {
            sampleList.appendChild(UIRenderer.createElement('li', {}, [sample]));
        }

        document.getElementById('date-format-select').value = detected.formatId || 'auto';
        this.updateDateParsePreview();
        document.getElementById('date-format-section').classList.remove('hidden');
    },

    updateDateParsePreview: function() {
        const preview = document.getElementById('date-parse-preview');
        // Use appropriate parser based on file type
        const parser = this.state.isExcelFile ? ExcelParser : CSVParser;
        const samples = parser.extractDateSamples(this.state.sampleRows, this.state.columnMappings.date);
        
        if (samples.length === 0 || !this.state.dateFormat) { preview.textContent = ''; return; }
        
        const parsed = PeriodUtils.parseDate(samples[0], this.state.dateFormat);
        if (parsed) {
            preview.textContent = `✓ Parsed as: ${PeriodUtils.formatDate(parsed, 'long')}`;
            preview.classList.remove('error');
        } else {
            preview.textContent = '✗ Could not parse - select correct format';
            preview.classList.add('error');
        }
    },

    validateColumnMappings: function() {
        const m = this.state.columnMappings;
        let valid = m.date && m.sales && m.quantity;
        if (this.state.mode === 'gm' && !m.cost) valid = false;
        if (!this.state.dateFormat) valid = false;
        document.getElementById('btn-to-periods').disabled = !valid;
    }
};

// Initialize on DOM ready
document.addEventListener('DOMContentLoaded', () => App.init());
