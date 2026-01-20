/**
 * ============================================
 * PVM Bridge Tool - Excel Parser
 * ============================================
 *
 * Parses Excel files (.xlsx, .xls) using SheetJS.
 * Converts Excel data to same format as CSV parser for compatibility.
 * ============================================
 */

const ExcelParser = {
    /**
     * Convert row array to object keyed by headers
     * @private
     */
    _rowToObject: function(row, headers) {
        const obj = {};
        for (let i = 0; i < headers.length; i++) {
            obj[headers[i]] = row[i] !== undefined ? String(row[i]) : '';
        }
        return obj;
    },

    /**
     * Scan Excel file and extract headers + sample rows
     * @param {File} file - Excel file
     * @param {number} sampleSize - Number of rows to sample
     * @returns {Promise<{headers: string[], sampleRows: Object[]}>}
     */
    scanFile: async function(file, sampleSize = 100) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Use first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // Convert to array of arrays
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    if (jsonData.length === 0) {
                        reject(new Error('Excel file is empty'));
                        return;
                    }

                    // First row is headers
                    const headers = jsonData[0].map(h => String(h).trim());

                    // Rest are data rows (up to sampleSize) - convert to objects
                    const sampleRows = jsonData.slice(1, sampleSize + 1).map(row => {
                        const rowArray = row.map(cell => String(cell));
                        return this._rowToObject(rowArray, headers);
                    });

                    resolve({ headers, sampleRows });
                } catch (error) {
                    reject(new Error('Failed to parse Excel file: ' + error.message));
                }
            };

            reader.onerror = function() {
                reject(new Error('Failed to read Excel file'));
            };

            reader.readAsArrayBuffer(file);
        });
    },

    /**
     * Parse entire Excel file in chunks (for large files)
     * Converts to row-by-row format compatible with CSV parser
     */
    parseFile: async function(file, options) {
        const {
            onRow,
            onProgress,
            onHeaders,
            onComplete,
            onError,
            shouldCancel = () => false
        } = options;

        try {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    if (shouldCancel()) {
                        onComplete && onComplete({ cancelled: true, rowCount: 0 });
                        return;
                    }

                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Use first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // Convert to array of arrays
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    if (jsonData.length === 0) {
                        onError && onError(new Error('Excel file is empty'));
                        return;
                    }

                    // First row is headers
                    const headers = jsonData[0].map(h => String(h).trim());
                    onHeaders && onHeaders(headers);

                    // Process data rows
                    let rowCount = 0;
                    for (let i = 1; i < jsonData.length; i++) {
                        if (shouldCancel()) {
                            onComplete && onComplete({ cancelled: true, rowCount });
                            return;
                        }

                        const rowArray = jsonData[i].map(cell => String(cell));
                        const rowObj = this._rowToObject(rowArray, headers);
                        onRow && onRow(rowObj, i);
                        rowCount++;

                        // Simulate progress (Excel is loaded all at once, but we can still report progress)
                        if (rowCount % 1000 === 0 && onProgress) {
                            const progressBytes = Math.floor((i / jsonData.length) * file.size);
                            onProgress(progressBytes, file.size, rowCount);
                        }
                    }

                    // Final progress
                    onProgress && onProgress(file.size, file.size, rowCount);
                    onComplete && onComplete({ cancelled: false, rowCount });

                } catch (error) {
                    onError && onError(new Error('Failed to parse Excel file: ' + error.message));
                }
            };

            reader.onerror = function() {
                onError && onError(new Error('Failed to read Excel file'));
            };

            reader.readAsArrayBuffer(file);

        } catch (error) {
            onError && onError(error);
        }
    },

    /**
     * Extract date samples from Excel data
     */
    extractDateSamples: function(sampleRows, dateColumn) {
        const samples = new Set();
        for (const row of sampleRows) {
            const value = row[dateColumn];
            if (value && value.trim && value.trim()) {
                samples.add(value.trim());
                if (samples.size >= 10) break;
            }
        }
        return Array.from(samples);
    },

    /**
     * Check if file is an Excel file
     */
    isExcelFile: function(filename) {
        const lower = filename.toLowerCase();
        return lower.endsWith('.xlsx') || lower.endsWith('.xls');
    }
};
