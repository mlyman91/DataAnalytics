/**
 * ============================================
 * PVM Bridge Tool - CSV Parser
 * ============================================
 * 
 * Streaming CSV parser designed for large files.
 * Key features:
 * - Chunked reading (never loads full file)
 * - Handles quoted fields with embedded commas/newlines
 * - Memory-efficient processing
 * - Progress callbacks for UI updates
 * ============================================
 */

const CSVParser = {
    /**
     * Parse a CSV file in streaming chunks
     */
    parseFile: async function(file, options) {
        const {
            onRow,
            onProgress,
            onHeaders,
            onComplete,
            onError,
            shouldCancel = () => false,
            chunkSize = CONFIG.CHUNK_SIZE
        } = options;

        const totalBytes = file.size;
        let bytesRead = 0;
        let rowCount = 0;
        let headers = null;
        let buffer = '';

        try {
            const reader = file.stream().getReader();
            const decoder = new TextDecoder('utf-8');

            while (true) {
                if (shouldCancel()) {
                    onComplete && onComplete({ cancelled: true, rowCount });
                    return;
                }

                const { done, value } = await reader.read();
                
                if (done) {
                    if (buffer.length > 0) {
                        const finalRows = this._extractCompleteRows(buffer + '\n').rows;
                        for (const row of finalRows) {
                            if (!headers) {
                                headers = row;
                                onHeaders && onHeaders(headers);
                            } else {
                                rowCount++;
                                onRow && onRow(this._rowToObject(row, headers), rowCount);
                            }
                        }
                    }
                    break;
                }

                const chunk = decoder.decode(value, { stream: true });
                bytesRead += value.length;
                buffer += chunk;

                const { rows, remainder } = this._extractCompleteRows(buffer);
                buffer = remainder;

                for (const row of rows) {
                    if (!headers) {
                        headers = row;
                        onHeaders && onHeaders(headers);
                    } else {
                        rowCount++;
                        onRow && onRow(this._rowToObject(row, headers), rowCount);
                        
                        if (rowCount % CONFIG.PROGRESS_UPDATE_INTERVAL === 0) {
                            onProgress && onProgress(bytesRead, totalBytes, rowCount);
                        }
                    }
                }

                onProgress && onProgress(bytesRead, totalBytes, rowCount);
            }

            onComplete && onComplete({ cancelled: false, rowCount, headers });

        } catch (error) {
            onError && onError(error);
        }
    },

    /**
     * Extract complete CSV rows from buffer
     * @private
     */
    _extractCompleteRows: function(buffer) {
        const rows = [];
        let currentRow = [];
        let currentField = '';
        let inQuotes = false;
        let i = 0;

        while (i < buffer.length) {
            const char = buffer[i];
            const nextChar = buffer[i + 1];

            if (inQuotes) {
                if (char === '"') {
                    if (nextChar === '"') {
                        currentField += '"';
                        i += 2;
                        continue;
                    } else {
                        inQuotes = false;
                        i++;
                        continue;
                    }
                } else {
                    currentField += char;
                    i++;
                    continue;
                }
            }

            if (char === '"') {
                inQuotes = true;
                i++;
                continue;
            }

            if (char === ',') {
                currentRow.push(currentField.trim());
                currentField = '';
                i++;
                continue;
            }

            if (char === '\r') {
                if (nextChar === '\n') i++;
                currentRow.push(currentField.trim());
                if (currentRow.length > 0 && !(currentRow.length === 1 && currentRow[0] === '')) {
                    rows.push(currentRow);
                }
                currentRow = [];
                currentField = '';
                i++;
                continue;
            }

            if (char === '\n') {
                currentRow.push(currentField.trim());
                if (currentRow.length > 0 && !(currentRow.length === 1 && currentRow[0] === '')) {
                    rows.push(currentRow);
                }
                currentRow = [];
                currentField = '';
                i++;
                continue;
            }

            currentField += char;
            i++;
        }

        let remainder = '';
        if (currentField || currentRow.length > 0) {
            currentRow.push(currentField);
            remainder = this._reconstructRow(currentRow, inQuotes);
        }

        return { rows, remainder };
    },

    _reconstructRow: function(row, inQuotes) {
        let result = row.map(f => {
            if (f.includes(',') || f.includes('"') || f.includes('\n')) {
                return '"' + f.replace(/"/g, '""') + '"';
            }
            return f;
        }).join(',');
        
        if (inQuotes) {
            const lastComma = result.lastIndexOf(',');
            if (lastComma >= 0) {
                result = result.substring(0, lastComma + 1) + '"' + result.substring(lastComma + 1);
            } else {
                result = '"' + result;
            }
        }
        return result;
    },

    _rowToObject: function(row, headers) {
        const obj = {};
        for (let i = 0; i < headers.length; i++) {
            obj[headers[i]] = row[i] !== undefined ? row[i] : '';
        }
        return obj;
    },

    /**
     * Quick scan of file to get headers and sample rows
     */
    scanFile: async function(file, maxRows = 100) {
        return new Promise((resolve, reject) => {
            const headers = [];
            const sampleRows = [];

            this.parseFile(file, {
                onHeaders: (h) => headers.push(...h),
                onRow: (row) => {
                    if (sampleRows.length < maxRows) sampleRows.push(row);
                },
                shouldCancel: () => sampleRows.length >= maxRows,
                onComplete: () => resolve({ headers, sampleRows }),
                onError: reject
            });
        });
    },

    /**
     * Detect column types from sample data
     */
    detectColumnMappings: function(headers, sampleRows) {
        const mappings = { date: null, sales: null, quantity: null, cost: null, dimensions: [] };
        const usedColumns = new Set();

        for (const field of ['date', 'sales', 'quantity', 'cost']) {
            const patterns = CONFIG.COLUMN_PATTERNS[field];
            for (const pattern of patterns) {
                for (const header of headers) {
                    if (!usedColumns.has(header) && pattern.test(header)) {
                        mappings[field] = header;
                        usedColumns.add(header);
                        break;
                    }
                }
                if (mappings[field]) break;
            }
        }

        for (const header of headers) {
            if (!usedColumns.has(header)) {
                const isNumeric = sampleRows.every(row => {
                    const val = row[header];
                    if (!val || val === '') return true;
                    return !isNaN(parseFloat(String(val).replace(/[,$]/g, '')));
                });
                if (!isNumeric) mappings.dimensions.push(header);
            }
        }

        return mappings;
    },

    extractDateSamples: function(sampleRows, dateColumn) {
        const samples = new Set();
        for (const row of sampleRows) {
            const value = row[dateColumn];
            if (value && value.trim()) {
                samples.add(value.trim());
                if (samples.size >= CONFIG.MAX_PREVIEW_ROWS) break;
            }
        }
        return Array.from(samples);
    },

    parseNumber: function(value) {
        if (value === null || value === undefined || value === '') return NaN;
        
        let str = String(value).trim();
        const isNegative = str.startsWith('(') && str.endsWith(')');
        if (isNegative) str = str.slice(1, -1);
        
        str = str.replace(/[$€£¥,]/g, '');
        const hasMinusSign = str.startsWith('-');
        if (hasMinusSign) str = str.slice(1);
        
        const num = parseFloat(str);
        if (isNaN(num)) return NaN;
        
        return (isNegative || hasMinusSign) ? -num : num;
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = CSVParser;
}
