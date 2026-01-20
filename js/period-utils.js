/**
 * ============================================
 * PVM Bridge Tool - Period Utilities
 * ============================================
 * 
 * Handles all date and period logic:
 * - Fiscal year calculations
 * - LTM (Last Twelve Months) range determination
 * - Period classification (PY vs CY/LTM)
 * - Date format detection and parsing
 * 
 * All calculations are deterministic and auditable.
 * ============================================
 */

const PeriodUtils = {
    /**
     * Detect date format from sample values
     * Returns the most likely format based on pattern matching
     * 
     * @param {string[]} samples - Array of date strings to analyze
     * @returns {Object} - { formatId, confidence, parser }
     */
    detectDateFormat: function(samples) {
        if (!samples || samples.length === 0) {
            return { formatId: null, confidence: 0, parser: null };
        }

        // Filter out empty/null values
        const validSamples = samples.filter(s => s && s.trim());
        if (validSamples.length === 0) {
            return { formatId: null, confidence: 0, parser: null };
        }

        // Try each format and count matches
        const formatScores = {};
        
        for (const format of CONFIG.DATE_FORMATS) {
            let matches = 0;
            let validDates = 0;
            
            for (const sample of validSamples) {
                const trimmed = sample.trim();
                if (format.pattern.test(trimmed)) {
                    matches++;
                    // Also verify it parses to a valid date
                    try {
                        const parsed = format.parse(trimmed);
                        if (parsed && !isNaN(parsed.getTime())) {
                            // Sanity check: year should be reasonable (1900-2100)
                            const year = parsed.getFullYear();
                            if (year >= 1900 && year <= 2100) {
                                validDates++;
                            }
                        }
                    } catch (e) {
                        // Parse failed, don't count as valid
                    }
                }
            }
            
            formatScores[format.id] = {
                matches: matches,
                validDates: validDates,
                score: validDates / validSamples.length
            };
        }

        // Find best format
        let bestFormat = null;
        let bestScore = 0;
        
        for (const format of CONFIG.DATE_FORMATS) {
            const score = formatScores[format.id];
            if (score.score > bestScore) {
                bestScore = score.score;
                bestFormat = format;
            }
        }

        // Handle MM/DD/YYYY vs DD/MM/YYYY ambiguity
        // If both match, check if any values have day > 12
        if (bestFormat && (bestFormat.id === 'MM/DD/YYYY' || bestFormat.id === 'DD/MM/YYYY')) {
            const mmddScore = formatScores['MM/DD/YYYY'];
            const ddmmScore = formatScores['DD/MM/YYYY'];
            
            if (mmddScore && ddmmScore && 
                Math.abs(mmddScore.score - ddmmScore.score) < 0.1) {
                // Ambiguous - look for values that disambiguate
                for (const sample of validSamples) {
                    const parts = sample.split('/');
                    if (parts.length === 3) {
                        const first = parseInt(parts[0], 10);
                        const second = parseInt(parts[1], 10);
                        
                        // If first part > 12, it must be day (DD/MM/YYYY)
                        if (first > 12 && first <= 31) {
                            bestFormat = CONFIG.DATE_FORMATS.find(f => f.id === 'DD/MM/YYYY');
                            break;
                        }
                        // If second part > 12, it must be day (MM/DD/YYYY)
                        if (second > 12 && second <= 31) {
                            bestFormat = CONFIG.DATE_FORMATS.find(f => f.id === 'MM/DD/YYYY');
                            break;
                        }
                    }
                }
            }
        }

        return {
            formatId: bestFormat ? bestFormat.id : null,
            confidence: bestScore,
            parser: bestFormat ? bestFormat.parse : null,
            format: bestFormat
        };
    },

    /**
     * Parse a date string using a specific format
     * 
     * @param {string} dateStr - Date string to parse
     * @param {string} formatId - Format ID (e.g., 'YYYY-MM-DD')
     * @returns {Date|null} - Parsed date or null if invalid
     */
    parseDate: function(dateStr, formatId) {
        if (!dateStr || !formatId) return null;
        
        const format = CONFIG.DATE_FORMATS.find(f => f.id === formatId);
        if (!format) {
            // Try auto-detect if format not found
            const detected = this.detectDateFormat([dateStr]);
            if (detected.parser) {
                return detected.parser(dateStr.trim());
            }
            return null;
        }
        
        try {
            const parsed = format.parse(dateStr.trim());
            if (parsed && !isNaN(parsed.getTime())) {
                return parsed;
            }
        } catch (e) {
            // Parse failed
        }
        
        return null;
    },

    /**
     * Calculate fiscal year for a given date
     * 
     * @param {Date} date - The date to classify
     * @param {number} fyEndMonth - Month when FY ends (1-12)
     * @returns {number} - Fiscal year (e.g., 2024)
     * 
     * Example: If FY ends in June (6), then:
     * - July 2023 -> FY 2024
     * - June 2024 -> FY 2024
     * - July 2024 -> FY 2025
     */
    getFiscalYear: function(date, fyEndMonth) {
        const month = date.getMonth() + 1; // 1-12
        const year = date.getFullYear();
        
        if (month > fyEndMonth) {
            // After FY end, belongs to next FY
            return year + 1;
        } else {
            // On or before FY end, belongs to current calendar year's FY
            return year;
        }
    },

    /**
     * Get fiscal year start and end dates
     * 
     * @param {number} fiscalYear - The fiscal year (e.g., 2024)
     * @param {number} fyEndMonth - Month when FY ends (1-12)
     * @returns {Object} - { start: Date, end: Date }
     * 
     * Example: FY 2024 with June year-end:
     * - Start: July 1, 2023
     * - End: June 30, 2024
     */
    getFiscalYearRange: function(fiscalYear, fyEndMonth) {
        let startYear, startMonth, endYear, endMonth;
        
        if (fyEndMonth === 12) {
            // Calendar year = fiscal year
            startYear = fiscalYear;
            startMonth = 1;
            endYear = fiscalYear;
            endMonth = 12;
        } else {
            // FY starts in month after FY end, previous calendar year
            startYear = fiscalYear - 1;
            startMonth = fyEndMonth + 1;
            endYear = fiscalYear;
            endMonth = fyEndMonth;
        }
        
        // Get last day of end month
        const endDate = new Date(endYear, endMonth, 0); // Day 0 = last day of previous month
        
        return {
            start: new Date(startYear, startMonth - 1, 1),
            end: endDate
        };
    },

    /**
     * Get LTM (Last Twelve Months) range ending on a specific date
     * 
     * @param {Date} endDate - LTM end date
     * @returns {Object} - { start: Date, end: Date }
     */
    getLTMRange: function(endDate) {
        // LTM starts 12 months before end date
        const start = new Date(endDate);
        start.setFullYear(start.getFullYear() - 1);
        start.setDate(start.getDate() + 1); // Day after to make it exactly 12 months
        
        return {
            start: start,
            end: new Date(endDate)
        };
    },

    /**
     * Determine prior fiscal year based on LTM end date
     * 
     * @param {Date} ltmEndDate - End date of LTM period
     * @param {number} fyEndMonth - Month when FY ends (1-12)
     * @returns {number} - Prior fiscal year number
     */
    getPriorFiscalYear: function(ltmEndDate, fyEndMonth) {
        const ltmFY = this.getFiscalYear(ltmEndDate, fyEndMonth);
        return ltmFY - 1;
    },

    /**
     * Classify a date into period: 'PY', 'CY', or null (outside both)
     * 
     * @param {Date} date - Date to classify
     * @param {Object} pyRange - { start: Date, end: Date } for Prior Year
     * @param {Object} cyRange - { start: Date, end: Date } for Current Year/LTM
     * @returns {string|null} - 'PY', 'CY', or null
     */
    classifyPeriod: function(date, pyRange, cyRange) {
        const time = date.getTime();
        
        if (time >= pyRange.start.getTime() && time <= pyRange.end.getTime()) {
            return 'PY';
        }
        
        if (time >= cyRange.start.getTime() && time <= cyRange.end.getTime()) {
            return 'CY';
        }
        
        return null;
    },

    /**
     * Format a date for display
     * 
     * @param {Date} date - Date to format
     * @param {string} style - 'short', 'medium', or 'long'
     * @returns {string} - Formatted date string
     */
    formatDate: function(date, style = 'medium') {
        if (!date || isNaN(date.getTime())) return '--';
        
        const day = date.getDate();
        const month = date.getMonth();
        const year = date.getFullYear();
        
        switch (style) {
            case 'short':
                return `${month + 1}/${day}/${year}`;
            case 'long':
                return `${CONFIG.MONTH_NAMES[month]} ${day}, ${year}`;
            case 'medium':
            default:
                return `${CONFIG.MONTH_NAMES_SHORT[month]} ${day}, ${year}`;
        }
    },

    /**
     * Format a date range for display
     * 
     * @param {Date} start - Start date
     * @param {Date} end - End date
     * @returns {string} - Formatted range string
     */
    formatDateRange: function(start, end) {
        return `${this.formatDate(start)} - ${this.formatDate(end)}`;
    },

    /**
     * Convert Date to ISO date string (YYYY-MM-DD) for input[type="date"]
     * 
     * @param {Date} date - Date to convert
     * @returns {string} - ISO date string
     */
    toISODateString: function(date) {
        if (!date || isNaN(date.getTime())) return '';
        
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        
        return `${year}-${month}-${day}`;
    },

    /**
     * Parse ISO date string (YYYY-MM-DD) from input[type="date"]
     * 
     * @param {string} isoString - ISO date string
     * @returns {Date|null} - Parsed date
     */
    fromISODateString: function(isoString) {
        if (!isoString) return null;
        
        const parts = isoString.split('-');
        if (parts.length !== 3) return null;
        
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const day = parseInt(parts[2], 10);
        
        const date = new Date(year, month, day);
        if (isNaN(date.getTime())) return null;
        
        return date;
    },

    /**
     * Check if two dates are the same day
     * 
     * @param {Date} date1 - First date
     * @param {Date} date2 - Second date
     * @returns {boolean}
     */
    isSameDay: function(date1, date2) {
        return date1.getFullYear() === date2.getFullYear() &&
               date1.getMonth() === date2.getMonth() &&
               date1.getDate() === date2.getDate();
    },

    /**
     * Get the last day of a month
     * 
     * @param {number} year - Year
     * @param {number} month - Month (1-12)
     * @returns {Date}
     */
    getLastDayOfMonth: function(year, month) {
        // Day 0 of next month = last day of current month
        return new Date(year, month, 0);
    },

    /**
     * Validate period configuration
     * Returns warnings/errors for the user
     * 
     * @param {Object} config - { fyEndMonth, ltmEndDate, pyRange, cyRange }
     * @returns {Object} - { valid: boolean, warnings: string[], errors: string[] }
     */
    validatePeriodConfig: function(config) {
        const warnings = [];
        const errors = [];
        
        // Check if LTM end date is valid
        if (!config.ltmEndDate || isNaN(config.ltmEndDate.getTime())) {
            errors.push('LTM end date is invalid.');
        }
        
        // Check if periods overlap
        if (config.pyRange && config.cyRange) {
            if (config.pyRange.end >= config.cyRange.start) {
                warnings.push('Prior Year and LTM periods overlap. Some data may be counted in both periods.');
            }
            
            // Check gap between periods
            const gapDays = (config.cyRange.start - config.pyRange.end) / (1000 * 60 * 60 * 24);
            if (gapDays > 365) {
                warnings.push(`There is a ${Math.round(gapDays)} day gap between Prior Year and LTM. Some data may be excluded.`);
            }
        }
        
        return {
            valid: errors.length === 0,
            warnings: warnings,
            errors: errors
        };
    }
};

// Export for use in other modules and web worker
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PeriodUtils;
}
