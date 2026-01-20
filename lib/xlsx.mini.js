/**
 * ============================================
 * Minimal XLSX Writer
 * ============================================
 * 
 * Lightweight Excel file generator.
 * Creates .xlsx files without external dependencies.
 * 
 * Based on the Office Open XML format.
 * Supports multiple sheets, basic formatting.
 * ============================================
 */

const XLSX = (function() {
    'use strict';

    // Helper: Convert array-of-arrays to sheet object
    function aoa_to_sheet(data) {
        const sheet = { '!ref': '', '!cols': [], data: data };
        
        if (data.length === 0) {
            sheet['!ref'] = 'A1:A1';
            return sheet;
        }

        let maxCol = 0;
        for (const row of data) {
            if (row.length > maxCol) maxCol = row.length;
        }

        const startCol = 'A';
        const endCol = colName(maxCol - 1);
        const endRow = data.length;
        sheet['!ref'] = `A1:${endCol}${endRow}`;

        return sheet;
    }

    // Helper: Column number to letter (0 = A, 25 = Z, 26 = AA)
    function colName(n) {
        let s = '';
        n++;
        while (n > 0) {
            n--;
            s = String.fromCharCode(65 + (n % 26)) + s;
            n = Math.floor(n / 26);
        }
        return s;
    }

    // Create new workbook
    function book_new() {
        return {
            SheetNames: [],
            Sheets: {}
        };
    }

    // Add sheet to workbook
    function book_append_sheet(wb, ws, name) {
        if (!name) name = 'Sheet' + (wb.SheetNames.length + 1);
        // Sanitize name (max 31 chars, no special chars)
        name = name.substring(0, 31).replace(/[:\\/?*\[\]]/g, '_');
        
        wb.SheetNames.push(name);
        wb.Sheets[name] = ws;
    }

    // Write workbook to file (triggers download)
    function writeFile(wb, filename) {
        const blob = write(wb, { type: 'blob' });
        
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // Write workbook to blob
    function write(wb, opts) {
        const zip = new JSZip();

        // [Content_Types].xml
        zip.file('[Content_Types].xml', generateContentTypes(wb));

        // _rels/.rels
        zip.file('_rels/.rels', generateRels());

        // xl/workbook.xml
        zip.file('xl/workbook.xml', generateWorkbook(wb));

        // xl/_rels/workbook.xml.rels
        zip.file('xl/_rels/workbook.xml.rels', generateWorkbookRels(wb));

        // xl/styles.xml
        zip.file('xl/styles.xml', generateStyles());

        // xl/sharedStrings.xml and sheets
        const { sharedStrings, stringMap } = collectSharedStrings(wb);
        zip.file('xl/sharedStrings.xml', generateSharedStrings(sharedStrings));

        for (let i = 0; i < wb.SheetNames.length; i++) {
            const sheetName = wb.SheetNames[i];
            const sheet = wb.Sheets[sheetName];
            zip.file(`xl/worksheets/sheet${i + 1}.xml`, generateSheet(sheet, stringMap));
        }

        return zip.generateAsync ? 
            zip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) :
            zip.generate({ type: 'blob' });
    }

    function generateContentTypes(wb) {
        let sheets = '';
        for (let i = 0; i < wb.SheetNames.length; i++) {
            sheets += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
${sheets}
</Types>`;
    }

    function generateRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    }

    function generateWorkbook(wb) {
        let sheets = '';
        for (let i = 0; i < wb.SheetNames.length; i++) {
            const name = escapeXml(wb.SheetNames[i]);
            sheets += `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheets}</sheets>
</workbook>`;
    }

    function generateWorkbookRels(wb) {
        let rels = '';
        for (let i = 0; i < wb.SheetNames.length; i++) {
            rels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
        }
        const styleId = wb.SheetNames.length + 1;
        const stringsId = wb.SheetNames.length + 2;

        rels += `<Relationship Id="rId${styleId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
        rels += `<Relationship Id="rId${stringsId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
    }

    function generateStyles() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
<numFmt numFmtId="164" formatCode="#,##0.00"/>
<numFmt numFmtId="165" formatCode="0.00%"/>
</numFmts>
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>`;
    }

    function collectSharedStrings(wb) {
        const strings = [];
        const map = new Map();

        for (const sheetName of wb.SheetNames) {
            const sheet = wb.Sheets[sheetName];
            if (!sheet.data) continue;

            for (const row of sheet.data) {
                for (const cell of row) {
                    if (typeof cell === 'string' && !map.has(cell)) {
                        map.set(cell, strings.length);
                        strings.push(cell);
                    }
                }
            }
        }

        return { sharedStrings: strings, stringMap: map };
    }

    function generateSharedStrings(strings) {
        let items = '';
        for (const s of strings) {
            items += `<si><t>${escapeXml(s)}</t></si>`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
${items}
</sst>`;
    }

    function generateSheet(sheet, stringMap) {
        const data = sheet.data || [];
        let rows = '';
        let cols = '';

        // Column widths
        if (sheet['!cols']) {
            cols = '<cols>';
            for (let i = 0; i < sheet['!cols'].length; i++) {
                const w = sheet['!cols'][i]?.wch || 10;
                cols += `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`;
            }
            cols += '</cols>';
        }

        for (let r = 0; r < data.length; r++) {
            const row = data[r];
            let cells = '';

            for (let c = 0; c < row.length; c++) {
                const cell = row[c];
                const ref = colName(c) + (r + 1);

                if (cell === null || cell === undefined || cell === '') {
                    continue;
                }

                if (typeof cell === 'number') {
                    // Check if it's a percentage (between -1 and 1 and has decimal)
                    const isPercent = Math.abs(cell) <= 1 && cell !== Math.floor(cell) && 
                                     r > 0 && data[r-1] && 
                                     (String(data[r-1][c] || '').includes('%') || 
                                      String(row[c-1] || '').includes('%'));
                    const style = isPercent ? ' s="2"' : ' s="1"';
                    cells += `<c r="${ref}"${style}><v>${cell}</v></c>`;
                } else if (typeof cell === 'string') {
                    const idx = stringMap.get(cell);
                    if (idx !== undefined) {
                        cells += `<c r="${ref}" t="s"><v>${idx}</v></c>`;
                    }
                } else if (typeof cell === 'boolean') {
                    cells += `<c r="${ref}" t="b"><v>${cell ? 1 : 0}</v></c>`;
                }
            }

            if (cells) {
                rows += `<row r="${r + 1}">${cells}</row>`;
            }
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${cols}
<sheetData>${rows}</sheetData>
</worksheet>`;
    }

    function escapeXml(s) {
        if (typeof s !== 'string') return s;
        return s.replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
    }

    // Return public API
    return {
        utils: {
            book_new: book_new,
            book_append_sheet: book_append_sheet,
            aoa_to_sheet: aoa_to_sheet
        },
        writeFile: writeFile,
        write: write
    };
})();

/**
 * Minimal JSZip implementation for XLSX generation
 */
const JSZip = (function() {
    'use strict';

    function JSZip() {
        this.files = {};
    }

    JSZip.prototype.file = function(name, content) {
        this.files[name] = content;
        return this;
    };

    JSZip.prototype.generate = function(options) {
        // Create ZIP file using native APIs
        const encoder = new TextEncoder();
        const parts = [];
        const centralDirectory = [];
        let offset = 0;

        for (const [name, content] of Object.entries(this.files)) {
            const nameBytes = encoder.encode(name);
            const contentBytes = encoder.encode(content);
            
            // Local file header
            const localHeader = createLocalHeader(nameBytes, contentBytes);
            parts.push(localHeader);
            parts.push(nameBytes);
            parts.push(contentBytes);

            // Central directory entry
            centralDirectory.push(createCentralEntry(nameBytes, contentBytes, offset));

            offset += localHeader.length + nameBytes.length + contentBytes.length;
        }

        // Central directory
        const cdStart = offset;
        for (const entry of centralDirectory) {
            parts.push(entry);
            offset += entry.length;
        }
        const cdSize = offset - cdStart;

        // End of central directory
        parts.push(createEndRecord(Object.keys(this.files).length, cdSize, cdStart));

        return new Blob(parts, { type: options?.mimeType || 'application/zip' });
    };

    function createLocalHeader(name, content) {
        const header = new Uint8Array(30);
        const view = new DataView(header.buffer);

        view.setUint32(0, 0x04034b50, true); // Signature
        view.setUint16(4, 20, true);          // Version needed
        view.setUint16(6, 0, true);           // Flags
        view.setUint16(8, 0, true);           // Compression (store)
        view.setUint16(10, 0, true);          // Mod time
        view.setUint16(12, 0, true);          // Mod date
        view.setUint32(14, crc32(content), true); // CRC32
        view.setUint32(18, content.length, true); // Compressed size
        view.setUint32(22, content.length, true); // Uncompressed size
        view.setUint16(26, name.length, true);    // Name length
        view.setUint16(28, 0, true);              // Extra length

        return header;
    }

    function createCentralEntry(name, content, offset) {
        const entry = new Uint8Array(46);
        const view = new DataView(entry.buffer);

        view.setUint32(0, 0x02014b50, true);  // Signature
        view.setUint16(4, 20, true);          // Version made by
        view.setUint16(6, 20, true);          // Version needed
        view.setUint16(8, 0, true);           // Flags
        view.setUint16(10, 0, true);          // Compression
        view.setUint16(12, 0, true);          // Mod time
        view.setUint16(14, 0, true);          // Mod date
        view.setUint32(16, crc32(content), true);
        view.setUint32(20, content.length, true);
        view.setUint32(24, content.length, true);
        view.setUint16(28, name.length, true);
        view.setUint16(30, 0, true);          // Extra length
        view.setUint16(32, 0, true);          // Comment length
        view.setUint16(34, 0, true);          // Disk number
        view.setUint16(36, 0, true);          // Internal attrs
        view.setUint32(38, 0, true);          // External attrs
        view.setUint32(42, offset, true);     // Offset

        return new Uint8Array([...entry, ...name]);
    }

    function createEndRecord(count, cdSize, cdOffset) {
        const record = new Uint8Array(22);
        const view = new DataView(record.buffer);

        view.setUint32(0, 0x06054b50, true);  // Signature
        view.setUint16(4, 0, true);           // Disk number
        view.setUint16(6, 0, true);           // CD disk
        view.setUint16(8, count, true);       // Entries on disk
        view.setUint16(10, count, true);      // Total entries
        view.setUint32(12, cdSize, true);     // CD size
        view.setUint32(16, cdOffset, true);   // CD offset
        view.setUint16(20, 0, true);          // Comment length

        return record;
    }

    // CRC32 table
    const crcTable = (function() {
        const table = new Uint32Array(256);
        for (let i = 0; i < 256; i++) {
            let c = i;
            for (let j = 0; j < 8; j++) {
                c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            }
            table[i] = c;
        }
        return table;
    })();

    function crc32(data) {
        let crc = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) {
            crc = crcTable[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    return JSZip;
})();
