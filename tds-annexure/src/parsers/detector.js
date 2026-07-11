'use strict';

/**
 * Format Detector
 * Determines which parser to use based on file type and content signals.
 *
 * Decision logic:
 *   1. If file is PDF → Format B (partner TDS PDF)
 *   2. If file is XLS/XLSX:
 *      a. If it has a sheet named "Annexure" with "Deductee code" header → Format C (passthrough)
 *      b. If it contains "Nature Of Payment" or "PAN No :" patterns → Format A (Galaxy style)
 *      c. Otherwise → try Format C as fallback
 */

const XLSX = require('xlsx');
const path = require('path');

function detectFormat(buffer, originalFilename) {
  const ext = path.extname(originalFilename || '').toLowerCase();

  // PDF → always Format B
  if (ext === '.pdf') return 'B';

  // Excel files (.xls, .xlsx)
  if (ext === '.xls' || ext === '.xlsx') {
    try {
      const workbook = XLSX.read(buffer, { type: 'buffer' });

      // Check each sheet for content signals
      for (const sheetName of workbook.SheetNames) {
        // Signal 1: Sheet named "Annexure" → likely Format C
        if (/annexure/i.test(sheetName)) {
          const sheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: '',
            blankrows: false,
          });
          // Confirm by looking for "Deductee code" or "Section code" headers
          const firstFewRows = rows.slice(0, 5).map(r =>
            r.map(c => String(c).trim()).join(' ')
          ).join(' ');

          if (/Deductee\s*code|Section\s*code|PAN\s*of\s*deductee/i.test(firstFewRows)) {
            return 'C';
          }
        }

        // Signal 2: "Nature Of Payment" → Format A
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          blankrows: false,
        });
        const allText = rows
          .map(r => r.map(c => String(c).trim()).join(' '))
          .join('\n');

        if (/Nature\s+Of\s+Payment/i.test(allText)) return 'A';
        if (/PAN\s+No\s*:/i.test(allText)) return 'A';

        // Signal 3: Has "Deductee code" anywhere → Format C
        if (/Deductee\s*code|Section\s*code/i.test(allText)) return 'C';

        // Signal 4: Has partner remuneration table → could be Format B content in Excel
        if (/Partner\s*(Interest|Remuneration)/i.test(allText)) return 'B_EXCEL';
      }
    } catch (e) {
      // Could not parse as Excel — may be corrupted
      return 'UNKNOWN';
    }

    // Default Excel fallback → Format C
    return 'C';
  }

  return 'UNKNOWN';
}

module.exports = { detectFormat };
