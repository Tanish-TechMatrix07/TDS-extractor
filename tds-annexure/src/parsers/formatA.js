'use strict';

/**
 * Format A Parser — Galaxy Enterprise / TDS Report style
 *
 * Known raw structures (all confirmed from debug logs):
 *
 *   January variant — 8 columns:
 *   [5] [null, null, "Party Name", "TDS Reason",
 *        "Total Voucher Assessable Amt", "TDS Assessable Amount", "TDS %", "Net TDS to be Paid"]
 *   [8] [null, null, "MADHUR GANESH TRADING CO.", null, 1724450.5, 1724450.5, 0.1, 1724]
 *
 *   February variant — 6 columns (no separate TDS Assessable or TDS % column):
 *   [5] [null, null, "Party Name", "TDS Reason",
 *        "Total Voucher Assessable Amt", "Net TDS to be Paid"]
 *   [8] [null, null, "MADHUR GANESH TRADING CO.", null, 1957800, 1958]
 *
 * Strategy:
 *   1. Scan all rows BEFORE the first data row for a header row
 *      (detected by keywords: "Party Name", "Total Voucher", "Net TDS", "TDS %")
 *   2. Map column indices semantically from header text — no hardcoded positions
 *   3. For every data row, validate tds/amount ratio; recalculate if corrupt
 *
 * Column semantic keywords (case-insensitive):
 *   name   → "Party Name"
 *   amount → "Total Voucher"  (prefer this over "TDS Assessable" which is a subset)
 *   rate   → "TDS %"
 *   tds    → "Net TDS"
 *
 * If no header row is found, fall back to the legacy fixed positions (col 2/4/5).
 */

const XLSX = require('xlsx');

// Legacy fallback column indices (February format, no explicit TDS % col)
const FALLBACK_NAME   = 2;
const FALLBACK_AMOUNT = 4;
const FALLBACK_TDS    = 5;

function parseFormatA(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });

  // ── Find first sheet with meaningful data ──────────────────────────────────
  let rawRows = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
      blankrows: true,   // keep blank rows so row indices stay aligned
      raw: true,         // numbers as numbers
    });
    const nonEmpty = rows.filter(r => r && r.some(c => c !== null && c !== ''));
    if (nonEmpty.length > 3) {
      rawRows = rows;
      break;
    }
  }

  const str = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  // ── Pass 1: scan header rows to extract deductor info + column positions ────
  let deductorName = '';
  let tan          = '';
  let toDate       = '';

  // Column indices — will be overridden if a header row is found
  let nameColIdx   = FALLBACK_NAME;
  let amtColIdx    = FALLBACK_AMOUNT;
  let assessableAmtColIdx = -1;
  let tdsColIdx    = FALLBACK_TDS;
  let rateColIdx   = -1;  // -1 = not present in this variant
  let headerRowIdx = -1;  // row index of the detected column header row

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row) continue;
    const cells = row.map(str);
    const joined = cells.join(' ').trim();
    if (!joined) continue;

    // Company name — first non-empty row that looks like a name
    if (!deductorName) {
      const candidate = cells.find(c => c.length > 2 && !/^\d/.test(c));
      if (candidate) deductorName = candidate;
      continue; // always continue after grabbing name on row 0
    }

    // TAN anywhere in header section
    if (!tan) {
      const m = joined.match(/\b([A-Z]{4}[0-9]{5}[A-Z])\b/);
      if (m) tan = m[1];
    }

    // "To" date — e.g. "From Date 01/01/2026 To 31/01/2026"
    if (!toDate) {
      const m = joined.match(/\bTo\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);
      if (m) toDate = normaliseDate(m[1]);
    }

    // Column header row detection:
    // Must contain at least one of these strong markers
    const isHeaderRow = (
      /Total\s*Voucher/i.test(joined) ||
      /Net\s*TDS/i.test(joined)       ||
      /TDS\s*%/i.test(joined)
    );

    if (isHeaderRow) {
      headerRowIdx = i;
      // Walk each cell and assign semantic meaning
      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        if (!cell) continue;

        // Amount: "Total Voucher Assessable Amt" — pick the FIRST match so
        // "TDS Assessable Amount" (which comes later) doesn't override it
        if (/Total\s*Voucher/i.test(cell) && amtColIdx === FALLBACK_AMOUNT) {
          amtColIdx = c;
        }

        // Fallback Amount: "TDS Assessable Amount"
        if (/TDS\s*Assessable/i.test(cell)) {
          assessableAmtColIdx = c;
        }

        // TDS: "Net TDS to be Paid"
        if (/Net\s*TDS/i.test(cell)) {
          tdsColIdx = c;
        }

        // Name: "Party Name"
        if (/Party\s*Name/i.test(cell)) {
          nameColIdx = c;
        }

        // Rate: "TDS %" — but NOT "Net TDS" which also contains "TDS"
        if (/TDS\s*%/i.test(cell) && !/Net\s*TDS/i.test(cell)) {
          rateColIdx = c;
        }
      }

      // Once we find the header row, stop scanning — data rows follow
      break;
    }

    // Safety: if we've passed the header zone (>15 rows) without finding headers,
    // stop scanning to avoid misidentifying data as headers
    if (i > 15) break;
  }

  // ── Heuristics Auto-Detection Fallback ────────────────────────────────────
  // If no header row was detected, scan first 30 rows to detect columns by content
  if (headerRowIdx === -1) {
    const colTypes = {};
    const scanLimit = Math.min(rawRows.length, 30);
    for (let i = 0; i < scanLimit; i++) {
      const row = rawRows[i] || [];
      for (let c = 0; c < row.length; c++) {
        const val = str(row[c]);
        if (!val) continue;

        if (!colTypes[c]) colTypes[c] = { name: 0, amount: 0, rate: 0 };

        const num = parseFloat(val.replace(/,/g, ''));
        if (!isNaN(num)) {
          if (num > 100) {
            colTypes[c].amount++;
          } else if (num > 0 && num <= 30) {
            colTypes[c].rate++;
          }
        } else if (val.length > 5 && !/^\d/.test(val) && !/PAN|TAN|Nature|Date/i.test(val)) {
          colTypes[c].name++;
        }
      }
    }

    // Assign name column
    for (const c of Object.keys(colTypes)) {
      const idx = parseInt(c, 10);
      const counts = colTypes[c];
      if (counts.name > 5) nameColIdx = idx;
    }

    // Determine numeric columns (Amount vs TDS) by average values
    const numericCols = Object.keys(colTypes)
      .map(Number)
      .filter(c => {
        const colVals = rawRows.slice(0, 30).map(r => r ? parseFloat(str(r[c]).replace(/,/g, '')) : NaN).filter(v => !isNaN(v));
        return colVals.length > 3;
      });

    if (numericCols.length >= 2) {
      const averages = numericCols.map(c => {
        const vals = rawRows.slice(0, 30).map(r => r ? parseFloat(str(r[c]).replace(/,/g, '')) : 0).filter(v => v > 0);
        const avg = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
        return { col: c, avg };
      });
      averages.sort((a, b) => b.avg - a.avg);
      if (averages.length >= 2) {
        amtColIdx = averages[0].col; // Larger numbers represent Amount
        tdsColIdx = averages[1].col; // Smaller numbers represent TDS
      }
    }
  }

  // ── Pass 2: parse data rows ─────────────────────────────────────────────────
  const records      = [];
  let currentSection = '';
  let currentPan     = '';
  let lastName       = '';

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row) continue;

    const cells = row.map(str);
    const joined = cells.join(' ').trim();
    if (!joined) continue;

    // Section line: "Nature Of Payment : TDS on Purchase of Goods (194Q)"
    const sectionMatch = joined.match(/Nature\s+Of\s+Payment\s*:\s*.*?\((\w+)\)/i);
    if (sectionMatch) {
      currentSection = sectionMatch[1].trim();
      currentPan = '';
      lastName = '';
      continue;
    }

    // PAN line: "PAN No : ABMFM3677L"
    const panMatch = joined.match(/PAN\s+No\s*:\s*([A-Z]{5}[0-9]{4}[A-Z])/i);
    if (panMatch) {
      currentPan = panMatch[1].toUpperCase();
      lastName = '';
      continue;
    }

    // Skip header and summary rows
    if (/^\s*Total\b/i.test(joined))                                       continue;
    if (/Nature\s+Of\s+Payment/i.test(joined))                             continue;
    if (/From\s+Date|TDS\s+Report/i.test(joined))                          continue;
    if (/Total\s*Voucher|Net\s*TDS|Party\s*Name.*TDS/i.test(joined))       continue;

    // Require at least section to have been seen (PAN is optional)
    if (!currentSection) continue;

    // ── Read values by detected column positions ─────────────────────────
    let partyName = str(row[nameColIdx]);
    if (partyName && !/^\s*Total\s*$/i.test(partyName)) {
      lastName = partyName;
    } else {
      partyName = lastName;
    }
    if (!partyName || /^\s*Total\s*$/i.test(partyName)) continue;

    let amount = 0;
    if (assessableAmtColIdx >= 0 && toNum(row[assessableAmtColIdx]) > 0) {
      amount = toNum(row[assessableAmtColIdx]);
    } else {
      amount = toNum(row[amtColIdx]);
    }
    const tds    = toNum(row[tdsColIdx]);
    if (!amount && !tds) continue;

    // Rate from explicit column if present, else derive
    const rawRate = (rateColIdx >= 0 && row[rateColIdx] != null)
      ? toNum(row[rateColIdx])
      : -1;

    // ── Sanitise TDS / rate ──────────────────────────────────────────────
    const { finalTds, finalRate } = sanitiseTds(tds, amount, rawRate, currentSection);

    records.push({
      deducteeCode:          '02',
      pan:                   currentPan,
      name:                  partyName,
      middleName:            '',             // not available in this format
      lastName:              '',             // not available in this format
      address1:              '',             // not available in this format
      address2:              '',             // not available in this format
      state:                 '',             // not available in this format
      pinCode:               '',             // not available in this format
      amount,
      date:                  toDate || endOfCurrentFY(),
      section:               currentSection,
      rate:                  finalRate,
      tds:                   finalTds,
      dateOfTdsDeduction:    '',             // not available in this format
      challanDetail:         '',             // not available in this format
      dateOfFurnishingCert:  '',             // not available in this format
      reasonForNonDeduction: '',             // not available in this format
      paidByBookEntry:       '',             // not available in this format
      certificateNo197:      '',             // not available in this format
      partyReferenceNo:      '',             // not available in this format
    });
  }

  return { deductorName, tan, records };
}

// ── Sanitise TDS & Rate ───────────────────────────────────────────────────────
// Handles three corrupt scenarios found in the wild:
//   1. tds === amount    (TDS column stored the full assessable amount)
//   2. tds >= amount*0.5 (clearly wrong for any normal TDS rate ≤ 30%)
//   3. rate === 30 for a 194Q section (should be 0.1%)
//
// Resolution priority:
//   a. If rawRate is a sane decimal (0.01–30) AND doesn't trigger #3, use it
//   b. Else derive from section code
//   c. Recalculate tds = ceil(amount * rate / 100)
function sanitiseTds(tds, amount, rawRate, section) {
  let finalRate = rawRate;
  let finalTds  = tds;

  const sectionUpper = (section || '').toUpperCase();

  // Determine the correct rate
  const rateIsKnownWrong = (
    rawRate < 0 ||          // not present
    rawRate > 30 ||         // impossible TDS rate
    (rawRate === 30 && /194Q/i.test(sectionUpper))  // 30% on 194Q is wrong
  );

  if (rateIsKnownWrong) {
    finalRate = defaultRateForSection(sectionUpper);
  }

  // Check if TDS value itself is corrupt (≥ 50% of amount means rate would be ≥50%)
  const tdsCorrupt = (amount > 0 && tds > 0 && tds >= amount * 0.5);

  if (tdsCorrupt) {
    // If rate was sane all along and we just had a bad tds column, use it
    if (!rateIsKnownWrong && rawRate > 0 && rawRate <= 30) {
      finalRate = rawRate;
    }
    finalTds = Math.ceil(amount * finalRate / 100);
  }

  // Edge: tds=0 but amount exists — recalculate
  if (finalTds === 0 && amount > 0 && finalRate > 0) {
    finalTds = Math.ceil(amount * finalRate / 100);
  }

  return { finalTds, finalRate };
}

// ── Section → default TDS rate ────────────────────────────────────────────────
function defaultRateForSection(sectionUpper) {
  if (/194Q/.test(sectionUpper))        return 0.1;
  if (/194H/.test(sectionUpper))        return 2;
  if (/194T/.test(sectionUpper))        return 10;
  if (/194[A-Z]/.test(sectionUpper))    return 10;  // 194A, 194C, 194I, etc.
  if (/^194$/.test(sectionUpper))       return 10;
  return 0.1; // safe default — better to understate than overstate
}

// ── Numeric helpers ───────────────────────────────────────────────────────────

function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function normaliseDate(s) {
  const parts = s.split(/[\/\-]/);
  if (parts.length !== 3) return s;
  const [d, m, y] = parts;
  return `${d.padStart(2, '0')}/${m.padStart(2, '0')}/${y}`;
}

function endOfCurrentFY() {
  const now  = new Date();
  const year = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  return `31/03/${year + 1}`;
}

module.exports = { parseFormatA, sanitiseTds, defaultRateForSection };
