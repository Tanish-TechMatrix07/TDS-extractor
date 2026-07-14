'use strict';

/**
 * Format C Parser — Annexure-style XLS/XLSX
 *
 * Handles two confirmed structures:
 *
 *   Shree Radhamadhav variant (8-column annexure):
 *   [0] ["Party Name :", null, "SHREE RADHAMADHAV AGRI ENTERPRISES LLP", null, null,
 *         "TAN :", "RKTS17423A", ...]
 *   [1] ["Deductee code","PAN of deductee","First Name","Amount of payment (Rs.)",
 *         "Date on which amount paid/credited","Section code",
 *         "Rate at which tax deducted","TDS(Rs.)", ...]
 *   [2] ["02","ABXFA5405F","AARAV ENTERPRISE","19974","31/03/2026","194H","2","400", ...]
 *
 *   Re-processed annexure variant (same column layout, but values can be corrupt):
 *   [9] ["02","ABMFM3677L","MADHUR GANESH TRADING CO.",1724450.5,"31/01/2026","194Q",30,1724450.5]
 *   ← rate=30 and tds=amount are both wrong for 194Q (should be rate=0.1, tds=1724)
 *
 * Strategy:
 *   1. Row 0: extract Party Name / TAN from free-text scan
 *   2. Scan up to row 10 for a header row containing "Deductee code" or
 *      "PAN of deductee" or "Amount of payment" — map columns semantically
 *   3. Parse data rows using mapped columns
 *   4. Run every record through sanitiseTds() to fix corrupt rate/tds pairs
 *
 * Column semantic keywords (case-insensitive, partial match):
 *   deducteeCode → "Deductee code"
 *   pan          → "PAN of deductee"
 *   name         → "First Name" | "Party Name" | "Name"
 *   amount       → "Amount of payment" | "Amount"
 *   date         → "Date on which" | "Date"
 *   section      → "Section code" | "Section"
 *   rate         → "Rate at which" | "Rate"
 *   tds          → "TDS" + "Rs" (e.g. "TDS(Rs.)", "TDS (Rs.)", "TDS Rs")
 */

const XLSX = require('xlsx');
const { sanitiseTds } = require('./formatA');

function parseFormatC(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', raw: true });

  // Prefer sheet named "Annexure" (case-insensitive), else use first sheet
  let sheet = null;
  for (const name of workbook.SheetNames) {
    if (/annexure/i.test(name)) { sheet = workbook.Sheets[name]; break; }
  }
  if (!sheet) sheet = workbook.Sheets[workbook.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    blankrows: false,
    raw: true,
  });

  const str = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  // ── Step 1: Extract deductor name + TAN ────────────────────────────────────
  // Scan first 5 rows for "Party Name : ..." and "TAN : ..." patterns.
  // These may appear on row 0 (Shree format) or not at all.
  let deductorName = '';
  let tan          = '';

  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const cells  = (rows[i] || []).map(str);
    const joined = cells.join(' ');

    if (!deductorName) {
      const m = joined.match(/Party\s*Name\s*[:\-]\s*(.+?)(?:\s{2,}TAN\s*[:\-]|$)/i);
      if (m) deductorName = m[1].trim();
    }
    if (!tan) {
      const m = joined.match(/TAN\s*[:\-]\s*([A-Z]{4}[0-9]{5}[A-Z])/i);
      if (m) tan = m[1].toUpperCase();
    }
    if (deductorName && tan) break;
  }

  // ── Step 2: Find the column header row ────────────────────────────────────
  // Default positions (match Shree/standard annexure layout)
  let colDeductee = 0;
  let colPan      = 1;
  let colName     = 2;
  let colAmount   = 3;
  let colDate     = 4;
  let colSection  = 5;
  let colRate     = 6;
  let colTds      = 7;
  let dataStartIdx = 2; // rows before this are header/meta

  // A header row must contain at least two of these strong markers
  const HEADER_MARKERS = [
    /Deductee\s*code/i,
    /PAN\s*of\s*deductee/i,
    /Amount\s*of\s*payment/i,
    /Section\s*code/i,
    /Rate\s*at\s*which/i,
    /TDS\s*[\(\[]?Rs/i,
  ];

  let headerFound = false;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const cells  = (rows[i] || []).map(str);
    const joined = cells.join(' ');

    const matchCount = HEADER_MARKERS.filter(re => re.test(joined)).length;
    if (matchCount >= 2) {
      headerFound = true;
      dataStartIdx = i + 1;

      // Map each cell to a semantic column
      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        if (!cell) continue;

        if (/Deductee\s*code/i.test(cell))                         colDeductee = c;
        if (/PAN\s*of\s*deductee/i.test(cell))                     colPan      = c;
        if (/First\s*Name/i.test(cell))                            colName     = c;
        else if (/\bName\b/i.test(cell) && !/Party/i.test(cell) && !/PAN/i.test(cell))
                                                                    colName     = c;
        if (/Amount\s*of\s*payment/i.test(cell))                   colAmount   = c;
        else if (/\bAmount\b/i.test(cell) && !/TDS/i.test(cell) && !/Date/i.test(cell))  colAmount   = c;
        if (/Date\s*on\s*which/i.test(cell))                       colDate     = c;
        else if (/\bDate\b/i.test(cell))                           colDate     = c;
        if (/Section\s*code/i.test(cell))                          colSection  = c;
        else if (/\bSection\b/i.test(cell))                        colSection  = c;
        if (/Rate\s*at\s*which/i.test(cell))                       colRate     = c;
        else if (/\bRate\b/i.test(cell) && !/TDS/i.test(cell))    colRate     = c;
        if (/TDS\s*[\(\[]?Rs/i.test(cell))                         colTds      = c;
        else if (/\bTDS\b/i.test(cell) && !/Rate/i.test(cell) && !/Section/i.test(cell))
                                                                    colTds      = c;
      }
      break;
    }
  }

  // ── Heuristics Auto-Detection Fallback ────────────────────────────────────
  // If no clear headers were found, scan first 30 rows to detect columns by content patterns
  if (!headerFound) {
    dataStartIdx = 0; // start parsing from the very beginning of the sheet
    const colTypes = {};

    const scanLimit = Math.min(rows.length, 30);
    for (let i = 0; i < scanLimit; i++) {
      const row = rows[i] || [];
      for (let c = 0; c < row.length; c++) {
        const val = str(row[c]);
        if (!val) continue;

        if (!colTypes[c]) colTypes[c] = { pan: 0, date: 0, section: 0, name: 0 };

        // 1. PAN pattern (5 letters, 4 digits, 1 letter)
        if (/^[A-Z]{5}[0-9]{4}[A-Z]$/i.test(val)) {
          colTypes[c].pan++;
        }
        // 2. Date pattern
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(val) || (/^\d{5}$/.test(val) && parseInt(val, 10) > 40000 && parseInt(val, 10) < 60000)) {
          colTypes[c].date++;
        }
        // 3. Section pattern (like 194, 194Q, 194H)
        if (/^194[A-Z0-9]*$/i.test(val)) {
          colTypes[c].section++;
        }
        // 4. Name pattern (alphabetical strings, excluding short codes and known markers)
        if (val.length > 4 && !/^\d/.test(val) && !/^[A-Z]{5}[0-9]{4}[A-Z]$/i.test(val) && !/194/i.test(val) && !/[\/\-]/.test(val)) {
          colTypes[c].name++;
        }
      }
    }

    // Assign columns based on maximum occurrences of type markers
    for (const c of Object.keys(colTypes)) {
      const idx = parseInt(c, 10);
      const counts = colTypes[c];
      if (counts.pan > 2)      colPan = idx;
      if (counts.date > 2)     colDate = idx;
      if (counts.section > 2)  colSection = idx;
      if (counts.name > 5 && colName === 2) colName = idx;
    }
  }

  // ── Step 3: Parse data rows ────────────────────────────────────────────────
  const records = [];

  for (let i = dataStartIdx; i < rows.length; i++) {
    const row   = rows[i] || [];
    const cells = row.map(str);

    // Skip fully blank rows
    if (cells.every(c => !c)) continue;

    const joined = cells.join(' ');

    // Skip repeated header rows or summary rows
    if (/Deductee\s*code/i.test(joined) && /PAN\s*of\s*deductee/i.test(joined)) continue;
    if (/Party\s*Name\s*:/i.test(joined))  continue;

    const pan     = cells[colPan]     || '';
    const name    = cells[colName]    || '';
    const section = cells[colSection] || '';

    // Need at least PAN or name, and a section code
    if ((!pan && !name) || !section) continue;

    const deducteeCode = cells[colDeductee] || '02';
    const amount       = parseNum(row[colAmount]);
    const date         = normaliseDate(str(row[colDate]));
    const rawRate      = parseNum(row[colRate]);
    const tds          = parseNum(row[colTds]);

    // Need at least one financial value
    if (!amount && !tds) continue;

    // ── Sanitise TDS / rate using shared helper ──────────────────────────
    const { finalTds, finalRate } = sanitiseTds(tds, amount, rawRate, section);

    records.push({
      deducteeCode:          deducteeCode || '02',
      pan,
      name,
      middleName:            '',             // not available in this format
      lastName:              '',             // not available in this format
      address1:              '',             // not available in this format
      address2:              '',             // not available in this format
      state:                 '',             // not available in this format
      pinCode:               '',             // not available in this format
      amount,
      date,
      section,
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

  // ── Step 4: Parse Challan sheet if present ────────────────────────────────
  const challanRecords = parseChallanSheet(workbook);

  return { deductorName, tan, records, challanRecords };
}

// ── Challan Sheet Parser ──────────────────────────────────────────────────────
// Extracts challan data from the "Challan" sheet (if it exists).
// The Challan sheet typically has:
//   Row 0: Party Name / TAN
//   Row 3: Column headers
//   Row 4+: Data rows
//
// Columns (0-indexed):
//   0:  S. No.
//   1:  Section Code
//   2:  TDS(Rs.)
//   3:  Surcharge (Rs.)
//   4:  Education Cess (Rs.)
//   5:  Higher Education Cess
//   6:  Interest (Rs.)
//   7:  Other (Rs.)
//   8:  Fees Amount (Rs.)
//   9:  Cheque/DD No.
//   10: BSR Code
//   11: Date on which Tax Deposited
//   12: Transfer Voucher/ Challan Serial No.
//   13: Whether TDS deposited by book entry?
//   14: Minor Head
function parseChallanSheet(workbook) {
  // Find the "Challan" sheet (case-insensitive)
  let sheet = null;
  for (const name of workbook.SheetNames) {
    if (/^challan$/i.test(name)) { sheet = workbook.Sheets[name]; break; }
  }
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    blankrows: false,
    raw: true,
  });

  const str = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  // Find header row containing "S. No." or "Section Code"
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const cells = (rows[i] || []).map(str);
    const joined = cells.join(' ');
    if (/S\.\s*No|Section\s*Code|BSR\s*Code/i.test(joined)) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) return [];

  // Map column indices
  const colMap = {
    srNo:       0,  // S. No.
    section:    1,  // Section Code
    tds:        2,  // TDS(Rs.)
    surcharge:  3,  // Surcharge (Rs.)
    eduCess:    4,  // Education Cess (Rs.)
    higherEduCess: 5,  // Higher Education Cess
    interest:   6,  // Interest (Rs.)
    other:      7,  // Other (Rs.)
    feesAmount: 8,  // Fees Amount (Rs.)
    chequeNo:   9,  // Cheque/DD No.
    bsrCode:    10, // BSR Code
    depositDate: 11, // Date on which Tax Deposited
    challanNo:  12, // Transfer Voucher/ Challan Serial No.
    bookEntry:  13, // Whether TDS deposited by book entry?
    minorHead:  14, // Minor Head
  };

  // Try to map columns semantically from header text
  const headerCells = (rows[headerRowIdx] || []).map(str);
  for (let c = 0; c < headerCells.length; c++) {
    const cell = headerCells[c];
    if (!cell) continue;

    if (/Section\s*Code/i.test(cell))                colMap.section = c;
    else if (/^TDS/i.test(cell) && !/Surcharge/i.test(cell)) colMap.tds = c;
    if (/Surcharge/i.test(cell))                     colMap.surcharge = c;
    if (/Education\s*Cess/i.test(cell))              colMap.eduCess = c;
    if (/Higher\s*Education/i.test(cell))             colMap.higherEduCess = c;
    if (/^Interest/i.test(cell))                     colMap.interest = c;
    if (/^Other/i.test(cell))                        colMap.other = c;
    if (/Fees/i.test(cell))                          colMap.feesAmount = c;
    if (/Cheque|DD\s*No/i.test(cell))                colMap.chequeNo = c;
    if (/BSR/i.test(cell))                           colMap.bsrCode = c;
    if (/Date.*Deposited/i.test(cell))               colMap.depositDate = c;
    if (/Challan.*Serial|Transfer.*Voucher/i.test(cell)) colMap.challanNo = c;
    if (/book\s*entry/i.test(cell))                  colMap.bookEntry = c;
    if (/Minor\s*Head/i.test(cell))                  colMap.minorHead = c;
  }

  const challanRecords = [];

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const cells = row.map(str);

    // Skip empty rows
    if (cells.every(c => !c)) continue;

    const joined = cells.join(' ');

    // Skip summary / total rows
    if (/^Total|Grand|Sum/i.test(joined)) continue;

    const srNo   = cells[colMap.srNo] || '';
    const section = cells[colMap.section] || '';
    const tds    = parseNum(row[colMap.tds]);

    // Need at least a Section code or TDS value
    if (!section && !tds) continue;

    // Skip the header row itself if it appears again
    if (/S\.\s*No/i.test(srNo) && /Section\s*Code/i.test(section)) continue;

    challanRecords.push({
      srNo,
      section,
      tds:        parseNum(row[colMap.tds]),
      surcharge:  parseNum(row[colMap.surcharge]),
      eduCess:    parseNum(row[colMap.eduCess]),
      higherEduCess: parseNum(row[colMap.higherEduCess]),
      interest:   parseNum(row[colMap.interest]),
      other:      parseNum(row[colMap.other]),
      feesAmount: parseNum(row[colMap.feesAmount]),
      chequeNo:   cells[colMap.chequeNo] || '',
      bsrCode:    cells[colMap.bsrCode] || '',
      depositDate: normaliseDate(str(row[colMap.depositDate])),
      challanNo:  cells[colMap.challanNo] || '',
      bookEntry:  cells[colMap.bookEntry] || '',
      minorHead:  cells[colMap.minorHead] || '',
    });
  }

  return challanRecords;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  const n = parseFloat(String(val).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function normaliseDate(val) {
  if (!val) return '';

  // Already DD/MM/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(val)) return val;

  // Excel serial number (numeric string like "46085")
  const serial = parseInt(val, 10);
  if (!isNaN(serial) && serial > 40000 && serial < 60000) {
    try {
      const d = XLSX.SSF.parse_date_code(serial);
      if (d) {
        return `${String(d.d).padStart(2, '0')}/${String(d.m).padStart(2, '0')}/${d.y}`;
      }
    } catch (e) { /* fall through */ }
  }

  // ISO or any JS-parseable date string
  const dt = new Date(val);
  if (!isNaN(dt.getTime())) {
    return `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}/${dt.getFullYear()}`;
  }

  return val; // return as-is if nothing worked
}

module.exports = { parseFormatC };
