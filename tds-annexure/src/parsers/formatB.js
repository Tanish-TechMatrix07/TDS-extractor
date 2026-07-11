'use strict';

/**
 * Format B Parser — Kalrav Industries / Partner TDS PDF
 *
 * Confirmed raw PDF text structure from debug:
 *   Line: "PARTNER NAMEPAN NO.PARTNER INTEREST PARTNER REMUNERATION TOTALTDS 10%"
 *   Line: "ASHWINBHAI CHHAGANBHAI VIRADIYA (36%)AAYPV7109F6011719704669821769822"
 *   Line: "HIRENBHAI DHIRAJLAL RAIYANI (10%)BZJPR6210G1471632695717412017412"
 *   Line: "JENTIBHAI MOHANBHAI MARAKANA  (18%)ADRPM6134B1202444852316876716877"
 *   Line: "SUMITBHAI BABUBHAI TIMBADIYA (36%)AURPT0192P3824859704647953147953"
 *   Line: "TOTAL12510632695721520635152064"
 *   Line: "KALRAV INDUSTRIES"
 *   Line: "TAN NO. RKTK05989E"
 *
 * Each data line = NAME + PAN (10 chars) + 4 numbers concatenated with no separator.
 * Numbers in order: INTEREST, REMUNERATION, TOTAL, TDS
 * We need: amount = TOTAL (3rd number), tds = TDS (4th number)
 *
 * Strategy:
 *   1. Find PAN in line (regex: 5 letters + 4 digits + 1 letter)
 *   2. Name = everything before the PAN
 *   3. Numbers string = everything after the PAN
 *   4. Parse numbers from the concatenated number string
 */

const pdfParse = require('pdf-parse');

// PAN pattern: exactly 5 uppercase letters, 4 digits, 1 uppercase letter
const PAN_RE = /([A-Z]{5}[0-9]{4}[A-Z])/;

async function parseFormatB(buffer, filename) {
  const data = await pdfParse(buffer);
  const text = data.text;

  const lines = text
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  // ── Extract deductor info ──────────────────────────────────────────────────
  let deductorName = '';
  let tan = '';
  let section = '';
  let fyYear = endFYYear();

  // Check filename for FY year first (e.g. "2025-26" or "2025-2026")
  if (filename) {
    const fyInName = filename.match(/20(\d{2})[–\-](?:20)?(\d{2})/);
    if (fyInName) fyYear = 2000 + parseInt(fyInName[2], 10);
  }

  for (const line of lines) {
    // TAN: "TAN NO. RKTK05989E"
    const tanMatch = line.match(/TAN\s*(?:NO\.?)?\s*[:\.]?\s*([A-Z]{4}[0-9]{5}[A-Z])/i);
    if (tanMatch && !tan) tan = tanMatch[1].toUpperCase();

    // Section code in line
    const secMatch = line.match(/\b(194[A-Z0-9]*)\b/);
    if (secMatch && !section) section = secMatch[1];

    // FY year e.g. "2025-26"
    const fyMatch = line.match(/20(\d{2})[–\-](?:20)?(\d{2})/);
    if (fyMatch) fyYear = 2000 + parseInt(fyMatch[2], 10);

    // Company name: appears AFTER the data rows (line 11 in debug = "KALRAV INDUSTRIES")
    // It's a line with only uppercase letters and spaces, no digits, no special chars
    if (!deductorName && /^[A-Z][A-Z\s]+$/.test(line) && line.length > 3
        && !/TOTAL|PARTNER|INTEREST|REMUN|TAN|PAN/i.test(line)) {
      deductorName = line.trim();
    }
  }

  if (!section) section = '194';
  const date = `31/03/${fyYear}`;

  // ── Parse data rows ────────────────────────────────────────────────────────
  const records = [];

  for (const line of lines) {
    const panMatch = line.match(PAN_RE);
    if (!panMatch) continue;

    const pan = panMatch[1];
    const panIdx = line.indexOf(pan);

    // Name = everything before PAN
    const name = line.substring(0, panIdx).trim();

    // Numbers string = everything after PAN
    const numStr = line.substring(panIdx + pan.length).trim();

    // Skip header and total lines
    if (/TOTAL|PARTNER\s*NAME|PAN\s*NO/i.test(name)) continue;
    if (!name || name.length < 3) continue;

    // Parse the concatenated number string into individual numbers
    // The 4 numbers are: interest, remuneration, total, tds
    const nums = splitConcatenatedNumbers(numStr);

    if (nums.length < 2) continue;

    // We need TOTAL (3rd) and TDS (4th)
    // If we got 4 numbers: [interest, remuneration, total, tds]
    // If we got 3 numbers: [interest+remuneration merged?, total, tds] — use last two
    // If we got 2 numbers: [total, tds]
    let total, tds;
    if (nums.length >= 4) {
      total = nums[2];
      tds   = nums[3];
    } else if (nums.length === 3) {
      total = nums[1];
      tds   = nums[2];
    } else {
      total = nums[0];
      tds   = nums[1];
    }

    if (!total || !tds) continue;

    // Sanity check: tds should be much less than total
    if (tds >= total) continue;

    const rate = calcRate(tds, total);

    records.push({
      deducteeCode: '02',
      pan,
      name,
      amount: total,
      date,
      section,
      rate,
      tds,
    });
  }

  return { deductorName, tan, records };
}

// ── Split a concatenated number string into 4 numbers ────────────────────────
// Confirmed example: "6011719704669821769822"
//   → interest=601171, remuneration=97046, total=698217, tds=69822
// Pattern: all 4 are integers. TDS ≈ 10% of TOTAL. TOTAL = interest + remuneration.
//
// Strategy: brute-force all split combinations (i1, i2, i3) where
//   n1 = digits[0..i1], n2 = digits[i1..i2], n3 = digits[i2..i3], n4 = digits[i3..]
//   and check: n3 ≈ n1+n2 (total = interest+remuneration) AND n4 ≈ 10%*n3
function splitConcatenatedNumbers(str) {
  const digits = str.replace(/[^\d]/g, '');
  const len = digits.length;
  if (!len) return [];

  // If there are natural separators (spaces etc), use them directly
  const parts = str.match(/\d+/g);
  if (parts && parts.length >= 2) {
    const nums = parts.map(Number);
    if (nums.length >= 4) return nums;
    if (nums.length === 2) return nums; // total + tds
  }

  // Brute-force split into 4 numbers
  // Each number is at least 1 digit and at most 10 digits
  const best = [];
  let bestScore = Infinity;

  for (let i1 = 1; i1 < len - 2; i1++) {
    for (let i2 = i1 + 1; i2 < len - 1; i2++) {
      for (let i3 = i2 + 1; i3 < len; i3++) {
        const n1 = parseInt(digits.slice(0, i1), 10);
        const n2 = parseInt(digits.slice(i1, i2), 10);
        const n3 = parseInt(digits.slice(i2, i3), 10);
        const n4 = parseInt(digits.slice(i3), 10);

        if (!n1 || !n2 || !n3 || !n4) continue;

        // Constraint 1: n3 should equal or be close to n1+n2 (total = int+remun)
        const sumDiff = Math.abs(n3 - (n1 + n2));
        const sumRatio = sumDiff / n3;

        // Constraint 2: n4/n3 should be close to a standard rate (0.1, 0.02, etc.)
        const impliedRate = (n4 / n3) * 100;
        const standardRates = [2, 5, 10, 20, 30];
        const rateDiff = Math.min(...standardRates.map(r => Math.abs(impliedRate - r)));

        // Score: lower is better. Weight sum constraint heavily.
        if (sumRatio > 0.05) continue; // total must be within 5% of sum
        if (rateDiff > 2) continue;    // rate must be within 2% of a standard rate

        const score = sumRatio + (rateDiff / 10);
        if (score < bestScore) {
          bestScore = score;
          best.length = 0;
          best.push(n1, n2, n3, n4);
        }
      }
    }
  }

  if (best.length === 4) return best;

  // Fallback: try just finding total+tds (2-number split)
  // total ≈ tds / 0.10, tds is last 4-6 digits
  for (let tdsLen = 4; tdsLen <= 7 && tdsLen < len; tdsLen++) {
    const tds   = parseInt(digits.slice(len - tdsLen), 10);
    const total = parseInt(digits.slice(0, len - tdsLen), 10);
    if (!tds || !total) continue;
    const rate = (tds / total) * 100;
    if (rate >= 1 && rate <= 35) return [total, tds];
  }

  return parts ? parts.map(Number).filter(n => n > 0) : [];
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function calcRate(tds, amount) {
  if (!amount || amount === 0) return 0;
  const raw = (tds / amount) * 100;
  const standard = [0.01, 0.1, 0.5, 1, 1.5, 2, 5, 10, 20, 30];
  let closest = standard[0];
  let minDiff = Math.abs(raw - standard[0]);
  for (const s of standard) {
    const diff = Math.abs(raw - s);
    if (diff < minDiff) { minDiff = diff; closest = s; }
  }
  return closest;
}

function endFYYear() {
  // Return the end year of the current financial year
  // FY runs April–March, so if we're in Jan 2026, FY end year is 2026
  const now = new Date();
  return now.getMonth() >= 3 ? now.getFullYear() + 1 : now.getFullYear();
}

module.exports = { parseFormatB };
