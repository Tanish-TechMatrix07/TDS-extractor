'use strict';

/**
 * Format B Parser — Kalrav Industries / Partner TDS PDF & TDS Report PDF
 *
 * This parser runs pdf-parse in an isolated child process to:
 *   1. Avoid event loop blocking for CPU-heavy PDF decoding.
 *   2. Prevent library conflict/global state contamination between older pdfjs-dist and SheetJS (xlsx).
 *
 * Supports two layouts:
 *   - Layout A (TDS Report style PDFs: "JAN.PDF", "FEB.PDF")
 *   - Layout B (Partner Remuneration PDFs: Kalrav Industries partner table)
 */

const { spawn } = require('child_process');

// PAN pattern: exactly 5 uppercase letters, 4 digits, 1 uppercase letter
const PAN_RE = /([A-Z]{5}[0-9]{4}[A-Z])/;

// ── Helper: Run pdf-parse in an isolated node process ────────────────────────
function extractPdfTextWithChildProcess(buffer) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [
      '-e',
      `
        const pdf = require('pdf-parse');
        const chunks = [];
        process.stdin.on('data', chunk => chunks.push(chunk));
        process.stdin.on('end', async () => {
          try {
            const data = await pdf(Buffer.concat(chunks));
            process.stdout.write('__JSON_START__' + JSON.stringify({ text: data.text }) + '__JSON_END__');
          } catch (err) {
            process.stdout.write('__JSON_START__' + JSON.stringify({ error: err.message }) + '__JSON_END__');
          }
        });
      `
    ]);

    const stdoutChunks = [];
    const stderrChunks = [];

    child.stdout.on('data', chunk => stdoutChunks.push(chunk));
    child.stderr.on('data', chunk => stderrChunks.push(chunk));

    child.on('close', code => {
      if (code !== 0) {
        const stderrMsg = Buffer.concat(stderrChunks).toString();
        return reject(new Error(`PDF worker exited with code ${code}: ${stderrMsg}`));
      }
      try {
        const stdoutStr = Buffer.concat(stdoutChunks).toString();
        const startIdx = stdoutStr.indexOf('__JSON_START__');
        const endIdx = stdoutStr.indexOf('__JSON_END__');
        if (startIdx === -1 || endIdx === -1) {
          return reject(new Error(`PDF worker output is missing JSON boundaries: ${stdoutStr}`));
        }
        const jsonStr = stdoutStr.substring(startIdx + '__JSON_START__'.length, endIdx);
        const output = JSON.parse(jsonStr);
        if (output.error) {
          return reject(new Error(output.error));
        }
        resolve(output.text);
      } catch (err) {
        reject(new Error(`Failed to parse PDF worker output: ${err.message}`));
      }
    });

    child.stdin.write(buffer);
    child.stdin.end();
  });
}

// ── Main Entry Point ────────────────────────────────────────────────────────
async function parseFormatB(buffer, filename) {
  const text = await extractPdfTextWithChildProcess(buffer);

  const lines = text
    .split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  // Detect which layout is in the PDF:
  // Layout A has "TDS Report" or "Nature Of Payment"
  const isLayoutA = lines.some(l => /TDS\s*Report|Nature\s*Of\s*Payment/i.test(l));

  if (isLayoutA) {
    return parsePdfLayoutA(lines, filename);
  } else {
    return parsePdfLayoutB(lines, filename);
  }
}

// ── Layout A Parser (TDS Report PDFs) ────────────────────────────────────────
function parsePdfLayoutA(lines, filename) {
  let deductorName = '';
  let tan          = '';
  let toDate       = '';

  for (const line of lines) {
    // Deductor name: first line that isn't a report header or digit
    if (!deductorName) {
      const isHeaderLine = /TDS\s*Report|Page|From\s*Date|Party\s*Name|Nature\s*Of\s*Payment|PAN\s*No/i.test(line);
      if (!isHeaderLine && line.length > 3 && !/\d/.test(line)) {
        deductorName = line.trim();
      }
    }

    // TAN
    if (!tan) {
      const m = line.match(/\b([A-Z]{4}[0-9]{5}[A-Z])\b/);
      if (m) tan = m[1].toUpperCase();
    }

    // To Date
    if (!toDate) {
      const m = line.match(/\bTo\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})/i);
      if (m) {
        const parts = m[1].split(/[\/\-]/);
        if (parts.length === 3) {
          toDate = `${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[2]}`;
        }
      }
    }
  }

  const records = [];
  let currentSection = '';
  let currentPan     = '';

  for (const line of lines) {
    // Section line
    const sectionMatch = line.match(/Nature\s+Of\s+Payment\s*:\s*.*?\((\w+)\)/i);
    if (sectionMatch) {
      currentSection = sectionMatch[1].trim();
      currentPan = '';
      continue;
    }

    // PAN line
    const panMatch = line.match(/PAN\s+No\s*:\s*([A-Z]{5}[0-9]{4}[A-Z])/i);
    if (panMatch) {
      currentPan = panMatch[1].toUpperCase();
      continue;
    }

    // Skip headers and summary totals
    if (/^\s*Total/i.test(line)) continue;
    if (/Nature\s+Of\s+Payment/i.test(line)) continue;
    if (/From\s+Date|TDS\s+Report/i.test(line)) continue;
    if (/Party\s*Name|TDS\s*Reason/i.test(line)) continue;

    if (!currentPan || !currentSection) continue;

    // Match name followed by decimal numbers (exactly 2 decimals each)
    const match = line.match(/^(.+?)((\d+\.\d{2})+)$/);
    if (match) {
      const partyName = match[1].trim();
      if (/^\s*Total\s*$/i.test(partyName)) continue;

      const nums = match[2].match(/\d+\.\d{2}/g);
      if (!nums || nums.length < 4) continue;

      const amount = parseFloat(nums[0]);
      const rate   = parseFloat(nums[2]);
      const tds    = parseFloat(nums[3]);

      // Sanitise TDS / rate
      const { finalTds, finalRate } = require('./formatA').sanitiseTds(tds, amount, rate, currentSection);

      records.push({
        deducteeCode: '02',
        pan:     currentPan,
        name:    partyName,
        amount,
        date:    toDate || endOfCurrentFY(),
        section: currentSection,
        rate:    finalRate,
        tds:     finalTds,
      });
    }
  }

  if (!deductorName && lines.length > 0) {
    deductorName = lines[0].trim();
  }

  return { deductorName, tan, records };
}

// ── Layout B Parser (Partner Remuneration PDFs) ──────────────────────────────
function parsePdfLayoutB(lines, filename) {
  let deductorName = '';
  let tan = '';
  let section = '';
  let fyYear = endFYYear();

  if (filename) {
    const fyInName = filename.match(/20(\d{2})[–\-](?:20)?(\d{2})/);
    if (fyInName) fyYear = 2000 + parseInt(fyInName[2], 10);
  }

  for (const line of lines) {
    const tanMatch = line.match(/TAN\s*(?:NO\.?)?\s*[:\.]?\s*([A-Z]{4}[0-9]{5}[A-Z])/i);
    if (tanMatch && !tan) tan = tanMatch[1].toUpperCase();

    const secMatch = line.match(/\b(194[A-Z0-9]*)\b/);
    if (secMatch && !section) section = secMatch[1];

    const fyMatch = line.match(/20(\d{2})[–\-](?:20)?(\d{2})/);
    if (fyMatch) fyYear = 2000 + parseInt(fyMatch[2], 10);

    if (!deductorName && /^[A-Z][A-Z\s]+$/.test(line) && line.length > 3
        && !/TOTAL|PARTNER|INTEREST|REMUN|TAN|PAN/i.test(line)) {
      deductorName = line.trim();
    }
  }

  if (!section) section = '194';
  const date = `31/03/${fyYear}`;

  const records = [];

  for (const line of lines) {
    const panMatch = line.match(PAN_RE);
    if (!panMatch) continue;

    const pan = panMatch[1];
    const panIdx = line.indexOf(pan);

    const name = line.substring(0, panIdx).trim();
    const numStr = line.substring(panIdx + pan.length).trim();

    if (/TOTAL|PARTNER\s*NAME|PAN\s*NO/i.test(name)) continue;
    if (!name || name.length < 3) continue;

    const nums = splitConcatenatedNumbers(numStr);

    if (nums.length < 2) continue;

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

// ── Concatenated Numbers Splitter ────────────────────────────────────────────
function splitConcatenatedNumbers(str) {
  const digits = str.replace(/[^\d]/g, '');
  const len = digits.length;
  if (!len) return [];

  const parts = str.match(/\d+/g);
  if (parts && parts.length >= 2) {
    const nums = parts.map(Number);
    if (nums.length >= 4) return nums;
    if (nums.length === 2) return nums;
  }

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

        const sumDiff = Math.abs(n3 - (n1 + n2));
        const sumRatio = sumDiff / n3;

        const impliedRate = (n4 / n3) * 100;
        const standardRates = [2, 5, 10, 20, 30];
        const rateDiff = Math.min(...standardRates.map(r => Math.abs(impliedRate - r)));

        if (sumRatio > 0.05) continue;
        if (rateDiff > 2) continue;

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

  for (let tdsLen = 4; tdsLen <= 7 && tdsLen < len; tdsLen++) {
    const tds   = parseInt(digits.slice(len - tdsLen), 10);
    const total = parseInt(digits.slice(0, len - tdsLen), 10);
    if (!tds || !total) continue;
    const rate = (tds / total) * 100;
    if (rate >= 1 && rate <= 35) return [total, tds];
  }

  return parts ? parts.map(Number).filter(n => n > 0) : [];
}

// ── Helpers ──────────────────────────────────────────────────────────────────
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
  const now = new Date();
  return now.getMonth() >= 3 ? now.getFullYear() + 1 : now.getFullYear();
}

function endOfCurrentFY() {
  const now  = new Date();
  const year = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  return `31/03/${year + 1}`;
}

module.exports = { parseFormatB };
