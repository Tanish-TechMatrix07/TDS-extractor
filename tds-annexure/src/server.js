'use strict';

const express  = require('express');
const multer   = require('multer');
const path     = require('path');
const crypto   = require('crypto');
const cors     = require('cors');

const { detectFormat }     = require('./parsers/detector');
const { parseFormatA }     = require('./parsers/formatA');
const { parseFormatB }     = require('./parsers/formatB');
const { parseFormatC }     = require('./parsers/formatC');
const { generateAnnexure } = require('./generator');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

// ── Multer — in-memory, accept any field name, up to 20 files ────────────────
// Using .any() so it accepts both "file" (old) and "files" (new) field names
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024, files: 20 },
});

// ── Helper: parse a single file buffer into {deductorName, tan, records} ─────
async function parseFile(buffer, filename) {
  const format = detectFormat(buffer, filename);
  if (format === 'UNKNOWN') {
    throw new Error(`Cannot recognise format of "${filename}". Use XLS, XLSX, or PDF.`);
  }
  let parsed;
  if (format === 'A') {
    parsed = parseFormatA(buffer);
  } else if (format === 'B' || format === 'B_EXCEL') {
    parsed = await parseFormatB(buffer, filename);
  } else {
    parsed = parseFormatC(buffer);
  }
  parsed._format = format;
  return parsed;
}

// ── POST /api/convert  (single OR multiple files) ────────────────────────────
// Field name: "files" (array) — also accepts legacy "file" (single)
// Returns: one merged Annexure Excel
app.post('/api/convert', upload.any(), async (req, res) => {
  try {
    // Collect files from any field name
    const uploadedFiles = (req.files || []).filter(f => {
      const ext = path.extname(f.originalname).toLowerCase();
      return ['.xls', '.xlsx', '.pdf'].includes(ext);
    });

    const rejected = (req.files || []).filter(f => {
      const ext = path.extname(f.originalname).toLowerCase();
      return !['.xls', '.xlsx', '.pdf'].includes(ext);
    });

    if (uploadedFiles.length === 0) {
      const badNames = rejected.map(f => f.originalname).join(', ');
      return res.status(400).json({
        error: rejected.length
          ? `Unsupported file type(s): ${badNames}. Please upload XLS, XLSX, or PDF.`
          : 'No files uploaded.',
      });
    }

    // ── Deduplicate files by content hash ──────────────────────────────────
    // Same file uploaded twice (even with different names) produces the same
    // MD5 hash — drop the duplicate before parsing.
    const seenHashes = new Set();
    const uniqueFiles = [];
    for (const file of uploadedFiles) {
      const hash = crypto.createHash('md5').update(file.buffer).digest('hex');
      if (!seenHashes.has(hash)) {
        seenHashes.add(hash);
        uniqueFiles.push(file);
      } else {
        console.log(`[dedup] Skipping duplicate file: ${file.originalname} (hash ${hash})`);
      }
    }

    // ── Parse every unique file ─────────────────────────────────────────────
    const results = [];
    const errors  = [];

    for (const file of uniqueFiles) {
      try {
        const parsed = await parseFile(file.buffer, file.originalname);
        results.push({ filename: file.originalname, parsed });
      } catch (err) {
        errors.push(`${file.originalname}: ${err.message}`);
      }
    }

    if (results.length === 0) {
      return res.status(422).json({
        error: 'No records could be extracted from the uploaded files.',
        details: errors,
      });
    }

    // ── Merge + deduplicate records ─────────────────────────────────────────
    // Deductor name + TAN: use from first successfully parsed file
    const primaryParsed = results[0].parsed;
    const deductorName  = primaryParsed.deductorName || '';
    const tan           = primaryParsed.tan || '';

    // Combine all records then deduplicate by content fingerprint.
    // Fingerprint = PAN + section + amount + TDS
    // This catches the same record appearing in multiple files (e.g. same
    // data exported in two different formats, or same file uploaded twice
    // with different names but identical content that slipped past hash check).
    const seenRecords = new Set();
    const allRecords  = [];

    for (const r of results) {
      for (const rec of (r.parsed.records || [])) {
        // Normalise each field to avoid false mismatches from whitespace/case
        const pan     = (rec.pan     || '').toUpperCase().trim();
        const section = (rec.section || '').toUpperCase().trim();
        const amount  = String(rec.amount || 0);
        const tds     = String(rec.tds    || 0);
        const date    = (rec.date || '').trim();

        const fingerprint = `${pan}|${section}|${amount}|${tds}|${date}`;

        if (!seenRecords.has(fingerprint)) {
          seenRecords.add(fingerprint);
          allRecords.push(rec);
        } else {
          console.log(`[dedup] Skipping duplicate record: ${pan} ${section} amt=${amount} tds=${tds}`);
        }
      }
    }

    if (allRecords.length === 0 && !results.some(r => r.parsed.challanRecords?.length > 0)) {
      return res.status(422).json({
        error: 'Files were parsed but no TDS records were found.',
        details: errors,
      });
    }

    // Collect challan records from all files (dedup by section+tds+bsr+challanNo)
    const seenChallan = new Set();
    const allChallanRecords = [];
    for (const r of results) {
      for (const rec of (r.parsed.challanRecords || [])) {
        const section   = (rec.section   || '').toUpperCase().trim();
        const tds       = String(rec.tds  || 0);
        const bsr       = (rec.bsrCode   || '').trim();
        const challanNo = (rec.challanNo || '').trim();
        const fingerprint = `${section}|${tds}|${bsr}|${challanNo}`;
        if (!seenChallan.has(fingerprint)) {
          seenChallan.add(fingerprint);
          allChallanRecords.push(rec);
        }
      }
    }

    // ── Generate Annexure ────────────────────────────────────────────────────
    const baseName = uniqueFiles.length === 1
      ? path.basename(uniqueFiles[0].originalname, path.extname(uniqueFiles[0].originalname))
      : 'TDS_Combined';
    const excelBuffer = await generateAnnexure(deductorName, tan, allRecords, baseName, allChallanRecords);

    // ── Output filename ──────────────────────────────────────────────────────
    const outName = `${baseName}_Annexure.xls`;

    // ── Send ─────────────────────────────────────────────────────────────────
    res.setHeader('Content-Type', 'application/vnd.ms-excel');
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('X-Records-Count',  String(allRecords.length));
    res.setHeader('X-Files-Count',    String(uniqueFiles.length));
    res.setHeader('X-Deductor-Name',  encodeURIComponent(deductorName));
    res.setHeader('X-Format-Detected', results.map(r => r.parsed._format).join(','));
    if (errors.length > 0) {
      res.setHeader('X-Parse-Warnings', encodeURIComponent(errors.join(' | ')));
    }
    res.send(excelBuffer);

  } catch (err) {
    console.error('[/api/convert] Error:', err);
    res.status(500).json({ error: 'An error occurred while processing.', detail: err.message });
  }
});

// ── GET /api/health ───────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ── Fallback to frontend ──────────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// ── Multer + global error handler ────────────────────────────────────────────
app.use((err, req, res, next) => {
  console.error('[Error]', err.message);
  // Always respond with JSON so the browser gets a readable error
  const status = err.status || err.statusCode || 500;
  res.status(status).json({ error: err.message || 'Internal server error' });
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`TDS Annexure Generator running at http://localhost:${PORT}`);
});

module.exports = app;
