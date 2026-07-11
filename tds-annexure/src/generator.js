'use strict';

/**
 * Output Generator
 * Takes the normalised internal record array and produces an Annexure Excel file
 * that matches the Shree Radhamadhav sample output exactly.
 *
 * Output sheet structure:
 *   Sheet name : Annexure
 *   Row 1      : Party Name : <deductorName>   [gap cells]   TAN : <tan>
 *   Row 2      : Column headers (8 columns)
 *   Row 3+     : Data rows, one per record
 *
 * Column layout (A–H):
 *   A  Deductee code
 *   B  PAN of deductee
 *   C  First Name
 *   D  Amount of payment (Rs.)
 *   E  Date on which amount paid/credited
 *   F  Section code
 *   G  Rate at which tax deducted
 *   H  TDS (Rs.)
 *
 * Styling (matching sample):
 *   - Header row: bold, light-blue fill, all borders
 *   - Data rows: all borders, alternating white background
 *   - Numeric columns (D, G, H): right-aligned, number format
 *   - Date column (E): centre-aligned
 *   - Column widths set to sensible values
 */

const ExcelJS = require('exceljs');

async function generateAnnexure(deductorName, tan, records) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'TDS Annexure Generator';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Annexure');

  // ── Column widths ──────────────────────────────────────────────────────────
  sheet.columns = [
    { key: 'deducteeCode', width: 16 },  // A
    { key: 'pan',          width: 18 },  // B
    { key: 'name',         width: 42 },  // C
    { key: 'amount',       width: 24 },  // D
    { key: 'date',         width: 28 },  // E
    { key: 'section',      width: 14 },  // F
    { key: 'rate',         width: 22 },  // G
    { key: 'tds',          width: 12 },  // H
  ];

  // ── Row 1 — Deductor header ─────────────────────────────────────────────────
  const headerRow = sheet.addRow([
    'Party Name :',
    deductorName || '',
    '', '', '',
    'TAN :',
    tan || '',
    '',
  ]);
  headerRow.font = { bold: true, size: 11 };
  headerRow.getCell(1).font = { bold: true, size: 11 };
  headerRow.getCell(6).font = { bold: true, size: 11 };
  headerRow.getCell(7).font = { bold: true, size: 11 };
  headerRow.height = 18;

  // ── Row 2 — Column headers ──────────────────────────────────────────────────
  const HEADERS = [
    'Deductee code',
    'PAN of deductee',
    'First Name',
    'Amount of payment (Rs.)',
    'Date on which amount paid/credited',
    'Section code',
    'Rate at which tax deducted',
    'TDS (Rs.)',
  ];

  const colHeaderRow = sheet.addRow(HEADERS);
  colHeaderRow.height = 30;

  colHeaderRow.eachCell((cell, colNumber) => {
    cell.font = { bold: true, size: 10, color: { argb: 'FF000000' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E1F2' }, // light blue (matches sample)
    };
    cell.border = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' },
    };
    cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' };
  });

  // ── Data rows ───────────────────────────────────────────────────────────────
  for (let i = 0; i < records.length; i++) {
    const rec = records[i];
    const row = sheet.addRow([
      rec.deducteeCode || '02',
      rec.pan          || '',
      rec.name         || '',
      rec.amount       || 0,
      rec.date         || '',
      rec.section      || '',
      rec.rate         || 0,
      rec.tds          || 0,
    ]);

    row.height = 16;

    row.eachCell((cell, colNumber) => {
      cell.border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        bottom: { style: 'thin' },
        right:  { style: 'thin' },
      };
      cell.font = { size: 10 };

      // Column-specific alignment and number format
      switch (colNumber) {
        case 1: // Deductee code — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 2: // PAN — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 3: // Name — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 4: // Amount — right, number format
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0';
          break;
        case 5: // Date — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 6: // Section — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 7: // Rate — right
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          break;
        case 8: // TDS — right, number format
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0';
          break;
      }
    });
  }

  // ── Return as buffer ────────────────────────────────────────────────────────
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}

module.exports = { generateAnnexure };
