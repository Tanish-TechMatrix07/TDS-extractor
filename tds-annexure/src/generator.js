'use strict';

/**
 * Output Generator
 * Takes the normalised internal record array and produces an Annexure Excel file
 * with the full 21-column government-format layout.
 *
 * Output sheet structure:
 *   Sheet name : Annexure
 *   Row 1      : Party Name : <deductorName>   [gap cells]   TAN : <tan>
 *   Row 2      : Column headers (21 columns)
 *   Row 3+     : Data rows, one per record
 *
 * Column layout (A–U):
 *   A  Deductee code
 *   B  PAN of deductee
 *   C  First Name
 *   D  Middle Name
 *   E  Last Name
 *   F  Address 1
 *   G  Address 2
 *   H  State
 *   I  Pin Code
 *   J  Amount of payment (Rs.)
 *   K  Date on which amount paid/credited
 *   L  Section code
 *   M  Rate at which tax deducted
 *   N  TDS (Rs.)
 *   O  Date on which tax deducted
 *   P  Challan Detail [Sr No (BSR, Date, Challan No.)]
 *   Q  Date of furnishing Tax Deduction Certificate
 *   R  Reason for non-deduction / lower deduction if any
 *   S  Paid by book entry or otherwise
 *   T  Certificate No u/s 197
 *   U  Deductee/Party Reference No
 *
 * Styling:
 *   - Header row: bold, light-purple fill, all borders
 *   - Data rows: all borders
 *   - Numeric columns: right-aligned, number format
 *   - Date/centre columns: centre-aligned
 *   - Column widths set to sensible values
 */

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

async function generateAnnexure(deductorName, tan, records, sheetName, challanRecords) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'TDS Annexure Generator';
  workbook.created = new Date();

  const name = sheetName || 'Annexure';
  const sheet = workbook.addWorksheet(name);

  // ── Column widths (21 columns A–U) ──────────────────────────────────────────
  sheet.columns = [
    { key: 'deducteeCode',          width: 16 },  // A
    { key: 'pan',                   width: 18 },  // B
    { key: 'name',                  width: 42 },  // C  First Name
    { key: 'middleName',            width: 18 },  // D
    { key: 'lastName',              width: 18 },  // E
    { key: 'address1',              width: 28 },  // F
    { key: 'address2',              width: 28 },  // G
    { key: 'state',                 width: 14 },  // H
    { key: 'pinCode',               width: 12 },  // I
    { key: 'amount',                width: 24 },  // J
    { key: 'date',                  width: 28 },  // K
    { key: 'section',               width: 14 },  // L
    { key: 'rate',                  width: 22 },  // M
    { key: 'tds',                   width: 12 },  // N
    { key: 'dateOfTdsDeduction',    width: 22 },  // O
    { key: 'challanDetail',         width: 32 },  // P
    { key: 'dateOfFurnishingCert',  width: 28 },  // Q
    { key: 'reasonForNonDeduction', width: 28 },  // R
    { key: 'paidByBookEntry',       width: 20 },  // S
    { key: 'certificateNo197',      width: 18 },  // T
    { key: 'partyReferenceNo',     width: 22 },  // U
  ];

  // ── Row 1 — Deductor header ─────────────────────────────────────────────────
  // Party Name at col 1, TAN at col 3 (after Party Name : + deductor name)
  const headerRow = sheet.addRow([
    'Party Name :',
    deductorName || '',
    'TAN :',
    tan || '',
  ]);
  // Fill remaining cells to column 21 width
  for (let c = 5; c <= 21; c++) {
    headerRow.getCell(c).value = '';
  }

  headerRow.font = { bold: true, size: 11 };
  headerRow.getCell(1).font = { bold: true, size: 11 };
  headerRow.getCell(3).font = { bold: true, size: 11 };
  headerRow.getCell(4).font = { bold: true, size: 11 };
  headerRow.height = 18;

  // ── Row 2 — Column headers (21 columns) ─────────────────────────────────────
  const HEADERS = [
    'Deductee code',
    'PAN of deductee',
    'First Name',
    'Middle Name',
    'Last Name',
    'Address 1',
    'Address 2',
    'State',
    'Pin Code',
    'Amount of payment (Rs.)',
    'Date on which amount paid/credited',
    'Section code',
    'Rate at which tax deducted',
    'TDS (Rs.)',
    'Date on which tax deducted',
    'Challan Detail [Sr No (BSR, Date, Challan No.)]',
    'Date of furnishing Tax Deduction Certificate',
    'Reason for non-deduction / lower deduction if any',
    'Paid by book entry or otherwise',
    'Certificate No u/s 197',
    'Deductee/Party Reference No',
  ];

  const colHeaderRow = sheet.addRow(HEADERS);
  colHeaderRow.height = 30;

  colHeaderRow.eachCell((cell) => {
    cell.font = { bold: true, size: 10, color: { argb: 'FF000000' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD4C5F9' }, // light purple
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
      rec.deducteeCode           || '02',          // A
      rec.pan                    || '',             // B
      rec.name                   || '',             // C  First Name
      rec.middleName             || '',             // D
      rec.lastName               || '',             // E
      rec.address1               || '',             // F
      rec.address2               || '',             // G
      rec.state                  || '',             // H
      rec.pinCode                || '',             // I
      rec.amount                 || 0,              // J
      rec.date                   || '',             // K
      rec.section                || '',             // L
      rec.rate                   || 0,              // M
      rec.tds                    || 0,              // N
      rec.dateOfTdsDeduction     || '',             // O
      rec.challanDetail          || '',             // P
      rec.dateOfFurnishingCert   || '',             // Q
      rec.reasonForNonDeduction  || '',             // R
      rec.paidByBookEntry        || '',             // S
      rec.certificateNo197       || '',             // T
      rec.partyReferenceNo      || '',             // U
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
        case 1:  // A — Deductee code — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 2:  // B — PAN — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 3:  // C — First Name — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 4:  // D — Middle Name — left
        case 5:  // E — Last Name — left
        case 6:  // F — Address 1 — left
        case 7:  // G — Address 2 — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 8:  // H — State — centre
        case 9:  // I — Pin Code — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 10: // J — Amount — right, number format
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0';
          break;
        case 11: // K — Date — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 12: // L — Section — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 13: // M — Rate — right
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          break;
        case 14: // N — TDS — right, number format
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0';
          break;
        case 15: // O — Date on which tax deducted — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 16: // P — Challan Detail — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 17: // Q — Date of furnishing cert — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 18: // R — Reason — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 19: // S — Paid by — left
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
          break;
        case 20: // T — Certificate No — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
        case 21: // U — Deductee/Party Reference No — centre
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          break;
      }
    });
  }

  // ── Sheet 2: Challan Details (always present, even if empty) ─────────────
  const challanSheet = workbook.addWorksheet('Challan');
  const chRecords = challanRecords || [];

  // Row 1: Header info
  const chHeaderRow = challanSheet.addRow([
    'Party Name :',
    deductorName || '',
    'TAN :',
    tan || '',
  ]);
  for (let c = 5; c <= 16; c++) { chHeaderRow.getCell(c).value = ''; }
  chHeaderRow.font = { bold: true, size: 11 };
  chHeaderRow.getCell(1).font = { bold: true, size: 11 };
  chHeaderRow.getCell(3).font = { bold: true, size: 11 };
  chHeaderRow.getCell(4).font = { bold: true, size: 11 };
  chHeaderRow.height = 18;

  // Row 2: Column headers
  const CHALLAN_HEADERS = [
    'S. No.',
    'Section Code',
    'TDS (Rs.)',
    'Surcharge (Rs.)',
    'Education Cess (Rs.)',
    'Higher Education Cess',
    'Interest (Rs.)',
    'Other (Rs.)',
    'Fees Amount (Rs.)',
    'Cheque/DD No.',
    'BSR Code',
    'Date on which Tax Deposited',
    'Challan Serial No.',
    'Book Entry?',
    'Minor Head',
  ];

  const chColRow = challanSheet.addRow(CHALLAN_HEADERS);
  chColRow.height = 28;
  chColRow.eachCell((cell) => {
    cell.font = { bold: true, size: 10, color: { argb: 'FF000000' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD4C5F9' }, // light purple
    };
    cell.border = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' },
    };
    cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' };
  });

  // Column widths
  challanSheet.columns = [
    { width: 10 },  // A  S. No.
    { width: 14 },  // B  Section Code
    { width: 14 },  // C  TDS (Rs.)
    { width: 16 },  // D  Surcharge (Rs.)
    { width: 18 },  // E  Education Cess (Rs.)
    { width: 20 },  // F  Higher Education Cess
    { width: 14 },  // G  Interest (Rs.)
    { width: 14 },  // H  Other (Rs.)
    { width: 16 },  // I  Fees Amount (Rs.)
    { width: 18 },  // J  Cheque/DD No.
    { width: 14 },  // K  BSR Code
    { width: 22 },  // L  Date on which Tax Deposited
    { width: 20 },  // M  Challan Serial No.
    { width: 14 },  // N  Book Entry?
    { width: 14 },  // O  Minor Head
  ];

  // Data rows (only if there are records)
  for (let i = 0; i < chRecords.length; i++) {
    const rec = chRecords[i];
    const row = challanSheet.addRow([
      rec.srNo        || '',
      rec.section     || '',
      rec.tds         || 0,
      rec.surcharge   || 0,
      rec.eduCess     || 0,
      rec.higherEduCess || 0,
      rec.interest    || 0,
      rec.other       || 0,
      rec.feesAmount  || 0,
      rec.chequeNo    || '',
      rec.bsrCode     || '',
      rec.depositDate || '',
      rec.challanNo   || '',
      rec.bookEntry   || '',
      rec.minorHead   || '',
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

      // Numeric columns right-aligned
      if ([3, 4, 5, 6, 7, 8, 9].includes(colNumber)) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0';
      } else if ([1, 2, 12, 15].includes(colNumber)) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }

  // ── Write as .xlsx buffer, then convert to .xls (biff8) ────────────────────
  const xlsxBuffer = await workbook.xlsx.writeBuffer();

  // Convert xlsx → xls (biff8) using the xlsx library
  const xlsWorkbook = XLSX.read(xlsxBuffer, { type: 'buffer' });
  const xlsBuffer = XLSX.write(xlsWorkbook, { bookType: 'biff8', type: 'buffer' });

  return xlsBuffer;
}

module.exports = { generateAnnexure };

