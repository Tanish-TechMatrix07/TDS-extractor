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

  // ── Sheet 1: Challan Details (always present, even if empty, and appears first) ─────────────
  const challanSheet = workbook.addWorksheet('Challan');
  const chRecords = challanRecords || [];

  // Column widths
  challanSheet.columns = [
    { key: 'srNo',                  width: 10 },  // A  S. No.
    { key: 'section',               width: 14 },  // B  Section Code
    { key: 'tds',                   width: 14 },  // C  TDS(Rs.)
    { key: 'interest',              width: 16 },  // D  Interest (Rs.)
    { key: 'other',                 width: 14 },  // E  Other (Rs.)
    { key: 'feesAmount',            width: 16 },  // F  Fees Amount (Rs.)
    { key: 'chequeNo',              width: 18 },  // G  Cheque/DD No.
    { key: 'bsrCode',               width: 14 },  // H  BSR Code
    { key: 'depositDate',           width: 22 },  // I  Date on which Tax Deposited
    { key: 'challanNo',             width: 24 },  // J  Transfer Voucher/Challan Serial No
    { key: 'bookEntry',             width: 24 },  // K  Whether TDS deposited by book entry
    { key: 'minorHead',             width: 14 },  // L  Minor Head
  ];

  // Row 1: Header info
  challanSheet.mergeCells('A1:B1');
  const chPartyLabel = challanSheet.getCell('A1');
  chPartyLabel.value = 'Party Name :';
  chPartyLabel.font = { bold: true, underline: true, name: 'Calibri', size: 11 };
  chPartyLabel.alignment = { horizontal: 'right', vertical: 'middle' };

  challanSheet.mergeCells('C1:E1');
  const chPartyVal = challanSheet.getCell('C1');
  chPartyVal.value = deductorName || '';
  chPartyVal.font = { bold: true, name: 'Calibri', size: 11 };
  chPartyVal.alignment = { horizontal: 'left', vertical: 'middle' };

  const chTanLabel = challanSheet.getCell('F1');
  chTanLabel.value = 'TAN :';
  chTanLabel.font = { bold: true, underline: true, name: 'Calibri', size: 11 };
  chTanLabel.alignment = { horizontal: 'right', vertical: 'middle' };

  const chTanVal = challanSheet.getCell('G1');
  chTanVal.value = tan || '';
  chTanVal.font = { bold: true, name: 'Calibri', size: 11 };
  chTanVal.alignment = { horizontal: 'left', vertical: 'middle' };

  challanSheet.getRow(1).height = 18;

  // Row 2: Warning info
  challanSheet.mergeCells('C2:H2');
  const chWarningCell = challanSheet.getCell('C2');
  chWarningCell.value = 'Please do not Cut/Copy/Paste (it may cause inconsitancy in data).';
  chWarningCell.font = { color: { argb: 'FFFF0000' }, italic: true, name: 'Calibri', size: 10 };
  chWarningCell.alignment = { horizontal: 'center', vertical: 'middle' };
  challanSheet.getRow(2).height = 18;

  // Row 3: Subheader
  const chRow3 = challanSheet.getRow(3);
  chRow3.getCell(1).value = 'Details of tax deducted and paid to the credit of the Central Government';
  chRow3.getCell(1).font = { bold: true, name: 'Calibri', size: 11 };
  chRow3.height = 18;

  // Row 4: Column headers
  const CHALLAN_HEADERS = [
    'S. No.',
    'Section Code',
    'TDS(Rs.)',
    'Interest (Rs.)',
    'Other (Rs.)',
    'Fees Amount (Rs.)',
    'Cheque/DD No.',
    'BSR Code',
    'Date on which Tax Deposited',
    'Transfer Voucher/Challan Serial No',
    'Whether TDS deposited by book entry',
    'Minor Head',
  ];

  const chColRow = challanSheet.getRow(4);
  chColRow.values = CHALLAN_HEADERS;
  chColRow.height = 30;
  chColRow.eachCell((cell) => {
    cell.font = { bold: true, size: 10, color: { argb: 'FF000000' }, name: 'Calibri' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFA6B5F7' }, // lighter periwinkle blue
    };
    cell.border = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' },
    };
    cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' };
  });

  // Enable AutoFilter on Challan sheet
  challanSheet.autoFilter = 'A4:L4';

  // Data rows (only if there are records)
  for (let i = 0; i < chRecords.length; i++) {
    const rec = chRecords[i];
    const row = challanSheet.getRow(5 + i);
    row.values = [
      rec.srNo        || '',
      rec.section     || '',
      rec.tds         || 0,
      rec.interest    || 0,
      rec.other       || 0,
      rec.feesAmount  || 0,
      rec.chequeNo    || '',
      rec.bsrCode     || '',
      rec.depositDate || '',
      rec.challanNo   || '',
      rec.bookEntry   || '',
      rec.minorHead   || '',
    ];

    row.height = 18;
    row.eachCell((cell, colNumber) => {
      cell.border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        bottom: { style: 'thin' },
        right:  { style: 'thin' },
      };
      cell.font = { size: 10, name: 'Calibri' };

      // Numeric columns right-aligned
      if ([3, 4, 5, 6].includes(colNumber)) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.numFmt = '#,##0';
      } else if ([1, 2, 9, 12].includes(colNumber)) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  }


  // ── Sheet 2: annexure (appears second) ──────────────────────────────────────────
  const sheet = workbook.addWorksheet('annexure');

  // Column widths (21 columns A–U)
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
    { key: 'partyReferenceNo',      width: 22 },  // U
  ];

  // Row 1 — Deductor header
  sheet.mergeCells('A1:B1');
  const annPartyLabel = sheet.getCell('A1');
  annPartyLabel.value = 'Party Name :';
  annPartyLabel.font = { bold: true, underline: true, name: 'Calibri', size: 11 };
  annPartyLabel.alignment = { horizontal: 'right', vertical: 'middle' };

  sheet.mergeCells('C1:E1');
  const annPartyVal = sheet.getCell('C1');
  annPartyVal.value = deductorName || '';
  annPartyVal.font = { bold: true, name: 'Calibri', size: 11 };
  annPartyVal.alignment = { horizontal: 'left', vertical: 'middle' };

  const annTanLabel = sheet.getCell('F1');
  annTanLabel.value = 'TAN :';
  annTanLabel.font = { bold: true, underline: true, name: 'Calibri', size: 11 };
  annTanLabel.alignment = { horizontal: 'right', vertical: 'middle' };

  const annTanVal = sheet.getCell('G1');
  annTanVal.value = tan || '';
  annTanVal.font = { bold: true, name: 'Calibri', size: 11 };
  annTanVal.alignment = { horizontal: 'left', vertical: 'middle' };

  sheet.getRow(1).height = 18;

  // Row 2: Warning info
  sheet.mergeCells('C2:F2');
  const annWarningCell = sheet.getCell('C2');
  annWarningCell.value = 'Please do not Cut/Copy/Paste (it may cause inconsitancy in data).';
  annWarningCell.font = { color: { argb: 'FFFF0000' }, italic: true, name: 'Calibri', size: 10 };
  annWarningCell.alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getRow(2).height = 18;

  // Row 3: Blank spacing row
  sheet.getRow(3).height = 18;

  // Row 4 — Column headers (21 columns)
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
    'TDS(Rs.)',
    'Date on which tax deducted',
    'Challan Detail [Sr No (BSR, Date, Challan No.)]',
    'Date of furnishing Tax Deduction Certificate',
    'Reason for non-deduction / lower deduction if any',
    'Paid by book entry or otherwise',
    'Certificate No u/s 197',
    'Deductee/Party Reference No',
  ];

  const annColRow = sheet.getRow(4);
  annColRow.values = HEADERS;
  annColRow.height = 30;

  annColRow.eachCell((cell) => {
    cell.font = { bold: true, size: 10, color: { argb: 'FF000000' }, name: 'Calibri' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFA6B5F7' }, // lighter periwinkle blue
    };
    cell.border = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' },
    };
    cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' };
  });

  // Enable AutoFilter on annexure sheet
  sheet.autoFilter = 'A4:U4';

  // Data rows
  for (let i = 0; i < records.length; i++) {
    const rec = records[i];
    const row = sheet.getRow(5 + i);
    row.values = [
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
      rec.partyReferenceNo       || '',             // U
    ];

    row.height = 18;

    row.eachCell((cell, colNumber) => {
      cell.border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        bottom: { style: 'thin' },
        right:  { style: 'thin' },
      };
      cell.font = { size: 10, name: 'Calibri' };

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

  // ── Write as .xlsx buffer directly to preserve all formatting and styles ────
  const xlsxBuffer = await workbook.xlsx.writeBuffer();
  return xlsxBuffer;
}

module.exports = { generateAnnexure };

