const test = require('node:test');
const assert = require('node:assert');
const { parseFormatC } = require('../src/parsers/formatC');
const XLSX = require('xlsx');

test('Format C - Robustness Tests', async (t) => {
  await t.test('Recovers corrupt TDS when amount and TDS were mapped to the same value', () => {
    const originalRead = XLSX.read;
    XLSX.read = () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
    const originalSheetToJson = XLSX.utils.sheet_to_json;
    // Format C's known corruption case: tds = amount
    XLSX.utils.sheet_to_json = () => [
      ['Party Name : Deductor Company', null, null, 'TAN : ABCD12345E'],
      ['Deductee code', 'PAN of deductee', 'First Name', 'Amount of payment (Rs.)', 'Date on which amount paid/credited', 'Section code', 'Rate at which tax deducted', 'TDS(Rs.)'],
      // The corrupt data row where TDS == Amount, and rate is completely wrong for 194Q (30 instead of 0.1)
      ['02', 'ABMFM3677L', 'MADHUR GANESH TRADING CO.', 1724450.5, '31/01/2026', '194Q', 30, 1724450.5]
    ];

    try {
      const buffer = Buffer.from('mock');
      const result = parseFormatC(buffer);

      assert.strictEqual(result.deductorName, 'Deductor Company');
      assert.strictEqual(result.tan, 'ABCD12345E');
      
      const records = result.records;
      assert.strictEqual(records.length, 1);
      
      const rec = records[0];
      // 194Q is 0.1%. So amount 1724450.5 * 0.1% = 1724.45 (ceil to 1725)
      assert.strictEqual(rec.amount, 1724450.5);
      assert.strictEqual(rec.rate, 0.1);
      assert.strictEqual(rec.tds, 1725);
    } finally {
      XLSX.utils.sheet_to_json = originalSheetToJson;
      XLSX.read = originalRead;
    }
  });

  await t.test('Heuristic fallback correctly maps columns when values are slightly offset (VALUE IS ABOVE/MERGED)', () => {
    const originalRead = XLSX.read;
    XLSX.read = () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
    const originalSheetToJson = XLSX.utils.sheet_to_json;
    XLSX.utils.sheet_to_json = () => [
      // No standard header. Values pushed to different columns!
      ['ABMFM3677L', null, 'Some Deductee Name', null, '10/10/2026', '194C', null, 50000, 1000],
      ['XBCDF1234F', 'Another Deductee', null, '11/10/2026', null, '194C', 20000, null, 400],
      ['YBCDF1234F', 'Third Deductee', null, '12/10/2026', null, '194C', 30000, null, 600],
      ['ZBCDF1234F', 'Fourth Deductee', null, '13/10/2026', null, '194C', 40000, null, 800],
      ['WBCDF1234F', 'Fifth Deductee', null, '14/10/2026', null, '194C', 50000, null, 1000],
      ['VBCDF1234F', 'Sixth Deductee', null, '15/10/2026', null, '194C', 60000, null, 1200]
    ];

    try {
      const buffer = Buffer.from('mock');
      const result = parseFormatC(buffer);
      
      assert.strictEqual(result.records.length, 6);
      
      // The heuristic should identify the first column as PAN, the second/third as Name, etc.
      // Format C heuristics:
      // PAN = col 0
      // Date = col 3, 4
      // Section = col 5
      // The exact mappings depend on how many occurrences fall into each column.
      const rec = result.records[1];
      assert.strictEqual(rec.pan, 'XBCDF1234F');
      assert.strictEqual(rec.section, '194C');
      
    } finally {
      XLSX.utils.sheet_to_json = originalSheetToJson;
      XLSX.read = originalRead;
    }
  });
});
