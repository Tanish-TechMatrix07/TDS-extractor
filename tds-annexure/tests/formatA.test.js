const test = require('node:test');
const assert = require('node:assert');
const { parseFormatA } = require('../src/parsers/formatA');
const XLSX = require('xlsx');

test('Format A - Robustness Tests', async (t) => {
  await t.test('Successfully parses when columns are completely swapped', () => {
    const originalRead = XLSX.read;
    XLSX.read = () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
    const originalSheetToJson = XLSX.utils.sheet_to_json;
    XLSX.utils.sheet_to_json = () => [
      ['Deductor XYZ', null, null, null],
      ['TAN', 'ABCD12345E', null, null],
      // Header row but columns are jumbled!
      ['TDS Assessable Amount', 'Net TDS to be Paid', 'TDS %', 'Total Voucher Assessable Amt', 'Party Name'],
      // Data row following the jumbled headers
      ['Nature Of Payment : TDS on Purchase of Goods (194Q)', null, null, null, null],
      [379161.10, 37916.07, 10.00, null, 'Bhavinbhai Vaghasiya'],
      [1219080.00, 1219.08, 0.1, null, 'Jagtat Trading Co']
    ];

    try {
      const buffer = Buffer.from('mock');
      const result = parseFormatA(buffer);

      assert.strictEqual(result.deductorName, 'Deductor XYZ');
      assert.strictEqual(result.tan, 'ABCD12345E');
      
      const records = result.records;
      assert.strictEqual(records.length, 2, 'Should extract 2 records');
      
      const rec1 = records[0];
      assert.strictEqual(rec1.name, 'Bhavinbhai Vaghasiya');
      assert.strictEqual(rec1.amount, 379161.10);
      assert.strictEqual(rec1.tds, 37916.07);
      
      const rec2 = records[1];
      assert.strictEqual(rec2.name, 'Jagtat Trading Co');
      assert.strictEqual(rec2.amount, 1219080.00);
      assert.strictEqual(rec2.tds, 1219.08);
    } finally {
      XLSX.utils.sheet_to_json = originalSheetToJson;
      XLSX.read = originalRead;
    }
  });

  await t.test('Heuristic fallback correctly maps columns when header is missing', () => {
    const originalRead = XLSX.read;
    XLSX.read = () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
    const originalSheetToJson = XLSX.utils.sheet_to_json;
    // No headers, just raw data that has "Party Name" traits and numeric traits
    XLSX.utils.sheet_to_json = () => [
      ['Deductor ABC', null, null, null],
      ['TAN: XBCD12345E', null, null, null],
      // No explicit headers!
      ['Nature Of Payment : TDS on Purchase of Goods (194C)', null, null, null, null, null],
      ['Shivmala Trading Co', null, 30001797.00, 25001797.00, 0.10, 25001.81],
      ['Another Trading Co', null, 50000.00, 50000.00, 10, 5000.00],
      ['Third Party LLC', null, 700000.00, 700000.00, 1, 7000.00],
      ['Fourth Party LLC', null, 800000.00, 800000.00, 2, 16000.00],
      ['Fifth Party LLC', null, 900000.00, 900000.00, 5, 45000.00],
      ['Sixth Party LLC', null, 100000.00, 100000.00, 5, 5000.00]
    ];

    try {
      const buffer = Buffer.from('mock');
      const result = parseFormatA(buffer);

      assert.strictEqual(result.deductorName, 'Deductor ABC');
      assert.strictEqual(result.records.length, 6, 'Should extract 6 records');
      
      const rec1 = result.records[0];
      assert.strictEqual(rec1.name, 'Shivmala Trading Co');
      assert.strictEqual(rec1.amount, 30001797.00);
    } finally {
      XLSX.utils.sheet_to_json = originalSheetToJson;
      XLSX.read = originalRead;
    }
  });
});
