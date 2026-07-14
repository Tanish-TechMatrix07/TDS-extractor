const test = require('node:test');
const assert = require('node:assert');
const { parseFormatB } = require('../src/parsers/formatB');
const child_process = require('child_process');

test('Format B - Robustness Tests', async (t) => {
  await t.test('Handles disconnected text via pendingName buffer correctly', async () => {
    // Mock the child_process to return our specific broken PDF layout
    const originalSpawn = child_process.spawn;
    child_process.spawn = () => {
      return {
        stdout: { on: (event, cb) => cb(Buffer.from('__JSON_START__' + JSON.stringify({
          text: `
Nature Of Payment : TDS on Purchase of Goods (194Q)
Jagtat Trading Co
 Pan No.:AAGFJ1554J
6219080.00 1219080.00 0.10 1219.08 1219.08
Shivmala Trading Co
30001797.00 25001797.00 0.10 25001.81 25001.81
          `
        }) + '__JSON_END__')) },
        stderr: { on: () => {} },
        on: (event, cb) => { if (event === 'close') cb(0); },
        stdin: { write: () => {}, end: () => {} }
      };
    };

    try {
      const buffer = Buffer.from('mock');
      const result = await parseFormatB(buffer, 'test.pdf');
      
      const records = result.records;
      assert.strictEqual(records.length, 2);
      
      // Jagtat Trading Co was separated by PAN No and a newline
      assert.strictEqual(records[0].name, 'Jagtat Trading Co');
      assert.strictEqual(records[0].amount, 1219080.00);
      
      // Shivmala was separated by a newline but no PAN
      assert.strictEqual(records[1].name, 'Shivmala Trading Co');
      assert.strictEqual(records[1].amount, 25001797.00);

    } finally {
      child_process.spawn = originalSpawn;
    }
  });

  await t.test('Dynamically corrects amounts that were prefix-joined by poor layout', async () => {
    const originalSpawn = child_process.spawn;
    child_process.spawn = () => {
      return {
        stdout: { on: (event, cb) => cb(Buffer.from('__JSON_START__' + JSON.stringify({
          text: `
Nature Of Payment : Remuneration/interest/commission to partners (194T)
Bhavinbhai Popatbhai Vaghasiya-50
379161.10 10.00 37916.07 37916.00 37916.00 37916.00
Popatbhai Ghusabhai vaghasiya-
50756832.36 10.00 75683.07 75683.00 75683.00 75683.00
          `
        }) + '__JSON_END__')) },
        stderr: { on: () => {} },
        on: (event, cb) => { if (event === 'close') cb(0); },
        stdin: { write: () => {}, end: () => {} }
      };
    };

    try {
      const buffer = Buffer.from('mock');
      const result = await parseFormatB(buffer, 'test.pdf');
      
      const records = result.records;
      assert.strictEqual(records.length, 2);
      
      // Due to the regex behavior without the 50 prefix logic, it would extract 50379161.10.
      // But our self-healing logic dynamically corrects it to 379161.10
      assert.strictEqual(records[0].name, 'Bhavinbhai Popatbhai Vaghasiya-50');
      // Verify amount is stripped of the "50"
      assert.strictEqual(records[0].amount, 379161.1);
      
      assert.strictEqual(records[1].name, 'Popatbhai Ghusabhai vaghasiya-');
      assert.strictEqual(records[1].amount, 756832.36);

    } finally {
      child_process.spawn = originalSpawn;
    }
  });
});
