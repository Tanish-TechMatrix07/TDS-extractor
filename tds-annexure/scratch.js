const text = `Nature Of Payment : Remuneration/interest/commission to partners (194T)
Bhavinbhai Popatbhai Vaghasiya-50               379161.10    10.00     37916.07       37916.00       37916.00                 37916.00
Popatbhai Ghusabhai vaghasiya-50                756832.36    10.00     75683.07       75683.00       75683.00                 75683.00
Nature Of Payment : TDS on Purchase of Goods (194Q)
Jagtat Trading Co
 Pan No.:AAGFJ1554J           6219080.00        1219080.00    0.10      1219.08        1219.08
Shivmala Trading Co
 Pan No.:AFKFS0738N           30001797.00       25001797.00   0.10      25001.81       25001.81`;

let pendingName = '';
const lines = text.split('\n');
for (const line of lines) {
  if (line.match(/Nature\s+Of\s+Payment/i) || line.match(/PAN\s+No/i)) {
    if (line.match(/Nature\s+Of\s+Payment/i)) pendingName = ''; // reset on new section
    continue;
  }
  const match = line.match(/^(.+?)((?:[\s]*\d+\.\d{2})+)$/);
  if (match) {
    let name = pendingName + ' ' + match[1].trim();
    console.log('Found:', name.trim(), match[2].trim());
    pendingName = '';
  } else {
    pendingName += ' ' + line.trim();
  }
}
