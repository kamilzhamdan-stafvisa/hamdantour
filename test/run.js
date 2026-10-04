// Uji cepat logika rekonsiliasi: node test/run.js manifest.xlsx mutamer.xlsx
const XLSX = require('xlsx');
const HTC = require('../js/core.js');
const [, , mf, uf] = process.argv;
if (!mf || !uf) { console.error('Pemakaian: node test/run.js manifest.xlsx mutamer.xlsx'); process.exit(1); }
const rd = (f) => { const wb = XLSX.readFile(f); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' }); };
const m = HTC.parseSheet(rd(mf), 'manifest');
const u = HTC.parseSheet(rd(uf), 'mutamer');
const r = HTC.reconcile(m.rows, u.rows);
console.log(r.summary);
r.rows.filter((x) => x.problem).forEach((x) => console.log(`${x.id}\t${x.nameManifest || x.nameMutamer}\t${x.category}\t${x.issues.join('; ')}`));
