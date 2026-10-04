/*
 * Hamdan Tour Check — membuka kembali laporan Excel hasil unduhan.
 * Mengembalikan KPU, nama keberangkatan, baris rekonsiliasi, ringkasan, dan catatan Action need.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.HTCImport = factory(root.HTC);
})(typeof self !== 'undefined' ? self : this, function (HTC) {
  'use strict';

  const DATA_SHEETS = ['Detail', 'Data Bermasalah'];
  const BULAN = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];

  const grid = (XLSX, ws) => XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  const cellText = (ws, addr) => (ws && ws[addr] && ws[addr].v !== undefined ? String(ws[addr].v) : '');
  const str = (v) => (v === undefined || v === null ? '' : String(v).trim());

  function isReport(wb) {
    if (wb.SheetNames.includes('_HTC')) return true;
    return (
      wb.SheetNames.includes('Ringkasan') &&
      DATA_SHEETS.some((n) => wb.SheetNames.includes(n)) &&
      /HAMDAN TOUR CHECK/i.test(cellText(wb.Sheets.Ringkasan, 'C1'))
    );
  }

  // "1 Oktober 2026, 10:33:06 UTC" -> Date (cadangan untuk laporan lama tanpa kolom ISO)
  function parseStampText(t) {
    const m = str(t).match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4}),\s*(\d{2}):(\d{2}):(\d{2})/);
    if (!m) return null;
    const mo = BULAN.indexOf(m[2].toLowerCase());
    if (mo < 0) return null;
    return new Date(+m[3], mo, +m[1], +m[4], +m[5], +m[6]);
  }

  function readMeta(wb, XLSX) {
    const meta = {};
    if (!wb.SheetNames.includes('_HTC')) return meta;
    grid(XLSX, wb.Sheets._HTC).forEach((r) => { if (r[0]) meta[str(r[0])] = r[1]; });
    return meta;
  }

  // Cadangan: ambil angka dari sheet Ringkasan bila meta tidak ada.
  function summaryFromSheet(wb, XLSX, rows) {
    const s = HTC.summarize(rows, rows.filter((r) => r.source === 'Manifest').length);
    const ws = wb.Sheets.Ringkasan;
    if (!ws) return s;
    const first = (v) => { const m = str(v).match(/\d+/); return m ? +m[0] : null; };
    grid(XLSX, ws).forEach((r) => {
      const label = str(r[2]);
      const v = first(r[3]);
      if (v === null) return;
      const map = {
        'Total Jamaah': 'total', 'Visa Printed': 'printed', 'Belum Visa': 'belum', 'Tidak Match': 'tidakMatch',
        'Data Bermasalah': 'bermasalah', 'Passport Match': 'passportMatch', 'Name Match': 'nameMatch',
        'Visa Number': 'visaNumber', 'Visa Status': 'visaPrinted',
      };
      if (map[label]) s[map[label]] = v;
      if (label === 'Name Match') {
        const sim = str(r[4]).match(/mirip:\s*(\d+)/i), dif = str(r[4]).match(/berbeda:\s*(\d+)/i);
        if (sim) s.nameSimilar = +sim[1];
        if (dif) s.nameDiff = +dif[1];
      }
    });
    return s;
  }

  /**
   * Mengembalikan null jika workbook bukan laporan Hamdan Tour Check.
   * Melempar Error berpesan jelas jika laporan rusak.
   */
  function fromWorkbook(wb, XLSX) {
    if (!isReport(wb)) return null;
    const sheetName = DATA_SHEETS.find((n) => wb.SheetNames.includes(n));
    if (!sheetName) throw new Error('Laporan tidak memuat sheet Detail atau Data Bermasalah.');
    const ws = wb.Sheets[sheetName];
    const m = grid(XLSX, ws);

    let h = -1;
    for (let r = 0; r < Math.min(30, m.length); r++) {
      if (str(m[r][0]) === 'No' && m[r].some((c) => str(c) === 'Status Visa')) { h = r; break; }
    }
    if (h < 0) throw new Error('Judul kolom laporan tidak ditemukan. File mungkin sudah diubah.');

    const idx = {};
    HTC.REPORT_COLS.forEach((c) => {
      const i = m[h].findIndex((x) => str(x) === c[0]);
      if (i >= 0) idx[c[1]] = i;
    });

    const rows = [];
    const actions = {};
    for (let r = h + 1; r < m.length; r++) {
      const line = m[r];
      const get = (f) => (idx[f] === undefined ? '' : str(line[idx[f]]));
      if (!get('no') && !get('nameManifest') && !get('nameMutamer')) continue;

      const no = parseInt(get('no'), 10) || rows.length + 1;
      const category = get('category');
      const source = category === HTC.CAT.EXTRA ? 'Mutamer List saja' : 'Manifest';
      const raw = get('issuesText');
      const issues = (raw.includes('\n') ? raw.split('\n') : raw.split('; ')).map((x) => x.trim()).filter(Boolean);
      const id = get('id') || (source === 'Manifest' ? 'M' + (no - 1) : 'U' + no);

      const row = {
        id, source, no,
        regCode: get('regCode'), trxCode: '',
        nameManifest: get('nameManifest'), passportManifest: get('passportManifest'), expiry: '',
        nameMutamer: get('nameMutamer'), passportMutamer: get('passportMutamer'),
        passportMatch: get('passportMatch') || 'Tidak', nameMatch: get('nameMatch') || '-', nameScore: null,
        visaNo: get('visaNo'), mofaNo: get('mofaNo'), visaStatus: get('visaStatus'),
        category, issues, problem: issues.length > 0,
      };
      rows.push(row);

      if (get('needText').toUpperCase() === 'YA') {
        const reason = get('reason');
        let at = get('atISO') ? new Date(get('atISO')) : parseStampText(get('atText'));
        if (!at || isNaN(at)) at = new Date();
        actions[id] = { need: true, reason: reason === '(alasan belum diisi)' ? '' : reason, at };
      }
    }
    if (!rows.length) throw new Error('Laporan tidak berisi baris jamaah.');

    const meta = readMeta(wb, XLSX);
    let summary = null;
    if (meta.summary) { try { summary = JSON.parse(meta.summary); } catch (e) { summary = null; } }
    if (!summary) summary = summaryFromSheet(wb, XLSX, rows);

    let kpu = str(meta.kpu), departure = str(meta.departure);
    if (!kpu && !departure) {
      const t = cellText(ws, 'C3').match(/Kode KPU:\s*(.*?)\s*\|\s*Keberangkatan:\s*(.*)$/);
      if (t) { kpu = t[1] === '-' ? '' : t[1]; departure = t[2] === '-' ? '' : t[2]; }
    }

    let printed = meta.printedISO ? new Date(str(meta.printedISO)) : parseStampText(cellText(ws, 'C4'));
    if (!printed || isNaN(printed)) printed = null;

    const kind = sheetName === 'Detail' ? 'detail' : 'bermasalah';
    const partial = kind === 'bermasalah' || str(meta.partial) === '1';
    return { kpu, departure, printed, kind, partial, summary, rows, actions };
  }

  return { fromWorkbook, isReport, parseStampText };
});
