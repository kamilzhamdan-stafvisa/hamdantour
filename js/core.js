/*
 * Hamdan Tour Check — inti rekonsiliasi.
 * Murni fungsi (tanpa DOM, tanpa jaringan, tanpa penyimpanan) sehingga bisa diuji di Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HTC = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- pembersih nilai ---------- */

  // Sel dari sistem lain sering berbentuk ="TEKS" — kupas jadi TEKS.
  function clean(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    let s = String(v).trim();
    const m = s.match(/^="([\s\S]*)"$/);
    if (m) s = m[1];
    return s.trim();
  }

  const normPass = (v) => clean(v).toUpperCase().replace(/[^A-Z0-9]/g, '');

  function normName(v) {
    return clean(v)
      .toUpperCase()
      .replace(/[^A-Z ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    let prev = new Array(b.length + 1);
    for (let j = 0; j <= b.length; j++) prev[j] = j;
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) {
        cur[j] = Math.min(
          prev[j] + 1,
          cur[j - 1] + 1,
          prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
        );
      }
      prev = cur;
    }
    return prev[b.length];
  }

  function similarity(a, b) {
    if (!a && !b) return 1;
    const max = Math.max(a.length, b.length);
    return max ? 1 - levenshtein(a, b) / max : 1;
  }

  const sortedTokens = (s) => s.split(' ').filter(Boolean).sort().join(' ');

  // Hasil: { status: 'Sama'|'Mirip'|'Beda', score: 0..1 }
  function compareNames(a, b) {
    const na = normName(a);
    const nb = normName(b);
    if (!na || !nb) return { status: 'Beda', score: 0 };
    if (na === nb) return { status: 'Sama', score: 1 };
    if (sortedTokens(na) === sortedTokens(nb)) return { status: 'Mirip', score: 0.99 };
    const score = similarity(na, nb);
    return { status: score >= 0.85 ? 'Mirip' : 'Beda', score };
  }

  function parseDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      return new Date(Math.round((v - 25569) * 86400 * 1000)); // serial Excel
    }
    const s = clean(v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  const fmtDate = (d) => (d ? d.toISOString().slice(0, 10) : '');

  /* ---------- deteksi kolom ---------- */

  const keyOf = (h) => clean(h).toLowerCase().replace(/[^a-z0-9]/g, '');

  const MANIFEST_COLS = {
    name: ['nama', 'namajamaah', 'namalengkap', 'name', 'fullname'],
    passport: ['nomorpaspor', 'nopaspor', 'nomerpaspor', 'paspor', 'passportnumber', 'passportno', 'passport'],
    expiry: ['tanggalkadaluarsapaspor', 'tanggalexpiredpaspor', 'masaberlakupaspor', 'passportexpiry', 'expirydate'],
    regCode: ['koderegistrasi', 'kodereg', 'registrationcode'],
    trxCode: ['kodetransaksi', 'kodetrx', 'transactioncode'],
    visaStatus: ['statusvisa'],
  };

  const MUTAMER_COLS = {
    name: ['mutamername', 'nama', 'name'],
    passport: ['passportnumber', 'passportno', 'nomorpaspor', 'passport'],
    visaStatus: ['visastatus', 'statusvisa'],
    visaNo: ['visanumber', 'visano', 'nomorvisa'],
    mofaNo: ['mofanumber', 'mofano', 'nomormofa'],
    agent: ['mainexternalagentname', 'externalagentname', 'subeaname'],
    biometric: ['biometricstatus'],
    age: ['mutamerage', 'age', 'umur'],
  };

  function mapColumns(headerRow, spec) {
    const keys = headerRow.map(keyOf);
    const map = {};
    for (const field in spec) {
      for (const alias of spec[field]) {
        const idx = keys.indexOf(alias);
        if (idx >= 0) { map[field] = idx; break; }
      }
    }
    return map;
  }

  // Cari baris header di 10 baris pertama (file ekspor kadang punya judul di atas).
  function detectHeader(matrix, spec) {
    let best = { row: 0, hits: -1, map: {} };
    for (let r = 0; r < Math.min(10, matrix.length); r++) {
      const map = mapColumns(matrix[r] || [], spec);
      const hits = Object.keys(map).length;
      if (hits > best.hits) best = { row: r, hits, map };
    }
    return best;
  }

  /**
   * Ubah matriks sheet (array of arrays) menjadi baris objek.
   * kind: 'manifest' | 'mutamer'. Melempar Error berpesan jelas jika kolom wajib tidak ada.
   */
  function parseSheet(matrix, kind) {
    const spec = kind === 'manifest' ? MANIFEST_COLS : MUTAMER_COLS;
    const label = kind === 'manifest' ? 'Manifest' : 'Mutamer List';
    const { row: hRow, map } = detectHeader(matrix, spec);
    const missing = ['name', 'passport'].filter((f) => map[f] === undefined);
    if (kind === 'mutamer' && map.visaStatus === undefined) missing.push('visaStatus');
    if (missing.length) {
      const human = { name: 'nama', passport: 'nomor paspor', visaStatus: 'status visa' };
      throw new Error(
        `File ${label}: kolom ${missing.map((m) => human[m]).join(', ')} tidak ditemukan. ` +
          `Pastikan baris judul kolom ada di 10 baris pertama.`
      );
    }
    const rows = [];
    for (let r = hRow + 1; r < matrix.length; r++) {
      const line = matrix[r] || [];
      const rec = { _line: r + 1 };
      let any = false;
      for (const f in map) {
        const raw = line[map[f]];
        rec[f] = f === 'expiry' ? raw : clean(raw);
        if (clean(raw) !== '') any = true;
      }
      if (any) rows.push(rec);
    }
    return { rows, columns: Object.keys(map), headerRow: hRow + 1 };
  }

  /* ---------- rekonsiliasi ---------- */

  const CAT = {
    PRINTED: 'Visa Printed',
    BELUM: 'Belum Visa',
    TIDAK_MATCH: 'Tidak Match',
    EXTRA: 'Hanya di Mutamer List',
  };

  const isPrinted = (s) => /printed/i.test(s || '');

  function reconcile(manifestRows, mutamerRows, opts) {
    opts = opts || {};
    const today = opts.today || new Date();
    const limit = new Date(today.getTime());
    limit.setMonth(limit.getMonth() + 6); // syarat umum: paspor berlaku min. 6 bulan

    // indeks mutamer per paspor
    const byPass = new Map();
    mutamerRows.forEach((u, i) => {
      const k = normPass(u.passport);
      if (!k) return;
      if (!byPass.has(k)) byPass.set(k, []);
      byPass.get(k).push(i);
    });
    // hitung paspor ganda di manifest
    const manifestPassCount = new Map();
    manifestRows.forEach((m) => {
      const k = normPass(m.passport);
      if (k) manifestPassCount.set(k, (manifestPassCount.get(k) || 0) + 1);
    });

    const used = new Set();
    const rows = [];

    manifestRows.forEach((m, idx) => {
      const issues = [];
      const pk = normPass(m.passport);
      const exp = parseDate(m.expiry);
      const row = {
        id: 'M' + idx,
        source: 'Manifest',
        no: idx + 1,
        regCode: m.regCode || '',
        trxCode: m.trxCode || '',
        nameManifest: m.name || '',
        passportManifest: clean(m.passport),
        expiry: fmtDate(exp),
        nameMutamer: '',
        passportMutamer: '',
        passportMatch: 'Tidak',
        nameMatch: '-',
        nameScore: null,
        visaNo: '',
        mofaNo: '',
        visaStatus: '',
        category: '',
        issues,
      };

      if (!pk) {
        issues.push('Nomor paspor kosong di Manifest');
      } else if (manifestPassCount.get(pk) > 1) {
        issues.push('Nomor paspor ganda di Manifest');
      }

      const hits = pk ? byPass.get(pk) : null;
      if (hits && hits.length) {
        const u = mutamerRows[hits[0]];
        hits.forEach((h) => used.add(h));
        row.passportMatch = 'Ya';
        row.nameMutamer = u.name;
        row.passportMutamer = u.passport;
        row.visaNo = u.visaNo || '';
        row.mofaNo = u.mofaNo || '';
        row.visaStatus = u.visaStatus || '';
        const nm = compareNames(m.name, u.name);
        row.nameMatch = nm.status;
        row.nameScore = nm.score;
        if (hits.length > 1) issues.push('Nomor paspor ganda di Mutamer List');
        if (nm.status === 'Beda') issues.push(`Nama berbeda dengan Mutamer List (${u.name})`);
        else if (nm.status === 'Mirip') issues.push(`Nama mirip, perlu dicek (${u.name})`);
        if (isPrinted(row.visaStatus)) {
          row.category = CAT.PRINTED;
          if (!row.visaNo) issues.push('Visa Printed tetapi nomor visa kosong');
        } else {
          row.category = CAT.BELUM;
          issues.push(`Status visa: ${row.visaStatus || 'kosong'}`);
        }
      } else {
        row.category = CAT.TIDAK_MATCH;
        row.visaStatus = 'Tidak ditemukan';
        row.nameMatch = '-';
        issues.push('Nomor paspor tidak ditemukan di Mutamer List');
        // kandidat: nama sama tetapi paspor beda
        const nmM = normName(m.name);
        let cand = -1;
        mutamerRows.forEach((u, i) => {
          if (cand >= 0 || used.has(i)) return;
          if (nmM && compareNames(m.name, u.name).status !== 'Beda') cand = i;
        });
        if (cand >= 0) {
          used.add(cand);
          const u = mutamerRows[cand];
          row.nameMutamer = u.name;
          row.passportMutamer = u.passport;
          issues.push(`Kandidat nama sama di Mutamer List dengan paspor ${u.passport || '-'}`);
        }
      }

      if (exp) {
        if (exp < today) issues.push(`Paspor sudah kedaluwarsa (${fmtDate(exp)})`);
        else if (exp < limit) issues.push(`Paspor berlaku kurang dari 6 bulan (${fmtDate(exp)})`);
      }
      rows.push(row);
    });

    // Ada di Mutamer List tetapi tidak di Manifest
    let n = manifestRows.length;
    mutamerRows.forEach((u, i) => {
      if (used.has(i)) return;
      n += 1;
      rows.push({
        id: 'U' + i,
        source: 'Mutamer List saja',
        no: n,
        regCode: '',
        trxCode: '',
        nameManifest: '',
        passportManifest: '',
        expiry: '',
        nameMutamer: u.name,
        passportMutamer: u.passport,
        passportMatch: 'Tidak',
        nameMatch: '-',
        nameScore: null,
        visaNo: u.visaNo || '',
        mofaNo: u.mofaNo || '',
        visaStatus: u.visaStatus || '',
        category: CAT.EXTRA,
        issues: ['Ada di Mutamer List tetapi tidak ada di Manifest'],
      });
    });

    rows.forEach((r) => { r.problem = r.issues.length > 0; });
    return { rows, summary: summarize(rows, manifestRows.length) };
  }

  function summarize(rows, totalManifest) {
    const man = rows.filter((r) => r.source === 'Manifest');
    const count = (fn) => man.filter(fn).length;
    const s = {
      total: totalManifest,
      printed: count((r) => r.category === CAT.PRINTED),
      belum: count((r) => r.category !== CAT.PRINTED),
      tidakMatch: count((r) => r.category === CAT.TIDAK_MATCH),
      extra: rows.filter((r) => r.category === CAT.EXTRA).length,
      bermasalah: rows.filter((r) => r.problem).length,
      // empat pemeriksaan rekonsiliasi
      passportMatch: count((r) => r.passportMatch === 'Ya'),
      nameMatch: count((r) => r.nameMatch === 'Sama'),
      nameSimilar: count((r) => r.nameMatch === 'Mirip'),
      nameDiff: count((r) => r.nameMatch === 'Beda'),
      visaNumber: count((r) => !!r.visaNo),
      visaPrinted: count((r) => r.category === CAT.PRINTED),
    };
    s.belumTermasukTidakMatch = s.tidakMatch;
    return s;
  }

  /* ---------- stempel waktu cetak ---------- */

  const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  function stamp(date) {
    const d = date || new Date();
    let tz = '';
    try {
      tz = new Intl.DateTimeFormat('id-ID', { timeZoneName: 'short' })
        .formatToParts(d).find((p) => p.type === 'timeZoneName').value;
    } catch (e) { /* abaikan */ }
    const p = (n) => String(n).padStart(2, '0');
    const text =
      `${d.getDate()} ${BULAN[d.getMonth()]} ${d.getFullYear()}, ` +
      `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}${tz ? ' ' + tz : ''}`;
    const file = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
    return { date: d, text, file };
  }


  /* ---------- kolom laporan (dipakai ekspor dan impor) ---------- */

  // [judul, field, lebar, tersembunyi]
  const REPORT_COLS = [
    ['No', 'no', 6],
    ['Kode Registrasi', 'regCode', 15],
    ['Nama (Manifest)', 'nameManifest', 30],
    ['No Paspor (Manifest)', 'passportManifest', 17],
    ['Nama (Mutamer List)', 'nameMutamer', 30],
    ['No Paspor (Mutamer List)', 'passportMutamer', 18],
    ['Passport Match', 'passportMatch', 13],
    ['Name Match', 'nameMatch', 12],
    ['Visa Number', 'visaNo', 14],
    ['Mofa Number', 'mofaNo', 13],
    ['Status Visa', 'visaStatus', 16],
    ['Kategori', 'category', 21],
    ['Catatan Sistem', 'issuesText', 48],
    ['Action Need', 'needText', 12],
    ['Alasan / Log Aksi', 'reason', 44],
    ['Waktu Log', 'atText', 26],
    ['ID', 'id', 8, true],
    ['Log ISO', 'atISO', 8, true],
  ];

  /* ---------- pesan WhatsApp ---------- */

  const VISA_LINK = 'https://visa.mofa.gov.sa/visaservices/searchvisa';
  const VISA_STEPS = [
    'Klik tombol *E* di pojok kiri atas untuk mengubah bahasa ke Inggris.',
    '*Device Type*: pilih *Barcode Reader*.',
    '*First Value*: pilih *Passport Number*, lalu ketik nomor paspor jamaah.',
    '*Second Value*: pilih *First Name*, lalu isi nama depan jamaah.',
    '*Nationality*: pilih *Indonesia*.',
    'Isi captcha setelah semua kolom terisi.',
    'Klik tombol *Inquire*.',
  ];

  const firstName = (n) => clean(n).split(/\s+/)[0] || '';

  // Teks siap tempel ke WhatsApp. Selalu memuat tautan dan tutorial cek visa.
  function waMessage(rows, ctx, when) {
    ctx = ctx || {};
    const st = stamp(when);
    const L = [
      '*HAMDAN TOUR - STATUS VISA*',
      'KPU: ' + (ctx.kpu || '-'),
      'Keberangkatan: ' + (ctx.departure || '-'),
      'Dicetak: ' + st.text,
      'Jumlah: ' + rows.length + ' jamaah',
      '',
    ];
    rows.forEach((r, i) => {
      const shown = r.nameManifest || r.nameMutamer || '-';
      const sys = r.passportMatch === 'Ya' && r.nameMutamer ? r.nameMutamer : shown;
      L.push(`${i + 1}. *${shown}*`);
      L.push('Paspor: ' + (r.passportManifest || r.passportMutamer || '-'));
      if (normName(sys) !== normName(shown)) L.push('Nama di sistem visa: ' + sys);
      L.push('Nama depan: ' + (firstName(sys) || '-'));
      L.push('Visa: ' + (r.visaNo || '-'));
      L.push('Status: ' + (r.visaStatus || '-'));
      L.push('');
    });
    L.push('*Cara cek visa sendiri*');
    L.push(VISA_LINK);
    L.push('');
    VISA_STEPS.forEach((t, i) => L.push(`${i + 1}. ${t}`));
    return L.join('\n');
  }

  return {
    REPORT_COLS, VISA_LINK, waMessage, firstName,
    clean, normPass, normName, compareNames, parseDate, fmtDate,
    parseSheet, reconcile, summarize, stamp, CAT, isPrinted,
  };
});
