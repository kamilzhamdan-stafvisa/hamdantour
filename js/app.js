/*
 * Hamdan Tour Check — antarmuka.
 * Semua state hanya di memori halaman ini. Tidak memakai localStorage, cookie, IndexedDB, atau jaringan.
 */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const esc = (s) =>
    String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const CAT = HTC.CAT;
  const PAGE = 100;

  const state = {
    manifest: null,  // { name, rows }
    mutamer: null,
    result: null,    // { rows, summary, partial? }
    imported: null,  // { name, kind, partial, printed } bila dibuka dari laporan lama
    selected: new Set(), // id jamaah yang dipilih untuk disalin ke WhatsApp
    actions: {},     // { [rowId]: { need, reason, at } }
    filter: 'all',
    q: '',
    shown: PAGE,
  };

  const QUICK = ['Menunggu proses visa', 'Hubungi jamaah', 'Paspor perlu dikoreksi', 'Cek ulang nama', 'Minta paspor baru'];

  /* ---------- upload ---------- */

  function parseWb(wb, kind) {
    let lastErr = null;
    for (const sn of wb.SheetNames) {
      const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
      try {
        return HTC.parseSheet(matrix, kind);
      } catch (e) { lastErr = e; }
    }
    // Mungkin file tertukar — beri petunjuk.
    const other = kind === 'manifest' ? 'mutamer' : 'manifest';
    for (const sn of wb.SheetNames) {
      const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
      try {
        HTC.parseSheet(matrix, other);
        throw new Error(
          `${lastErr.message} File ini tampak seperti ${other === 'manifest' ? 'Manifest' : 'Mutamer List'}; coba unggah di kotak yang satunya.`
        );
      } catch (e) {
        if (/tampak seperti/.test(e.message)) throw e;
      }
    }
    throw lastErr || new Error('File tidak bisa dibaca.');
  }

  const BOX = {
    manifest: ['dropManifest', 'stateManifest'],
    mutamer: ['dropMutamer', 'stateMutamer'],
    report: ['dropReport', 'stateReport'],
  };
  const IDLE = { manifest: 'Tarik file atau klik', mutamer: 'Tarik file atau klik', report: 'Excel hasil unduhan dari sini' };

  function boxReset(kind) {
    $(BOX[kind][0]).classList.remove('ok', 'fail');
    $(BOX[kind][1]).textContent = IDLE[kind];
  }
  function boxOk(kind, name, note) {
    $(BOX[kind][0]).classList.remove('fail'); $(BOX[kind][0]).classList.add('ok');
    $(BOX[kind][1]).innerHTML = `${esc(name)}<small>${esc(note)}</small>`;
  }
  function boxFail(kind, msg) {
    $(BOX[kind][0]).classList.remove('ok'); $(BOX[kind][0]).classList.add('fail');
    $(BOX[kind][1]).textContent = 'Ditolak';
    $('uploadErr').textContent = msg;
    $('uploadErr').hidden = false;
  }

  const okToReplace = () => !hasActions() || confirm('Catatan Action need yang ada sekarang akan terhapus. Lanjutkan?');

  function resetView() {
    state.filter = 'all'; state.q = ''; state.shown = PAGE;
    state.selected = new Set();
    $('search').value = '';
  }

  // Buka laporan Excel lama: isi KPU, keberangkatan, tabel, dan catatan Action need.
  function applyReport(rep, file) {
    if (!okToReplace()) return;
    state.manifest = state.mutamer = null;
    boxReset('manifest'); boxReset('mutamer');
    state.imported = { name: file.name, kind: rep.kind, partial: rep.partial, printed: rep.printed };
    state.actions = rep.actions;
    state.downloaded = false;
    $('kpu').value = rep.kpu; $('dep').value = rep.departure;
    ['kpu', 'dep'].forEach((id) => { $(id).classList.remove('bad'); $(id + 'Err').hidden = true; });
    state.result = { rows: rep.rows, summary: rep.summary, partial: rep.partial };
    boxOk('report', file.name, `${rep.rows.length} baris${rep.printed ? ' · dicetak ' + HTC.stamp(rep.printed).text : ''}`);
    resetView();
    renderAll();
    $('s4').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function onFile(kind, file) {
    if (!file) return;
    $('uploadErr').hidden = true;
    let wb;
    try {
      wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    } catch (e) {
      boxFail(kind, `File ${file.name} tidak bisa dibaca.`);
      return;
    }

    // Laporan Hamdan Tour Check dikenali otomatis, di kotak mana pun dijatuhkan.
    let rep = null;
    try { rep = HTCImport.fromWorkbook(wb, XLSX); } catch (e) { boxFail('report', e.message); return; }
    if (rep) { applyReport(rep, file); return; }
    if (kind === 'report') { boxFail('report', 'File ini bukan laporan Hamdan Tour Check.'); return; }

    if (!okToReplace()) return;
    const hadReal = !!(state.manifest || state.mutamer);
    try {
      const parsed = parseWb(wb, kind);
      if (!parsed.rows.length) throw new Error(`File ${file.name} tidak berisi baris data.`);
      state[kind] = { name: file.name, rows: parsed.rows };
      boxOk(kind, file.name, `${parsed.rows.length} baris`);
    } catch (e) {
      state[kind] = null;
      boxFail(kind, e.message);
      if (!hadReal && state.imported) return; // jangan hapus laporan lama yang sedang dibuka
    }
    state.imported = null; boxReset('report');
    state.actions = {};
    recompute();
  }

  function wireDrop(kind, boxId, inputId) {
    const box = $(boxId), input = $(inputId);
    input.addEventListener('change', () => { onFile(kind, input.files[0]); input.value = ''; });
    ['dragenter', 'dragover'].forEach((ev) => box.addEventListener(ev, (e) => { e.preventDefault(); box.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => box.addEventListener(ev, (e) => { e.preventDefault(); box.classList.remove('over'); }));
    box.addEventListener('drop', (e) => onFile(kind, e.dataTransfer.files[0]));
  }

  /* ---------- hitung ---------- */

  function recompute() {
    if (state.manifest && state.mutamer) {
      state.result = HTC.reconcile(state.manifest.rows, state.mutamer.rows, { today: new Date() });
    } else {
      state.result = null;
    }
    resetView();
    renderAll();
  }

  const hasActions = () => Object.values(state.actions).some((a) => a.need);

  /* ---------- render ---------- */

  function renderAll() {
    const r = state.result;
    ['reconEmpty', 'cpEmpty', 'outEmpty'].forEach((id) => { $(id).hidden = !!r; });
    ['reconBody', 'cpBody', 'outBody'].forEach((id) => { $(id).hidden = !r; });
    ['s3', 's4', 's5'].forEach((id) => $(id).classList.toggle('idle', !r));
    updateDock();
    renderImpNote();
    renderSel();
    if (!r) return;
    renderChecks(r.summary);
    renderBand();
    renderTable();
    updateWarn();
  }

  function updateDock() {
    const done = {
      1: !!($('kpu').value.trim() && $('dep').value.trim()),
      2: !!(state.manifest && state.mutamer) || !!state.imported,
      3: !!state.result, 4: !!state.result, 5: !!state.downloaded,
    };
    document.querySelectorAll('.dock a[data-step]').forEach((a) => a.classList.toggle('done', !!done[a.dataset.step]));
  }

  function renderImpNote() {
    const el = $('impNote');
    const im = state.imported;
    el.hidden = !im;
    if (!im) return;
    el.innerHTML = `Dibuka dari laporan <b>${esc(im.name)}</b><span>${im.printed ? ' · dicetak ' + esc(HTC.stamp(im.printed).text) : ''}${im.partial ? ' · hanya memuat data bermasalah' : ''}</span>`;
  }

  function renderChecks(s) {
    const t = s.total || 1;
    const items = [
      ['Passport match', s.passportMatch, '', s.passportMatch < s.total],
      ['Name match', s.nameMatch, s.nameSimilar + s.nameDiff ? `Mirip ${s.nameSimilar}, beda ${s.nameDiff}` : '', s.nameMatch < s.total],
      ['Visa number', s.visaNumber, '', s.visaNumber < s.total],
      ['Visa status', s.visaPrinted, '', s.visaPrinted < s.total],
    ];
    $('checks').innerHTML = items.map(([k, n, note, warn]) =>
      `<div class="check ${warn ? 'warn' : ''}"><div class="k">${k}</div>
        <div class="v">${n}<small> / ${s.total}</small></div>
        <div class="bar"><i style="width:${Math.round((n / t) * 100)}%"></i></div>
        <div class="n">${esc(note)}</div></div>`).join('');
  }

  function actionCount() { return Object.values(state.actions).filter((a) => a.need).length; }

  function renderBand() {
    const s = state.result.summary;
    const t = s.total || 1;
    const matchedBelum = s.belum - s.tidakMatch;
    const defs = [
      ['all', 'Semua', s.total, ''],
      ['printed', 'Visa printed', s.printed, 'a'],
      ['belum', 'Belum visa', s.belum, 'b'],
      ['tidak', 'Tidak match', s.tidakMatch, 'c'],
      ['problem', 'Data bermasalah', s.bermasalah, 'c'],
      ['action', 'Action need', actionCount(), 'o'],
    ];
    $('band').innerHTML =
      `<div class="total"><b>${s.total}</b><span>Total jamaah</span></div>
       <div class="split" aria-hidden="true">
         <i class="a" style="flex:${s.printed}"></i><i class="b" style="flex:${matchedBelum}"></i><i class="c" style="flex:${s.tidakMatch}"></i>
       </div>
       <div class="filters" role="group" aria-label="Filter">` +
      defs.map(([k, l, v, c]) =>
        `<button type="button" class="fbtn" data-f="${k}" data-c="${c}" aria-pressed="${state.filter === k}">${l}<b data-v="${k}">${v}</b></button>`).join('') +
      `</div>`;
  }

  function visibleRows() {
    const q = state.q.trim().toLowerCase();
    return state.result.rows.filter((r) => {
      switch (state.filter) {
        case 'printed': if (r.category !== CAT.PRINTED) return false; break;
        case 'belum': if (r.source !== 'Manifest' || r.category === CAT.PRINTED) return false; break;
        case 'tidak': if (r.category !== CAT.TIDAK_MATCH) return false; break;
        case 'problem': if (!r.problem) return false; break;
        case 'action': if (!(state.actions[r.id] && state.actions[r.id].need)) return false; break;
        default: break;
      }
      if (!q) return true;
      return [r.nameManifest, r.nameMutamer, r.passportManifest, r.passportMutamer, r.visaNo, r.regCode]
        .some((v) => String(v || '').toLowerCase().includes(q));
    });
  }

  const tagClass = (c) => ({ [CAT.PRINTED]: 't-printed', [CAT.BELUM]: 't-belum', [CAT.TIDAK_MATCH]: 't-tidak', [CAT.EXTRA]: 't-extra' }[c]);

  const mk = (v) => {
    if (v === 'Ya' || v === 'Sama') return `<span class="mk y" role="img" aria-label="${v}">&#10003;</span>`;
    if (v === 'Mirip') return `<span class="mk m" role="img" aria-label="Mirip">&asymp;</span>`;
    if (v === 'Tidak' || v === 'Beda') return `<span class="mk n" role="img" aria-label="${v}">&#10005;</span>`;
    return `<span class="mk x" aria-label="Tidak ada">&ndash;</span>`;
  };

  function noteRowHtml(r) {
    const a = state.actions[r.id];
    const empty = !a.reason.trim();
    return `<tr class="note has-need" data-note="${r.id}"><td></td><td></td><td colspan="9"><div class="note-box">
      <textarea id="n-${r.id}" data-id="${r.id}" aria-label="Alasan dan log" placeholder="Alasan / log tindak lanjut">${esc(a.reason)}</textarea>
      <div class="chips">${QUICK.map((q) => `<button type="button" class="chip" data-id="${r.id}" data-q="${esc(q)}">${esc(q)}</button>`).join('')}</div>
      <div class="note-meta"><span>Dicatat <b data-at="${r.id}">${esc(HTC.stamp(a.at).text)}</b></span>
        <span class="need" data-empty="${r.id}" ${empty ? '' : 'hidden'}>Alasan belum diisi</span></div>
    </div></td></tr>`;
  }

  function rowHtml(r) {
    const a = state.actions[r.id];
    const need = !!(a && a.need);
    const nm = r.nameManifest || r.nameMutamer;
    const nameCell = r.source === 'Manifest'
      ? `${esc(r.nameManifest)}${r.nameMutamer && r.nameMatch !== 'Sama' ? `<span class="sub">Mutamer: ${esc(r.nameMutamer)}</span>` : ''}`
      : `${esc(r.nameMutamer)}<span class="sub">Hanya di Mutamer List</span>`;
    const sel = state.selected.has(r.id);
    return `<tr class="main ${need ? 'has-need' : ''} ${sel ? 'is-sel' : ''}" data-row="${r.id}">
      <td class="c-sel"><input type="checkbox" data-sel="${r.id}" ${sel ? 'checked' : ''} aria-label="Pilih ${esc(nm)}"></td>
      <td class="c-act"><input type="checkbox" data-id="${r.id}" ${need ? 'checked' : ''} aria-label="Action need: ${esc(nm)}"></td>
      <td class="num">${r.no}</td>
      <td>${nameCell}</td>
      <td class="mono">${esc(r.passportManifest || r.passportMutamer)}</td>
      <td class="c-ico">${mk(r.passportMatch)}</td>
      <td class="c-ico">${mk(r.nameMatch)}</td>
      <td class="mono">${esc(r.visaNo)}</td>
      <td>${esc(r.visaStatus)}</td>
      <td><span class="tag ${tagClass(r.category)}">${esc(r.category)}</span></td>
      <td class="issues">${r.issues.map((i) => `<div>${esc(i)}</div>`).join('')}</td>
    </tr>${need ? noteRowHtml(r) : ''}`;
  }

  function renderTable() {
    const rows = visibleRows();
    const slice = rows.slice(0, state.shown);
    $('tbody').innerHTML = slice.length
      ? slice.map(rowHtml).join('')
      : `<tr><td colspan="11" class="num" style="padding:18px">Tidak ada baris.</td></tr>`;
    $('rowCount').textContent = `${slice.length} dari ${rows.length} baris`;
    const more = $('more');
    more.hidden = rows.length <= slice.length;
    more.textContent = `Tampilkan ${Math.min(PAGE, rows.length - slice.length)} lagi`;
    syncSelAll();
  }

  /* ---------- pilih jamaah + salin ke WhatsApp ---------- */

  function syncSelAll() {
    const box = $('selAll');
    const vis = state.result ? visibleRows() : [];
    const n = vis.filter((r) => state.selected.has(r.id)).length;
    box.checked = vis.length > 0 && n === vis.length;
    box.indeterminate = n > 0 && n < vis.length;
  }

  const selectedRows = () => state.result.rows.filter((r) => state.selected.has(r.id));
  const looseCtx = () => ({ kpu: $('kpu').value.trim(), departure: $('dep').value.trim() });

  function renderBubble(text) {
    $('bubble').innerHTML = esc(text)
      .replace(/\*([^*\n]+)\*/g, '<b>$1</b>')
      .replace(/(https:\/\/[^\s<]+)/g, '<span class="lnk">$1</span>');
  }

  function renderSel() {
    const n = state.result ? state.selected.size : 0;
    $('sel').hidden = n === 0;
    document.body.classList.toggle('has-sel', n > 0);
    if (!n) return;
    $('selCount').textContent = `${n} jamaah dipilih`;
    renderBubble(HTC.waMessage(selectedRows(), looseCtx(), new Date()));
  }

  async function copyText(text) {
    try {
      if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(text); return true; }
    } catch (e) { /* lanjut ke cadangan */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(ta); ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    return ok;
  }

  $('selCopy').addEventListener('click', async () => {
    const btn = $('selCopy'), label = btn.querySelector('span');
    const text = HTC.waMessage(selectedRows(), looseCtx(), new Date());
    renderBubble(text);
    const ok = await copyText(text);
    label.textContent = ok ? 'Tersalin' : 'Gagal, salin manual dari kotak hijau';
    btn.classList.toggle('done', ok);
    clearTimeout(btn._t);
    btn._t = setTimeout(() => { label.textContent = 'Salin ke WhatsApp'; btn.classList.remove('done'); }, 2400);
  });
  $('selClear').addEventListener('click', () => { state.selected = new Set(); renderTable(); renderSel(); });
  $('selToggle').addEventListener('click', () => {
    const body = $('selBody');
    body.hidden = !body.hidden;
    $('selToggle').textContent = body.hidden ? 'Tampilkan' : 'Sembunyikan';
    $('selToggle').setAttribute('aria-expanded', String(!body.hidden));
  });
  $('selAll').addEventListener('change', (e) => {
    visibleRows().forEach((r) => { if (e.target.checked) state.selected.add(r.id); else state.selected.delete(r.id); });
    renderTable(); renderSel();
  });

  function updateCounters() {
    const el = document.querySelector('[data-v="action"]');
    if (el) el.textContent = actionCount();
    updateWarn();
  }
  function updateWarn() { $('noteWarn').hidden = !hasActions(); }

  /* ---------- aksi di tabel ---------- */

  $('band').addEventListener('click', (e) => {
    const b = e.target.closest('.fbtn');
    if (!b) return;
    state.filter = state.filter === b.dataset.f && b.dataset.f !== 'all' ? 'all' : b.dataset.f;
    state.shown = PAGE;
    renderBand(); renderTable();
  });
  $('search').addEventListener('input', (e) => { state.q = e.target.value; state.shown = PAGE; renderTable(); });
  $('more').addEventListener('click', () => { state.shown += PAGE; renderTable(); });

  $('tbody').addEventListener('change', (e) => {
    const cb = e.target.closest('input[type="checkbox"]');
    if (!cb) return;
    if (cb.dataset.sel !== undefined) {
      if (cb.checked) state.selected.add(cb.dataset.sel); else state.selected.delete(cb.dataset.sel);
      cb.closest('tr').classList.toggle('is-sel', cb.checked);
      renderSel(); syncSelAll();
      return;
    }
    const id = cb.dataset.id;
    const cur = state.actions[id] || { need: false, reason: '', at: new Date() };
    cur.need = cb.checked;
    if (cb.checked) cur.at = new Date();
    state.actions[id] = cur;
    const tr = cb.closest('tr');
    const r = state.result.rows.find((x) => x.id === id);
    tr.classList.toggle('has-need', cb.checked);
    const existing = $('tbody').querySelector(`tr[data-note="${id}"]`);
    if (existing) existing.remove();
    if (cb.checked) {
      tr.insertAdjacentHTML('afterend', noteRowHtml(r));
      const ta = $('n-' + id); if (ta) ta.focus();
    } else if (state.filter === 'action') {
      renderTable();
    }
    updateCounters();
  });

  function setReason(id, text) {
    const a = state.actions[id];
    if (!a) return;
    a.reason = text; a.at = new Date();
    const at = document.querySelector(`[data-at="${id}"]`);
    if (at) at.textContent = HTC.stamp(a.at).text;
    const em = document.querySelector(`[data-empty="${id}"]`);
    if (em) em.hidden = !!text.trim();
  }
  $('tbody').addEventListener('input', (e) => {
    if (e.target.matches('textarea[data-id]')) setReason(e.target.dataset.id, e.target.value);
  });
  $('tbody').addEventListener('click', (e) => {
    const chip = e.target.closest('.chip');
    if (!chip) return;
    const id = chip.dataset.id;
    const ta = $('n-' + id);
    const sep = ta.value.trim() ? '; ' : '';
    ta.value = ta.value.trim() + sep + chip.dataset.q;
    setReason(id, ta.value);
    ta.focus();
  });

  /* ---------- output ---------- */

  function context() {
    const kpu = $('kpu').value.trim(), dep = $('dep').value.trim();
    $('kpu').classList.toggle('bad', !kpu); $('dep').classList.toggle('bad', !dep);
    $('kpuErr').hidden = !!kpu; $('depErr').hidden = !!dep;
    if (!kpu || !dep) {
      $('s1').scrollIntoView({ behavior: 'smooth', block: 'center' });
      (kpu ? $('dep') : $('kpu')).focus({ preventScroll: true });
      return null;
    }
    return { kpu, departure: dep };
  }

  function download(out) {
    const blob = new Blob([out.buffer], { type: out.mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = out.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  async function run(btn, fn) {
    const ctx = context();
    if (!ctx) return;
    $('outErr').hidden = true;
    btn.disabled = true;
    const label = btn.querySelector('strong').textContent;
    btn.querySelector('strong').textContent = 'Menyiapkan…';
    try {
      // setTimeout agar teks "Menyiapkan..." sempat tampil sebelum proses berat.
      await new Promise((r) => setTimeout(r, 30));
      download(await fn(ctx, new Date()));
      state.downloaded = true; updateDock();
    } catch (e) {
      console.error(e);
      $('outErr').textContent = 'Gagal membuat laporan: ' + e.message;
      $('outErr').hidden = false;
    } finally {
      btn.disabled = false;
      btn.querySelector('strong').textContent = label;
    }
  }

  $('outDetail').addEventListener('click', (e) => run(e.currentTarget, (ctx, when) => HTCExport.buildExcel('detail', ctx, state.result, state.actions, when)));
  $('outProblem').addEventListener('click', (e) => run(e.currentTarget, (ctx, when) => HTCExport.buildExcel('bermasalah', ctx, state.result, state.actions, when)));
  $('outPdf').addEventListener('click', (e) => run(e.currentTarget, async (ctx, when) => HTCExport.buildPdf(ctx, state.result, state.actions, when)));

  /* ---------- lain-lain ---------- */

  ['kpu', 'dep'].forEach((id) => $(id).addEventListener('input', () => {
    $(id).classList.remove('bad'); $(id + 'Err').hidden = true; updateDock(); renderSel();
  }));

  $('btnReset').addEventListener('click', () => {
    if ((state.manifest || state.mutamer || state.imported) && !confirm('Hapus semua data dan catatan?')) return;
    state.manifest = state.mutamer = null; state.imported = null; state.actions = {}; state.downloaded = false;
    ['manifest', 'mutamer', 'report'].forEach(boxReset);
    $('uploadErr').hidden = true;
    $('kpu').value = ''; $('dep').value = '';
    recompute();
    window.scrollTo({ top: 0 });
  });

  window.addEventListener('beforeunload', (e) => {
    if (hasActions()) { e.preventDefault(); e.returnValue = ''; }
  });

  // stempel waktu hidup di bagian Output
  const live = $('liveStamp');
  const tick = () => { live.textContent = HTC.stamp(new Date()).text; };
  tick(); setInterval(tick, 1000);

  // penanda langkah aktif di navigasi
  const links = [...document.querySelectorAll('.dock a[data-step]')];
  const io = new IntersectionObserver((entries) => {
    entries.forEach((en) => {
      if (en.isIntersecting) links.forEach((a) => a.classList.toggle('on', a.getAttribute('href') === '#' + en.target.id));
    });
  }, { rootMargin: '-25% 0px -60% 0px' });
  document.querySelectorAll('main .sec').forEach((p) => io.observe(p));

  wireDrop('manifest', 'dropManifest', 'fileManifest');
  wireDrop('mutamer', 'dropMutamer', 'fileMutamer');
  wireDrop('report', 'dropReport', 'fileReport');
  renderAll();
})();
