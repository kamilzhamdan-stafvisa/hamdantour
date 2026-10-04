/*
 * Hamdan Tour Check — pembuat laporan Excel (ExcelJS) dan PDF (jsPDF).
 * Semua dibuat di memori browser lalu diunduh; tidak ada yang dikirim atau disimpan.
 * Logo Hamdan dan stempel waktu cetak ada di setiap sheet Excel dan setiap halaman PDF.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.HTCExport = factory(root.HTC);
})(typeof self !== 'undefined' ? self : this, function (HTC) {
  'use strict';

  const G = typeof self !== 'undefined' ? self : global;
  const GREEN = '0B6735';
  const GREEN_RGB = [9, 103, 53];
  const ORANGE = 'F08A1C';
  const ORANGE_RGB = [240, 138, 28];

  const CAT_FILL = {
    'Visa Printed': 'DCF2E3',
    'Belum Visa': 'FDE9CC',
    'Tidak Match': 'F9D5D3',
    'Hanya di Mutamer List': 'E6E3F4',
  };
  const CAT_RGB = {
    'Visa Printed': [220, 242, 227],
    'Belum Visa': [253, 233, 204],
    'Tidak Match': [249, 213, 211],
    'Hanya di Mutamer List': [230, 227, 244],
  };

  const colLetter = (n) => {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  const safe = (s) => String(s || '').replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '');

  function fileName(ctx, label, ext, st) {
    const kpu = safe(ctx.kpu) || 'TanpaKPU';
    return `HamdanTourCheck_${kpu}_${label}_${st.file}.${ext}`;
  }

  // Siapkan baris + catatan aksi menjadi satu bentuk datar.
  function decorate(rows, actions) {
    return rows.map((r) => {
      const a = actions[r.id];
      const need = !!(a && a.need);
      return Object.assign({}, r, {
        issuesText: r.issues.join('\n'),
        needText: need ? 'YA' : '',
        reason: need ? (a.reason || '(alasan belum diisi)') : '',
        atText: need && a.at ? HTC.stamp(a.at).text : '',
        atISO: need && a.at ? new Date(a.at).toISOString() : '',
      });
    });
  }

  function actionList(rows, actions) {
    return decorate(rows, actions).filter((r) => r.needText === 'YA');
  }

  /* ======================= EXCEL ======================= */

  const DETAIL_COLS = HTC.REPORT_COLS;

  function brandHeader(wb, ws, ctx, st, title, lastCol) {
    const imgId = wb.addImage({ base64: G.HAMDAN_LOGO || ctx.logo, extension: 'png' });
    for (let r = 1; r <= 4; r++) ws.getRow(r).height = 21;
    ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 150, height: 80 } });
    const L = colLetter(lastCol);
    const put = (row, text, font) => {
      ws.mergeCells(`C${row}:${L}${row}`);
      const c = ws.getCell(`C${row}`);
      c.value = text;
      c.font = font;
      c.alignment = { vertical: 'middle' };
    };
    put(1, 'HAMDAN TOUR CHECK', { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF' + GREEN } });
    put(2, title, { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF333333' } });
    put(3, `Kode KPU: ${ctx.kpu || '-'}    |    Keberangkatan: ${ctx.departure || '-'}`, { name: 'Calibri', size: 11, color: { argb: 'FF333333' } });
    put(4, `Dicetak: ${st.text}`, { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF' + ORANGE } });
    ws.getColumn(1).width = Math.max(ws.getColumn(1).width || 0, 8);
    ws.getColumn(2).width = Math.max(ws.getColumn(2).width || 0, 15);
    // garis hijau di bawah kepala
    for (let c = 1; c <= lastCol; c++) {
      ws.getCell(5, c).border = { top: { style: 'medium', color: { argb: 'FF' + GREEN } } };
    }
    ws.getRow(5).height = 6;
    ws.headerFooter.oddFooter = `&L&8Hamdan Tour Check - Dicetak: ${st.text}&R&8Halaman &P dari &N`;
    ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.3, footer: 0.3 } };
  }

  function headerRow(ws, rowNo, labels) {
    const row = ws.getRow(rowNo);
    labels.forEach((l, i) => {
      const c = row.getCell(i + 1);
      c.value = l;
      c.font = { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GREEN } };
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      c.border = { bottom: { style: 'thin', color: { argb: 'FF' + ORANGE } } };
    });
    row.height = 30;
  }

  const thin = { style: 'thin', color: { argb: 'FFD9D9D9' } };

  function sheetSummary(wb, ctx, st, summary, nAction) {
    const ws = wb.addWorksheet('Ringkasan', { views: [{ showGridLines: false }] });
    ws.getColumn(1).width = 8; ws.getColumn(2).width = 12;
    ws.getColumn(3).width = 34; ws.getColumn(4).width = 14; ws.getColumn(5).width = 62;
    brandHeader(wb, ws, ctx, st, 'Ringkasan Rekonsiliasi Visa', 5);
    headerRow(ws, 7, ['', '', 'Control Panel', 'Jumlah', 'Keterangan']);
    const items = [
      ['Total Jamaah', summary.total, 'Jumlah baris pada Manifest'],
      ['Visa Printed', summary.printed, 'Paspor cocok dan status Visa Printed'],
      ['Belum Visa', summary.belum, `Belum berstatus Visa Printed (termasuk ${summary.tidakMatch} tidak match)`],
      ['Tidak Match', summary.tidakMatch, 'Nomor paspor Manifest tidak ada di Mutamer List'],
      ['Data Bermasalah', summary.bermasalah, 'Baris dengan minimal satu catatan sistem' + (summary.extra ? ` (termasuk ${summary.extra} hanya di Mutamer List)` : '')],
      ['Action Need', nAction, 'Baris yang Anda tandai untuk ditindaklanjuti'],
    ];
    items.forEach((it, i) => {
      const r = 8 + i;
      ws.getCell(r, 3).value = it[0];
      ws.getCell(r, 4).value = it[1];
      ws.getCell(r, 5).value = it[2];
      ws.getCell(r, 3).font = { bold: true };
      ws.getCell(r, 4).font = { bold: true, size: 13, color: { argb: 'FF' + GREEN } };
      ws.getCell(r, 4).alignment = { horizontal: 'center' };
      for (let c = 3; c <= 5; c++) ws.getCell(r, c).border = { bottom: thin };
    });
    const r0 = 8 + items.length + 1;
    headerRow(ws, r0, ['', '', 'Pemeriksaan Rekonsiliasi', 'Hasil', 'Keterangan']);
    const checks = [
      ['Passport Match', `${summary.passportMatch} / ${summary.total}`, 'Nomor paspor Manifest ditemukan di Mutamer List'],
      ['Name Match', `${summary.nameMatch} / ${summary.total}`, `Nama sama persis; mirip: ${summary.nameSimilar}, berbeda: ${summary.nameDiff}`],
      ['Visa Number', `${summary.visaNumber} / ${summary.total}`, 'Nomor visa tersedia'],
      ['Visa Status', `${summary.visaPrinted} / ${summary.total}`, 'Berstatus Visa Printed'],
    ];
    checks.forEach((it, i) => {
      const r = r0 + 1 + i;
      ws.getCell(r, 3).value = it[0];
      ws.getCell(r, 4).value = it[1];
      ws.getCell(r, 5).value = it[2];
      ws.getCell(r, 3).font = { bold: true };
      ws.getCell(r, 4).alignment = { horizontal: 'center' };
      for (let c = 3; c <= 5; c++) ws.getCell(r, c).border = { bottom: thin };
    });
    const rn = r0 + checks.length + 2;
    ws.mergeCells(`C${rn}:E${rn}`);
    ws.getCell(`C${rn}`).value = 'Laporan dibuat di browser. File Manifest dan Mutamer List tidak disimpan oleh sistem.';
    ws.getCell(`C${rn}`).font = { italic: true, color: { argb: 'FF666666' }, size: 10 };
    return ws;
  }

  function sheetData(wb, name, title, ctx, st, data) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 6, xSplit: 3, showGridLines: false }] });
    DETAIL_COLS.forEach((c, i) => { ws.getColumn(i + 1).width = c[2]; if (c[3]) ws.getColumn(i + 1).hidden = true; });
    brandHeader(wb, ws, ctx, st, title, DETAIL_COLS.length);
    headerRow(ws, 6, DETAIL_COLS.map((c) => c[0]));
    data.forEach((r, i) => {
      const row = ws.getRow(7 + i);
      DETAIL_COLS.forEach((c, j) => {
        const cell = row.getCell(j + 1);
        const v = r[c[1]];
        cell.value = v === null || v === undefined ? '' : (c[1] === 'no' ? v : String(v));
        cell.border = { bottom: thin };
        cell.alignment = { vertical: 'top', wrapText: ['issuesText', 'reason'].includes(c[1]) };
        cell.font = { name: 'Calibri', size: 10 };
        if (c[1] === 'category') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + (CAT_FILL[r.category] || 'FFFFFF') } };
          cell.font = { name: 'Calibri', size: 10, bold: true };
        }
        if (c[1] === 'issuesText' && r.issuesText) cell.font = { name: 'Calibri', size: 10, color: { argb: 'FFB3261E' } };
        if (c[1] === 'needText' && r.needText) {
          cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ORANGE } };
          cell.alignment = { horizontal: 'center', vertical: 'top' };
        }
        if ((c[1] === 'passportMatch' || c[1] === 'nameMatch') && v === 'Tidak') cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFB3261E' } };
        if (c[1] === 'nameMatch' && (v === 'Beda' || v === 'Mirip')) cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFB3261E' } };
      });
    });
    ws.autoFilter = { from: { row: 6, column: 1 }, to: { row: 6 + Math.max(data.length, 1), column: DETAIL_COLS.length } };
    ws.pageSetup.printTitlesRow = '6:6';
    return ws;
  }

  function sheetActions(wb, ctx, st, acts) {
    const ws = wb.addWorksheet('Action Need', { views: [{ state: 'frozen', ySplit: 6, showGridLines: false }] });
    const cols = [
      ['No', 'no', 6], ['Nama', 'nameShow', 32], ['No Paspor', 'passShow', 16], ['Kategori', 'category', 21],
      ['Catatan Sistem', 'issuesText', 44], ['Alasan / Log Aksi', 'reason', 52], ['Waktu Log', 'atText', 26],
    ];
    cols.forEach((c, i) => { ws.getColumn(i + 1).width = c[2]; });
    brandHeader(wb, ws, ctx, st, 'Catatan Action Need', cols.length);
    headerRow(ws, 6, cols.map((c) => c[0]));
    if (!acts.length) {
      ws.mergeCells('A7:G7');
      ws.getCell('A7').value = 'Tidak ada baris yang ditandai Action Need.';
      ws.getCell('A7').font = { italic: true, color: { argb: 'FF666666' } };
      return ws;
    }
    acts.forEach((r, i) => {
      const o = Object.assign({}, r, { nameShow: r.nameManifest || r.nameMutamer, passShow: r.passportManifest || r.passportMutamer });
      const row = ws.getRow(7 + i);
      cols.forEach((c, j) => {
        const cell = row.getCell(j + 1);
        cell.value = c[1] === 'no' ? o.no : String(o[c[1]] || '');
        cell.border = { bottom: thin };
        cell.alignment = { vertical: 'top', wrapText: true };
        cell.font = { name: 'Calibri', size: 10 };
        if (c[1] === 'category') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + (CAT_FILL[o.category] || 'FFFFFF') } };
      });
    });
    ws.pageSetup.printTitlesRow = '6:6';
    return ws;
  }

  // Sheet tersembunyi: dipakai saat laporan diunggah kembali.
  function sheetMeta(wb, kind, ctx, st, result) {
    const ws = wb.addWorksheet('_HTC', { state: 'hidden' });
    [
      ['app', 'HamdanTourCheck'],
      ['version', '2'],
      ['kpu', ctx.kpu || ''],
      ['departure', ctx.departure || ''],
      ['printedISO', st.date.toISOString()],
      ['kind', kind],
      ['partial', kind === 'bermasalah' || result.partial ? '1' : '0'],
      ['summary', JSON.stringify(result.summary)],
    ].forEach((r, i) => { ws.getCell(i + 1, 1).value = r[0]; ws.getCell(i + 1, 2).value = r[1]; });
  }

  /**
   * kind: 'detail' | 'bermasalah'
   * ctx: { kpu, departure }, result: { rows, summary }, actions: { [rowId]: {need, reason, at} }
   * Mengembalikan { buffer, filename, mime }
   */
  async function buildExcel(kind, ctx, result, actions, when) {
    const ExcelJS = G.ExcelJS || require('exceljs');
    const st = HTC.stamp(when);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Hamdan Tour Check';
    wb.created = st.date;
    wb.title = 'Hamdan Tour Check - ' + (kind === 'detail' ? 'Detail' : 'Data Bermasalah');

    const all = decorate(result.rows, actions);
    const acts = all.filter((r) => r.needText === 'YA');
    sheetSummary(wb, ctx, st, result.summary, acts.length);
    if (kind === 'detail') {
      sheetData(wb, 'Detail', 'Detail Rekonsiliasi Visa' + (result.partial ? ' (hanya data bermasalah)' : ''), ctx, st, all);
    } else {
      sheetData(wb, 'Data Bermasalah', 'Data Bermasalah', ctx, st, all.filter((r) => r.problem));
    }
    sheetActions(wb, ctx, st, acts);
    sheetMeta(wb, kind, ctx, st, result);
    const buffer = await wb.xlsx.writeBuffer();
    return {
      buffer,
      filename: fileName(ctx, kind === 'detail' ? 'Detail' : 'DataBermasalah', 'xlsx', st),
      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    };
  }

  /* ======================= PDF ======================= */

  function buildPdf(ctx, result, actions, when) {
    const jsPDF = (G.jspdf && G.jspdf.jsPDF) || require('jspdf').jsPDF;
    if (!G.jspdf) require('jspdf-autotable');
    const st = HTC.stamp(when);
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const W = doc.internal.pageSize.getWidth();
    const logo = G.HAMDAN_LOGO || ctx.logo;
    const s = result.summary;
    const all = decorate(result.rows, actions);
    const acts = all.filter((r) => r.needText === 'YA');

    // Ringkasan halaman pertama
    let y = 30;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(0);
    doc.text('Laporan Rekonsiliasi Visa', 10, y);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(60);
    doc.text(`Kode KPU: ${ctx.kpu || '-'}      Keberangkatan: ${ctx.departure || '-'}`, 10, y + 6);
    y += 11;
    const boxes = [
      ['Total Jamaah', s.total, [60, 60, 60]],
      ['Visa Printed', s.printed, GREEN_RGB],
      ['Belum Visa', s.belum, [196, 112, 10]],
      ['Tidak Match', s.tidakMatch, [179, 38, 30]],
      ['Data Bermasalah', s.bermasalah, [179, 38, 30]],
      ['Action Need', acts.length, ORANGE_RGB],
    ];
    const gap = 3;
    const bw = (W - 20 - gap * (boxes.length - 1)) / boxes.length;
    boxes.forEach((b, i) => {
      const x = 10 + i * (bw + gap);
      doc.setDrawColor(200); doc.setFillColor(250, 250, 248);
      doc.roundedRect(x, y, bw, 16, 1.5, 1.5, 'FD');
      doc.setFillColor(...b[2]); doc.rect(x, y, 1.6, 16, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(...b[2]);
      doc.text(String(b[1]), x + 5, y + 8.5);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(80);
      doc.text(b[0], x + 5, y + 13.5);
    });
    y += 21;
    doc.setFontSize(8.5); doc.setTextColor(60);
    doc.text(
      `Passport Match ${s.passportMatch}/${s.total}   |   Name Match ${s.nameMatch}/${s.total} (mirip ${s.nameSimilar}, beda ${s.nameDiff})   |   ` +
        `Visa Number ${s.visaNumber}/${s.total}   |   Visa Status Printed ${s.visaPrinted}/${s.total}` +
        (s.extra ? `   |   Hanya di Mutamer List ${s.extra}` : ''),
      10, y
    );
    y += 4;

    const head = [['No', 'Nama (Manifest)', 'Paspor Manifest', 'Nama (Mutamer List)', 'Paspor Mutamer', 'Passport Match', 'Name Match', 'Visa Number', 'Status Visa', 'Kategori', 'Catatan Sistem']];
    const body = all.map((r) => [
      r.no, r.nameManifest, r.passportManifest, r.nameMutamer, r.passportMutamer,
      r.passportMatch, r.nameMatch, r.visaNo, r.visaStatus, r.category,
      r.issuesText + (r.needText ? `${r.issuesText ? '\n' : ''}[ACTION NEED] ${r.reason}` : ''),
    ]);
    doc.autoTable({
      head, body, startY: y,
      margin: { top: 28, bottom: 16, left: 10, right: 10 },
      styles: { font: 'helvetica', fontSize: 6.8, cellPadding: 1.3, overflow: 'linebreak', valign: 'top' },
      headStyles: { fillColor: GREEN_RGB, textColor: 255, fontStyle: 'bold', halign: 'left' },
      columnStyles: {
        0: { cellWidth: 8 }, 1: { cellWidth: 37 }, 2: { cellWidth: 19 }, 3: { cellWidth: 37 }, 4: { cellWidth: 19 },
        5: { cellWidth: 15 }, 6: { cellWidth: 13 }, 7: { cellWidth: 20 }, 8: { cellWidth: 19 }, 9: { cellWidth: 22 },
      },
      didParseCell: (d) => {
        if (d.section !== 'body') return;
        const r = all[d.row.index];
        if (d.column.index === 9 && CAT_RGB[r.category]) { d.cell.styles.fillColor = CAT_RGB[r.category]; d.cell.styles.fontStyle = 'bold'; }
        if (d.column.index === 10 && r.issuesText) d.cell.styles.textColor = [179, 38, 30];
        if ((d.column.index === 5 || d.column.index === 6) && ['Tidak', 'Beda', 'Mirip'].includes(d.cell.raw)) {
          d.cell.styles.textColor = [179, 38, 30]; d.cell.styles.fontStyle = 'bold';
        }
      },
    });

    // Bagian Action Need
    if (acts.length) {
      let ay = doc.lastAutoTable.finalY + 8;
      if (ay > doc.internal.pageSize.getHeight() - 40) { doc.addPage(); ay = 30; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(0);
      doc.text('Catatan Action Need', 10, ay);
      doc.autoTable({
        startY: ay + 3,
        margin: { top: 28, bottom: 16, left: 10, right: 10 },
        head: [['No', 'Nama', 'Paspor', 'Kategori', 'Alasan / Log Aksi', 'Waktu Log']],
        body: acts.map((r) => [r.no, r.nameManifest || r.nameMutamer, r.passportManifest || r.passportMutamer, r.category, r.reason, r.atText]),
        styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.6, valign: 'top' },
        headStyles: { fillColor: ORANGE_RGB, textColor: 255, fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 10 }, 1: { cellWidth: 55 }, 2: { cellWidth: 25 }, 3: { cellWidth: 28 }, 5: { cellWidth: 45 } },
      });
    }

    // Kepala (logo) dan kaki (stempel cetak, nomor halaman) di setiap halaman
    const total = doc.getNumberOfPages();
    const H = doc.internal.pageSize.getHeight();
    for (let p = 1; p <= total; p++) {
      doc.setPage(p);
      if (logo) doc.addImage(logo, 'PNG', 10, 6, 34, 18.2);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(...GREEN_RGB);
      doc.text('Hamdan Tour Check', 48, 14);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(90);
      doc.text(`KPU ${ctx.kpu || '-'}  |  ${ctx.departure || '-'}`, 48, 19.5);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(...ORANGE_RGB);
      doc.text(`Dicetak: ${st.text}`, W - 10, 14, { align: 'right' });
      doc.setDrawColor(...GREEN_RGB); doc.setLineWidth(0.6); doc.line(10, 25, W - 10, 25);
      doc.setDrawColor(200); doc.setLineWidth(0.2); doc.line(10, H - 12, W - 10, H - 12);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(110);
      doc.text(`Hamdan Tour Check  |  Dicetak: ${st.text}  |  File tidak disimpan oleh sistem`, 10, H - 7.5);
      doc.text(`Halaman ${p} dari ${total}`, W - 10, H - 7.5, { align: 'right' });
    }

    return {
      buffer: doc.output('arraybuffer'),
      filename: fileName(ctx, 'Laporan', 'pdf', st),
      mime: 'application/pdf',
    };
  }

  return { buildExcel, buildPdf, actionList, decorate };
});
