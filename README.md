# Hamdan Tour Check

Cek visa jamaah Hamdan Tour: cocokkan **Manifest** dengan **Mutamer List** lewat nomor paspor, lalu unduh laporan Excel/PDF berlogo dan bercap waktu cetak.

Aplikasi statis (HTML + JS). Tanpa server, database, atau upload. File hanya dibaca di memori browser dan tidak disimpan. Koneksi keluar diblokir lewat Content-Security-Policy, semua library ada di `vendor/`.

## Deploy ke GitHub Pages

```bash
git init && git add . && git commit -m "Hamdan Tour Check"
git branch -M main
git remote add origin https://github.com/<akun>/<repo>.git
git push -u origin main
```

Lalu di repo: **Settings > Pages > Source > GitHub Actions**. Workflow `.github/workflows/pages.yml` jalan otomatis.

Coba lokal: buka `index.html`, atau `python3 -m http.server 8080`.

## Aturan

- **Passport match**: nomor paspor Manifest ada di Mutamer List.
- **Name match**: Sama / Mirip (urutan kata beda atau kemiripan >= 85%) / Beda.
- **Visa printed**: paspor cocok dan status mengandung "Printed".
- **Belum visa**: semua yang belum Visa Printed, termasuk Tidak match.
- **Tidak match**: paspor tidak ada di Mutamer List.
- **Data bermasalah**: ada minimal satu catatan sistem (tidak match, nama beda/mirip, belum printed, tanpa nomor visa, paspor ganda, paspor kedaluwarsa atau < 6 bulan, hanya ada di Mutamer List).

## Salin ke WhatsApp

Centang kolom **Pilih** (bisa banyak, atau centang kotak di judul kolom untuk semua baris yang tampil). Panel di bawah menampilkan pesan berisi KPU, nama keberangkatan, lalu per jamaah: nama, paspor, nama depan, nomor visa, status visa. Tombol **Salin ke WhatsApp** menyalin pesan. Setiap salinan selalu diakhiri tautan dan 7 langkah cek visa di `visa.mofa.gov.sa`.

## Buka laporan lama

Unggah Excel hasil unduhan ke kotak **Buka laporan lama** (atau ke kotak mana pun, dikenali otomatis). KPU, nama keberangkatan, tabel, ringkasan, dan catatan Action need terisi kembali. Catatan bisa dilanjutkan lalu diunduh ulang. Laporan **Excel bermasalah** hanya memuat baris bermasalah, sehingga yang terbuka juga hanya baris itu.

## Action need

Centang baris di tabel, isi alasan. Waktu tercatat otomatis dan ikut ke Excel (kolom + sheet "Action Need") dan PDF. Catatan hilang saat halaman ditutup.

## Struktur

```
index.html  css/style.css
js/core.js     rekonsiliasi (bisa diuji di Node)
js/export.js   Excel (ExcelJS) dan PDF (jsPDF)
js/import.js   buka kembali laporan Excel lama
js/app.js      antarmuka
js/logo.js     logo tertanam untuk laporan
assets/        logo
vendor/        SheetJS, ExcelJS, jsPDF, AutoTable
```

Uji logika: `npm i xlsx && node test/run.js manifest.xlsx mutamer.xlsx`
