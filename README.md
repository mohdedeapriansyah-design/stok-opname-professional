# STOK OPNAME PROFESSIONAL

Aplikasi web responsif untuk Reconcile Automation dan Analisa Overstok.

## Struktur Project

- `index.html` — tampilan aplikasi
- `style.css` — tampilan biru responsif
- `app.js` — pengendali aplikasi
- `js/file-engine.js` — membaca file Excel
- `js/reconcile-engine.js` — engine rekonsiliasi template Komper
- `js/export-engine.js` — download hasil XLSX
- `js/storage-engine.js` — riwayat proses
- `js/overstok-engine.js` — analisa overstok
- `manifest.json` — konfigurasi PWA
- `sw.js` — service worker
- `README.md` — panduan

## Logika Reconcile

### File SO
Engine mencari:
- `productSapCode`
- `qtyFix`

Mendukung:
- banyak file
- banyak sheet
- QtyFix desimal
- penggabungan QtyFix jika Kode SAP sama

Contoh:
`3.0` → `3`

### Template Komper
Untuk template Komper:
- `KODE ITEM` = kode pencocokan
- `HASIL SO` = kolom yang diisi QtyFix

Engine memproses semua sheet yang memiliki struktur tersebut.

Template asli tidak ditimpa. Workbook baru dibuat dari file yang dipilih pengguna, lalu hanya cell `HASIL SO` yang cocok yang diubah.

Kolom formula lain tidak disentuh oleh engine.

## Analisa Overstok

Rumus:

`(Omset Juni + Omset Juli) × Faktor`

Faktor default:

`1.5`

Jika hasil lebih besar dari Target Agustus:
`OVERSTOK`

Jika hasil sama atau lebih kecil:
`AMAN`

## Menjalankan

1. Extract ZIP.
2. Buka folder project di Visual Studio Code.
3. Install extension Live Server.
4. Klik kanan `index.html`.
5. Pilih `Open with Live Server`.

Untuk penggunaan Android/PWA, aplikasi sebaiknya dijalankan melalui HTTPS saat sudah di-hosting.

## Catatan

Aplikasi menggunakan SheetJS melalui CDN pada `index.html`. Karena itu proses pembacaan XLSX membutuhkan koneksi internet saat library belum tersedia secara lokal.

Untuk produksi offline penuh, simpan `xlsx.full.min.js` ke folder `js` lalu ubah pemanggilan CDN menjadi file lokal.
