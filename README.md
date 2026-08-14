# Reconcile Stock Pro V10

Clean rebuild untuk GitHub Pages dan Android.

## Fitur
- Template Komper sebagai dasar hasil.
- Banyak file SO.
- Deteksi header otomatis: Odoo, SAP, Kode Item, Product Name, Qty.
- Prioritas Odoo, fallback SAP, lalu Kode Item.
- Fallback SAP -> Kode Item hanya jika kecocokan suffix unik.
- Hasil ditulis ke salinan template.
- Download Hasil XLSX.
- Download SKU Tidak Ditemukan.
- Analisa Overstok: (Bulan 1 + Bulan 2) x faktor 1,5–2,0 dibanding Target Sekarang.
- Riwayat lokal di browser.
- UI responsif Android.

## GitHub Pages
Upload `index.html`, `style.css`, dan `app.js` ke root repository. Tidak memerlukan manifest.json.

Catatan: aplikasi menggunakan SheetJS dari CDN untuk membaca/menulis spreadsheet.
