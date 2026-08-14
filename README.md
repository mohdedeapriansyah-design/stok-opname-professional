# Reconcile Stock Pro V10.1

Versi perbaikan dari V10 untuk GitHub Pages / Android.

## Perbaikan utama
- Semua sheet template brand dibaca, bukan hanya sheet pertama.
- Header Odoo / SAP / Kode Item / Product Name / Qty Fix dideteksi otomatis.
- `qtyFix` dari file SO dipetakan ke kolom **HASIL SO** pada template.
- Jika beberapa file SO memiliki SKU yang sama, Qty dijumlahkan, bukan ditimpa.
- Odoo diprioritaskan, SAP menjadi fallback, lalu Kode Item.
- Kode numerik dinormalisasi sehingga `05880` dan `5880` dapat dicocokkan.
- Template asli tidak diubah secara langsung; hasil dibuat sebagai file baru.
- SKU yang tidak ditemukan dapat diekspor terpisah.
- Menu Tentang/penjelasan panjang dihapus agar UI lebih profesional.
- Analisa Overstok: `(Bulan 1 + Bulan 2) × faktor 1,5–2,0`, lalu dibandingkan dengan Target Sekarang.


## V10.2
- Explicit `HASIL SO` destination; `qtyFix` is source only.
- Uses actual worksheet cell ranges to avoid scanning declared million-row ranges.
