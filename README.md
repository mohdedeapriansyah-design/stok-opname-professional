# Reconcile Stock Pro V9
Clean rebuild untuk GitHub Pages / Android.

## Fitur
- Beranda, Reconcile Otomatis, Analisa Overstok, Riwayat, Tentang.
- Upload Template Komper.
- Upload banyak file SO.
- Pembacaan header `productSapCode/Kode SAP`, `ProductName`, `ProductCode`, `qtyFix`.
- Normalisasi kode SAP numerik dengan menghapus leading zero.
- Hasil SO ditulis ke salinan template, bukan file template asli.
- Overstok manual per brand: `(Bulan 1 + Bulan 2) × Faktor`, faktor 1,5–2,0, dibandingkan Target Sekarang.

## Catatan
V9 adalah fondasi baru. Jalur produksi masih dapat dioptimalkan lagi untuk template mendekati batas 1.048.576 baris; tahap berikutnya sebaiknya mengganti scan template penuh dengan indeks/streaming yang benar-benar terarah.
