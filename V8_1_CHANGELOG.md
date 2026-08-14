# V8.1 — Android memory stabilization

- Membersihkan `sourceMap` setelah proses rekonsiliasi selesai.
- Melepaskan referensi workbook dan worksheet setiap file SO segera setelah diproses.
- Memberi jeda 16 ms antar file agar UI Android tetap responsif.
- Menampilkan jumlah SKU yang sudah digabung setelah parsing SO.
- Tetap tanpa batas jumlah baris.
- Rumus overstok tetap: (A + B) × faktor 1,5–2,0 dibandingkan Target Sekarang.
