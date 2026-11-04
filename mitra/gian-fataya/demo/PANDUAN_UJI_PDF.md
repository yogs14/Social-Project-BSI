# Latihan tombol PDF Gian Fataya

Paket ini berisi file **SIMULASI**, satu modul VBA dan panduan. Semua harga, transaksi, outlet/PIC dan saldo latihan tidak boleh dipakai sebagai data/tagihan nyata. Workbook operasional tidak diubah.

1. Ekstrak `Paket_SIMULASI_Uji_PDF.zip` ke folder lokal yang mudah ditemukan.
2. Buka `SIMULASI_Uji_PDF_Gian_Fataya.xlsx` dengan **Microsoft Excel desktop**.
3. File → Save As → pilih **Excel Macro-Enabled Workbook (*.xlsm)**, beri nama `SIMULASI_Gian_Fataya.xlsm`.
4. Tekan **Alt+F11**. Pada panel kiri, pilih `VBAProject (SIMULASI_Gian_Fataya.xlsm)`.
5. File → Import File → pilih `Modul_GianFataya.bas` dari folder yang diekstrak.
6. Pada editor VBA: Debug → Compile VBAProject. Jika ada kesalahan kompilasi, catat pesannya; pengujian VBA desktop memang belum dijalankan di lingkungan Linux.
7. Kembali ke Excel dengan **Alt+F11**. Tekan **Alt+F8**, pilih `PasangTombol` milik workbook latihan, lalu Run. Simpan Ctrl+S. Bila diminta izin menjalankan macro, gunakan hanya pada file latihan yang Anda percayai ini, tanpa mengaktifkan semua macro secara global.
8. Buka sheet **CETAK INVOICE**. B5 sudah berisi `GF-DEMO-001`, indikator A14 harus **SIAP CETAK**.
9. Klik **EKSPOR INVOICE PDF**, di dekat sel D5. Pilih folder tujuan dan tekan **Save**.
10. Excel menampilkan pesan **PDF tersimpan**. Buka `GF-DEMO-001.pdf` di folder pilihan Anda.

PDF harus memuat tanda **SIMULASI / BUKAN TAGIHAN NYATA** dengan hasil berikut:

| Komponen | Nilai latihan |
| --- | ---: |
| Jumlah | 30 pouch |
| Harga / pouch | Rp10.000 |
| Bruto | Rp300.000 |
| Diskon 10% | Rp30.000 |
| Ongkir | Rp20.000 |
| Total tagihan | **Rp290.000** |
| Pembayaran | Rp0 |
| Sisa | Rp290.000 |

Tautan navigasi **Cetak Invoice** hanya membuka lembar. Tombol **EKSPOR INVOICE PDF** yang dipasang oleh VBA membuat file setelah lokasi dipilih; PDF tidak dibuka otomatis oleh macro.

Jika tombol belum muncul, jalankan `PasangTombol` lagi pada workbook latihan yang benar. Jika `PasangTombol` tidak ada pada daftar Alt+F8, periksa modul sudah diimpor ke project file `.xlsm`, bukan workbook lain. Simpan semua hasil latihan dengan nama SIMULASI.

Data dan rumus invoice latihan sudah dihitung ulang dengan LibreOffice. Ini memverifikasi status SIAP CETAK serta total Rp290.000, **bukan verifikasi eksekusi tombol VBA di Excel desktop**. Untuk membuat ulang demo dari checkout: `python3 mitra/gian-fataya/src/build_invoice_demo.py` (memerlukan openpyxl dan soffice).
