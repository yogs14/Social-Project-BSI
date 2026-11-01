# Social Project BSI — template operasional UMKM

Implementasi pertama: **Gian Fataya**, stok konsinyasi dan invoice per wilayah berdasarkan MoM 22 Oktober 2026.

Lihat [panduan](mitra/gian-fataya/docs/PANDUAN_GIAN_FATAYA.md). Workbook dan paket distribusi ada di `mitra/gian-fataya/output/`.

## Menghasilkan dan memvalidasi workbook

Gunakan checkout yang ada; cloud task sudah terisolasi, tidak perlu worktree baru.

```sh
cd /workspace/Social-Project-BSI/mitra/gian-fataya
python3 -m pip install -r requirements.txt
python3 src/build_workbook.py
python3 tests/validate_workbook.py
```

Python 3.11+; generator memakai XlsxWriter, pengujian memakai openpyxl dan `soffice` (LibreOffice headless). Tidak ada server atau database untuk dijalankan. Pengujian memakai file simulasi terpisah di `/tmp`, menghitung ulang rumus yang benar-benar digunakan workbook, lalu menyimpan hasil validasi JSON. File operasional tetap tanpa transaksi contoh.

Workbook utama `.xlsx` bekerja dengan jurnal manual. Sumber VBA tambahan tersedia di `vba/Modul_GianFataya.bas`; impor ke salinan `.xlsm` di Excel desktop untuk tombol FORM/PDF. VBA tidak dikompilasi atau dieksekusi di Linux. Modul restoran pada lampiran awal tidak dibawa ke file konsinyasi.

Data referensi yang berhasil dibaca tersimpan lokal bersama metadata sumber di `data/`. Harga jual, outlet, PIC tagihan, saldo awal, tanggal expired dan rekening belum tersedia dan harus diisi dari data mitra yang telah dikonfirmasi.
