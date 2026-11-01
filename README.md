# Social Project BSI — template operasional UMKM

Workbook operasional setiap mitra mengikuti MoM dan data pendukung dashboard Bengkel UMKM.

| Mitra | Alur yang disiapkan | Workbook | Panduan |
| --- | --- | --- | --- |
| Gian Fataya — Bu Santi | Stok konsinyasi, distribusi outlet, invoice wilayah dan pembayaran bertahap | [Excel](mitra/gian-fataya/output/POS_Gian_Fataya.xlsx) · [Paket ZIP](mitra/gian-fataya/output/Paket_POS_Gian_Fataya.zip) | [Panduan](mitra/gian-fataya/docs/PANDUAN_GIAN_FATAYA.md) |
| Murita Crepes — Pak Syarif | POS 34 menu, premiks tujuh bahan, adonan basah, kalibrasi topping, kas usaha/pribadi dan CRM/poin | [Excel](mitra/murita-crepes/output/POS_Murita_Crepes.xlsx) · [Paket ZIP](mitra/murita-crepes/output/Paket_POS_Murita_Crepes.zip) | [Panduan](mitra/murita-crepes/docs/PANDUAN_MURITA_CREPES.md) |
| Pizza Fuzta — Bu Elok | Kasir 49 menu Medokan, rekap Majoo Karah, gudang dan transfer CSV offline dengan konfirmasi | [Medokan](mitra/pizza-fuzta/output/POS_Pizza_Fuzta_Medokan.xlsx) · [Karah](mitra/pizza-fuzta/output/Rekap_Pizza_Fuzta_Karah.xlsx) · [Gudang](mitra/pizza-fuzta/output/Gudang_Pizza_Fuzta.xlsx) · [Paket ZIP](mitra/pizza-fuzta/output/Paket_POS_Pizza_Fuzta.zip) | [Panduan](mitra/pizza-fuzta/docs/PANDUAN_PIZZA_FUZTA.md) |

File utama tidak memuat transaksi/pelanggan/saldo fiktif. Data yang belum tersedia ditandai untuk diisi mitra. Pada Murita, harga SKM dari rumus sumber tidak dipakai karena salah referensi; HPP tetap ditandai belum lengkap sampai harga sebenarnya diisi. Loyalty awalnya nonaktif karena kebijakan hadiah belum diberikan.

## Menghasilkan dan memvalidasi workbook

Gunakan checkout yang ada; cloud task sudah terisolasi, tidak perlu worktree baru.

```sh
cd /workspace/Social-Project-BSI/mitra/gian-fataya
python3 -m pip install -r requirements.txt
python3 src/build_workbook.py
python3 tests/validate_workbook.py
python3 src/package_delivery.py

cd /workspace/Social-Project-BSI/mitra/pizza-fuzta
python3 -m pip install -r requirements.txt
python3 src/build_workbooks.py
python3 tests/validate_workbooks.py
python3 src/package_delivery.py

cd /workspace/Social-Project-BSI/mitra/murita-crepes
python3 -m pip install -r requirements.txt
python3 src/build_workbook.py
python3 tests/validate_workbook.py
python3 src/package_delivery.py
```

Python 3.11+; generator memakai XlsxWriter, pengujian memakai openpyxl dan `soffice` (LibreOffice headless). Tidak ada server atau database untuk dijalankan. Pengujian memakai file simulasi terpisah di `/tmp`, menghitung ulang rumus yang benar-benar digunakan workbook, lalu menyimpan hasil validasi JSON. File operasional tetap tanpa transaksi contoh.

Workbook utama `.xlsx` bekerja dengan rumus dan jurnal manual. Sumber VBA tersedia pada direktori `vba/` masing-masing mitra; impor ke salinan `.xlsm` di Excel desktop untuk tombol pencatatan dan ekspor PDF. VBA belum dikompilasi atau dieksekusi di Linux; periksa Debug → Compile VBAProject dan alur tombol sebelum operasional. Modul restoran lama tidak digunakan untuk alur konsinyasi atau produksi bertahap.

Data referensi yang berhasil dibaca tersimpan bersama metadata sumber di `data/` masing-masing mitra. Hasil uji ada pada `output/hasil_validasi.json`. Pengujian menghitung rumus workbook sebenarnya: stok, perbedaan penjualan/kas, cutoff tanggal, snapshot, pembatalan dan kebutuhan khusus tiap mitra. File simulasi tidak masuk ke workbook operasional.

Untuk Murita, alurnya **bahan mentah → premiks → adonan basah → penjualan**: tujuh bahan mentah tidak dipotong dua kali saat kasir mencatat penjualan. Harga, kuantitas bahan, HPP dan poin disimpan sebagai snapshot oleh VBA; perubahan harga/kalibrasi tidak mengubah jurnal lama. Kontak pelanggan opsional, dan tidak ada pengiriman pesan otomatis.

Pizza Fuzta memakai **tiga file mandiri** sesuai MoM 27 Oktober 2026. Karah tetap memakai Majoo sebagai kasir utama; workbook Karah merekap laporan final dan tidak menebak resep/harga Karah dari Medokan. Gudang mengelola permintaan, pembelian, pengiriman serta batch produksi aktual. Tidak ada sinkronisasi otomatis, tautan Excel eksternal, API Majoo atau pengiriman file lewat jaringan.

Pertukaran CSV Pizza Fuzta memeriksa schema PF1, lokasi, kode bahan, satuan, tanggal, qty dan ID unik. Impor ulang identik tidak menggandakan stok; konfirmasi penerimaan tidak menjadi mutasi stok kedua. Selisih kirim/terima tetap terlihat. CLI opsional hanya menerima `.xlsx` dan menghasilkan file baru; untuk `.xlsm` bertombol gunakan kedua modul VBA yang disertakan.

Seluruh BOM Pizza Fuzta awalnya memerlukan persetujuan. Penjualan dengan resep belum lengkap tetap dicatat, tetapi belum memotong stok otomatis dan ditandai **BELUM LENGKAP**; lengkapi pemakaian satu kali sebelum memakai saldo untuk belanja. Harga/stok bahan yang belum diberikan tidak dibuat-buat. Alias resep, kuantitas desimal koma dan baris resep tanpa menu diaudit pada `mitra/pizza-fuzta/data/audit_sumber.json`.
