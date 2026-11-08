# Social Project BSI — template operasional UMKM

Workbook operasional setiap mitra mengikuti MoM dan data pendukung dashboard Bengkel UMKM.

| Mitra | Alur yang disiapkan | Workbook | Panduan |
| --- | --- | --- | --- |
| Gian Fataya — Bu Santi | Enam sheet: stok konsinyasi, distribusi PO, invoice wilayah dan cicilan | [Excel](mitra/gian-fataya/output/POS_Gian_Fataya.xlsx) · [Paket ZIP](mitra/gian-fataya/output/Paket_POS_Gian_Fataya.zip) | [Panduan](mitra/gian-fataya/docs/PANDUAN_GIAN_FATAYA.md) |
| Murita Crepes — Pak Syarif | POS 34 menu, premiks tujuh bahan, adonan basah, kalibrasi topping, kas usaha/pribadi dan CRM/poin | [Excel](mitra/murita-crepes/output/POS_Murita_Crepes.xlsx) · [Paket ZIP](mitra/murita-crepes/output/Paket_POS_Murita_Crepes.zip) | [Panduan](mitra/murita-crepes/docs/PANDUAN_MURITA_CREPES.md) |
| Pizza Fuzta — Bu Elok | Kasir 49 menu Medokan, rekap Majoo Karah, gudang dan transfer CSV offline dengan konfirmasi | [Medokan](mitra/pizza-fuzta/output/POS_Pizza_Fuzta_Medokan.xlsx) · [Karah](mitra/pizza-fuzta/output/Rekap_Pizza_Fuzta_Karah.xlsx) · [Gudang](mitra/pizza-fuzta/output/Gudang_Pizza_Fuzta.xlsx) · [Paket ZIP](mitra/pizza-fuzta/output/Paket_POS_Pizza_Fuzta.zip) | [Panduan](mitra/pizza-fuzta/docs/PANDUAN_PIZZA_FUZTA.md) |
| Chicken Al-Haqi — Mba Gita | Kasir potongan + rasa, nasi, belanja kg/pcs aktual, stok dan pemisahan kas ayam/pribadi/usaha lain | [Excel](mitra/chicken-al-haqi/output/POS_Chicken_Al_Haqi.xlsx) · [Paket ZIP](mitra/chicken-al-haqi/output/Paket_POS_Chicken_Al_Haqi.zip) | [Panduan](mitra/chicken-al-haqi/docs/PANDUAN_CHICKEN_AL_HAQI.md) |

File utama tidak memuat transaksi/pelanggan/saldo fiktif. Data yang belum tersedia ditandai untuk diisi mitra. Pada Murita, harga SKM dari rumus sumber tidak dipakai karena salah referensi; HPP tetap ditandai belum lengkap sampai harga sebenarnya diisi. Loyalty awalnya nonaktif karena kebijakan hadiah belum diberikan.

Gian Fataya v2 mengikuti kolom **Updated** di MoM: 14 sheet digabung menjadi **6 sheet nyata** (BERANDA, TRANSAKSI, STOK, TAGIHAN, CETAK INVOICE, DATA USAHA). PO/ongkir kini berdekatan dengan qty; invoice dan pembayaran berada pada satu sheet dengan tautan bagian. Mitra dibedakan Amanda/Independen; PO minimal 10 pcs dapat mencakup beberapa lot. Beranda menampilkan grafik batang penjualan delapan wilayah dan grafik area tren 12 bulan dengan pilihan wilayah. Kategori tinggi/sedang/datar memiliki warna dan label teks; ambang relatif dapat diubah. Contoh terpisah berlabel SIMULASI disertakan dalam paket. [Riset UI/UX dan pemetaan kebutuhan](mitra/gian-fataya/docs/RISET_UI_UX_GIAN_FATAYA.md) menyertakan sumber NN/G dan Microsoft. Tombol v2 memerlukan modul VBA v2; jangan memakai modul lama pada tata letak baru.

Untuk meninjau grafik Gian Fataya, buka [pratinjau dashboard](mitra/gian-fataya/demo/Pratinjau_Grafik_Wilayah.png) atau [Excel simulasi delapan wilayah](mitra/gian-fataya/demo/SIMULASI_Grafik_Wilayah.xlsx). Semua nilai pada contoh adalah data fiktif. Pilih wilayah di BERANDA!I18 untuk mengganti grafik area.

Untuk mencoba seluruh alur Gian Fataya, tersedia [Excel berisi data DUMMY](mitra/gian-fataya/latihan/POS_Gian_Fataya_DUMMY.xlsx), [panduan praktik PDF](mitra/gian-fataya/latihan/PANDUAN_LATIHAN_GIAN_FATAYA.pdf), [versi Markdown](mitra/gian-fataya/latihan/PANDUAN_LATIHAN_GIAN_FATAYA.md), dan [paket latihan lengkap](mitra/gian-fataya/latihan/Paket_Latihan_Gian_Fataya_DUMMY.zip). Contoh mencakup empat outlet, stok dua bulan, empat invoice dan empat pembayaran. Panduan menunjukkan sel input, urutan kerja dan angka hasil; paket menyertakan template kosong untuk latihan ulang serta contoh invoice PDF. Gunakan file DUMMY untuk melihat hasil akhir; ketik ulang latihan pada salinan template kosong.

Navigasi Gian Fataya sudah diperbaiki untuk error Excel **Cannot open the specified file**: tombol menuju sheet dalam workbook yang sama dan tetap dapat digunakan setelah file dipindah atau diganti namanya. Perbaikan mencakup file operasional, contoh grafik, file DUMMY, serta template kosong dalam paket latihan.

Versi **Mufrita/Murita Crepes ringkas** menyatukan 20 sheet menjadi 8 tab kerja dan 1 sheet rumus tersembunyi: [Excel ringkas](mitra/murita-crepes/output/ringkas/POS_Mufrita_Crepes_Ringkas.xlsx) · [Paket dengan VBA dan panduan](mitra/murita-crepes/output/ringkas/Paket_Mufrita_Crepes_Ringkas.zip) · [Panduan](mitra/murita-crepes/docs/PANDUAN_MUFRITA_RINGKAS.md). Jurnal tetap menyimpan tabel nota, item, gerak bahan, keuangan dan poin terpisah dengan pintasan. Gunakan modul VBA versi ringkas yang disertakan, karena alamat blok telah berubah. Model lama dipertahankan. STOK mengikuti layar Stok Realtime POS Pak Seger; TUTUP KASIR menjadi tab tersendiri dengan rekonsiliasi tunai dan arsip revisi. Master bahan berada di KATALOG A225.

## Menghasilkan dan memvalidasi workbook

Gunakan checkout yang ada; cloud task sudah terisolasi, tidak perlu worktree baru.

```sh
cd /workspace/Social-Project-BSI/mitra/gian-fataya
python3 -m pip install -r requirements.txt
python3 src/build_workbook.py
python3 tests/validate_workbook.py
python3 src/build_chart_demo.py
python3 src/package_delivery.py
python3 src/build_full_demo.py
python3 tests/validate_navigation.py

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

cd /workspace/Social-Project-BSI/mitra/chicken-al-haqi
python3 -m pip install -r requirements.txt
python3 src/build_workbook.py
python3 tests/validate_workbook.py
python3 tests/validate_ux.py
python3 src/package_delivery.py
```

Untuk menghasilkan versi Crepes ringkas: dari `mitra/murita-crepes`, jalankan `python3 src/build_compact.py`, `python3 tests/validate_compact.py`, lalu `python3 src/package_compact.py`. Pengujian menjalankan kembali skenario bisnis model awal pada alamat fisik workbook ringkas dan menghitung ulang rumusnya.

Python 3.11+; generator memakai XlsxWriter, pengujian memakai openpyxl dan `soffice` (LibreOffice headless). Tidak ada server atau database untuk dijalankan. Pengujian memakai file simulasi terpisah di `/tmp`, menghitung ulang rumus yang benar-benar digunakan workbook, lalu menyimpan hasil validasi JSON. File operasional tetap tanpa transaksi contoh.

Workbook utama `.xlsx` bekerja dengan rumus dan jurnal manual. Sumber VBA tersedia pada direktori `vba/` masing-masing mitra; impor ke salinan `.xlsm` di Excel desktop untuk tombol pencatatan dan ekspor PDF. VBA belum dikompilasi atau dieksekusi di Linux; periksa Debug → Compile VBAProject dan alur tombol sebelum operasional. Modul restoran lama tidak digunakan untuk alur konsinyasi atau produksi bertahap.

Data referensi yang berhasil dibaca tersimpan bersama metadata sumber di `data/` masing-masing mitra. Hasil uji ada pada `output/hasil_validasi.json`. Pengujian menghitung rumus workbook sebenarnya: stok, perbedaan penjualan/kas, cutoff tanggal, snapshot, pembatalan dan kebutuhan khusus tiap mitra. File simulasi tidak masuk ke workbook operasional.

Untuk Murita, alurnya **bahan mentah → premiks → adonan basah → penjualan**: tujuh bahan mentah tidak dipotong dua kali saat kasir mencatat penjualan. Harga, kuantitas bahan, HPP dan poin disimpan sebagai snapshot oleh VBA; perubahan harga/kalibrasi tidak mengubah jurnal lama. Kontak pelanggan opsional, dan tidak ada pengiriman pesan otomatis.

Pizza Fuzta memakai **tiga file mandiri** sesuai MoM 27 Oktober 2026. Karah tetap memakai Majoo sebagai kasir utama; workbook Karah merekap laporan final dan tidak menebak resep/harga Karah dari Medokan. Gudang mengelola permintaan, pembelian, pengiriman serta batch produksi aktual. Tidak ada sinkronisasi otomatis, tautan Excel eksternal, API Majoo atau pengiriman file lewat jaringan.

Pertukaran CSV Pizza Fuzta memeriksa schema PF1, lokasi, kode bahan, satuan, tanggal, qty dan ID unik. Impor ulang identik tidak menggandakan stok; konfirmasi penerimaan tidak menjadi mutasi stok kedua. Selisih kirim/terima tetap terlihat. CLI opsional hanya menerima `.xlsx` dan menghasilkan file baru; untuk `.xlsm` bertombol gunakan kedua modul VBA yang disertakan.

Seluruh BOM Pizza Fuzta awalnya memerlukan persetujuan. Penjualan dengan resep belum lengkap tetap dicatat, tetapi belum memotong stok otomatis dan ditandai **BELUM LENGKAP**; lengkapi pemakaian satu kali sebelum memakai saldo untuk belanja. Harga/stok bahan yang belum diberikan tidak dibuat-buat. Alias resep, kuantitas desimal koma dan baris resep tanpa menu diaudit pada `mitra/pizza-fuzta/data/audit_sumber.json`.

Chicken Al-Haqi mengikuti permintaan penyederhanaan pada MoM 23 Oktober 2026: sembilan tab utama. Kasir mengikuti referensi POS sate dengan dropdown 17 kombinasi potongan/rasa/nasi, 12 baris keranjang, panel omzet/stok dan tab TUTUP KASIR; buku NOTA terlihat untuk audit. Perkiraan 9/10 pcs per kg hanya pembanding; stok bertambah dari **pcs aktual**, bukan kg dan perkiraan sekaligus. Saldo awal yang belum diketahui tetap kosong/ditandai BELUM AWAL. Kas ayam dipisahkan dari prive dan arus ke/dari usaha lain seperti obat, tanpa menambahkan penjualan obat ke kasir ayam.

HPP umum Chicken Al-Haqi Rp7.000/pcs adalah acuan sumber yang belum dikonfirmasi, bukan biaya lengkap setiap potongan. HPP saus/nasi tidak dikarang, dan biaya sumber kg tidak ditambahkan lagi ke biaya per pcs. Harga/HPP transaksi disimpan sebagai snapshot. Jam weekend 06.30 diterima; belanja di atas 3 kg tetap diperbolehkan karena batas tersebut untuk ukuran batch marinasi, bukan jumlah pembelian.

Adaptasi Chicken Al-Haqi terhadap referensi POS sate mempertahankan jurnal pcs aktual dan snapshot harga/HPP. Tanggal, modal awal dan uang fisik kini diisi pada TUTUP KASIR C5/I8/I11; dashboard HARIAN mengikutinya. Petunjuk migrasi file yang sudah berisi transaksi tersedia pada panduan mitra.

UX Chicken Al-Haqi memakai gaya dan ukuran sel asli POS Sate Ayam Ponorogo: merah bata/oranye, input putih, panel hijau, kontrol transaksi terlihat, dan KASIR terbuka langsung. Tampilan dihitung ulang pada skenario kasir/tutup kasir; XML tombol Excel dipertahankan ketika cache rumus diperbarui.
