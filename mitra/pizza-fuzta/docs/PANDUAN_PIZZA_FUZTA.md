# Pizza Fuzta — Bu Elok

Implementasi mengikuti MoM **27 Oktober 2026** dan lampiran `POS Pizza Fuzta.xlsm` dari pengguna. Tiga file bekerja mandiri, sesuai kebutuhan dua outlet dan satu gudang:

| File | Kegunaan |
| --- | --- |
| POS_Pizza_Fuzta_Medokan.xlsx | Kasir 49 menu/varian, struk, pembelian, stok, kas dan transfer Medokan |
| Rekap_Pizza_Fuzta_Karah.xlsx | Rekap laporan final Majoo, pembelian/stok Karah, rekonsiliasi kas dan transfer |
| Gudang_Pizza_Fuzta.xlsx | Bahan di rumah, pembelian, permintaan outlet, distribusi, konfirmasi dan batch production house aktual |

**Tidak ada sinkronisasi otomatis antar laptop, link Excel eksternal, API Majoo atau pengiriman data melalui jaringan.** Karah tetap memakai Majoo sebagai kasir utama. Gudang tidak otomatis menggabungkan omzet outlet. Dua file yang dibuka pada laptop yang sama tetap mempunyai identitas lokasi berbeda; jangan mengganti CONFIG!B6 untuk memakai satu file sebagai lokasi lain.

## Perbaikan dan audit sumber

- Katalog sumber memuat **49 menu**, tetapi `MENU_NAMA` lama masih menunjuk 20 baris. Modul baru menyediakan kapasitas 80 menu, termasuk pizza L/M/S, burger dan lima minuman.
- Teks sisa template sate, nomor kontaknya, batas menu lama, harga/HPP contoh dan tiga transaksi pada template tidak dimasukkan ke file operasional baru.
- Nama resep yang berbeda dari katalog dipetakan dengan alias eksplisit, dicatat di `data/audit_sumber.json`. Contoh: `CHEESE LOVER` ↔ `Cheese Lovers`, `VEGAN` ↔ `Veggie`, `BEEF MUSHROM` ↔ `Beef Mushroom`, dan ejaan burger ORY/ORI.
- Kuantitas teks `1,50` dibaca sebagai `1.5`, **bukan 1.500 gram**. Satuan gram/pcs asli dipertahankan; angka yang tampak tidak wajar tidak ditebak menjadi satuan lain.
- Terdapat 17 baris resep tanpa menu yang cocok, seperti Tuna Melt, Meat Lover Slice dan `Huruf Beef Cheese`. Baris tersebut dipertahankan dalam audit, bukan dijadikan produk baru tanpa harga.
- Burger Telur dan Burger Telur Cheese belum memiliki resep sumber yang cocok. Tidak dibuatkan takaran telur/bahan fiktif.
- Resep minuman satu unit dibuat sebagai draf satu botol/porsi dari katalog, tetap perlu persetujuan. Lemon Tea tidak otomatis dianggap minuman kemasan botol.
- Harga bahan dan stok sumber tidak tersedia. Semua harga/isi kemasan dan saldo awal sengaja tidak diisi. Daftar BAHAN awal adalah **kandidat dari Medokan**, bukan pernyataan semua bahan tersedia di gudang/Karah.
- Harga menu sumber tetap ditampilkan sebagai acuan. Cheese Lovers S serta Smoked Beef M/S memiliki urutan harga ukuran yang perlu dikonfirmasi. Catatan harga asumsi lama juga diberi penanda; tidak dibetulkan dengan harga tebakan.

File sumber XLSM disertakan untuk audit. Gunakan workbook baru untuk operasional; macro restoran lama tidak dipindahkan.

## Persiapan sebelum operasional

1. Buat salinan kerja setiap file dan mulai dari **MULAI**.
2. Periksa kode bahan dan satuan dasar pada ketiga file. Untuk bahan yang berpindah lokasi, kodenya harus sama dan satuannya identik. Tambahkan bahan gudang/Karah yang benar-benar digunakan; jangan otomatis menyamakan supplier/resep Karah dengan Medokan.
3. Isi harga dan isi kemasan bila sudah diketahui. Biaya per satuan adalah harga kemasan ÷ isi kemasan. HPP tetap berstatus belum lengkap jika harga belum tersedia.
4. Isi AWAL pada MUTASI berdasarkan hitung fisik, per kode bahan dan tanggal cut-off. Qty adalah satuan dasar: gram, pcs, botol atau porsi; bukan selalu jumlah kemasan.
5. Stok tidak direset tiap bulan. Bila membuat arsip/file baru, bawa saldo cut-off sekali dan hindari menghitung transaksi lama lagi. Selesaikan konfirmasi transfer pada file yang menyimpan pengiriman aslinya; file baru tidak memiliki histori kirim tersebut.
6. Konfirmasi harga menu dan takaran resep. Pada **Medokan**, isi BOM yang benar, ubah persetujuan setiap baris menjadi YA, lalu MENU → Resep Disetujui menjadi YA. Status resep harus LENGKAP.
7. Isi saldo pembukaan laci dan data kas aktual. Tidak ada saldo/pembelian/penjualan fiktif yang dimuat.

Sel biru untuk input; abu-abu berisi rumus. Jangan menimpa rumus atau mengganti kode bahan/menu yang sudah digunakan. Jumlah stok gram dan pcs tidak dijumlahkan menjadi satu KPI kuantitas.

## Tombol VBA di Excel desktop

File `.xlsx` bekerja dengan rumus dan jurnal manual. Untuk tombol:

1. Simpan salinan setiap file sebagai **Excel Macro-Enabled Workbook (.xlsm)**.
2. `Alt+F11` → pilih project workbook yang benar → File → Import File.
3. Impor **keduanya**: `Modul_PizzaFuzta.bas` dan `Modul_TransferPizza.bas`.
4. Debug → Compile VBAProject, kemudian `Alt+F8` → PasangTombol → Run. Simpan dengan Ctrl+S.
5. Aktifkan macro hanya pada file yang dipercaya. Tombol mengacu ke workbook pemiliknya, sehingga tiga file dapat dibuka tanpa sengaja menulis ke lokasi lain.

VBA disertakan sebagai sumber, **belum tertanam di `.xlsx` dan belum dikompilasi/dijalankan di Excel desktop**. Rumus dan alur pertukaran CSV diuji dengan LibreOffice/Python. Uji kompilasi, tombol dan ekspor PDF/CSV pada versi Excel mitra sebelum operasional. Ekspor CSV UTF-8 memakai format Excel 2016+ (`xlCSVUTF8`), tanpa library tambahan. Tutup CSV yang sedang terbuka sebelum mengimpor agar perubahan pengguna tidak ikut tertutup.

## Kasir Medokan

1. Isi tanggal, jam, nama kasir, metode, SKU, qty dan harga pada KASIR. Jam sumber 09.00–21.00; tidak menggunakan batas template sate yang tidak sesuai.
2. Diskon rupiah dicatat sekali per nota. Untuk Tunai, isi uang diterima; kembalian = diterima − netto. Untuk GoPay/QRIS/Transfer/OVO/DANA/Lainnya, penerimaan sama dengan netto.
3. SIMPAN NOTA menyimpan header NOTA, baris ITEM, harga dan HPP sebagai nilai tetap. Maksimum 12 baris per nota. Beberapa pcs menu sama dapat dijadikan satu baris.
4. Jika **semua** resep dalam keranjang sudah disetujui, pemakaian bahan diakumulasi dan disimpan sebagai nilai tetap di MUTASI. Stok terlacak tidak boleh negatif.
5. Jika ada satu resep yang belum lengkap, **seluruh nota belum memotong stok otomatis**. Penjualan kasir tetap tercatat dengan status stok **BELUM LENGKAP**. Ini memisahkan pencatatan uang dari persetujuan resep, bukan klaim stok sudah aktual.
6. Setelah kuantitas aktual dapat dikonfirmasi, lengkapi pemakaian seluruh nota tersebut satu kali sebagai MUTASI jenis PAKAI, tanggal sesuai nota dan No Nota terkait. Periksa tidak ada pemakaian ganda, lalu ubah status stok NOTA ke LENGKAP. Jangan merekam ulang penjualan hanya untuk memperbaiki stok.
7. Dashboard menunjukkan jumlah nota stok belum lengkap; **jangan mengandalkan saldo untuk belanja selama masih ada pemakaian yang tertunda**.
8. Pilih nomor nota pada STRUK untuk mencetak/PDF. Harga cetak berasal dari snapshot, bukan katalog terbaru.

HPP bahan yang kosong tidak diakui sebagai nol biaya yang pasti. Status HPP menjadi BELUM LENGKAP, dan dashboard tidak menyajikan laba bersih fiktif. HPP yang sudah tersimpan tidak otomatis berubah ketika harga/recipe diubah. Ini biaya acuan bahan, bukan pembukuan FIFO dan bukan biaya lengkap tenaga kerja/sewa/pajak.

Pembatalan nota: pilih baris NOTA → `Alt+F8` → VoidNota → isi alasan. Efek penjualan dan pemakaian terkait dinonaktifkan, riwayat dipertahankan. Laksanakan dan dokumentasikan pengembalian uang yang sebenarnya; macro tidak mengirim refund ke GoPay/QRIS atau bank.

## Rekap Karah dengan Majoo

Karah tidak memakai kasir Medokan. Isi **REKAP MAJOO** dari laporan harian final:

- ID laporan unik, tanggal, bruto dan diskon rupiah.
- Penerimaan **netto setelah diskon** untuk Tunai, QRIS, Transfer, GoPay, OVO, DANA dan Lainnya. Isi `0` bila tidak ada penerimaan pada suatu metode.
- Status AKTIF dan catatan referensi laporan.

Netto = bruto − diskon, dan harus sama dengan jumlah metode. Selisih ditandai SELISIH METODE dan tidak dihitung sebagai penjualan valid. Satu laporan AKTIF per tanggal; koreksi dengan VOID laporan lama, lalu catat versi final dengan ID baru. Jangan menduplikasi ekspor yang sama.

Jika laporan bank/payment gateway sudah dikurangi fee, jangan menganggap fee tersebut diskon penjualan kedua: rekonsiliasikan penerimaan transaksi bruto gateway dengan laporannya, lalu catat fee sebagai BIAYA USAHA sesuai tanggal/metode aktual. Belum ada import/API Majoo karena format ekspor belum diberikan.

ACUAN MEDOKAN hanya pembanding. Harga, resep dan pembelian roti Karah harus memakai data Karah. Pemakaian stok diisi dari pemakaian nyata/laporan yang dikonfirmasi; rekap omzet tidak otomatis menebak kuantitas bahan.

## Pembelian, keuangan dan tutup kasir

FORM menangani AWAL, BELI, KIRIM, PAKAI tanpa nota, RUSAK serta KOREKSI MASUK/KELUAR. Qty selalu positif; arah stok ditentukan jenis. Koreksi/rusak perlu alasan.

Contoh alur pembelian Santri: catat BELI dengan jumlah botol dan **nominal total pembelian** serta metode. Stok bertambah dan pembelian Tunai mengurangi laci. Jangan mencatat pembelian yang sama lagi sebagai BIAYA USAHA.

KEUANGAN memisahkan BIAYA USAHA, PRIVE, MODAL MASUK dan PENDAPATAN LAIN. Modal/prive bukan omzet penjualan. TUTUP KASIR:

**Saldo pembukaan + penjualan Tunai netto − pembelian Tunai + kas lain Tunai bertanda.**

QRIS/GoPay/Transfer/OVO/DANA tidak dimasukkan ke uang fisik laci. Saldo pembukaan tidak boleh dicatat lagi sebagai modal masuk pada tanggal yang sama. Selisih fisik tidak otomatis mengubah stok atau membuat pengeluaran fiktif.

## Gudang, permintaan dan production house

PERMINTAAN mencatat satu ID per bahan/outlet, tanggal, qty dan catatan. Ini rencana; stok tidak berubah. Pada gudang, realisasi dihitung dari KIRIM dengan referensi permintaan dan bahan/tujuan yang sama. Pada outlet, realisasi mengikuti TERIMA. Qty belum terpenuhi tetap terlihat.

Untuk production house, isi tanggal serta bahan/qty aktual pada sisi input dan kode hasil/qty aktual pada sisi output sheet PRODUKSI. Minimal satu input dan satu output. Tidak ada rumus resep gudang yang diasumsikan; gram dan pcs tidak dibandingkan seolah unitnya sama. Tombol mencatat kedua sisi dengan ID dokumen yang sama dan menolak stok negatif. Harga/HPP hasil diisi dari biaya yang dikonfirmasi.

Pembatalan manual batch harus mengubah **seluruh baris ID produksi yang sama** menjadi VOID; jangan membatalkan satu sisi. Pasangan produksi yang tidak lengkap ditandai dan tidak memengaruhi stok valid.

## Transfer offline antar file

1. Pengirim mencatat KIRIM: asal lokasi file ini, tujuan berbeda, kode bahan, qty, tanggal dan ID transfer unik. Bila E12 kosong, macro membuat ID. No Permintaan opsional.
2. Pada TRANSFER klik EKSPOR KIRIM CSV dan pilih kode lokasi lawan. Satu CSV hanya untuk satu lokasi. Serahkan file lewat cara yang disepakati; aplikasi tidak mengirim pesan/file ke pihak lain.
3. Penerima membuka file lokasi yang benar, IMPOR KIRIM CSV, lalu isi tanggal terima aktual (ISO `yyyy-mm-dd`). Tujuan, satuan, kode bahan, tanggal dan ID diperiksa.
4. Import ulang dengan metadata kirim yang sama dilewati dan **tidak mengubah qty/tanggal terima yang sudah dicatat**. ID sama dengan metadata kirim berubah ditolak. Batch salah tidak meninggalkan sebagian baris baru.
5. Saat diterima, periksa fisik. Jika qty berbeda, ubah Qty TERIMA sesuai aktual sebelum konfirmasi. Tgl Kirim dan Qty Kirim sumber jangan diubah. Uraikan selisih pada catatan; jangan mengonfirmasi jumlah yang tidak diterima.
6. Penerima mengekspor TERIMA CSV ke lokasi pengirim. Pengirim melakukan IMPOR KONFIRMASI.
7. KONFIRMASI adalah bukti rekonsiliasi saja: **tidak menambah stok pengirim**, tidak memotong stok penerima kedua kali. Jika Qty Terima ≠ Qty Kirim, kolom Cek menunjukkan SELISIH dan delta tetap terlihat.
8. Retur adalah transfer baru dengan ID baru dan asal/tujuan terbalik. Jangan menghapus/VOID sepihak transfer yang sudah diterima file lain.

Kirim 20 botol, terima aktual 19 → stok pengirim berkurang 20, penerima bertambah 19, konfirmasi menunjukkan selisih −1. Tidak ada penyesuaian otomatis yang menyembunyikan perbedaan. Penanganan selisih harus mengikuti pemeriksaan aktual.

CSV memakai schema PF1, UTF-8, tanggal ISO dan desimal titik tanpa pemisah ribuan. Kode/ID memakai huruf, angka, garis bawah atau minus. Nomor transfer unik per baris bahan; beberapa bahan dalam surat jalan yang sama tetap mempunyai ID baris berbeda. Kode/satuan harus disamakan di ketiga master sebelum pertukaran. Tidak ada konversi satuan otomatis.

## CLI pengembang (opsional; `.xlsx` saja)

Pengguna Excel dapat memakai tombol VBA. CLI tersedia untuk pengujian/pengelolaan `.xlsx` dan sengaja menolak `.xlsm` agar kontrol tombol tidak terhapus oleh library.

```sh
cd /workspace/Social-Project-BSI/mitra/pizza-fuzta
python3 src/transfer_cli.py export-shipments output/Gudang_Pizza_Fuzta.xlsx --csv kirim-medokan.csv --peer MEDOKAN
python3 src/transfer_cli.py import-shipments output/POS_Pizza_Fuzta_Medokan.xlsx --csv kirim-medokan.csv --received-on 2026-10-28 --output Medokan_setelah_terima.xlsx
python3 src/transfer_cli.py export-receipts Medokan_setelah_terima.xlsx --csv terima-medokan.csv --peer GUDANG
python3 src/transfer_cli.py import-receipts output/Gudang_Pizza_Fuzta.xlsx --csv terima-medokan.csv --output Gudang_setelah_konfirmasi.xlsx
```

Tanggal di atas contoh argumen; gunakan tanggal aktual. CSV hanya dapat diekspor jika jurnal sudah berisi KIRIM/TERIMA yang valid; workbook kosong tidak menghasilkan pengiriman contoh. File keluaran wajib baru, berbeda dari sumber. CLI tidak menimpa file asli.

Untuk jumlah terima berbeda, `import-shipments` dapat diberi `--received jumlah-terima.json`, berisi misalnya `{"ID-TRANSFER-AKTUAL": 19}`. Gunakan ID yang benar-benar ada dalam CSV, bukan teks contoh itu. Penyesuaian qty aktual terhadap ID yang sudah diimpor tidak dilakukan diam-diam; konflik ditolak. Semua hasil perlu dibuka di Excel/LibreOffice untuk menghitung ulang cache rumus, lalu periksa kolom Cek.

## Kapasitas, validasi dan data lanjutan

Kapasitas umum: 80 bahan, 4.000 gerak stok, 1.000 keuangan, 1.000 permintaan dan 2.000 konfirmasi. Medokan: 80 menu, 600 BOM, 1.000 nota, 3.000 item dan 12 baris per nota. Karah: 1.000 rekap. Satu operator per file; tidak ada audit trail perubahan sel, penguncian periode atau database bersama.

Data lanjutan: daftar bahan yang benar-benar disimpan di rumah, stok/biaya awal, persetujuan BOM/satuan, harga menu yang perlu konfirmasi, rincian pembelian/resep Karah, format ekspor Majoo bila ingin integrasi berikutnya, serta resep dan hasil production house. Pengujian memakai simulasi di direktori sementara; tidak masuk ke file operasional. Mengubah master tidak mengubah harga/kuantitas snapshot transaksi lama, tetapi jangan mengganti kode/satuan bahan yang sudah dipakai.
