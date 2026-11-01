# Murita Crepes — Pak Syarif

Workbook: `POS_Murita_Crepes.xlsx`. Modul tombol: `Modul_MuritaCrepes.bas`.

File utama tidak berisi penjualan, stok awal, pelanggan, pengeluaran, atau program hadiah fiktif. Tiga puluh empat menu dan harga bersumber dari file biaya mitra; periksa kembali harga terkini sebelum dipakai.

## Apa yang disesuaikan dari template POS

- Kasir mendukung 34 menu sumber dengan kapasitas 50 menu, 12 baris per nota, diskon rupiah, pembayaran Tunai/QRIS/Transfer/E-wallet, kembalian dan struk.
- Jam 23.00 diterima; tidak memakai batas pukul 22.00 pada modul restoran sebelumnya. MoM menyebut buka 17.00–23.00 dan Senin libur rutin. Transaksi pada hari libur tidak dipaksakan menjadi nol karena mungkin ada penjualan khusus.
- Tujuh bahan mentah direkombinasikan menjadi **PREMIX**, lalu PREMIX + mentega + air menjadi **WET** (adonan basah). Penjualan hanya memakai WET dan pelengkap/topping, sehingga tidak memotong bahan mentah dua kali.
- Takaran topping dapat memakai acuan sumber atau hasil kalibrasi sendok. Konsumsi tetap estimasi; lakukan hitung fisik berkala.
- KEUANGAN memisahkan biaya usaha, modal, prive dan pendapatan lain. Pembelian stok tidak dihitung lagi sebagai biaya kedua setelah HPP acuan.
- CRM menghitung kunjungan per nota dan belanja netto; program poin opsional, awalnya tidak aktif.

## Data sumber dan koreksi rumus

MoM: 24 Oktober 2026. Folder referensi tercantum pada REFERENSI dan metadata URL/hash pada `data/sumber.json`.

Komposisi resep premiks sumber adalah 1.000 g tepung beras, 500 g terigu, 500 g maizena, 400 g fullcream, 500 g gula, 6,5 g garam dan 2 g vanili. Totalnya **2.908,5 gram**, dengan biaya acuan **Rp56.731,25** sesuai harga pada file sumber.

File sumber memakai angka 2.900 g dalam sebagian rumus mentega/air, dan harga SKM salah merujuk ke total biaya adonan. Workbook baru:

- Menormalisasi tujuh bahan dengan total berat sebenarnya.
- Menggunakan rasio mentega `0,065` dan air `0,83` terhadap premiks yang benar-benar dipakai.
- Menggunakan asumsi kerja MoM **30 crepes per kg premiks**, dapat diubah ke hasil ukur aktual (MoM menyebut kisaran 28–30).
- Mengosongkan harga beli SKM. Isi harga sebenarnya; tidak mengarang nominal pengganti.
- Tidak menyalin total biaya adonan lama yang tidak konsisten dengan penjumlahan komponen.

Dengan 1.000 g premiks: mentega 65 g, air 830 g, adonan basah 1.895 g, sekitar 30 porsi. Takaran WET per crepes adalah 63,1667 g. Semua ini kalkulator rencana, bukan transaksi atau saldo awal.

HPP acuan juga memuat SKM 5,63 g, gas 4 menit dan kemasan 0,25 lembar per crepes sesuai sumber. Takaran tersebut perlu dikonfirmasi, termasuk satuan pisang **buah** dan marshmallow **pcs**. Harga sumber adalah acuan, bukan biaya aktual FIFO. Biaya tetap pada file presentasi tidak dimuat sebagai pengeluaran nyata; termasuk biaya pegawai yang belum sesuai dengan MoM usaha dikelola sendiri.

## Memulai

1. Buka **MULAI**, lalu periksa MENU, BAHAN, PREMIKS, ADONAN dan REFERENSI.
2. Isi harga beli SKM di **BAHAN**, kode `SKM`, kolom Harga Kemasan. Isi bersih sumber adalah 490 gram; periksa kemasan aktual.
3. Konfirmasikan harga menu, harga/isi kemasan bahan, serta takaran pelengkap. Jangan mengganti kode bahan/menu yang sudah digunakan pada jurnal.
4. Catat stok awal berdasarkan hitung fisik pada GERAK BAHAN, jenis `AWAL`, satu baris per bahan. Masukkan dalam **satuan dasar**, misalnya 500 gram, bukan 1 bungkus.
5. Bahan PREMIX/WET boleh diberi AWAL hanya bila memang sudah ada saat cut-off awal. Jangan membuat AWAL premiks yang sama sekaligus memproduksinya lagi dari bahan yang sudah tercatat.
6. Air dan gas awalnya tidak dilacak sebagai stok fisik. Biaya per menit gas tetap digunakan dalam HPP acuan. Bila hendak melacak, isi stok/satuan yang sesuai lalu aktifkan Lacak Stok.
7. Tentukan batas belanja tiap bahan jika sudah ada kebijakan mitra. Batas ini sengaja tidak ditebak.

Sel biru adalah input. Sel abu-abu adalah rumus; jangan ditimpa. Batas tabel tetap: 50 menu, 60 bahan, 25 kalibrasi topping, 1.000 nota, 3.000 item, 6.000 gerak bahan, 300 pelanggan, 1.000 gerak poin dan 1.000 catatan keuangan. Jangan menambahkan data di luar tabel. Arsipkan file sebelum penuh. Satu operator per file; tidak ada sinkronisasi antar laptop.

## Memasang tombol di Excel desktop

Workbook utama `.xlsx` memakai rumus dan jurnal manual. VBA **diberikan sebagai sumber terpisah**, belum tertanam dalam `.xlsx`.

1. Simpan salinan sebagai **Excel Macro-Enabled Workbook (.xlsm)**.
2. `Alt+F11` → pilih workbook yang benar → File → Import File → `Modul_MuritaCrepes.bas`.
3. Debug → Compile VBAProject untuk memeriksa pada versi Excel Anda.
4. Kembali ke Excel → `Alt+F8` → `PasangTombol` → Run. Simpan file.
5. Aktifkan macro hanya untuk file yang dipercaya; jangan menonaktifkan keamanan macro secara global.

Tombol terpasang di KASIR, PRODUKSI, FORM dan STRUK. Pemasangan dapat diulang tanpa menggandakan tombol. Modul tidak menggunakan library tambahan, jaringan, pengiriman pesan atau kredensial. Tekan **Ctrl+S** setelah pencatatan. Macro memeriksa stok/poin historis dan mengembalikan pencatatan baru jika gagal validasi.

Lingkungan pengerjaan Linux tidak memiliki Excel desktop. Rumus dihitung ulang dan diuji dengan LibreOffice. **Kompilasi VBA, tombol, dan ekspor PDF VBA belum diuji di Excel desktop** dan perlu diperiksa sebelum pemakaian operasional.

## Produksi dua tahap

### 1. PREMIKS

Pada PRODUKSI pilih `PREMIKS`, tanggal dan jumlah premiks yang hendak dihasilkan dalam gram. Tabel memperlihatkan pemakaian ketujuh bahan. Klik SIMPAN PRODUKSI.

Jurnal berisi tujuh `PRODUKSI PAKAI` yang mengurangi bahan mentah dan satu `PRODUKSI MASUK` untuk menambah PREMIX. Semua memakai ID dokumen produksi yang sama. Stok mentah harus cukup.

### 2. ADONAN BASAH

Pilih `ADONAN BASAH`. Angka gram pada input adalah **premiks yang dipakai**, bukan berat hasil basah. Dengan input 1.000 g, jurnal mengurangi PREMIX 1.000 g, mentega 65 g dan air 830 g, lalu menambah WET 1.895 g.

Produksi bukan penjualan dan tidak menghasilkan penerimaan kas. Jangan langsung menambahkan WET tanpa pasangan pemakaian bahan, kecuali itu saldo awal yang memang sudah ada atau koreksi hitung fisik dengan alasan.

## Kalibrasi sendok topping

Untuk setiap topping pada KALIBRASI:

1. Ukur tiga sendok/takaran dengan teknik yang sama dan masukkan dalam satuan dasar yang ditampilkan. Untuk topping gram, timbang gram; jangan mengisi angka ml tanpa konversi yang diukur.
2. Isi jumlah sendok per crepes.
3. Rerata tiga sampel × sendok per crepes menjadi pemakaian standar baru.

Sebelum lengkap, sistem memakai takaran `Acuan / Porsi` dari sumber dan memberi label `ACUAN SUMBER`. Isian sampel parsial ditandai untuk dilengkapi. Pisang dan marshmallow mempunyai satuan buah/pcs; takaran langsung per porsi dapat tetap dipakai bila sendok tidak relevan.

Resep dan HPP baru berlaku untuk transaksi berikutnya. Kuantitas dan biaya yang sudah disimpan sebagai snapshot tidak berubah. Catat selisih terhadap hasil hitung fisik sebagai koreksi beralasan, bukan mengubah transaksi lama.

## Penjualan melalui KASIR

1. Isi tanggal, jam, metode pembayaran dan kasir. Pelanggan opsional; kosong berarti transaksi anonim.
2. Pilih SKU, qty dan harga. Harga awal mengikuti katalog; boleh disesuaikan pada transaksi yang disetujui. Maksimal 12 baris per nota; beberapa pcs menu yang sama boleh berada dalam satu baris.
3. Isi diskon rupiah (`0` bila tidak ada). Untuk Tunai, isi uang diterima; kembalian dihitung dari netto. Untuk QRIS/Transfer/E-wallet, nominal tercatat sama dengan netto.
4. SIMPAN NOTA menyimpan header NOTA, baris ITEM serta pemakaian bahan teragregasi per kode ke GERAK BAHAN. Harga, HPP, poin dan kuantitas bahan disimpan sebagai **nilai tetap**.
5. Stok bahan terlacak tidak boleh negatif. Bila WET tidak cukup, periksa produksi; jangan menghapus pemeriksaan stok.
6. Pilih nomor nota pada STRUK untuk cetak/PDF. Cetakan menampilkan diskon sekali per nota, total, diterima, kembalian dan poin.

Jika harga SKM belum diisi, penjualan tetap dapat dicatat, tetapi HPP ditandai **BELUM LENGKAP** dan kontribusi pada dashboard tidak ditampilkan sebagai angka laba. Melengkapi harga master di kemudian hari tidak otomatis mengarang biaya historis: perbaiki snapshot HPP jurnal hanya jika biaya untuk tanggal tersebut benar-benar dapat dikonfirmasi.

## Pembelian, kas dan pemisahan uang pribadi

FORM bagian BARANG mencatat AWAL, BELI, RUSAK, KOREKSI MASUK dan KOREKSI KELUAR. Kuantitas selalu positif; arah gerakan ditentukan jenis. Isi nominal pembelian total dan metode untuk BELI. Koreksi/rusak wajib beralasan. Jangan mencatat jenis produksi atau penjualan melalui form barang ini.

FORM bagian KAS memisahkan:

| Jenis | Dampak kas | Dampak hasil operasional |
| --- | --- | --- |
| BIAYA USAHA | Keluar | Mengurangi kontribusi setelah biaya |
| PRIVE | Keluar | Tidak mengurangi hasil usaha |
| MODAL MASUK | Masuk | Bukan penjualan/laba |
| PENDAPATAN LAIN | Masuk | Ditampilkan pada kas, terpisah dari penjualan crepes |

Pembelian bahan dicatat di GERAK BAHAN saja. Dashboard menampilkan pembelian sebagai arus keluar tersendiri; tidak mengurangkannya lagi dari kontribusi yang sudah memakai HPP bahan. Kontribusi = penjualan netto − HPP acuan snapshot − BIAYA USAHA. Ini bukan laba bersih audit/FIFO.

TUTUP KASIR: isi tanggal, modal awal laci dan uang fisik. Perkiraan isi laci = modal awal + penjualan Tunai netto − BELI Tunai + arus KEUANGAN Tunai bertanda. QRIS/Transfer/E-wallet tidak dimasukkan ke laci. Jangan memasukkan saldo awal yang sama lagi sebagai transaksi MODAL MASUK pada tanggal tersebut.

## CRM dan loyalty

Daftarkan pelanggan hanya bila bersedia, dengan ID unik, nama dan izin `YA`. Kontak opsional; tanpa kontak fitur kunjungan dan poin tetap bekerja. Untuk pelanggan yang tidak bersedia, gunakan transaksi anonim. Tidak ada pelanggan nyata yang dimuat dari sumber.

PELANGGAN menghitung kunjungan **per nota**, total belanja netto, kunjungan terakhir dan poin sampai tanggal akhir dashboard. Membeli tiga menu dalam satu nota berarti satu kunjungan.

Program poin awalnya `TIDAK` pada REFERENSI!B16, ambang rupiah B17 kosong. Setelah aturan disepakati:

1. Isi rupiah netto untuk satu poin, lalu aktifkan `YA`.
2. Poin dihitung bulat ke bawah dari netto per nota terdaftar dan disimpan sebagai snapshot. Tidak ada poin untuk transaksi anonim.
3. Penukaran/koreksi melalui FORM bagian POIN. Isi hadiah/alasan; saldo tidak boleh negatif.
4. Penukaran **tidak otomatis memberikan hadiah atau diskon rupiah**. Jalankan sesuai kebijakan mitra dan catat diskon pada kasir bila memang disepakati.
5. Mengubah ambang poin hanya memengaruhi nota baru. Tidak ada broadcast atau pesan otomatis kepada pelanggan.

## Pembatalan dan jurnal manual

Untuk nota salah: pilih baris pada NOTA → `Alt+F8` → `VoidNota` → isi alasan. Status VOID meniadakan penjualan, pemakaian bahan serta poin dari nota, sementara riwayat dipertahankan. Jika poinnya sudah dipakai, selesaikan pembatalan penukaran terkait lebih dahulu. Jangan menghapus baris ITEM/GERAK untuk menyembunyikan nota.

Untuk koreksi stok non-penjualan, gunakan KOREKSI MASUK/KELUAR dengan alasan. Bila membatalkan produksi secara manual, ubah status **seluruh baris dalam satu ID dokumen** menjadi VOID dan pastikan tidak ada saldo historis negatif. Jangan membatalkan satu sisi produksi saja. Pembatalan atau pengembalian uang tidak otomatis membuat bukti refund; dokumentasikan pelaksanaan pengembalian yang sebenarnya.

Tanpa macro:

1. Isi NOTA kolom A:I serta N (poin snapshot; `0` bila tidak aktif/anonim). Diskon wajib angka, termasuk nol.
2. Isi ITEM kolom A:H dengan ID item unik, nomor nota, SKU, qty, harga snapshot dan HPP/status. HPP kosong harus ditandai `BELUM LENGKAP`, bukan diakui nol.
3. Dari RESEP, salin pemakaian per crepes × jumlah item ke GERAK BAHAN sebagai jenis `JUAL PAKAI`. Isi tanggal yang sama dan No Nota. Pastikan seluruh bahan yang dipakai tercatat, lalu **tempel nilai**, bukan rumus ke resep yang bisa berubah.
4. Nota tanpa pemakaian bahan belum dianggap valid. Semua kolom Cek harus `OK`; tidak boleh ada saldo stok/poin historis negatif.
5. Produksi manual menggunakan pasangan gerakan masuk/pakai per dokumen sebagaimana dijelaskan di atas. Angka rencana pada sheet kalkulator tidak otomatis menjadi transaksi.

Status saat ini memengaruhi laporan historis. File ini tidak memiliki penguncian periode, audit trail perubahan sel atau sinkronisasi. Jangan mengubah kode, tanggal, harga dan jumlah nota final; simpan arsip sebelum koreksi dan sebelum penutupan periode.

## Yang belum tersedia atau perlu dikonfirmasi

- Harga beli/isi bersih SKM dan harga terbaru bahan/menu.
- Stok awal serta saldo kas aktual.
- Kalibrasi sendok dan hasil aktual crepes per kg premiks.
- Takaran pelengkap, terutama SKM, gas dan kemasan.
- Aturan hadiah/poin dan persetujuan pelanggan bila CRM diaktifkan.
- Batas minimum belanja bahan.

Data historis tally/30 hari tetap disediakan sebagai referensi sumber, bukan diubah otomatis menjadi nota: tanggal lintas bulan, pembayaran per nota, stok dan identitas pelanggan belum cukup untuk rekonstruksi tanpa asumsi.
