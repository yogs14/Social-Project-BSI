# Chicken Al-Haqi — Mba Gita

Workbook utama: `POS_Chicken_Al_Haqi.xlsx`. Tombol tambahan: `Modul_ChickenAlHaqi.bas`.

MoM **23 Oktober 2026** meminta dashboard lebih sederhana karena menu sedikit, dengan pencatatan yang memisahkan uang ayam dari uang pribadi dan usaha obat. Karena itu, tampilan hanya mempunyai tujuh tab utama: **MULAI, KASIR, HARIAN, BELANJA, STOK, KAS USAHA, SETUP**. Jurnal NOTA dan ITEM disembunyikan untuk audit, bukan dihapus. Tidak ada modul gudang, CRM, loyalty atau peramalan.

File utama tidak memuat transaksi, stok awal atau saldo kas fiktif. Data sintetis pada pengujian tidak masuk ke file operasional.

## Menu dan sumber harga

| Potongan/menu | Harga dasar |
| --- | ---: |
| Sayap | Rp8.000 |
| Paha Bawah | Rp8.000 |
| Paha Atas | Rp9.000 |
| Dada | Rp9.000 |
| Nasi | Rp3.000/porsi |

Rasa Original memakai harga dasar. Sambal Geprek, Saus Keju dan Hot Spicy masing-masing menambah **Rp3.000 per potong**. Nilai tersebut berasal dari sel harga yang digabung untuk ketiga rasa pada file sumber, bukan menebak harga kosong.

Contoh perhitungan: 2 dada saus keju = 2 × (Rp9.000 + Rp3.000) = Rp24.000. Dua nasi menambah Rp6.000. Saus tidak ditambahkan sekali saja per nota.

Harga dapat diperiksa pada SETUP. Nominal di nota lama disimpan sebagai snapshot dan tidak ikut berubah ketika harga menu atau saus diganti.

## Persiapan singkat

1. Buka MULAI dan periksa harga pada SETUP.
2. Isi STOK: tanggal saldo awal dan jumlah potongan fisik untuk empat potongan ayam. Gunakan awal hari pada tanggal itu. Kedua isian dibiarkan kosong bila belum diketahui; sistem menampilkan **BELUM AWAL**, bukan nol yang dianggap pasti.
3. Nasi awalnya **TIDAK DILACAK**, karena konversi beras menjadi porsi belum tersedia. Bila ingin melacak nasi, isi jumlah porsi fisik dan aktifkan Lacak Stok pada SETUP; jangan memasukkan kg beras sebagai porsi nasi.
4. Masukkan nama kasir sekali pada KASIR. Default metode pembayaran adalah Tunai, dengan pilihan QRIS, Transfer dan Lainnya.
5. Untuk tutup hari, isi tanggal, kas awal usaha ayam dan uang fisik usaha ayam pada HARIAN. Saat mengganti tanggal laporan, **isi ulang kas awal dan fisik sesuai tanggal tersebut**; dua nilai tersebut adalah kontrol laporan, bukan jurnal saldo tersimpan per hari.

Waktu sumber: hari biasa 09.00–21.00, weekend mulai 06.30/07.00. Jam pagi tersebut diterima. Tidak ada pembatas pukul 07.00 yang menghalangi penjualan weekend. Pengingat MoM: usahakan belanja/pesan ke pemasok sebelum lewat 07.00 agar pilihan potongan tidak habis. Modul tidak mengirim pesanan atau pesan otomatis.

## Mengaktifkan tombol di Excel desktop

Workbook `.xlsx` memakai rumus dan jurnal manual; VBA disertakan terpisah dan **belum tertanam**.

1. Simpan salinan sebagai **Excel Macro-Enabled Workbook (.xlsm)**.
2. `Alt+F11` → pilih workbook yang benar → File → Import File → pilih `Modul_ChickenAlHaqi.bas`.
3. Debug → Compile VBAProject untuk memeriksa kompilasi pada versi Excel Anda.
4. Kembali ke Excel, `Alt+F8` → PasangTombol → Run, lalu simpan.
5. Aktifkan macro hanya untuk file yang dipercaya. Tidak ada library tambahan, koneksi jaringan atau kredensial dalam modul.

Lima tombol berada pada KASIR: SIMPAN NOTA, NOTA BARU, BATAL NOTA, SIMPAN BELANJA dan SIMPAN KAS. Pemasangan dapat diulang tanpa menggandakan tombol. Gunakan **Ctrl+S** untuk menyimpan file setelah pencatatan. Tombol menyimpan ke jurnal workbook, bukan mengirim data ke pihak lain.

Rumus diuji dengan LibreOffice. **Kompilasi dan eksekusi VBA belum diuji di Excel desktop** karena lingkungan pengerjaan Linux tidak memiliki Excel desktop; lakukan pemeriksaan tersebut sebelum operasional.

## Penjualan

1. Isi tanggal, jam, kasir dan metode bayar pada KASIR.
2. Pilih potongan dan rasa pada salah satu dari enam baris ayam, kemudian isi jumlah pcs. Harga dihitung dari potongan + rasa. Harga pada keranjang dapat disesuaikan jika ada harga transaksi yang benar-benar disepakati; nilai itu disimpan tetap.
3. Isi nasi pada kotak **Nasi (porsi)** bila ada. Nasi tidak diberi pilihan saus, sehingga tidak menjadi menu nasi saus fiktif.
4. Isi diskon rupiah (`0` bila tidak ada) sekali per nota.
5. Untuk Tunai, isi uang diterima; kembalian dihitung dari netto. Untuk non-tunai, penerimaan tercatat sebesar netto, tidak sebagai isi laci.
6. Klik SIMPAN NOTA. Sistem menyimpan header NOTA dan ITEM secara bersama, serta mengosongkan keranjang setelah validasi. Batas per nota enam baris ayam + satu baris nasi.

Jika stok awal belum diisi, pencatatan uang tetap berjalan tetapi kecukupan stok belum dapat diverifikasi. Periksa peringatan BELUM AWAL. Setelah saldo awal nyata tersedia, stok dihitung dari pembukaan + belanja aktual − penjualan + koreksi; stok yang diketahui tidak boleh negatif. Tidak ada pemisahan stok ayam mentah, dimarinasi atau matang. Semua dinyatakan sebagai potongan yang perlu direkonsiliasi secara fisik.

## Belanja kiloan dan pcs aktual

Gunakan panel BELANJA AYAM di kanan KASIR atau jurnal BELANJA:

- Tanggal, potongan/menu, kg jika diketahui, jumlah pcs/porsi **aktual**, nominal total dan metode.
- Dada/paha atas memakai perkiraan 9 pcs/kg dari MoM; sayap/paha bawah 10 pcs/kg.
- Perkiraan hanya pembanding. **2 kg dada yang ternyata berisi 17 pcs menambah stok 17, bukan 18, bukan 19, dan bukan 2 + 17.** Selisih terhadap perkiraan ditampilkan dan bukan otomatis dianggap kesalahan.
- Acuan nominal kg memakai harga ayam sumber Rp35.000/kg; nominal belanja tetap harus diisi sesuai pembayaran yang sebenarnya.
- Belanja lebih dari 3 kg diperbolehkan. Batas 3 kg pada MoM adalah ukuran batch membumbui, bukan batas pembelian.
- Bila membeli per pcs tanpa data kg, biarkan kg kosong. Jika nasi dibeli sebagai porsi siap jual, gunakan jumlah porsi; kolom kg harus kosong/nol.

Untuk **nota pembelian campur beberapa potongan**, buat satu baris per potongan dengan pcs aktual. Jangan menyalin kg total ke setiap baris. Jika berat masing-masing potongan tidak diketahui, biarkan kg kosong dan catat berat total pada catatan bukti. Alokasikan nominal berdasarkan bukti; jika hanya total invoice yang diketahui, catat nominal total sekali dan gunakan `0` pada baris potongan lain yang termasuk invoice yang sama. Tuliskan referensi bukti yang sama. Ini mencegah uang keluar terhitung berkali-kali; modul tidak menebak biaya per potongan dari invoice campuran.

Pembelian bumbu, minyak, beras atau bahan pendukung yang tidak mempunyai konversi stok di modul ini dicatat sebagai **BIAYA USAHA** dengan catatan yang jelas, bukan pcs ayam tambahan. Pembelian ayam/nasi yang sudah masuk BELANJA tidak dicatat lagi sebagai biaya kas kedua.

## HPP acuan dan rencana marinasi

File sumber mencantumkan biaya per pcs: ayam Rp3.500, terigu dan bumbu Rp2.000, marinasi Rp500, minyak Rp1.000; jumlahnya **Rp7.000**. Nilai itu dimuat sebagai **acuan umum belum dikonfirmasi**, bukan pernyataan setiap potongan memiliki HPP lengkap Rp7.000.

Harga/takaran 1 kg pada sumber (ayam Rp35.000, racikan terigu Rp15.000 dan marinasi Rp3.000) tidak otomatis dijumlahkan lagi ke biaya per pcs. Takaran fisik/rincian peruntukannya belum cukup untuk membuat BOM; modul tidak menghitung bahan yang sama dua kali.

Harga pokok saus dan nasi belum tersedia dan dibiarkan kosong. Biaya kemasan, gas dan overhead juga belum lengkap. Pada SETUP, konfirmasi HPP per potongan/rasa hanya setelah nilainya dapat dipastikan. Sebelum seluruh item pada nota dikonfirmasi, perbandingan netto–HPP pada HARIAN ditampilkan sebagai **HPP BELUM DIKONFIRMASI**. Ini tetap **bukan laba bersih**; biaya usaha lain ditampilkan terpisah.

HPP tiap item disimpan sebagai snapshot. Mengisi harga pokok master belakangan tidak mengarang biaya nota lama; koreksi snapshot hanya bila biaya pada tanggal transaksi benar-benar dapat dibuktikan.

Bagian bawah SETUP menyediakan kalkulator ringan rencana kg: perkiraan 9/10 pcs per kg dan jumlah batch paling banyak 3 kg. Rencana 5 kg menjadi 2 batch. Kalkulator bukan transaksi produksi dan tidak menambah stok. Takaran gram racikan/marinasi tidak dibuat-buat dari angka harga.

## Pisahkan kas ayam, pribadi dan usaha obat

Panel KAS USAHA mencatat arus yang benar-benar masuk/keluar dari uang usaha ayam:

| Jenis | Dampak kas ayam | Perlakuan |
| --- | --- | --- |
| BIAYA USAHA | Keluar | Bumbu, minyak, beras, biaya operasional aktual |
| MODAL MASUK | Masuk | Bukan penjualan ayam |
| PRIVE | Keluar | Pengambilan pribadi; bukan biaya operasional ayam |
| KE USAHA LAIN | Keluar | Misalnya uang ayam dipakai untuk usaha obat; bukan penjualan/biaya ayam |
| DARI USAHA LAIN | Masuk | Pengembalian/transfer antar-usaha; bukan omzet ayam |

Nominal selalu positif; arah arus ditentukan jenis. Metode menentukan apakah uang masuk/keluar laci tunai atau non-tunai. Tidak ada produk maupun penjualan obat pada kasir ayam.

HARIAN menghitung:

**Kas awal ayam + penjualan Tunai netto − BELANJA Tunai + KAS USAHA Tunai bertanda.**

QRIS/Transfer/Lainnya ditampilkan terpisah. Uang diterima sebelum kembalian tidak dianggap seluruhnya sebagai pendapatan. Hitung uang fisik bagian usaha ayam saja. Jika uang obat/pribadi masih berada pada laci yang sama, identifikasi dan catat perpindahannya; uang tersebut tidak otomatis menjadi omzet ayam. Jangan mencatat kas awal lagi sebagai modal masuk pada tanggal yang sama.

## Stok rusak/koreksi dan pembatalan

STOK menampilkan saldo sampai tanggal HARIAN. Pada tabel bawahnya (mulai baris 36), catat RUSAK, TAMBAH atau KURANG, tanggal, kode, jumlah positif, alasan dan status AKTIF. Koreksi bukan pembelian atau penjualan tambahan. Jangan memakai stok untuk menutup selisih uang.

Saldo dihitung pada tingkat **tanggal**: belanja/koreksi satu tanggal dianggap tersedia pada hari itu, bukan urutan jam. Saldo awal adalah pembukaan sebelum gerakan pada tanggal pembukaan. Transaksi sebelum tanggal awal stok tetap muncul pada laporan uang, tetapi tidak dipotong lagi dari saldo cut-off. Jangan mengganti tanggal/jumlah saldo awal setelah digunakan tanpa rekonsiliasi fisik dan arsip.

BATAL NOTA menampilkan nomor nota terakhir sebagai pilihan awal; periksa nomor yang benar dan isi alasan. Status VOID meniadakan penjualan dan pemakaian potongan terkait, sambil menyimpan riwayat. Pelaksanaan refund aktual harus dilakukan dan didokumentasikan; modul tidak mengirim refund ke bank/QRIS. Status VOID saat ini memengaruhi laporan historis; arsipkan sebelum koreksi periode lama.

## Jurnal manual dan batas

Tanpa macro, buka NOTA/ITEM melalui klik kanan tab → **Unhide**. Isi NOTA A:I, serta ITEM A:H dengan ID unik. Nilai harga/HPP harus ditempel sebagai nilai, bukan tautan ke master. Gunakan status HPP `DIKONFIRMASI` atau `BELUM DIKONFIRMASI`. ITEM untuk nasi harus memakai rasa `Original` (label internal tanpa saus). Kolom lain adalah rumus.

BELANJA dapat diisi langsung A:I. KAS USAHA A:G. Koreksi STOK A:G pada baris 36–235. ID tidak boleh ganda, dan jangan menghapus riwayat untuk menyembunyikan kesalahan. Gunakan status VOID bila relevan. `Alt+F8 > PeriksaData` memeriksa jurnal dan saldo historis yang diketahui.

Kapasitas: 1.500 nota, 6.000 item, 1.500 belanja, 1.500 arus kas, 20 slot master/stok (kasir awal empat potongan + nasi), 200 koreksi. Jangan menambah baris di luar tabel; arsipkan dan bawa saldo cut-off sekali bila file sudah penuh. Kasir enam baris ayam + satu nasi; perlu pengembangan terpisah untuk menu baru di luar pola tersebut. Tidak ada sinkronisasi antar-laptop, audit trail perubahan sel atau penguncian periode.

Yang masih perlu diisi/dikonfirmasi: stok fisik awal, kas awal/fisik, harga terkini, HPP per potongan, HPP saus/nasi dan rincian takaran bahan bila kelak ingin mengelola stok bumbu. Data sumber tetap disimpan beserta URL dan SHA256 pada `data/profile.json`.
