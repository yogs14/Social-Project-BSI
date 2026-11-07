# Gian Fataya — Excel konsinyasi ringkas v2

Mulai dari **BERANDA**. Versi ini menggabungkan 14 sheet lama menjadi **6 sheet nyata**, semuanya terlihat. File operasional tidak berisi transaksi, harga jual, saldo, outlet, atau rekening fiktif. Simpan salinan kerja dan cadangkan sebelum penagihan.

Dasar revisi: [MoM/Notulensi!G4:G6, kolom Updated](https://docs.google.com/spreadsheets/d/1MCGsQK3ILZBpxMRH2-fWdnkxv1I8Cfq83BnCVYbJqmQ/edit?gid=52890219#gid=52890219), dibaca ulang 7 November 2026. Rapat asli bertanggal 22 Oktober 2026.

| Sheet | Dipakai untuk |
| --- | --- |
| BERANDA | Periode, angka ringkas, peringatan, grafik dan performa area; form VBA opsional di bawah |
| TRANSAKSI | Satu jurnal stok awal, barang masuk, kirim PO, jual, retur, rusak dan koreksi |
| STOK | Saldo otomatis per lokasi/lot sampai tanggal akhir Beranda |
| TAGIHAN | Invoice wilayah dan pembayaran/cicilan dalam dua tabel |
| CETAK INVOICE | Pilih invoice, periksa kesiapan, cetak atau ekspor PDF |
| DATA USAHA | Wilayah, outlet, produk, lot, HPP acuan dan pengaturan |

Tautan baris 4 tersedia pada semua sheet. Baris 5 DATA USAHA menuju setiap bagian. TAGIHAN memiliki tautan langsung ke pembayaran; tidak perlu menggulir ratusan baris. Ctrl+G → tblPembayaran juga menuju tabel tersebut.

Sel biru untuk input; kolom otomatis memakai header dan latar abu-abu. Peringatan memiliki teks selain warna. Jangan menimpa rumus.

## Persiapan sekali di DATA USAHA

1. **Wilayah — A6:G26**: isi PJ/PIC, alamat invoice dan kontak. Delapan wilayah mengikuti Updated: Jawa Timur, Jawa Tengah, Medan–Aceh, Sulawesi Selatan, Kalimantan, Sumatera Selatan, Batam, Jabodetabek.
2. **Outlet — A34:I95**: isi kode unik, nama outlet aktual, wilayah penagihan, status AKTIF, dan kelompok **Amanda** atau **Independen**. Toko independen tetap memiliki wilayah penagihan. GDG adalah gudang rumah Bu Santi, kelompok PUSAT.
3. **Produk — A102:H152**: isi harga jual/net konsinyasi aktual. Pastel (pouch), kode internal GF-PST-001, sudah disediakan. **Rp9.410 adalah acuan bahan + kemasan, bukan harga jual atau HPP lengkap.**
4. **Lot — A160:F240**: isi kode lot unik, SKU dan kedaluwarsa aktual. Satu kode hanya untuk satu SKU/tanggal.
5. **Pengaturan**: info pembayaran **B270**, alamat pengirim **B272**, pengingat expired dalam hari **B274**, ambang performa relatif **B276**. Jangan mengarang rekening atau masa simpan.
6. Catat stok fisik pembukaan per lokasi/lot sebagai AWAL di TRANSAKSI. Jangan ulangi AWAL setiap bulan.

HPP acuan berisi enam komponen sumber: tepung, minyak, abon, margarin, telur dan pouch. Hanya referensi; tidak ada pengelolaan pembelian/stok bahan baku.

## Pemakaian harian tanpa makro

Semua rumus langsung bekerja di .xlsx. Isi tabel TRANSAKSI mulai baris 7.

| Kolom | Isi |
| --- | --- |
| A–F | ID unik, tanggal, jenis, kode lokasi, kode lot, qty positif bulat |
| G | Harga jual sebagai **nilai tetap**, wajib hanya untuk JUAL |
| H | Nomor PO/surat jalan, wajib untuk KIRIM |
| I | Ongkir aktual; tidak otomatis ditagih ke mitra |
| J–L | Otomatis: Cek, Cek minimum PO, saldo asal sesudah |
| M–P | Detail: ID invoice, ekspedisi, status, catatan/alasan |
| Q–X | Perhitungan pendukung; jangan diubah |

Kolom M:X dikelompokkan agar tampilan harian ringkas. Klik **+** di atas kolom untuk membukanya. Status O sudah AKTIF; baris tanpa ID tetap tidak dihitung. Alasan retur/koreksi pada **P** wajib diisi.

| Jenis | Lokasi | Akibat |
| --- | --- | --- |
| AWAL | Lokasi stok aktual | Saldo pembukaan |
| PRODUKSI | GDG | Produk jadi masuk; tidak memotong bahan baku |
| KIRIM | Outlet tujuan | Gudang berkurang, outlet bertambah; belum penjualan |
| JUAL | Outlet pelapor | Stok outlet berkurang, menjadi penjualan |
| RETUR | Outlet asal | Kembali ke gudang; khusus kerusakan fisik kemasan, wajib alasan |
| RUSAK | Lokasi barang | Mengurangi stok dengan alasan; bukan penjualan |
| TAMBAH / KURANG | Lokasi koreksi | Koreksi stok dengan alasan |

**PO minimal 10 pcs per nomor PO + outlet**, boleh beberapa lot. Entri pertama 6 pcs tetap mencatat perpindahan fisik dan menampilkan PO < 10 PCS; lot kedua 4 pcs dengan PO/outlet sama membuatnya OK. Jangan gunakan ulang nomor yang sama untuk pesanan lain. Penjualan/retur boleh kurang dari 10 pcs. Pastikan **BERANDA!I57 = 0** sebelum penagihan; tombol invoice/PDF menolak PO belum lengkap.

Kirim 40 → terjual 30 → sisa 10; bulan berikutnya kirim 40 → sisa 50. Tidak ada reset bulanan. Pengiriman/penjualan setelah expired ditolak rumus. Barang expired tetap stok fisik sampai ada tindakan retur/rusak; tidak otomatis dianggap terjual.

J harus OK, K tanpa peringatan, L tidak negatif. Input tidak valid dikecualikan dari stok/penjualan; saldo negatif manual ditandai. Selesaikan sebelum menagih. Excel mewajibkan alasan retur, tetapi kondisi fisik kemasan tetap diperiksa operator.

## Invoice wilayah dan pembayaran

1. Di **TAGIHAN**, mulai baris 7, isi ID invoice unik, tanggal terbit, wilayah, periode awal/akhir, jatuh tempo, diskon persen, ongkir **ditagih**, status DRAF.
2. Di TRANSAKSI buka grup (+). Filter JUAL, wilayah/periode sesuai, status AKTIF, J OK, dan **M (ID invoice) kosong**.
3. Isi ID invoice di M untuk penjualan terpilih saja. Jangan tautkan KIRIM atau mengganti ID pada penjualan yang sudah ditagih.
4. Pada TAGIHAN, salin **Q Bruto Sumber** ke **J Bruto Snapshot** dengan **Paste Values / Tempel Nilai**. Jangan pakai rumus. Diskon ditulis 10%, bukan 10; ongkir 0 bila tidak ada.
5. Pastikan P OK dan peringatan Beranda selesai, lalu ubah I menjadi TERBIT.
6. Pilih ID pada **CETAK INVOICE!B5** sampai SIAP CETAK. Gunakan File → Print/Export PDF. Sembunyikan baris detail qty nol untuk cetakan manual ringkas; makro melakukannya otomatis.
7. Klik **Pembayaran / cicilan** pada TAGIHAN. Header tabel kedua **A436:J436**, input **baris 437**. Isi A:G: ID pembayaran, tanggal, ID invoice TERBIT, metode, nominal positif, status AKTIF, catatan. H:J otomatis. Cicilan diperbolehkan.

Tagihan = **bruto − pembulatan(bruto × diskon%) + ongkir ditagih**. Ongkir aktual pengiriman terpisah dari ongkir yang ditagih. Tidak ada pajak/komisi tambahan yang diasumsikan.

Status bayar: BELUM BAYAR, SEBAGIAN, LUNAS, LEBIH BAYAR. Periksa kelebihan bayar; makro menolaknya. Harga cetak menjadi rata-rata jika satu SKU memakai beberapa harga dalam invoice. Format invoice masih perlu dicocokkan dengan contoh asli yang disetujui mitra.

## Tombol otomatis dan PDF di Excel desktop

1. Simpan salinan sebagai **Excel Macro-Enabled Workbook (.xlsm)**.
2. Alt+F11 → pilih project salinan → File → Import File → Modul_GianFataya.bas **v2 dari paket ini**.
3. Debug → Compile VBAProject. Kembali ke Excel: Alt+F8 → PasangTombol → Run. Simpan; aktifkan makro hanya pada file terpercaya.
4. Pada BERANDA, tautan **Form barang / Form invoice / Form pembayaran** di baris 72 menuju form di bawah. Tidak ada sheet form tambahan.
5. SimpanMutasi mencatat nilai tetap dan memeriksa stok. PO beberapa lot dapat dilengkapi bertahap.
6. **TERBITKAN + PDF** mengambil JUAL aktif, valid, belum ditagih pada area/periode terpilih. Setelah terbit dan siap cetak, dialog penyimpanan PDF terbuka. Jika PIC/alamat belum lengkap, lengkapi dahulu, lalu gunakan **EKSPOR INVOICE PDF** di CETAK INVOICE.
7. SimpanPembayaran mencatat cicilan/pelunasan. Tekan Ctrl+S.

Membatalkan dialog PDF tidak membatalkan invoice yang sudah terbit; ekspor ulang dari CETAK INVOICE.

**VBA terpisah sebagai .bas, belum tertanam di .xlsx.** Rumus diuji dengan LibreOffice. Lingkungan Linux ini tidak menjalankan Excel desktop; kompilasi, tombol, dialog dan PDF VBA belum dieksekusi. Jangan gunakan modul v1 untuk tata letak v2.

## Membaca Beranda

- **B5/E5**: awal/akhir periode. Penjualan memakai tanggal JUAL; tagihan tanggal terbit; kas tanggal pembayaran.
- Piutang/stok sampai tanggal akhir; pembayaran bulan berikutnya tidak mengurangi piutang historis.
- **Grafik batang di A20** membandingkan bruto delapan wilayah pada periode B5–E5. Warna mengikuti kategori; nama wilayah juga menampilkan teks TINGGI/SEDANG/DATAR dan angka nol tetap terlihat. Toko independen masuk wilayahnya.
- **Grafik area di G20** menunjukkan tren 12 bulan. Pilih nama wilayah pada sel biru **I18**. Rentang tren selalu 12 bulan sampai E5; B5 hanya mengubah perbandingan, bukan rentang tren. Bulan terakhir dihitung sampai tanggal E5 dan bisa belum penuh.
- Hanya transaksi JUAL yang valid dan AKTIF yang menjadi penjualan; KIRIM, retur, VOID, dan input tidak valid tidak menambah grafik. Bulan tanpa penjualan bernilai nol.
- **Rincian A36:H44** memuat qty, bruto, tagihan, pembayaran, stok dan kategori. Kapasitas wilayah tambahan dilipat di baris 45–56; klik + jika diperlukan.
- File utama kosong menampilkan petunjuk “Belum ada penjualan”. Untuk mencoba grafik, gunakan **contoh/SIMULASI_Grafik_Wilayah.xlsx** dalam paket; data dan harga di sana sepenuhnya fiktif.
- **DATAR** = bruto nol. **TINGGI** = bruto minimal ambang × bruto area teratas. **SEDANG** = positif di bawahnya. Default **50%** dapat diubah di DATA USAHA!B276; ini pilihan desain, bukan target notulensi atau ukuran pertumbuhan. File kosong menampilkan nol, bukan transaksi contoh.
- Periksa input salah, saldo negatif, invoice/pembayaran bermasalah, expired, penjualan belum ditagih, dan PO belum lengkap.

## Koreksi, kapasitas dan arsip

Gunakan VOID dengan alasan; jangan hapus riwayat. Jangan mengubah qty, harga, lot, lokasi, tanggal/master pada penjualan yang sudah ditagih. SELISIH SUMBER menghentikan PDF makro.

Batalkan invoice dengan VOID seluruh pembayaran terkait, lepaskan ID invoice pada penjualan sambil mencatat riwayatnya, kemudian VOID invoice. Untuk tombol, pilih baris tabel TRANSAKSI/TAGIHAN → Alt+F8 → VoidMutasi, VoidInvoice atau VoidPembayaran.

Grafik mengikuti delapan wilayah resmi dari Updated. Jika memperluas daftar wilayah, perbarui data/profile.json dan hasilkan ulang workbook agar grafik ikut diperluas.

Kapasitas tabel: 20 wilayah, 60 outlet + gudang, 50 produk, 80 lot, 2.000 gerakan, 400 invoice, 1.000 pembayaran. Jangan menambah di luar tabel. Satu operator; tanpa sinkronisasi atau penguncian periode. Perubahan status sekarang memengaruhi laporan historis; arsipkan bulanan sebelum koreksi. Jangan mengganti kode/identitas master yang sudah dipakai.

Riset dan pemetaan kebutuhan: RISET_UI_UX_GIAN_FATAYA.md. Latihan PDF dipisahkan dari file operasional.
