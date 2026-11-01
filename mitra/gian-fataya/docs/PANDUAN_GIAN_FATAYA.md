# Gian Fataya — stok konsinyasi dan invoice wilayah

File utama: `POS_Gian_Fataya.xlsx`. Tidak ada penjualan atau saldo awal fiktif di dalamnya.

## Sebelum mulai

1. Buka di Excel desktop, simpan salinan kerja, dan baca sheet **MULAI**. Gunakan satu operator per file; cadangkan sebelum penagihan. Jangan menimpa sel abu-abu yang berisi rumus.
2. Isi **WILAYAH**: PIC penagihan, alamat, kontak, serta cakupan wilayah. Nama delapan wilayah mengikuti foto sumber; beberapa berbeda dari rangkuman MoM, sehingga perlu konfirmasi.
3. Isi **OUTLET** dengan kode unik, nama outlet aktual, kode wilayah, dan status `AKTIF`. `GDG` sudah disediakan untuk gudang rumah Bu Santi. Jangan mengganti kode yang sudah dipakai pada transaksi. `NONAKTIF` mencegah pengiriman/penjualan baru melalui tombol; riwayatnya tetap dihitung. Jika memakai jurnal manual, jangan menambahkan KIRIM/JUAL baru ke outlet NONAKTIF.
4. Isi harga jual/net konsinyasi yang benar di **PRODUK**. Satu produk `GF-PST-001 / Pastel (pouch)` disiapkan dari referensi mitra; kode adalah kode internal baru, bukan kode dari mitra. Jangan memakai Rp9.410 sebagai harga jual: itu hanya jumlah biaya bahan + kemasan dari foto.
5. Isi **LOT**: kode lot unik, SKU, dan kedaluwarsa aktual dari kemasan/catatan. Satu lot hanya boleh menunjuk satu SKU dan satu tanggal kedaluwarsa. Jangan menebak masa simpan.
6. Isi informasi pembayaran pada **REFERENSI!B8**, misalnya instruksi rekening yang sudah diverifikasi. Tidak ada nomor rekening contoh di file.

## Pemakaian langsung tanpa makro

Semua perhitungan, dashboard, dan invoice berjalan dengan rumus Excel. Sheet **FORM** memerlukan VBA untuk tombol simpan; tanpa VBA, isi langsung jurnal berikut.

### MUTASI — satu gerakan per baris

Isi ID unik, tanggal, jenis, lokasi, lot, jumlah positif, status `AKTIF`, dan catatan. Kolom `K:T` adalah rumus. `Harga Jual Snapshot` hanya diisi untuk `JUAL` dan harus berupa **nilai tetap**; jangan menghubungkannya dengan rumus ke harga katalog. `ID Invoice` kosong sampai penjualan ditagih.

| Jenis | Kolom Kode Lokasi | Efek |
| --- | --- | --- |
| AWAL | Lokasi stok aktual, termasuk GDG | Menambah saldo pembukaan lokasi itu |
| PRODUKSI | GDG | Produk jadi masuk ke gudang; tidak memotong bahan baku |
| KIRIM | Outlet tujuan | Gudang berkurang, outlet bertambah; tidak menjadi penjualan |
| JUAL | Outlet yang melaporkan penjualan | Stok outlet berkurang, bruto = qty × harga snapshot |
| RETUR | Outlet asal | Stok outlet berkurang, gudang bertambah |
| RUSAK | Lokasi barang rusak | Stok lokasi berkurang; tidak menjadi penjualan |
| TAMBAH | Lokasi yang dikoreksi | Koreksi tambah, wajib tuliskan alasan |
| KURANG | Lokasi yang dikoreksi | Koreksi kurang, wajib tuliskan alasan |

Untuk `KIRIM`, isi **No PO / Surat Jalan** di kolom U. Kolom V dapat diisi ekspedisi, kolom W ongkir aktual. Ongkir aktual adalah catatan pengeluaran pengiriman; **tidak otomatis ditagihkan kepada mitra**. Ongkir pada invoice diisi tersendiri sesuai kesepakatan.

Gunakan tanggal gerakan yang benar dan hindari mencampur beberapa lot dalam satu baris. Jangan mengulang saldo AWAL setiap awal bulan. Saldo dihitung kumulatif: kirim 40, terjual 30, bulan berikutnya kirim 40 → stok tersisa 50.

Lihat kolom **Cek** dan **Saldo Asal Sesudah**. Baris yang tidak valid dikecualikan dari stok/penjualan, sedangkan kekurangan stok pada jurnal manual ditandai merah. Semua kekurangan harus diselesaikan sebelum penagihan. Macro menolak penyimpanan yang menyebabkan saldo historis negatif.

### STOK dan barang kedaluwarsa

Atur tanggal akhir pada **DASHBOARD!E5**. Sheet STOK menampilkan kombinasi lokasi dan lot yang terdaftar, sehingga tidak perlu membuat rekap stok manual. Filter berdasarkan kode lokasi, lot, atau status.

Barang expired **tetap menjadi stok fisik** sampai dicatat tindakan `RUSAK` atau `RETUR`. Pengiriman dan penjualan melewati tanggal expired tidak valid. Ambang pengingat 30 hari dapat diubah pada **REFERENSI!B12**; itu bukan masa simpan produk.

### Membuat invoice wilayah secara manual

1. Buat satu baris pada **INVOICE**: ID unik, tanggal terbit, kode wilayah, periode awal/akhir, jatuh tempo, diskon (misalnya `10%`, bukan `10`), ongkir (`0` bila tidak ada), status `DRAF`, dan catatan.
2. Pada MUTASI, filter `JUAL`, status `AKTIF`, wilayah/periode yang sesuai, Cek `OK`, dan ID Invoice kosong.
3. Tautkan ID invoice tersebut ke penjualan terpilih di kolom H. **Jangan tautkan baris KIRIM dan jangan menimpa ID invoice penjualan yang sudah ditagih.**
4. Lihat `Bruto Sumber` pada INVOICE (kolom Q). Salin nilainya ke `Bruto Snapshot` (kolom J) memakai **Paste Values / Tempel Nilai**, bukan rumus. Pastikan kolom Cek menjadi `OK`.
5. Periksa diskon dan ongkir, lalu ganti status menjadi `TERBIT`.
6. Pada **CETAK INVOICE!B5**, pilih ID tersebut. Lengkapi PIC dan alamat hingga indikator `SIAP CETAK`.
7. Gunakan File → Print atau Export PDF. Detail mencakup seluruh slot produk; sembunyikan baris produk yang qty-nya nol bila ingin cetakan ringkas. Macro PDF mengerjakannya otomatis lalu mengembalikan tampilan semula.

Rumus invoice: **bruto − pembulatan(bruto × diskon%) + ongkir ditagih**. Tidak ada pajak, komisi konsinyasi, atau biaya tambahan tersembunyi. Format ini sementara karena contoh invoice asli belum tersedia. Harga cetak adalah harga rata-rata jika ada beberapa harga untuk SKU yang sama dalam satu invoice; rinciannya dapat dilihat dengan filter ID invoice di MUTASI.

### PEMBAYARAN — cicilan diperbolehkan

Isi ID pembayaran unik, tanggal, ID invoice yang sudah `TERBIT`, metode, nominal positif, status `AKTIF`, dan catatan. Invoice menghitung sisa piutang dan status `BELUM BAYAR`, `SEBAGIAN`, `LUNAS`, atau `LEBIH BAYAR` secara otomatis. Periksa kelebihan bayar; macro menolaknya.

### Koreksi dan pembatalan

- Jangan menghapus baris riwayat. Gunakan status `VOID` dan alasan.
- Jangan mengubah qty, harga, lot, lokasi, tanggal, atau master terkait penjualan yang sudah ditagih. `SELISIH SUMBER` menandakan snapshot invoice berbeda dari sumbernya dan menghentikan pencetakan melalui macro.
- Untuk membatalkan invoice: VOID semua pembayaran terkait, lepaskan ID invoice pada mutasi terkait sambil mencatat riwayatnya, kemudian ubah invoice menjadi VOID. Penjualan yang dilepas dapat ditagih kembali.
- Sebelum VOID mutasi barang, pastikan tidak menyebabkan stok historis negatif. Transaksi yang sudah memiliki ID invoice harus diselesaikan invoicenya lebih dahulu.
- Master outlet, produk, dan lot tidak dikunci. Jangan mengubah kode/area/identitas setelah digunakan. Ini alat operasional, bukan sistem audit dengan penguncian periode.

## Memasang VBA dan tombol di Excel desktop

1. Simpan salinan file sebagai **Excel Macro-Enabled Workbook (.xlsm)**, misalnya `POS_Gian_Fataya_Operasional.xlsm`.
2. Tekan `Alt+F11` → pilih workbook yang benar → File → Import File → pilih `Modul_GianFataya.bas`.
3. Pilih Debug → Compile VBAProject untuk memeriksa kompilasi pada versi Excel Anda.
4. Kembali ke Excel, `Alt+F8` → `PasangTombol` → Run. Simpan file. Aktifkan macro hanya untuk file yang dipercaya; jangan mengubah keamanan macro global.
5. Gunakan FORM untuk `SimpanMutasi`, `TerbitkanInvoice`, dan `SimpanPembayaran`. Harga jual transaksi harus diketik eksplisit agar tidak berubah ketika katalog berubah. Harga katalog dan stok tersedia ditampilkan sebagai referensi.
6. Pada CETAK INVOICE, pilih ID dan klik `EKSPOR INVOICE PDF`.
7. Pembatalan: pilih baris pada jurnalnya, lalu `Alt+F8` → `VoidMutasi`, `VoidInvoice`, atau `VoidPembayaran`. Alasan wajib diisi. Tombol pemasangan dapat dijalankan ulang tanpa menggandakan tombol.

Macro membuat ID unik, memeriksa input, menolak stok negatif, menautkan hanya penjualan belum ditagih, dan mengembalikan perubahan pencatatan baru jika validasi gagal. Data tersimpan di workbook setelah **Ctrl+S**. Tidak ada koneksi jaringan, pengiriman email, library tambahan atau kredensial di modul VBA.

VBA disertakan sebagai sumber `.bas`, **belum tertanam dalam file `.xlsx`**. Lingkungan pengerjaan Linux tidak memiliki Excel desktop. Rumus diuji dengan LibreOffice; kompilasi, tombol, dan ekspor PDF VBA harus diuji di Excel desktop sebelum pemakaian operasional.

## Membaca dashboard

- Penjualan bruto: hanya `JUAL` valid dalam periode.
- Tagihan terbit: snapshot invoice `TERBIT`, berdasarkan tanggal terbit. Snapshot tetap tampil jika sumber berubah; lihat peringatan invoice bermasalah.
- Pembayaran masuk: pembayaran valid berdasarkan tanggal penerimaan.
- Piutang sampai akhir: invoice terbit sampai tanggal akhir dikurangi pembayaran sampai tanggal itu; pembayaran bulan berikutnya tidak mengurangi piutang historis.
- Stok: seluruh gerakan valid sampai tanggal akhir, tidak direset tiap bulan.
- Grafik: penjualan bruto per wilayah, sesuai kebutuhan visualisasi area di MoM.

Status VOID saat ini memengaruhi laporan historis. File ini tidak menyimpan histori perubahan status berdasarkan waktu. Jika memerlukan pembukuan final per periode, simpan arsip bulanan sebelum melakukan koreksi.

## Kapasitas dan data yang masih dibutuhkan

20 wilayah; 60 outlet + gudang; 50 produk; 80 lot; 2.000 gerakan; 400 invoice; 1.000 pembayaran. Jangan menambah baris di luar area tabel: rumus/validasi dan macro memakai batas tetap. File sumber generator disediakan untuk pengembangan kapasitas berikutnya.

Yang perlu dilengkapi oleh mitra: harga jual atau harga net konsinyasi; daftar outlet dan PIC; saldo awal per lot/lokasi; kedaluwarsa aktual; kesepakatan diskon/ongkir; informasi pembayaran; format invoice asli. Tidak ada asumsi harga, stok atau outlet fiktif dalam file utama.
