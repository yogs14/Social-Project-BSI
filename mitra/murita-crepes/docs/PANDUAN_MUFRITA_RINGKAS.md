# Mufrita Crepes — workbook ringkas

Workbook ini menyatukan model Murita yang tersedia di workspace dari 20 sheet menjadi **9 sheet: 8 terlihat, 1 rumus tersembunyi**. Nama tampilan mengikuti permintaan “Mufrita Crepes”; sumber MoM 24 Oktober 2026 dan file biaya menyebut **Murita Crepes — Pak Syarif**. Sel hijau untuk input, abu-abu untuk rumus. File operasional tidak berisi penjualan, pelanggan, pembelian atau saldo fiktif.

## Peta kerja

| Tab | Isi / pintasan |
| --- | --- |
| RINGKASAN | KPI periode di atas; panduan mulai di A81 |
| KASIR | Form penjualan di atas; pilih nota struk di B51; area cetak A53:F81 |
| HARIAN | Produksi di A1; pembelian/stok di A43; kas/prive di A57; poin di A67 |
| STOK | Layar Stok Realtime dari kerangka POS Pak Seger: kartu ringkasan, arus bahan, saldo, status warna dan saran beli |
| TUTUP KASIR | Rekap harian, rincian metode bayar, pencocokan laci, menu laris dan simpan revisi penutupan |
| KATALOG | Menu di A1; kalibrasi topping A63; premiks A101; adonan A135; parameter/sumber A167; master bahan A225 |
| PELANGGAN | Daftar pelanggan opsional, kunjungan per nota, belanja dan saldo poin |
| JURNAL | Nota A1; item A1015; gerak bahan A4029; keuangan A10043; poin A11057; riwayat tutup kasir A12071 |
| _RESEP | Matriks rumus tersembunyi; tidak diisi operator |

Klik tautan baris 3 untuk pindah bagian, baris 4 untuk pindah tab. Jurnal menyimpan tabel terpisah, masing-masing memiliki filter. Gunakan pintasan agar tidak menggulir ribuan baris. Seluruh rumus, nama range, dropdown, grafik, area struk dan sumber VBA sudah disesuaikan ke lokasi baru. File lama tetap tersedia.

## Sebelum dipakai

1. Periksa harga 34 menu pada KATALOG dan harga bahan pada KATALOG > Master Bahan. **Harga SKM di KATALOG!E240 harus diisi dengan harga sebenarnya**; sumber semula merujuk keliru ke total biaya adonan. HPP berstatus BELUM LENGKAP sampai harga lengkap. Margin transaksi dengan HPP belum lengkap tidak dianggap laba.
2. Catat stok fisik awal melalui HARIAN bagian BARANG, jenis AWAL. Kuantitas dalam satuan dasar: gram, buah, pcs atau lembar, bukan jumlah kemasan. Stok nol awal workbook bukan hasil opname.
3. Periksa takaran sumber: SKM 5,63 g, gas 4 menit dan kemasan 0,25 lembar per crepes. Air/gas awalnya tidak dilacak sebagai stok fisik. Tetapkan batas belanja bila sudah diketahui.
4. Kalibrasi topping dengan tiga sampel takaran dan jumlah sendok per porsi pada KATALOG A63. Sebelum lengkap, pemakaian memakai acuan sumber dan tetap merupakan estimasi.
5. File ini untuk satu operator. Cadangkan setiap hari dan arsipkan sebelum tabel penuh.

## Mengaktifkan tombol di Excel desktop

File `.xlsx` tidak berisi macro. Agar form dapat menyimpan otomatis:

1. Simpan salinan sebagai **Excel Macro-Enabled Workbook (.xlsm)**.
2. Alt+F11 → File → Import File → pilih **Modul_Mufrita_Ringkas.bas** dari paket yang sama. Bila salinan `.xlsm` sudah berisi modul versi sebelumnya, simpan cadangan lalu ganti modul lama terlebih dahulu agar tidak ada nama prosedur ganda.
3. Debug → Compile VBAProject, lalu kembali ke Excel.
4. Alt+F8 → PasangTombol → Run; simpan workbook.

Gunakan modul ringkas ini karena alamat selnya berbeda dari modul lama. Tombol mencatat nota, pemakaian stok dan poin bersama, memeriksa jurnal dan saldo, lalu mengembalikan pencatatan baru bila validasi gagal. Tekan Ctrl+S untuk menyimpan ke disk. VBA belum dikompilasi/dijalankan dalam Microsoft Excel; uji tombol dan ekspor PDF di Excel desktop sebelum operasional.

Tanpa macro, rumus dan laporan tetap bekerja, tetapi form tidak otomatis menyimpan. Pengisian manual harus mencatat seluruh tabel terkait sesuai bagian jurnal manual di bawah.

## Produksi dan penjualan

Alur stok: **7 bahan mentah → PREMIX → WET → crepes terjual**.

- HARIAN: pilih PREMIKS, tanggal dan gram hasil kering, lalu SIMPAN PRODUKSI. Tujuh bahan berkurang dan PREMIX bertambah dalam satu dokumen.
- Pilih ADONAN BASAH. Angka gram adalah **premiks yang dipakai**, bukan berat hasil basah. Contoh 1.000 g premiks + 65 g mentega + 830 g air menghasilkan 1.895 g WET untuk sekitar 30 crepes.
- KASIR: isi tanggal/jam, metode bayar, SKU dan qty. Pelanggan opsional. Diskon rupiah dicatat sekali per nota. Untuk Tunai, isi uang diterima. SIMPAN NOTA mengarsipkan harga, HPP, kuantitas bahan dan poin sebagai nilai tetap.
- Penjualan mengurangi WET, pelengkap dan topping. Bahan mentah tidak dipotong lagi sehingga tidak terjadi pemotongan ganda.
- Di KASIR B51 pilih nomor nota, lalu cetak atau EKSPOR PDF. Cetak hanya bagian struk; total diskon tampil sekali.

Resep premiks sumber berjumlah **2.908,5 g**, bukan pembulatan 2.900 g: tepung beras 1.000 g, terigu 500 g, maizena 500 g, fullcream 400 g, gula 500 g, garam 6,5 g dan vanili 2 g. Biaya campuran menurut harga sumber Rp56.731,25. Asumsi 30 crepes/kg dapat diubah pada **KATALOG!B141**; rasio mentega **B142**, air **B143**. Produksi memakai parameter ini; perubahan master tidak mengubah snapshot transaksi yang sudah dicatat.

## Pembelian, keuangan dan tutup hari

HARIAN A43 mencatat AWAL, BELI, RUSAK dan koreksi. BELI memakai total nominal pembelian dan metode bayar; jangan dicatat lagi sebagai BIAYA USAHA.

HARIAN A57 memisahkan BIAYA USAHA, PRIVE, MODAL MASUK dan PENDAPATAN LAIN. Modal/prive memengaruhi kas tetapi bukan hasil operasional. Kontribusi pada Ringkasan = penjualan netto − HPP snapshot − biaya usaha. Pembelian stok tidak dikurangkan lagi setelah HPP. Ini estimasi kontribusi, bukan laporan laba bersih FIFO/audit.

Pada tab **TUTUP KASIR**, isi tanggal **C5**, modal awal laci **I8**, uang fisik **I11** dan nama kasir **C26**. Isi angka **0** bila memang nol; sel kosong berarti belum diisi. Laci seharusnya = modal awal + penjualan tunai netto − pembelian tunai − biaya usaha tunai − prive tunai + modal/pendapatan lain tunai. QRIS/transfer/e-wallet ditampilkan terpisah. Saldo pembukaan yang sama jangan dicatat lagi sebagai MODAL MASUK.

Selisih **I12** = uang fisik − kas seharusnya; pesan menunjukkan cocok, lebih atau kurang. Bila ada selisih, catatan alasan pada **B20** wajib diisi sebelum menyimpan. Ringkasan menghitung hanya nota aktif yang valid pada tanggal pilihan, diskon sekali per nota, jumlah item, HPP acuan dan nota VOID. HPP tidak lengkap tidak ditampilkan sebagai kontribusi kotor. Menu terlaris menampilkan qty dan **bruto sebelum diskon**, dengan penanganan jumlah terjual yang sama agar nama tidak berulang. Jam ramai dihitung dari jumlah nota per jam.

Klik **SIMPAN TUTUP KASIR** setelah modul VBA versi terbaru dipasang. Sistem menambahkan snapshot ke JURNAL, tidak mengubah baris lama. Menyimpan tanggal yang sama lagi membuat nomor revisi berikutnya. Riwayat berada di **JURNAL!A12071**, data pertama **A12077**; tersedia 366 slot penutupan/revisi. Penutupan tidak mengunci periode atau menghentikan perubahan transaksi; simpan revisi baru bila transaksi dikoreksi, lalu tekan Ctrl+S. Tanpa VBA, rekonsiliasi tetap bekerja dan snapshot dapat dicatat sebagai nilai tetap pada tabel riwayat (A:Q), dengan kolom R harus OK.

## Memakai STOK seperti POS Pak Seger

STOK merupakan layar pantau, dengan kartu nilai persediaan, jumlah status merah/kuning, dan estimasi belanja. Tanggal mengikuti akhir periode pada RINGKASAN. Master bahan, harga kemasan, isi bersih, satuan dan batas belanja dipindahkan ke **KATALOG!A225** (data mulai baris 231); klik kode bahan pada kolom N di STOK untuk langsung ke master.

- Stok awal berasal dari AWAL. Masuk mencakup BELI, PRODUKSI MASUK dan KOREKSI MASUK.
- Terpakai mencakup JUAL PAKAI, PRODUKSI PAKAI dan KOREKSI KELUAR. Terbuang berasal dari RUSAK.
- Stok kini = awal + masuk − terpakai − terbuang. Produksi premiks/adonan tidak memotong bahan mentah lagi pada saat penjualan.
- Merah berarti habis/negatif. Kuning berarti stok mencapai batas belanja yang diisi. Hijau berarti di atas batas. BELUM AWAL dan ISI BATAS mengingatkan data yang belum tersedia; tidak dianggap aman. Air/gas yang tidak dilacak ditampilkan terpisah sebagai TIDAK DILACAK.
- Pakai/hari adalah rerata pemakaian jual/produksi selama tujuh hari kalender; cukup (hari) adalah estimasi berdasarkan rerata itu. Tanpa pemakaian, kecukupan dibiarkan kosong.
- Saran beli dihitung saat stok mencapai batas: kekurangan terhadap batas dibulatkan ke kemasan pembelian, minimal satu kemasan. Ini saran, bukan transaksi BELI otomatis. Batas yang belum diisi tidak ditebak.
- Nilai stok dan belanja memakai harga acuan kini, bukan FIFO. Harga yang belum lengkap ditandai CEK HARGA. Baris satuan dapat berbeda (gram, pcs, buah, lembar); jangan menjumlahkan kuantitas lintas bahan sebagai satu stok total.

## Pelanggan dan poin

Pelanggan boleh kosong. Bila pelanggan bersedia dicatat, buat ID unik, nama dan izin YA pada PELANGGAN. Kontak opsional. Kunjungan dihitung per nota, bukan per item. Tidak ada pengiriman pesan otomatis.

Poin awalnya **TIDAK** aktif pada KATALOG!B182; ambang rupiah per poin **B183** kosong. Tetapkan aturan sebelum mengaktifkan. Poin dibulatkan turun dari netto per nota dan disimpan sebagai snapshot. Penukaran dicatat melalui HARIAN A67; tidak otomatis memberikan diskon atau hadiah. Saldo negatif ditolak pada pemeriksaan jurnal.

## Jurnal manual dan pembatalan

Tanpa VBA, tulis nilai tetap pada kolom input tabel berikut. Rumus dan laporan akan menghitung otomatis; form Kasir/Harian bukan jurnal tersimpan.

- NOTA: baris pertama **JURNAL!7**, input A:I dan N (poin snapshot, isi 0 jika nonaktif). Satu baris per nota.
- ITEM: baris pertama **JURNAL!1021**, input A:H. Satu baris per SKU pada nota, harga/HPP/status biaya disalin sebagai nilai tetap.
- GERAK BAHAN: baris pertama **JURNAL!4035**, input A:K. Catat JUAL PAKAI per bahan dari resep × qty, dengan tanggal dan nomor nota yang sama. Tanpa pemakaian, nota ditandai tidak valid. Produksi harus memiliki pasangan PRODUKSI PAKAI dan PRODUKSI MASUK dalam ID dokumen yang sama.
- KEUANGAN: baris pertama **JURNAL!10049**, input A:G.
- POIN: baris pertama **JURNAL!11063**, input A:G.

Semua kolom Cek harus OK; stok/poin historis tidak boleh negatif. Jangan mengubah kode yang sudah dipakai pada jurnal. Pembatalan nota: pilih baris nota pada JURNAL, Alt+F8 → VoidNota, isi alasan. Efek penjualan, stok dan poin dibatalkan tanpa menghapus riwayat. Bila poin telah ditukar, selesaikan penukaran terkait terlebih dahulu. Pembatalan mengubah laporan historis saat ini; arsipkan periode sebelum koreksi. Pembatalan produksi manual harus mencakup semua baris satu ID dokumen. Refund aktual dicatat dan dilaksanakan secara terpisah.

Kapasitas tetap: 50 menu, 60 bahan, 25 kalibrasi, 1.000 nota, 3.000 item, 6.000 gerak bahan, 300 pelanggan, 1.000 gerak poin, 1.000 keuangan, serta 366 snapshot tutup kasir/revisi. Jangan menambah data di luar tabel. Tidak ada sinkronisasi antarlaptop atau penguncian periode otomatis.

## Sumber dan validasi

- Referensi desain: POS_Sate_Ayam_Ponorogo_Pak_Seger.xlsm yang dilampirkan pengguna; kerangka dasar kasir → jurnal → Stok Realtime → Tutup Kasir → dashboard. Layar STOK dan TUTUP KASIR menyalin tata letak, warna dan struktur panel dari file referensi, lalu mengganti rumus restoran dengan model produksi Crepes.
- MoM Murita, 24 Oktober 2026: https://docs.google.com/spreadsheets/d/1MCGsQK3ILZBpxMRH2-fWdnkxv1I8Cfq83BnCVYbJqmQ/edit?gid=52890219
- Data biaya/menu serta rekap: https://drive.google.com/drive/folders/1nqkXUGihEdjwkEBubi3W3uETu7ioPvM9
- Model 20 sheet dan data pendukung yang telah tersedia di workspace dipertahankan sebagai dasar. Database 30 hari/tally tidak dimasukkan sebagai nota karena identitas pembayaran, nota, periode dan stok belum lengkap.

Rumus workbook ringkas dihitung ulang menggunakan LibreOffice dengan skenario sintetis terpisah: produksi dua tahap, penjualan tunai/QRIS, diskon, stok, perubahan resep/harga, cutoff tanggal, VOID, HPP tidak lengkap, penukaran poin, ID duplikat, stok negatif, jumlah item berlebih dan nota tanpa pemakaian bahan. Rincian ada pada hasil_validasi.json. Data uji tidak masuk workbook operasional. Uji tersebut tidak mengeksekusi VBA atau membuktikan kompatibilitas tombol Microsoft Excel.

Revisi UI POS: uji tambahan mencakup saldo STOK versus jurnal untuk seluruh bahan terlacak, produksi dua tahap, merah/kuning/hijau, pembulatan saran belanja, kas nol, kas kurang, HPP tidak lengkap, nota bermasalah dan snapshot revisi penutupan. VBA SimpanTutupKasir disertakan sebagai sumber dan belum dijalankan di Excel desktop.
