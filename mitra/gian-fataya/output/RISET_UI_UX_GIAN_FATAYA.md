# Riset dan keputusan UI/UX — Gian Fataya v2

Penelaahan awal: 4 November 2026; verifikasi ulang MoM dan revisi grafik: 7 November 2026. Lingkup: Excel konsinyasi Gian Fataya, satu operator, tanpa layanan berlangganan.

## Dasar kebutuhan

[MoM/Notulensi!G4:G6, kolom Updated](https://docs.google.com/spreadsheets/d/1MCGsQK3ILZBpxMRH2-fWdnkxv1I8Cfq83BnCVYbJqmQ/edit?gid=52890219#gid=52890219) menjadi dasar revisi. Teks tersimpan di data/updated_mom.txt. Notulensi hanya dibaca.

Fokusnya stok produk jadi → distribusi PO → penjualan → stok akhir → invoice area. Delapan area resmi dan toko independen harus tertampung. HPP enam bahan cukup sebagai referensi. Ongkir aktual, ongkir ditagih dan diskon dibedakan; kirim belum berarti penjualan.

Audit awal menemukan 14 sheet, empat master terpisah, panduan/dashboard/form terpisah, dan input PO/ongkir jauh di kolom U:W. Pengguna mengeluhkan terlalu banyak sheet.

## Sumber primer dan penerapan

| Sumber | Temuan yang digunakan | Penerapan |
| --- | --- | --- |
| [NN/G — Progressive Disclosure](https://www.nngroup.com/articles/progressive-disclosure/) | Tampilkan fungsi utama; rincian jarang dipakai memiliki jalur masuk yang jelas. | Enam tab berdasarkan tugas; detail Transaksi dibuka lewat grup (+); pengaturan satu sheet dengan tautan bagian. |
| [NN/G — Recognition and Recall](https://www.nngroup.com/articles/recognition-and-recall/) | Pilihan terlihat dan petunjuk kontekstual mengurangi kebutuhan mengingat. | Navigasi baris 4, dropdown, label input/otomatis, tautan langsung pembayaran, petunjuk dekat tabel. |
| [Microsoft — Organizing and formatting data](https://support.microsoft.com/en-us/excel/guidelines-for-organizing-and-formatting-data-on-a-worksheet) | Data sejenis memakai kolom/header konsisten; Excel Table membantu filter. | Delapan tabel tetap terpisah secara skema meskipun master/pembayaran digabung secara fisik dalam sheet. |
| [Microsoft — Excel accessibility](https://support.microsoft.com/en-us/accessibility/excel/accessibility-best-practices-with-excel-spreadsheets) | Nama sheet bermakna, tabel sederhana, kontras, dan label tautan jelas membantu navigasi. | Nama tab sesuai tugas, teks di A1, peringatan teks selain warna, tanpa merge dalam tabel. Merge hanya pada judul, petunjuk dan format cetak. |

Ini kajian pustaka dan evaluasi desain. Belum ada pengujian kegunaan dengan Bu Santi; tidak ada klaim pengurangan waktu yang sudah diukur.

## Perubahan terukur

| Sebelumnya | Sesudah |
| --- | --- |
| 14 tab | 6 tab nyata, semua terlihat |
| MULAI + DASHBOARD + FORM | BERANDA; form opsional di bawah |
| MUTASI, PO/ongkir U:W | TRANSAKSI, input utama A:I dan pemeriksaan J:L |
| INVOICE + PEMBAYARAN | TAGIHAN, dua tabel dengan tautan langsung |
| WILAYAH + OUTLET + PRODUK + LOT + BAHAN + REFERENSI | DATA USAHA, enam tautan bagian |
| Outlet tanpa kelompok | Amanda / Independen |
| Grafik di bawah tabel panjang | Grafik batang delapan wilayah + grafik area tren 12 bulan di atas rincian; dropdown wilayah, label angka dan TINGGI/SEDANG/DATAR |

Operasionalisasi performa: DATAR = bruto 0; TINGGI ≥ ambang × bruto area teratas; SEDANG = positif di bawahnya. Default 50% dapat diubah. Angka ini pilihan desain, bukan ketentuan MoM, target usaha, atau tren pertumbuhan.

## Pemetaan Updated ke implementasi

| Kebutuhan | Lokasi / perilaku |
| --- | --- |
| Stok pusat rumah | GDG pada Data Usaha; PRODUKSI masuk GDG |
| Konsinyasi | KIRIM hanya memindahkan stok; invoice menautkan JUAL |
| PO minimal 10 pcs | Transaksi K menjumlah nomor PO + outlet termasuk beberapa lot; Beranda I57 menunjukkan baris bermasalah |
| Sisa terbawa bulan berikutnya | Saldo kumulatif sampai cutoff, tanpa reset AWAL |
| Retur kerusakan fisik kemasan | Alasan wajib pada RETUR; kebijakan di form/panduan |
| 8 area + independen | Wilayah sumber; kelompok Amanda/Independen per outlet |
| Ongkir dan diskon area | Ongkir aktual di jurnal; ongkir ditagih/diskon pada invoice wilayah |
| HPP referensi | Enam komponen, total Rp9.410; bukan harga jual |
| Invoice PDF | Layout cetak dan VBA v2 TERBITKAN + PDF / ekspor ulang; pemasangan .xlsm diperlukan |
| Tinggi/sedang/datar | Grafik batang A20 berwarna sesuai kategori, label teks, ringkasan jumlah area, tabel A36:H44, ambang relatif B276 pada Data Usaha |

## Revisi visualisasi wilayah — 7 November 2026

MoM Updated G6 meminta “grafik area yang memetakan tingkat penjualan setiap kota/wilayah” dan kategori tinggi/sedang/datar. Grafik sebelumnya di A47 kurang terlihat. Beranda kini menempatkan grafik lebih dahulu, mulai A20/G20, sebelum tabel rincian. Grafik batang horizontal memudahkan perbandingan delapan nama wilayah yang panjang. Grafik area kedua menunjukkan urutan waktu pada satu wilayah yang dipilih, sehingga tidak menumpuk delapan kurva.

Perbandingan mengikuti B5–E5; tren memakai 12 bulan sampai E5 dan pilihan wilayah I18. Label rentang ditampilkan untuk menjelaskan perbedaan itu. Kategori menggunakan teks selain warna; wilayah tanpa penjualan tetap bernama, berstatus DATAR dan bernilai nol. Angka penjualan pada batang dan tabel memungkinkan pembacaan tanpa mengandalkan warna. Tidak ada data sintetis di file operasional; contoh grafik diberi label SIMULASI.

Prinsip sumber NN/G dan Microsoft di atas diterapkan pada prioritas informasi, pilihan wilayah yang terlihat, label sumbu/satuan, dan penjelasan keadaan kosong. Grafik bukan ukuran laba, penerimaan kas, pencapaian target, atau pertumbuhan otomatis. Ambang 50% tetap pilihan konfigurasi, bukan angka yang ditentukan MoM.

## Verifikasi dan uji penerimaan

Rumus dihitung ulang pada simulasi terpisah: stok 40−30+40=50, cutoff, pemisahan area, diskon/ongkir/cicilan, retur/rusak, snapshot berubah, VOID, outlet nonaktif, PO kecil/beberapa lot, serta kategori performa. Pemeriksaan grafik mencakup delapan wilayah, 12 bulan, penggantian wilayah, pemisahan periode perbandingan/tren, cutoff tanggal, VOID, serta perubahan warna kategori setelah ambang diubah. Diperiksa pula jumlah sheet, tabel, dropdown/nama range, dan ketiadaan transaksi fiktif.

Pemeriksaan render berfokus pada navigasi, header, angka, grafik, input dan cetak invoice. Kompilasi/tombol/PDF VBA tetap membutuhkan Excel desktop; Linux tidak menjalankan Microsoft Excel.

Uji berikutnya bersama Bu Santi: isi outlet/harga/stok awal; kirim 40, jual 30, kirim 40; temukan saldo 50; terbitkan invoice dan catat cicilan. Amati apakah tautan pembayaran dan grup (+) ditemukan serta perbedaan kirim/jual dan dua jenis ongkir dipahami. Ukur waktu dan kesalahan pada sesi tersebut; hasil pengguna belum tersedia.
