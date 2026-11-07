"""One source for the Indonesian PDF and Markdown training walkthrough."""
from __future__ import annotations

from html import escape
from pathlib import Path
import re

import fitz
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether

GUIDE = "PANDUAN_LATIHAN_GIAN_FATAYA"
NAVY, TEAL, PALE = "#123044", "#087F8C", "#F0F5F7"


def money(value):
    return f"{value:,.0f}".replace(",", ".")


def date(value):
    year, month, day = value.split("-")
    return f"{day}/{month}/{year}"


def content(data):
    pages = []

    def page(title, subtitle):
        result = [("title", title), ("subtitle", subtitle)]
        pages.append(result)
        return result

    def p(blocks, value): blocks.append(("p", value))
    def note(blocks, value): blocks.append(("note", value))
    def h(blocks, value): blocks.append(("h", value))
    def table(blocks, headers, rows, widths=None): blocks.append(("table", headers, rows, widths))

    b = page("Mulai di sini", "Gian Fataya • latihan input sampai laporan akhir • versi 7 November 2026")
    note(b, "**DATA DUMMY.** Outlet, PIC, harga jual, lot, transaksi dan pembayaran pada contoh adalah fiktif. Nama delapan wilayah mengikuti MoM. Latihan memakai tabel Excel dan cetak PDF manual, tanpa pemasangan VBA.")
    h(b, "Pilih cara menggunakan paket")
    table(b, ["Tujuan", "File dan tindakan"], [
        ["Melihat hasil akhir", "Buka POS_Gian_Fataya_DUMMY.xlsx. Semua contoh sudah diisi. Jangan ketik ulang transaksi yang sama."],
        ["Mencoba dari nol", "Buka template_kosong/POS_Gian_Fataya.xlsx → File → Save As → LATIHAN_SAYA.xlsx. Ikuti halaman 2–10 pada salinan ini."],
        ["Melihat hasil PDF", "Buka CONTOH_Invoice_DUMMY_Jatim.pdf. Ini contoh cetakan invoice, bukan tagihan nyata."]
    ], [0.26, 0.74])
    h(b, "Urutan kerja")
    p(b, "**DATA USAHA → BERANDA (periode) → TRANSAKSI → STOK → TAGIHAN → PEMBAYARAN di TAGIHAN → CETAK INVOICE → BERANDA (hasil).**")
    p(b, "Setiap sheet memiliki navigasi baris 4. Untuk ke sel tertentu: klik kotak alamat di kiri bilah rumus, ketik alamat seperti A103, lalu Enter. Untuk berpindah ke pembayaran gunakan tautan pada TAGIHAN baris 5.")
    h(b, "Hasil yang akan Anda cocokkan")
    table(b, ["Periode 01/10–07/11/2026", "Hasil akhir"], [
        ["Penjualan / total tagihan", "Rp590.000 / Rp580.000"],
        ["Pembayaran masuk / piutang", "Rp490.000 / Rp90.000"],
        ["Stok gudang / semua outlet", "60 / 81 pouch"],
        ["Isi contoh", "4 outlet, 11 transaksi, 4 invoice, 4 pembayaran"]
    ], [0.56, 0.44])
    note(b, "**Saat mengetik:** masukkan nominal tanpa pemisah ribuan, misalnya 10000. Tanggal panduan memakai DD/MM/YYYY; masukkan sebagai tanggal Excel, bukan teks. Diskon 10% diketik 10%, bukan 10. Tampilan 10,000 dan 10.000 bisa berbeda menurut pengaturan regional.")

    b = page("1. Lengkapi DATA USAHA", "Mulai dari master: lokasi dan lot harus tersedia sebelum transaksi dibuat.")
    p(b, "Buka **DATA USAHA**. Sel biru untuk input. Kode wilayah A7:A14 dan namanya sudah tersedia; lengkapi PJ/alamat pada tiga wilayah yang dipakai. Kontak boleh kosong pada latihan.")
    table(b, ["Baris", "A/B sudah ada", "C · PIC Tagihan", "D · Alamat Invoice"],
          [[r, code + " / " + name, pic, address] for r, code, name, pic, address in data["regions"]], [0.08, 0.28, 0.30, 0.34])
    h(b, "Outlet — mulai A36, pertahankan GDG di baris 35")
    table(b, ["Baris", "A · Kode", "B · Nama", "C · Wilayah", "I · Kelompok"],
          [[r, code, name, region, group] for r, code, name, region, group in data["outlets"]], [0.08, 0.17, 0.36, 0.17, 0.22])
    p(b, "Untuk setiap outlet isi **D = PIC DUMMY**, **E = Alamat outlet DUMMY**, **G = AKTIF**. F/kontak kosong. **H otomatis** dan harus OK. Gudang GDG tetap berwilayah PUSAT; untuk salinan latihan Anda boleh mengganti nama/PIC gudang menjadi DUMMY.")
    h(b, "Produk, lot dan pengaturan")
    table(b, ["Alamat", "Isi"], [
        ["A103 / B103 / C103", "GF-PST-001 / Pastel DUMMY (pouch) / pouch"],
        ["D103 / F103", "10000 / AKTIF. Harga Rp10.000 hanya untuk latihan."],
        ["A161 / B161 / D161", "LOT-DUMMY / GF-PST-001 / 31/01/2027"],
        ["B268", "Pemilik DUMMY"],
        ["B270", "SIMULASI — tidak ada rekening atau permintaan transfer nyata"],
        ["B272", "Alamat pengirim DUMMY — latihan saja"],
        ["B274 / B276", "30 hari / 50% (ambang kategori TINGGI)"]
    ], [0.30, 0.70])
    note(b, "**Cek sebelum lanjut:** G7:G9 pada wilayah, H36:H39 pada outlet, H103 pada produk dan F161 pada lot harus OK. C161/nama produk serta kolom HPP acuan otomatis; jangan ditimpa. HPP bawaan bukan harga jual.")

    b = page("2–3. Periode, stok awal dan distribusi", "Catat stok fisik satu kali, kemudian catat kiriman sesuai PO.")
    h(b, "2. Atur periode di BERANDA")
    p(b, "Isi **B5 = 01/10/2026**, **E5 = 07/11/2026**. Periode ini dipakai untuk hasil gabungan latihan. Biarkan **I18 = Jawa Timur** untuk grafik tren. Tanggal akhir menentukan stok/piutang sampai tanggal tersebut.")
    h(b, "3. Buka TRANSAKSI dan isi baris berikut")
    p(b, "Semua baris menggunakan **E / Kode Lot = LOT-DUMMY**. Kolom G/harga jual dibiarkan kosong untuk AWAL dan KIRIM. **O / Status = AKTIF** sudah menjadi default. J:L otomatis; jangan diisi manual.")
    table(b, ["Baris", "A · ID", "B · Tgl", "C · Jenis", "D · Lokasi", "F · Qty", "H · PO", "I · Ongkir"],
          [[m["row"], m["id"], date(m["date"]), m["kind"], m["location"], m["qty"], m.get("po", "kosong"), money(m["freight"]) if "freight" in m else "kosong"] for m in data["movements"][:5]],
          [0.08, 0.10, 0.14, 0.11, 0.14, 0.07, 0.14, 0.12])
    p(b, "Isi P/catatan dengan DUMMY bila ingin menandai latihan. N/ekspedisi opsional, misalnya Ekspedisi DUMMY. Klik **+** pada grup kolom M:X bila perlu membuka status/catatan, kemudian tutup kembali agar input ringkas.")
    h(b, "Apa hasilnya sampai tahap ini?")
    table(b, ["Pemeriksaan", "Angka yang harus terlihat"], [
        ["BERANDA!E13 / I13", "Stok gudang 100 / stok outlet 100 pouch"],
        ["BERANDA!A8 / E8 / I8", "Penjualan 0 / tagihan 0 / pembayaran 0"],
        ["TRANSAKSI!J7:J11", "OK semua; KIRIM bukan penjualan"],
        ["Kolom K untuk empat PO", "OK; setiap PO minimal 10 pouch"]
    ], [0.43, 0.57])
    note(b, "**STOK adalah hasil otomatis.** Jangan mengetik saldo di sana. Pada STOK buka filter kolom D/Lot dan pilih LOT-DUMMY untuk melihat lima lokasi. File DUMMY sudah disaring ke lot itu; Data → Clear membuka kembali filter.")

    b = page("4. Catat penjualan dan stok bulan berikutnya", "Hanya JUAL menambah penjualan; stok yang tersisa dibawa terus.")
    h(b, "Penjualan Oktober — TRANSAKSI baris 12–15")
    p(b, "Isi C = **JUAL**, E = **LOT-DUMMY**, G/harga jual = **10000**, O = **AKTIF** pada empat baris ini. H/PO dan I/ongkir kosong. **M/ID invoice masih kosong**; pengaitan dilakukan setelah invoice dibuat pada langkah 5.")
    table(b, ["Baris", "A · ID", "B · Tanggal", "D · Lokasi", "F · Qty", "G · Harga"],
          [[m["row"], m["id"], date(m["date"]), m["location"], m["qty"], "10.000"] for m in data["movements"][5:9]],
          [0.07, 0.13, 0.23, 0.22, 0.10, 0.15])
    note(b, "**Cek setelah empat penjualan:** terjual 55 pouch, BERANDA!A8 = Rp550.000; gudang 100, outlet 45 pouch. Tagihan/kas masih 0. Bruto belum ditagih (K62) = Rp550.000; ini wajar sampai langkah invoice selesai.")
    h(b, "Lanjut November — tambah dua baris")
    table(b, ["Sel / baris", "Isi"], [
        ["A16:F16", "D-M10 / 01/11/2026 / KIRIM / D-JT01 / LOT-DUMMY / 40"],
        ["G16 / H16 / I16", "kosong / PO-D-005 / 15000"],
        ["A17:F17", "D-M11 / 05/11/2026 / JUAL / D-JT02 / LOT-DUMMY / 4"],
        ["G17 / H17 / I17", "10000 / kosong / kosong"],
        ["O16:O17 / M17", "AKTIF / ID invoice M17 masih kosong dulu"]
    ], [0.29, 0.71])
    p(b, "**D-JT01: 40 dikirim − 30 terjual + 40 kiriman baru = 50 pouch.** Tidak perlu mencatat AWAL lagi pada 1 November. D-JT02 tersisa 20 − 10 − 4 = 6 pouch.")
    p(b, "Setelah 11 transaksi: BERANDA!A8 = **Rp590.000**; stok gudang **60**, outlet **81**. Penjualan belum ditagih menjadi Rp590.000 hingga seluruh JUAL dikaitkan ke invoice. Semua J7:J17 harus OK dan L/saldo asal tidak negatif.")

    b = page("5. Buat invoice per wilayah", "Satu invoice menghimpun penjualan beberapa outlet dalam wilayah/periode yang sama.")
    p(b, "Buka **TAGIHAN**, tabel pertama mulai baris 7. Isi A:F berikut. Pada keempat baris isi **I/status = DRAF** dan R/catatan = DUMMY. Periode invoice November hanya sampai 05/11, sesuai penjualan yang ditagih.")
    table(b, ["Baris", "A · ID Invoice", "B · Terbit", "C · Area", "D–E · Periode", "F · Tempo"],
          [[i["row"], i["id"], date(i["date"]), i["region"], date(i["start"]) + " – " + date(i["end"]), date(i["due"])] for i in data["invoices"]],
          [0.08, 0.23, 0.17, 0.08, 0.27, 0.17])
    table(b, ["Baris", "G · Diskon", "H · Ongkir ditagih", "Q yang diharapkan = J/snapshot"],
          [[i["row"], f'{i["discount"]:.0%}', money(i["freight"]), money(i["gross"])] for i in data["invoices"]],
          [0.08, 0.18, 0.29, 0.45])
    h(b, "Lakukan berurutan untuk setiap invoice")
    p(b, "**1.** Setelah DRAF dibuat, buka TRANSAKSI → klik **+** untuk M:X. Isi **M12 dan M13 = INV-D-JT-1026**, **M14 = INV-D-JTG-1026**, **M15 = INV-D-MA-1026**, **M17 = INV-D-JT-1126**. M pada AWAL/KIRIM tetap kosong.")
    p(b, "**2.** Kembali ke TAGIHAN. Q/Bruto Sumber menghitung penjualan tertaut. Salin **Q7 → J7** dengan **Copy → Paste Special → Values / Tempel Nilai**. Ulangi Q8→J8, Q9→J9, Q10→J10. Jangan menaruh rumus =Q7 di J7.")
    p(b, "**3.** Pastikan P7:P10 = **OK** dan J cocok dengan Q. Nilai J boleh kosong/menunjukkan CEK NOMINAL selama DRAF belum selesai; jangan diterbitkan sebelum cocok.")
    p(b, "**4.** Ubah I7:I10 dari DRAF menjadi **TERBIT**. Jaga angka snapshot tetap. Invoice DRAF belum dihitung sebagai tagihan terbit dan belum siap dicetak.")
    note(b, "**Hasil sebelum pembayaran:** total tagihan Rp580.000; piutang Rp580.000. Invoice Jatim Oktober = 400.000 − (10% × 400.000) + 20.000 = **Rp380.000**, gabungan 30 + 10 pouch dari dua outlet. Ongkir aktual di TRANSAKSI berbeda dari ongkir yang diputuskan untuk ditagih pada invoice.")

    b = page("6. Catat pembayaran dan cicilan", "Pembayaran mengurangi piutang, tanpa mengubah jumlah produk terjual.")
    p(b, "Pada **TAGIHAN** klik **Pembayaran / cicilan** di baris 5, atau gunakan kotak alamat **A437**. Header tabel di baris 436; isi pembayaran mulai baris 437. Pada semua baris D/metode = **Transfer**, F/status = **AKTIF**, G/catatan = **DUMMY**. H:J otomatis.")
    table(b, ["Baris", "A · ID", "B · Tanggal", "C · ID Invoice", "E · Nominal"],
          [[p["row"], p["id"], date(p["date"]), p["invoice"], money(p["amount"])] for p in data["payments"]],
          [0.08, 0.14, 0.23, 0.34, 0.21])
    p(b, "Catat sesuai tanggal uang benar-benar diterima, walaupun data diinput belakangan. Dua pembayaran boleh menunjuk invoice yang sama; ID pembayarannya harus berbeda. I437:I440 harus OK.")
    h(b, "Kembali ke tabel invoice: lihat M, N dan O")
    table(b, ["Invoice", "L · Total", "M · Dibayar", "N · Piutang", "O · Status"], [
        ["INV-D-JT-1026", "380.000", "380.000", "0", "LUNAS"],
        ["INV-D-JTG-1026", "110.000", "60.000", "50.000", "SEBAGIAN"],
        ["INV-D-MA-1026", "50.000", "50.000", "0", "LUNAS"],
        ["INV-D-JT-1126", "40.000", "0", "40.000", "BELUM BAYAR"]
    ], [0.29, 0.17, 0.17, 0.17, 0.20])
    note(b, "**Hasil akhir pembayaran:** Rp100.000 + Rp280.000 + Rp60.000 + Rp50.000 = **Rp490.000**. Sisa piutang = Rp580.000 − Rp490.000 = **Rp90.000**. Jangan mengisi M/N/status bayar manual; semuanya hasil rumus.")
    h(b, "Mengapa total tagihan berbeda dari bruto penjualan?")
    p(b, "Bruto Rp590.000 dikurangi diskon Rp40.000 lalu ditambah ongkir ditagih Rp30.000 menghasilkan tagihan Rp580.000. Pembayaran masuk Rp490.000 adalah penerimaan pelanggan, bukan laba atau saldo kas bersih setelah seluruh biaya usaha.")

    b = page("7. Lihat dan ekspor invoice PDF", "Hasil cetak terpisah per wilayah; contoh ini memakai Jatim Oktober.")
    p(b, "**1.** Buka **CETAK INVOICE**. **2.** Pilih **B5 = INV-D-JT-1026**. **3.** A14 harus **SIAP CETAK**; penerima adalah PJ DUMMY Jatim dan alamat dummy, bukan setiap outlet.")
    p(b, "**4.** Cocokkan qty **40**, bruto **Rp400.000**, diskon **Rp40.000**, ongkir **Rp20.000**, total **Rp380.000**, dibayar **Rp380.000**, sisa **0**. File DUMMY sudah menyembunyikan baris produk kosong 18–66; saat latihan dari template kosong, pilih baris 18–66 → klik kanan → Hide agar cetakan ringkas.")
    p(b, "**5.** Excel: **File → Print**, pilih **Print Active Sheets**, periksa satu halaman A4 portrait, lalu **Microsoft Print to PDF / Save as PDF**. Alternatif: File → Export → Create PDF, pilih Active sheet(s). LibreOffice: File → Export As → PDF, gunakan area cetak sheet invoice. **6.** Simpan sebagai **Invoice_DUMMY_Jatim.pdf**.")
    b.append(("image", "gambar/03_Invoice_Dummy.png", 450, "Contoh cetakan aktual dari workbook; label DUMMY tetap terlihat."))
    note(b, "Ganti B5 untuk invoice lain. Contoh Jateng: total Rp110.000, dibayar Rp60.000, sisa Rp50.000. Jangan mengubah isi cetakan langsung; perbaiki data sumber jika ada kesalahan.")

    b = page("8. Baca dashboard dan grafik wilayah", "Kembali ke BERANDA dengan periode 01/10–07/11/2026.")
    table(b, ["Sel", "Angka yang diharapkan", "Artinya"], [
        ["A8", "Rp590.000", "Bruto seluruh JUAL valid pada periode"],
        ["E8", "Rp580.000", "Invoice TERBIT dalam periode"],
        ["I8", "Rp490.000", "Pembayaran masuk dalam periode"],
        ["A13", "Rp90.000", "Piutang sampai tanggal akhir"],
        ["E13 / I13", "60 / 81 pouch", "Stok gudang / semua outlet sampai tanggal akhir"]
    ], [0.14, 0.30, 0.56])
    b.append(("image", "gambar/02_Grafik_Dummy.png", 510, "Grafik Excel asli pada data latihan: perbandingan delapan wilayah dan tren wilayah pilihan."))
    p(b, "**Grafik kiri:** Jawa Timur Rp440.000 (**TINGGI**), Jawa Tengah Rp100.000 dan Medan–Aceh Rp50.000 (**SEDANG**). Lima wilayah lain **DATAR (0)**. Ringkasan: 1 tinggi, 2 sedang, 5 datar. Toko independen Semarang ikut Jawa Tengah.")
    p(b, "**Grafik kanan:** pilih wilayah pada **I18**. Untuk Jawa Timur, Oktober Rp400.000 dan November sampai 7 November Rp40.000. Rentang selalu 12 bulan sampai E5; bulan sebelum ada penjualan bernilai 0. B5 mengubah perbandingan kiri, bukan panjang rentang tren.")
    note(b, "TINGGI = bruto minimal 50% area teratas; SEDANG = positif di bawah ambang; DATAR = nol. Ambang dapat diubah di DATA USAHA!B276. Ini pembanding antarwilayah, bukan ukuran pertumbuhan atau target laba. Data November belum satu bulan penuh.")

    b = page("9. Coba filter periode dan wilayah", "Ubah hanya sel biru pada salinan latihan, lalu pulihkan ke periode awal.")
    p(b, "Ganti **BERANDA!B5/E5** sesuai tabel. Angka di bawah adalah hasil hitung ulang dari transaksi yang sama, tanpa menambah/menghapus jurnal.")
    table(b, ["B5–E5 (2026)", "Bruto", "Tagihan", "Kas masuk", "Piutang", "Gudang", "Outlet"], [
        ["01/10–07/11", "590.000", "580.000", "490.000", "90.000", "60", "81"],
        ["01/10–31/10", "550.000", "540.000", "100.000", "440.000", "100", "45"],
        ["01/11–07/11", "40.000", "40.000", "390.000", "90.000", "60", "81"]
    ], [0.20, 0.15, 0.15, 0.15, 0.15, 0.10, 0.10])
    note(b, "Bruto, tagihan terbit dan penerimaan memakai rentang B5–E5. Stok dan piutang memakai **akumulasi sampai E5**. Karena itu piutang November tetap mencakup invoice Oktober yang belum lunas; pembayaran November tidak mengurangi piutang historis pada 31 Oktober.")
    p(b, "**TAGIHAN** menampilkan status bayar terkini dari seluruh pembayaran yang sudah dicatat. Status tersebut tidak ikut menjadi historis ketika filter Beranda diganti. Untuk piutang pada tanggal tertentu, baca BERANDA!A13.")
    h(b, "Ganti pilihan wilayah untuk tren")
    table(b, ["Pilihan I18", "Oktober", "November s.d. 07/11"], [
        ["Jawa Timur", "400.000", "40.000"],
        ["Jawa Tengah", "100.000", "0"],
        ["Medan - Aceh", "50.000", "0"]
    ], [0.42, 0.29, 0.29])
    p(b, "Pulihkan **B5 = 01/10/2026**, **E5 = 07/11/2026**, **I18 = Jawa Timur** setelah mencoba. Ctrl+S untuk menyimpan.")
    h(b, "Bulan berikutnya mulai dari mana?")
    p(b, "Gunakan file kerja yang sama dan lanjutkan di baris TRANSAKSI berikutnya (**18** pada contoh). Catat PO/penjualan baru; jangan ulangi AWAL atau menyalin saldo sebagai transaksi baru. Invoice berikutnya di TAGIHAN baris **11**, pembayaran berikutnya di baris **441**. Beri ID baru yang unik, lalu ubah periode laporan sesuai kebutuhan.")
    p(b, "Sebelum mengoreksi data periode lama, simpan arsip bulanan. Untuk penggunaan usaha nyata, mulai dari template kosong dan isi data aktual milik mitra; jangan memakai stok/harga/outlet DUMMY sebagai saldo pembukaan usaha.")

    b = page("10. Cocokkan hasil dan selesaikan latihan", "Gunakan daftar ini untuk memastikan seluruh alur sudah berjalan.")
    h(b, "Stok akhir per lokasi — STOK, filter LOT-DUMMY")
    table(b, ["Lokasi", "Perhitungan", "Saldo akhir"], [
        ["GDG", "200 − (40 + 20 + 20 + 20 + 40)", "60"],
        ["D-JT01", "40 − 30 + 40", "50"],
        ["D-JT02", "20 − 10 − 4", "6"],
        ["D-JTG01", "20 − 10", "10"],
        ["D-MA01", "20 − 5", "15"],
        ["TOTAL", "200 awal − 59 terjual", "141 = 60 gudang + 81 outlet"]
    ], [0.22, 0.45, 0.33])
    p(b, "**Pemeriksaan akhir:** J7:J17 di TRANSAKSI, P7:P10 di TAGIHAN, I437:I440 pada pembayaran semuanya OK. BERANDA!I57/C60/G60/K60/C62/G62/K62 semuanya 0. Sisa piutang Rp90.000 adalah saldo sah dari dua invoice, bukan kesalahan input.")
    h(b, "Jika hasil belum cocok")
    table(b, ["Gejala", "Periksa dan perbaiki"], [
        ["Cannot open the specified file", "Gunakan file versi perbaikan. Untuk satu tombol pada salinan lama: klik kanan → Edit Hyperlink → Place in This Document → sheet tujuan, misalnya STOK!A1."],
        ["J bukan OK / angka tidak masuk", "Cek kode outlet dan lot di master, tanggal valid, qty positif, status AKTIF; harga wajib hanya untuk JUAL, nomor PO wajib untuk KIRIM."],
        ["PO < 10 PCS / saldo negatif", "Lengkapi PO minimal 10 pcs per nomor PO + outlet; cek stok sebelum kirim/jual. Jangan menimpa rumus atau saldo STOK."],
        ["CEK NOMINAL / SELISIH SUMBER", "Cek diskon (10%, bukan 10), ongkir 0 bila nihil, tautan ID invoice di TRANSAKSI M, serta snapshot J sama dengan sumber Q. Jangan mengubah penjualan yang sudah ditagih."],
        ["PDF belum SIAP CETAK", "Invoice harus TERBIT dan P = OK; lengkapi PIC/alamat wilayah. Pilih ID yang benar di CETAK INVOICE!B5."],
        ["Pembayaran tidak dihitung", "Pilih invoice TERBIT, tanggal tidak sebelum tanggal terbit, nominal positif, metode Transfer, status AKTIF; cek kolom I pembayaran."],
        ["Grafik tampak nol / angka tidak berubah", "Cek B5/E5 dan I18; pilih Formulas → Calculation Options → Automatic, lalu F9. Jangan mengubah sumber/rumus grafik."]
    ], [0.29, 0.71])
    note(b, "**Hasil yang dibawa pulang:** Excel berisi jurnal, stok berjalan, empat invoice beserta status bayar, dua grafik wilayah, serta contoh PDF invoice. Simpan LATIHAN_SAYA.xlsx dan PDF hasil latihan. Rumus, tujuh tahap dan render contoh diuji dengan LibreOffice; Microsoft Excel desktop dan VBA belum dieksekusi di lingkungan ini.")
    return pages


def inline(text):
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escape(str(text))).replace("\n", "<br/>")


def build(dest: Path, data):
    pages = content(data)
    # Embedded local fonts keep tables and Indonesian punctuation consistent.
    folder = Path("/usr/share/fonts/truetype/liberation")
    for name, file in [("Guide", "LiberationSans-Regular.ttf"), ("GuideBold", "LiberationSans-Bold.ttf"), ("GuideItalic", "LiberationSans-Italic.ttf")]:
        pdfmetrics.registerFont(TTFont(name, str(folder / file)))
    pdfmetrics.registerFontFamily("Guide", normal="Guide", bold="GuideBold", italic="GuideItalic", boldItalic="GuideBold")
    base = ParagraphStyle("body", fontName="Guide", fontSize=10.1, leading=13.8, textColor=colors.HexColor(NAVY), spaceAfter=8)
    styles = {
        "p": base,
        "title": ParagraphStyle("title", parent=base, fontName="GuideBold", fontSize=21, leading=25, spaceAfter=9),
        "subtitle": ParagraphStyle("subtitle", parent=base, fontSize=10, leading=14, textColor=colors.HexColor("#526674"), spaceAfter=16),
        "h": ParagraphStyle("section", parent=base, fontName="GuideBold", fontSize=12.2, leading=16, spaceBefore=5, spaceAfter=8),
        "cell": ParagraphStyle("cell", parent=base, fontSize=8.6, leading=11.3, spaceAfter=0),
        "th": ParagraphStyle("th", parent=base, fontName="GuideBold", fontSize=8.6, leading=11.3, textColor=colors.white, spaceAfter=0),
        "caption": ParagraphStyle("caption", parent=base, fontName="GuideItalic", fontSize=8.5, leading=11, textColor=colors.HexColor("#526674")),
    }
    width = A4[0] - 80
    story, markdown = [], ["# Gian Fataya — panduan latihan data DUMMY", ""]
    for number, blocks in enumerate(pages):
        if number:
            story.append(PageBreak())
        for block in blocks:
            kind = block[0]
            if kind in {"title", "subtitle", "p", "h"}:
                story.append(Paragraph(inline(block[1]), styles[kind]))
                prefix = "## " if kind == "title" else "### " if kind == "h" else ""
                markdown.extend([prefix + block[1], ""])
            elif kind == "note":
                box = Table([[Paragraph(inline(block[1]), base)]], colWidths=[width])
                box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(PALE)),
                    ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9E5EA")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                    ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
                story.extend([KeepTogether(box), Spacer(1, 9)])
                markdown.extend(["> " + block[1], ""])
            elif kind == "table":
                _, headers, rows, widths = block
                table_data = [[Paragraph(inline(v), styles["th"]) for v in headers]]
                table_data += [[Paragraph(inline(v), styles["cell"]) for v in row] for row in rows]
                col_widths = [width * x for x in widths] if widths else [width / len(headers)] * len(headers)
                table = Table(table_data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
                table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TEAL)),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor(PALE), colors.white]),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor(TEAL)),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
                story.extend([table, Spacer(1, 10)])
                markdown.append("| " + " | ".join(headers) + " |")
                markdown.append("| " + " | ".join(["---"] * len(headers)) + " |")
                markdown.extend("| " + " | ".join(str(v) for v in row) + " |" for row in rows)
                markdown.append("")
            elif kind == "image":
                _, filename, size, caption = block
                image = Image(str(dest / filename))
                factor = min(size, width) / image.imageWidth
                image.drawWidth, image.drawHeight = image.imageWidth * factor, image.imageHeight * factor
                image.hAlign = "CENTER"
                story.extend([image, Spacer(1, 4), Paragraph(inline(caption), styles["caption"])])
                markdown.extend([f"![{caption}]({filename})", ""])
    (dest / (GUIDE + ".md")).write_text("\n".join(markdown), encoding="utf-8")

    def furniture(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor(NAVY))
        canvas.rect(0, A4[1] - 23, A4[0], 23, fill=1, stroke=0)
        canvas.setFont("GuideBold", 8)
        canvas.setFillColor(colors.white)
        canvas.drawString(40, A4[1] - 15, "GIAN FATAYA  /  PANDUAN PRAKTIK")
        canvas.setFillColor(colors.HexColor("#526674"))
        canvas.setFont("Guide", 8)
        canvas.drawString(40, 25, "DUMMY • untuk latihan, bukan tagihan nyata")
        canvas.drawRightString(A4[0] - 40, 25, f"Halaman {doc.page}")
        canvas.restoreState()

    path = dest / (GUIDE + ".pdf")
    document = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=40, rightMargin=40, topMargin=42, bottomMargin=44,
                                 title="Gian Fataya — panduan latihan data DUMMY", author="Panduan latihan Gian Fataya")
    document.build(story, onFirstPage=furniture, onLaterPages=furniture)
    with fitz.open(path) as pdf:
        text = "\n".join(page.get_text() for page in pdf)
        for expected in ["Mulai di sini", "D-JTG01", "INV-D-JT-1026", "590.000", "90.000", "Tempel Nilai", "10. Cocokkan"]:
            assert expected in text, expected
        assert len(pdf) == len(pages), f"Guide overflow: {len(pdf)} pages, expected {len(pages)}"
        for i, page in enumerate(pdf):
            for block in page.get_text("blocks"):
                assert block[0] >= 35 and block[2] <= A4[0] - 34, f"Text clipped horizontally on page {i+1}"
        print(f"Guide verified: {len(pdf)} pages; PDF/Markdown share the same instructions.", flush=True)
