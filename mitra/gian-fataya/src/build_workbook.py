"""Build Gian Fataya's consignment workbook without inheriting restaurant VBA.

Run from any directory: python3 src/build_workbook.py
All formulas use Excel 2016-compatible functions. No actual sales are seeded.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col

ROOT = Path(__file__).resolve().parents[1]
FIRST = 7
CAP = {"WILAYAH": 20, "OUTLET": 61, "PRODUK": 50, "LOT": 80,
       "MUTASI": 2000, "INVOICE": 400, "PEMBAYARAN": 1000}
TYPES = ["AWAL", "PRODUKSI", "KIRIM", "JUAL", "RETUR", "RUSAK", "TAMBAH", "KURANG"]
NAV = ["MULAI", "DASHBOARD", "FORM", "STOK", "INVOICE", "CETAK INVOICE"]


def rng(sheet, column):
    return f"'{sheet}'!${column}${FIRST}:${column}${FIRST + CAP[sheet] - 1}"


def lookup(cell, sheet, key, result, default='""'):
    return f'IFERROR(INDEX({rng(sheet,result)},MATCH({cell},{rng(sheet,key)},0)),{default})'


def build(dest: Path):
    profile = json.loads((ROOT / "data/profile.json").read_text())
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(dest))
    wb.set_properties({"title": "Gian Fataya | Konsinyasi & Invoice", "author": "Social Project BSI",
                       "subject": "Stok, distribusi, penjualan dan invoice per wilayah",
                       "comments": "Data sumber: MoM 22 Oktober 2026 dan dua foto referensi. Tidak berisi transaksi asli."})
    wb.set_calc_mode("auto")
    navy, teal, gold = "#123044", "#087F8C", "#DFA94A"
    ink, muted, pale, blue = "#203B4A", "#657986", "#F0F5F7", "#145DA0"
    fmt = {}
    def f(name, props):
        fmt[name] = wb.add_format({"font_name": "Calibri", "font_size": 11, "font_color": ink, **props})
    f("title", {"bold": True, "font_size": 23, "bg_color": navy, "font_color": "white", "valign": "vcenter"})
    f("sub", {"font_size": 10, "font_color": muted, "text_wrap": True, "valign": "vcenter"})
    f("nav", {"font_size": 10, "bold": True, "font_color": teal, "underline": True})
    f("section", {"bold": True, "bg_color": teal, "font_color": "white", "valign": "vcenter"})
    f("input", {"font_color": blue, "bg_color": "#EDF5FC", "bottom": 1, "bottom_color": "#CFDFEC"})
    f("input_date", {"font_color": blue, "bg_color": "#EDF5FC", "num_format": "dd mmm yyyy", "bottom": 1, "bottom_color": "#CFDFEC"})
    f("input_money", {"font_color": blue, "bg_color": "#EDF5FC", "num_format": '#,##0;[Red](#,##0);"–"', "bottom": 1, "bottom_color": "#CFDFEC"})
    f("input_pct", {"font_color": blue, "bg_color": "#EDF5FC", "num_format": "0.0%", "bottom": 1, "bottom_color": "#CFDFEC"})
    f("calc", {"bg_color": "#F0F5F7"})
    f("money", {"bg_color": "#F0F5F7", "num_format": '#,##0;[Red](#,##0);"–"'})
    f("qty", {"bg_color": "#F0F5F7", "num_format": '#,##0;[Red](#,##0);"–"'})
    f("date", {"bg_color": "#F0F5F7", "num_format": "dd mmm yyyy"})
    f("pct", {"num_format": "0.0%", "bg_color": "#F0F5F7"})
    f("label", {"bold": True, "font_color": muted})
    f("body", {"text_wrap": True, "valign": "top"})
    f("note", {"font_size": 10, "font_color": muted, "text_wrap": True, "valign": "top"})
    f("warning", {"bg_color": "#FFF2D5", "font_color": "#80590D", "text_wrap": True, "valign": "vcenter"})
    f("error", {"bg_color": "#FDE9E7", "font_color": "#9D2E28"})
    f("good", {"bg_color": "#E0F1E8", "font_color": "#216A44"})
    f("kpi_label", {"bg_color": navy, "font_color": "#BBCED8", "font_size": 10})
    f("kpi", {"bg_color": navy, "font_color": "white", "bold": True, "font_size": 23, "num_format": '#,##0;[Red](#,##0);"0"'})
    f("header", {"bold": True, "bg_color": navy, "font_color": "white", "text_wrap": True, "valign": "vcenter"})
    f("print_title", {"bold": True, "font_size": 25, "font_color": navy})
    f("print_money", {"num_format": '#,##0;[Red](#,##0);"–"', "bottom": 1, "bottom_color": "#DFE7EB"})
    f("print_cell", {"font_size": 10, "bottom": 1, "bottom_color": "#DFE7EB"})
    sheets = {}
    def base(name, subtitle, width=12):
        s = wb.add_worksheet(name)
        sheets[name] = s
        s.hide_gridlines(2)
        s.set_default_row(21)
        s.set_tab_color(teal if name in NAV else (gold if name in ["WILAYAH", "OUTLET", "PRODUK", "LOT"] else navy))
        s.set_column(0, width-1, 16)
        s.set_column(0, 0, 19)
        s.merge_range(0, 0, 0, width-1, name.replace("MULAI", "GIAN FATAYA  /  KONSINYASI"), fmt["title"])
        s.set_row(0, 44)
        s.merge_range(1, 0, 1, width-1, subtitle, fmt["sub"])
        s.set_row(1, 32)
        for i, n in enumerate(NAV):
            s.write_url(3, i*2, f"internal:'{n}'!A1", fmt["nav"], n.title())
        s.freeze_panes(6, 2)
        s.set_landscape()
        s.set_paper(9)
        s.fit_to_pages(1, 0)
        s.set_margins(0.25, 0.25, 0.4, 0.4)
        s.set_header("&LGIAN FATAYA  |  Konsinyasi & Invoice")
        s.set_footer("&LInternal • data operasional mitra&RHalaman &P / &N")
        s.repeat_rows(0, 5)
        s.set_zoom(85)
        return s

    def table(s, name, headers, n, formulas=None, inputs=None, formats=None):
        formulas, inputs, formats = formulas or {}, inputs or set(), formats or {}
        columns = []
        for i,h in enumerate(headers):
            letter = col(i)
            columns.append({"header": h, "format": fmt[formats.get(letter, "input" if letter in inputs else "calc")]})
        s.add_table(5, 0, 5+n, len(headers)-1, {"name": name, "style": "Table Style Medium 2", "columns": columns})
        s.set_row(5, 35)
        for r in range(FIRST, FIRST+n):
            for letter, make in formulas.items():
                s.write_formula(f"{letter}{r}", "="+make(r), fmt[formats.get(letter,"calc")], "")
        s.print_area(0,0,min(5+n,60),len(headers)-1)

    def listv(s, cells, source):
        s.data_validation(cells, {"validate":"list", "source":source, "error_type":"stop", "error_title":"Pilih dari daftar", "error_message":"Gunakan kode atau pilihan yang tersedia.", "input_title":"Kolom input", "input_message":"Pilih dari dropdown; sel biru diisi pengguna."})

    def audit(s, column, end):
        s.set_column(f"{column}:{column}", 26)
        s.conditional_format(f"{column}{FIRST}:{column}{end}", {"type":"formula", "criteria":f'AND({column}{FIRST}<>"",{column}{FIRST}<>"OK")', "format":fmt["error"]})

    # Shared validation lists have explicit capacities; never depend on Excel dynamic arrays.
    for name, sheet, c in [("DaftarWilayah","WILAYAH","A"),("DaftarLokasi","OUTLET","A"),("DaftarSKU","PRODUK","A"),("DaftarLot","LOT","A"),("DaftarInvoice","INVOICE","A")]:
        wb.define_name(name, "="+rng(sheet,c))

    s = base("MULAI", "Pencatatan distribusi konsinyasi, stok per outlet, penjualan dan tagihan per wilayah. Versi 1.0 • MoM 22 Oktober 2026.")
    s.merge_range("A6:L6", "01  SIAPKAN DATA, LALU CATAT ALUR BARANG", fmt["section"])
    steps = [
        ("1. Wilayah & outlet", "Delapan wilayah bersumber dari foto mitra. Isi PIC tagihan, alamat, serta kode outlet sebenarnya di WILAYAH dan OUTLET."),
        ("2. Produk & lot", "Pastel per pouch sudah tersedia. Isi harga jual. Daftarkan batch/lot dan tanggal kedaluwarsa aktual; jangan menebak masa simpan."),
        ("3. Stok awal", "Catat AWAL di MUTASI untuk setiap lokasi dan lot sesuai hitung fisik pada tanggal mulai. Jangan ulangi AWAL setiap bulan."),
        ("4. Distribusi & penjualan", "KIRIM memindahkan stok gudang ke outlet, bukan pendapatan. JUAL hanya untuk barang yang sudah dilaporkan terjual."),
        ("5. Invoice per wilayah", "Buat ID unik di INVOICE (DRAF), tautkan ID ke baris JUAL terpilih di MUTASI, salin nilai bruto sumber sebagai nilai tetap, lalu TERBIT."),
        ("6. Pembayaran", "Catat pembayaran parsial/penuh beserta ID invoice di PEMBAYARAN. Buka CETAK INVOICE dan pilih ID untuk mencetak."),
    ]
    for row,(title,body) in zip([7,10,13,16,19,22],steps):
        s.merge_range(row,0,row,2,title,fmt["label"])
        s.merge_range(row,3,row+1,11,body,fmt["body"])
        s.set_row(row,25)
    s.merge_range("A27:L28", "Sel biru = input. Sel abu-abu = rumus, jangan ditimpa. Satu baris MUTASI = satu jenis gerakan, satu lokasi, satu lot. Baris tidak valid dikecualikan dari perhitungan; periksa kolom Cek dan saldo negatif sebelum menerbitkan invoice.", fmt["warning"])
    s.merge_range("A30:L30", "02  PEMERIKSAAN PERSIAPAN",fmt["section"])
    setup = [
        ("Harga jual produk",f'IF(COUNTIF({rng("PRODUK","D")},">0")=0,"BELUM DIISI","TERSEDIA; CEK SEMUA PRODUK")'),
        ("Outlet mitra",f'IF(COUNTA({rng("OUTLET","A")})<2,"BELUM DIISI","TERSEDIA")'),
        ("Lot & kedaluwarsa",f'IF(COUNTA({rng("LOT","A")})=0,"BELUM DIISI","CEK KOLOM VALIDASI LOT")'),
        ("HPP bahan acuan / pouch", "SUM(BAHAN!D7:D12)"),
    ]
    for row,(label,formula) in enumerate(setup,31):
        s.merge_range(row,0,row,3,label,fmt["label"])
        s.merge_range(row,4,row,9,"",fmt["calc"])
        s.write_formula(row,4,"="+formula,fmt["money"] if row==34 else fmt["calc"],9410 if row==34 else "BELUM DIISI")
    s.merge_range("A37:L39", "CATATAN: harga jual, outlet terperinci, stok awal, tanggal kedaluwarsa, PIC/alamat tagihan, rekening dan contoh invoice belum tersedia dari sumber. Tidak ada transaksi asli maupun nominal penjualan contoh di file utama. HPP Rp9.410 hanya bahan dan kemasan, bukan HPP lengkap.",fmt["warning"])
    s.merge_range("A41:L43", "AUTOMASI OPSIONAL: file ini langsung dapat digunakan dengan mengisi jurnal. Untuk tombol FORM dan ekspor PDF, simpan salinan sebagai .xlsm, impor Modul_GianFataya.bas melalui Alt+F11 > File > Import, lalu Alt+F8 > PasangTombol. Aktifkan makro hanya untuk file yang Anda percayai. Modul kasir restoran lama tidak disertakan.",fmt["note"])
    s.merge_range("A45:L47", "Kapasitas versi ini: 20 wilayah, 60 outlet + gudang, 50 produk, 80 lot, 2.000 mutasi, 400 invoice, 1.000 pembayaran. Jangan memasukkan data di luar tabel. Satu operator per file; file ini tidak menyediakan sinkronisasi antar laptop. Cadangkan file sebelum setiap periode penagihan.",fmt["note"])
    s.print_area("A1:L47")

    s=base("WILAYAH","Master area penagihan. Nama delapan wilayah mengikuti foto Outlet Area_GianFataya; cakupan dan PIC tetap perlu dikonfirmasi.")
    table(s,"tblWilayah",["Kode Wilayah","Nama Wilayah","PIC Tagihan","Alamat Invoice","Kontak","Catatan","Cek"],CAP["WILAYAH"],{
        "G":lambda r:f'IF(A{r}="","",IF(COUNTIF({rng("WILAYAH","A")},A{r})>1,"ID GANDA",IF(B{r}="","ISI NAMA","OK")))'},set("ABCDEF"))
    for i,name in enumerate(profile["regions"],FIRST-1):
        s.write(i,0,f"W{i-FIRST+2:02}",fmt["input"]);s.write(i,1,name,fmt["input"])
        s.write(i,5,"Sumber: foto mitra; PIC dan cakupan perlu konfirmasi",fmt["input"])
    s.set_column("B:D",28);s.set_column("F:F",52);audit(s,"G",26)

    s=base("OUTLET","Daftarkan outlet sebenarnya; jangan menggunakan nama wilayah sebagai pengganti outlet. GDG adalah stok pusat di rumah Bu Santi.")
    table(s,"tblOutlet",["Kode Lokasi","Nama Lokasi","Kode Wilayah","PIC Outlet","Alamat","Kontak","Status","Cek"],CAP["OUTLET"],{
        "H":lambda r:f'IF(A{r}="","",IF(COUNTIF({rng("OUTLET","A")},A{r})>1,"ID GANDA",IF(B{r}="","ISI NAMA",IF(AND(A{r}<>"GDG",COUNTIF(DaftarWilayah,C{r})<>1),"CEK WILAYAH",IF(AND(G{r}<>"AKTIF",G{r}<>"NONAKTIF"),"CEK STATUS","OK")))))'},set("ABCDEFG"))
    s.write_row("A7",["GDG","Gudang Rumah Bu Santi","PUSAT","Bu Santi","","","AKTIF"],fmt["input"])
    listv(s,"C8:C67","=DaftarWilayah");listv(s,"G7:G67",["AKTIF","NONAKTIF"])
    s.set_column("B:B",30);s.set_column("D:F",28);audit(s,"H",67)

    s=base("PRODUK","Harga jual belum diberikan. HPP bahan acuan tidak mencakup tenaga kerja, overhead atau ongkir. Kode SKU dibuat untuk keperluan sistem.")
    table(s,"tblProduk",["SKU","Nama Produk","Satuan","Harga Jual","HPP Bahan Acuan","Status","Catatan","Cek"],CAP["PRODUK"],{
        "H":lambda r:f'IF(A{r}="","",IF(COUNTIF(DaftarSKU,A{r})>1,"SKU GANDA",IF(OR(B{r}="",C{r}=""),"ISI NAMA/SATUAN",IF(OR(NOT(ISNUMBER(D{r})),D{r}<=0),"ISI HARGA JUAL",IF(AND(F{r}<>"AKTIF",F{r}<>"NONAKTIF"),"CEK STATUS","OK")))))'},set("ABCDEFG"),{"D":"input_money","E":"input_money"})
    s.write_row("A7",[profile["product"]["sku"],profile["product"]["name"],"pouch"],fmt["input"])
    s.write_formula("E7","=SUM(BAHAN!D7:D12)",fmt["money"],9410)
    s.write("F7","AKTIF",fmt["input"]);s.write("G7","Nama dinormalisasi dari foto; harga jual perlu konfirmasi",fmt["input"])
    listv(s,"F7:F56",["AKTIF","NONAKTIF"]);s.set_column("B:B",30);s.set_column("G:G",55);audit(s,"H",56)

    s=base("LOT","Satu kode lot unik untuk satu produk dan satu tanggal kedaluwarsa. Isi dari kemasan/catatan aktual. Stok dicatat melalui MUTASI.")
    table(s,"tblLot",["Kode Lot","SKU","Nama Produk","Kedaluwarsa","Catatan","Cek"],CAP["LOT"],{
        "C":lambda r:f'IF(A{r}="","",{lookup(f"B{r}","PRODUK","A","B")})',
        "F":lambda r:f'IF(A{r}="","",IF(COUNTIF(DaftarLot,A{r})>1,"LOT GANDA",IF(COUNTIF(DaftarSKU,B{r})<>1,"CEK SKU",IF(OR(NOT(ISNUMBER(D{r})),D{r}<=0),"ISI TANGGAL EXP","OK"))))'},set("ABDE"),{"D":"input_date"})
    listv(s,"B7:B86","=DaftarSKU");s.set_column("C:C",30);s.set_column("E:E",45);audit(s,"F",86)

    # Movement audit is independent of stock arithmetic, avoiding a circular reference.
    def move_check(r):
        invoice_region=lookup(f'H{r}',"INVOICE","A","C")
        inv_start=lookup(f'H{r}',"INVOICE","A","D","0")
        inv_end=lookup(f'H{r}',"INVOICE","A","E","0")
        inv_state=lookup(f'H{r}',"INVOICE","A","I")
        checks = [
            (f'COUNTIF({rng("MUTASI","A")},A{r})>1',"ID GANDA"),
            (f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())',"CEK TANGGAL"),
            (f'COUNTIF(JenisMutasi,C{r})<>1',"CEK JENIS"),
            (f'COUNTIF(DaftarLokasi,D{r})<>1',"CEK LOKASI"),
            (f'{lookup(f"D{r}","OUTLET","A","H")}<>"OK"',"MASTER OUTLET ERROR"),
            (f'COUNTIF(DaftarLot,E{r})<>1',"CEK LOT"),
            (f'{lookup(f"E{r}","LOT","A","F")}<>"OK"',"MASTER LOT ERROR"),
            (f'OR(NOT(ISNUMBER(F{r})),F{r}<=0,F{r}<>INT(F{r}))',"CEK QTY"),
            (f'AND(I{r}<>"AKTIF",I{r}<>"VOID")',"CEK STATUS"),
            (f'AND(OR(C{r}="KIRIM",C{r}="JUAL",C{r}="RETUR"),D{r}="GDG")',"PILIH OUTLET"),
            (f'AND(C{r}="PRODUKSI",D{r}<>"GDG")',"PRODUKSI DI GDG"),
            (f'AND(OR(C{r}="KIRIM",C{r}="JUAL"),B{r}>N{r})',"SUDAH KEDALUWARSA"),
            (f'AND(C{r}="JUAL",OR(NOT(ISNUMBER(G{r})),G{r}<=0))',"ISI HARGA JUAL"),
            (f'AND(C{r}="KIRIM",U{r}="")',"ISI NO PO/SJ"),
            (f'AND(W{r}<>"",OR(NOT(ISNUMBER(W{r})),W{r}<0))',"CEK ONGKIR AKTUAL"),
            (f'AND(H{r}<>"",C{r}<>"JUAL")',"INVOICE HANYA JUAL"),
            (f'AND(H{r}<>"",COUNTIF(DaftarInvoice,H{r})<>1)',"CEK ID INVOICE"),
            (f'AND(H{r}<>"",OR({invoice_region}<>M{r},B{r}<{inv_start},B{r}>{inv_end},{inv_state}="VOID"))',"CEK AREA/PERIODE INV"),
        ]
        expr='"OK"'
        for condition, message in reversed(checks):
            expr=f'IF({condition},"{message}",{expr})'
        return f'IF(A{r}="","",{expr})'
    s=base("MUTASI","Jurnal barang: AWAL/PRODUKSI masuk; KIRIM ke outlet; JUAL menjadi pendapatan; RETUR kembali ke gudang; RUSAK/KURANG mengurangi stok. Jangan hapus riwayat.",23)
    forms={
        "K":lambda r:f'IF(A{r}="","",{lookup(f"E{r}","LOT","A","B")})',
        "L":lambda r:f'IF(A{r}="","",{lookup(f"K{r}","PRODUK","A","B")})',
        "M":lambda r:f'IF(A{r}="","",{lookup(f"D{r}","OUTLET","A","C")})',
        "N":lambda r:f'IF(A{r}="","",{lookup(f"E{r}","LOT","A","D","0")})',
        "O":lambda r:f'IF(A{r}="","",IF(C{r}="KIRIM","GDG",IF(OR(C{r}="JUAL",C{r}="RETUR",C{r}="RUSAK",C{r}="KURANG"),D{r},"")))',
        "P":lambda r:f'IF(A{r}="","",IF(C{r}="RETUR","GDG",IF(OR(C{r}="AWAL",C{r}="PRODUKSI",C{r}="KIRIM",C{r}="TAMBAH"),D{r},"")))',
        "Q":lambda r:f'IF(AND(C{r}="JUAL",S{r}>0),F{r}*G{r},0)',
        "R":move_check,
        "S":lambda r:f'IF(AND(A{r}<>"",I{r}="AKTIF",R{r}="OK"),F{r},0)',
        "T":lambda r:(f'IF(OR(A{r}="",O{r}="",S{r}=0),"",'
          f'SUMIFS({rng("MUTASI","S")},{rng("MUTASI","P")},O{r},{rng("MUTASI","E")},E{r},{rng("MUTASI","B")},"<"&B{r})'
          f'-SUMIFS({rng("MUTASI","S")},{rng("MUTASI","O")},O{r},{rng("MUTASI","E")},E{r},{rng("MUTASI","B")},"<"&B{r})'
          f'+SUMIFS($S$7:S{r},$P$7:P{r},O{r},$E$7:E{r},E{r},$B$7:B{r},B{r})'
          f'-SUMIFS($S$7:S{r},$O$7:O{r},O{r},$E$7:E{r},E{r},$B$7:B{r},B{r}))'),
    }
    table(s,"tblMutasi",["ID Mutasi","Tanggal","Jenis","Kode Lokasi","Kode Lot","Qty","Harga Jual Snapshot","ID Invoice","Status","Catatan","SKU","Produk","Wilayah","Kedaluwarsa","Asal","Tujuan","Bruto Penjualan","Cek","Qty Efektif","Saldo Asal Sesudah","No PO / Surat Jalan","Ekspedisi","Ongkir Aktual"],CAP["MUTASI"],forms,set("ABCDEFGHIJUVW"),{"B":"input_date","G":"input_money","N":"date","Q":"money","S":"qty","T":"qty","W":"input_money"})
    for c,source in [("C","=JenisMutasi"),("D","=DaftarLokasi"),("E","=DaftarLot"),("H","=DaftarInvoice"),("I",["AKTIF","VOID"])]:listv(s,f"{c}7:{c}2006",source)
    s.set_column("J:L",27);s.set_column("G:G",22);s.set_column("S:T",21);audit(s,"R",2006)
    s.conditional_format("T7:T2006",{"type":"cell","criteria":"<","value":0,"format":fmt["error"]})
    s.write_comment("H6","Hanya tautkan penjualan yang belum ditagih. Jangan mengganti harga, area, atau kuantitas penjualan yang sudah masuk invoice TERBIT; gunakan proses koreksi dan dokumentasikan.")

    s=base("STOK","Saldo kumulatif sampai tanggal akhir DASHBOARD, otomatis per lokasi dan lot. Stok tersisa terbawa antarbulan; expired tidak otomatis dianggap terjual atau dihapus.",12)
    n=CAP["OUTLET"]*CAP["LOT"]
    stock={
        "A":lambda r:f'IF(OR(INDEX({rng("OUTLET","A")},INT((ROW()-7)/80)+1)="",INDEX({rng("LOT","A")},MOD(ROW()-7,80)+1)=""),"",INDEX({rng("OUTLET","A")},INT((ROW()-7)/80)+1))',
        "B":lambda r:f'IF(A{r}="","",{lookup(f"A{r}","OUTLET","A","B")})',
        "C":lambda r:f'IF(A{r}="","",{lookup(f"A{r}","OUTLET","A","C")})',
        "D":lambda r:f'IF(A{r}="","",INDEX({rng("LOT","A")},MOD(ROW()-7,80)+1))',
        "E":lambda r:f'IF(A{r}="","",{lookup(f"D{r}","LOT","A","B")})',
        "F":lambda r:f'IF(A{r}="","",{lookup(f"E{r}","PRODUK","A","B")})',
        "G":lambda r:f'IF(A{r}="","",{lookup(f"D{r}","LOT","A","D","0")})',
        "H":lambda r:f'IF(A{r}="","",SUMIFS({rng("MUTASI","S")},{rng("MUTASI","P")},A{r},{rng("MUTASI","E")},D{r},{rng("MUTASI","B")},"<="&TglAkhir))',
        "I":lambda r:f'IF(A{r}="","",SUMIFS({rng("MUTASI","S")},{rng("MUTASI","O")},A{r},{rng("MUTASI","E")},D{r},{rng("MUTASI","B")},"<="&TglAkhir))',
        "J":lambda r:f'IF(A{r}="","",H{r}-I{r})',
        "K":lambda r:f'IF(A{r}="","",IF(J{r}<0,"STOK NEGATIF",IF(J{r}=0,"KOSONG",IF(G{r}<TglAkhir,"KEDALUWARSA",IF(G{r}<=TglAkhir+REFERENSI!$B$12,"MENDEKATI EXP","TERSEDIA")))))',
        "L":lambda r:f'IF(A{r}="","",IF(G{r}=0,"",G{r}-TglAkhir))',
    }
    # STOK is a read-only Cartesian register, not a manually maintained list.
    table(s,"tblStok",["Kode Lokasi","Lokasi","Wilayah","Lot","SKU","Produk","Kedaluwarsa","Total Masuk","Total Keluar","Saldo","Status","Hari ke Exp"],n,stock,formats={"G":"date","H":"qty","I":"qty","J":"qty"})
    s.set_column("B:B",30);s.set_column("F:F",26);s.set_column("K:K",24)
    for state in ["STOK NEGATIF","KEDALUWARSA"]:
        s.conditional_format(f"K7:K{6+n}",{"type":"text","criteria":"containing","value":state,"format":fmt["error"]})
    s.conditional_format(f"K7:K{6+n}",{"type":"text","criteria":"containing","value":"MENDEKATI","format":fmt["warning"]})

    def inv_check(r):
        return (f'IF(A{r}="","",IF(COUNTIF(DaftarInvoice,A{r})>1,"ID GANDA",'
                f'IF(COUNTIF(DaftarWilayah,C{r})<>1,"CEK WILAYAH",'
                f'IF(OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY(),NOT(ISNUMBER(D{r})),D{r}<=0,NOT(ISNUMBER(E{r})),E{r}<D{r},E{r}>B{r},NOT(ISNUMBER(F{r})),F{r}<B{r}),"CEK TANGGAL",'
                f'IF(OR(NOT(ISNUMBER(G{r})),G{r}<0,G{r}>1,NOT(ISNUMBER(H{r})),H{r}<0,NOT(ISNUMBER(J{r})),J{r}<=0),"CEK NOMINAL",'
                f'IF(AND(I{r}<>"DRAF",I{r}<>"TERBIT",I{r}<>"VOID"),"CEK STATUS",'
                f'IF(I{r}="VOID","VOID",IF(ABS(J{r}-Q{r})>0.5,"SELISIH SUMBER","OK"))))))))')
    s=base("INVOICE","Satu invoice per wilayah dan kumpulan penjualan yang ditautkan. Bruto snapshot tetap; selisih sumber mendeteksi penjualan yang berubah setelah penagihan.",18)
    invf={
        "K":lambda r:f'IF(A{r}="","",ROUND(N(J{r})*N(G{r}),0))',
        "L":lambda r:f'IF(A{r}="","",ROUND(N(J{r})-N(K{r})+N(H{r}),0))',
        "M":lambda r:f'IF(A{r}="","",SUMIFS({rng("PEMBAYARAN","J")},{rng("PEMBAYARAN","C")},A{r}))',
        "N":lambda r:f'IF(A{r}="","",IF(I{r}="TERBIT",L{r}-M{r},0))',
        "O":lambda r:f'IF(A{r}="","",IF(I{r}<>"TERBIT",I{r},IF(M{r}>L{r},"LEBIH BAYAR",IF(N{r}=0,"LUNAS",IF(M{r}=0,"BELUM BAYAR","SEBAGIAN")))))',
        "P":inv_check,
        "Q":lambda r:f'IF(A{r}="","",SUMIFS({rng("MUTASI","Q")},{rng("MUTASI","H")},A{r}))',
    }
    table(s,"tblInvoice",["ID Invoice","Tanggal Terbit","Kode Wilayah","Periode Awal","Periode Akhir","Jatuh Tempo","Diskon %","Ongkir Ditagih","Status","Bruto Snapshot","Diskon Rp","Total Tagihan","Sudah Dibayar","Sisa Piutang","Status Bayar","Cek","Bruto Sumber","Catatan"],CAP["INVOICE"],invf,set("ABCDEFGHIJR"),{"B":"input_date","D":"input_date","E":"input_date","F":"input_date","G":"input_pct","H":"input_money","J":"input_money","K":"money","L":"money","M":"money","N":"money","Q":"money"})
    listv(s,"C7:C406","=DaftarWilayah");listv(s,"I7:I406",["DRAF","TERBIT","VOID"])
    s.data_validation("G7:G406",{"validate":"decimal","criteria":"between","minimum":0,"maximum":1,"error_type":"stop","error_message":"Isi persentase antara 0% dan 100%."})
    s.set_column("J:Q",23);s.set_column("R:R",40);audit(s,"P",406)

    s=base("PEMBAYARAN","Pembayaran dapat dicicil. Nominal masuk hanya untuk invoice TERBIT. Pembayaran berlebih ditandai pada INVOICE; jangan dianggap tambahan penjualan.")
    payf={
        "H":lambda r:f'IF(A{r}="","",{lookup(f"C{r}","INVOICE","A","C")})',
        "I":lambda r:(f'IF(A{r}="","",IF(COUNTIF({rng("PEMBAYARAN","A")},A{r})>1,"ID GANDA",'
          f'IF(OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY()),"CEK TANGGAL",'
          f'IF(COUNTIF(DaftarInvoice,C{r})<>1,"CEK INVOICE",'
          f'IF({lookup(f"C{r}","INVOICE","A","I")}<>"TERBIT","INVOICE BELUM TERBIT",'
          f'IF(B{r}<{lookup(f"C{r}","INVOICE","A","B","0")},"SEBELUM INVOICE",'
          f'IF(COUNTIF(MetodeBayar,D{r})<>1,"CEK METODE",'
          f'IF(OR(NOT(ISNUMBER(E{r})),E{r}<=0),"CEK NOMINAL",'
          f'IF(AND(F{r}<>"AKTIF",F{r}<>"VOID"),"CEK STATUS","OK")))))))))'),
        "J":lambda r:f'IF(AND(A{r}<>"",F{r}="AKTIF",I{r}="OK"),E{r},0)',
    }
    table(s,"tblPembayaran",["ID Pembayaran","Tanggal","ID Invoice","Metode","Nominal","Status","Catatan","Wilayah","Cek","Nominal Efektif"],CAP["PEMBAYARAN"],payf,set("ABCDEFG"),{"B":"input_date","E":"input_money","J":"money"})
    listv(s,"C7:C1006","=DaftarInvoice");listv(s,"D7:D1006","=MetodeBayar");listv(s,"F7:F1006",["AKTIF","VOID"])
    s.set_column("G:G",40);audit(s,"I",1006)

    # Operator form: optional VBA posts values into the same auditable ledgers.
    s=base("FORM","Opsional: gunakan tombol setelah impor VBA ke salinan .xlsm. Tanpa makro, isi tabel MUTASI / INVOICE / PEMBAYARAN secara langsung.")
    s.set_column("B:B",25);s.set_column("E:E",27)
    sections=[(5,"01  CATAT MUTASI BARANG"),(19,"02  TERBITKAN INVOICE WILAYAH"),(36,"03  CATAT PEMBAYARAN")]
    for row,text in sections:s.merge_range(row-1,0,row-1,9,text,fmt["section"])
    inputs=[("A6","Tanggal","B6","date"),("D6","Jenis Mutasi","E6","text"),("A8","Kode Lokasi","B8","text"),("D8","Kode Lot","E8","text"),("A10","Qty (pouch)","B10","qty"),("D10","Harga jual / pouch","E10","money"),("A12","Catatan","B12","text"),
            ("A20","Tanggal Terbit","B20","date"),("D20","Kode Wilayah","E20","text"),("A22","Periode Awal","B22","date"),("D22","Periode Akhir","E22","date"),("A24","Jatuh Tempo","B24","date"),("D24","Diskon %","E24","pct"),("A26","Ongkir Ditagih","B26","money"),("A28","Catatan","B28","text"),
            ("A37","Tanggal Bayar","B37","date"),("D37","ID Invoice","E37","text"),("A39","Metode","B39","text"),("D39","Nominal","E39","money"),("A41","Catatan","B41","text")]
    for label,text,cell,kind in inputs:
        s.write(label,text,fmt["label"])
        s.write_blank(cell,None,fmt[{"date":"input_date","money":"input_money","pct":"input_pct"}.get(kind,"input")])
    today=dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    for c in ["B6","B20","B37"]:s.write_datetime(c,today,fmt["input_date"])
    s.write("E6","KIRIM",fmt["input"]);s.write("E24",0,fmt["input_pct"]);s.write("B26",0,fmt["input_money"])
    for label,text,cell in [("G8","No PO / Surat Jalan","H8"),("G10","Ekspedisi","H10"),("G12","Ongkir Aktual","H12")]:
        s.write(label,text,fmt["label"]);s.write_blank(cell,None,fmt["input_money"] if cell=="H12" else fmt["input"])
    for c,source in [("E6","=JenisMutasi"),("B8","=DaftarLokasi"),("E8","=DaftarLot"),("E20","=DaftarWilayah"),("E37","=DaftarInvoice"),("B39","=MetodeBayar")]:listv(s,c,source)
    s.write("A14","Harga katalog",fmt["label"])
    s.write_formula("B14",f'=IF(E8="","",{lookup(lookup("E8","LOT","A","B"),"PRODUK","A","D","0")})',fmt["money"],"")
    s.write("D14","Stok asal s/d tanggal",fmt["label"])
    source='IF(E6="KIRIM","GDG",B8)'
    s.write_formula("E14",f'=IF(OR(B8="",E8=""),"",SUMIFS({rng("MUTASI","S")},{rng("MUTASI","P")},{source},{rng("MUTASI","E")},E8,{rng("MUTASI","B")},"<="&B6)-SUMIFS({rng("MUTASI","S")},{rng("MUTASI","O")},{source},{rng("MUTASI","E")},E8,{rng("MUTASI","B")},"<="&B6))',fmt["qty"],"")
    s.merge_range("A16:J17","SimpanMutasi: harga jual disimpan sebagai nilai tetap. AWAL hanya untuk saldo pembukaan; stok tidak direset saat berganti bulan. Untuk transaksi yang sudah ditagih, lakukan koreksi terdokumentasi.",fmt["note"])
    s.write("A30","Bruto belum ditagih",fmt["label"])
    s.write_formula("B30",f'=SUMIFS({rng("MUTASI","Q")},{rng("MUTASI","M")},E20,{rng("MUTASI","B")},">="&B22,{rng("MUTASI","B")},"<="&E22,{rng("MUTASI","H")},"")',fmt["money"],0)
    s.merge_range("A32:J34","TerbitkanInvoice mengambil hanya JUAL aktif yang valid, belum memiliki ID invoice, sesuai wilayah dan periode. Diskon dihitung dari bruto; ongkir ditambahkan setelah diskon. Barang dikirim tetapi belum terjual tidak ditagihkan.",fmt["note"])
    s.merge_range("A43:J45","SimpanPembayaran menautkan penerimaan ke invoice TERBIT. Pilih invoice di CETAK INVOICE untuk cetak/PDF. Pastikan PIC dan alamat tagihan sudah lengkap sebelum mengirim invoice.",fmt["note"])
    s.print_area("A1:J45")

    # Management view: period sales, current stock and invoice issue-date cash metrics are explicit.
    s=base("DASHBOARD","Ringkasan konsinyasi. Penjualan memakai tanggal JUAL; tagihan memakai tanggal terbit; kas memakai tanggal pembayaran; stok kumulatif sampai tanggal akhir.")
    s.freeze_panes(6,0)
    s.write("A5","Periode awal",fmt["label"]);s.write_datetime("B5",today.replace(day=1),fmt["input_date"])
    s.write("D5","Periode akhir",fmt["label"]);s.write_datetime("E5",today,fmt["input_date"])
    wb.define_name("TglAwal","='DASHBOARD'!$B$5");wb.define_name("TglAkhir","='DASHBOARD'!$E$5")
    s.merge_range("G5:L5","Ubah sel biru untuk periode laporan",fmt["note"])
    def between(sheet,datecol):return f'{rng(sheet,datecol)},">="&TglAwal,{rng(sheet,datecol)},"<="&TglAkhir'
    kpis=[("PENJUALAN BRUTO",f'SUMIFS({rng("MUTASI","Q")},{between("MUTASI","B")})'),
          ("TAGIHAN TERBIT",f'SUMIFS({rng("INVOICE","L")},{rng("INVOICE","I")},"TERBIT",{between("INVOICE","B")})'),
          ("PEMBAYARAN MASUK",f'SUMIFS({rng("PEMBAYARAN","J")},{between("PEMBAYARAN","B")})'),
          ("PIUTANG S/D AKHIR",f'SUMIFS({rng("INVOICE","L")},{rng("INVOICE","I")},"TERBIT",{rng("INVOICE","B")},"<="&TglAkhir)-SUMIFS({rng("PEMBAYARAN","J")},{rng("PEMBAYARAN","B")},"<="&TglAkhir)'),
          ("STOK GUDANG • POUCH",f'SUMIFS(STOK!$J$7:$J${6+n},STOK!$A$7:$A${6+n},"GDG")'),
          ("STOK OUTLET • POUCH",f'SUMIFS(STOK!$J$7:$J${6+n},STOK!$A$7:$A${6+n},"<>GDG",STOK!$A$7:$A${6+n},"<>")')]
    for i,(label,formula) in enumerate(kpis):
        rr=6+(i//3)*5; cc=(i%3)*4
        s.merge_range(rr,cc,rr,cc+2,label,fmt["kpi_label"])
        s.merge_range(rr+1,cc,rr+2,cc+2,"",fmt["kpi"])
        s.write_formula(rr+1,cc,"="+formula,fmt["kpi"],0)
    s.merge_range("A17:L17","PERIKSA SEBELUM MENGIRIM TAGIHAN",fmt["section"])
    alerts=[("Mutasi tidak valid",f'COUNTIFS({rng("MUTASI","A")},"<>",{rng("MUTASI","R")},"<>OK")'),
            ("Saldo historis negatif",f'COUNTIF({rng("MUTASI","T")},"<0")'),
            ("Invoice bermasalah",f'COUNTIFS({rng("INVOICE","A")},"<>",{rng("INVOICE","P")},"<>OK",{rng("INVOICE","P")},"<>VOID")'),
            ("Pembayaran tidak valid",f'COUNTIFS({rng("PEMBAYARAN","A")},"<>",{rng("PEMBAYARAN","I")},"<>OK")'),
            ("Stok expired (pouch)",f'SUMIFS(STOK!$J$7:$J${6+n},STOK!$K$7:$K${6+n},"KEDALUWARSA")'),
            ("Bruto belum ditagih",f'SUMIFS({rng("MUTASI","Q")},{rng("MUTASI","H")},"",{rng("MUTASI","B")},"<="&TglAkhir)')]
    for i,(label,formula) in enumerate(alerts):
        rr=18+i//3*2;cc=i%3*4
        s.merge_range(rr,cc,rr,cc+1,label,fmt["label"])
        s.write_formula(rr,cc+2,"="+formula,fmt["money"],0)
        s.conditional_format(rr,cc+2,rr,cc+2,{"type":"cell","criteria":">","value":0,"format":fmt["warning"]})
    headers=["Kode","Wilayah","Terjual (pouch)","Bruto","Tagihan Terbit","Pembayaran","Stok Akhir"]
    s.write_row("A24",headers,fmt["header"]);s.set_row(23,32)
    for i in range(CAP["WILAYAH"]):
        r=25+i;master=FIRST+i
        s.write_formula(f"A{r}",f'=IF(WILAYAH!A{master}="","",WILAYAH!A{master})',fmt["calc"],"")
        s.write_formula(f"B{r}",f'=IF(A{r}="","",WILAYAH!B{master})',fmt["calc"],"")
        fs={"C":f'SUMIFS({rng("MUTASI","S")},{rng("MUTASI","C")},"JUAL",{rng("MUTASI","M")},A{r},{between("MUTASI","B")})',
            "D":f'SUMIFS({rng("MUTASI","Q")},{rng("MUTASI","M")},A{r},{between("MUTASI","B")})',
            "E":f'SUMIFS({rng("INVOICE","L")},{rng("INVOICE","C")},A{r},{rng("INVOICE","I")},"TERBIT",{between("INVOICE","B")})',
            "F":f'SUMIFS({rng("PEMBAYARAN","J")},{rng("PEMBAYARAN","H")},A{r},{between("PEMBAYARAN","B")})',
            "G":f'SUMIFS(STOK!$J$7:$J${6+n},STOK!$C$7:$C${6+n},A{r})'}
        for c,formula in fs.items():s.write_formula(f"{c}{r}",f'=IF(A{r}="","",{formula})',fmt["qty" if c in ["C","G"] else "money"],"")
    s.set_column("B:B",26)
    chart=wb.add_chart({"type":"bar"})
    chart.add_series({"name":"Penjualan bruto per wilayah","categories":"='DASHBOARD'!$B$25:$B$44","values":"='DASHBOARD'!$D$25:$D$44","fill":{"color":teal},"border":{"none":True}})
    chart.set_title({"name":"Penjualan menurut wilayah"});chart.set_legend({"none":True})
    chart.set_x_axis({"num_format":"#,##0","major_gridlines":{"visible":True,"line":{"color":"#E3E9ED"}}})
    chart.set_y_axis({"reverse":True});chart.set_chartarea({"border":{"none":True}})
    s.insert_chart("I24",chart,{"x_scale":1.1,"y_scale":1.25})
    s.merge_range("A47:L49","Piutang dan stok memakai tanggal akhir; status invoice saat ini memengaruhi laporan historis. Tagihan TERBIT tetap memakai snapshot walaupun ada selisih sumber: periksa alert sebelum mengirim tagihan. Ini alat satu pengguna tanpa penguncian periode. Bruto penjualan berbeda dari kas masuk. Rupiah; qty pouch.",fmt["note"])
    s.print_area("A1:Q48")

    # Invoice layout contains every product slot; unused lines can be hidden by the PDF macro.
    s=base("CETAK INVOICE","Pilih ID invoice. DRAF, VOID, data tidak lengkap, atau selisih sumber akan ditandai. Contoh invoice asli belum diberikan.")
    s.set_portrait();s.set_column("A:A",6);s.set_column("B:B",24);s.set_column("C:C",20);s.set_column("D:D",13);s.set_column("E:E",21);s.set_column("F:F",22)
    s.write("A5","ID",fmt["label"]);s.write_blank("B5",None,fmt["input"]);listv(s,"B5","=DaftarInvoice")
    s.merge_range("A7:F7","GIAN FATAYA",fmt["print_title"]);s.merge_range("A8:F8","INVOICE KONSINYASI  •  berdasarkan barang terjual",fmt["label"])
    s.write("A9","No.",fmt["label"])
    s.merge_range("B9:F9","",fmt["label"])
    s.write_formula("B9",'=IF($B$5="","",$B$5)',fmt["label"],"")
    s.merge_range("B12:C13","",fmt["body"])
    s.write("D12","Periode",fmt["label"])
    s.merge_range("E12:F13","",fmt["body"])
    s.write_formula("E12",f'=IF($B$5="","",TEXT({lookup("$B$5","INVOICE","A","D","0")},"dd mmm yyyy")&" - "&TEXT({lookup("$B$5","INVOICE","A","E","0")},"dd mmm yyyy"))',fmt["body"],"")
    fields=[("A10","Wilayah","B10","C","WILAYAH","B"),("A11","Kepada / PIC","B11","C","WILAYAH","C"),("A12","Alamat","B12","C","WILAYAH","D"),("D10","Terbit","E10","B",None,None),("D11","Jatuh tempo","E11","F",None,None)]
    for lab,text,cell,c,extra,res in fields:
        s.write(lab,text,fmt["label"])
        expr=lookup("$B$5","INVOICE","A",c)
        if extra:expr=lookup(expr,extra,"A",res)
        s.write_formula(cell,f'=IF($B$5="","",{expr})',fmt["date"] if cell.startswith("E") else (fmt["body"] if cell=="B12" else fmt["calc"]),"")
    s.merge_range("A14:F14","",fmt["warning"])
    s.write_formula("A14",f'=IF(B5="","PILIH ID INVOICE",IF({lookup("B5","INVOICE","A","I")}<>"TERBIT","BUKAN INVOICE TERBIT",IF({lookup("B5","INVOICE","A","P")}<>"OK","STOP: PERIKSA INVOICE DAN SUMBER",IF(OR(B11="",B12=""),"LENGKAPI PIC DAN ALAMAT TAGIHAN","SIAP CETAK"))))',fmt["warning"],"PILIH ID INVOICE")
    s.write_row("A16",["No","SKU","Produk","Qty","Harga rata-rata","Bruto (Rp)"],fmt["header"])
    for i in range(CAP["PRODUK"]):
        r=17+i;master=FIRST+i
        s.write_formula(f"A{r}",f'=IF(D{r}>0,COUNTIF($D$17:D{r},">0"),"")',fmt["print_cell"],"")
        s.write_formula(f"B{r}",f'=IF(D{r}>0,PRODUK!A{master},"")',fmt["print_cell"],"")
        s.write_formula(f"C{r}",f'=IF(D{r}>0,PRODUK!B{master},"")',fmt["print_cell"],"")
        s.write_formula(f"D{r}",f'=IF(OR($B$5="",PRODUK!A{master}=""),0,SUMIFS({rng("MUTASI","S")},{rng("MUTASI","H")},$B$5,{rng("MUTASI","K")},PRODUK!A{master},{rng("MUTASI","C")},"JUAL"))',fmt["print_money"],0)
        s.write_formula(f"F{r}",f'=IF(D{r}=0,0,SUMIFS({rng("MUTASI","Q")},{rng("MUTASI","H")},$B$5,{rng("MUTASI","K")},PRODUK!A{master}))',fmt["print_money"],0)
        s.write_formula(f"E{r}",f'=IF(D{r}=0,"",F{r}/D{r})',fmt["print_money"],"")
    for r,label,c in [(68,"Bruto","J"),(69,"Diskon (Rp)","K"),(70,"Ongkir ditagih","H"),(71,"TOTAL TAGIHAN","L"),(72,"Sudah dibayar","M"),(73,"Sisa piutang","N")]:
        s.merge_range(r-1,2,r-1,4,label,fmt["label"])
        s.write_formula(f"F{r}",f'=IF($B$5="","",{lookup("$B$5","INVOICE","A",c,"0")})',fmt["money"],"")
    s.merge_range("A75:F76","",fmt["note"])
    s.write_formula("A75",'=IF(REFERENSI!B8="","Informasi pembayaran belum diisi. Lengkapi REFERENSI sebelum mengirim invoice.","Pembayaran: "&REFERENSI!B8)',fmt["note"],"Informasi pembayaran belum diisi.")
    s.merge_range("A78:F79","Harga rata-rata = bruto dibagi kuantitas jika ada perubahan harga dalam periode. Rincian outlet/lot tersedia di MUTASI dengan filter ID invoice. Invoice ini tidak menambahkan pajak karena ketentuan pajak belum diberikan.",fmt["note"])
    s.print_area("A7:F79");s.repeat_rows(6,15);s.fit_to_pages(1,0)

    s=base("BAHAN","Referensi foto Bahan Baku & Harga_GianFataya: penggunaan dan biaya bahan per satu pouch. Bukan pencatatan stok bahan baku atau produksi otomatis.")
    s.write_row("A6",["Bahan","Penggunaan / pouch","Satuan","Biaya / pouch (Rp)"],fmt["header"])
    s.set_column("A:B",25);s.set_column("D:D",26)
    for i,row in enumerate(profile["ingredients"],6):
        s.write_row(i,0,row,fmt["input"])
        s.write_number(i,3,row[3],fmt["input_money"])
    s.merge_range("A14:C14","TOTAL BAHAN + KEMASAN",fmt["section"]);s.write_formula("D14","=SUM(D7:D12)",fmt["money"],9410)
    s.merge_range("A17:H20","Angka disalin dari foto mitra: tepung 60 g; minyak 30 ml; abon 25 g; margarin 60 g; telor 35 g; pouch 1 lembar. Nilai ini tidak membuktikan takaran telah dikonfirmasi ulang. Harga jual tetap kosong. Sistem utama berfokus pada stok produk jadi, distribusi, penjualan dan invoice sesuai prioritas MoM.",fmt["warning"])

    s=base("REFERENSI","Sumber, keputusan desain, konfigurasi dan data yang masih harus dilengkapi. Tidak menyimpan kredensial atau data rekening yang dibuat-buat.")
    s.write("A6","Pemilik",fmt["label"]);s.write("B6","Bu Santi",fmt["input"])
    s.write("A8","Info pembayaran",fmt["label"]);s.write_blank("B8",None,fmt["input"]);s.set_column("B:B",45)
    s.write("A10","Alamat pengirim",fmt["label"]);s.write_blank("B10",None,fmt["input"])
    s.write("A12","Peringatan exp (hari)",fmt["label"]);s.write("B12",30,fmt["input"])
    s.write_comment("B12","30 hari adalah ambang pengingat operasional yang dapat diubah, bukan asumsi masa simpan produk.")
    s.write_url("A15",profile["source_sheet"],fmt["nav"],"MoM • 22 Oktober 2026")
    s.write_url("A17",profile["source_folder"],fmt["nav"],"Folder data Gian Fataya")
    notes=[
        "Wilayah mengikuti foto sumber: Jawa Timur; Jawa Tengah; Medan - Aceh; Sulawesi Selatan; Kalimantan; Sumatera Selatan; Batam; Jabodetabek. Nama/cakupan ini berbeda sebagian dari catatan rapat dan perlu konfirmasi.",
        "MoM menyebut konsinyasi: barang dikirim belum menjadi penjualan; outlet membayar hanya barang yang terjual. Invoice ditagih kepada penanggung jawab wilayah.",
        "Minimum PO 10 pcs adalah informasi MoM, bukan pembatas mutasi: sistem tetap menerima retur atau penjualan parsial di bawah 10 pouch.",
        "Belum tersedia: harga jual/net konsinyasi; daftar outlet dan PIC wilayah; stok awal per lot; masa simpan/kedaluwarsa; aturan ongkir dan diskon; rekening; nomor/format invoice asli.",
        "Diskon berlaku sekali pada bruto invoice; ongkir ditagih ditambahkan setelah diskon. Tidak ada pajak atau potongan komisi tersembunyi. Konfirmasi jika mitra memerlukan harga/komisi berbeda per outlet.",
        "Barang expired tetap tercatat sebagai stok sampai ada keputusan dan mutasi RUSAK/RETUR. Tidak dibuat otomatis sebagai stok baru atau dianggap penjualan.",
        "Harga JUAL dan bruto invoice berupa snapshot nilai tetap. Jangan mengubah penjualan yang sudah ditagih. Pembatalan invoice harus melepas tautan mutasi dan membatalkan pembayaran terkait terlebih dahulu.",
        "Baris ber-ID kosong tidak diproses. Gunakan ID unik dan status AKTIF; semua gerakan qty positif. Koreksi TAMBAH/KURANG harus disertai catatan. Kapasitas tabel tidak otomatis meluas dalam versi ini.",
    ]
    for i,text in enumerate(notes):
        rr=19+i*3;s.merge_range(rr,0,rr+1,11,text,fmt["body"])
    s.write_column("N2",TYPES);wb.define_name("JenisMutasi","='REFERENSI'!$N$2:$N$9")
    s.write_column("O2",["Tunai","Transfer","QRIS","Lainnya"]);wb.define_name("MetodeBayar","='REFERENSI'!$O$2:$O$5")
    s.set_column("N:O",None,None,{"hidden":True})
    s.print_area("A1:L44")
    order=["MULAI","DASHBOARD","FORM","WILAYAH","OUTLET","PRODUK","LOT","MUTASI","STOK","INVOICE","PEMBAYARAN","CETAK INVOICE","BAHAN","REFERENSI"]
    wb.worksheets_objs=[sheets[name] for name in order]
    for index,sheet in enumerate(wb.worksheets_objs):
        sheet.index=index
    sheets["MULAI"].activate()
    wb.close()
    return dest


if __name__ == "__main__":
    parser=argparse.ArgumentParser()
    parser.add_argument("--output",type=Path,default=ROOT/"output/POS_Gian_Fataya.xlsx")
    args=parser.parse_args()
    print(build(args.output))
