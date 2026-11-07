"""Six physical sheets, with explicit relocation of accounting sections.

The staging workbook is a formula engine, not a user-facing extra workbook.
Formula tokens, names, tables, validation and conditional formatting are moved
together. No sheets are hidden to make the tab count appear smaller.
"""
from __future__ import annotations

from copy import copy, deepcopy
from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties


NAV = ["BERANDA", "TRANSAKSI", "STOK", "TAGIHAN", "CETAK INVOICE", "DATA USAHA"]
NAV_LABELS = ["Beranda", "Catat barang", "Lihat stok", "Tagihan & bayar", "Cetak invoice", "Data usaha"]
# Logical source column -> compact TRANSAKSI column. Inputs A:I; checks J:L.
MOVE_COLS = {
    1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 21: 8, 23: 9,
    18: 10, 24: 11, 20: 12, 8: 13, 22: 14, 9: 15, 10: 16,
    11: 17, 12: 18, 13: 19, 14: 20, 15: 21, 16: 22, 17: 23, 19: 24,
}


@dataclass(frozen=True)
class Section:
    sheet: str
    offset: int = 0


SECTIONS = {
    "MULAI": Section("BERANDA"),
    "DASHBOARD": Section("BERANDA"),
    "FORM": Section("BERANDA", 70),
    "MUTASI": Section("TRANSAKSI"),
    "STOK": Section("STOK"),
    "INVOICE": Section("TAGIHAN"),
    "PEMBAYARAN": Section("TAGIHAN", 430),
    "CETAK INVOICE": Section("CETAK INVOICE"),
    "WILAYAH": Section("DATA USAHA"),
    "OUTLET": Section("DATA USAHA", 28),
    "PRODUK": Section("DATA USAHA", 96),
    "LOT": Section("DATA USAHA", 154),
    "BAHAN": Section("DATA USAHA", 242),
    "REFERENSI": Section("DATA USAHA", 262),
}

CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")
NAVY, TEAL, INK, MUTED = "123044", "087F8C", "203B4A", "526674"
PALE, BLUE, GOLD = "F0F5F7", "EDF5FC", "FFF2D5"


def address(logical: str, ref: str) -> str:
    """Map a cell or rectangular range. Accounting ranges are column-local."""
    def one(value):
        match = CELL.fullmatch(value)
        if not match:
            return value
        ca, letters, ra, row = match.groups()
        column = column_index_from_string(letters)
        if logical == "MUTASI":
            column = MOVE_COLS[column]
        return f"{ca}{get_column_letter(column)}{ra}{int(row) + SECTIONS[logical].offset}"
    return ":".join(one(value) for value in ref.split(":"))


def formula(value: str, logical: str) -> str:
    """Rewrite references only, leaving names and quoted text intact."""
    leading_equal = value.startswith("=")
    tokens = Tokenizer(value if leading_equal else "=" + value).items
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        reference = token.value
        if "!" in reference:
            source, reference = reference.rsplit("!", 1)
            source = source.strip("'").replace("''", "'")
            if source not in SECTIONS:
                continue
        else:
            source = logical
            if not all(CELL.fullmatch(part) for part in reference.split(":")):
                continue  # Named range, not an A1 reference.
        target = SECTIONS[source].sheet.replace("'", "''")
        # Every reference is explicit. Sections on the same sheet remain distinct.
        token.value = f"'{target}'!{address(source, reference)}"
    result = "".join(token.value for token in tokens)
    return ("=" if leading_equal else "") + result


def destination(logical: str, ref: str) -> str:
    return f"'{SECTIONS[logical].sheet}'!{address(logical, ref)}"


def compact(source: Path, dest: Path) -> Path:
    original = load_workbook(source)
    book = Workbook()
    book.remove(book.active)
    for title in NAV:
        book.create_sheet(title)
    book.properties = copy(original.properties)
    book.properties.title = "Gian Fataya | Konsinyasi Ringkas v2"
    book.properties.description = "Enam sheet; mengikuti MoM/Notulensi kolom Updated, dibaca ulang 7 November 2026; grafik perbandingan dan tren wilayah."
    book.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
    styles = {}

    def copy_style(cell, target):
        if cell.style_id in styles:
            target._style = copy(styles[cell.style_id])
            return
        for attr in ("font", "fill", "border", "alignment", "protection"):
            setattr(target, attr, copy(getattr(cell, attr)))
        target.number_format = cell.number_format
        styles[cell.style_id] = copy(target._style)

    def copy_section(logical, first, last, max_col):
        old = original[logical]
        new = book[SECTIONS[logical].sheet]
        offset = SECTIONS[logical].offset
        for row in old.iter_rows(min_row=first, max_row=last, max_col=max_col):
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                target = new[address(logical, cell.coordinate)]
                target.value = formula(cell.value, logical) if cell.data_type == "f" else cell.value
                copy_style(cell, target)
                if cell.comment:
                    target.comment = copy(cell.comment)
                if cell.hyperlink:
                    link = copy(cell.hyperlink)
                    link.ref = target.coordinate
                    if link.location and "!" in link.location:
                        sheet, ref = link.location.rsplit("!", 1)
                        sheet = sheet.strip("'")
                        if sheet in SECTIONS:
                            link.location = destination(sheet, ref)
                    target._hyperlink = link
        for row, dimension in old.row_dimensions.items():
            if first <= row <= last:
                new.row_dimensions[row + offset].height = dimension.height
                new.row_dimensions[row + offset].hidden = dimension.hidden
        for merged in old.merged_cells.ranges:
            if merged.min_row >= first and merged.max_row <= last:
                new.merge_cells(address(logical, str(merged)))
        for table in old.tables.values():
            # XlsxWriter stores table body formats in dataDxfId rather than in
            # empty Cell records. Materialize them before relocating the table.
            # Source workbook style IDs must never leak into the destination.
            for column in table.tableColumns:
                if column.dataDxfId is None:
                    continue
                dxf = original._differential_styles.styles[column.dataDxfId]
                first_row, last_row = range_boundaries(table.ref)[1], range_boundaries(table.ref)[3]
                for r in range(first_row + 1, last_row + 1):
                    target = new[address(logical, f"{get_column_letter(column.id)}{r}")]
                    for attr in ("font", "fill", "border", "alignment", "protection"):
                        if getattr(dxf, attr, None) is not None:
                            setattr(target, attr, copy(getattr(dxf, attr)))
                    if dxf.numFmt is not None:
                        target.number_format = dxf.numFmt.formatCode
                    # Differential fonts can omit the face and inherit Cambria
                    # when converted. Keep table text consistent with the UI.
                    font = copy(target.font)
                    font.name, font.scheme = "Calibri", None
                    font.sz = font.sz or 11
                    target.font = font
            moved = deepcopy(table)
            # Reordered movement table still covers A:X, with technical fields last.
            moved.ref = address(logical, table.ref)
            if logical == "MUTASI":
                moved.ref = "A6:X2006"
                moved.tableColumns = sorted(moved.tableColumns, key=lambda c: MOVE_COLS[c.id])
                for index, column in enumerate(moved.tableColumns, 1):
                    column.id = index
            for column in moved.tableColumns:
                column.dataDxfId = column.headerRowDxfId = column.totalsRowDxfId = None
            moved.dataDxfId = moved.headerRowDxfId = moved.totalsRowDxfId = None
            if moved.autoFilter:
                moved.autoFilter.ref = moved.ref
            new.add_table(moved)
        for validation in old.data_validations.dataValidation:
            moved = deepcopy(validation)
            moved.sqref = MultiCellRange(" ".join(address(logical, str(r)) for r in validation.sqref.ranges))
            for attr in ("formula1", "formula2"):
                value = getattr(moved, attr)
                if value and not value.startswith('"'):
                    setattr(moved, attr, formula(value, logical))
            new.add_data_validation(moved)
        for rule_range, rules in old.conditional_formatting._cf_rules.items():
            ranges = " ".join(address(logical, str(r)) for r in rule_range.sqref.ranges)
            for rule in rules:
                moved = deepcopy(rule)
                moved.formula = [formula(v, logical) for v in moved.formula]
                new.conditional_formatting.add(ranges, moved)

    # Keep accounting sections, removing duplicate start pages and navigation.
    specs = [
        ("DASHBOARD", 5, 44, 12), ("FORM", 5, 45, 10),
        ("MUTASI", 6, 2006, 24), ("STOK", 6, 4886, 12),
        ("INVOICE", 6, 406, 18), ("PEMBAYARAN", 6, 1006, 10),
        ("CETAK INVOICE", 5, 79, 6),
        ("WILAYAH", 6, 26, 7), ("OUTLET", 6, 67, 9),
        ("PRODUK", 6, 56, 8), ("LOT", 6, 86, 6),
        ("BAHAN", 6, 20, 8), ("REFERENSI", 6, 44, 15),
    ]
    # Additional PO control is distinct from stock validity: split-lot shipments
    # can be entered consecutively without losing the first partial line.
    old = original["MUTASI"]
    from openpyxl.worksheet.table import TableColumn
    old.tables["tblMutasi"].tableColumns.append(TableColumn(id=24, name="Cek minimum PO"))
    old["X6"] = "Cek minimum PO"
    old["X6"]._style = copy(old["R6"]._style)
    for r in range(7, 2007):
        old[f"X{r}"] = (
            f'=IF(AND(A{r}<>"",C{r}="KIRIM",I{r}="AKTIF"),'
            f'IF(SUMIFS($F$7:$F$2006,$C$7:$C$2006,"KIRIM",$D$7:$D$2006,D{r},'
            f'$U$7:$U$2006,U{r},$I$7:$I$2006,"AKTIF")<10,"PO < 10 PCS","OK"),"")'
        )
        old[f"X{r}"]._style = copy(old[f"R{r}"]._style)
        old[f"I{r}"] = "AKTIF"  # Default status; an empty ID remains an unused row.
    for spec in specs:
        copy_section(*spec)
    # Reference lists sit above the old content region; they are moved too.
    for column, last in [("N", 9), ("O", 5)]:
        for r in range(2, last + 1):
            book["DATA USAHA"][address("REFERENSI", f"{column}{r}")] = original["REFERENSI"][f"{column}{r}"].value
    for defined in original.defined_names.values():
        moved = copy(defined)
        moved.attr_text = formula(defined.attr_text, "DASHBOARD")
        book.defined_names.add(moved)

    def band(ws, ref, text, bg=TEAL, fg="FFFFFF", size=11, bold=True):
        ws.merge_cells(ref)
        for row in ws[ref]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell = ws[ref.split(":")[0]]
        cell.value = text
        cell.font = Font(name="Calibri", size=size, color=fg, bold=bold)
        return cell

    def link(ws, cell, label, sheet, anchor="A1", span=None):
        if span:
            ws.merge_cells(f"{cell}:{span}")
        c = ws[cell]
        c.value = label
        c.hyperlink = f"#'{sheet}'!{anchor}"
        c.font = Font(name="Calibri", size=11, bold=True, color=TEAL, underline="single")
        c.alignment = Alignment(vertical="center", wrap_text=True)

    subtitles = {
        "BERANDA": "Konsinyasi ringkas v2  |  Stok • penjualan • invoice wilayah  |  MoM kolom Updated",
        "TRANSAKSI": "Isi A:I sesuai jenis gerakan. J:L otomatis. Detail penagihan/status ada di grup kolom M:X (+).",
        "STOK": "Saldo otomatis sampai tanggal akhir Beranda. Filter lokasi/lot. Sisa stok terbawa ke bulan berikutnya.",
        "TAGIHAN": "Invoice per wilayah berdasarkan JUAL. Pembayaran dan cicilan berada di bagian kedua pada sheet ini.",
        "CETAK INVOICE": "Pilih ID. Cetak/PDF hanya jika SIAP CETAK. Tombol PDF tersedia setelah pemasangan VBA.",
        "DATA USAHA": "Isi sekali sebelum mulai; buka bagian lewat tautan baris 5. Semua pengaturan berada di satu sheet.",
    }
    for ws in book:
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85
        ws.sheet_view.tabSelected = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_properties.tabColor = TEAL if ws.title != "DATA USAHA" else "DFA94A"
        ws.sheet_format.defaultRowHeight = 22
        for i in range(1, 25):
            ws.column_dimensions[get_column_letter(i)].width = 16
        band(ws, "A1:L1", "GIAN FATAYA  /  " + ws.title, NAVY, size=23)
        ws.row_dimensions[1].height = 43
        band(ws, "A2:L2", subtitles[ws.title], "FFFFFF", MUTED, 11, False)
        ws.row_dimensions[2].height = 32
        for i, (title, label) in enumerate(zip(NAV, NAV_LABELS)):
            link(ws, f"{get_column_letter(1+i*2)}4", label, title)
        ws.row_dimensions[4].height = 29
        ws.freeze_panes = "A5"
        ws.sheet_properties.outlinePr.summaryRight = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.oddHeader.center.text = "GIAN FATAYA | Konsinyasi ringkas"
        ws.oddFooter.right.text = "Halaman &P / &N"
        ws.print_title_rows = "1:6"

    ws = book["BERANDA"]
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["K"].width = 21
    ws.column_dimensions["L"].width = 16
    from regional_dashboard import configure as configure_regional_dashboard
    areas = sum(original["WILAYAH"].cell(r, 1).value is not None for r in range(7, 27))
    configure_regional_dashboard(book, areas, band, link)
    band(ws, "A66:L67", "Tanpa makro: isi tabel Transaksi dan Tagihan. Tombol otomatis memerlukan salinan .xlsm dan Modul_GianFataya.bas; lihat panduan pemasangan.", PALE, INK, 11, False)
    band(ws, "A69:L70", "INPUT CEPAT OPSIONAL  /  setelah VBA dipasang", TEAL)
    link(ws, "A72", "Form barang", "BERANDA", "A75")
    link(ws, "E72", "Form invoice", "BERANDA", "A89")
    link(ws, "I72", "Form pembayaran", "BERANDA", "A106")
    # Form sections retain accounting coordinates shifted by 70 rows.
    ws["A86"] = "RETUR hanya untuk kerusakan fisik kemasan; isi alasannya. PO minimal 10 pcs per nomor PO + outlet (boleh beberapa lot)."
    ws["A102"] = "Hanya JUAL aktif yang valid dan belum ditagih. Diskon sekali dari bruto; ongkir ditambahkan setelah diskon."
    ws["A113"] = "Pilih invoice TERBIT, catat cicilan/pelunasan. Gunakan CETAK INVOICE untuk PDF setelah indikator SIAP CETAK."
    for cell in ["A86", "A102", "A113"]:
        ws[cell].alignment = Alignment(wrap_text=True, vertical="top")
    ws.print_area = "A1:L64"
    ws = book["TRANSAKSI"]
    widths = {"A":21,"B":17,"C":15,"D":17,"E":18,"F":9,"G":18,"H":22,"I":18,"J":25,"K":18,"L":18,"M":23,"N":22,"O":14,"P":42}
    for c, width in widths.items():
        ws.column_dimensions[c].width = width
    ws.freeze_panes = "F7"
    ws.column_dimensions.group("M", "X", outline_level=1, hidden=True)
    ws["A5"] = "INPUT A:I"
    ws["D5"] = "KIRIM: wajib PO"
    ws["G5"] = "JUAL: isi harga"
    ws["J5"] = "OTOMATIS J:L"
    link(ws, "M5", "Detail invoice / status / alasan (+)", "TRANSAKSI", "M6")
    from openpyxl.comments import Comment
    ws["H6"].comment = Comment("Minimum 10 pcs per nomor PO dan outlet; beberapa lot dapat memakai PO yang sama. Baris K menandai PO yang belum lengkap.", "Panduan")
    ws["O6"].comment = Comment("AKTIF sudah diisi sebagai default. Buka grup (+) untuk mengganti menjadi VOID; jangan hapus riwayat.", "Panduan")
    ws["P6"].comment = Comment("Alasan wajib untuk retur kerusakan fisik kemasan, rusak dan koreksi TAMBAH/KURANG.", "Panduan")
    ws["M6"].comment = Comment("Isi ID invoice hanya pada JUAL belum ditagih. Buka grup (+), lalu filter wilayah/periode sebelum menautkan.", "Panduan")
    from openpyxl.formatting.rule import CellIsRule
    ws.conditional_formatting.add("K7:K2006", CellIsRule(operator="equal", formula=['"PO < 10 PCS"'], fill=PatternFill("solid", fgColor=GOLD), font=Font(color="80590D")))
    ws.print_area = "A1:L40"
    ws.sheet_view.selection[0].activeCell = "B7"
    ws.sheet_view.selection[0].sqref = "B7"

    ws = book["TAGIHAN"]
    for c, width in {"A":25,"B":18,"C":18,"D":18,"E":18,"F":18,"G":14,"H":19,"I":15,"J":20,"K":18,"L":20,"M":20,"N":20,"O":19,"P":25,"Q":20,"R":40}.items():
        ws.column_dimensions[c].width = width
    link(ws, "A5", "Invoice wilayah (bagian 1)", "TAGIHAN", "A6", "D5")
    link(ws, "E5", "Pembayaran / cicilan (bagian 2)", "TAGIHAN", "A436", "H5")
    link(ws, "I5", "Form otomatis", "BERANDA", "A89", "L5")
    band(ws, "A410:L412", "CARA MANUAL: buat DRAF di tabel atas. Tautkan ID ke JUAL yang sesuai di Transaksi kolom M (buka +). Tempel nilai Bruto Sumber (Q) ke Bruto Snapshot (J), cek OK, lalu TERBIT. Catat penerimaan pada tabel pembayaran di bawah.", PALE, INK, 11, False)
    link(ws, "A414", "Kembali ke invoice", "TAGIHAN", "A6")
    link(ws, "E414", "Ke pembayaran", "TAGIHAN", "A436")
    band(ws, "A432:L433", "02  PEMBAYARAN DAN CICILAN")
    band(ws, "A434:L434", "Isi A:G; H:J otomatis. Pilih invoice TERBIT, nominal positif. Gunakan VOID untuk membatalkan.", PALE, INK, 11, False)
    link(ws, "A435", "Kembali ke invoice", "TAGIHAN", "A6")
    link(ws, "E435", "Buka form pembayaran", "BERANDA", "A106", "H435")
    ws.print_area = "A1:R32"
    ws.freeze_panes = "C5"

    ws = book["DATA USAHA"]
    for c, width in {"A":22,"B":30,"C":18,"D":30,"E":24,"F":26,"G":17,"H":22,"I":18}.items():
        ws.column_dimensions[c].width = width
    for start, end in [(7,26),(35,95),(103,152),(161,240)]:
        for row in ws.iter_rows(min_row=start,max_row=end,max_col=9):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True,vertical="center")
            ws.row_dimensions[row[0].row].height = 40
    ws.freeze_panes = "A6"
    for row in range(7,15):
        ws[f"F{row}"] = "Sumber: MoM Updated; lengkapi PIC dan alamat tagihan"
    for cell, label, anchor in [
        ("A5", "Wilayah", "A6"), ("C5", "Outlet", "A34"), ("E5", "Produk", "A102"),
        ("G5", "Lot / expired", "A160"), ("I5", "HPP acuan", "A248"), ("K5", "Pengaturan", "A268"),
    ]:
        link(ws, cell, label, "DATA USAHA", anchor)
    for ref, text in [
        ("A29:I30", "02  OUTLET  /  isi lokasi aktual; pilih Amanda atau Independen"),
        ("A98:H99", "03  PRODUK  /  isi harga jual, bukan biaya acuan"),
        ("A156:H157", "04  LOT DAN KEDALUWARSA  /  isi dari kemasan aktual"),
        ("A244:H245", "05  HPP BAHAN ACUAN  /  referensi saja, bukan pengelolaan produksi"),
        ("A266:L266", "06  PENGATURAN DAN SUMBER"),
    ]:
        band(ws, ref, text)
    ws["A264"] = "Daftar internal untuk dropdown tersimpan di kolom N:O."
    ws["A274"] = "Peringatan exp (hari)"
    ws["A276"] = "Ambang tinggi / area teratas"
    ws["B276"] = 0.5
    ws["B276"].number_format = "0%"
    ws["B276"].fill = PatternFill("solid", fgColor=BLUE)
    ws["B276"].font = Font(name="Calibri", size=11, color="145DA0")
    threshold = DataValidation(type="decimal", operator="between", formula1=0.01, formula2=1, allow_blank=False)
    threshold.errorTitle, threshold.error = "Ambang 1–100%", "Masukkan persentase antara 1% dan 100%."
    threshold.showErrorMessage = True
    ws.add_data_validation(threshold)
    threshold.add("B276")
    ws["B276"].comment = Comment("Default desain: 50% dari bruto wilayah tertinggi pada periode yang dipilih. Bukan target yang diberikan oleh mitra. DATAR berarti bruto nol.", "Panduan")
    ws["A282"] = "Delapan wilayah resmi mengikuti kolom Updated. Daftarkan toko independen dengan kelompok Independen dan wilayah penagihan yang disepakati."
    ws["A285"] = "Konsinyasi: kirim bukan penjualan; hanya barang terjual ditagih kepada PJ area."
    ws["A288"] = "PO minimal 10 pcs per nomor PO + outlet; beberapa baris lot boleh digabung. Retur hanya untuk kerusakan fisik kemasan, wajib alasan."
    ws["A291"] = "Lengkapi harga jual, outlet, PIC, stok awal, kedaluwarsa, diskon/ongkir, informasi pembayaran dan format invoice yang disetujui."
    ws["A294"] = "Diskon sekali dari bruto invoice; ongkir ditagih ditambah sesudahnya. Ongkir aktual pengiriman dicatat terpisah."
    ws["A297"] = "Expired tetap stok fisik sampai tindakan tercatat; tidak boleh dikirim/dijual melewati tanggal kedaluwarsa."
    ws["A300"] = "Harga jual dan bruto invoice disimpan sebagai nilai tetap. Jangan ubah penjualan yang sudah ditagih; gunakan prosedur pembatalan."
    ws["A303"] = "ID kosong tidak dihitung. Kapasitas tetap: 2.000 gerakan, 400 invoice, 1.000 pembayaran. Jangan menambah baris di luar tabel."
    # Old informational notes may be merged at these starts.
    ws["A277"] = "MoM • kolom Updated • dibaca 4 November 2026"
    ws.column_dimensions.group("N", "O", outline_level=1, hidden=True)
    for ref in ("A268", "A270", "A272", "A274", "A276"):
        ws[ref].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[ws[ref].row].height = 34
    for ref in ("B270:H271", "B272:H273"):
        ws.merge_cells(ref)
        ws[ref.split(":")[0]].alignment = Alignment(wrap_text=True,vertical="center")
    ws.print_area = "A1:I26"

    ws = book["STOK"]
    for c, width in {"A":20,"B":30,"C":18,"D":22,"E":19,"F":28,"G":20,"H":17,"I":17,"J":13,"K":24,"L":16}.items():
        ws.column_dimensions[c].width = width
    ws.freeze_panes = "C7"
    link(ws, "A5", "Ubah tanggal laporan di Beranda", "BERANDA", "B5", "D5")
    ws.print_area = "A1:L35"

    ws = book["CETAK INVOICE"]
    # Invoice printout omits navigation and selector, as before.
    ws.unmerge_cells("A1:L1")
    ws.unmerge_cells("A2:L2")
    band(ws, "A1:F1", "GIAN FATAYA  /  CETAK INVOICE", NAVY, size=23)
    band(ws, "A2:F2", subtitles["CETAK INVOICE"], "FFFFFF", MUTED, 11, False)
    for i, (title, label) in enumerate(zip(NAV, ["Beranda", "Transaksi", "Stok", "Tagihan", "Cetak", "Data usaha"]), 1):
        link(ws, f"{get_column_letter(i)}4", label, title)
    for cell in ("G4", "I4", "K4"):
        ws[cell].value, ws[cell].hyperlink = None, None
    for c, width in {"A":14,"B":24,"C":27,"D":14,"E":21,"F":22}.items():
        ws.column_dimensions[c].width = width
    ws["A75"] = '=IF(\'DATA USAHA\'!B270="","Informasi pembayaran belum diisi. Lengkapi Data Usaha sebelum mengirim invoice.","Pembayaran: "&\'DATA USAHA\'!B270)'
    ws["A78"] = "Harga rata-rata = bruto / qty bila ada perubahan harga. Rincian tersedia di Transaksi dengan filter ID invoice (kolom M)."
    ws.page_setup.orientation = "portrait"
    ws.print_area = "A7:F79"
    ws.print_title_rows = "7:16"
    ws.freeze_panes = "A5"
    book.active = 0
    book["BERANDA"].sheet_view.tabSelected = True
    dest.parent.mkdir(parents=True, exist_ok=True)
    book.save(dest)
    return dest
