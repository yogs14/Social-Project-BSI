"""Build and verify a complete, isolated dummy workbook and its training pack."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

import fitz
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "output/POS_Gian_Fataya.xlsx"
DEST = ROOT / "latihan"
NAME = "POS_Gian_Fataya_DUMMY.xlsx"
GUIDE = "PANDUAN_LATIHAN_GIAN_FATAYA"
DATA = json.loads((ROOT / "data/latihan_dummy.json").read_text())
CHECKS = []


def fix_guide_link(path):
    """Keep the PDF link relative to the delivered file, preserving caches.

    LibreOffice rebases links from raw/ to calculated/. Copying that file into
    the delivery directory otherwise leaves a link back into the temporary run.
    Edit only the relationship, without rewriting any formulas or chart caches.
    """
    sibling = GUIDE + ".pdf"
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    part = "xl/worksheets/_rels/sheet1.xml.rels"
    with ZipFile(path) as source:
        sheet = ET.fromstring(source.read("xl/worksheets/sheet1.xml"))
        link = sheet.find("s:hyperlinks/s:hyperlink[@ref='A64']", ns)
        assert link is not None
        target_id = link.attrib[rel_id]
        relationships = ET.fromstring(source.read(part))
        target = next(r for r in relationships if r.attrib["Id"] == target_id)
        target.set("Target", sibling)
        replacement = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
        staged = path.with_suffix(".links.xlsx")
        with ZipFile(staged, "w", ZIP_DEFLATED) as out:
            for info in source.infolist():
                out.writestr(info, replacement if info.filename == part else source.read(info.filename))
    staged.replace(path)
    with ZipFile(path) as source:
        fixed = ET.fromstring(source.read(part))
        actual = next(r.attrib["Target"] for r in fixed if r.attrib["Id"] == target_id)
    equal("PDF guide hyperlink points to sibling file", actual, sibling)


def day(value):
    return dt.datetime.fromisoformat(value)


def equal(label, actual, expected):
    assert actual == expected, f"{label}: {actual!r} != {expected!r}"
    CHECKS.append({"check": label, "result": "PASS", "actual": actual})


def put(sheet, row, values):
    for column, value in values.items():
        sheet.cell(row, column, value)


def filled(stage="final"):
    w = load_workbook(SOURCE)
    assert w["TRANSAKSI"].tables["tblMutasi"].ref == "A6:X2006"
    assert not any(w["TRANSAKSI"].cell(r, 1).value for r in range(7, 2007))
    assert not any(w["TAGIHAN"].cell(r, 1).value for r in range(7, 407))
    assert not any(w["TAGIHAN"].cell(r, 1).value for r in range(437, 1437))
    w.properties.title = "Gian Fataya — DUMMY alur lengkap"
    w.properties.subject = DATA["label"]
    w.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
    for s in w:
        s["A1"] = "GIAN FATAYA / DUMMY / " + s.title
        s["A2"] = "DUMMY | " + s["A2"].value
    master, home, moves, bills = w["DATA USAHA"], w["BERANDA"], w["TRANSAKSI"], w["TAGIHAN"]
    for row, code, name, pic, address in DATA["regions"]:
        put(master, row, {3: pic, 4: address, 6: "DUMMY — hanya untuk latihan"})
    put(master, 35, {2: "Gudang DUMMY", 4: "Pemilik DUMMY", 5: "Alamat gudang DUMMY"})
    for row, code, name, region, group in DATA["outlets"]:
        put(master, row, {1: code, 2: name, 3: region, 4: "PIC DUMMY", 5: "Alamat outlet DUMMY", 7: "AKTIF", 9: group})
    put(master, 103, {2: "Pastel DUMMY (pouch)", 4: DATA["price"], 7: "Harga latihan, bukan harga jual mitra."})
    put(master, 161, {1: DATA["lot"], 2: DATA["sku"], 4: day(DATA["expiry"]), 5: "Batch DUMMY"})
    master["B268"] = "Pemilik DUMMY"
    master["B270"] = "SIMULASI — tidak ada rekening atau permintaan transfer nyata"
    master["B272"] = "Alamat pengirim DUMMY — latihan saja"
    master["B276"] = 0.5
    home["A1"] = "GIAN FATAYA / LATIHAN DATA DUMMY"
    home["A2"] = "HASIL CONTOH | Data fiktif. Mulai latihan ulang dari template_kosong; ikuti PANDUAN_LATIHAN_GIAN_FATAYA.pdf."
    home["B5"], home["E5"] = day(DATA["period_start"]), day(DATA["period_end"])
    if stage == "october":
        home["E5"] = day("2026-10-31")
    if stage == "november":
        home["B5"] = day("2026-11-01")
    home["I18"] = "Jawa Timur"
    home["A64"] = "Buka panduan latihan (PDF)"
    home["A64"].hyperlink = GUIDE + ".pdf"
    limit = 5 if stage == "send" else 9 if stage == "sale" else 11
    invoice_stage = stage not in {"send", "sale"}
    for movement in DATA["movements"][:limit]:
        r = movement["row"]
        values = {1: movement["id"], 2: day(movement["date"]), 3: movement["kind"],
                  4: movement["location"], 5: DATA["lot"], 6: movement["qty"],
                  15: "AKTIF", 16: "DUMMY — bukan transaksi nyata"}
        if movement["kind"] == "JUAL":
            values[7] = DATA["price"]
            if invoice_stage:
                values[13] = movement["invoice"]
        if movement["kind"] == "KIRIM":
            values.update({8: movement["po"], 9: movement["freight"], 14: "Ekspedisi DUMMY"})
        put(moves, r, values)
    if invoice_stage:
        for invoice in DATA["invoices"]:
            put(bills, invoice["row"], {1: invoice["id"], 2: day(invoice["date"]), 3: invoice["region"],
                4: day(invoice["start"]), 5: day(invoice["end"]), 6: day(invoice["due"]), 7: invoice["discount"],
                8: invoice["freight"], 9: "DRAF" if stage == "draft" else "TERBIT", 10: invoice["gross"],
                18: "DUMMY — bukan tagihan nyata"})
    if stage not in {"send", "sale", "draft", "issue"}:
        for payment in DATA["payments"]:
            put(bills, payment["row"], {1: payment["id"], 2: day(payment["date"]), 3: payment["invoice"],
                4: "Transfer", 5: payment["amount"], 6: "AKTIF", 7: "DUMMY — bukan penerimaan nyata"})
    printed = w["CETAK INVOICE"]
    printed["B5"] = "INV-D-JT-1026" if invoice_stage else None
    printed["A7"] = "GIAN FATAYA — DUMMY"
    printed["A8"] = "SIMULASI | BUKAN TAGIHAN NYATA"
    for r in range(18, 67):
        printed.row_dimensions[r].hidden = True
    # A useful initial stock view: the one populated lot, at five locations.
    stock = w["STOK"]
    stock.tables["tblStok"].autoFilter.add_filter_column(3, [DATA["lot"]])
    visible_stock = {7 + i * 80 for i in range(5)}
    for r in range(7, 4887):
        stock.row_dimensions[r].hidden = r not in visible_stock
    stock.print_area = "A1:L327"
    stock["A2"] = "DUMMY | Filter Lot = LOT-DUMMY. Saldo otomatis sampai E5 di Beranda. Data → Clear untuk membuka filter."
    w.active = 0
    for s in w:
        s.sheet_view.tabSelected = s.title == "BERANDA"
    return w


def recalculate(sources, folder):
    result = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(folder),
                             *[str(p) for p in sources]], capture_output=True, text=True, timeout=180)
    assert result.returncode == 0, result.stdout + result.stderr
    assert all((folder / p.name).exists() for p in sources), result.stdout + result.stderr


def stock_values(w):
    return {w["STOK"].cell(r, 1).value: w["STOK"].cell(r, 10).value for r in [7, 87, 167, 247, 327]}


def validate(books):
    for name, w in books.items():
        errors = [f"{s.title}!{c.coordinate}={c.value}" for s in w for row in s for c in row if c.data_type == "e"]
        equal(name + ": no formula errors", errors, [])
    home = books["final"]["BERANDA"]
    equal("six visible sheets", [(s.title, s.sheet_state) for s in books["final"]],
          [(n, "visible") for n in ["BERANDA", "TRANSAKSI", "STOK", "TAGIHAN", "CETAK INVOICE", "DATA USAHA"]])
    for cell, expected in {"A8": 590000, "E8": 580000, "I8": 490000, "A13": 90000, "E13": 60, "I13": 81}.items():
        equal("final dashboard " + cell, home[cell].value, expected)
    for cell in ["I57", "C60", "G60", "K60", "C62", "G62", "K62"]:
        equal("final alert clear " + cell, home[cell].value, 0)
    equal("final stock balances", stock_values(books["final"]),
          {"GDG": 60, "D-JT01": 50, "D-JT02": 6, "D-JTG01": 10, "D-MA01": 15})
    equal("stock conservation: 200 - 59 sold", sum(stock_values(books["final"]).values()), 141)
    equal("regional quantity", [home[f"C{r}"].value for r in range(37, 45)], [44, 10, 5, 0, 0, 0, 0, 0])
    equal("regional gross", [home[f"D{r}"].value for r in range(37, 45)], [440000, 100000, 50000, 0, 0, 0, 0, 0])
    equal("regional classification", [home[f"H{r}"].value for r in range(37, 45)], ["TINGGI", "SEDANG", "SEDANG"] + ["DATAR"] * 5)
    equal("trend October", home["O16"].value, 400000)
    equal("trend November through 7th", home["O17"].value, 40000)
    equal("both chart types", [type(c).__name__ for c in home._charts], ["BarChart", "AreaChart"])
    equal("chart cached high region value", home._charts[0].series[0].val.numRef.numCache.pt[0].v, 440000)
    equal("trend cache has twelve months", home._charts[1].series[0].val.numRef.numCache.ptCount, 12)
    equal("final invoice totals", [books["final"]["TAGIHAN"][f"L{r}"].value for r in range(7, 11)], [380000, 110000, 50000, 40000])
    equal("final payments by invoice", [books["final"]["TAGIHAN"][f"M{r}"].value for r in range(7, 11)], [380000, 60000, 50000, 0])
    equal("final outstanding by invoice", [books["final"]["TAGIHAN"][f"N{r}"].value for r in range(7, 11)], [0, 50000, 0, 40000])
    equal("invoice payment states", [books["final"]["TAGIHAN"][f"O{r}"].value for r in range(7, 11)], ["LUNAS", "SEBAGIAN", "LUNAS", "BELUM BAYAR"])
    equal("discount is ten percent", books["final"]["TAGIHAN"]["K7"].value, 40000)
    equal("gross discount freight reconciliation", 590000 - 40000 + 30000, home["E8"].value)
    equal("invoice less receipts equals receivables", home["E8"].value - home["I8"].value, home["A13"].value)
    for r in range(7, 18):
        equal(f"movement {r} valid", books["final"]["TRANSAKSI"][f"J{r}"].value, "OK")
    for r in range(7, 11):
        equal(f"invoice {r} reconciles to source", books["final"]["TAGIHAN"][f"P{r}"].value, "OK")
    for r in range(437, 441):
        equal(f"payment {r} valid", books["final"]["TAGIHAN"][f"I{r}"].value, "OK")
    equal("shipping-only warehouse stock", books["send"]["BERANDA"]["E13"].value, 100)
    equal("shipping-only outlet stock", books["send"]["BERANDA"]["I13"].value, 100)
    equal("shipping not revenue", books["send"]["BERANDA"]["A8"].value, 0)
    equal("October sale before invoice", books["sale"]["BERANDA"]["A8"].value, 550000)
    equal("unbilled sale remains visible", books["sale"]["BERANDA"]["K62"].value, 550000)
    equal("unbilled sale not an invoice", books["sale"]["BERANDA"]["E8"].value, 0)
    equal("draft not counted as issued", books["draft"]["BERANDA"]["E8"].value, 0)
    equal("draft cannot print", books["draft"]["CETAK INVOICE"]["A14"].value, "BUKAN INVOICE TERBIT")
    equal("issued before payment receivable", books["issue"]["BERANDA"]["A13"].value, 580000)
    equal("issued before payment has no receipts", books["issue"]["BERANDA"]["I8"].value, 0)
    for stage, expected in [("october", [550000, 540000, 100000, 440000, 100, 45]),
                            ("november", [40000, 40000, 390000, 90000, 60, 81])]:
        equal(stage + " view", [books[stage]["BERANDA"][c].value for c in ["A8", "E8", "I8", "A13", "E13", "I13"]], expected)
    printed = books["final"]["CETAK INVOICE"]
    equal("print ready", printed["A14"].value, "SIAP CETAK")
    equal("two outlets on one area invoice", printed["D17"].value, 40)
    equal("printed total", printed["F71"].value, 380000)
    equal("printed paid", printed["F72"].value, 380000)
    equal("printed outstanding", printed["F73"].value, 0)
    for r in [7, 87, 167, 247, 327]:
        equal("populated stock row visible " + str(r), books["final"]["STOK"].row_dimensions[r].hidden, False)


def render_review(source, temp):
    w = load_workbook(source)
    for s in w:
        s.print_area = "A1:L1"
        s.print_title_rows = None
        s.page_setup.fitToWidth = s.page_setup.fitToHeight = 1
        s.oddHeader.center.text = ""
        s.oddFooter.right.text = ""
    home = w["BERANDA"]
    home.print_area, home.page_setup.orientation, home.page_setup.paperSize = "A1:L44", "portrait", home.PAPERSIZE_A3
    invoice = w["CETAK INVOICE"]
    invoice.print_area, invoice.page_setup.orientation, invoice.page_setup.paperSize = "A7:F79", "portrait", invoice.PAPERSIZE_A4
    review = temp / "Review_Dummy.xlsx"
    w.save(review)
    result = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(review)],
                            capture_output=True, text=True, timeout=120)
    assert result.returncode == 0 and review.with_suffix(".pdf").exists(), result.stdout + result.stderr
    with fitz.open(review.with_suffix(".pdf")) as pdf:
        first = pdf[0]
        assert "LATIHAN DATA DUMMY" in first.get_text()
        bounds = fitz.Rect()
        for block in first.get_text("blocks"):
            bounds |= fitz.Rect(block[:4])
        first.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), clip=(bounds + (-12, -12, 12, 12)) & first.rect).save(DEST / "gambar/01_Beranda_Dummy.png")
        # A second crop concentrates on controls and both charts in the guide.
        blocks = first.get_text("blocks")
        top = min(b[1] for b in blocks if "PERBANDINGAN PENJUALAN" in b[4]) - 6
        bottom = min(b[1] for b in blocks if "RINCIAN PER WILAYAH" in b[4]) - 3
        first.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), clip=fitz.Rect(bounds.x0, top, bounds.x1, bottom)).save(DEST / "gambar/02_Grafik_Dummy.png")
        pages = [p.number for p in pdf if "INV-D-JT-1026" in p.get_text() and "SIAP CETAK" in p.get_text()]
        equal("one-page invoice PDF", len(pages), 1)
        page = pdf[pages[0]]
        equal("PDF has dummy label", "BUKAN TAGIHAN NYATA" in page.get_text(), True)
        for expected in ["380,000", "400,000", "40,000", "20,000"]:
            equal("invoice PDF contains " + expected, expected in page.get_text(), True)
        invoice_bounds = fitz.Rect()
        for block in page.get_text("blocks"):
            invoice_bounds |= fitz.Rect(block[:4])
        page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), clip=(invoice_bounds + (-12, -12, 12, 12)) & page.rect).save(DEST / "gambar/03_Invoice_Dummy.png")
        export = fitz.open()
        export.insert_pdf(pdf, from_page=pages[0], to_page=pages[0])
        export.save(DEST / "CONTOH_Invoice_DUMMY_Jatim.pdf")
        export.close()


def main():
    fingerprint = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    DEST.mkdir(exist_ok=True)
    (DEST / "gambar").mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gian-full-demo-") as temp:
        temp = Path(temp)
        raw, calculated = temp / "raw", temp / "calculated"
        raw.mkdir(); calculated.mkdir()
        files = []
        stages = ["send", "sale", "draft", "issue", "final", "october", "november"]
        for stage in stages:
            path = raw / (stage + ".xlsx")
            filled(stage).save(path)
            files.append(path)
        print("Recalculating seven workflow checkpoints...", flush=True)
        recalculate(files, calculated)
        books = {stage: load_workbook(calculated / (stage + ".xlsx"), data_only=True) for stage in stages}
        validate(books)
        shutil.copyfile(calculated / "final.xlsx", DEST / NAME)
        fix_guide_link(DEST / NAME)
        print(f"Workflow checks: {len(CHECKS)} passed. Rendering dashboard and invoice...", flush=True)
        render_review(DEST / NAME, temp)
    equal("operational template unchanged", hashlib.sha256(SOURCE.read_bytes()).hexdigest(), fingerprint)
    (DEST / "template_kosong").mkdir(exist_ok=True)
    shutil.copyfile(SOURCE, DEST / "template_kosong/POS_Gian_Fataya.xlsx")
    from training_guide import build as build_guide
    build_guide(DEST, DATA)
    report = {"label": DATA["label"], "passed": len(CHECKS), "failed": 0,
              "validation": "Seven checkpoints recalculated with LibreOffice; actual Excel charts and invoice rendered to PDF",
              "source_operational_file_unchanged": True, "vba_executed": False, "excel_desktop_tested": False, "checks": CHECKS}
    (DEST / "hasil_validasi_latihan.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    (DEST / "MULAI_DI_SINI.txt").write_text(
        "GIAN FATAYA — PAKET LATIHAN DATA DUMMY\n\n"
        "1. Ekstrak seluruh ZIP ke satu folder.\n"
        f"2. Buka {GUIDE}.pdf, mulai halaman pertama.\n"
        f"3. {NAME} sudah terisi: gunakan untuk melihat contoh hasil akhir.\n"
        "4. Untuk latihan dari nol, buka template_kosong/POS_Gian_Fataya.xlsx,\n"
        "   Save As LATIHAN_SAYA.xlsx, lalu ikuti langkah dalam panduan.\n"
        "   Jangan mengetik ulang transaksi di file DUMMY yang sudah terisi.\n"
        "5. Latihan memakai input tabel dan ekspor PDF manual; tidak perlu VBA.\n"
        "6. Semua nama outlet, PIC, harga, tanggal kedaluwarsa, transaksi,\n"
        "   ongkir dan pembayaran DUMMY adalah contoh fiktif, bukan data mitra.\n"
        "   Delapan nama wilayah mengikuti MoM.\n\n"
        "Hasil: bruto Rp590.000; tagihan Rp580.000; pembayaran Rp490.000;\n"
        "piutang Rp90.000; stok gudang 60 dan stok outlet 81 pouch.\n"
        "Contoh invoice PDF DUMMY Jatim: tagihan Rp380.000, sudah lunas.\n"
        "Rumus dan render diuji di LibreOffice; Excel desktop/VBA belum dieksekusi.\n", encoding="utf-8")
    package()
    print(DEST / NAME)
    print(DEST / (GUIDE + ".pdf"))
    print(f"{len(CHECKS)} checks passed; ZIP verified; template unchanged.")


def package():
    files = sorted(p for p in DEST.rglob("*") if p.is_file() and p.suffix != ".zip" and p.name != "SHA256.json")
    manifest = {str(p.relative_to(DEST)): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    (DEST / "SHA256.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    package = DEST / "Paket_Latihan_Gian_Fataya_DUMMY.zip"
    with ZipFile(package, "w", ZIP_DEFLATED) as z:
        for p in files + [DEST / "SHA256.json"]:
            z.write(p, str(p.relative_to(DEST)))
    with ZipFile(package) as z:
        assert z.testzip() is None
        for path, expected in manifest.items():
            assert hashlib.sha256(z.read(path)).hexdigest() == expected, path


if __name__ == "__main__":
    main()
