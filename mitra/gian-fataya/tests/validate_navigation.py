"""Regression checks for Excel's 'Cannot open the specified file' navigation bug."""
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.cell import quote_sheetname
from openpyxl.worksheet.hyperlink import Hyperlink

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from workbook_links import financial_fingerprint, internal_link, repair_navigation, validate_navigation

TABS = ["BERANDA", "TRANSAKSI", "STOK", "TAGIHAN", "CETAK INVOICE", "DATA USAHA"]
DELIVERED = [ROOT / "output/POS_Gian_Fataya.xlsx", ROOT / "demo/SIMULASI_Grafik_Wilayah.xlsx",
             ROOT / "latihan/POS_Gian_Fataya_DUMMY.xlsx", ROOT / "latihan/template_kosong/POS_Gian_Fataya.xlsx"]
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def fixture(broken):
    book = Workbook()
    book.remove(book.active)
    for title in TABS:
        sheet = book.create_sheet(title)
        for i, target in enumerate(TABS):
            cell = sheet.cell(4, 1 + 2 * i, target)
            if broken:
                cell.hyperlink = Hyperlink(ref=cell.coordinate,
                    target="../../../workspace/Social-Project-BSI/mitra/gian-fataya/output/POS_Gian_Fataya.xlsx",
                    location=f"{quote_sheetname(target.replace(' ', '%20'))}!A1")
            else:
                internal_link(cell, target)
    book["BERANDA"]["A10"] = "Web source"
    book["BERANDA"]["A10"].hyperlink = "https://docs.google.com/spreadsheets/d/source/edit#gid=52890219"
    book["BERANDA"]["A11"] = "Guide PDF"
    book["BERANDA"]["A11"].hyperlink = "PANDUAN_LATIHAN_GIAN_FATAYA.pdf"
    for r, qty in [(20, 30), (21, 40), (22, 50)]:
        book["BERANDA"].cell(r, 1, qty)
    book["BERANDA"]["B20"] = "=SUM(A20:A22)"
    chart = BarChart()
    chart.add_data(Reference(book["BERANDA"], min_col=1, min_row=20, max_row=22))
    book["BERANDA"].add_chart(chart, "D20")
    return book


def check_navigation(path):
    validate_navigation(path)
    # Read native bookmarks independently, without the repair routine's parser.
    book = load_workbook(path)
    for sheet in book:
        columns = [1, 2, 3, 4, 5, 6] if sheet.title == "CETAK INVOICE" else [1, 3, 5, 7, 9, 11]
        for col, target in zip(columns, TABS):
            link = sheet.cell(4, col).hyperlink
            assert link and link.target is None and link.id is None, f"{sheet.title}!{link.ref if link else col}: external navigation"
            assert link.location == f"{quote_sheetname(target)}!A1", (sheet.title, col, link.location)
    for sheet, ref, target in [("BERANDA", "I16", "'TAGIHAN'!A436"), ("DATA USAHA", "C5", "'DATA USAHA'!A34")]:
        assert book[sheet][ref].hyperlink.location == target


class NavigationTests(unittest.TestCase):
    def test_fresh_build_writes_native_navigation(self):
        from build_workbook import build
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "New workbook.xlsx"
            build(path)
            check_navigation(path)

    def test_all_delivered_files_have_native_bookmarks(self):
        for path in DELIVERED:
            with self.subTest(file=str(path.relative_to(ROOT))):
                check_navigation(path)

    def test_legacy_file_references_are_fixed_without_changing_data(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "Example.xlsx"
            fixture(True).save(path)
            original = financial_fingerprint(path)
            with self.assertRaises(AssertionError):
                validate_navigation(path)
            result = repair_navigation(path)
            self.assertEqual(result["repaired"], 36)
            self.assertEqual(validate_navigation(path)["internal"], 36)
            self.assertEqual(financial_fingerprint(path), original)
            book = load_workbook(path)
            self.assertEqual(book["BERANDA"]["A10"].hyperlink.target, "https://docs.google.com/spreadsheets/d/source/edit#gid=52890219")
            self.assertEqual(book["BERANDA"]["A11"].hyperlink.target, "PANDUAN_LATIHAN_GIAN_FATAYA.pdf")
            self.assertEqual(repair_navigation(path)["repaired"], 0)

    def test_genuine_external_workbook_link_is_not_retargeted(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "Example.xlsx"
            book = fixture(False)
            book["BERANDA"]["A12"] = "External workbook"
            book["BERANDA"]["A12"].hyperlink = Hyperlink(ref="A12", target="Different.xlsx", location="'STOK'!A1")
            book.save(path)
            repair_navigation(path)
            self.assertEqual(load_workbook(path)["BERANDA"]["A12"].hyperlink.target, "Different.xlsx")

    def test_bookmarks_survive_recalculation_spaces_apostrophes_and_rename(self):
        with tempfile.TemporaryDirectory(prefix="gian-navigation-") as td:
            folder = Path(td)
            raw = folder / "Example.xlsx"
            book = fixture(False)
            book.create_sheet("Area O'Brien")
            book["BERANDA"]["A12"] = "Apostrophe destination"
            internal_link(book["BERANDA"]["A12"], "Area O'Brien", "B8")
            book.save(raw)
            self.assertEqual(validate_navigation(raw)["internal"], 37)
            out = folder / "calculated"
            out.mkdir()
            result = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(raw)],
                                    capture_output=True, text=True, timeout=60)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            recalculated = out / raw.name
            repair_navigation(recalculated)
            self.assertEqual(validate_navigation(recalculated)["internal"], 37)
            self.assertEqual(load_workbook(recalculated)["BERANDA"]["A12"].hyperlink.location, "'Area O''Brien'!B8")
            moved = folder / "Folder pengguna dengan spasi" / "Latihan saya.xlsx"
            moved.parent.mkdir()
            shutil.copyfile(recalculated, moved)
            self.assertEqual(validate_navigation(moved)["internal"], 37)

    def test_zipped_workbooks_are_fixed_too(self):
        with tempfile.TemporaryDirectory() as td:
            for archive, name in [("output/Paket_POS_Gian_Fataya.zip", "POS_Gian_Fataya.xlsx"),
                                  ("output/Paket_POS_Gian_Fataya.zip", "contoh/SIMULASI_Grafik_Wilayah.xlsx"),
                                  ("latihan/Paket_Latihan_Gian_Fataya_DUMMY.zip", "POS_Gian_Fataya_DUMMY.xlsx"),
                                  ("latihan/Paket_Latihan_Gian_Fataya_DUMMY.zip", "template_kosong/POS_Gian_Fataya.xlsx")]:
                with self.subTest(archive=archive, file=name):
                    path = Path(td) / Path(name).name
                    with ZipFile(ROOT / archive) as z:
                        path.write_bytes(z.read(name))
                    check_navigation(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
