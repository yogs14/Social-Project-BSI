"""Render the actual dashboard with clearly labelled synthetic regional sales."""
from pathlib import Path
import datetime as dt
import hashlib
import json
import subprocess
import tempfile

import fitz
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties
from workbook_links import repair_navigation, validate_navigation

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "output/POS_Gian_Fataya.xlsx"
DEST = ROOT / "demo"
NAME = "SIMULASI_Grafik_Wilayah.xlsx"


def convert(source, folder, extension):
    result = subprocess.run(["soffice", "--headless", "--convert-to", extension,
                             "--outdir", str(folder), str(source)], capture_output=True, text=True, timeout=120)
    target = folder / (source.stem + "." + extension)
    assert result.returncode == 0 and target.exists(), result.stdout + result.stderr
    return target


def main():
    fingerprint = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    w = load_workbook(SOURCE)
    assert not any(w["TRANSAKSI"].cell(r, 1).value for r in range(7, 2007)), "Use the empty template only."
    master, moves, home = w["DATA USAHA"], w["TRANSAKSI"], w["BERANDA"]
    master["D103"] = 10000
    master["B103"] = "Pastel SIMULASI"
    for c, v in {1: "LOT-GRAFIK", 2: "GF-PST-001", 4: dt.datetime(2027, 1, 31), 5: "DATA SINTETIS"}.items():
        master.cell(161, c, v)
    quantities = [[25, 35, 30, 45, 65, 80], [20, 25, 25, 35, 45, 60],
                  [10, 15, 20, 15, 25, 35], [15, 12, 18, 20, 25, 30],
                  [8, 12, 10, 18, 20, 22], [5, 8, 12, 10, 13, 15],
                  [2, 5, 8, 6, 10, 8], [0, 0, 0, 0, 0, 0]]
    next_row = 7

    def movement(date, kind, location, qty):
        nonlocal next_row
        values = {1: f"SIM-{next_row}", 2: date, 3: kind, 4: location, 5: "LOT-GRAFIK", 6: qty,
                  15: "AKTIF", 16: "SIMULASI GRAFIK — BUKAN TRANSAKSI NYATA"}
        if kind == "JUAL":
            values[7] = 10000
        if kind == "KIRIM":
            values[8] = f"PO-SIM-{next_row}"
        for c, v in values.items():
            moves.cell(next_row, c, v)
        next_row += 1

    movement(dt.datetime(2026, 5, 31), "AWAL", "GDG", 4000)
    for area, monthly in enumerate(quantities, 1):
        code = f"SIM-W{area:02}"
        for c, v in {1: code, 2: "Outlet SIMULASI", 3: f"W{area:02}", 4: "PIC SIMULASI", 5: "Alamat simulasi",
                     7: "AKTIF", 9: "Independen" if area == 2 else "Amanda"}.items():
            master.cell(35 + area, c, v)
        movement(dt.datetime(2026, 6, 1), "KIRIM", code, 500)
        for month, qty in enumerate(monthly, 6):
            if qty:
                movement(dt.datetime(2026, month, 5), "JUAL", code, qty)
    home["A1"] = "GIAN FATAYA / SIMULASI GRAFIK"
    home["A2"] = "CONTOH DATA FIKTIF — bukan performa nyata Gian Fataya. Ubah periode B5/E5 atau pilihan wilayah I18 untuk mencoba grafik."
    home["B5"], home["E5"] = dt.datetime(2026, 11, 1), dt.datetime(2026, 11, 7)
    w.properties.title = "SIMULASI grafik wilayah — bukan data usaha nyata"
    w.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
    w.active = 0
    DEST.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gian-chart-demo-") as temp:
        temp = Path(temp)
        raw = temp / NAME
        w.save(raw)
        result = convert(raw, DEST, "xlsx")
        repair_navigation(result)
        validate_navigation(result)
        cached = load_workbook(result, data_only=True)
        h = cached["BERANDA"]
        assert [h[f"D{r}"].value for r in range(37, 45)] == [800000, 600000, 350000, 300000, 220000, 150000, 80000, 0]
        assert [h[f"H{r}"].value for r in range(37, 45)] == ["TINGGI"] * 2 + ["SEDANG"] * 5 + ["DATAR"]
        assert [h[f"O{r}"].value for r in range(12, 18)] == [250000, 350000, 300000, 450000, 650000, 800000]
        assert all(cached["TRANSAKSI"][f"J{r}"].value == "OK" for r in range(7, next_row))
        assert len(h._charts) == 2
        # Review copy only: export the real sheet, with a short print region.
        review = load_workbook(result)
        for sheet in review:
            sheet.print_area = "A1:L1"
            sheet.print_title_rows = None
            sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
        page = review["BERANDA"]
        page.print_area = "A1:L44"
        page.page_setup.paperSize = page.PAPERSIZE_A3
        page.page_setup.orientation = "portrait"
        page.oddHeader.center.text = ""
        page.oddFooter.right.text = ""
        preview_source = temp / "Pratinjau_Grafik_Wilayah.xlsx"
        review.save(preview_source)
        pdf = convert(preview_source, temp, "pdf")
        document = fitz.open(pdf)
        first = document[0]
        assert "SIMULASI GRAFIK" in first.get_text()
        assert "PERBANDINGAN PENJUALAN" in first.get_text()
        bounds = fitz.Rect(first.get_text("blocks")[0][:4])
        for block in first.get_text("blocks")[1:]:
            bounds |= fitz.Rect(block[:4])
        bounds = (bounds + (-12, -12, 12, 12)) & first.rect
        first.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), clip=bounds).save(DEST / "Pratinjau_Grafik_Wilayah.png")
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest() == fingerprint
    (DEST / "hasil_validasi_grafik.json").write_text(json.dumps({
        "label": "SIMULASI — BUKAN PERFORMA NYATA", "charts": 2, "regions": 8, "trend_months": 12,
        "high": 2, "medium": 5, "flat": 1, "source_operational_file_unchanged": True,
        "validation": "Formulas and chart caches recalculated; actual workbook rendered by LibreOffice",
        "excel_desktop_tested": False,
    }, ensure_ascii=False, indent=2) + "\n")
    print(result)
    print(DEST / "Pratinjau_Grafik_Wilayah.png")


if __name__ == "__main__":
    main()
