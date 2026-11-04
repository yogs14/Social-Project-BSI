"""Create a clearly labelled PDF training copy; never change the operational file.

The invoice formulas are recalculated with LibreOffice. Excel VBA is not executed.
"""
import datetime as dt
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'output/POS_Gian_Fataya.xlsx'
DEST = ROOT / 'demo'
NAME = 'SIMULASI_Uji_PDF_Gian_Fataya.xlsx'


def put(sheet, row, data):
    for column, value in data.items():
        sheet.cell(row, column, value)


def main():
    fingerprint = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    w = load_workbook(SOURCE)
    for title, last in [('MUTASI', 2006), ('INVOICE', 406), ('PEMBAYARAN', 1006)]:
        assert not any(w[title].cell(r, 1).value for r in range(7, last + 1)), 'Use the blank repository template, not a live operating file.'
    assert sum(bool(w['OUTLET'].cell(r, 1).value) for r in range(7, 68)) == 1
    assert not any(w['LOT'].cell(r, 1).value for r in range(7, 87))
    today = dt.datetime(2026, 11, 4)
    invoice = 'GF-DEMO-001'
    w.properties.title = 'SIMULASI Gian Fataya - latihan ekspor PDF'
    w.properties.subject = 'Data latihan; bukan transaksi atau tagihan nyata'
    w.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
    w['WILAYAH']['C7'] = 'PIC SIMULASI - bukan pelanggan asli'
    w['WILAYAH']['D7'] = 'Alamat latihan, bukan alamat penagihan nyata'
    w['OUTLET']['B7'] = 'Gudang SIMULASI'
    put(w['OUTLET'], 8, {1: 'DEMO-JT', 2: 'Outlet SIMULASI', 3: 'W01', 4: 'PIC LATIHAN', 5: 'Alamat latihan', 7: 'AKTIF'})
    w['PRODUK']['B7'] = 'Pastel - DATA SIMULASI'
    w['PRODUK']['D7'] = 10000
    w['PRODUK']['G7'] = 'Harga fiktif untuk latihan PDF, bukan harga mitra.'
    put(w['LOT'], 7, {1: 'LOT-DEMO', 2: 'GF-PST-001', 4: dt.datetime(2027, 1, 31), 5: 'Batch fiktif untuk latihan'})
    for r, key, kind, location, qty in [(7, 'DEMO-M1', 'AWAL', 'GDG', 40), (8, 'DEMO-M2', 'KIRIM', 'DEMO-JT', 40), (9, 'DEMO-M3', 'JUAL', 'DEMO-JT', 30)]:
        put(w['MUTASI'], r, {1: key, 2: today, 3: kind, 4: location, 5: 'LOT-DEMO', 6: qty, 9: 'AKTIF', 10: 'DATA SIMULASI - bukan transaksi nyata'})
    w['MUTASI']['U8'] = 'PO-DEMO-001'
    w['MUTASI']['G9'] = 10000
    w['MUTASI']['H9'] = invoice
    put(w['INVOICE'], 7, {1: invoice, 2: today, 3: 'W01', 4: today, 5: today, 6: dt.datetime(2026, 11, 11), 7: .1, 8: 20000, 9: 'TERBIT', 10: 300000, 18: 'SIMULASI PDF - BUKAN TAGIHAN NYATA'})
    w['DASHBOARD']['B5'] = today
    w['DASHBOARD']['E5'] = today
    w['FORM']['B6'] = today
    w['FORM']['B20'] = today
    w['FORM']['B37'] = today
    w['REFERENSI']['B8'] = 'SIMULASI - tidak ada rekening atau permintaan pembayaran nyata'
    print_sheet = w['CETAK INVOICE']
    print_sheet['B5'] = invoice
    print_sheet['A7'] = 'GIAN FATAYA - SIMULASI'
    print_sheet['A8'] = 'BUKAN TAGIHAN NYATA | LATIHAN EKSPOR PDF'
    for r in range(18, 67):
        print_sheet.row_dimensions[r].hidden = True
    start = w['MULAI']
    start['A1'] = 'GIAN FATAYA / SIMULASI PDF'
    start['A2'] = 'File latihan terpisah; satu invoice GF-DEMO-001. Total yang diharapkan Rp290.000. Seluruh transaksi/harga/PIC di sini adalah simulasi.'
    start['A6'] = 'UJI TOMBOL EKSPOR PDF DI EXCEL DESKTOP'
    steps = [
        ('1. Simpan salinan .xlsm', 'File > Save As > Excel Macro-Enabled Workbook (*.xlsm). Gunakan nama SIMULASI_Gian_Fataya.xlsm.'),
        ('2. Impor VBA', 'Alt+F11 > pilih project file latihan > File > Import File > Modul_GianFataya.bas. Modul disertakan pada paket latihan.'),
        ('3. Pasang tombol', 'Debug > Compile VBAProject. Kembali ke Excel: Alt+F8 > PasangTombol > Run. Simpan Ctrl+S.'),
        ('4. Buka invoice latihan', 'Sheet CETAK INVOICE: B5 sudah GF-DEMO-001 dan indikator harus SIAP CETAK. Tidak perlu mengisi data transaksi dari awal.'),
        ('5. Ekspor', 'Klik EKSPOR INVOICE PDF di dekat D5, pilih lokasi penyimpanan dan tekan Save. Nama awal GF-DEMO-001.pdf.'),
        ('6. Periksa hasil', 'PDF harus bertanda SIMULASI, qty 30, bruto Rp300.000, diskon Rp30.000, ongkir Rp20.000, total Rp290.000.'),
    ]
    for r, (label, body) in zip([8, 11, 14, 17, 20, 23], steps):
        start[f'A{r}'] = label
        start[f'D{r}'] = body
    start['A37'] = 'SELURUH DATA DI FILE INI UNTUK LATIHAN. Bukan harga, stok, pelanggan atau tagihan nyata. File operasional POS_Gian_Fataya.xlsx tetap kosong dan tidak diubah. VBA belum diuji di Excel desktop.'
    w.active = w.sheetnames.index('CETAK INVOICE')
    for sheet in w:
        sheet.sheet_view.tabSelected = sheet.title == 'CETAK INVOICE'
    DEST.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='gian-invoice-demo-') as td:
        raw = Path(td) / NAME
        w.save(raw)
        result = subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', str(DEST), str(raw)], capture_output=True, text=True, timeout=120)
        assert result.returncode == 0 and (DEST / NAME).exists(), result.stdout + result.stderr
    result = load_workbook(DEST / NAME, data_only=True)
    assert result['CETAK INVOICE']['A14'].value == 'SIAP CETAK'
    assert result['INVOICE']['P7'].value == 'OK'
    assert result['CETAK INVOICE']['F68'].value == 300000
    assert result['CETAK INVOICE']['F69'].value == 30000
    assert result['CETAK INVOICE']['F70'].value == 20000
    assert result['CETAK INVOICE']['F71'].value == 290000
    assert result['CETAK INVOICE']['F73'].value == 290000
    assert all(result['MUTASI'][f'R{r}'].value == 'OK' for r in range(7, 10))
    assert all(result['MUTASI'][f'T{r}'].value >= 0 for r in [8, 9])
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest() == fingerprint
    source_module = ROOT / 'output/Modul_GianFataya.bas'
    shutil.copyfile(source_module, DEST / source_module.name)
    report = {
        'invoice': invoice, 'label': 'SIMULASI - BUKAN TAGIHAN NYATA',
        'print_status': 'SIAP CETAK', 'gross': 300000, 'discount': 30000,
        'freight': 20000, 'total': 290000, 'source_operational_file_unchanged': True,
        'validation': 'Actual demo formulas recalculated with LibreOffice',
        'vba_executed': False, 'excel_desktop_tested': False,
    }
    (DEST / 'hasil_validasi_demo.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    files = [NAME, source_module.name, 'PANDUAN_UJI_PDF.md', 'hasil_validasi_demo.json']
    (DEST / 'SHA256.json').write_text(json.dumps({name: hashlib.sha256((DEST / name).read_bytes()).hexdigest() for name in files}, indent=2) + '\n')
    with ZipFile(DEST / 'Paket_SIMULASI_Uji_PDF.zip', 'w', ZIP_DEFLATED) as z:
        for name in files + ['SHA256.json']:
            z.write(DEST / name, name)
    print(DEST / NAME)
    print('Verified: SIAP CETAK; total Rp290.000; operational source unchanged. Excel VBA not executed.')


if __name__ == '__main__':
    main()
