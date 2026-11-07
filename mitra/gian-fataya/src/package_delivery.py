"""Package only the tested deliverables; never include temporary test fixtures."""
from pathlib import Path
import hashlib
import json
import shutil
from zipfile import ZipFile,ZIP_DEFLATED
from workbook_links import validate_navigation

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'output'
validate_navigation(OUT/'POS_Gian_Fataya.xlsx')
validate_navigation(ROOT/'demo/SIMULASI_Grafik_Wilayah.xlsx')
report=json.loads((OUT/'hasil_validasi.json').read_text())
assert report['failed']==0 and report['passed']>=70
assert (OUT/'hasil_validasi.json').stat().st_mtime >= (OUT/'POS_Gian_Fataya.xlsx').stat().st_mtime, 'Re-run validation after rebuilding the workbook.'
# Excel's .bas import uses ANSI; the module deliberately contains only ASCII.
source=(ROOT/'vba/Modul_GianFataya.bas').read_text()
assert source.isascii()
(OUT/'Modul_GianFataya.bas').write_bytes(source.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copyfile(ROOT/'docs/PANDUAN_GIAN_FATAYA.md',OUT/'PANDUAN_GIAN_FATAYA.md')
shutil.copyfile(ROOT/'docs/RISET_UI_UX_GIAN_FATAYA.md',OUT/'RISET_UI_UX_GIAN_FATAYA.md')
(OUT/'BACA_DULU.txt').write_text('''GIAN FATAYA - KONSINYASI DAN INVOICE WILAYAH

1. Buka POS_Gian_Fataya.xlsx, mulai dari BERANDA. Kini hanya 6 sheet nyata.
2. Isi DATA USAHA: wilayah/PIC, outlet dan kelompok Amanda/Independen,
   harga jual, lot/expired, info pembayaran; catat stok AWAL aktual.
3. TRANSAKSI: input A:I, pemeriksaan J:L. Klik + untuk detail M:X.
   TAGIHAN memuat invoice dan cicilan; klik tautan pembayaran pada baris 5.
4. Tombol form/PDF opsional: simpan sebagai .xlsm, Alt+F11 > File > Import
   Modul_GianFataya.bas, Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol.
5. Simpan file dengan Ctrl+S dan cadangkan sebelum penagihan bulanan.
6. BERANDA: grafik batang A20 mengikuti periode B5/E5. Grafik area G20
   menampilkan 12 bulan sampai E5; pilih wilayah di I18. Rincian A36:H44.
   contoh/SIMULASI_Grafik_Wilayah.xlsx berisi DATA FIKTIF untuk mencoba grafik.
   contoh/Pratinjau_Grafik_Wilayah.png adalah tampilan contoh, bukan performa nyata.
7. PO minimal 10 pcs per nomor PO + outlet (beberapa lot boleh digabung).
   RETUR hanya untuk kerusakan fisik kemasan, dengan alasan.

Tidak ada transaksi atau harga jual fiktif di file utama.
Biaya bahan Rp9.410/pouch berasal dari foto mitra, bukan harga jual atau HPP lengkap.
VBA diberikan terpisah; tidak tertanam dalam .xlsx, dan belum dijalankan di Excel desktop.
Rumus dihitung ulang dan diuji dengan LibreOffice pada data simulasi terpisah.
Format invoice sementara karena contoh invoice asli belum tersedia.
Panduan lengkap: PANDUAN_GIAN_FATAYA.md.
Riset UI/UX dan pemetaan kolom Updated: RISET_UI_UX_GIAN_FATAYA.md.
''',encoding='utf-8')
names=['POS_Gian_Fataya.xlsx','Modul_GianFataya.bas','PANDUAN_GIAN_FATAYA.md','RISET_UI_UX_GIAN_FATAYA.md','BACA_DULU.txt','hasil_validasi.json','hasil_validasi_navigasi.json']
examples=['SIMULASI_Grafik_Wilayah.xlsx','Pratinjau_Grafik_Wilayah.png','hasil_validasi_grafik.json']
manifest={n:hashlib.sha256((OUT/n).read_bytes()).hexdigest() for n in names}
manifest.update({'contoh/'+n:hashlib.sha256((ROOT/'demo'/n).read_bytes()).hexdigest() for n in examples})
(OUT/'SHA256.json').write_text(json.dumps(manifest,indent=2))
dest=OUT/'Paket_POS_Gian_Fataya.zip'
with ZipFile(dest,'w',ZIP_DEFLATED) as z:
    for n in names+['SHA256.json']:z.write(OUT/n,n)
    for name in examples:
        z.write(ROOT/'demo'/name, 'contoh/'+name)
    for directory in ['data','src','tests']:
        for p in (ROOT/directory).rglob('*'):
            if p.is_file() and '__pycache__' not in p.parts:
                z.write(p,'pengembang/'+str(p.relative_to(ROOT)))
    z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
    z.write(ROOT/'vba/Modul_GianFataya.bas','pengembang/vba/Modul_GianFataya.bas')
    z.write(ROOT/'docs/PANDUAN_GIAN_FATAYA.md','pengembang/docs/PANDUAN_GIAN_FATAYA.md')
    z.write(ROOT/'docs/RISET_UI_UX_GIAN_FATAYA.md','pengembang/docs/RISET_UI_UX_GIAN_FATAYA.md')
with ZipFile(dest) as z:
    assert z.testzip() is None
print(dest)
print('Validation:',report['passed'],'passed; 0 failed')
