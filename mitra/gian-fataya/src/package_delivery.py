"""Package only the tested deliverables; never include temporary test fixtures."""
from pathlib import Path
import hashlib
import json
import shutil
from zipfile import ZipFile,ZIP_DEFLATED

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'output'
report=json.loads((OUT/'hasil_validasi.json').read_text())
assert report['failed']==0 and report['passed']>=70
assert (OUT/'hasil_validasi.json').stat().st_mtime >= (OUT/'POS_Gian_Fataya.xlsx').stat().st_mtime, 'Re-run validation after rebuilding the workbook.'
# Excel's .bas import uses ANSI; the module deliberately contains only ASCII.
source=(ROOT/'vba/Modul_GianFataya.bas').read_text()
assert source.isascii()
(OUT/'Modul_GianFataya.bas').write_bytes(source.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copyfile(ROOT/'docs/PANDUAN_GIAN_FATAYA.md',OUT/'PANDUAN_GIAN_FATAYA.md')
(OUT/'BACA_DULU.txt').write_text('''GIAN FATAYA - KONSINYASI DAN INVOICE WILAYAH

1. Buka POS_Gian_Fataya.xlsx, mulai dari sheet MULAI.
2. Isi harga jual, outlet, PIC/alamat tagihan, lot/expired, dan stok awal aktual.
3. Jurnal manual langsung bekerja tanpa macro. Sel biru untuk input;
   sel abu-abu berisi rumus. Jangan mengubah kode master yang sudah dipakai.
4. Tombol FORM/PDF opsional: simpan sebagai .xlsm, Alt+F11 > File > Import
   Modul_GianFataya.bas, Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol.
5. Simpan file dengan Ctrl+S dan cadangkan sebelum penagihan bulanan.

Tidak ada transaksi atau harga jual fiktif di file utama.
Biaya bahan Rp9.410/pouch berasal dari foto mitra, bukan harga jual atau HPP lengkap.
VBA diberikan terpisah; tidak tertanam dalam .xlsx, dan belum dijalankan di Excel desktop.
Rumus dihitung ulang dan diuji dengan LibreOffice pada data simulasi terpisah.
Format invoice sementara karena contoh invoice asli belum tersedia.
Panduan lengkap: PANDUAN_GIAN_FATAYA.md.
''',encoding='utf-8')
names=['POS_Gian_Fataya.xlsx','Modul_GianFataya.bas','PANDUAN_GIAN_FATAYA.md','BACA_DULU.txt','hasil_validasi.json']
manifest={n:hashlib.sha256((OUT/n).read_bytes()).hexdigest() for n in names}
(OUT/'SHA256.json').write_text(json.dumps(manifest,indent=2))
dest=OUT/'Paket_POS_Gian_Fataya.zip'
with ZipFile(dest,'w',ZIP_DEFLATED) as z:
    for n in names+['SHA256.json']:z.write(OUT/n,n)
    for directory in ['data','src','tests']:
        for p in (ROOT/directory).rglob('*'):
            if p.is_file() and '__pycache__' not in p.parts:
                z.write(p,'pengembang/'+str(p.relative_to(ROOT)))
    z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
    z.write(ROOT/'vba/Modul_GianFataya.bas','pengembang/vba/Modul_GianFataya.bas')
    z.write(ROOT/'docs/PANDUAN_GIAN_FATAYA.md','pengembang/docs/PANDUAN_GIAN_FATAYA.md')
with ZipFile(dest) as z:
    assert z.testzip() is None
print(dest)
print('Validation:',report['passed'],'passed; 0 failed')
