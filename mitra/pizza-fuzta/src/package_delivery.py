"""Package all three verified workbooks, both VBA modules and source provenance."""
import hashlib,json,shutil
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'output'
FILES=['POS_Pizza_Fuzta_Medokan.xlsx','Rekap_Pizza_Fuzta_Karah.xlsx','Gudang_Pizza_Fuzta.xlsx']
report=json.loads((OUT/'hasil_validasi.json').read_text())
assert report['failed']==0 and report['passed']>=70
for file in FILES:assert (OUT/'hasil_validasi.json').stat().st_mtime>=(OUT/file).stat().st_mtime,'Validate rebuilt workbooks first.'
for file in ['Modul_PizzaFuzta.bas','Modul_TransferPizza.bas']:
    s=(ROOT/'vba'/file).read_text();assert s.isascii()
    (OUT/file).write_bytes(s.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copyfile(ROOT/'docs/PANDUAN_PIZZA_FUZTA.md',OUT/'PANDUAN_PIZZA_FUZTA.md')
(OUT/'BACA_DULU.txt').write_text('''PIZZA FUZTA - BU ELOK

Tiga workbook mandiri:
- POS_Pizza_Fuzta_Medokan.xlsx: kasir 49 menu, stok dan kas Medokan.
- Rekap_Pizza_Fuzta_Karah.xlsx: rekap Majoo dan stok/kas pendamping Karah.
- Gudang_Pizza_Fuzta.xlsx: stok rumah, permintaan, distribusi dan production house.

Tidak ada sinkronisasi otomatis/API Majoo. Pertukaran stok via CSV PF1 offline.
Semua file mulai tanpa transaksi/saldo fiktif. Isi harga bahan/stok aktual dan
setujui resep/satuan sebelum mengandalkan pemakaian otomatis. Nota dengan resep
belum disetujui tercatat sebagai uang, tetapi status stok BELUM LENGKAP.

Tombol: simpan salinan .xlsm, impor KEDUA modul .bas, Debug > Compile VBAProject,
lalu Alt+F8 > PasangTombol. VBA belum diuji di Excel desktop, belum tertanam di
.xlsx. Rumus diuji dengan LibreOffice dan pertukaran CSV diuji dengan Python.

Gunakan PANDUAN_PIZZA_FUZTA.md untuk langkah kasir, rekap, transfer dan koreksi.
Jangan memakai macro lama dalam file sumber sebagai modul baru ini.
''',encoding='utf-8')
names=FILES+['Modul_PizzaFuzta.bas','Modul_TransferPizza.bas','PANDUAN_PIZZA_FUZTA.md','BACA_DULU.txt','hasil_validasi.json']
(OUT/'SHA256.json').write_text(json.dumps({n:hashlib.sha256((OUT/n).read_bytes()).hexdigest() for n in names},indent=2)+'\n')
dest=OUT/'Paket_POS_Pizza_Fuzta.zip'
with ZipFile(dest,'w',ZIP_DEFLATED) as z:
    for n in names+['SHA256.json']:z.write(OUT/n,n)
    for directory in ['data','src','tests','vba','docs']:
        for p in (ROOT/directory).rglob('*'):
            if p.is_file() and '__pycache__' not in p.parts:z.write(p,'pengembang/'+str(p.relative_to(ROOT)))
    z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
with ZipFile(dest) as z:assert z.testzip() is None
print(dest);print('Validation:',report['passed'],'passed;',report['failed'],'failed')
