"""Package the validated blank operating workbook and reproducible sources."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json,shutil
ROOT=Path(__file__).resolve().parents[1];OUT=ROOT/'output'
report=json.loads((OUT/'hasil_validasi.json').read_text())
assert report['failed']==0 and report['passed']>=70
assert (OUT/'hasil_validasi.json').stat().st_mtime >= (OUT/'POS_Chicken_Al_Haqi.xlsx').stat().st_mtime,'Revalidate after rebuilding.'
source=(ROOT/'vba/Modul_ChickenAlHaqi.bas').read_text();assert source.isascii()
(OUT/'Modul_ChickenAlHaqi.bas').write_bytes(source.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copyfile(ROOT/'docs/PANDUAN_CHICKEN_AL_HAQI.md',OUT/'PANDUAN_CHICKEN_AL_HAQI.md')
(OUT/'BACA_DULU.txt').write_text('''CHICKEN AL-HAQI - MBA GITA

Buka POS_Chicken_Al_Haqi.xlsx. Mulai dari MULAI lalu KASIR/HARIAN.
Empat potongan ayam + pilihan rasa, nasi terpisah, tujuh tab utama.
Belanja: isi jumlah pcs AKTUAL. Kg dan perkiraan 9/10 pcs bukan stok tambahan.
Isi saldo awal/fisik nyata; stok yang belum diisi tetap BELUM AWAL, bukan nol.
HPP Rp7.000/pcs dari sumber adalah acuan belum dikonfirmasi, bukan laba pasti.
Pisahkan biaya ayam, pribadi dan perpindahan uang ke/dari usaha lain (obat).

Tombol: simpan salinan .xlsm, Alt+F11 > File > Import Modul_ChickenAlHaqi.bas,
Debug > Compile VBAProject, Alt+F8 > PasangTombol. Simpan Ctrl+S.
VBA belum tertanam dalam .xlsx dan belum diuji di Excel desktop.
Rumus diuji dengan LibreOffice pada simulasi terpisah, bukan transaksi mitra.
Panduan lengkap: PANDUAN_CHICKEN_AL_HAQI.md.
''',encoding='utf-8')
names=['POS_Chicken_Al_Haqi.xlsx','Modul_ChickenAlHaqi.bas','PANDUAN_CHICKEN_AL_HAQI.md','BACA_DULU.txt','hasil_validasi.json']
(OUT/'SHA256.json').write_text(json.dumps({n:hashlib.sha256((OUT/n).read_bytes()).hexdigest() for n in names},indent=2)+'\n')
dest=OUT/'Paket_POS_Chicken_Al_Haqi.zip'
with ZipFile(dest,'w',ZIP_DEFLATED) as z:
    for n in names+['SHA256.json']:z.write(OUT/n,n)
    for directory in ['src','data','tests','docs','vba']:
        for path in (ROOT/directory).rglob('*'):
            if path.is_file() and '__pycache__' not in path.parts:z.write(path,'pengembang/'+str(path.relative_to(ROOT)))
    z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
with ZipFile(dest) as z:assert z.testzip() is None
print(dest);print('Validation:',report['passed'],'passed;',report['failed'],'failed')
