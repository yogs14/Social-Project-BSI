"""Package validated operational files and reproducible sources, never test records."""
import hashlib,json,shutil
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'output'
report=json.loads((OUT/'hasil_validasi.json').read_text())
assert report['failed']==0 and report['passed']>=70
assert (OUT/'hasil_validasi.json').stat().st_mtime >= (OUT/'POS_Murita_Crepes.xlsx').stat().st_mtime,'Validate the rebuilt workbook before packaging.'
source=(ROOT/'vba/Modul_MuritaCrepes.bas').read_text()
assert source.isascii()
(OUT/'Modul_MuritaCrepes.bas').write_bytes(source.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copyfile(ROOT/'docs/PANDUAN_MURITA_CREPES.md',OUT/'PANDUAN_MURITA_CREPES.md')
(OUT/'BACA_DULU.txt').write_text('''MURITA CREPES - PAK SYARIF

1. Buka POS_Murita_Crepes.xlsx, mulai dari MULAI.
2. Periksa 34 menu/harga sumber; isi harga SKM, stok awal, saldo kas aktual.
3. Produksi dua tahap: bahan mentah -> PREMIX -> WET -> kasir.
4. Kalibrasi sendok topping bila datanya sudah ada; sebelum itu takaran adalah acuan.
5. Loyalty awalnya TIDAK aktif. Tetapkan aturan terlebih dahulu; pelanggan opsional.
6. Untuk tombol: simpan salinan .xlsm, impor Modul_MuritaCrepes.bas via Alt+F11,
   Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol. Simpan dengan Ctrl+S.
7. Tanpa macro, jurnal NOTA, ITEM dan GERAK BAHAN diisi bersama sesuai panduan.

Tidak ada penjualan, saldo stok atau pelanggan fiktif dalam workbook utama.
Rumus sumber SKM yang salah tidak dipakai; HPP belum lengkap sampai harga diketahui.
Rumus diuji dengan LibreOffice pada simulasi terpisah. VBA belum dikompilasi/
dijalankan di Excel desktop dan belum tertanam di .xlsx; uji sebelum operasional.
Panduan lengkap: PANDUAN_MURITA_CREPES.md.
''',encoding='utf-8')
names=['POS_Murita_Crepes.xlsx','Modul_MuritaCrepes.bas','PANDUAN_MURITA_CREPES.md','BACA_DULU.txt','hasil_validasi.json']
manifest={name:hashlib.sha256((OUT/name).read_bytes()).hexdigest() for name in names}
(OUT/'SHA256.json').write_text(json.dumps(manifest,indent=2)+'\n')
dest=OUT/'Paket_POS_Murita_Crepes.zip'
with ZipFile(dest,'w',ZIP_DEFLATED) as z:
    for name in names+['SHA256.json']:z.write(OUT/name,name)
    for directory in ['data','src','tests','vba','docs']:
        for path in (ROOT/directory).rglob('*'):
            if path.is_file() and '__pycache__' not in path.parts:z.write(path,'pengembang/'+str(path.relative_to(ROOT)))
    z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
with ZipFile(dest) as z:assert z.testzip() is None
print(dest)
print('Validation:',report['passed'],'passed;',report['failed'],'failed')
