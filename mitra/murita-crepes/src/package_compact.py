"""Deliver the compact workbook plus its matching VBA, guide and evidence."""
import hashlib, json, shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from build_compact import ROOT, OUT, DEST

def main():
    report=json.loads((OUT/'hasil_validasi.json').read_text())
    assert report['failed']==0 and report['passed']>=87
    assert (OUT/'hasil_validasi.json').stat().st_mtime >= DEST.stat().st_mtime
    shutil.copyfile(ROOT/'docs/PANDUAN_MUFRITA_RINGKAS.md',OUT/'PANDUAN_MUFRITA_RINGKAS.md')
    (OUT/'BACA_DULU.txt').write_text('''MUFRITA CREPES - VERSI RINGKAS

1. Buka POS_Mufrita_Crepes_Ringkas.xlsx. Terdapat 8 tab kerja + 1 rumus tersembunyi.
2. Periksa menu/harga, isi harga SKM pada KATALOG!E240 dan stok fisik awal.
3. Produksi: 7 bahan -> PREMIX -> WET; penjualan hanya memakai WET + pelengkap/topping.
4. Untuk menyimpan melalui tombol, simpan salinan .xlsm lalu impor
   Modul_Mufrita_Ringkas.bas. Debug > Compile VBAProject, jalankan PasangTombol.
5. Ganti modul VBA versi sebelumnya; jangan impor dua modul dengan prosedur ganda.
6. Tanpa VBA: isi jurnal nota, item dan pemakaian bahan bersama sesuai panduan.
7. Tab STOK memakai tampilan Stok Realtime POS Pak Seger; master di KATALOG A225.
8. TUTUP KASIR: tanggal C5, modal I8, fisik I11. Simpan melalui tombol untuk arsip revisi.
9. Poin awalnya nonaktif; jangan mengarang harga, stok, pelanggan atau kebijakan reward.

Rumus telah dihitung ulang dan diuji di LibreOffice. VBA tidak dijalankan atau
dikompilasi di Microsoft Excel; uji tombol dan ekspor PDF sebelum operasional.
Panduan: PANDUAN_MUFRITA_RINGKAS.md. Hasil uji: hasil_validasi.json.
Nama sumber: Murita Crepes, Pak Syarif; nama tampilan mengikuti permintaan Mufrita.
''',encoding='utf-8')
    provenance={
        'pos_reference':json.loads((ROOT/'data/pos_reference.json').read_text()),
        'mom':'https://docs.google.com/spreadsheets/d/1MCGsQK3ILZBpxMRH2-fWdnkxv1I8Cfq83BnCVYbJqmQ/edit?gid=52890219',
        'mom_date':'2026-10-24',
        'source_data':json.loads((ROOT/'data/sumber.json').read_text()),
        'original_operational_workbook_preserved':True,
    }
    (OUT/'sumber_ringkas.json').write_text(json.dumps(provenance,indent=2))
    names=[DEST.name,'Modul_Mufrita_Ringkas.bas','PANDUAN_MUFRITA_RINGKAS.md','BACA_DULU.txt','hasil_validasi.json','peta_sheet.json','sumber_ringkas.json']
    hashes={name:hashlib.sha256((OUT/name).read_bytes()).hexdigest() for name in names}
    (OUT/'SHA256.json').write_text(json.dumps(hashes,indent=2))
    target=OUT/'Paket_Mufrita_Crepes_Ringkas.zip'
    with ZipFile(target,'w',ZIP_DEFLATED) as z:
        for name in names+['SHA256.json']:z.write(OUT/name,name)
        for folder in ['src','tests','vba','data']:
            for file in (ROOT/folder).rglob('*'):
                if file.is_file() and '__pycache__' not in file.parts:z.write(file,'pengembang/'+str(file.relative_to(ROOT)))
        z.write(ROOT/'requirements.txt','pengembang/requirements.txt')
    with ZipFile(target) as z:assert z.testzip() is None
    print(target)

if __name__=='__main__':main()
