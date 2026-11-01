"""Three standalone workbooks; explicit offline transfers, no cross-file formulas."""
from pathlib import Path
import argparse,datetime as dt,json
from layout import Layout,CAP,rng,look,checks,col
ROOT=Path(__file__).resolve().parents[1]
FILES={'MEDOKAN':'POS_Pizza_Fuzta_Medokan.xlsx','KARAH':'Rekap_Pizza_Fuzta_Karah.xlsx','GUDANG':'Gudang_Pizza_Fuzta.xlsx'}
KINDS=['AWAL','BELI','KIRIM','TERIMA','PAKAI','RUSAK','KOREKSI MASUK','KOREKSI KELUAR','PRODUKSI MASUK','PRODUKSI PAKAI']
METHODS=['Tunai','QRIS','Transfer','GoPay','OVO','DANA','Lainnya']

def build(role,path):
    p=json.loads((ROOT/'data/profile.json').read_text());path.parent.mkdir(parents=True,exist_ok=True)
    b=Layout(path);w=b.w;f=b.f;today=dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    for name,sheet,c in [('KodeBahan','BAHAN','A'),('KodeNota','NOTA','A'),('KodeMenu','MENU','A')]:
        if role=='MEDOKAN' or sheet=='BAHAN':w.define_name(name,'='+rng(sheet,c))
    s=b.base('MULAI',f'Bu Elok • {p["locations"][role]} • MoM 27 Oktober 2026 • File mandiri, tanpa sinkronisasi otomatis antar laptop.')
    s.merge_range('A6:L6',role+'  /  ALUR OPERASIONAL',f['head'])
    descriptions={
      'MEDOKAN':['Periksa 49 menu/varian L-M-S, burger dan minuman. Harga sumber dipertahankan; tandai yang belum dikonfirmasi. Tidak lagi terbatas pada 20 menu.','BAHAN dan BOM adalah acuan/draf. Isi harga dan stok awal; verifikasi nama, satuan dan takaran. Setujui BOM dan MENU hanya setelah benar. Harga bahan kosong berarti HPP belum lengkap.','KASIR menyimpan nota, item, harga/HPP snapshot dan pemakaian bahan. Jika ada resep yang belum disetujui, seluruh nota ditandai stok BELUM LENGKAP dan belum dipotong otomatis. Lengkapi pemakaian nota itu sekali saja sebelum mengandalkan saldo.','Pembelian Santri dan bahan langsung dicatat BELI, bukan biaya lain kedua. Tunai, QRIS, Transfer, GoPay, OVO dan DANA dipisahkan di tutup kasir.'],
      'KARAH':['Majoo tetap menjadi kasir utama. File ini mendampingi rekap harian, pembelian, kas, stok dan transfer; tidak meniru API atau format ekspor Majoo yang belum tersedia.','Isi REKAP MAJOO dari laporan harian final: bruto, diskon, serta penerimaan netto per metode. Selisih antara netto dan total metode harus nol. Jangan menghitung ulang discount/fee yang sudah termasuk laporan.','Resep, harga dan pembelian Karah belum diberikan. ACUAN MEDOKAN hanya pembanding, bukan katalog/recipe aktif Karah. Pemakaian stok Karah dicatat dari data nyata, bukan otomatis dari menu Medokan.','Menerima barang dari gudang melalui CSV; catat pembelian sendiri sesuai supplier Karah. Tidak ada data penjualan contoh atau saldo awal fiktif.'],
      'GUDANG':['Isi daftar bahan yang benar-benar disimpan di rumah serta stok fisik, harga dan satuan. Daftar awal hanya kandidat dari BOM Medokan; belum berarti semua bahan tersedia di gudang.','Catat PERMINTAAN dari outlet, BELI untuk barang masuk dan KIRIM untuk distribusi. Transfer bukan penjualan dan tidak menambah pendapatan.','Ekspor pengiriman CSV ke file outlet; setelah diterima, impor konfirmasi penerimaan. Konfirmasi hanya rekonsiliasi dan tidak menambah stok gudang kembali.','PRODUKSI menyediakan input/output batch aktual untuk rencana production house. Takaran dan hasil produksi tidak dibuat-buat. Input dan output disimpan dengan satu ID dokumen; tidak boleh dibatalkan satu sisi saja.']}
    for i,text in enumerate(descriptions[role]):s.merge_range(7+i*4,0,9+i*4,11,f'{i+1}. {text}',f['body'])
    s.merge_range('A26:L29','TRANSFER OFFLINE: KIRIM mengurangi stok pengirim. TERIMA menambah stok penerima pada tanggal diterima. ID transfer tidak boleh dipakai dua kali. Konfirmasi penerimaan disimpan terpisah; selisih kuantitas tetap terlihat. Tiga file ini tidak terhubung langsung dan tidak menyinkronkan kas/penjualan.',f['warn'])
    s.merge_range('A31:L34','TOMBOL: simpan salinan .xlsm, impor Modul_PizzaFuzta.bas melalui Alt+F11, Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol. Jurnal dapat diisi manual tanpa macro. CLI pengembang untuk CSV tersedia pada src/transfer_cli.py dan menerima file .xlsx; jangan memakai CLI pada .xlsm bertombol.',f['note'])
    s.merge_range('A36:L39','Data yang masih diperlukan: harga bahan/stok awal, persetujuan resep dan satuan, harga menu anomali, data khusus Karah serta bahan/hasil production house. File utama tidak mengimpor tiga transaksi dari template sumber. Catatan dan nomor kontak sisa template sate tidak digunakan di tampilan operasional.',f['warn'])
    s.merge_range('A41:L43','Kapasitas: 80 bahan, 4.000 gerak stok, 1.000 keuangan/permintaan, 2.000 konfirmasi. Medokan: 80 menu, 600 BOM, 1.000 nota, 3.000 item, 12 baris per nota. Karah: 1.000 rekap. Jangan menambah data di luar tabel; cadangkan dan arsipkan sebelum penuh.',f['note']);s.print_area('A1:L43')

    s=b.base('BAHAN','Kandidat bahan dari sumber Medokan; verifikasi sebelum dipakai. Harga/stok tidak diwarisi dari asumsi template lama. Satuan gram, pcs, botol/porsi tidak dicampur.',14)
    bf={'G':lambda r:f'IF(A{r}="","",IFERROR(N(E{r})/N(F{r}),0))',
      'K':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeBahan,A{r})<>1','KODE GANDA'),(f'OR(B{r}="",D{r}="")','ISI NAMA/SATUAN'),(f'AND(H{r}<>"YA",H{r}<>"TIDAK")','CEK LACAK'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<0,NOT(ISNUMBER(F{r})),F{r}<=0)','ISI HARGA/ISI BERSIH')]),
      'L':lambda r:f'IF(A{r}="","",IF(H{r}<>"YA","",ROUND(SUMIFS({rng("MUTASI","P")},{rng("MUTASI","D")},A{r},{rng("MUTASI","B")},"<="&TglAkhir),6)))',
      'M':lambda r:f'IF(A{r}="","",IF(H{r}<>"YA","TIDAK DILACAK",IF(L{r}<0,"STOK NEGATIF",IF(L{r}=0,"HABIS / BELUM AWAL",IF(AND(ISNUMBER(I{r}),L{r}<=I{r}),"BELANJA","TERSEDIA")))))'}
    b.table(s,'tblBahan',['Kode','Bahan','Kelompok','Satuan Dasar','Harga Kemasan','Isi Kemasan','Biaya / Satuan','Lacak Stok','Batas Belanja','Catatan','Cek Harga','Saldo','Status'],80,bf,'ABCDEFHIJ',{'E':'in_money','F':'in_qty','G':'qty','I':'in_qty','L':'qty'})
    for r,m in enumerate(p['materials'],7):
        for c,v in [(0,m['code']),(1,m['name']),(2,'ACUAN MEDOKAN'),(3,m['unit']),(7,'YA'),(9,m['source'])]:s.write(r-1,c,v,f['input'])
    s.set_column('B:B',33);s.set_column('J:J',50);b.audit(s,'K',86);b.validate(s,'H7:H86',['YA','TIDAK'])

    s=b.base('MUTASI','Satu baris per bahan/gerakan. ID Transfer sama di pengirim/penerima. Tanggal TERIMA adalah tanggal aktual diterima; Tgl Kirim dan Qty Kirim menyimpan sumber.',22)
    def mc(r):
        cond=[(f'COUNTIF({rng("MUTASI","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(JenisGerak,C{r})<>1','CEK JENIS'),(f'COUNTIF(KodeBahan,D{r})<>1','CEK BAHAN'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0)','CEK QTY'),(f'AND(M{r}<>"AKTIF",M{r}<>"VOID")','CEK STATUS'),(f'AND(C{r}="BELI",OR(NOT(ISNUMBER(F{r})),F{r}<0,COUNTIF(MetodeBayar,G{r})<>1))','CEK PEMBELIAN'),(f'AND(OR(C{r}="KIRIM",C{r}="TERIMA"),OR(H{r}="",COUNTIF({rng("MUTASI","H")},H{r})<>1,COUNTIF(KodeLokasi,I{r})<>1,COUNTIF(KodeLokasi,J{r})<>1,I{r}=J{r}))','CEK TRANSFER'),(f'AND(C{r}="KIRIM",I{r}<>CONFIG!$B$6)','ASAL BUKAN LOKASI INI'),(f'AND(C{r}="TERIMA",J{r}<>CONFIG!$B$6)','TUJUAN BUKAN LOKASI INI'),(f'AND(C{r}="TERIMA",OR(NOT(ISNUMBER(T{r})),T{r}<=0,T{r}>B{r},NOT(ISNUMBER(U{r})),U{r}<=0))','CEK DATA KIRIM'),(f'AND(OR(C{r}="PRODUKSI MASUK",C{r}="PRODUKSI PAKAI"),V{r}="")','ISI DOKUMEN PRODUKSI'),(f'AND(OR(C{r}="RUSAK",C{r}="KOREKSI MASUK",C{r}="KOREKSI KELUAR"),N{r}="")','ISI ALASAN')]
        cond += [(f'AND(OR(C{r}="PRODUKSI MASUK",C{r}="PRODUKSI PAKAI"),OR(COUNTIFS({rng("MUTASI","V")},V{r},{rng("MUTASI","C")},"PRODUKSI MASUK",{rng("MUTASI","M")},M{r})=0,COUNTIFS({rng("MUTASI","V")},V{r},{rng("MUTASI","C")},"PRODUKSI PAKAI",{rng("MUTASI","M")},M{r})=0))','LENGKAPI DUA SISI PRODUKSI')]
        if role=='MEDOKAN':cond += [(f'AND(L{r}<>"",OR(C{r}<>"PAKAI",COUNTIF(KodeNota,L{r})<>1))','CEK NO NOTA')]
        else:cond += [(f'L{r}<>""','NOTA TIDAK DIPAKAI DI SINI')]
        return checks(f'A{r}',cond)
    notevalid=f'OR(L{{r}}="",AND({look("L{r}","NOTA","H")}="AKTIF",{look("L{r}","NOTA","O")}="OK"))' if role=='MEDOKAN' else 'TRUE'
    # Format literal row placeholders only after the complete expression is built.
    mf={'O':lambda r:f'IF(A{r}="","",{look(f"D{r}","BAHAN","D")})',
      'P':lambda r:f'IF(AND(A{r}<>"",Q{r}="OK",M{r}="AKTIF",{notevalid.replace("{r}",str(r))}),N(E{r})*IF(OR(C{r}="AWAL",C{r}="BELI",C{r}="TERIMA",C{r}="KOREKSI MASUK",C{r}="PRODUKSI MASUK"),1,-1),0)',
      'Q':mc,
      'R':lambda r:f'IF(OR(A{r}="",{look(f"D{r}","BAHAN","H")}<>"YA",P{r}=0),"",ROUND(SUMIFS({rng("MUTASI","P")},{rng("MUTASI","D")},D{r},{rng("MUTASI","B")},"<"&B{r})+SUMIFS($P$7:P{r},$D$7:D{r},D{r},$B$7:B{r},B{r}),6))',
      'S':lambda r:f'IF(AND(A{r}<>"",C{r}="BELI",M{r}="AKTIF",Q{r}="OK"),N(F{r}),0)'}
    b.table(s,'tblMutasi',['ID Gerak','Tanggal','Jenis','Kode Bahan','Qty','Nominal Beli','Metode','ID Transfer','Asal','Tujuan','No Permintaan','No Nota','Status','Catatan','Satuan','Qty Bertanda','Cek','Saldo Historis','Beli Efektif','Tgl Kirim','Qty Kirim','Dokumen Produksi'],4000,mf,'ABCDEFGHIJKLMNTUV',{'B':'in_date','E':'in_qty','F':'in_money','P':'qty','R':'qty','S':'money','T':'in_date','U':'in_qty'})
    for c,src in [('C','=JenisGerak'),('D','=KodeBahan'),('G','=MetodeBayar'),('I','=KodeLokasi'),('J','=KodeLokasi'),('M',['AKTIF','VOID'])]:b.validate(s,c+'7:'+c+'4006',src)
    b.audit(s,'Q',4006);s.conditional_format('R7:R4006',{'type':'cell','criteria':'<','value':-.000001,'format':f['bad']});s.set_column('N:N',40)

    s=b.base('KEUANGAN','BIAYA USAHA, PRIVE, MODAL MASUK dan PENDAPATAN LAIN dipisahkan. BELI bahan hanya dicatat pada MUTASI; jangan diduplikasi di sini.')
    kf={'H':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("KEUANGAN","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(JenisKas,C{r})<>1','CEK JENIS'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0)','CEK NOMINAL'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(G{r}<>"AKTIF",G{r}<>"VOID")','CEK STATUS')]),'I':lambda r:f'IF(AND(A{r}<>"",H{r}="OK",G{r}="AKTIF"),N(D{r})*IF(OR(C{r}="MODAL MASUK",C{r}="PENDAPATAN LAIN"),1,-1),0)'}
    b.table(s,'tblKeuangan',['ID Kas','Tanggal','Jenis','Nominal','Metode','Catatan','Status','Cek','Kas Bertanda'],1000,kf,'ABCDEFG',{'B':'in_date','D':'in_money','I':'money'})
    for c,src in [('C','=JenisKas'),('E','=MetodeBayar'),('G',['AKTIF','VOID'])]:b.validate(s,c+'7:'+c+'1006',src)
    b.audit(s,'H',1006)

    s=b.base('PERMINTAAN','Satu ID permintaan per bahan/outlet. Permintaan adalah rencana, tidak menambah/mengurangi stok. Di file outlet, realisasi tampak setelah transfer penerimaan dicatat.')
    movement='KIRIM' if role=='GUDANG' else 'TERIMA'
    pf={'G':lambda r:f'IF(A{r}="","",SUMIFS({rng("MUTASI","E")},{rng("MUTASI","K")},A{r},{rng("MUTASI","D")},D{r},{rng("MUTASI","J")},C{r},{rng("MUTASI","C")},"{movement}",{rng("MUTASI","M")},"AKTIF",{rng("MUTASI","Q")},"OK"))','H':lambda r:f'IF(A{r}="","",MAX(0,N(E{r})-G{r}))','I':lambda r:f'IF(A{r}="","",IF(G{r}>E{r},"LEBIH",IF(G{r}=E{r},"TERPENUHI",IF(G{r}=0,"BELUM","SEBAGIAN"))))','J':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("PERMINTAAN","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0)','CEK TANGGAL'),(f'OR(C{r}="GUDANG",COUNTIF(KodeLokasi,C{r})<>1)','PILIH OUTLET'),(f'COUNTIF(KodeBahan,D{r})<>1','CEK BAHAN'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0)','CEK QTY')])}
    b.table(s,'tblPermintaan',['ID Permintaan','Tanggal','Outlet','Kode Bahan','Qty Diminta','Catatan','Realisasi','Sisa','Status','Cek'],1000,pf,'ABCDEF',{'B':'in_date','E':'in_qty','G':'qty','H':'qty'});b.validate(s,'C7:C1006',['MEDOKAN','KARAH']);b.validate(s,'D7:D1006','=KodeBahan');b.audit(s,'J',1006)

    s=b.base('KONFIRMASI','Bukti penerimaan atas KIRIM dari lokasi ini. Tidak mengubah stok lagi. Kuantitas berbeda ditandai SELISIH; data pengiriman tidak ditimpa.',12)
    def ack_check(r):
        key=f'A{r}';source=lambda c:look(key,'MUTASI',c,'""','H')
        return checks(key,[(f'COUNTIF({rng("KONFIRMASI","A")},A{r})<>1','ID GANDA'),(f'COUNTIF({rng("MUTASI","H")},A{r})<>1','KIRIM TIDAK ADA'),(f'OR({source("C")}<>"KIRIM",{source("M")}<>"AKTIF")','KIRIM TIDAK AKTIF'),(f'OR(C{r}<>CONFIG!$B$6,C{r}<>{source("I")},D{r}<>{source("J")},E{r}<>{source("D")},H{r}<>{source("O")},ABS(N(F{r})-N({source("E")}))>0.000001,I{r}<>{source("B")})','DATA KIRIM BERBEDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<I{r},B{r}>TODAY(),NOT(ISNUMBER(G{r})),G{r}<=0)','CEK PENERIMAAN'),(f'ABS(N(G{r})-N(F{r}))>0.000001','SELISIH')])
    b.table(s,'tblKonfirmasi',['ID Transfer','Tgl Terima','Asal','Tujuan','Kode Bahan','Qty Kirim','Qty Terima','Satuan','Tgl Kirim','Catatan','Cek','Selisih'],2000,{'K':ack_check,'L':lambda r:f'IF(A{r}="","",N(G{r})-N(F{r}))'},'ABCDEFGHIJ',{'B':'in_date','F':'in_qty','G':'in_qty','I':'in_date','L':'qty'});b.audit(s,'K',2006)

    s=b.base('FORM','Input barang, pengiriman dan keuangan. Tombol memerlukan VBA; kolom jurnal tetap dapat diisi manual. Penerimaan CSV memakai validasi tujuan/unit/ID ganda.')
    fields=[('A6','Tanggal','B6','date'),('D6','Jenis','E6','text'),('A8','Kode Bahan','B8','text'),('D8','Qty','E8','qty'),('A10','Nominal Beli','B10','money'),('D10','Metode Beli','E10','text'),('A12','Tujuan KIRIM','B12','text'),('D12','ID Transfer','E12','text'),('A14','No Permintaan','B14','text'),('D14','Catatan/Alasan','E14','text'),('A21','Tanggal Kas','B21','date'),('D21','Jenis Kas','E21','text'),('A23','Nominal','B23','money'),('D23','Metode','E23','text'),('A25','Catatan','B25','text')]
    for label,text,cell,kind in fields:s.write(label,text,f['label']);s.write_blank(cell,None,f['input' if kind=='text' else 'in_'+kind])
    for cell in ['B6','B21']:s.write_datetime(cell,today,f['in_date'])
    s.write('E6','BELI',f['input']);s.write('E10','Tunai',f['input']);s.write('E21','BIAYA USAHA',f['input']);s.write('E23','Tunai',f['input'])
    for cell,src in [('E6',['AWAL','BELI','KIRIM','PAKAI','RUSAK','KOREKSI MASUK','KOREKSI KELUAR']),('B8','=KodeBahan'),('E10','=MetodeBayar'),('B12','=KodeLokasi'),('E21','=JenisKas'),('E23','=MetodeBayar')]:b.validate(s,cell,src)
    s.merge_range('A29:L32','KIRIM membutuhkan ID transfer unik (boleh dikosongkan agar macro membuat ID), tujuan berbeda dari lokasi ini, dan stok cukup. TERIMA hanya dari CSV atau jurnal penerimaan aktual. Jangan menambah TERIMA ke file pengirim. Transfer tidak dicatat sebagai pendapatan.',f['warn']);s.print_area('A1:L32')

    if role=='MEDOKAN':build_pos(b,p,today)
    elif role=='KARAH':
        s=b.base('REKAP MAJOO','Isi dari laporan final Majoo; satu ID laporan/hari unik. Netto = bruto − diskon dan harus sama dengan jumlah penerimaan per metode. Biaya payment gateway dicatat terpisah.',18)
        rf={'M':lambda r:f'IF(A{r}="","",N(C{r})-N(D{r}))','N':lambda r:f'IF(A{r}="","",SUM(E{r}:K{r}))','O':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("REKAP MAJOO","A")},A{r})<>1','ID GANDA'),(f'COUNTIFS({rng("REKAP MAJOO","B")},B{r},{rng("REKAP MAJOO","L")},"AKTIF")>1','TANGGAL GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'OR(COUNT(C{r}:K{r})<>9,MIN(C{r}:K{r})<0,D{r}>C{r})','CEK NOMINAL'),(f'AND(L{r}<>"AKTIF",L{r}<>"VOID")','CEK STATUS'),(f'ABS(M{r}-N{r})>0.5','SELISIH METODE')]),'P':lambda r:f'IF(AND(A{r}<>"",O{r}="OK",L{r}="AKTIF"),M{r},0)','Q':lambda r:f'IF(AND(A{r}<>"",O{r}="OK",L{r}="AKTIF"),N(E{r}),0)'}
        b.table(s,'tblRekap',['ID Laporan','Tanggal','Bruto','Diskon','Tunai','QRIS','Transfer','GoPay','OVO','DANA','Lainnya','Status','Netto','Total Metode','Cek','Netto Efektif','Tunai Efektif','Catatan'],1000,rf,'ABCDEFGHIJKLR',{'B':'in_date',**{c:'in_money' for c in 'CDEFGHIJK'},**{c:'money' for c in 'MNPQ'}});b.audit(s,'O',1006);b.validate(s,'L7:L1006',['AKTIF','VOID'])
        s=b.base('ACUAN MEDOKAN','Katalog ini hanya referensi Medokan. Bukan harga atau resep aktif Karah; tidak digunakan untuk menghitung penjualan/pemakaian Karah.')
        s.write_row('A6',['SKU','Menu Medokan','Varian','Kategori','Harga Acuan'],f['head'])
        for r,m in enumerate(p['menus'],6):s.write_row(r,0,[m['sku'],m['name'],m['size'],m['category'],m['price']],f['calc'])
        s.set_column('B:B',37);s.print_area('A1:E55')
    else:
        s=b.base('PRODUKSI','Rencana production house: input/output batch aktual. Tidak ada BOM produksi yang dibuat-buat. Isi bahan dan hasil dalam satuan dasar; kedua sisi disimpan dengan ID dokumen sama.')
        s.write('A6','Tanggal',f['label']);s.write_datetime('B6',today,f['in_date']);s.write('D6','ID Dokumen',f['label']);s.write_blank('E6',None,f['input'])
        s.write_row('A10',['Bahan Dipakai','Qty','Satuan'],f['head']);s.write_row('E10',['Hasil Produksi','Qty','Satuan'],f['head'])
        for r in range(11,23):
            for c in ['A','B','E','F']:s.write_blank(c+str(r),None,f['input'])
            for c,ref in [('C','A'),('G','E')]:b.formula(s,c+str(r),f'IF({ref}{r}="","",{look(f"{ref}{r}","BAHAN","D")})')
        b.validate(s,'A11:A22','=KodeBahan');b.validate(s,'E11:E22','=KodeBahan');s.merge_range('A26:L29','Minimal satu bahan dipakai dan satu hasil. Jangan menjumlahkan gram dan pcs untuk menentukan hasil. Setiap output harus sudah terdaftar pada BAHAN dengan satuan yang benar. Harga/HPP hasil produksi perlu diisi dari biaya aktual, tidak dihitung dari resep yang belum ada.',f['note']);s.print_area('A1:L29')

    s=b.base('DASHBOARD',f'{p["locations"][role]} • Saldo kumulatif sampai tanggal akhir; laporan kas per periode. Stok perlu dilengkapi bila ada nota dengan resep belum disetujui.')
    s.write('A5','Awal',f['label']);s.write_datetime('B5',today.replace(day=1),f['in_date']);s.write('D5','Akhir',f['label']);s.write_datetime('E5',today,f['in_date']);w.define_name('TglAwal',"='DASHBOARD'!$B$5");w.define_name('TglAkhir',"='DASHBOARD'!$E$5")
    def per(sheet,c='B'):return f'{rng(sheet,c)},">="&TglAwal,{rng(sheet,c)},"<="&TglAkhir'
    sales=f'SUMIFS({rng("NOTA","P")},{per("NOTA")})' if role=='MEDOKAN' else (f'SUMIFS({rng("REKAP MAJOO","P")},{per("REKAP MAJOO")})' if role=='KARAH' else '0')
    cards=[('PENJUALAN NETTO',sales),('PEMBELIAN BAHAN',f'SUMIFS({rng("MUTASI","S")},{per("MUTASI")})'),('BIAYA USAHA',f'-SUMIFS({rng("KEUANGAN","I")},{rng("KEUANGAN","C")},"BIAYA USAHA",{per("KEUANGAN")})'),('BARIS KIRIM',f'COUNTIFS({rng("MUTASI","C")},"KIRIM",{rng("MUTASI","M")},"AKTIF",{rng("MUTASI","Q")},"OK",{per("MUTASI")})'),('TERKONFIRMASI TANPA SELISIH',f'COUNTIFS({rng("KONFIRMASI","K")},"OK",{per("KONFIRMASI","I")})'),('SKU BERSALDO',f'COUNTIF({rng("BAHAN","L")},">0")')]
    for i,(label,formula) in enumerate(cards):
        rr=6+i//3*5;cc=i%3*4;s.merge_range(rr,cc,rr,cc+2,label,f['kpi_label']);s.merge_range(rr+1,cc,rr+2,cc+2,'',f['kpi']);s.write_formula(rr+1,cc,'='+formula,f['kpi'],0)
    s.merge_range('A17:L17','PERIKSA DATA DAN TRANSFER',f['head'])
    alerts=[('Mutasi tidak valid',f'COUNTIFS({rng("MUTASI","A")},"<>",{rng("MUTASI","Q")},"<>OK")'),('Saldo historis negatif',f'COUNTIF({rng("MUTASI","R")},"<-0.000001")'),('Konfirmasi selisih',f'COUNTIF({rng("KONFIRMASI","K")},"SELISIH")')]
    if role=='MEDOKAN':alerts += [('Nota stok belum lengkap',f'COUNTIFS({rng("NOTA","A")},"<>",{rng("NOTA","H")},"AKTIF",{rng("NOTA","Q")},"<>LENGKAP")')]
    if role=='KARAH':alerts += [('Rekap bermasalah',f'COUNTIFS({rng("REKAP MAJOO","A")},"<>",{rng("REKAP MAJOO","O")},"<>OK")')]
    for i,(label,formula) in enumerate(alerts):s.write(18+i,0,label,f['label']);s.write_formula(18+i,3,'='+formula,f['qty'],0)
    s.merge_range('F19:L23','Tidak ada total kuantitas campuran gram/pcs/botol. Transfer hanya stok antar lokasi, bukan omzet. Gudang tidak mengimpor penjualan outlet secara otomatis. Konfirmasi tidak memotong atau menambah stok kedua kali.',f['warn'])
    s.write_row('A25',['Kode','Bahan','Satuan','Saldo','Status'],f['head'])
    for r in range(26,106):
        mr=r-19
        for c,src in [('A','A'),('B','B'),('C','D'),('D','L'),('E','M')]:b.formula(s,c+str(r),f'IF(BAHAN!A{mr}="","",BAHAN!{src}{mr})','qty' if c=='D' else 'calc')
    s.set_column('A:A',25);s.set_column('B:B',34);s.set_column('E:E',27);s.print_area('A1:Q61')
    if role in ['MEDOKAN','KARAH']:
        s.write_row('G25',['Metode','Penerimaan Netto'],f['head'])
        for r,method in enumerate(METHODS,26):
            s.write('G'+str(r),method,f['calc'])
            if role=='MEDOKAN':expr=f'SUMIFS({rng("NOTA","P")},{rng("NOTA","E")},G{r},{per("NOTA")})'
            else:expr=f'SUMIFS({rng("REKAP MAJOO",col(4+r-26))},{rng("REKAP MAJOO","L")},"AKTIF",{rng("REKAP MAJOO","O")},"OK",{per("REKAP MAJOO")})'
            b.formula(s,'H'+str(r),expr,'money')
        chart=w.add_chart({'type':'column'});chart.add_series({'name':'Netto per metode','categories':"='DASHBOARD'!$G$26:$G$32",'values':"='DASHBOARD'!$H$26:$H$32",'fill':{'color':'#9A3026'},'border':{'none':True}});chart.set_title({'name':'Penerimaan per metode'});chart.set_legend({'none':True});s.insert_chart('G35',chart,{'x_scale':1.1,'y_scale':1.1})
    else:
        s.write_row('G25',['Tujuan','Baris Pengiriman'],f['head'])
        for r,loc in [(26,'MEDOKAN'),(27,'KARAH')]:s.write('G'+str(r),loc,f['calc']);b.formula(s,'H'+str(r),f'COUNTIFS({rng("MUTASI","C")},"KIRIM",{rng("MUTASI","J")},G{r},{rng("MUTASI","M")},"AKTIF",{rng("MUTASI","Q")},"OK",{per("MUTASI")})','qty')
        chart=w.add_chart({'type':'column'});chart.add_series({'name':'Baris pengiriman','categories':"='DASHBOARD'!$G$26:$G$27",'values':"='DASHBOARD'!$H$26:$H$27",'fill':{'color':'#21624C'}});chart.set_title({'name':'Distribusi ke outlet'});chart.set_legend({'none':True});s.insert_chart('G30',chart)
    if role=='MEDOKAN':
        s.write_row('G65',['SKU','Menu','Qty','Bruto','Varian/Kategori'],f['head'])
        for r in range(66,146):
            mr=r-59;b.formula(s,'G'+str(r),f'IF(MENU!A{mr}="","",MENU!A{mr})');b.formula(s,'H'+str(r),f'IF(G{r}="","",MENU!B{mr})');b.formula(s,'I'+str(r),f'IF(G{r}="","",SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},G{r},{per("ITEM","I")}))','qty');b.formula(s,'J'+str(r),f'IF(G{r}="","",SUMIFS({rng("ITEM","K")},{rng("ITEM","C")},G{r},{rng("ITEM","N")},">0",{per("ITEM","I")}))','money');b.formula(s,'K'+str(r),f'IF(G{r}="","",IF(MENU!C{mr}="",MENU!D{mr},MENU!C{mr}))')
        s.write_row('L25',['Varian/Kategori','Qty','Bruto'],f['head'])
        for r,size in enumerate(['L','M','S','Burger','Minuman'],26):s.write('L'+str(r),size,f['calc']);b.formula(s,'M'+str(r),f'SUMIF($K$66:$K$145,L{r},$I$66:$I$145)','qty');b.formula(s,'N'+str(r),f'SUMIF($K$66:$K$145,L{r},$J$66:$J$145)','money')

    s=b.base('TUTUP KASIR','Rekonsiliasi uang fisik. Penjualan tunai netto, pembelian dan kas lain saja; QRIS/GoPay/transfer tidak masuk laci. Saldo pembukaan jangan dicatat dua kali.')
    for cell,label in [('A5','Tanggal'),('A7','Modal Awal Laci'),('D7','Uang Fisik')]:s.write(cell,label,f['label'])
    s.write_datetime('B5',today,f['in_date']);s.write_blank('B7',None,f['in_money']);s.write_blank('E7',None,f['in_money'])
    cash=f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5,{rng("NOTA","E")},"Tunai")' if role=='MEDOKAN' else (f'SUMIFS({rng("REKAP MAJOO","Q")},{rng("REKAP MAJOO","B")},$B$5)' if role=='KARAH' else '0')
    for r,label,formula in [(10,'Penjualan Tunai Netto',cash),(11,'Pembelian Tunai',f'SUMIFS({rng("MUTASI","S")},{rng("MUTASI","B")},$B$5,{rng("MUTASI","G")},"Tunai")'),(12,'Kas Lain Bertanda',f'SUMIFS({rng("KEUANGAN","I")},{rng("KEUANGAN","B")},$B$5,{rng("KEUANGAN","E")},"Tunai")'),(14,'Seharusnya di Laci','N(B7)+B10-B11+B12'),(15,'Selisih','IF(ISNUMBER(E7),E7-B14,"")')]:s.write('A'+str(r),label,f['label']);b.formula(s,'B'+str(r),formula,'money')
    s.set_column('A:A',30);s.print_area('A1:L18')

    s=b.base('TRANSFER','Pengiriman/penerimaan CSV offline. Tidak menggabungkan tiga file, tidak membuat tautan Excel eksternal, tidak mengirim data melalui jaringan.')
    steps=['1. Catat KIRIM di file pengirim: ID unik, asal lokasi ini, tujuan, bahan, qty dan tanggal.', '2. Ekspor CSV KIRIM, serahkan ke operator tujuan lewat cara yang disepakati. Ini bukan sinkronisasi otomatis.', '3. Import CSV di file tujuan dengan tanggal terima aktual. Impor ulang yang identik dilewati; ID/data konflik ditolak tanpa menimpa catatan.', '4. Jika jumlah fisik berbeda, ubah qty TERIMA sebelum ekspor konfirmasi; Qty Kirim sumber tetap dipertahankan. Jangan menutupi selisih dengan konfirmasi palsu.', '5. Ekspor CSV TERIMA, impor sebagai KONFIRMASI pada pengirim. Ack tidak menjadi mutasi masuk. Selisih penerimaan ditandai SELISIH.', '6. Retur dikirim sebagai transfer baru dengan asal/tujuan terbalik dan ID baru. Jangan menghapus pengiriman yang sudah diterima di lokasi lain.']
    for i,text in enumerate(steps):s.merge_range(6+i*3,0,7+i*3,11,text,f['body'])
    s.merge_range('A27:L30','CLI pengembang: export-shipments / import-shipments / export-receipts / import-receipts. Schema CSV PF1 memeriksa kode lokasi, bahan, satuan, tanggal, kuantitas dan ID. Opsi --received menampung jumlah terima aktual berbeda; lihat panduan. File keluaran baru wajib berbeda agar sumber tidak tertimpa.',f['note']);s.print_area('A1:L30')

    s=b.base('CONFIG','Identitas file dan sumber. Jangan mengganti lokasi file yang sudah dipakai. Tiga lokasi harus memakai kode bahan/satuan yang sama untuk pertukaran CSV.')
    for cell,label,value in [('A6','Kode Lokasi',role),('A7','Schema','PF1'),('A8','Pemilik','Bu Elok'),('A9','Jam','09.00–21.00'),('A10','Alamat Medokan',p['medokan_address'] if role=='MEDOKAN' else ''),('A12','Info Pembayaran','')]:s.write(cell,label,f['label']);s.write('B'+cell[1:],value,f['input'])
    s.write_url('A15',p['source_sheet'],f['nav'],'MoM Pizza Fuzta • 27 Oktober 2026')
    s.merge_range('A18:L21','Sumber XLSM memuat beberapa nama resep berbeda dari katalog, teks qty desimal koma, harga ukuran tidak berurutan, resep tanpa menu dan catatan sisa sate. Audit/alias disimpan pada data/audit_sumber.json. Alias eksplisit adalah pemetaan kandidat, bukan persetujuan resep. Semua BOM/menu awalnya TIDAK disetujui.',f['warn'])
    s.merge_range('A23:L26','Belum ada format ekspor/API Majoo, data bahan gudang, stok nyata, resep produksi house atau konfigurasi harga Karah. Karah merekap laporan final secara manual. Tidak ada pajak/komisi tersembunyi; biaya payment gateway dicatat terpisah sesuai laporan aktual.',f['note'])
    lists=[('N','JenisGerak',KINDS),('O','MetodeBayar',METHODS),('P','JenisKas',['BIAYA USAHA','PRIVE','MODAL MASUK','PENDAPATAN LAIN']),('Q','KodeLokasi',['MEDOKAN','KARAH','GUDANG'])]
    for c,name,vals in lists:s.write_column(c+'2',vals);w.define_name(name,f"='CONFIG'!${c}$2:${c}${len(vals)+1}")
    s.set_column('A:A',29);s.set_column('B:B',44);s.set_column('N:Q',None,None,{'hidden':True});s.print_area('A1:L26')
    order=['MULAI','DASHBOARD']+(['KASIR'] if role=='MEDOKAN' else ['REKAP MAJOO'] if role=='KARAH' else ['PRODUKSI'])+['FORM','BAHAN']+(['MENU','BOM','RESEP','NOTA','ITEM','STRUK'] if role=='MEDOKAN' else ['ACUAN MEDOKAN'] if role=='KARAH' else [])+['MUTASI','KEUANGAN','TUTUP KASIR','TRANSFER','PERMINTAAN','KONFIRMASI','CONFIG']
    b.close(order);return path

def build_pos(b,p,today):
    f=b.f
    s=b.base('BOM','Draf dari sumber; alias eksplisit diaudit. Persetujuan YA diberikan setelah bahan, satuan dan qty dikonfirmasi. Tidak memakai resep yatim tanpa menu aktif.',12)
    bf={'G':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeMenu,A{r})<>1','CEK SKU'),(f'COUNTIF(KodeBahan,B{r})<>1','CEK BAHAN'),(f'OR(NOT(ISNUMBER(C{r})),C{r}<=0)','CEK QTY'),(f'D{r}<>{look(f"B{r}","BAHAN","D")}','SATUAN BERBEDA'),(f'AND(E{r}<>"YA",E{r}<>"TIDAK")','ISI PERSETUJUAN')])}
    b.table(s,'tblBOM',['SKU','Kode Bahan','Qty / Item','Satuan','Disetujui','Catatan','Cek'],600,bf,'ABCDEF',{'C':'in_qty'})
    for r,x in enumerate(p['bom'],7):s.write_row(r-1,0,[x['sku'],x['material'],x['qty'],x['unit'],'TIDAK',f'Sumber baris {x["row"]}; verifikasi sebelum setuju'],f['input'])
    b.validate(s,'A7:A606','=KodeMenu');b.validate(s,'B7:B606','=KodeBahan');b.validate(s,'E7:E606',['YA','TIDAK']);b.audit(s,'G',606)
    s=b.base('MENU','49 menu sumber + kapasitas 80. Harga anomali tidak diperbaiki dengan asumsi. Semua resep awalnya belum disetujui; gunakan harga/takaran yang dikonfirmasi.',13)
    mf={'H':lambda r:f'IF(A{r}="","",RESEP!CE{r})','I':lambda r:f'IF(A{r}="","",RESEP!CF{r})','K':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeMenu,A{r})<>1','SKU GANDA'),(f'B{r}=""','ISI NAMA'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0)','ISI HARGA'),(f'AND(F{r}<>"AKTIF",F{r}<>"NONAKTIF")','CEK STATUS'),(f'AND(G{r}<>"YA",G{r}<>"TIDAK")','CEK PERSETUJUAN')]),'L':lambda r:f'IF(A{r}="","",IF(AND(G{r}="YA",COUNTIF({rng("BOM","A")},A{r})>0,COUNTIFS({rng("BOM","A")},A{r},{rng("BOM","E")},"<>YA")=0,COUNTIFS({rng("BOM","A")},A{r},{rng("BOM","G")},"<>OK")=0),"LENGKAP","BELUM LENGKAP"))'}
    b.table(s,'tblMenu',['SKU','Menu','Varian','Kategori','Harga','Status','Resep Disetujui','HPP Acuan','Status HPP','Catatan Harga','Cek','Status Resep'],80,mf,'ABCDEFGJ',{'E':'in_money','H':'money'})
    for r,m in enumerate(p['menus'],7):
        note='Harga acuan sumber; konfirmasi harga aktual.'
        if 'Asumsi' in m['price_note'] or m['name'] in ['Cheese Lovers S','Smoked Beef M','Smoked Beef S']:note='PERLU KONFIRMASI: catatan/harga ukuran sumber tidak konsisten.'
        s.write_row(r-1,0,[m['sku'],m['name'],m['size'],m['category'],m['price'],'AKTIF','TIDAK'],f['input']);s.write('J'+str(r),note,f['input'])
    b.validate(s,'F7:F86',['AKTIF','NONAKTIF']);b.validate(s,'G7:G86',['YA','TIDAK']);s.set_column('B:B',33);s.set_column('J:J',55);b.audit(s,'K',86)
    s=b.base('RESEP','Matriks takaran draf per item. Hanya resep LENGKAP yang boleh memotong stok. HPP tanpa harga semua bahan tidak disebut lengkap.',84)
    headers=['SKU','Menu']+[p['materials'][i]['code'] if i<len(p['materials']) else f'Cadangan {i+1}' for i in range(80)]+['HPP','Status HPP']
    for i in range(80):c=col(i+2);mr=i+7;b.formula(s,c+'4',f'IF(BAHAN!K{mr}="OK",0,1)');b.formula(s,c+'5',f'N(BAHAN!G{mr})','qty')
    s.set_row(3,None,None,{'hidden':True})
    rf={'A':lambda r:f'IF(MENU!A{r}="","",MENU!A{r})','B':lambda r:f'IF(A{r}="","",MENU!B{r})','CE':lambda r:f'IF(A{r}="","",SUMPRODUCT(C{r}:CD{r},$C$5:$CD$5))','CF':lambda r:f'IF(A{r}="","",IF(OR(MENU!L{r}<>"LENGKAP",SUMPRODUCT(--(C{r}:CD{r}>0),$C$4:$CD$4)>0),"BELUM LENGKAP","LENGKAP"))'}
    for i in range(80):
        c=col(i+2);mr=i+7;rf[c]=lambda r,mr=mr:f'IF(A{r}="","",IF(BAHAN!A{mr}="",0,SUMIFS({rng("BOM","C")},{rng("BOM","A")},A{r},{rng("BOM","B")},BAHAN!A{mr})))'
    b.table(s,'tblResep',headers,80,rf)
    s=b.base('NOTA','Satu nota, diskon dan pembayaran sekali. Status stok snapshot membedakan kasir yang tercatat dari pemakaian bahan yang belum lengkap.',18)
    nf={'J':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","K")},{rng("ITEM","B")},A{r}))','K':lambda r:f'IF(A{r}="","",J{r}-N(F{r}))','L':lambda r:f'IF(A{r}="","",IF(E{r}="Tunai",N(G{r})-K{r},0))','M':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","L")},{rng("ITEM","B")},A{r}))','N':lambda r:f'IF(A{r}="","",IF(COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","G")},"<>LENGKAP")>0,"BELUM LENGKAP","LENGKAP"))','O':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeNota,A{r})<>1','NOTA GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY(),NOT(ISNUMBER(C{r})),C{r}<0,C{r}>=1)','CEK TANGGAL/JAM'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(H{r}<>"AKTIF",H{r}<>"VOID")','CEK STATUS'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<0,F{r}>J{r})','CEK DISKON'),(f'J{r}<=0','ISI ITEM'),(f'COUNTIF({rng("ITEM","B")},A{r})>12','MAKS 12 ITEM'),(f'COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","M")},"<>OK")>0','CEK ITEM'),(f'OR(NOT(ISNUMBER(G{r})),G{r}<K{r},AND(E{r}<>"Tunai",G{r}<>K{r}))','CEK PEMBAYARAN'),(f'AND(Q{r}<>"LENGKAP",Q{r}<>"BELUM LENGKAP")','ISI STATUS STOK'),(f'AND(Q{r}="LENGKAP",COUNTIFS({rng("MUTASI","L")},A{r},{rng("MUTASI","C")},"PAKAI",{rng("MUTASI","M")},"AKTIF")=0)','ISI PEMAKAIAN')]),'P':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",O{r}="OK"),K{r},0)'}
    b.table(s,'tblNota',['No Nota','Tanggal','Jam','Kasir','Metode','Diskon','Diterima','Status','Catatan','Bruto','Netto','Kembalian','HPP Total','Status HPP','Cek','Netto Efektif','Status Stok'],1000,nf,'ABCDEFGHIQ',{'B':'in_date','C':'in_time','F':'in_money','G':'in_money',**{c:'money' for c in 'JKLMP'}});b.audit(s,'O',1006);b.validate(s,'E7:E1006','=MetodeBayar');b.validate(s,'H7:H1006',['AKTIF','VOID']);b.validate(s,'Q7:Q1006',['LENGKAP','BELUM LENGKAP'])
    s=b.base('ITEM','Satu baris per menu/nota; harga dan biaya berupa snapshot tetap. Tidak dibatasi 20 menu. Jangan mengubah transaksi final untuk mengganti harga katalog.',14)
    itf={'I':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","B","0")})','J':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","H")})','K':lambda r:f'IF(AND(A{r}<>"",M{r}="OK"),N(D{r})*N(E{r}),0)','L':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",G{r}="LENGKAP"),N(D{r})*N(F{r}),0)','M':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("ITEM","A")},A{r})<>1','ID GANDA'),(f'COUNTIF(KodeNota,B{r})<>1','CEK NOTA'),(f'COUNTIF(KodeMenu,C{r})<>1','CEK SKU'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0,D{r}<>INT(D{r}))','CEK QTY'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0)','CEK HARGA'),(f'AND(G{r}<>"LENGKAP",G{r}<>"BELUM LENGKAP")','CEK STATUS HPP'),(f'AND(G{r}="LENGKAP",OR(NOT(ISNUMBER(F{r})),F{r}<0))','ISI HPP')]),'N':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",J{r}="AKTIF",{look(f"B{r}","NOTA","O")}="OK"),N(D{r}),0)'}
    b.table(s,'tblItem',['ID Item','No Nota','SKU','Qty','Harga Snapshot','HPP Snapshot','Status HPP','Catatan','Tanggal','Status Nota','Bruto','HPP Total','Cek','Qty Efektif'],3000,itf,'ABCDEFGH',{'D':'in_qty','E':'in_money','F':'in_money','I':'date','K':'money','L':'money','N':'qty'});b.audit(s,'M',3006)
    s=b.base('KASIR','Medokan • Pizza L/M/S, burger dan minuman. Metode GoPay tersedia. Penjualan resep belum disetujui tetap tercatat, tetapi tidak dianggap pemakaian stok lengkap.')
    for cell,text in [('A6','Tanggal'),('D6','Jam'),('G6','Metode'),('A8','Kasir')]:s.write(cell,text,f['label'])
    s.write_datetime('B6',today,f['in_date']);s.write('E6',9/24,f['in_time']);s.write('H6','Tunai',f['input']);s.write_blank('B8',None,f['input']);b.validate(s,'H6','=MetodeBayar')
    s.write_row('A12',['No','SKU','Menu','Qty','Harga','Bruto','Status Resep','Status HPP'],f['head']);s.set_column('C:C',34)
    for r in range(13,25):
        s.write('A'+str(r),r-12,f['calc']);s.write_blank('B'+str(r),None,f['input']);s.write_blank('D'+str(r),None,f['in_qty'])
        for c,src in [('C','B'),('E','E'),('G','L'),('H','I')]:b.formula(s,c+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU",src,"0" if c=="E" else chr(34)+chr(34))})','in_money' if c=='E' else 'calc')
        b.formula(s,'F'+str(r),f'IF(B{r}="","",N(D{r})*N(E{r}))','money')
    b.validate(s,'B13:B24','=KodeMenu')
    for r,label in [(26,'Bruto'),(27,'Diskon Rp'),(28,'TOTAL'),(29,'Diterima'),(30,'Kembalian')]:s.write('E'+str(r),label,f['label'])
    b.formula(s,'F26','SUM(F13:F24)','money',0);s.write('F27',0,f['in_money']);b.formula(s,'F28','F26-N(F27)','money',0);s.write_blank('F29',None,f['in_money']);b.formula(s,'F30','IF(H6="Tunai",N(F29)-F28,0)','money',0);s.merge_range('A33:H35','Semua resep awalnya perlu persetujuan. Nota dengan status stok BELUM LENGKAP belum memotong bahan; sebelum memakai saldo stok untuk belanja, lengkapi pemakaian aktual satu kali dan tandai nota LENGKAP. Harga bahan kosong selalu berarti HPP belum lengkap.',f['warn']);s.print_area('A1:L35')
    s=b.base('STRUK','Struk Medokan; pilih nota AKTIF dan valid. Harga/diskon berasal dari snapshot, bukan harga katalog terbaru.')
    s.set_portrait();s.write('A5','No Nota',f['label']);s.write_blank('B5',None,f['input']);b.validate(s,'B5','=KodeNota');s.merge_range('A7:F7','PIZZA FUZTA MEDOKAN',f['title']);b.formula(s,'A8','IF(B5="","PILIH NOTA",B5)');b.formula(s,'D8',look('$B$5','NOTA','B'),'date');b.formula(s,'F8',look('$B$5','NOTA','C'),'time');b.formula(s,'A10',f'IF(B5="","PILIH NOTA",IF(AND({look("B5","NOTA","H")}="AKTIF",{look("B5","NOTA","O")}="OK"),"SIAP CETAK","NOTA VOID / TIDAK VALID"))','warn');s.write_row('A12',['No','SKU','Menu','Qty','Harga','Jumlah'],f['head'])
    for r in range(13,25):
        b.formula(s,'G'+str(r),f'IF($B$5="","",IFERROR(_xlfn.AGGREGATE(15,6,(ROW({rng("ITEM","B")})-6)/({rng("ITEM","B")}=$B$5),{r-12}),""))')
        b.formula(s,'A'+str(r),f'IF(G{r}="","",{r-12})')
        for c,src in [('B','C'),('D','D'),('E','E'),('F','K')]:b.formula(s,c+str(r),f'IF($G{r}="","",INDEX({rng("ITEM",src)},$G{r}))','money' if c in ['E','F'] else 'calc')
        b.formula(s,'C'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","B")})')
    for r,label,src in [(26,'Bruto','J'),(27,'Diskon','F'),(28,'TOTAL','K'),(29,'Diterima','G'),(30,'Kembalian','L')]:s.write('D'+str(r),label,f['label']);b.formula(s,'F'+str(r),f'IF($B$5="","",{look("$B$5","NOTA",src,"0")})','money')
    s.set_column('C:C',34);s.set_column('G:G',None,None,{'hidden':True});s.print_area('A7:F32');s.repeat_rows(6,11);s.fit_to_pages(1,1)

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--role',choices=list(FILES));a=parser.parse_args()
    for role in [a.role] if a.role else FILES:print(build(role,ROOT/'output'/FILES[role]))
