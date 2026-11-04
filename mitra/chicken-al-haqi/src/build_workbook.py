"""A small cashier workbook: cuts + flavors, actual pcs stock, separate business cash."""
import argparse,datetime as dt,json
from pathlib import Path
from layout import Layout,CAP,rng,look,checks,col
ROOT=Path(__file__).resolve().parents[1]
CORR=lambda c:f"'STOK'!${c}$36:${c}$235"
FLAV=lambda c:f"'SETUP'!${c}$31:${c}$34"
METHODS=['Tunai','QRIS','Transfer','Lainnya']
CASHKINDS=['BIAYA USAHA','MODAL MASUK','PRIVE','KE USAHA LAIN','DARI USAHA LAIN']
def flavor(value,c,default='""'):return f'IFERROR(INDEX({FLAV(c)},MATCH({value},{FLAV("A")},0)),{default})'
def stock_at(code,when):
    opening=look(code,'STOK','C','0');qty=look(code,'STOK','D','0')
    return (f'N({qty})+SUMIFS({rng("BELANJA","M")},{rng("BELANJA","C")},{code},{rng("BELANJA","B")},">="&{opening},{rng("BELANJA","B")},"<="&{when})'
        f'+SUMIFS({CORR("I")},{CORR("C")},{code},{CORR("B")},">="&{opening},{CORR("B")},"<="&{when})'
        f'-SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},{code},{rng("ITEM","I")},">="&{opening},{rng("ITEM","I")},"<="&{when})')

def build(path):
    p=json.loads((ROOT/'data/profile.json').read_text());path.parent.mkdir(parents=True,exist_ok=True)
    b=Layout(path);w=b.w;f=b.f;today=dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    for name,sheet,c in [('KodeMenu','SETUP','A'),('KodeNota','NOTA','A')]:w.define_name(name,'='+rng(sheet,c))
    w.define_name('PotonganAyam',"='SETUP'!$A$7:$A$10");w.define_name('RasaAyam','='+FLAV('A'))
    w.define_name('TglLaporan',"='HARIAN'!$B$5")

    s=b.base('MULAI','Mba Gita • MoM 23 Oktober 2026 • Kasir sederhana, belanja kg → pcs aktual, stok potongan dan uang usaha ayam.')
    s.merge_range('A6:L6','CUKUP MULAI DARI KASIR DAN HARIAN',f['head'])
    steps=[('1. SETUP sekali','Harga sumber: sayap/paha bawah Rp8.000; paha atas/dada Rp9.000; tambahan saus Rp3.000; nasi Rp3.000. Periksa harga terkini. HPP sumber adalah acuan belum dikonfirmasi.'),('2. STOK awal','Isi tanggal dan jumlah potongan fisik pada STOK. Jika belum diketahui, biarkan keduanya kosong: stok ditampilkan BELUM AWAL, bukan angka nol yang dianggap pasti.'),('3. Catat kasir','Pilih potongan, rasa dan qty. Nasi diisi terpisah. Diskon satu kali per nota; masukkan uang diterima untuk Tunai. Tombol memerlukan impor VBA.'),('4. Catat belanja','Masukkan kg jika diketahui, jumlah pcs AKTUAL dan nominal total. Perkiraan 9/10 pcs per kg hanya pembanding; tidak ditambahkan lagi sebagai stok.'),('5. Pisahkan uang','KAS USAHA memisahkan biaya usaha ayam, modal, prive dan perpindahan ke/dari usaha lain seperti obat. Tidak ada produk obat di kasir ayam.'),('6. Tutup hari','HARIAN: pilih tanggal, isi kas awal ayam dan uang fisik khusus usaha ayam. Periksa selisih tunai, stok minus dan baris yang belum valid; simpan Ctrl+S.')]
    for i,(label,text) in enumerate(steps):r=7+i*3;s.merge_range(r,0,r,2,label,f['label']);s.merge_range(r,3,r+1,11,text,f['body'])
    s.merge_range('A28:L30','JAM: hari biasa 09.00–21.00; weekend mulai 06.30/07.00. Jam pagi weekend diterima. Pengingat dari MoM: siapkan belanja/pesanan pemasok sebelum lewat 07.00 agar potongan tidak kehabisan.',f['note'])
    s.merge_range('A32:L35','TOMBOL: simpan salinan .xlsm, Alt+F11 > File > Import > Modul_ChickenAlHaqi.bas; Debug > Compile VBAProject; Alt+F8 > PasangTombol. Jurnal NOTA/ITEM disembunyikan agar tampilan ringan dan dapat dibuka untuk audit/manual. VBA belum diuji di Excel desktop.',f['warn'])
    s.merge_range('A37:L40','Tidak ada transaksi, uang fisik atau stok fiktif. Ayam dilacak dalam pcs; tidak memisahkan tahap mentah/matang. Nasi awalnya tidak dilacak karena resep konversi beras belum tersedia. Tanpa stok awal, pencatatan uang bisa berjalan tetapi kecukupan stok belum dapat diverifikasi.',f['warn']);s.print_area('A1:L40')

    s=b.base('SETUP','Lima menu dasar + empat rasa. Harga dari file mitra, perkiraan hasil kg dari MoM. HPP tidak dianggap lengkap sebelum dikonfirmasi.',12)
    mf={'J':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeMenu,A{r})<>1','KODE GANDA'),(f'OR(B{r}="",D{r}="")','ISI MENU/SATUAN'),(f'OR(NOT(ISNUMBER(C{r})),C{r}<=0)','ISI HARGA'),(f'AND(G{r}<>"YA",G{r}<>"TIDAK")','CEK HPP'),(f'AND(G{r}="YA",OR(NOT(ISNUMBER(F{r})),F{r}<0))','ISI HPP'),(f'AND(H{r}<>"YA",H{r}<>"TIDAK")','CEK LACAK'),(f'AND(I{r}<>"AKTIF",I{r}<>"NONAKTIF")','CEK STATUS')])}
    b.table(s,'tblMenu',['Kode','Menu','Harga Dasar','Satuan','Perkiraan Pcs/Kg','HPP Acuan/Pcs','HPP Dikonfirmasi','Lacak Stok','Status','Cek','Catatan'],20,mf,'ABCDEFGHIK',{'C':'in_money','E':'in_qty','F':'in_money'})
    for r,m in enumerate(p['items'],7):
        for c,v in [(0,m['code']),(1,m['name']),(2,m['price']),(3,m['unit']),(4,m['estimate_per_kg']),(6,'TIDAK'),(7,m['track']),(8,'AKTIF'),(10,'HPP umum sumber, bukan biaya lengkap; konfirmasi per potongan.' if m['code']!='NASI' else 'Harga diketahui; HPP/resep nasi belum tersedia.')]:
            if v is not None:s.write(r-1,c,v,f['input'])
        if m['code']!='NASI':b.formula(s,'F'+str(r),'SUM($C$51:$C$54)','in_money',7000)
    b.audit(s,'J',26);s.set_column('B:B',26);s.set_column('K:K',55)
    for c in ['G','H']:b.validate(s,c+'7:'+c+'26',['YA','TIDAK'])
    b.validate(s,'I7:I26',['AKTIF','NONAKTIF'])
    s.write_row('A30',['Rasa','Tambahan Harga','HPP Tambahan','HPP Dikonfirmasi','Cek'],f['head'])
    for r,x in enumerate(p['flavors'],31):
        for c,v in [(0,x['name']),(1,x['extra']),(2,x['hpp']),(3,x['confirmed'])]:
            if v is not None:s.write(r-1,c,v,f['input'])
        b.formula(s,'E'+str(r),f'IF(OR(COUNTIF($A$31:$A$34,A{r})<>1,NOT(ISNUMBER(B{r})),B{r}<0,AND(D{r}<>"YA",D{r}<>"TIDAK"),AND(D{r}="YA",OR(NOT(ISNUMBER(C{r})),C{r}<0))),"CEK RASA","OK")')
    b.validate(s,'D31:D34',['YA','TIDAK'])
    s.write_row('A39',['Acuan 1 kg sumber','Nominal','Catatan'],f['head'])
    for r,(name,price) in enumerate(p['per_kg_source'],40):s.write_row(r-1,0,[name,price,'Harga/takaran sumber; bukan BOM kuantitas otomatis.'],f['calc'])
    s.merge_range('A44:L47','Harga ayam sumber Rp35.000/kg dipakai hanya sebagai pembanding nominal belanja. Racikan terigu Rp15.000 dan marinasi Rp3.000 tidak otomatis dijumlahkan ke HPP per pcs lagi: satuan/pemakaian rinci belum lengkap. Tidak ada penggandaan biaya kg dan pcs.',f['note'])
    s.write_row('A50',['No','Acuan Per Pcs Sumber','Nominal'],f['head'])
    for r,(name,value) in enumerate(p['per_piece_source'],51):s.write_row(r-1,0,[r-50,name,value],f['calc'])
    s.write('B55','Jumlah daftar per pcs',f['label']);b.formula(s,'C55','SUM(C51:C54)','money',7000)
    s.write('A59','Rencana kg ayam',f['label']);s.write('B59',1,f['in_qty'])
    for r,label,formula in [(60,'Perkiraan 9 pcs/kg','N(B59)*9'),(61,'Perkiraan 10 pcs/kg','N(B59)*10'),(62,'Batch marinasi maks 3 kg','IF(B59>0,ROUNDUP(B59/3,0),0)')]:s.write('A'+str(r),label,f['label']);b.formula(s,'B'+str(r),formula,'qty')
    s.merge_range('D59:L63','Kalkulator rencana, bukan pencatatan stok/produksi. Belanja 5 kg boleh, kemudian dibagi menjadi batch bumbu paling banyak 3 kg. Tidak menebak gram bumbu dari angka harga. Gunakan jumlah potongan nyata sebagai stok.',f['note']);s.print_area('A1:L63')

    s=b.base('BELANJA','Belanja ayam per potongan: kg opsional, pcs AKTUAL wajib. Stok hanya bertambah dari pcs aktual. Bumbu/minyak/beras masuk KAS USAHA sebagai biaya pendukung.')
    bf={'J':lambda r:f'IF(OR(A{r}="",D{r}="",D{r}=0),"",N(D{r})*{look(f"C{r}","SETUP","E","0")})','K':lambda r:f'IF(OR(A{r}="",J{r}=""),"",N(E{r})-J{r})',
        'L':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("BELANJA","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(KodeMenu,C{r})<>1','CEK MENU'),(f'AND(D{r}<>"",OR(NOT(ISNUMBER(D{r})),D{r}<0))','CEK KG'),(f'AND(C{r}="NASI",N(D{r})>0)','NASI PAKAI PORSI'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0,E{r}<>INT(E{r}))','ISI PCS AKTUAL'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<0)','CEK NOMINAL'),(f'COUNTIF(MetodeBayar,G{r})<>1','CEK METODE'),(f'AND(H{r}<>"AKTIF",H{r}<>"VOID")','CEK STATUS')]),
        'M':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",L{r}="OK"),N(E{r}),0)','N':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",L{r}="OK"),N(F{r}),0)'}
    b.table(s,'tblBelanja',['ID Belanja','Tanggal','Potongan / Nasi','Kg Opsional','Pcs/Porsi Aktual','Nominal Total','Metode','Status','Catatan Bukti','Pcs Perkiraan','Selisih Pcs','Cek','Pcs Efektif','Nominal Efektif'],1500,bf,'ABCDEFGHI',{'B':'in_date','D':'in_qty','E':'in_qty','F':'in_money','J':'qty','K':'qty','M':'qty','N':'money'})
    for c,source in [('C','=KodeMenu'),('G','=MetodeBayar'),('H',['AKTIF','VOID'])]:b.validate(s,c+'7:'+c+'1506',source)
    b.audit(s,'L',1506)

    s=b.base('KAS USAHA','Catat hanya arus uang yang memengaruhi kas ayam. Penjualan obat bukan menu/omzet ayam. Modal, prive dan antar-usaha tidak menjadi biaya operasional ayam.')
    kf={'H':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("KAS USAHA","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(JenisKas,C{r})<>1','CEK JENIS'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0)','CEK NOMINAL'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(G{r}<>"AKTIF",G{r}<>"VOID")','CEK STATUS')]),'I':lambda r:f'IF(AND(A{r}<>"",H{r}="OK",G{r}="AKTIF"),N(D{r})*IF(OR(C{r}="MODAL MASUK",C{r}="DARI USAHA LAIN"),1,-1),0)'}
    b.table(s,'tblKas',['ID Kas','Tanggal','Jenis','Nominal','Metode','Catatan','Status','Cek','Kas Bertanda'],1500,kf,'ABCDEFG',{'B':'in_date','D':'in_money','I':'money'})
    for c,source in [('C','=JenisKas'),('E','=MetodeBayar'),('G',['AKTIF','VOID'])]:b.validate(s,c+'7:'+c+'1506',source)
    b.audit(s,'H',1506);s.set_column('F:F',45)

    s=b.base('STOK','Saldo potongan dalam pcs; tidak memisahkan mentah/matang. Stok belum diisi ditampilkan kosong. Saldo nasi tidak dilacak sampai diaktifkan dengan data nyata.')
    sf={'B':lambda r:f'IF(A{r}="","",{look(f"A{r}","SETUP","B")})',
        'E':lambda r:f'IF(OR(A{r}="",J{r}<>"OK"),"",SUMIFS({rng("BELANJA","M")},{rng("BELANJA","C")},A{r},{rng("BELANJA","B")},">="&C{r},{rng("BELANJA","B")},"<="&TglLaporan))',
        'F':lambda r:f'IF(OR(A{r}="",J{r}<>"OK"),"",SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},A{r},{rng("ITEM","I")},">="&C{r},{rng("ITEM","I")},"<="&TglLaporan))',
        'G':lambda r:f'IF(OR(A{r}="",J{r}<>"OK"),"",SUMIFS({CORR("I")},{CORR("C")},A{r},{CORR("B")},">="&C{r},{CORR("B")},"<="&TglLaporan))',
        'H':lambda r:f'IF(OR(A{r}="",J{r}<>"OK",{look(f"A{r}","SETUP","H")}<>"YA",C{r}>TglLaporan),"",N(D{r})+N(E{r})-N(F{r})+N(G{r}))',
        'I':lambda r:f'IF(A{r}="","",IF({look(f"A{r}","SETUP","H")}<>"YA","TIDAK DILACAK",IF(J{r}<>"OK",J{r},IF(C{r}>TglLaporan,"BELUM MULAI",IF(H{r}<0,"STOK NEGATIF",IF(H{r}=0,"HABIS","TERSEDIA"))))))',
        'J':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("STOK","A")},A{r})<>1','KODE GANDA'),(f'COUNTIF(KodeMenu,A{r})<>1','CEK MENU'),(f'AND({look(f"A{r}","SETUP","H")}="YA",C{r}="",D{r}="")','BELUM AWAL'),(f'AND({look(f"A{r}","SETUP","H")}="YA",OR(NOT(ISNUMBER(C{r})),C{r}<=0,C{r}>TODAY(),NOT(ISNUMBER(D{r})),D{r}<0,D{r}<>INT(D{r})))','ISI TGL/QTY AWAL')])}
    b.table(s,'tblStok',['Kode','Menu','Tanggal Awal','Pcs/Porsi Awal','Belanja Pcs','Terjual Pcs','Koreksi','Saldo','Status','Cek'],20,sf,'ACD',{'C':'in_date',**{c:'qty' for c in 'EFGH'}})
    for r,m in enumerate(p['items'],7):s.write('A'+str(r),m['code'],f['input'])
    s.merge_range('A29:L32','Koreksi di bawah hanya untuk jumlah fisik rusak/tambah/kurang, bukan pembelian kedua atau omzet. Isi alasan. Jika kas ayam dan obat masih bercampur, catat perpindahan kas pada KAS USAHA; jangan mengubah stok untuk menutup selisih uang.',f['note'])
    cf={'H':lambda r:checks(f'A{r}',[(f'COUNTIF({CORR("A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(KodeMenu,C{r})<>1','CEK MENU'),(f'COUNTIF(JenisKoreksi,D{r})<>1','CEK JENIS'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0,E{r}<>INT(E{r}))','CEK PCS'),(f'F{r}=""','ISI ALASAN'),(f'AND(G{r}<>"AKTIF",G{r}<>"VOID")','CEK STATUS')]),'I':lambda r:f'IF(AND(A{r}<>"",G{r}="AKTIF",H{r}="OK"),N(E{r})*IF(D{r}="TAMBAH",1,-1),0)',
        'J':lambda r:f'IF(OR(A{r}="",{look(f"C{r}","SETUP","H")}<>"YA",{look(f"C{r}","STOK","J")}<>"OK",B{r}<{look(f"C{r}","STOK","C","0")}),"",{stock_at(f"C{r}",f"B{r}")})'}
    b.table(s,'tblKoreksi',['ID Koreksi','Tanggal','Kode','Jenis','Pcs/Porsi','Alasan','Status','Cek','Qty Bertanda','Saldo Hari Koreksi'],200,cf,'ABCDEFG',{'B':'in_date','E':'in_qty','I':'qty','J':'qty'},first=36)
    b.validate(s,'C36:C235','=KodeMenu');b.validate(s,'D36:D235','=JenisKoreksi');b.validate(s,'G36:G235',['AKTIF','VOID']);s.print_area('A1:L32')
    s.conditional_format('H7:H26',{'type':'cell','criteria':'<','value':0,'format':f['bad']});s.conditional_format('J36:J235',{'type':'cell','criteria':'<','value':0,'format':f['bad']})

    s=b.base('NOTA','Jurnal audit: satu baris satu nota. Harga/HPP disimpan tetap. Dapat dibuka melalui Unhide jika perlu koreksi atau pencatatan manual.',16)
    nf={'J':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","K")},{rng("ITEM","B")},A{r}))','K':lambda r:f'IF(A{r}="","",J{r}-N(F{r}))','L':lambda r:f'IF(A{r}="","",IF(E{r}="Tunai",N(G{r})-K{r},0))','M':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","L")},{rng("ITEM","B")},A{r}))','N':lambda r:f'IF(A{r}="","",IF(COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","H")},"<>DIKONFIRMASI")>0,"BELUM DIKONFIRMASI","DIKONFIRMASI"))',
        'O':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeNota,A{r})<>1','NOTA GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY(),NOT(ISNUMBER(C{r})),C{r}<0,C{r}>=1)','CEK TANGGAL/JAM'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(H{r}<>"AKTIF",H{r}<>"VOID")','CEK STATUS'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<0,F{r}>J{r})','CEK DISKON'),(f'J{r}<=0','ISI ITEM'),(f'COUNTIF({rng("ITEM","B")},A{r})>7','MAKS 7 BARIS'),(f'COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","M")},"<>OK")>0','CEK ITEM'),(f'OR(NOT(ISNUMBER(G{r})),G{r}<K{r},AND(E{r}<>"Tunai",G{r}<>K{r}))','CEK PEMBAYARAN')]),'P':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",O{r}="OK"),K{r},0)'}
    b.table(s,'tblNota',['No Nota','Tanggal','Jam','Kasir','Metode','Diskon','Diterima','Status','Catatan','Bruto','Netto','Kembalian','HPP Acuan','Status HPP','Cek','Netto Efektif'],1500,nf,'ABCDEFGHI',{'B':'in_date','C':'in_time','F':'in_money','G':'in_money',**{c:'money' for c in 'JKLMP'}});b.audit(s,'O',1506);s.hide()

    s=b.base('ITEM','Jurnal audit per item. Saus merupakan tambahan per potong; nasi tanpa tambahan saus. Stok bergerak dalam pcs/porsi aktual.',15)
    itf={'I':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","B","0")})','J':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","H")})','K':lambda r:f'IF(AND(A{r}<>"",M{r}="OK"),N(E{r})*N(F{r}),0)','L':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",H{r}="DIKONFIRMASI"),N(E{r})*N(G{r}),0)',
        'M':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("ITEM","A")},A{r})<>1','ID GANDA'),(f'COUNTIF(KodeNota,B{r})<>1','CEK NOTA'),(f'COUNTIF(KodeMenu,C{r})<>1','CEK MENU'),(f'COUNTIF(RasaAyam,D{r})<>1','CEK RASA'),(f'AND(C{r}="NASI",D{r}<>"Original")','NASI TANPA SAUS'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0,E{r}<>INT(E{r}))','CEK QTY'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<=0)','CEK HARGA'),(f'AND(H{r}<>"DIKONFIRMASI",H{r}<>"BELUM DIKONFIRMASI")','CEK STATUS HPP'),(f'AND(H{r}="DIKONFIRMASI",OR(NOT(ISNUMBER(G{r})),G{r}<0))','ISI HPP SNAPSHOT')]),
        'N':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",J{r}="AKTIF",{look(f"B{r}","NOTA","O")}="OK"),N(E{r}),0)'}
    def item_balance(r):
        code=f'C{r}';opening=look(code,'STOK','C','0');qty=look(code,'STOK','D','0')
        return (f'IF(OR(A{r}="",N{r}=0,{look(code,"SETUP","H")}<>"YA",{look(code,"STOK","J")}<>"OK",I{r}<{opening}),"",N({qty})'
          f'+SUMIFS({rng("BELANJA","M")},{rng("BELANJA","C")},C{r},{rng("BELANJA","B")},">="&{opening},{rng("BELANJA","B")},"<="&I{r})'
          f'+SUMIFS({CORR("I")},{CORR("C")},C{r},{CORR("B")},">="&{opening},{CORR("B")},"<="&I{r})'
          f'-SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},C{r},{rng("ITEM","I")},">="&{opening},{rng("ITEM","I")},"<"&I{r})'
          f'-SUMIFS($N$7:N{r},$C$7:C{r},C{r},$I$7:I{r},I{r}))')
    itf['O']=item_balance
    b.table(s,'tblItem',['ID Item','No Nota','Menu','Rasa','Qty','Harga Snapshot','HPP Snapshot','Status HPP','Tanggal','Status Nota','Bruto','HPP Total','Cek','Qty Efektif','Saldo Historis'],6000,itf,'ABCDEFGH',{'E':'in_qty','F':'in_money','G':'in_money','I':'date','K':'money','L':'money','N':'qty','O':'qty'});b.audit(s,'M',6006);s.hide()

    s=b.base('KASIR','Pilih potongan + rasa, masukkan jumlah; nasi terpisah. Pembelian dan arus uang dapat dicatat dari panel kanan. Tidak ada menu obat di kasir ini.')
    for cell,text in [('A6','Tanggal'),('D6','Jam'),('G6','Metode'),('A8','Kasir')]:s.write(cell,text,f['label'])
    s.write_datetime('B6',today,f['in_date']);s.write('E6',9/24,f['in_time']);s.write('H6','Tunai',f['input']);s.write_blank('B8',None,f['input']);b.validate(s,'H6','=MetodeBayar')
    s.write_row('A12',['No','Potongan','Rasa','Qty','Harga / Pcs','Bruto'],f['head']);s.set_column('B:C',23)
    for r in range(13,19):
        s.write('A'+str(r),r-12,f['calc']);s.write_blank('B'+str(r),None,f['input']);s.write('C'+str(r),'Original',f['input']);s.write_blank('D'+str(r),None,f['in_qty'])
        b.formula(s,'E'+str(r),f'IF(B{r}="","",{look(f"B{r}","SETUP","C","0")}+{flavor(f"C{r}","B","0")})','in_money');b.formula(s,'F'+str(r),f'IF(B{r}="","",N(D{r})*N(E{r}))','money')
    b.validate(s,'B13:B18','=PotonganAyam');b.validate(s,'C13:C18','=RasaAyam')
    s.write('D21','Nasi (porsi)',f['label']);s.write_blank('E21',None,f['in_qty']);s.write('D22','Subtotal Nasi',f['label']);b.formula(s,'F22',f'N(E21)*{look(chr(34)+"NASI"+chr(34),"SETUP","C","0")}','money')
    for r,label in [(24,'Bruto'),(25,'Diskon Rp'),(26,'TOTAL BAYAR'),(27,'Diterima'),(28,'Kembalian')]:s.write('E'+str(r),label,f['label'])
    b.formula(s,'F24','SUM(F13:F18)+N(F22)','money',0);s.write('F25',0,f['in_money']);b.formula(s,'F26','F24-N(F25)','money',0);s.write_blank('F27',None,f['in_money']);b.formula(s,'F28','IF(H6="Tunai",N(F27)-F26,0)','money',0)
    s.write('H12','BELANJA AYAM',f['head']);s.write('H24','KAS USAHA',f['head'])
    fields=[('H13','Tanggal','I13','date'),('H14','Potongan / Nasi','I14','text'),('H15','Kg (opsional)','I15','qty'),('H16','Pcs/Porsi AKTUAL','I16','qty'),('H17','Nominal total','I17','money'),('H18','Metode','I18','text'),('H25','Tanggal kas','I25','date'),('H26','Jenis kas','I26','text'),('H27','Nominal','I27','money'),('H28','Metode','I28','text'),('H29','Catatan','I29','text')]
    for label,text,cell,kind in fields:s.write(label,text,f['label']);s.write_blank(cell,None,f['input' if kind=='text' else 'in_'+kind])
    for cell in ['I13','I25']:s.write_datetime(cell,today,f['in_date'])
    s.write('I18','Tunai',f['input']);s.write('I26','BIAYA USAHA',f['input']);s.write('I28','Tunai',f['input'])
    for cell,src in [('I14','=KodeMenu'),('I18','=MetodeBayar'),('I26','=JenisKas'),('I28','=MetodeBayar')]:b.validate(s,cell,src)
    s.write('H19','Perkiraan pcs',f['label']);b.formula(s,'I19',f'IF(OR(I14="",I15=""),"",N(I15)*{look("I14","SETUP","E","0")})','qty')
    s.write('H20','Acuan nominal',f['label']);b.formula(s,'I20','IF(OR(I14="",I14="NASI",I15=""),"",N(I15)*SETUP!B40)','money')
    s.merge_range('A32:F35','Saus ditambahkan per potong, bukan sekali per nota. Nasi tanpa pilihan saus. Harga/HPP lama disimpan sebagai snapshot. Ketika stok belum diisi, kasir mencatat uang tetapi belum dapat memverifikasi kecukupan stok.',f['warn']);s.set_column('H:H',25);s.set_column('I:I',26);s.print_area('A1:L35')

    s=b.base('HARIAN','Laporan satu tanggal: kas ayam, pembelian, stok dan potongan terjual. QRIS/transfer tidak dihitung sebagai uang fisik. Gunakan kas usaha ayam saja.')
    for cell,label in [('A5','Tanggal'),('D5','Kas awal ayam'),('H5','Uang fisik ayam')]:s.write(cell,label,f['label'])
    s.write_datetime('B5',today,f['in_date']);s.write_blank('E5',None,f['in_money']);s.write_blank('I5',None,f['in_money'])
    valid=f'{rng("NOTA","B")},$B$5,{rng("NOTA","H")},"AKTIF",{rng("NOTA","O")},"OK"'
    total=f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5)'
    expenses=f'-SUMIFS({rng("KAS USAHA","I")},{rng("KAS USAHA","B")},$B$5,{rng("KAS USAHA","C")},"BIAYA USAHA")'
    incomplete=f'COUNTIFS({valid},{rng("NOTA","N")},"<>DIKONFIRMASI")'
    cards=[('PENJUALAN NETTO',total),('NOTA VALID',f'COUNTIFS({valid})'),('PCS AYAM TERJUAL',f'SUMIFS({rng("ITEM","N")},{rng("ITEM","I")},$B$5,{rng("ITEM","C")},"<>NASI")'),('BELANJA AYAM / NASI',f'SUMIFS({rng("BELANJA","N")},{rng("BELANJA","B")},$B$5)'),('BIAYA USAHA LAIN',expenses),('SELISIH NETTO - HPP ACUAN',f'IF({incomplete}>0,"HPP BELUM DIKONFIRMASI",{total}-SUMIFS({rng("NOTA","M")},{valid}))')]
    status_card=w.add_format({'font_name':'Calibri','font_size':16,'bold':True,'bg_color':'#302D29','font_color':'white','text_wrap':True,'valign':'vcenter','num_format':'#,##0;[Red](#,##0);"0"'})
    for i,(label,formula) in enumerate(cards):
        rr=6+i//3*5;cc=i%3*4;style=status_card if i==5 else f['kpi']
        s.merge_range(rr,cc,rr,cc+2,label,f['kpi_label']);s.merge_range(rr+1,cc,rr+2,cc+2,'',style);s.write_formula(rr+1,cc,'='+formula,style,0)
    for r,label,formula in [(18,'Penjualan Tunai Netto',f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5,{rng("NOTA","E")},"Tunai")'),(19,'Belanja Tunai',f'SUMIFS({rng("BELANJA","N")},{rng("BELANJA","B")},$B$5,{rng("BELANJA","G")},"Tunai")'),(20,'Kas Lain Tunai Bertanda',f'SUMIFS({rng("KAS USAHA","I")},{rng("KAS USAHA","B")},$B$5,{rng("KAS USAHA","E")},"Tunai")'),(22,'Kas Ayam Seharusnya','IF(ISNUMBER(E5),E5+B18-B19+B20,"")'),(23,'Selisih Fisik','IF(AND(ISNUMBER(I5),ISNUMBER(B22)),I5-B22,"")')]:s.write('A'+str(r),label,f['label']);b.formula(s,'B'+str(r),formula,'money')
    s.write_row('E18',['Metode','Netto'],f['head'])
    for r,method in enumerate(METHODS,19):s.write('E'+str(r),method,f['calc']);b.formula(s,'F'+str(r),f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5,{rng("NOTA","E")},E{r})','money')
    s.write_row('A26',['Potongan / Menu','Qty','Bruto Sebelum Diskon','Saldo Pcs','Status'],f['head'])
    for r,m in enumerate(p['items'],27):
        s.write('A'+str(r),m['name'],f['calc']);key='"'+m['code']+'"'
        b.formula(s,'B'+str(r),f'SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},{key},{rng("ITEM","I")},$B$5)','qty')
        b.formula(s,'C'+str(r),f'SUMIFS({rng("ITEM","K")},{rng("ITEM","C")},{key},{rng("ITEM","I")},$B$5,{rng("ITEM","N")},">0")','money')
        for c,source in [('D','H'),('E','I')]:b.formula(s,c+str(r),look(key,'STOK',source),'qty' if c=='D' else 'calc')
    chart=w.add_chart({'type':'column'});chart.add_series({'name':'Pcs ayam terjual','categories':"='HARIAN'!$A$27:$A$30",'values':"='HARIAN'!$B$27:$B$30",'fill':{'color':'#17634F'},'border':{'none':True}});chart.set_title({'name':'Potongan terjual hari ini'});chart.set_legend({'none':True});s.insert_chart('G26',chart,{'x_scale':1.05,'y_scale':1.0})
    s.merge_range('A34:F34','PERIKSA SEBELUM MENUTUP HARI',f['head'])
    alerts=[('Nota bermasalah',f'COUNTIFS({rng("NOTA","A")},"<>",{rng("NOTA","O")},"<>OK")'),('Stok belum awal',f'COUNTIF({rng("STOK","J")},"BELUM AWAL")'),('Saldo historis negatif',f'COUNTIF({rng("ITEM","O")},"<0")+COUNTIF({CORR("J")},"<0")'),('Belanja bermasalah',f'COUNTIFS({rng("BELANJA","A")},"<>",{rng("BELANJA","L")},"<>OK")')]
    for r,(label,formula) in enumerate(alerts,35):s.write('A'+str(r),label,f['label']);b.formula(s,'D'+str(r),formula,'qty');s.conditional_format('D'+str(r),{'type':'cell','criteria':'>','value':0,'format':f['warn']})
    s.merge_range('A41:L44','Selisih netto − HPP acuan bukan laba bersih dan belum mengurangi biaya usaha lain. Pembelian stok tidak dikurangkan lagi dari selisih HPP tersebut. Uang pribadi/usaha obat bukan omzet ayam. Saldo stok dihitung pada tingkat tanggal: pembelian/koreksi satu tanggal dianggap tersedia pada hari itu, bukan urutan jam.',f['note']);s.set_column('A:A',30);s.set_column('C:C',29);s.set_column('E:E',27);s.print_area('A1:M44')

    for c,name,vals in [('N','MetodeBayar',METHODS),('O','JenisKas',CASHKINDS),('P','JenisKoreksi',['RUSAK','TAMBAH','KURANG'])]:
        b.s['SETUP'].write_column(c+'2',vals);w.define_name(name,f"='SETUP'!${c}$2:${c}${len(vals)+1}")
    b.s['SETUP'].set_column('N:P',None,None,{'hidden':True})
    b.close(['MULAI','KASIR','HARIAN','BELANJA','STOK','KAS USAHA','SETUP','NOTA','ITEM']);return path

if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('--output',type=Path,default=ROOT/'output/POS_Chicken_Al_Haqi.xlsx');print(build(a.parse_args().output))
