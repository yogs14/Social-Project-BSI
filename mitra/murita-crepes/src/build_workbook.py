"""Murita POS: explicit stock transformations, per-receipt CRM and frozen journals."""
import argparse,datetime as dt,json
from pathlib import Path
from layout import Layout,rng,look,checks,CAP,col

ROOT=Path(__file__).resolve().parents[1]

def build(path):
    p=json.loads((ROOT/'data/profile.json').read_text());path.parent.mkdir(parents=True,exist_ok=True)
    b=Layout(path);f=b.f;w=b.w
    material_row={m['code']:i+7 for i,m in enumerate(p['materials'])}
    def cost(code):return f"BAHAN!$G${material_row[code]}"
    def named(name,sheet,c):w.define_name(name,'='+rng(sheet,c))
    for name,sheet,c in [('KodeBahan','BAHAN','A'),('KodeMenu','MENU','A'),('KodeTop','KALIBRASI','A'),('KodeNota','NOTA','A'),('KodePelanggan','PELANGGAN','A')]:named(name,sheet,c)
    today=dt.datetime.now().replace(hour=0,minute=0,second=0,microsecond=0)
    movements=['AWAL','BELI','PRODUKSI MASUK','PRODUKSI PAKAI','JUAL PAKAI','RUSAK','KOREKSI MASUK','KOREKSI KELUAR']
    methods=['Tunai','QRIS','Transfer','E-wallet']

    s=b.base('MULAI','POS, premiks tujuh bahan, adonan basah, stok topping dan pelanggan • Pak Syarif • MoM 24 Oktober 2026.')
    s.merge_range('A6:L6','01  SIAPKAN BAHAN, PRODUKSIKAN ADONAN, LALU JUAL',f['head'])
    sections=[('1. Periksa data sumber','34 menu dan harga dari file biaya mitra sudah dimuat. Periksa harga terbaru; harga SKM belum diketahui dan tidak diambil dari rumus sumber yang keliru.'),('2. Isi stok awal','Isi AWAL per kode bahan di GERAK BAHAN berdasarkan hitung fisik. Satuan dasar gram, buah/pcs, lembar atau menit; bukan jumlah kemasan.'),('3. Buat premiks','PRODUKSI > PREMIKS mengubah tujuh bahan mentah menjadi stok premiks. PRODUKSI > ADONAN BASAH menggunakan premiks, mentega dan air. Jangan memotong tujuh bahan mentah lagi saat penjualan.'),('4. Jalankan kasir','Isi KASIR: menu, qty, pelanggan opsional, metode bayar, diskon dan uang diterima. Dengan VBA, nota, item, pemakaian bahan dan poin disimpan bersama sebagai nilai tetap.'),('5. Pisahkan uang usaha','BELI mencatat pembelian bahan. KEUANGAN memisahkan BIAYA USAHA, PRIVE, MODAL MASUK dan PENDAPATAN LAIN. Modal/prive bukan laba/rugi operasional.'),('6. Tutup dan evaluasi','TUTUP KASIR mencocokkan uang tunai. DASHBOARD merangkum penjualan, menu laris, stok dan biaya. CRM menghitung kunjungan per nota; program poin awalnya NONAKTIF.')]
    for i,(title,body) in enumerate(sections):
        r=7+i*3;s.merge_range(r,0,r,2,title,f['label']);s.merge_range(r,3,r+1,11,body,f['body'])
    s.merge_range('A28:L30','Sel biru = input, abu-abu = rumus. Workbook utama tidak berisi penjualan, pelanggan, pembelian atau stok fiktif. HPP yang belum lengkap ditandai; dashboard tidak menganggap biaya yang hilang sebagai laba.',f['warn'])
    s.merge_range('A32:L34','TOMBOL OPSIONAL: simpan salinan sebagai .xlsm, Alt+F11 > File > Import > Modul_MuritaCrepes.bas, Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol. Tanpa macro, jurnal dapat diisi langsung sesuai panduan; FORM/KASIR tidak dapat menyimpan otomatis.',f['note'])
    s.merge_range('A36:L39','Selesaikan sebelum operasional: harga beli SKM; stok fisik; kalibrasi sendok topping; konfirmasi takaran SKM 5,63 g, gas 4 menit dan kemasan 0,25 lembar/porsi; serta aturan poin bila ingin diaktifkan. Jam sumber 17.00–23.00, Senin libur rutin. File 30 hari/tally sumber disimpan sebagai referensi, tidak diimpor otomatis sebagai nota.',f['warn'])
    s.merge_range('A41:L43','Kapasitas: 50 menu, 60 bahan, 25 kalibrasi topping, 1.000 nota, 3.000 item, 6.000 gerak bahan, 300 pelanggan, 1.000 gerak poin dan 1.000 catatan keuangan. Satu operator per file. Cadangkan dan arsipkan sebelum tabel penuh; jangan menambah data di luar batas tabel.',f['note']);s.print_area('A1:L43')

    s=b.base('BAHAN','Harga acuan per kemasan diubah menjadi biaya per satuan dasar. Stok dihitung dari GERAK BAHAN sampai tanggal akhir dashboard. HPP hanya estimasi resep.',14)
    def material_check(r):
        ordinary=f'IF(OR(NOT(ISNUMBER(E{r})),E{r}<0,NOT(ISNUMBER(F{r})),F{r}<=0),"ISI HARGA/ISI BERSIH","OK")'
        branch=f'IF(A{r}="PREMIX",IF(PREMIKS!$B$21="OK","OK","CEK PREMIKS"),IF(A{r}="WET",IF(ADONAN!$B$21="OK","OK","CEK ADONAN"),{ordinary}))'
        return checks(f'A{r}',[(f'COUNTIF(KodeBahan,A{r})<>1','KODE GANDA'),(f'OR(B{r}="",D{r}="")','ISI NAMA/SATUAN'),(f'AND(H{r}<>"YA",H{r}<>"TIDAK")','CEK LACAK STOK')],branch)
    bf={'G':lambda r:f'IF(A{r}="","",IF(A{r}="PREMIX",PREMIKS!$B$20,IF(A{r}="WET",ADONAN!$B$18,IFERROR(N(E{r})/N(F{r}),0))))','K':material_check,
        'L':lambda r:f'IF(A{r}="","",IF(H{r}<>"YA","",ROUND(SUMIFS({rng("GERAK BAHAN","M")},{rng("GERAK BAHAN","E")},A{r},{rng("GERAK BAHAN","B")},"<="&TglAkhir),6)))',
        'M':lambda r:f'IF(A{r}="","",IF(H{r}<>"YA","TIDAK DILACAK",IF(L{r}<0,"STOK NEGATIF",IF(L{r}=0,"HABIS / BELUM AWAL",IF(AND(ISNUMBER(I{r}),L{r}<=I{r}),"BELANJA","TERSEDIA")))))'}
    b.table(s,'tblBahan',['Kode','Nama Bahan','Kelompok','Satuan Dasar','Harga Kemasan','Isi Kemasan','Biaya / Satuan','Lacak Stok','Batas Belanja','Catatan','Cek','Stok Akhir','Status Stok'],60,bf,'ABCDEFHIJ',{'E':'in_money','F':'in_qty','G':'qty','I':'in_qty','L':'qty'})
    for i,m in enumerate(p['materials'],7):
        for c,v in [(0,m['code']),(1,m['name']),(2,m['group']),(3,m['unit']),(4,m['price']),(5,m['pack']),(7,m['track'])]:
            if v is not None:s.write(i-1,c,v,f['in_money'] if c==4 else f['input'])
        note='Harga SKM sumber salah referensi; isi harga beli sebenarnya.' if m['code']=='SKM' else ('Biaya dihitung dari produksi; bukan harga beli.' if m['group']=='ANTARA' else 'Acuan file biaya mitra; konfirmasi harga/takaran aktual.')
        s.write(i-1,9,note,f['input'])
    b.validate(s,'H7:H66',['YA','TIDAK']);s.set_column('B:B',35);s.set_column('J:J',48);b.audit(s,'K',66)
    s.conditional_format('L7:L66',{'type':'cell','criteria':'<','value':0,'format':f['bad']})

    s=b.base('PREMIKS','Tujuh komponen dinormalisasi berdasarkan total berat sebenarnya, bukan angka pembulatan 2.900 g. Kalkulator ini tidak menambah stok tanpa jurnal produksi.')
    s.write('D5','Target premiks (g)',f['label']);s.write('E5',1000,f['in_qty'])
    s.write_row('A6',['Kode','Bahan','Bobot Resep (g)','Biaya / g','Biaya Resep','Proporsi','Untuk Target (g)'],f['head'])
    for r,m in enumerate(p['materials'][:7],7):
        s.write('A'+str(r),m['code'],f['input']);b.formula(s,'B'+str(r),look('A'+str(r),'BAHAN','B'))
        s.write('C'+str(r),m['recipe_weight'],f['in_qty']);b.formula(s,'D'+str(r),look('A'+str(r),'BAHAN','G','0'),'qty')
        b.formula(s,'E'+str(r),f'N(C{r})*N(D{r})','money');b.formula(s,'F'+str(r),f'IFERROR(C{r}/$B$18,0)','pct');b.formula(s,'G'+str(r),f'N($E$5)*F{r}','qty')
    for cell,text in [('A18','Total berat (g)'),('A19','Biaya resep (Rp)'),('A20','Biaya / gram'),('A21','Validasi resep')]:s.write(cell,text,f['label'])
    b.formula(s,'B18','SUM(C7:C13)','qty',2908.5);b.formula(s,'B19','SUM(E7:E13)','money',56731.25);b.formula(s,'B20','IFERROR(B19/B18,0)','qty')
    rawerrors='+'.join(f'IF({look(f"A{r}","BAHAN","K")}<>"OK",1,0)' for r in range(7,14))
    b.formula(s,'B21',f'IF(OR(COUNT(C7:C13)<>7,MIN(C7:C13)<=0,({rawerrors})>0),"CEK DATA","OK")')
    s.merge_range('A24:L27','Sumber: 1.000 g tepung beras + 500 g terigu + 500 g maizena + 400 g fullcream + 500 g gula + 6,5 g garam + 2 g vanili = 2.908,5 g. Bahan mentah berkurang hanya saat membuat premiks. Biaya acuan campuran asli Rp56.731,25, sebelum mentega/air/pelengkap.',f['note']);s.set_column('B:B',30);s.print_area('A1:L27')

    s=b.base('ADONAN','Air dan mentega mengikuti jumlah premiks yang dipakai. Hasil 30 crepes/kg adalah asumsi kerja yang disepakati di MoM; dapat diubah setelah pengukuran.')
    labels={5:'Premiks target (g)',7:'Crepes per kg premiks',8:'Mentega / premiks',9:'Air / premiks',11:'Mentega (g)',12:'Air (g)',13:'Adonan basah (g)',14:'Perkiraan porsi',15:'Basah / porsi (g)',17:'Biaya adonan batch',18:'Biaya basah / g',19:'Biaya adonan / porsi',21:'Validasi'}
    for r,label in labels.items():s.write('A'+str(r),label,f['label'])
    for r,v in [(5,1000),(7,p['yield_per_kg']),(8,p['butter_ratio']),(9,p['water_ratio'])]:s.write('B'+str(r),v,f['in_qty'])
    for cell,formula in {'B11':'N(B5)*N(B8)','B12':'N(B5)*N(B9)','B13':'N(B5)+B11+B12','B14':'N(B5)/1000*N(B7)','B15':'IFERROR(B13/B14,0)','B17':f'N(B5)*PREMIKS!$B$20+B11*{cost("BTR")}+B12*{cost("WTR")}','B18':'IFERROR(B17/B13,0)','B19':'IFERROR(B17/B14,0)'}.items():b.formula(s,cell,formula,'qty' if cell not in ['B17','B19'] else 'money')
    b.formula(s,'B21',f'IF(OR(PREMIKS!$B$21<>"OK",COUNT(B5,B7:B9)<>4,B5<=0,B7<=0,B8<0,B9<0,BAHAN!K{material_row["BTR"]}<>"OK",BAHAN!K{material_row["WTR"]}<>"OK"),"CEK DATA","OK")')
    s.merge_range('D6:L10','Contoh kalkulator: 1.000 g premiks + 65 g mentega + 830 g air = 1.895 g adonan basah, diperkirakan 30 porsi. 1 porsi menggunakan 63,1667 g adonan basah. Ini rencana, bukan stok awal atau transaksi produksi.',f['warn'])
    s.merge_range('D13:L18','Penjualan mengurangi stok WET, SKM, gas/kemasan dan topping sesuai resep. Penjualan tidak lagi mengurangi premiks atau tujuh bahan mentah; pemotongan tersebut sudah terjadi pada produksi sebelumnya.',f['body']);s.set_column('A:A',30);s.print_area('A1:L23')

    s=b.base('KALIBRASI','Timbang tiga sendok/takaran dengan teknik yang sama. Isi satuan dasar per sendok dan sendok per porsi. Sebelum kalibrasi, gunakan standar sumber sebagai ESTIMASI.',12)
    kf={'B':lambda r:f'IF(A{r}="","",{look(f"A{r}","BAHAN","B")})','C':lambda r:f'IF(A{r}="","",{look(f"A{r}","BAHAN","D")})',
        'I':lambda r:f'IF(A{r}="","",IF(AND(COUNT(E{r}:G{r})=3,MIN(E{r}:G{r})>0),AVERAGE(E{r}:G{r}),""))',
        'J':lambda r:f'IF(A{r}="","",IF(AND(ISNUMBER(I{r}),ISNUMBER(H{r}),H{r}>0),I{r}*H{r},N(D{r})))',
        'K':lambda r:f'IF(A{r}="","",IF(AND(ISNUMBER(I{r}),ISNUMBER(H{r}),H{r}>0),"TERKALIBRASI","ACUAN SUMBER"))',
        'L':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeBahan,A{r})<>1','CEK KODE'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0)','ISI ACUAN'),(f'AND(COUNTA(E{r}:H{r})>0,OR(COUNT(E{r}:G{r})<>3,MIN(E{r}:G{r})<=0,NOT(ISNUMBER(H{r})),H{r}<=0))','LENGKAPI KALIBRASI')])}
    b.table(s,'tblKalibrasi',['Kode Topping','Nama','Satuan Dasar','Acuan / Porsi','Sampel 1 / Sendok','Sampel 2 / Sendok','Sampel 3 / Sendok','Sendok / Porsi','Rerata / Sendok','Pakai / Porsi','Dasar Takaran','Cek'],25,kf,'ADEFGH',{c:'qty' for c in 'IJ'}|{c:'in_qty' for c in 'DEFGH'})
    for r,m in enumerate([m for m in p['materials'] if m['group']=='TOPPING'],7):s.write('A'+str(r),m['code'],f['input']);s.write('D'+str(r),m['portion'],f['in_qty'])
    s.set_column('B:B',30);b.audit(s,'L',31)

    s=b.base('MENU','34 menu sumber; kapasitas 50. Harga dan HPP harus disalin sebagai nilai tetap saat transaksi. HPP tetap diberi label acuan, bukan biaya aktual FIFO.',12)
    mf={'H':lambda r:f'IF(A{r}="","",RESEP!BK{r})','I':lambda r:f'IF(A{r}="","",RESEP!BL{r})','K':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeMenu,A{r})<>1','SKU GANDA'),(f'B{r}=""','ISI MENU'),(f'OR(NOT(ISNUMBER(C{r})),C{r}<=0)','ISI HARGA'),(f'AND(G{r}<>"AKTIF",G{r}<>"NONAKTIF")','CEK STATUS'),(f'SUMPRODUCT(--(D{r}:F{r}<>""),--(COUNTIF(KodeTop,D{r}:F{r})<>1))>0','CEK TOPPING')])}
    b.table(s,'tblMenu',['SKU','Nama Menu','Harga Jual','Topping 1','Topping 2','Topping 3','Status','HPP Acuan','Kelengkapan HPP','Catatan','Cek'],50,mf,'ABCDEFGJ',{'C':'in_money','H':'money'})
    for r,m in enumerate(p['menus'],7):
        s.write_row(r-1,0,[m['sku'],m['name'],m['price'],*(m['toppings']+['']*(3-len(m['toppings']))),'AKTIF'],f['input'])
        s.write('J'+str(r),'Harga/topping dari file biaya; verifikasi harga terkini.',f['input'])
    for c in 'DEF':b.validate(s,c+'7:'+c+'56','=KodeTop')
    b.validate(s,'G7:G56',['AKTIF','NONAKTIF']);s.set_column('B:B',39);s.set_column('I:J',35);b.audit(s,'K',56)

    s=b.base('RESEP','Matriks pemakaian satu crepes: adonan basah + pelengkap + topping. Biaya dan kuantitas saat transaksi disalin sebagai snapshot melalui VBA.',64)
    s.set_row(3,None,None,{'hidden':True});s.write('A5','Biaya/satuan',f['label'])
    headers=['SKU','Menu']
    for i in range(60):
        c=col(i+2);mr=i+7;headers.append(p['materials'][i]['code'] if i<len(p['materials']) else 'Cadangan '+str(i+1))
        b.formula(s,c+'4',f'IF(BAHAN!K{mr}="OK",0,1)');b.formula(s,c+'5',f'N(BAHAN!G{mr})','qty')
    headers+=['HPP Acuan','Kelengkapan']
    rf={'A':lambda r:f'IF(MENU!A{r}="","",MENU!A{r})','B':lambda r:f'IF(A{r}="","",MENU!B{r})',
        'BK':lambda r:f'IF(A{r}="","",SUMPRODUCT(C{r}:BJ{r},$C$5:$BJ$5))',
        'BL':lambda r:f'IF(A{r}="","",IF(SUMPRODUCT(--(C{r}:BJ{r}>0),$C$4:$BJ$4)>0,"BELUM LENGKAP","LENGKAP"))'}
    for i in range(60):
        c=col(i+2);mr=i+7
        rf[c]=lambda r,mr=mr:f'IF(A{r}="","",IF(BAHAN!$A${mr}="WET",ADONAN!$B$15,IF(BAHAN!$A${mr}="SKM",REFERENSI!$B$10,IF(BAHAN!$A${mr}="GAS",REFERENSI!$B$11,IF(BAHAN!$A${mr}="PACK",REFERENSI!$B$12,COUNTIF(MENU!D{r}:F{r},BAHAN!$A${mr})*{look(f"BAHAN!$A${mr}","KALIBRASI","J","0")})))))'
    b.table(s,'tblResep',headers,50,rf);s.set_column('B:B',36);s.set_column('C:BJ',13)

    # Sales headers and lines deliberately avoid circular validation dependencies.
    s=b.base('NOTA','Satu baris = satu nota. Diskon dan pembayaran dicatat sekali per nota; jumlah kunjungan pelanggan juga dihitung per nota, bukan per item.',18)
    nf={
      'J':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","K")},{rng("ITEM","B")},A{r}))',
      'K':lambda r:f'IF(A{r}="","",J{r}-N(F{r}))',
      'L':lambda r:f'IF(A{r}="","",IF(E{r}="Tunai",N(G{r})-K{r},0))',
      'M':lambda r:f'IF(A{r}="","",SUMIFS({rng("ITEM","L")},{rng("ITEM","B")},A{r}))',
      'O':lambda r:checks(f'A{r}',[(f'COUNTIF(KodeNota,A{r})<>1','NOTA GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY(),NOT(ISNUMBER(C{r})),C{r}<0,C{r}>=1)','CEK TANGGAL/JAM'),(f'AND(D{r}<>"",COUNTIF(KodePelanggan,D{r})<>1)','CEK PELANGGAN'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(H{r}<>"AKTIF",H{r}<>"VOID")','CEK STATUS'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<0,F{r}>J{r})','CEK DISKON'),(f'J{r}<=0','ISI ITEM'),(f'COUNTIF({rng("ITEM","B")},A{r})>12','MAKS 12 ITEM'),(f'COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","M")},"<>OK")>0','CEK ITEM'),(f'COUNTIFS({rng("GERAK BAHAN","I")},A{r},{rng("GERAK BAHAN","D")},"JUAL PAKAI",{rng("GERAK BAHAN","J")},"AKTIF")=0','ISI PEMAKAIAN'),(f'OR(NOT(ISNUMBER(G{r})),G{r}<K{r},AND(E{r}<>"Tunai",G{r}<>K{r}))','CEK PEMBAYARAN'),(f'OR(NOT(ISNUMBER(N{r})),N{r}<0,N{r}<>INT(N{r}),AND(D{r}="",N{r}<>0))','CEK POIN')]),
      'P':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",O{r}="OK"),K{r},0)',
      'Q':lambda r:f'IF(A{r}="","",IF(COUNTIFS({rng("ITEM","B")},A{r},{rng("ITEM","G")},"<>LENGKAP")>0,"BELUM LENGKAP","LENGKAP"))',
      'R':lambda r:f'IF(AND(A{r}<>"",H{r}="AKTIF",O{r}="OK"),N(N{r}),0)',
    }
    b.table(s,'tblNota',['No Nota','Tanggal','Jam','ID Pelanggan','Metode','Diskon Rp','Diterima','Status','Catatan','Bruto','Netto','Kembalian','HPP Acuan Total','Poin Snapshot','Cek','Netto Efektif','Status HPP','Poin Efektif'],1000,nf,'ABCDEFGHI N'.replace(' ',''),{'B':'in_date','C':'in_time','F':'in_money','G':'in_money','J':'money','K':'money','L':'money','M':'money','P':'money'})
    b.validate(s,'D7:D1006','=KodePelanggan');b.validate(s,'E7:E1006','=MetodeBayar');b.validate(s,'H7:H1006',['AKTIF','VOID']);b.audit(s,'O',1006)

    s=b.base('ITEM','Satu baris = satu menu di dalam nota. Harga dan HPP merupakan snapshot nilai tetap, bukan tautan ke harga saat ini. Catat pemakaian bahan di GERAK BAHAN.',14)
    itf={
      'I':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","B","0")})',
      'J':lambda r:f'IF(A{r}="","",{look(f"B{r}","NOTA","H")})',
      'K':lambda r:f'IF(AND(A{r}<>"",M{r}="OK"),N(D{r})*N(E{r}),0)',
      'L':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",G{r}="LENGKAP"),N(D{r})*N(F{r}),0)',
      'M':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("ITEM","A")},A{r})<>1','ID GANDA'),(f'COUNTIF(KodeNota,B{r})<>1','CEK NOTA'),(f'COUNTIF(KodeMenu,C{r})<>1','CEK MENU'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0,D{r}<>INT(D{r}))','CEK QTY'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0)','CEK HARGA'),(f'AND(G{r}<>"LENGKAP",G{r}<>"BELUM LENGKAP")','CEK STATUS HPP'),(f'AND(G{r}="LENGKAP",OR(NOT(ISNUMBER(F{r})),F{r}<0))','ISI HPP SNAPSHOT')]),
      'N':lambda r:f'IF(AND(A{r}<>"",M{r}="OK",J{r}="AKTIF",{look(f"B{r}","NOTA","O")}="OK"),N(D{r}),0)'}
    b.table(s,'tblItem',['ID Item','No Nota','SKU','Qty','Harga Snapshot','HPP Snapshot','Status HPP','Catatan','Tanggal','Status Nota','Bruto','HPP Total','Cek','Qty Efektif'],3000,itf,'ABCDEFGH',{'D':'in_qty','E':'in_money','F':'in_money','I':'date','K':'money','L':'money','N':'qty'})
    b.validate(s,'B7:B3006','=KodeNota');b.validate(s,'C7:C3006','=KodeMenu');b.validate(s,'G7:G3006',['LENGKAP','BELUM LENGKAP']);b.audit(s,'M',3006)

    s=b.base('GERAK BAHAN','Semua kuantitas memakai satuan dasar bahan. BELI adalah pengeluaran pembelian; jangan dicatat lagi sebagai BIAYA USAHA. Pemakaian JUAL terkait nota mengikuti status VOID nota.',16)
    gf={
      'L':lambda r:f'IF(A{r}="","",{look(f"E{r}","BAHAN","D")})',
      'M':lambda r:f'IF(AND(A{r}<>"",N{r}="OK",J{r}="AKTIF",OR(I{r}="",AND({look(f"I{r}","NOTA","H")}="AKTIF",{look(f"I{r}","NOTA","O")}="OK"))),N(F{r})*IF(OR(D{r}="AWAL",D{r}="BELI",D{r}="PRODUKSI MASUK",D{r}="KOREKSI MASUK"),1,-1),0)',
      'N':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("GERAK BAHAN","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(JenisGerak,D{r})<>1','CEK JENIS'),(f'COUNTIF(KodeBahan,E{r})<>1','CEK BAHAN'),(f'OR(NOT(ISNUMBER(F{r})),F{r}<=0)','CEK QTY'),(f'AND(J{r}<>"AKTIF",J{r}<>"VOID")','CEK STATUS'),(f'AND(D{r}="BELI",OR(NOT(ISNUMBER(G{r})),G{r}<0,COUNTIF(MetodeBayar,H{r})<>1))','CEK PEMBELIAN'),(f'AND(D{r}="JUAL PAKAI",OR(COUNTIF(KodeNota,I{r})<>1,{look(f"I{r}","NOTA","B","0")}<>B{r}))','CEK NOTA JUAL'),(f'AND(D{r}<>"JUAL PAKAI",I{r}<>"")','NOTA HANYA JUAL'),(f'AND(OR(D{r}="PRODUKSI MASUK",D{r}="PRODUKSI PAKAI"),C{r}="")','ISI ID PRODUKSI'),(f'AND(OR(D{r}="RUSAK",D{r}="KOREKSI MASUK",D{r}="KOREKSI KELUAR"),K{r}="")','ISI ALASAN')]),
      'O':lambda r:f'IF(OR(A{r}="",{look(f"E{r}","BAHAN","H")}<>"YA",M{r}=0),"",ROUND(SUMIFS({rng("GERAK BAHAN","M")},{rng("GERAK BAHAN","E")},E{r},{rng("GERAK BAHAN","B")},"<"&B{r})+SUMIFS($M$7:M{r},$E$7:E{r},E{r},$B$7:B{r},B{r}),6))',
      'P':lambda r:f'IF(AND(A{r}<>"",D{r}="BELI",J{r}="AKTIF",N{r}="OK"),N(G{r}),0)'}
    b.table(s,'tblGerak',['ID Gerak','Tanggal','ID Dokumen','Jenis','Kode Bahan','Qty','Nominal Beli','Metode','No Nota','Status','Catatan','Satuan','Qty Bertanda','Cek','Saldo Historis','Beli Efektif'],6000,gf,'ABCDEFGHIJK',{'B':'in_date','F':'in_qty','G':'in_money','M':'qty','O':'qty','P':'money'})
    for c,source in [('D','=JenisGerak'),('E','=KodeBahan'),('H','=MetodeBayar'),('I','=KodeNota'),('J',['AKTIF','VOID'])]:b.validate(s,c+'7:'+c+'6006',source)
    b.audit(s,'N',6006);s.conditional_format('O7:O6006',{'type':'cell','criteria':'<','value':-.000001,'format':f['bad']});s.set_column('K:K',40)

    s=b.base('KEUANGAN','Pisahkan uang usaha dan pribadi. BIAYA USAHA mengurangi hasil operasional; PRIVE dan MODAL bukan biaya/pendapatan usaha. Pembelian stok hanya dicatat di GERAK BAHAN.')
    kf={'H':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("KEUANGAN","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(JenisKas,C{r})<>1','CEK JENIS'),(f'OR(NOT(ISNUMBER(D{r})),D{r}<=0)','CEK NOMINAL'),(f'COUNTIF(MetodeBayar,E{r})<>1','CEK METODE'),(f'AND(G{r}<>"AKTIF",G{r}<>"VOID")','CEK STATUS')]),'I':lambda r:f'IF(AND(A{r}<>"",H{r}="OK",G{r}="AKTIF"),N(D{r})*IF(OR(C{r}="MODAL MASUK",C{r}="PENDAPATAN LAIN"),1,-1),0)'}
    b.table(s,'tblKeuangan',['ID Kas','Tanggal','Jenis','Nominal','Metode','Catatan','Status','Cek','Kas Bertanda'],1000,kf,'ABCDEFG',{'B':'in_date','D':'in_money','I':'money'})
    b.validate(s,'C7:C1006','=JenisKas');b.validate(s,'E7:E1006','=MetodeBayar');b.validate(s,'G7:G1006',['AKTIF','VOID']);b.audit(s,'H',1006);s.set_column('F:F',44)

    s=b.base('PELANGGAN','CRM sederhana: jumlah kunjungan per nota, belanja, kunjungan terakhir dan poin. Data pelanggan opsional; tanpa pelanggan terdaftar transaksi tetap bisa dicatat.',13)
    cf={'G':lambda r:f'IF(A{r}="","",COUNTIFS({rng("NOTA","D")},A{r},{rng("NOTA","H")},"AKTIF",{rng("NOTA","O")},"OK",{rng("NOTA","B")},"<="&TglAkhir))',
      'H':lambda r:f'IF(A{r}="","",SUMIFS({rng("NOTA","P")},{rng("NOTA","D")},A{r},{rng("NOTA","B")},"<="&TglAkhir))',
      'I':lambda r:f'IF(OR(A{r}="",G{r}=0),"",IFERROR(_xlfn.AGGREGATE(14,6,{rng("NOTA","B")}/({rng("NOTA","D")}=A{r})/({rng("NOTA","H")}="AKTIF")/({rng("NOTA","O")}="OK")/({rng("NOTA","B")}<=TglAkhir),1),""))',
      'J':lambda r:f'IF(A{r}="","",SUMIFS({rng("NOTA","R")},{rng("NOTA","D")},A{r},{rng("NOTA","B")},"<="&TglAkhir))',
      'K':lambda r:f'IF(A{r}="","",SUMIFS({rng("POIN","I")},{rng("POIN","C")},A{r},{rng("POIN","B")},"<="&TglAkhir))',
      'L':lambda r:f'IF(A{r}="","",J{r}+K{r})',
      'M':lambda r:checks(f'A{r}',[(f'COUNTIF(KodePelanggan,A{r})<>1','ID GANDA'),(f'B{r}=""','ISI NAMA'),(f'AND(D{r}<>"YA",D{r}<>"TIDAK")','ISI IZIN'),(f'L{r}<0','POIN NEGATIF')])}
    b.table(s,'tblPelanggan',['ID Pelanggan','Nama','Kontak Opsional','Izin CRM','Tanggal Daftar','Catatan','Kunjungan','Belanja Netto','Terakhir Datang','Poin dari Nota','Gerak Poin','Saldo Poin','Cek'],300,cf,'ABCDEF',{'E':'in_date','H':'money','I':'date'})
    b.validate(s,'D7:D306',['YA','TIDAK']);b.audit(s,'M',306);s.set_column('B:C',27)

    s=b.base('POIN','Program poin awalnya NONAKTIF. Tentukan aturan di REFERENSI sebelum digunakan. Penukaran dicatat terpisah; tidak otomatis memberikan diskon rupiah.')
    pf={'H':lambda r:checks(f'A{r}',[(f'COUNTIF({rng("POIN","A")},A{r})<>1','ID GANDA'),(f'OR(NOT(ISNUMBER(B{r})),B{r}<=0,B{r}>TODAY())','CEK TANGGAL'),(f'COUNTIF(KodePelanggan,C{r})<>1','CEK PELANGGAN'),(f'COUNTIF(JenisPoin,D{r})<>1','CEK JENIS'),(f'OR(NOT(ISNUMBER(E{r})),E{r}<=0,E{r}<>INT(E{r}))','CEK POIN'),(f'F{r}=""','ISI ALASAN/HADIAH'),(f'AND(G{r}<>"AKTIF",G{r}<>"VOID")','CEK STATUS')]),
        'I':lambda r:f'IF(AND(A{r}<>"",H{r}="OK",G{r}="AKTIF"),N(E{r})*IF(D{r}="KOREKSI MASUK",1,-1),0)',
        'J':lambda r:f'IF(A{r}="","",SUMIFS({rng("NOTA","R")},{rng("NOTA","D")},C{r},{rng("NOTA","B")},"<="&B{r})+SUMIFS({rng("POIN","I")},{rng("POIN","C")},C{r},{rng("POIN","B")},"<"&B{r})+SUMIFS($I$7:I{r},$C$7:C{r},C{r},$B$7:B{r},B{r}))'}
    b.table(s,'tblPoin',['ID Poin','Tanggal','ID Pelanggan','Jenis','Jumlah Poin','Hadiah / Alasan','Status','Cek','Poin Bertanda','Saldo Historis'],1000,pf,'ABCDEFG',{'B':'in_date'})
    b.validate(s,'C7:C1006','=KodePelanggan');b.validate(s,'D7:D1006','=JenisPoin');b.validate(s,'G7:G1006',['AKTIF','VOID']);b.audit(s,'H',1006);s.conditional_format('J7:J1006',{'type':'cell','criteria':'<','value':0,'format':f['bad']})

    s=b.base('KASIR','Kasir 34 menu; pelanggan opsional. Tombol simpan memerlukan impor VBA. Tidak ada batas jam 22.00: operasional sumber sampai pukul 23.00.',12)
    for cell,text in [('A6','Tanggal'),('D6','Jam'),('G6','Metode'),('A8','Pelanggan opsional'),('D8','Nama pelanggan'),('G8','Kasir')]:s.write(cell,text,f['label'])
    s.write_datetime('B6',today,f['in_date']);s.write_number('E6',17/24,f['in_time']);s.write('H6','Tunai',f['input']);s.write_blank('B8',None,f['input']);s.write('H8','Pak Syarif',f['input'])
    b.formula(s,'E8',f'IF(B8="","Tanpa pelanggan",{look("B8","PELANGGAN","B")})');b.validate(s,'B8','=KodePelanggan');b.validate(s,'H6','=MetodeBayar')
    s.write_row('A12',['No','SKU','Menu','Qty','Harga / Pcs','Bruto','HPP Acuan/Pcs','Status HPP'],f['head']);s.set_column('C:C',37)
    for r in range(13,25):
        s.write_number(r-1,0,r-12,f['calc']);s.write_blank('B'+str(r),None,f['input']);s.write_blank('D'+str(r),None,f['in_qty'])
        b.formula(s,'C'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","B")})')
        b.formula(s,'E'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","C","0")})','in_money')
        b.formula(s,'F'+str(r),f'IF(B{r}="","",N(D{r})*N(E{r}))','money')
        b.formula(s,'G'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","H","0")})','money')
        b.formula(s,'H'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","I")})')
    b.validate(s,'B13:B24','=KodeMenu')
    for cell,text in [('E26','Bruto'),('E27','Diskon Rp'),('E28','TOTAL BAYAR'),('E29','Diterima'),('E30','Kembalian'),('G28','Poin nota'),('A33','Catatan')]:s.write(cell,text,f['label'])
    b.formula(s,'F26','SUM(F13:F24)','money',0);s.write('F27',0,f['in_money']);b.formula(s,'F28','F26-N(F27)','money',0);s.write_blank('F29',None,f['in_money']);b.formula(s,'F30','IF(H6="Tunai",N(F29)-F28,0)','money',0)
    b.formula(s,'H28','IF(AND(REFERENSI!B16="YA",ISNUMBER(REFERENSI!B17),REFERENSI!B17>0,B8<>""),IFERROR(INT(MAX(0,F28)/REFERENSI!B17),0),0)','qty',0)
    s.merge_range('B33:H34','',f['input']);s.merge_range('A37:L39','Harga di keranjang boleh disesuaikan sesuai transaksi yang disetujui; akan disimpan sebagai snapshot. Stok topping menggunakan takaran acuan/kalibrasi, bukan hasil timbang aktual tiap crepes. Harga SKM kosong membuat HPP belum lengkap tetapi pencatatan penjualan tetap dapat berjalan.',f['warn']);s.print_area('A1:L39')

    s=b.base('PRODUKSI','Catat dua tahap terpisah. Menghasilkan stok antara tidak dihitung sebagai penjualan. Tombol memerlukan VBA; jurnal manual memakai pasangan PRODUKSI PAKAI dan PRODUKSI MASUK.',12)
    for cell,text in [('A6','Tahap'),('A8','Tanggal'),('A10','Premiks (g)'),('D10','Bahan masuk per tahap')]:s.write(cell,text,f['label'])
    s.write('B6','PREMIKS',f['input']);s.write_datetime('B8',today,f['in_date']);s.write('B10',1000,f['in_qty']);b.validate(s,'B6',['PREMIKS','ADONAN BASAH'])
    s.write_row('A13',['Kode Bahan','Bahan','Qty Dipakai','Satuan'],f['head'])
    for i,r in enumerate(range(14,21)):
        dry=p['materials'][i]['code'];wet=['PREMIX','BTR','WTR'][i] if i<3 else ''
        b.formula(s,'A'+str(r),f'IF($B$6="PREMIKS","{dry}","{wet}")')
        b.formula(s,'B'+str(r),f'IF(A{r}="","",{look(f"A{r}","BAHAN","B")})')
        wetqty=['N($B$10)','N($B$10)*ADONAN!$B$8','N($B$10)*ADONAN!$B$9'][i] if i<3 else '0'
        b.formula(s,'C'+str(r),f'IF($B$6="PREMIKS",N($B$10)*PREMIKS!F{i+7},{wetqty})','qty')
        b.formula(s,'D'+str(r),f'IF(A{r}="","",{look(f"A{r}","BAHAN","D")})')
    s.write('A24','Kode Hasil',f['label']);b.formula(s,'B24','IF(B6="PREMIKS","PREMIX","WET")')
    s.write('D24','Berat Hasil (g)',f['label']);b.formula(s,'E24','IF(B6="PREMIKS",N(B10),N(B10)*(1+ADONAN!B8+ADONAN!B9))','qty')
    s.merge_range('A28:L31','Stok bahan mentah harus cukup sebelum membuat premiks; stok premiks dan mentega harus cukup sebelum membuat adonan basah. Kuantitas rencana menjadi gerakan stok hanya setelah disimpan. Input gram adalah premiks output untuk tahap PREMIKS, atau premiks input untuk tahap ADONAN BASAH.',f['note']);s.print_area('A1:L31')

    s=b.base('FORM','Input pendukung untuk pembelian/stok, keuangan usaha dan poin. Tombol memerlukan VBA; jurnal tetap dapat diisi langsung sesuai panduan.')
    for row,title in [(5,'01  BARANG / PEMBELIAN'),(19,'02  KAS USAHA / PRIBADI'),(29,'03  PENUKARAN / KOREKSI POIN')]:s.merge_range(row-1,0,row-1,9,title,f['head'])
    fields=[('A6','Tanggal','B6','date'),('D6','Jenis gerak','E6','text'),('A8','Kode Bahan','B8','text'),('D8','Qty satuan dasar','E8','qty'),('A10','Nominal beli total','B10','money'),('D10','Metode beli','E10','text'),('A12','Catatan / Alasan','B12','text'),('A14','Referensi bukti','B14','text'),('A20','Tanggal','B20','date'),('D20','Jenis kas','E20','text'),('A22','Nominal','B22','money'),('D22','Metode','E22','text'),('A24','Catatan','B24','text'),('A30','Tanggal','B30','date'),('D30','ID Pelanggan','E30','text'),('A32','Jenis poin','B32','text'),('D32','Jumlah poin','E32','qty'),('A34','Hadiah / Alasan','B34','text')]
    for label,title,cell,kind in fields:s.write(label,title,f['label']);s.write_blank(cell,None,f['in_'+kind] if kind!='text' else f['input'])
    for cell in ['B6','B20','B30']:s.write_datetime(cell,today,f['in_date'])
    s.write('E6','BELI',f['input']);s.write('E10','Tunai',f['input']);s.write('E20','BIAYA USAHA',f['input']);s.write('E22','Tunai',f['input']);s.write('B32','TUKAR',f['input'])
    for cell,source in [('E6',['AWAL','BELI','RUSAK','KOREKSI MASUK','KOREKSI KELUAR']),('B8','=KodeBahan'),('E10','=MetodeBayar'),('E20','=JenisKas'),('E22','=MetodeBayar'),('E30','=KodePelanggan'),('B32','=JenisPoin')]:b.validate(s,cell,source)
    s.merge_range('A37:L40','BELI dicatat sekali sebagai arus pembelian stok. Jangan masukkan pembelian yang sama lagi sebagai BIAYA USAHA. Penukaran poin tidak otomatis mengubah harga atau memberi diskon; catat hadiah sesuai kebijakan yang sudah disetujui mitra. Produksi dan pemakaian penjualan memakai tombol khusus, bukan form barang ini.',f['note']);s.set_column('B:B',29);s.print_area('A1:L40')

    # Reports use independent dates and distinguish profit from stock purchases and owner cash.
    s=b.base('DASHBOARD','Penjualan netto, biaya usaha, pembelian bahan, pemakaian stok, menu laris dan pelanggan. Laba kontribusi tidak ditampilkan bila HPP transaksi belum lengkap.',12)
    s.write('A5','Awal',f['label']);s.write_datetime('B5',today.replace(day=1),f['in_date']);s.write('D5','Akhir',f['label']);s.write_datetime('E5',today,f['in_date'])
    w.define_name('TglAwal',"='DASHBOARD'!$B$5");w.define_name('TglAkhir',"='DASHBOARD'!$E$5")
    def period(sheet,c='B'):return f'{rng(sheet,c)},">="&TglAwal,{rng(sheet,c)},"<="&TglAkhir'
    net=f'SUMIFS({rng("NOTA","P")},{period("NOTA")})'
    valid=f'{rng("NOTA","H")},"AKTIF",{rng("NOTA","O")},"OK",{period("NOTA")}'
    incomplete=f'COUNTIFS({valid},{rng("NOTA","Q")},"<>LENGKAP")'
    expenses=f'-SUMIFS({rng("KEUANGAN","I")},{rng("KEUANGAN","C")},"BIAYA USAHA",{period("KEUANGAN")})'
    hpp=f'SUMIFS({rng("NOTA","M")},{valid})'
    cards=[('PENJUALAN NETTO',net),('NOTA VALID',f'COUNTIFS({valid})'),('ITEM TERJUAL',f'SUMIFS({rng("ITEM","N")},{period("ITEM","I")})'),('BIAYA USAHA',expenses),('PEMBELIAN BAHAN',f'SUMIFS({rng("GERAK BAHAN","P")},{period("GERAK BAHAN")})'),('KONTRIBUSI SETELAH BIAYA',f'IF({incomplete}>0,"HPP BELUM LENGKAP",{net}-{hpp}-({expenses}))')]
    for i,(label,formula) in enumerate(cards):
        rr=6+i//3*5;cc=i%3*4;s.merge_range(rr,cc,rr,cc+2,label,f['kpi_label']);s.merge_range(rr+1,cc,rr+2,cc+2,'',f['kpi']);s.write_formula(rr+1,cc,'='+formula,f['kpi'],0)
    s.merge_range('A17:L17','PERIKSA SEBELUM MENUTUP HARI',f['head'])
    alerts=[('Nota bermasalah',f'COUNTIFS({rng("NOTA","A")},"<>",{rng("NOTA","O")},"<>OK")'),('Saldo bahan negatif',f'COUNTIF({rng("GERAK BAHAN","O")},"<-0.000001")'),('Nota HPP belum lengkap',incomplete),('Gerak bahan tidak valid',f'COUNTIFS({rng("GERAK BAHAN","A")},"<>",{rng("GERAK BAHAN","N")},"<>OK")'),('Poin historis negatif',f'COUNTIF({rng("POIN","J")},"<0")'),('Topping belum kalibrasi','COUNTIF(KALIBRASI!K7:K31,"ACUAN SUMBER")')]
    for i,(label,formula) in enumerate(alerts):
        rr=18+i//3*2;cc=i%3*4;s.merge_range(rr,cc,rr,cc+1,label,f['label']);s.write_formula(rr,cc+2,'='+formula,f['qty'],0)
        s.conditional_format(rr,cc+2,rr,cc+2,{'type':'cell','criteria':'>','value':0,'format':f['warn']})
    s.write_row('A24',['SKU','Menu','Qty','Bruto Sebelum Diskon'],f['head']);s.set_column('B:B',34);s.set_column('D:D',27)
    for r in range(25,75):
        mr=r-18;b.formula(s,'A'+str(r),f'IF(MENU!A{mr}="","",MENU!A{mr})');b.formula(s,'B'+str(r),f'IF(A{r}="","",MENU!B{mr})')
        b.formula(s,'C'+str(r),f'IF(A{r}="","",SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},A{r},{period("ITEM","I")}))','qty')
        b.formula(s,'D'+str(r),f'IF(A{r}="","",SUMIFS({rng("ITEM","K")},{rng("ITEM","C")},A{r},{rng("ITEM","N")},">0",{period("ITEM","I")}))','money')
    chart=w.add_chart({'type':'bar'});chart.add_series({'name':'Qty per menu','categories':"='DASHBOARD'!$B$25:$B$58",'values':"='DASHBOARD'!$C$25:$C$58",'fill':{'color':'#A05423'},'border':{'none':True}});chart.set_title({'name':'Penjualan 34 menu'});chart.set_legend({'none':True});chart.set_y_axis({'reverse':True,'num_font':{'size':8}});s.insert_chart('F24',chart,{'x_scale':1.3,'y_scale':2})
    s.merge_range('F68:L72','Pembelian bahan adalah arus uang keluar, bukan biaya kedua setelah HPP resep. Modal/prive dipisahkan. HPP acuan bukan laporan laba bersih final: belum mencakup biaya yang belum dicatat. Harga dan takaran baru tidak mengubah snapshot transaksi lama.',f['note']);s.print_area('A1:L74')

    s=b.base('TUTUP KASIR','Rekonsiliasi laci tunai per hari. QRIS/transfer/e-wallet ditampilkan terpisah dan tidak dimasukkan ke uang fisik laci. Modal awal laci adalah saldo pembukaan, bukan transaksi modal masuk baru.')
    for cell,text in [('A5','Tanggal'),('A7','Modal awal laci'),('D7','Uang fisik laci')]:s.write(cell,text,f['label'])
    s.write_datetime('B5',today,f['in_date']);s.write_blank('B7',None,f['in_money']);s.write_blank('E7',None,f['in_money'])
    rows={10:('Penjualan tunai',f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5,{rng("NOTA","E")},"Tunai")'),11:('Pembelian tunai',f'SUMIFS({rng("GERAK BAHAN","P")},{rng("GERAK BAHAN","B")},$B$5,{rng("GERAK BAHAN","H")},"Tunai")'),12:('Kas lain bertanda',f'SUMIFS({rng("KEUANGAN","I")},{rng("KEUANGAN","B")},$B$5,{rng("KEUANGAN","E")},"Tunai")'),14:('Perkiraan isi laci','N(B7)+B10-B11+B12'),15:('Selisih fisik','IF(ISNUMBER(E7),E7-B14,"")')}
    for r,(label,formula) in rows.items():s.write('A'+str(r),label,f['label']);b.formula(s,'B'+str(r),formula,'money')
    for r,method in enumerate(['QRIS','Transfer','E-wallet'],19):s.write('A'+str(r),method,f['label']);b.formula(s,'B'+str(r),f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$B$5,{rng("NOTA","E")},"{method}")','money')
    s.merge_range('D10:L16','Penjualan tunai dihitung sebesar netto yang menjadi hak usaha, bukan uang diterima sebelum kembalian. BELI dan biaya/prive hanya mengurangi laci bila metode Tunai. Isi uang fisik setelah dihitung; selisih tidak otomatis membuat koreksi stok/biaya.',f['note']);s.set_column('A:A',29);s.print_area('A1:L23')

    s=b.base('STRUK','Pilih No Nota; cetak/PDF melalui Excel desktop. Struk hanya untuk nota AKTIF yang valid. Tidak memakai data pribadi pelanggan sebagai kewajiban transaksi.')
    s.set_portrait();s.write('A5','No Nota',f['label']);s.write_blank('B5',None,f['input']);b.validate(s,'B5','=KodeNota')
    s.merge_range('A7:F7','MURITA CREPES',f['title']);b.formula(s,'A8','IF(B5="","PILIH NOTA",B5)');b.formula(s,'D8',look('$B$5','NOTA','B'),'date');b.formula(s,'F8',look('$B$5','NOTA','C'),'time')
    b.formula(s,'A10',f'IF(B5="","PILIH NOTA",IF(AND({look("B5","NOTA","H")}="AKTIF",{look("B5","NOTA","O")}="OK"),"SIAP CETAK","NOTA VOID / TIDAK VALID"))','warn')
    s.write_row('A12',['No','SKU','Menu','Qty','Harga','Jumlah'],f['head'])
    # Unique item IDs are selected in source row order; every slot of a 12-line cart is supported.
    for r in range(13,25):
        pos=r-12
        b.formula(s,'G'+str(r),f'IF($B$5="","",IFERROR(_xlfn.AGGREGATE(15,6,(ROW({rng("ITEM","B")})-6)/({rng("ITEM","B")}=$B$5),{pos}),""))')
        b.formula(s,'A'+str(r),f'IF(G{r}="","",{pos})')
        for c,source in [('B','C'),('D','D'),('E','E'),('F','K')]:b.formula(s,c+str(r),f'IF($G{r}="","",INDEX({rng("ITEM",source)},$G{r}))','money' if c in ['E','F'] else 'calc')
        b.formula(s,'C'+str(r),f'IF(B{r}="","",{look(f"B{r}","MENU","B")})')
    for r,label,src in [(26,'Bruto','J'),(27,'Diskon','F'),(28,'TOTAL','K'),(29,'Diterima','G'),(30,'Kembalian','L'),(31,'Poin didapat','N')]:s.write('D'+str(r),label,f['label']);b.formula(s,'F'+str(r),f'IF($B$5="","",{look("$B$5","NOTA",src,"0")})','money')
    s.merge_range('A34:F35','Terima kasih. Sampai jumpa kembali di Murita Crepes.',f['note']);s.set_column('C:C',35);s.set_column('G:G',None,None,{'hidden':True});s.print_area('A7:F35');s.repeat_rows(6,11);s.fit_to_pages(1,1)

    s=b.base('REFERENSI','Parameter, sumber dan keputusan implementasi. Input berasal dari catatan mitra; kebijakan loyalty tidak dikarang dan kontak pelanggan tidak dimuat otomatis.')
    params={6:('Pemilik','Pak Syarif'),8:('Jam operasional','17.00–23.00; Senin libur rutin'),10:('SKM / crepes (g)',p['skm_per_crepe']),11:('Gas / crepes (menit)',p['gas_minutes_per_crepe']),12:('Kemasan / crepes (lembar)',p['pack_per_crepe']),16:('Aktifkan poin (YA/TIDAK)','TIDAK'),17:('Rupiah netto untuk 1 poin',None)}
    for r,(label,value) in params.items():s.write('A'+str(r),label,f['label']);s.write('B'+str(r),value,f['input'])
    b.validate(s,'B16',['YA','TIDAK']);s.set_column('A:A',34);s.set_column('B:B',43)
    s.write_url('A20',p['source_sheet'],f['nav'],'MoM • 24 Oktober 2026');s.write_url('A22',p['source_folder'],f['nav'],'Folder referensi Murita Crepes')
    notes=['Sumber biaya memiliki kesalahan referensi: harga SKM mengambil total biaya adonan; angka biaya adonan lama juga tidak sama dengan hasil penjumlahan resep. Harga SKM dikosongkan; HPP belum lengkap sampai diisi.',
      'Premiks terdiri dari tujuh bahan dengan total 2.908,5 g. Mentega dan air memakai rasio 0,065 dan 0,83 terhadap premiks aktual. Hasil 30 crepes/kg mengikuti asumsi kerja MoM (kisaran aktual 28–30).',
      'Produksi premiks memotong tujuh bahan mentah. Produksi adonan memotong premiks, mentega, air. Penjualan memotong hanya WET + SKM + gas/kemasan + topping. Tidak ada pemotongan ganda bahan mentah saat penjualan.',
      'Kalibrasi topping memerlukan tiga takaran dan jumlah sendok per porsi. Sebelum kalibrasi, pakai angka acuan sumber. Satuan pisang adalah buah; marshmallow pcs; jangan disamakan dengan gram.',
      'File database 30 hari dan rekap tally adalah referensi historis. Judul/batas bulan perlu dikonfirmasi dan tidak mengandung identitas nota/pelanggan lengkap; tidak dimuat sebagai transaksi operasional agar tidak dibuat-buat atau terduplikasi.',
      'Izin CRM tidak berarti izin mengirim pesan otomatis. Tidak ada broadcast, email atau pengiriman data pelanggan dalam modul. Program poin awalnya mati; bila aktif, poin dibulatkan turun per NOTA netto dan disimpan sebagai snapshot.',
      'HPP dan konsumsi snapshot berasal dari resep saat transaksi. Mengubah harga/kalibrasi tidak mengubah jurnal lama. Biaya pembelian stok tidak dihitung lagi sebagai biaya di samping HPP acuan. Ini alat operasional, bukan pembukuan audit/FIFO.',
      'Stok fisik, saldo kas, daftar pelanggan, kebijakan reward dan nominal biaya harian belum diberikan. File utama tidak mengimpor saldo atau transaksi contoh. Bahan terlacak tidak boleh negatif; air/gas awalnya tidak dilacak secara fisik.',
      'Jangan mengganti kode bahan/menu/pelanggan yang telah dipakai. Jangan mengubah nota yang sudah final; VOID harus dilakukan untuk seluruh nota beserta dampak stok/poinnya. Status VOID saat ini memengaruhi laporan historis; arsipkan file per periode.' ]
    for i,text in enumerate(notes):r=24+i*3;s.merge_range(r,0,r+1,11,text,f['body'])
    for column,name,values in [('N','JenisGerak',movements),('O','MetodeBayar',methods),('P','JenisKas',['BIAYA USAHA','PRIVE','MODAL MASUK','PENDAPATAN LAIN']),('Q','JenisPoin',['TUKAR','KOREKSI MASUK','KOREKSI KELUAR'])]:
        s.write_column(column+'2',values);w.define_name(name,f"='REFERENSI'!${column}$2:${column}${len(values)+1}")
    s.set_column('N:Q',None,None,{'hidden':True});s.print_area('A1:L52')
    order=['MULAI','DASHBOARD','KASIR','PRODUKSI','FORM','MENU','BAHAN','PREMIKS','ADONAN','KALIBRASI','RESEP','GERAK BAHAN','NOTA','ITEM','KEUANGAN','TUTUP KASIR','PELANGGAN','POIN','STRUK','REFERENSI']
    b.close(order);return path

if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('--output',type=Path,default=ROOT/'output/POS_Murita_Crepes.xlsx');print(build(a.parse_args().output))
