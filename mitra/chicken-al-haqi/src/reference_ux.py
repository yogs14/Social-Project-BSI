"""Use the uploaded Ponorogo workbook's actual cell styles and geometry."""
import json
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED


def finish_native_buttons(path):
    """Keep the reference's 11 pt Calibri on native Excel form controls."""
    with ZipFile(path) as source:
        parts={name:source.read(name) for name in source.namelist()}
    for name,data in parts.items():
        if name.endswith('.vml'):
            parts[name]=re.sub(rb'<font face="[^"]*" size="[^"]*"',
                               b'<font face="Calibri" size="220"',data)
    temporary=path.with_suffix('.controls.xlsx')
    with ZipFile(temporary,'w',ZIP_DEFLATED) as out:
        for name,data in parts.items():out.writestr(name,data)
    temporary.replace(path)


class ReferenceView:
    def __init__(self, builder, name, reference):
        self.b = builder
        spec = json.loads((Path(__file__).resolve().parents[1] / 'data/ux_pos_ponorogo.json').read_text())
        self.spec = spec['sheets'][reference]
        self.formats = {key: builder.w.add_format(value) for key, value in spec['styles'].items()}
        self.s = builder.w.add_worksheet(name)
        builder.s[name] = self.s
        s = self.s
        s.hide_gridlines(2)
        s.set_default_row(self.spec['default_row_height'])
        s.set_zoom(self.spec['zoom'])
        s.set_tab_color('#902A20' if name == 'KASIR' else '#C96B1D')
        for column, value in self.spec['columns'].items():
            s.set_column(value['min'] - 1, value['max'] - 1,
                         max(0, value['width'] - 0.7109375), None, {'hidden': value['hidden']})
        for row, height in self.spec['rows'].items():
            s.set_row(int(row) - 1, height)
        for cell in self.spec['cells']:
            s.write_blank(cell, None, self.f(cell))
        for area in self.spec['merges']:
            s.merge_range(area, '', self.f(area.split(':')[0]))
        s.set_landscape()
        s.set_paper(9)
        s.fit_to_pages(1, 1)
        s.set_margins(.25, .25, .3, .3)
        s.set_footer('&LChicken Al Haqi&R&P / &N')

    def f(self, cell):
        return self.formats.get(self.spec['cells'].get(cell), self.b.f['calc'])

    def text(self, cell, value):
        self.s.write(cell, value, self.f(cell))

    def formula(self, cell, value, cached=''):
        self.s.write_formula(cell, '=' + value, self.f(cell), cached)

    def date(self, cell, value):
        self.s.write_datetime(cell, value, self.f(cell))

    def link(self, cell, tab, caption):
        self.s.write_url(cell, f"internal:'{tab}'!A1", self.f(cell), caption)


def cashier_view(b, profile, today, stock_at, menu_lookup, rng, look):
    v = ReferenceView(b, 'KASIR', 'KASIR')
    s = v.s
    v.text('B1', 'KASIR - Chicken Al Haqi')
    v.text('B2', 'Pilih menu, isi jumlah, lalu simpan nota. Harga, kembalian, dan stok potongan dihitung otomatis.')
    for cell, tab, label in [('B3','HARIAN','📊 Dashboard'),('C3','NOTA','📋 Buku Transaksi'),
                              ('D3','STOK','📦 Stok Realtime'),('E3','TUTUP KASIR','🔒 Tutup Kasir'),
                              ('F3','MULAI','📖 Panduan')]:
        v.link(cell, tab, label)
    v.text('B5', '1️⃣   INFO NOTA')
    for row, label in [(6,'No Nota'),(7,'Tanggal'),(8,'Jam'),(9,'Nama Kasir'),
                        (10,'Metode Bayar'),(11,'Channel'),(12,'Jumlah Pelanggan')]:
        v.text(f'B{row}', label)
    v.text('C6', 'Otomatis saat disimpan')
    v.date('C7', today)
    # Excel time values support the 06:30 weekend opening.
    s.write('C8', 9/24, b.f['ref_time'])
    v.text('C9', 'Kasir 1')
    v.text('C10', 'Tunai')
    v.text('C11', 'Dine-in')
    v.text('C12', 1)
    v.text('E7', '← Ctrl+; untuk tanggal hari ini')
    v.text('E8', '← contoh jam: 06:30 atau 09:00')
    for cell, source in [('C10','=MetodeBayar'),('C11','=ChannelPenjualan'),('B16:B27','=MenuKasir')]:
        b.validate(s, cell, source)
    for area in ['C12','C16:C27']:
        s.data_validation(area, {'validate':'integer','criteria':'>','value':0,
                                'input_title':'Jumlah','input_message':'Isi jumlah bulat, misalnya 1 atau 2.'})
    v.text('H5', '📈   PENJUALAN HARI INI')
    for row,label,formula in [
        (6,'Omzet berjalan',f'SUMIFS({rng("NOTA","P")},{rng("NOTA","B")},$C$7)'),
        (7,'Jumlah nota',f'COUNTIFS({rng("NOTA","B")},$C$7,{rng("NOTA","H")},"AKTIF",{rng("NOTA","O")},"OK")'),
        (8,'Item terjual',f'SUMIFS({rng("ITEM","N")},{rng("ITEM","I")},$C$7)'),
        (9,'Jumlah pelanggan',f'SUMIFS({rng("NOTA","R")},{rng("NOTA","B")},$C$7,{rng("NOTA","H")},"AKTIF",{rng("NOTA","O")},"OK")'),
        (10,'Rata-rata per nota','IF(I7=0,"",I6/I7)')]:
        v.text(f'H{row}',label)
        v.formula(f'I{row}',formula,0 if row!=10 else '')
    v.text('B14','2️⃣   KERANJANG BELANJA PEMBELI')
    for cell,label in [('B15','Menu (pilih dari dropdown)'),('C15','Jumlah'),('D15','Harga'),('E15','Subtotal'),('F15','Cek Stok')]:
        v.text(cell,label)
    for row in range(16,28):
        v.formula(f'D{row}',f'IF(B{row}="","",{menu_lookup(f"B{row}","D","0")})')
        v.formula(f'E{row}',f'IF(OR(B{row}="",C{row}=""),"",N(C{row})*N(D{row}))')
        v.formula(f'X{row}',menu_lookup(f'B{row}','B'))
        v.formula(f'Y{row}',menu_lookup(f'B{row}','C'))
        balance = stock_at(f'X{row}','$C$7')
        need = f'SUMIF($X$16:$X$27,X{row},$C$16:$C$27)'
        status=(f'IF(B{row}="","",IF(X{row}="","CEK MENU",IF({look(f"X{row}","SETUP","H")}<>"YA",'
                f'"TIDAK DILACAK",IF({look(f"X{row}","STOK","J")}<>"OK","BELUM AWAL",'
                f'IF($C$7<{look(f"X{row}","STOK","C","0")},"BELUM MULAI",'
                f'IF({balance}<{need},"STOK KURANG","CUKUP"))))))')
        v.formula(f'F{row}',status)
    for area in ['F16:F27','K15:K18']:
        for text,style in [('CUKUP','good'),('TERSEDIA','good'),('KURANG','bad'),('BELUM','warn')]:
            s.conditional_format(area,{'type':'text','criteria':'containing','value':text,'format':b.f[style]})
    for row,label in [(29,'Subtotal'),(30,'Diskon / Promo (Rp)'),(31,'TOTAL BAYAR'),(32,'Uang Diterima'),(33,'Kembalian')]:
        v.text(f'B{row}',label)
    v.formula('C29','SUM(E16:E27)',0)
    v.text('C30',0)
    v.formula('C31','MAX(0,C29-N(C30))',0)
    v.formula('C33','IF(OR(C32="",C10<>"Tunai"),"",N(C32)-C31)')
    v.formula('B35','IF(COUNTA(B16:B27)=0,"Pilih menu dan isi jumlah untuk mulai melayani.",IF(COUNTIF(F16:F27,"STOK KURANG")>0,"⚠ Stok kurang — periksa stok sebelum menyimpan.",IF(COUNTIF(F16:F27,"BELUM AWAL")>0,"⚠ Isi stok awal agar jumlahnya bisa diperiksa.","✓ Stok cukup untuk pesanan ini.")))')
    v.text('B37','3️⃣   SIMPAN TRANSAKSI')
    v.text('B38','Klik SIMPAN TRANSAKSI setelah pesanan selesai. Nota tersimpan, stok mengikuti penjualan, dan keranjang siap untuk pembeli berikutnya.')
    v.text('H13','⚠️   STOK POTONGAN AYAM')
    for cell,label in [('H14','Potongan'),('I14','Stok'),('J14','Dipesan'),('K14','Status')]:v.text(cell,label)
    for row,item in enumerate(profile['items'][:4],15):
        code='"'+item['code']+'"'
        opened=f'AND({look(code,"STOK","J")}="OK",$C$7>={look(code,"STOK","C","0")})'
        v.text(f'H{row}',item['name'])
        v.formula(f'I{row}',f'IF({opened},{stock_at(code,"$C$7")},"")')
        v.formula(f'J{row}',f'SUMIF($X$16:$X$27,{code},$C$16:$C$27)',0)
        v.formula(f'K{row}',f'IF(NOT({opened}),"BELUM AWAL",IF(I{row}<J{row},"KURANG",IF(I{row}=0,"HABIS","TERSEDIA")))')
    v.text('H24','🍽️   MENU TERLARIS HARI INI')
    for cell,label in [('H25','Menu'),('I25','Qty'),('J25','Omzet')]:v.text(cell,label)
    for row,item in enumerate(profile['items'],23):
        code='"'+item['code']+'"'
        v.text(f'AC{row}',item['name'])
        v.formula(f'AD{row}',f'SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},{code},{rng("ITEM","I")},$C$7)')
        v.formula(f'AE{row}',f'SUMIFS({rng("ITEM","K")},{rng("ITEM","C")},{code},{rng("ITEM","I")},$C$7,{rng("ITEM","N")},">0")')
        v.formula(f'AF{row}',f'AD{row}+(28-ROW())/1000')
    for row in range(26,31):
        rank=row-25
        match=f'MATCH(LARGE($AF$23:$AF$27,{rank}),$AF$23:$AF$27,0)'
        qty=f'INDEX($AD$23:$AD$27,{match})'
        for column,formula in [('H',f'INDEX($AC$23:$AC$27,{match})'),('I',qty),('J',f'INDEX($AE$23:$AE$27,{match})')]:
            v.formula(f'{column}{row}',f'IF({qty}=0,"",{formula})')
    v.text('H33','AKTIFKAN TOMBOL (sekali saja)')
    for cell,label in [('H34','1. Simpan salinan sebagai .xlsm.'),('H35','2. Impor Modul_ChickenAlHaqi.bas.'),
                        ('H36','3. Jalankan PasangTombol.'),('H37','4. Simpan file, lalu mulai melayani.')]:v.text(cell,label)
    v.text('H38','Belanja dan kas usaha tersedia pada tab masing-masing atau panel di bawah kasir.')
    v.text('B50','Chicken Al Haqi  •  Harga menu dapat disesuaikan pada SETUP  •  Simpan file dengan Ctrl+S')
    s.set_column('N:AF',None,None,{'hidden':True})
    s.freeze_panes(15,1)
    s.set_top_left_cell('A1')
    s.set_selection('B16')
    s.print_area('B1:K48')
    # Native Excel form controls, already visible before the module is imported.
    for index,(caption,macro) in enumerate([('SIMPAN TRANSAKSI','SimpanTransaksi'),('NOTA BARU','NotaBaru'),
                                           ('VOID NOTA TERAKHIR','VoidNotaTerakhir'),('TUTUP KASIR','TutupKasir')]):
        rect=v.spec['buttons'][index]
        s.insert_button('A1',{'caption':caption,'macro':macro,
                              'width':round(rect['width']*4/3),'height':round(rect['height']*4/3),
                              'x_offset':round(rect['margin-left']*4/3),'y_offset':round(rect['margin-top']*4/3)})
    return s


def closing_view(b,today):
    v=ReferenceView(b,'TUTUP KASIR','Tutup Kasir')
    s=v.s
    v.text('B1','🔒  TUTUP KASIR — Chicken Al Haqi')
    v.text('B2','Isi tanggal, modal awal laci, dan uang fisik. Rekap serta selisih kas dihitung otomatis.')
    for cell,tab,label in [('B3','KASIR','🧾 Kasir'),('C3','HARIAN','📊 Dashboard'),('E3','STOK','📦 Stok'),('F3','NOTA','📋 Transaksi')]:v.link(cell,tab,label)
    v.text('B4','⚙️   PILIH HARI')
    v.text('B5','Tanggal')
    v.date('C5',today)
    v.text('B7','💰   RINGKASAN PENJUALAN')
    for row,label,formula in [(8,'Jumlah Nota',"'HARIAN'!E8"),(9,'Jumlah Pelanggan',"SUMIFS('NOTA'!$R$7:$R$1506,'NOTA'!$B$7:$B$1506,$C$5,'NOTA'!$H$7:$H$1506,\"AKTIF\",'NOTA'!$O$7:$O$1506,\"OK\")"),
                              (10,'Item Terjual',"SUMIFS('ITEM'!$N$7:$N$6006,'ITEM'!$I$7:$I$6006,$C$5)"),
                              (13,'Omzet Bersih',"'HARIAN'!A8"),(14,'Rata-rata per Nota','IF(C8=0,"",C13/C8)'),
                              (15,'HPP Tercatat',"IF(COUNTIFS('NOTA'!$B$7:$B$1506,$C$5,'NOTA'!$H$7:$H$1506,\"AKTIF\",'NOTA'!$O$7:$O$1506,\"OK\",'NOTA'!$N$7:$N$1506,\"<>DIKONFIRMASI\")>0,\"Belum lengkap\",SUMIFS('NOTA'!$M$7:$M$1506,'NOTA'!$B$7:$B$1506,$C$5,'NOTA'!$H$7:$H$1506,\"AKTIF\",'NOTA'!$O$7:$O$1506,\"OK\"))"),
                              (16,'Selisih Omzet - HPP','IF(ISNUMBER(C15),C13-C15,"HPP belum lengkap")')]:
        v.text(f'B{row}',label);v.formula(f'C{row}',formula)
    v.text('E7','💳   RINCIAN PEMBAYARAN')
    for row,label,formula in [(8,'Tunai',"'HARIAN'!B18"),(9,'QRIS',"'HARIAN'!F20"),(10,'Transfer',"'HARIAN'!F21"),(11,'Lainnya',"'HARIAN'!F22")]:
        v.text(f'E{row}',label);v.formula(f'F{row}',formula,0)
    v.text('H7','🧮   COCOKKAN UANG LACI')
    for row,label in [(8,'Modal Awal Laci'),(9,'Penjualan Tunai'),(10,'Kas Seharusnya'),(11,'Uang Fisik di Laci'),(12,'SELISIH KAS')]:v.text(f'H{row}',label)
    for cell,formula in [('I9',"'HARIAN'!B18"),('I10','K14'),('I12','K15')]:v.formula(cell,formula)
    # One-way aliases keep the established dashboard and historical calculations intact.
    for cell,formula in [('K8','IF(ISNUMBER(I8),I8,"")'),('K11','IF(ISNUMBER(I11),I11,"")'),
                         ('K9',"'HARIAN'!B18"),('K10',"'HARIAN'!B19"),('K12',"'HARIAN'!B20"),
                         ('K14','IF(ISNUMBER(K8),K8+K9-K10+K12,"")'),
                         ('K15','IF(AND(ISNUMBER(K11),ISNUMBER(K14)),K11-K14,"")')]:v.formula(cell,formula)
    v.text('H14','Kas seharusnya sudah memperhitungkan belanja dan arus kas usaha tunai.')
    # No empty data table beneath the three closing panels.
    plain=b.w.add_format({'font_name':'Calibri','font_size':9,'font_color':'#808080','bg_color':'white'})
    for row in range(17,24):
        for column in range(7,10):s.write_blank(row-1,column,None,plain)
    s.write('H17','Hijau = sesuai  •  Merah = periksa lagi',plain)
    v.text('B19','SELESAI MELAYANI HARI INI')
    v.text('B20','Periksa nota dan stok, cocokkan uang laci, lalu simpan file. Isi ulang modal dan uang fisik saat tanggal diganti. Simpan PDF/arsip untuk rekap harian.')
    s.set_column('K:L',None,None,{'hidden':True})
    s.conditional_format('I12',{'type':'formula','criteria':'AND(ISNUMBER(I12),I12<>0)','format':b.f['bad']})
    s.conditional_format('I12',{'type':'formula','criteria':'AND(ISNUMBER(I12),I12=0)','format':b.f['good']})
    for cell in ['I8','I11']:
        s.data_validation(cell,{'validate':'decimal','criteria':'>=','value':0,'input_title':'Uang tunai','input_message':'Isi nominal rupiah sesuai uang fisik.'})
    s.print_area('B1:J24')
    s.set_selection('I8')
    return s
