"""Stock and closing screens reuse the actual Pak Seger POS sheet layouts."""
import copy,datetime as dt
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule,CellIsRule
from openpyxl.utils import column_index_from_string,get_column_letter
from compact_layout import BLOCKS

FIRST_CLOSE=12077
LAST_CLOSE=12442
MONEY='"Rp" #,##0;[Red]("Rp" #,##0);"Rp 0"'
QTY='#,##0.###;[Red](#,##0.###);"0"'
ROOT=Path(__file__).resolve().parents[1]

def area(logical,col):
    sheet,off=BLOCKS[logical]
    cap={'NOTA':1000,'ITEM':3000,'GERAK BAHAN':6000,'KEUANGAN':1000,'MENU':50,'BAHAN':60}[logical]
    return f"'{sheet}'!${col}${7+off}:${col}${6+off+cap}"

def go(ws,cell,text,target):
    c=ws[cell];c.value=text;c.hyperlink='#'+target
    c.font=Font(name='Calibri',size=10,bold=True,color='FFFFFF')
    c.fill=PatternFill('solid',fgColor='C96B1D')
    c.alignment=Alignment(vertical='center')

def copy_screen(dst,ref,title):
    # Retain the source's UI, never its restaurant data/formula dependencies.
    if title in dst:dst.remove(dst[title])
    new=dst.create_sheet(title)
    for row in ref:
        for c in row:
            if c.value is None and not c.has_style:continue
            d=new.cell(c.row,c.column)
            d.value=None if c.data_type=='f' else c.value
            for attr in ['font','fill','border','alignment','protection']:
                setattr(d,attr,copy.copy(getattr(c,attr)))
            d.number_format=c.number_format
    for merged in ref.merged_cells.ranges:new.merge_cells(str(merged))
    for key,dim in ref.column_dimensions.items():
        lo=dim.min or column_index_from_string(key)
        hi=dim.max or lo
        for index in range(lo,hi+1):
            x=copy.copy(dim);x.parent=new;x.min=index;x.max=index;x.index=get_column_letter(index)
            new.column_dimensions[x.index]=x
    for key,dim in ref.row_dimensions.items():
        x=copy.copy(dim);x.parent=new;new.row_dimensions[key]=x
    new.sheet_view.showGridLines=False;new.sheet_view.zoomScale=85
    new.sheet_properties.tabColor='902A20'
    new.page_setup.orientation='landscape';new.page_setup.paperSize=new.PAPERSIZE_A4
    new.sheet_properties.pageSetUpPr.fitToPage=True
    new.page_setup.fitToWidth=1;new.page_setup.fitToHeight=1
    new.page_margins.left=.25;new.page_margins.right=.25
    return new

def set_formula(ws,cell,text,num=None):
    ws[cell]='='+text
    if num:ws[cell].number_format=num

def style_source_palette(dst):
    fills={'302D29':'902A20','8C4D28':'A63B25','EDF4FB':'E2EFDA','F5F2EC':'FBF3E7'}
    fonts={'A05423':'C96B1D','17578B':'1B5E20'}
    for ws in dst:
        for row in ws:
            for c in row:
                if c.fill.fgColor.type=='rgb' and str(c.fill.fgColor.rgb)[-6:] in fills:
                    v=copy.copy(c.fill);v.fgColor=fills[str(c.fill.fgColor.rgb)[-6:]];c.fill=v
                if c.font.color and c.font.color.type=='rgb' and str(c.font.color.rgb)[-6:] in fonts:
                    v=copy.copy(c.font);v.color=fonts[str(c.font.color.rgb)[-6:]];c.font=v
                if c.data_type!='f' and isinstance(c.value,str):
                    c.value=c.value.replace('Sel biru','Sel hijau').replace('sel biru','sel hijau')

def stock_screen(dst,source):
    ws=copy_screen(dst,source['Stok Realtime'],'STOK')
    ws['B1']='STOK REALTIME | MUFRITA CREPES'
    ws['B2']='Stok awal + beli/hasil produksi/koreksi masuk - pemakaian jual/produksi/koreksi keluar - rusak. Tanggal mengikuti Ringkasan.'
    for a,t in [('B4','Nilai stok (acuan)'),('E4','Bahan status merah'),('H4','Bahan status kuning'),('K4','Estimasi belanja')]:ws[a]=t
    for a,t,target in [('B3','Kasir',"'KASIR'!A1"),('E3','Dashboard',"'RINGKASAN'!A1"),('H3','Tutup Kasir',"'TUTUP KASIR'!B1"),('K3','Master Bahan',"'KATALOG'!A225")]:go(ws,a,t,target)
    headers=['Bahan Baku','Satuan','Stok Awal','Masuk','Terpakai','Terbuang','STOK KINI','Nilai (Rp)','Pakai / Hari','Cukup (Hari)','Status','Saran Beli','Kode Bahan']
    for c,h in enumerate(headers,2):ws.cell(7,c,h)
    for r in range(8,68):
        # Expand the source's 20 slots to the existing 60-material capacity.
        for c in range(2,15):
            d=ws.cell(r,c);base=source['Stok Realtime'].cell(8,c)
            for attr in ['font','fill','border','alignment']:
                setattr(d,attr,copy.copy(getattr(base,attr)))
            d.alignment=Alignment(vertical='center',wrap_text=c in (2,12))
            if c>=4:d.number_format=QTY
        ws.row_dimensions[r].height=30
        mr=231+r-8
        code=f'$N{r}'
        qty=area('GERAK BAHAN','M');material=area('GERAK BAHAN','E');day=area('GERAK BAHAN','B');kind=area('GERAK BAHAN','D')
        def moves(kinds,recent=False):
            terms=[]
            for k in kinds:
                tail=f',{day},">="&TglAkhir-6' if recent else ''
                terms.append(f'SUMIFS({qty},{material},{code},{day},"<="&TglAkhir,{kind},"{k}"{tail})')
            return '+'.join(terms)
        wrap=lambda s:f'IF({code}="","",{s})'
        set_formula(ws,f'N{r}',f'IF(KATALOG!A{mr}="","",KATALOG!A{mr})')
        ws[f'N{r}'].hyperlink=f"#'KATALOG'!A{mr}"
        set_formula(ws,f'B{r}',wrap(f'KATALOG!B{mr}'))
        set_formula(ws,f'C{r}',wrap(f'KATALOG!D{mr}'))
        tracked=f'KATALOG!H{mr}="YA"'
        for c,expr in [('D',moves(['AWAL'])),('E',moves(['BELI','PRODUKSI MASUK','KOREKSI MASUK'])),('F','-('+moves(['PRODUKSI PAKAI','JUAL PAKAI','KOREKSI KELUAR'])+')'),('G','-('+moves(['RUSAK'])+')')]:
            set_formula(ws,f'{c}{r}',wrap(f'IF({tracked},{expr},"")'))
        set_formula(ws,f'H{r}',wrap(f'IF({tracked},ROUND(D{r}+E{r}-F{r}-G{r},6),"")'))
        set_formula(ws,f'I{r}',wrap(f'IF({tracked},IF(KATALOG!K{mr}="OK",H{r}*KATALOG!G{mr},"CEK HARGA"),"")'),MONEY)
        set_formula(ws,f'J{r}',wrap(f'IF({tracked},-({moves(["JUAL PAKAI","PRODUKSI PAKAI"],True)})/7,"")'))
        set_formula(ws,f'K{r}',wrap(f'IF(AND(ISNUMBER(J{r}),J{r}>0),ROUND(H{r}/J{r},1),"")'))
        lim=f'KATALOG!I{mr}'
        status=f'IF(NOT({tracked}),"TIDAK DILACAK",IF(H{r}<0,"MERAH",IF(D{r}+E{r}=0,"BELUM AWAL",IF(H{r}=0,"MERAH",IF(OR(NOT(ISNUMBER({lim})),{lim}<0),"ISI BATAS",IF(H{r}<={lim},"KUNING","HIJAU"))))))'
        set_formula(ws,f'L{r}',wrap(status))
        pack=f'KATALOG!F{mr}'
        recommendation=f'IF(AND({tracked},ISNUMBER({lim}),{lim}>=0,ISNUMBER({pack}),{pack}>0,H{r}<={lim}),MAX(1,ROUNDUP(({lim}-H{r})/{pack},0))*{pack},0)'
        set_formula(ws,f'M{r}',wrap(recommendation))
        set_formula(ws,f'O{r}',f'IF(AND(N{r}<>"",M{r}>0),M{r}*KATALOG!G{mr},0)',MONEY)
        set_formula(ws,f'P{r}',f'IF(AND(N{r}<>"",KATALOG!H{mr}="YA",H{r}>0,KATALOG!K{mr}<>"OK"),1,0)')
        set_formula(ws,f'Q{r}',f'IF(AND(N{r}<>"",M{r}>0,KATALOG!K{mr}<>"OK"),1,0)')
    set_formula(ws,'B5','IF(SUM(P8:P67)>0,"CEK HARGA",SUM(I8:I67))',MONEY)
    set_formula(ws,'E5','COUNTIF(L8:L67,"MERAH")','0')
    set_formula(ws,'H5','COUNTIF(L8:L67,"KUNING")','0')
    set_formula(ws,'K5','IF(SUM(Q8:Q67)>0,"CEK HARGA",SUM(O8:O67))',MONEY)
    ws['B69']='Merah = habis/negatif. Kuning = mencapai batas belanja. Hijau = di atas batas. BELUM AWAL / ISI BATAS perlu dilengkapi.'
    ws.merge_cells('B69:N69');ws['B69'].alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[69].height=30
    ws['B70']='Pakai/hari memakai rerata 7 hari kalender; kecukupan hanya estimasi. Saran beli dibulatkan ke kemasan. Nilai stok memakai harga acuan saat ini.'
    ws.merge_cells('B70:N70');ws['B70'].alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[70].height=30
    for a,t,target in [('B72','Input Belanja / Opname',"'HARIAN'!A43"),('F72','Produksi Adonan',"'HARIAN'!A1"),('J72','Atur Batas / Harga',"'KATALOG'!A225")]:go(ws,a,t,target)
    # The source has a footer at row 30, now occupied by the larger table.
    widths={'A':2.5,'B':32,'C':8,'D':11,'E':12,'F':12,'G':10,'H':13,'I':15,'J':11,'K':11,'L':18,'M':13,'N':12}
    for c,v in widths.items():ws.column_dimensions[c].width=v
    for c in 'OPQ':ws.column_dimensions[c].hidden=True
    table=Table(displayName='tblStokRealtime',ref='B7:N67');table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=False)
    ws.add_table(table)
    for status,bg,fg in [('MERAH','FCE4E4','9C0006'),('KUNING','FFF2CC','9C6500'),('HIJAU','E2EFDA','1B5E20'),('BELUM AWAL','FFF2CC','9C6500'),('ISI BATAS','FFF2CC','9C6500')]:
        ws.conditional_formatting.add('L8:L67',FormulaRule(formula=[f'$L8="{status}"'],fill=PatternFill('solid',fgColor=bg),font=Font(color=fg,bold=True)))
    ws.conditional_formatting.add('H8:H67',CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FCE4E4')))
    ws.freeze_panes='D8';ws.print_area='B1:N72';ws.print_title_rows='1:7'
    ws.page_setup.fitToHeight=0

def closing_history(dst):
    ws=dst['JURNAL'];top=FIRST_CLOSE-6
    ws.cell(top,1,'RIWAYAT TUTUP KASIR | SNAPSHOT PER REVISI')
    ws.cell(top,1).font=Font(name='Calibri',size=16,bold=True,color='902A20')
    go(ws,f'A{top+2}','Kembali ke Tutup Kasir',"'TUTUP KASIR'!B1")
    headers=['ID Tutup','Tanggal','Revisi','Kasir','Modal Awal','Penjualan Tunai','Beli Tunai','Kas Lain Bertanda','Kas Seharusnya','Uang Fisik','Selisih','Omzet Netto','Nota Valid','Item Terjual','Status HPP','Catatan','Dicatat Pada','Cek']
    for c,h in enumerate(headers,1):ws.cell(FIRST_CLOSE-1,c,h)
    table=Table(displayName='tblTutupKasir',ref=f'A{FIRST_CLOSE-1}:R{LAST_CLOSE}');table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True)
    ws.add_table(table)
    for r in range(FIRST_CLOSE,LAST_CLOSE+1):
        ws.cell(r,18,f'=IF(A{r}="","",IF(COUNTIF($A${FIRST_CLOSE}:$A${LAST_CLOSE},A{r})<>1,"ID GANDA",IF(COUNT(B{r},C{r},E{r}:N{r},Q{r})<>13,"CEK INPUT",IF(OR(B{r}<=0,B{r}>TODAY(),E{r}<0,F{r}<0,G{r}<0,J{r}<0,L{r}<0,M{r}<0,N{r}<0,C{r}<1,C{r}<>INT(C{r}),M{r}<>INT(M{r}),N{r}<>INT(N{r}),D{r}="",AND(O{r}<>"LENGKAP",O{r}<>"BELUM LENGKAP")),"CEK INPUT",IF(COUNTIFS($B${FIRST_CLOSE}:$B${LAST_CLOSE},B{r},$C${FIRST_CLOSE}:$C${LAST_CLOSE},C{r})<>1,"REVISI GANDA",IF(OR(ABS(I{r}-(E{r}+F{r}-G{r}+H{r}))>0.01,ABS(K{r}-(J{r}-I{r}))>0.01),"CEK HITUNG","OK"))))))')
        ws.cell(r,2).number_format='dd mmm yyyy';ws.cell(r,17).number_format='dd mmm yyyy hh:mm'
        for c in range(5,13):ws.cell(r,c).number_format=MONEY

def closing_screen(dst,source):
    ws=copy_screen(dst,source['Tutup Kasir'],'TUTUP KASIR')
    ws['B1']='TUTUP KASIR | MUFRITA CREPES'
    ws['B2']='Pilih tanggal, isi modal awal dan hitung uang fisik. Rekap otomatis dari jurnal aktif; simpan hasil penutupan sebagai arsip.'
    for a,t,target in [('B3','Kasir',"'KASIR'!A1"),('C3','Dashboard',"'RINGKASAN'!A1"),('E3','Stok Realtime',"'STOK'!B1"),('H3','Buku Transaksi',"'JURNAL'!A1")]:go(ws,a,t,target)
    ws['D3']=None;ws['D3'].fill=PatternFill(fill_type=None)
    for r,l in [(8,'Jumlah Nota'),(9,'Nota Pelanggan'),(10,'Item Terjual'),(11,'Penjualan Kotor'),(12,'Total Diskon'),(13,'OMZET BERSIH'),(14,'Rata-rata per Nota'),(15,'HPP Acuan'),(16,'KONTRIBUSI KOTOR'),(17,'Nota Dibatalkan')]:ws[f'B{r}']=l
    ws['B19']='Catatan / Alasan Selisih'
    set_formula(ws,'C5','TODAY()','dd mmm yyyy')
    ws['C5'].fill=PatternFill('solid',fgColor='E2EFDA')
    ws['B20'].fill=PatternFill('solid',fgColor='E2EFDA')
    nday=area('NOTA','B');nstatus=area('NOTA','H');ncheck=area('NOTA','O')
    valid=f'{nday},$C$5,{nstatus},"AKTIF",{ncheck},"OK"'
    net=area('NOTA','P')
    for cell,expr in {
        'C8':f'COUNTIFS({valid})','C9':f'COUNTIFS({valid},{area("NOTA","D")},"<>")',
        'C10':f'SUMIFS({area("ITEM","N")},{area("ITEM","I")},$C$5)',
        'C11':f'SUMIFS({area("NOTA","J")},{valid})','C12':f'SUMIFS({area("NOTA","F")},{valid})',
        'C13':f'SUMIFS({net},{nday},$C$5)','C14':'IFERROR(C13/C8,0)',
        'C15':f'IF(COUNTIFS({valid},{area("NOTA","Q")},"<>LENGKAP")>0,"BELUM LENGKAP",SUMIFS({area("NOTA","M")},{valid}))',
        'C16':'IF(ISNUMBER(C15),C13-C15,"HPP BELUM LENGKAP")',
        'C17':f'COUNTIFS({nday},$C$5,{nstatus},"VOID")',
    }.items():set_formula(ws,cell,expr,MONEY if cell not in ('C8','C9','C10','C17') else '0')
    ws['B18']='Nota Perlu Diperiksa'
    set_formula(ws,'C18',f'COUNTIFS({nday},$C$5,{nstatus},"<>VOID",{ncheck},"<>OK")','0')
    for r,m in [(8,'Tunai'),(9,'QRIS'),(10,'Transfer'),(11,'E-wallet')]:
        ws[f'E{r}']=m;set_formula(ws,f'F{r}',f'SUMIFS({net},{nday},$C$5,{area("NOTA","E")},"{m}")',MONEY)
    for cell,label,expr in [
        ('F12','Pembelian Tunai',f'SUMIFS({area("GERAK BAHAN","P")},{area("GERAK BAHAN","B")},$C$5,{area("GERAK BAHAN","H")},"Tunai")'),
        ('F13','Biaya Usaha Tunai',f'-SUMIFS({area("KEUANGAN","I")},{area("KEUANGAN","B")},$C$5,{area("KEUANGAN","E")},"Tunai",{area("KEUANGAN","C")},"BIAYA USAHA")'),
        ('F14','Prive Tunai',f'-SUMIFS({area("KEUANGAN","I")},{area("KEUANGAN","B")},$C$5,{area("KEUANGAN","E")},"Tunai",{area("KEUANGAN","C")},"PRIVE")'),
        ('F15','Modal / Lain Tunai',f'SUMIFS({area("KEUANGAN","I")},{area("KEUANGAN","B")},$C$5,{area("KEUANGAN","E")},"Tunai",{area("KEUANGAN","I")},">0")'),
    ]:ws['E'+cell[1:]]=label;set_formula(ws,cell,expr,MONEY)
    ws['E16']='Jam Ramai (Nota)';ws['E17']='Menu Terlaris'
    ws['I8']=None;ws['I11']=None
    for cell in ['I8','I11']:
        ws[cell].fill=PatternFill('solid',fgColor='E2EFDA');ws[cell].font=Font(name='Calibri',size=12,bold=True,color='1B5E20');ws[cell].number_format=MONEY
        dv=DataValidation(type='decimal',operator='greaterThanOrEqual',formula1=0,allow_blank=True);dv.errorTitle='Nominal tidak valid';dv.error='Isi angka nol atau lebih.';dv.showErrorMessage=True;dv.errorStyle='stop';ws.add_data_validation(dv);dv.add(cell)
    dv=DataValidation(type='date',operator='between',formula1=1,formula2='TODAY()',allow_blank=False);dv.showErrorMessage=True;dv.error='Pilih tanggal sampai hari ini.';ws.add_data_validation(dv);dv.add('C5')
    set_formula(ws,'I9','F8',MONEY)
    set_formula(ws,'I10','IF(AND(ISNUMBER(I8),I8>=0),I8+F8-F12-F13-F14+F15,"")',MONEY)
    set_formula(ws,'I12','IF(AND(ISNUMBER(I8),I8>=0,ISNUMBER(I11),I11>=0),I11-I10,"")',MONEY)
    set_formula(ws,'H14','IF(OR(NOT(ISNUMBER(C5)),C5<=0,C5>TODAY()),"CEK TANGGAL",IF(OR(NOT(ISNUMBER(I8)),I8<0),"ISI MODAL AWAL (0 BILA TIDAK ADA)",IF(OR(NOT(ISNUMBER(I11)),I11<0),"ISI UANG FISIK (0 BILA KOSONG)",IF(ABS(I12)<0.01,"KAS COCOK - TIDAK ADA SELISIH",IF(I12>0,"KAS LEBIH Rp "&TEXT(I12,"#,##0"),"KAS KURANG Rp "&TEXT(-I12,"#,##0"))))))')
    for cell in ['H14','C16']:ws[cell].alignment=Alignment(wrap_text=True,vertical='center')
    ws['H18']='Menu';ws['I18']='Qty';ws['J18']='Bruto'
    # Stable, unique ranking when two menus have identical sales volumes.
    for r in range(8,58):
        mr=r-1
        set_formula(ws,f'N{r}',f'IF(KATALOG!A{mr}="","",KATALOG!A{mr})')
        set_formula(ws,f'O{r}',f'IF(N{r}="","",KATALOG!B{mr})')
        set_formula(ws,f'P{r}',f'IF(N{r}="",0,SUMIFS({area("ITEM","N")},{area("ITEM","C")},N{r},{area("ITEM","I")},$C$5))')
        set_formula(ws,f'Q{r}',f'IF(P{r}>0,P{r}+(58-ROW())/100000,0)')
        set_formula(ws,f'R{r}',f'IF(N{r}="",0,SUMIFS({area("ITEM","K")},{area("ITEM","C")},N{r},{area("ITEM","I")},$C$5,{area("ITEM","N")},">0"))')
    for r in range(19,24):
        k=r-18
        idx=f'MATCH(LARGE($Q$8:$Q$57,{k}),$Q$8:$Q$57,0)'
        set_formula(ws,f'H{r}',f'IF(LARGE($Q$8:$Q$57,{k})=0,"",INDEX($O$8:$O$57,{idx}))')
        set_formula(ws,f'I{r}',f'IF(H{r}="","",INDEX($P$8:$P$57,{idx}))','0')
        set_formula(ws,f'J{r}',f'IF(H{r}="","",INDEX($R$8:$R$57,{idx}))',MONEY)
        ws[f'H{r}'].alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[r].height=36
    set_formula(ws,'F17','IF(H19="","-",H19)')
    ws['F17'].alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[17].height=34
    for r in range(2,26):
        ws.cell(r,21,r-2)
        set_formula(ws,f'V{r}',f'COUNTIFS({valid},{area("NOTA","C")},">="&U{r}/24,{area("NOTA","C")},"<"&(U{r}+1)/24)')
    set_formula(ws,'F16','IF(MAX(V2:V25)=0,"-",TEXT(INDEX(U2:U25,MATCH(MAX(V2:V25),V2:V25,0)),"00")&":00")')
    ws['B25']='Rekap Tersimpan';set_formula(ws,'C25',f'COUNTIF(JURNAL!$B${FIRST_CLOSE}:$B${LAST_CLOSE},C5)','0')
    ws['B26']='Nama Kasir';ws['C26']='Pak Syarif';ws['C26'].fill=PatternFill('solid',fgColor='E2EFDA')
    go(ws,'H26','Riwayat Penutupan',f"'JURNAL'!A{FIRST_CLOSE-6}")
    ws['B28']='Isi 0 bila modal atau uang fisik memang nol. Selisih tidak otomatis dicatat sebagai koreksi uang.';ws.merge_cells('B28:J28')
    ws['B29']='Simpan tutup kasir melalui tombol VBA. Penyimpanan ulang membuat revisi baru; transaksi dan periode tidak dikunci.';ws.merge_cells('B29:J30');ws['B29'].alignment=Alignment(wrap_text=True,vertical='center')
    for a in ['B28','B29']:ws[a].font=Font(name='Calibri',size=10,color='756B60')
    ws.row_dimensions[28].height=26
    for c in ['N','O','P','Q','R','S','T','U','V']:ws.column_dimensions[c].hidden=True
    ws.column_dimensions['A'].width=2.5;ws.column_dimensions['H'].width=32;ws.column_dimensions['I'].width=10;ws.column_dimensions['J'].width=16
    for cell in ['I10','I12']:
        ws.conditional_formatting.add(cell,CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FCE4E4'),font=Font(color='9C0006',bold=True)))
    ws.conditional_formatting.add('I12',CellIsRule(operator='equal',formula=['0'],fill=PatternFill('solid',fgColor='E2EFDA'),font=Font(color='1B5E20',bold=True)))
    ws.freeze_panes='B8';ws.print_area='B1:J30'

def apply_pos_ui(dst):
    style_source_palette(dst)
    source=load_workbook(ROOT/'data/pos_reference.xlsm',data_only=False)
    stock_screen(dst,source);closing_history(dst);closing_screen(dst,source)
    source.close()
