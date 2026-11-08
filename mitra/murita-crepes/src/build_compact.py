"""Compact Crepes business model with Pak Seger POS as the UI foundation."""
import copy, json, re, tempfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from build_workbook import build
from compact_layout import BLOCKS, ORDER, formula, reference, shift_range
from pos_ui import apply_pos_ui

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT/'output/ringkas'
DEST = OUT/'POS_Mufrita_Crepes_Ringkas.xlsx'

def rewrite_chart(obj, context, seen=None):
    seen=set() if seen is None else seen
    if id(obj) in seen:return
    seen.add(id(obj))
    if hasattr(obj, '__dict__'):
        for k,v in vars(obj).items():
            if k == 'f' and isinstance(v,str): setattr(obj,k,reference(v,context))
            elif isinstance(v,list):
                for x in v: rewrite_chart(x,context,seen)
            elif hasattr(v,'__dict__'): rewrite_chart(v,context,seen)

def link(ws, address, label, target):
    c=ws[address];c.value=label;c.hyperlink='#'+target
    c.font=Font(name='Calibri',size=10,color='A05423',bold=True,underline='single')
    c.alignment=Alignment(vertical='center')

def finish_layout(dst):
    """Keep operator instructions and the matching macro filename in sync."""
    ws=dst['RINGKASAN']
    bodies=[
        'Periksa 34 menu dan master bahan pada KATALOG. STOK menampilkan saldo realtime seperti POS Pak Seger. Isi harga SKM yang belum diketahui pada master bahan.',
        'HARIAN > Belanja / stok: pilih AWAL dan isi kuantitas hasil hitung fisik per bahan dalam satuan dasar. Jurnal gerak bahan tersedia melalui pintasan JURNAL.',
        'HARIAN > Produksi: buat PREMIKS dari tujuh bahan, lalu ADONAN BASAH dari premiks, mentega dan air. Penjualan hanya memakai adonan basah dan pelengkap/topping.',
        'KASIR: pilih SKU, qty, pelanggan opsional, metode, diskon dan uang diterima. Modul VBA ringkas menyimpan nota, item, pemakaian dan poin sebagai nilai tetap.',
        'HARIAN > Kas / prive: pisahkan biaya usaha, modal, prive dan pendapatan lain. BELI bahan dicatat sekali pada bagian barang; tidak menjadi biaya usaha kedua.',
        'TUTUP KASIR: pilih tanggal, isi modal awal dan uang fisik, periksa selisih, lalu simpan rekap. RINGKASAN menampilkan KPI; PELANGGAN menghitung kunjungan per nota.'
    ]
    for i,body in enumerate(bodies):ws.cell(88+i*3,4,body)
    ws['A112']='TOMBOL: simpan salinan .xlsm, Alt+F11 > File > Import > Modul_Mufrita_Ringkas.bas, Debug > Compile VBAProject, lalu Alt+F8 > PasangTombol. Tanpa VBA, isi tabel JURNAL bersama sesuai panduan; form tidak menyimpan otomatis.'
    ws['A116']='Lengkapi harga SKM, stok fisik, kalibrasi topping dan takaran pelengkap sebelum operasional. Jam 17.00-23.00, Senin libur rutin. Parameter poin tersedia pada KATALOG > Parameter; awalnya nonaktif. Data historis sumber tidak dibuat menjadi nota baru.'
    dst['KASIR']['A2']='Penjualan 34 menu. Pelanggan opsional. Operasional sumber 17.00-23.00, Senin libur rutin. Simpan otomatis melalui modul VBA ringkas.'
    dst['KASIR'].column_dimensions['B'].width=20
    dst['KASIR'].column_dimensions['C'].width=38
    for col in 'NOPQ':dst['KATALOG'].column_dimensions[col].hidden=True
    # Receipt selection helpers share columns with the cart, so hide values
    # only in the receipt's helper cells, without hiding cart columns.
    for row in range(59,71):dst['KASIR'].cell(row,7).number_format=';;;'
    cell=dst['RINGKASAN']['I13']
    cell.alignment=Alignment(wrap_text=True,vertical='center')

def build_vba():
    source=(ROOT/'vba/Modul_MuritaCrepes.bas').read_text()
    # A Range supports relative .Cells/.Range; every logical block keeps its
    # original coordinates. Direct worksheet actions use .Worksheet explicitly.
    cases='\n'.join(f'        Case "{old}": Set S = ThisWorkbook.Worksheets("{new}").Range("A{off+1}")' for old,(new,off) in BLOCKS.items())
    function='Private Function S(ByVal title As String) As Range\n    Select Case title\n'+cases+'\n        Case Else: Err.Raise vbObjectError + 604, , "Blok tidak dikenal: " & title\n    End Select\nEnd Function'
    source=re.sub(r'Private Function S\(.*?End Function',function,source,count=1,flags=re.S)
    for var in ['k','p','f','t']:
        source=source.replace(f'Dim {var} As Worksheet',f'Dim {var} As Range')
    # The loop over all real worksheets in PasangTombol still uses Worksheet.
    source=source.replace('Dim t As Range, i As Long\n    On Error GoTo Failed\n    For Each t In ThisWorkbook.Worksheets','Dim t As Worksheet, i As Long\n    On Error GoTo Failed\n    For Each t In ThisWorkbook.Worksheets')
    source=source.replace('(ActiveSheet Is t)', '(ActiveSheet Is t.Worksheet)')
    source=source.replace('r = ActiveCell.Row:', 'r = ActiveCell.Row - t.Row + 1:')
    source=source.replace('S(title).Buttons.Add', 'S(title).Worksheet.Buttons.Add')
    source=source.replace('t.ExportAsFixedFormat Type:=xlTypePDF', 't.Range("A7:F35").ExportAsFixedFormat Type:=xlTypePDF')
    source=source.replace('Pilih baris pada sheet NOTA.', 'Pilih baris nota di JURNAL bagian NOTA.')
    source=source.replace('pilih baris NOTA lalu', 'pilih baris nota pada JURNAL lalu')
    # The only formula assembled by VBA refers to the old MENU worksheet.
    source=source.replace('MENU!$C$7:$C$56','KATALOG!$C$7:$C$56').replace('MENU!$A$7:$A$56','KATALOG!$A$7:$A$56')
    source=source.replace('Murita Crepes','Mufrita Crepes')
    source=source.replace('Array("POIN", "H", 1006))','Array("POIN", "H", 1006), Array("ARSIP TUTUP", "R", 372))')
    source=source.replace('    MsgBox "Tujuh tombol terpasang.', '    AddButton "TUTUP KASIR", "H5", "SIMPAN TUTUP KASIR", "SimpanTutupKasir"\n    MsgBox "Delapan tombol terpasang.')
    source+='\n'+(ROOT/'vba/Tutup_Mufrita.inc').read_text()
    assert source.isascii()
    (OUT/'Modul_Mufrita_Ringkas.bas').write_bytes(source.replace('\n','\r\n').encode('ascii'))

def main():
    OUT.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='mufrita-build-') as td:
        src=load_workbook(build(Path(td)/'model.xlsx'))
        dst=Workbook();dst.remove(dst.active)
        for name in ORDER: dst.create_sheet(name)
        for old in src:
            name,offset=BLOCKS[old.title];new=dst[name]
            for row in old:
                for c in row:
                    if c.value is None and not c.has_style: continue
                    d=new.cell(c.row+offset,c.column)
                    d.value=formula(c.value,old.title) if c.data_type=='f' else c.value
                    if c.data_type=='f' and old.title=='STRUK':
                        # ROW returns physical row numbers after consolidation.
                        item_offset=BLOCKS['ITEM'][1]
                        d.value=d.value.replace(
                            "ROW('JURNAL'!$B$1021:$B$4020)-6",
                            f"ROW('JURNAL'!$B$1021:$B$4020)-{6+item_offset}")
                    if isinstance(d.value,str) and c.data_type!='f':
                        d.value=d.value.replace('MURITA CREPES','MUFRITA CREPES').replace('Murita Crepes','Mufrita Crepes').replace('Modul_MuritaCrepes.bas','Modul_Mufrita_Ringkas.bas')
                    if c.has_style:
                        for attr in ['font','fill','border','alignment','protection']:
                            setattr(d,attr,copy.copy(getattr(c,attr)))
                        d.number_format=c.number_format
                    if c.comment:d.comment=copy.copy(c.comment)
                    if c.hyperlink:
                        d._hyperlink=copy.copy(c.hyperlink);d._hyperlink.ref=d.coordinate
                        if d.hyperlink.location: d.hyperlink.location=reference(d.hyperlink.location,old.title)
                        if d.hyperlink.target and d.hyperlink.target.startswith('#'):
                            d.hyperlink.target='#'+reference(d.hyperlink.target[1:],old.title)
            for rng in old.merged_cells.ranges:new.merge_cells(shift_range(str(rng),offset))
            for row,dim in old.row_dimensions.items():
                nd=copy.copy(dim);nd.index=row+offset;nd.parent=new;new.row_dimensions[row+offset]=nd
            for dv in old.data_validations.dataValidation:
                nd=copy.copy(dv);nd.sqref=' '.join(shift_range(str(r),offset) for r in dv.sqref.ranges)
                nd.formula1=formula(dv.formula1,old.title);nd.formula2=formula(dv.formula2,old.title)
                new.add_data_validation(nd)
            for cf,rules in old.conditional_formatting._cf_rules.items():
                dest=' '.join(shift_range(str(r),offset) for r in cf.sqref.ranges)
                for rule in rules:
                    nr=copy.deepcopy(rule);nr.formula=[formula(v,old.title,local=True) for v in rule.formula]
                    new.conditional_formatting.add(dest,nr)
            for table in old.tables.values():
                nt=copy.deepcopy(table);nt.ref=shift_range(table.ref,offset)
                if nt.autoFilter:nt.autoFilter.ref=nt.ref
                if nt.sortState:nt.sortState.ref=nt.ref
                new.add_table(nt)
            for ch in old._charts:
                nc=copy.deepcopy(ch);rewrite_chart(nc,old.title)
                nc.anchor._from.row+=offset;nc.anchor.to.row+=offset
                new.add_chart(nc)
        for n in src.defined_names.values():
            dn=copy.copy(n);dn.attr_text=formula(n.attr_text);dst.defined_names.add(dn)
        dst.properties=copy.copy(src.properties)
        dst.properties.title='Mufrita Crepes | POS Ringkas'
        dst.properties.description='8 tab kerja + 1 sheet rumus. Kerangka UI POS Pak Seger; model Murita Crepes, Pak Syarif, MoM 24 Oktober 2026.'
        for ws in dst:
            ws.sheet_view.showGridLines=False;ws.sheet_view.zoomScale=75
            ws.sheet_format.defaultRowHeight=22
            ws.column_dimensions['A'].width=22
            ws.column_dimensions['B'].width=34
            for c in 'CDEFGHIJKLMNOPQR':ws.column_dimensions[c].width=18
            ws.freeze_panes='C7';ws.sheet_properties.pageSetUpPr.fitToPage=True
            ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
            ws.sheet_properties.tabColor='A05423';ws.print_options.horizontalCentered=True
        # Replace repeated old worksheet menus with compact block navigation.
        for old,(name,offset) in BLOCKS.items():
            ws=dst[name]
            for i,target in enumerate(ORDER[:6]):link(ws,f'{chr(65+i*2)}{4+offset}',target.title(),f"'{target}'!A1")
            link(ws,f'M{4+offset}','Jurnal',"'JURNAL'!A1")
        nav={
          'RINGKASAN': [('A3','Panduan mulai',"'RINGKASAN'!A81"),('E3','Katalog & parameter',"'KATALOG'!A167"),('I3','Tutup Kasir',"'TUTUP KASIR'!B1")],
          'KASIR': [('A3','Input nota',"'KASIR'!A6"),('D3','Struk / PDF',"'KASIR'!A47"),('H3','Tutup kasir',"'TUTUP KASIR'!B1")],
          'HARIAN': [('A3','Produksi',"'HARIAN'!A1"),('D3','Belanja / stok',"'HARIAN'!A43"),('H3','Kas / prive',"'HARIAN'!A57"),('K3','Tutup hari',"'TUTUP KASIR'!B1")],
          'KATALOG': [('A3','Menu',"'KATALOG'!A1"),('D3','Topping',"'KATALOG'!A63"),('G3','Premiks',"'KATALOG'!A101"),('I3','Adonan',"'KATALOG'!A135"),('K3','Parameter',"'KATALOG'!A167")],
          'JURNAL': [('A3','Nota',"'JURNAL'!A1"),('D3','Item',"'JURNAL'!A1015"),('G3','Gerak bahan',"'JURNAL'!A4029"),('J3','Keuangan',"'JURNAL'!A10043"),('M3','Riwayat Tutup',"'JURNAL'!A12071")],
        }
        for name,links in nav.items():
            for addr,label,target in links:link(dst[name],addr,label,target)
        # Meaningful section names, no need to know source worksheet titles.
        for old,title in [('DASHBOARD','MUFRITA CREPES | RINGKASAN'),('MULAI','PANDUAN MULAI'),('PRODUKSI','HARIAN | PRODUKSI'),('FORM','HARIAN | BELANJA, KAS & POIN'),('BAHAN','STOK | BAHAN & SALDO'),('MENU','KATALOG | MENU'),('REFERENSI','PARAMETER & SUMBER')]:
            name,off=BLOCKS[old];dst[name].cell(off+1,1,title)
        note=dst['RINGKASAN']['A127'];note.value='Nama tampilan mengikuti permintaan: Mufrita Crepes. Pada sumber MoM dan file biaya tertulis Murita Crepes, Pak Syarif.'
        note.font=Font(name='Calibri',size=10,color='756B60');note.alignment=Alignment(wrap_text=True)
        dst['RINGKASAN'].merge_cells('A127:L128')
        for ws in dst:
            ws.print_area=f'A1:L{min(ws.max_row,74)}'
        # Print only the actual receipt; Range.ExportAsFixedFormat uses the same bounds.
        dst['KASIR'].print_area='A53:F81';dst['KASIR'].print_title_rows=None
        dst['KASIR'].page_setup.orientation='portrait';dst['KASIR'].page_setup.fitToHeight=1
        dst['STOK'].print_area='A1:M66';dst['PELANGGAN'].print_area='A1:M36'
        dst['_RESEP'].sheet_state='hidden'
        finish_layout(dst)
        apply_pos_ui(dst)
        link(dst['KATALOG'],'A223','MASTER BAHAN | harga, satuan & batas belanja',"'KATALOG'!A225")
        dst._sheets=[dst[n] for n in ORDER]
        dst.active=0;dst.calculation.fullCalcOnLoad=True
        dst.save(DEST)
    build_vba()
    (OUT/'peta_sheet.json').write_text(json.dumps(BLOCKS,ensure_ascii=False,indent=2))
    print(DEST)

if __name__=='__main__':main()
