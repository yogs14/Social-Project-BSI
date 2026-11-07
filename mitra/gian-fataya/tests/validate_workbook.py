"""Functional tests recalculate actual workbook formulas using LibreOffice.

All seeded data is synthetic and stays in /tmp. The production file stays blank.
This does not execute VBA or establish Excel desktop macro compatibility.
"""
from __future__ import annotations
import datetime as dt
import json
import subprocess
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import sys

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'src'))
from workbook_links import repair_navigation, validate_navigation
SOURCE=ROOT/'output/POS_Gian_Fataya.xlsx'
checks=[]

# Independent physical-coordinate expectations catch section relocation errors.
OFFSETS = {
    "PRODUK": ("DATA USAHA",96), "WILAYAH": ("DATA USAHA",0),
    "OUTLET": ("DATA USAHA",28), "LOT": ("DATA USAHA",154),
    "BAHAN": ("DATA USAHA",242), "REFERENSI": ("DATA USAHA",262),
    "DASHBOARD": ("BERANDA",0), "FORM": ("BERANDA",70),
    "MUTASI": ("TRANSAKSI",0), "INVOICE": ("TAGIHAN",0),
    "PEMBAYARAN": ("TAGIHAN",430),
}
COLS = {8:13,9:15,10:16,11:17,12:18,13:19,14:20,15:21,16:22,17:23,18:10,19:24,20:12,21:8,22:14,23:9,24:11}


class SectionView:
    def __init__(self, book, name):
        self.logical = name
        title, self.offset = OFFSETS.get(name,(name,0))
        self.ws = book[title]

    def cell(self, row, column, value=None):
        if self.logical == "MUTASI":
            column = COLS.get(column,column)
        return self.ws.cell(row+self.offset,column,value)

    def __getitem__(self, ref):
        match = re.fullmatch(r"([A-Z]+)([0-9]+)",ref)
        return self.cell(int(match[2]),column_index_from_string(match[1]))

    def __setitem__(self, ref, value):
        self[ref].value=value

    def __getattr__(self, name):
        return getattr(self.ws,name)


class BookView:
    def __init__(self, book):
        self.book=book

    def __getitem__(self,name):
        return SectionView(self.book,name)

    def __iter__(self):
        return iter(self.book)

    def __getattr__(self,name):
        return getattr(self.book,name)


def day(s):return dt.datetime.fromisoformat(s)


def equal(name,actual,expected):
    assert actual==expected,f'{name}: {actual!r} != {expected!r}'
    checks.append({'check':name,'result':'PASS','actual':actual})


def base():
    w=BookView(load_workbook(SOURCE))
    w['PRODUK']['D7']=10000
    for r,code,area in [(8,'UJI-JT','W01'),(9,'UJI-JG','W02')]:
        for c,v in enumerate([code,'SIMULASI - BUKAN OUTLET ASLI',area,'PIC UJI','ALAMAT UJI','','AKTIF'],1):w['OUTLET'].cell(r,c,v)
        w['OUTLET'].cell(r,9,"Amanda" if r==8 else "Independen")
    for r,lot,expiry in [(7,'UJI-LOT','2027-01-31'),(8,'UJI-EXP','2026-10-05')]:
        for c,v in [(1,lot),(2,'GF-PST-001'),(4,day(expiry)),(5,'DATA UJI SINTETIS')]:w['LOT'].cell(r,c,v)
    w['WILAYAH']['C7']='PIC SIMULASI';w['WILAYAH']['D7']='ALAMAT SIMULASI'
    w['DASHBOARD']['B5']=day('2026-10-01');w['DASHBOARD']['E5']=day('2026-11-01')
    moves=[
      ('M1','2026-10-01','AWAL','GDG','UJI-LOT',200,None,None),
      ('M2','2026-10-01','KIRIM','UJI-JT','UJI-LOT',40,None,None),
      ('M3','2026-10-10','JUAL','UJI-JT','UJI-LOT',30,10000,'INV-UJI'),
      ('M4','2026-11-01','KIRIM','UJI-JT','UJI-LOT',40,None,None),
      ('M5','2026-10-01','KIRIM','UJI-JG','UJI-LOT',20,None,None),
      ('M6','2026-10-10','JUAL','UJI-JG','UJI-LOT',5,10000,None),
    ]
    for r,vals in enumerate(moves,7):
        id,d,t,loc,lot,qty,price,invoice=vals
        for c,v in enumerate([id,day(d),t,loc,lot,qty,price,invoice,'AKTIF','SIMULASI'],1):
            if v is not None:w['MUTASI'].cell(r,c,v)
        if t=='KIRIM':w['MUTASI'].cell(r,21,'PO-UJI-'+id)
    vals=['INV-UJI',day('2026-10-31'),'W01',day('2026-10-01'),day('2026-10-31'),day('2026-11-15'),.1,20000,'TERBIT',300000]
    for c,v in enumerate(vals,1):w['INVOICE'].cell(7,c,v)
    for r,id,d,amount in [(7,'PAY1','2026-10-31',100000),(8,'PAY2','2026-11-01',190000)]:
        for c,v in enumerate([id,day(d),'INV-UJI','Transfer',amount,'AKTIF','SIMULASI'],1):w['PEMBAYARAN'].cell(r,c,v)
    w['CETAK INVOICE']['B5']='INV-UJI'
    return w


def stock(w,location,lot):
    for row in w['STOK'].iter_rows(min_row=7,values_only=True):
        if row[0]==location and row[3]==lot:return row[9]
    raise AssertionError(f'Stock row not found: {location}/{lot}')


def no_errors(name,w):
    errors=[]
    for s in w:
        for row in s:
            for c in row:
                if c.data_type=='e':errors.append(f'{s.title}!{c.coordinate}={c.value}')
    equal(name+' / no formula errors',errors,[])


def main():
    equal('native internal navigation before recalculation',validate_navigation(SOURCE)['internal'],62)
    original=BookView(load_workbook(SOURCE,data_only=False))
    equal('invoice print area belongs to invoice',str(original['CETAK INVOICE'].print_area),"'CETAK INVOICE'!$A$7:$F$79")
    equal('dashboard print area belongs to dashboard',str(original['DASHBOARD'].print_area),"'BERANDA'!$A$1:$L$64")
    equal('movement table includes PO and freight',original['MUTASI'].tables['tblMutasi'].ref,'A6:X2006')
    equal('VBA not falsely embedded in xlsx',original.vba_archive,None)
    equal('exactly six physical sheets',original.sheetnames,['BERANDA','TRANSAKSI','STOK','TAGIHAN','CETAK INVOICE','DATA USAHA'])
    equal('all six sheets are visible',all(s.sheet_state=='visible' for s in original),True)
    equal('invoice table retained',original['INVOICE'].tables['tblInvoice'].ref,'A6:R406')
    equal('payments moved into same sheet',original['PEMBAYARAN'].tables['tblPembayaran'].ref,'A436:J1436')
    equal('outlet group dropdown in master section',any('I36:I95' in str(v.sqref) and v.formula1=='"Amanda,Independen"' for v in original.book['DATA USAHA'].data_validations.dataValidation),True)
    equal('payment selector moved with form',any('E107' in str(v.sqref) and v.formula1=='DaftarInvoice' for v in original.book['BERANDA'].data_validations.dataValidation),True)
    equal('named invoice range points to new sheet',original.defined_names['DaftarInvoice'].attr_text,"'TAGIHAN'!$A$7:$A$406")
    equal('named lot range points to new section',original.defined_names['DaftarLot'].attr_text,"'DATA USAHA'!$A$161:$A$240")
    equal('master data remain actual source regions',[original['WILAYAH'].cell(r,2).value for r in range(7,15)],['Jawa Timur','Jawa Tengah','Medan - Aceh','Sulawesi Selatan','Kalimantan','Sumatera Selatan','Batam','Jabodetabek'])
    equal('advanced transaction fields grouped',original.book['TRANSAKSI'].column_dimensions['M'].hidden,True)
    equal('PO control has visible header',original.book['TRANSAKSI']['K6'].value,'Cek minimum PO')
    equal('input dates are formatted even while empty',original.book['TRANSAKSI']['B7'].number_format,'dd mmm yyyy')
    equal('invoice discount uses percent format',original.book['TAGIHAN']['G7'].number_format,'0.0%')
    equal('all table headers are populated',all(s.cell(int(re.search(r'\d+',t.ref)[0]),c).value is not None for s in original for t in s.tables.values() for c in range(1,len(t.tableColumns)+1)),True)
    with tempfile.TemporaryDirectory(prefix='gian-fataya-tests-') as td:
        td=Path(td);out=td/'recalculated';out.mkdir()
        files=[]
        for name,modify in [
            ('normal',lambda w:None),
            ('historical',lambda w:w['DASHBOARD'].__setitem__('E5',day('2026-10-31'))),
            ('partial',lambda w:w['PEMBAYARAN'].__setitem__('F8','VOID')),
            ('returns_damage',lambda w:None),
            ('invalid',lambda w:None),
            ('changed_invoice',lambda w:w['MUTASI'].__setitem__('G9',11000)),
            ('draft',lambda w:w['INVOICE'].__setitem__('I7','DRAF')),
            ('retired_outlet',lambda w:w['OUTLET'].__setitem__('G8','NONAKTIF')),
            ('void',lambda w:None),
            ('small_po',lambda w:None),
            ('split_po',lambda w:None),
            ('retur_without_reason',lambda w:None),
            ('custom_threshold',lambda w:w['REFERENSI'].__setitem__('B14',.1)),
            ('chart_selected',lambda w:w['DASHBOARD'].__setitem__('I18','Jawa Tengah')),
            ('chart_period',lambda w:w['DASHBOARD'].__setitem__('B5',day('2026-10-11'))),
            ('chart_cutoff',lambda w:w['DASHBOARD'].__setitem__('E5',day('2026-10-09'))),
            ('chart_void',lambda w:w['MUTASI'].__setitem__('I9','VOID')),
        ]:
            w=base();modify(w)
            if name=='returns_damage':
                for r,vals in [(13,['M7',day('2026-10-20'),'RETUR','UJI-JT','UJI-LOT',5,None,None,'AKTIF','SIMULASI']),
                               (14,['M8',day('2026-10-21'),'RUSAK','UJI-JT','UJI-LOT',2,None,None,'AKTIF','SIMULASI'])]:
                    for c,v in enumerate(vals,1):
                        if v is not None:w['MUTASI'].cell(r,c,v)
            if name=='invalid':
                cases=[
                  ['BAD1',day('2026-10-01'),'AWAL','UJI-JT','UJI-EXP',10,None,None,'AKTIF'],
                  ['BAD2',day('2026-10-10'),'JUAL','UJI-JT','UJI-EXP',2,10000,None,'AKTIF'],
                  ['BAD3',day('2026-10-10'),'JUAL','UJI-JT','UJI-LOT',1,None,None,'AKTIF'],
                  ['DUP',day('2026-10-10'),'AWAL','GDG','UJI-LOT',100,None,None,'AKTIF'],
                  ['DUP',day('2026-10-10'),'AWAL','GDG','UJI-LOT',100,None,None,'AKTIF'],
                  ['BAD4',day('2026-10-10'),'JUAL','UJI-JT','UJI-LOT',9999,10000,None,'AKTIF'],
                  ['BAD5',day('2026-10-10'),'KIRIM','UJI-JT','UJI-LOT',1,None,None,'AKTIF'],
                  ['BAD6',day('2026-10-10'),'AWAL','GDG','UJI-LOT',-1,None,None,'AKTIF'],
                ]
                for r,vals in enumerate(cases,13):
                    for c,v in enumerate(vals,1):
                        if v is not None:w['MUTASI'].cell(r,c,v)
            if name=='void':
                w['INVOICE']['I7']='VOID';w['MUTASI']['H9']=None
                w['PEMBAYARAN']['F7']='VOID';w['PEMBAYARAN']['F8']='VOID'
            if name in ['small_po','split_po']:
                w['MUTASI']['F8']=6
                w['MUTASI']['F9']=3
                w['INVOICE']['J7']=30000
                w['PEMBAYARAN']['E7']=10000
                w['PEMBAYARAN']['E8']=19000
            if name=='split_po':
                for c,v in [(1,'UJI-SPLIT'),(2,'GF-PST-001'),(4,day('2027-01-31'))]:w['LOT'].cell(9,c,v)
                for r,kind,loc,qty in [(13,'AWAL','GDG',100),(14,'KIRIM','UJI-JT',4)]:
                    vals=['SPLIT'+str(r),day('2026-10-01'),kind,loc,'UJI-SPLIT',qty,None,None,'AKTIF','SIMULASI']
                    for c,v in enumerate(vals,1):
                        if v is not None:w['MUTASI'].cell(r,c,v)
                w['MUTASI']['U14']='PO-UJI-M2'
            if name=='retur_without_reason':
                for c,v in enumerate(['RET-UJI',day('2026-10-20'),'RETUR','UJI-JT','UJI-LOT',1,None,None,'AKTIF'],1):
                    if v is not None:w['MUTASI'].cell(13,c,v)
            dest=td/(name+'.xlsx');w.save(dest);files.append(dest)
        # Test the real blank workbook too, without editing or adding fixtures.
        files.append(SOURCE)
        command=['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in files]]
        result=subprocess.run(command,capture_output=True,text=True,timeout=240)
        if result.returncode:raise RuntimeError(result.stdout+'\n'+result.stderr)
        equal('Recalculated expected files',len(list(out.glob('*.xlsx'))),len(files))
        loaded={p.stem:BookView(load_workbook(out/p.name,data_only=True)) for p in files}
        for name,w in loaded.items():no_errors(name,w)
        charts=loaded['normal'].book['BERANDA']._charts
        equal('comparison and area trend charts retained',len(charts),2)
        chart,trend=charts
        equal('comparison retains high category source',chart.series[0].val.numRef.f,'BERANDA!$O$21:$O$28')
        equal('comparison contains all eight official areas',chart.series[0].val.numRef.numCache.ptCount,8)
        equal('comparison retains actual high sales',chart.series[0].val.numRef.numCache.pt[0].v,300000)
        equal('comparison retains actual medium sales',chart.series[1].val.numRef.numCache.pt[1].v,50000)
        equal('trend retains concrete monthly source',trend.series[0].val.numRef.f,'BERANDA!$O$6:$O$17')
        equal('trend contains twelve months',trend.series[0].val.numRef.numCache.ptCount,12)
        equal('trend chart retains actual October sales',trend.series[0].val.numRef.numCache.pt[10].v,300000)
        equal('charts include hidden source columns',all(c.visible_cells_only is False for c in charts),True)
        equal('comparison appears above details',chart.anchor._from.row < 35,True)
        equal('trend appears above details',trend.anchor._from.row < 35,True)
        w=loaded['normal']
        equal('comparison labels include region and category',w['DASHBOARD']['N21'].value,'Jawa Timur · TINGGI')
        equal('zero sales category visible in chart label',w['DASHBOARD']['N23'].value,'Medan - Aceh · DATAR (0)')
        equal('trend starts eleven months before cutoff',w['DASHBOARD']['N6'].value.isoformat(),day('2025-12-01').isoformat())
        equal('month without sales displays zero',w['DASHBOARD']['O17'].value,0)
        equal('selector changes trend region',loaded['chart_selected']['DASHBOARD']['O2'].value,'W02')
        equal('selected region trend includes independent outlets',loaded['chart_selected']['DASHBOARD']['O16'].value,50000)
        equal('comparison respects period start',loaded['chart_period']['DASHBOARD']['D37'].value,0)
        equal('trend deliberately retains twelve-month history',loaded['chart_period']['DASHBOARD']['O16'].value,300000)
        equal('trend respects day cutoff in last month',loaded['chart_cutoff']['DASHBOARD']['O17'].value,0)
        equal('void sale excluded from comparison',loaded['chart_void']['DASHBOARD']['D37'].value,0)
        equal('void sale excluded from trend',loaded['chart_void']['DASHBOARD']['O16'].value,0)
        equal('threshold moves medium bar into high series',loaded['custom_threshold']['DASHBOARD']['O22'].value,50000)
        equal('threshold removes old medium bar',loaded['custom_threshold']['DASHBOARD']['P22'].value,0)
        w=loaded['normal']
        for r in range(7,13):equal(f'movement {r} valid',w['MUTASI'][f'R{r}'].value,'OK')
        equal('carryover: 40 - 30 + 40 = 50',stock(w,'UJI-JT','UJI-LOT'),50)
        equal('central stock after three shipments',stock(w,'GDG','UJI-LOT'),100)
        equal('other region stock separate',stock(w,'UJI-JG','UJI-LOT'),15)
        equal('shipment not counted as revenue',w['DASHBOARD']['A8'].value,350000)
        equal('invoice only linked region sales',w['INVOICE']['Q7'].value,300000)
        equal('invoice percentage discount',w['INVOICE']['K7'].value,30000)
        equal('freight after discount',w['INVOICE']['L7'].value,290000)
        equal('full payment',w['INVOICE']['M7'].value,290000)
        equal('paid invoice balance',w['INVOICE']['N7'].value,0)
        equal('paid invoice status',w['INVOICE']['O7'].value,'LUNAS')
        equal('invoice source reconciles',w['INVOICE']['P7'].value,'OK')
        equal('invoice print recipient ready',w['CETAK INVOICE']['A14'].value,'SIAP CETAK')
        equal('invoice line quantity',w['CETAK INVOICE']['D17'].value,30)
        equal('invoice line gross',w['CETAK INVOICE']['F17'].value,300000)
        equal('printed invoice number present',w['CETAK INVOICE']['B9'].value,'INV-UJI')
        equal('unbilled other region',w['DASHBOARD']['K62'].value,50000)
        equal('top area classified high',w['DASHBOARD']['H37'].value,'TINGGI')
        equal('smaller positive area classified medium',w['DASHBOARD']['H38'].value,'SEDANG')
        equal('area without sales classified flat',w['DASHBOARD']['H39'].value,'DATAR')
        equal('independent outlets retained in area sales',w['DASHBOARD']['D38'].value,50000)
        w=loaded['historical']
        equal('historical outlet stock excludes next month shipment',stock(w,'UJI-JT','UJI-LOT'),10)
        equal('historical central stock',stock(w,'GDG','UJI-LOT'),140)
        equal('historical cash excludes future settlement',w['DASHBOARD']['I8'].value,100000)
        equal('historical receivable excludes future settlement',w['DASHBOARD']['A13'].value,190000)
        w=loaded['partial']
        equal('partial payment balance',w['INVOICE']['N7'].value,190000)
        equal('partial payment status',w['INVOICE']['O7'].value,'SEBAGIAN')
        w=loaded['returns_damage']
        equal('return and spoilage reduce outlet stock',stock(w,'UJI-JT','UJI-LOT'),43)
        equal('return increases warehouse stock',stock(w,'GDG','UJI-LOT'),105)
        equal('return and spoilage do not become revenue',w['DASHBOARD']['A8'].value,350000)
        w=loaded['invalid']
        for row,expected in [(14,'SUDAH KEDALUWARSA'),(15,'ISI HARGA JUAL'),(16,'ID GANDA'),(17,'ID GANDA'),(19,'ISI NO PO/SJ'),(20,'CEK QTY')]:
            equal(f'invalid movement {row} detected',w['MUTASI'][f'R{row}'].value,expected)
            equal(f'invalid movement {row} excluded',w['MUTASI'][f'S{row}'].value,0)
        equal('expired stock retained until disposal',stock(w,'UJI-JT','UJI-EXP'),10)
        equal('negative historical balance detected',w['MUTASI']['T18'].value<0,True)
        w=loaded['changed_invoice']
        equal('post-issue source edit detected',w['INVOICE']['P7'].value,'SELISIH SUMBER')
        equal('snapshot remains fixed after price edit',w['INVOICE']['J7'].value,300000)
        equal('invalid invoice print stopped',w['CETAK INVOICE']['A14'].value,'STOP: PERIKSA INVOICE DAN SUMBER')
        w=loaded['draft']
        equal('draft invoice payments excluded',w['INVOICE']['M7'].value,0)
        equal('draft invoice excluded from issued totals',w['DASHBOARD']['E8'].value,0)
        equal('draft invoice not printable',w['CETAK INVOICE']['A14'].value,'BUKAN INVOICE TERBIT')
        w=loaded['retired_outlet']
        equal('retiring outlet preserves historical sales',w['DASHBOARD']['A8'].value,350000)
        equal('retiring outlet preserves its stock',stock(w,'UJI-JT','UJI-LOT'),50)
        w=loaded['void']
        equal('void invoice has no receivable',w['INVOICE']['N7'].value,0)
        equal('void invoice sales available for rebilling',w['DASHBOARD']['K62'].value,350000)
        w=loaded['small_po']
        equal('PO below ten has explicit warning',w['MUTASI']['X8'].value,'PO < 10 PCS')
        equal('small PO remains physical stock, not silently dropped',w['MUTASI']['S8'].value,6)
        equal('dashboard surfaces incomplete PO',w['DASHBOARD']['I57'].value,1)
        w=loaded['split_po']
        equal('same PO first lot accepted at six',w['MUTASI']['X8'].value,'OK')
        equal('same PO second lot completes ten',w['MUTASI']['X14'].value,'OK')
        equal('split PO leaves no incomplete warning',w['DASHBOARD']['I57'].value,0)
        w=loaded['retur_without_reason']
        equal('return needs damage explanation',w['MUTASI']['R13'].value,'ISI ALASAN')
        equal('return without reason excluded',w['MUTASI']['S13'].value,0)
        w=loaded['custom_threshold']
        equal('configurable performance threshold recalculates',w['DASHBOARD']['H38'].value,'TINGGI')
        equal('performance legend follows active threshold','10%' in w['DASHBOARD']['I37'].value,True)
        w=loaded['POS_Gian_Fataya']
        equal('production workbook has no invented sale price',w['PRODUK']['D7'].value,None)
        equal('source material cost sum',w['BAHAN']['D14'].value,9410)
        equal('production workbook no sales',w['DASHBOARD']['A8'].value,0)
        equal('production workbook no invoice selected',w['CETAK INVOICE']['A14'].value,'PILIH ID INVOICE')
        equal('production workbook no transaction IDs',sum(bool(w['MUTASI'].cell(r,1).value) for r in range(7,2007)),0)
        equal('blank workbook has no false high area',w['DASHBOARD']['H37'].value,'DATAR')
        equal('production default active status is not a transaction',w['MUTASI']['I7'].value,'AKTIF')
        # Keep a recalculated ready-to-open production artifact, without changing formulas via openpyxl.
        (ROOT/'output/POS_Gian_Fataya.xlsx').write_bytes((out/SOURCE.name).read_bytes())
        repair_navigation(SOURCE)
        equal('native internal navigation after recalculation',validate_navigation(SOURCE)['internal'],62)
    report={'passed':len(checks),'failed':0,'engine':'LibreOffice headless 26.8 alpha; real formula recalculation',
            'vba_runtime_tested':False,'excel_desktop_tested':False,'checks':checks}
    (ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps({'passed':len(checks),'failed':0},indent=2))


if __name__=='__main__':main()
