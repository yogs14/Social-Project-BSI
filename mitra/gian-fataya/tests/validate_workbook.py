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

ROOT=Path(__file__).resolve().parents[1]
SOURCE=ROOT/'output/POS_Gian_Fataya.xlsx'
checks=[]


def day(s):return dt.datetime.fromisoformat(s)


def equal(name,actual,expected):
    assert actual==expected,f'{name}: {actual!r} != {expected!r}'
    checks.append({'check':name,'result':'PASS','actual':actual})


def base():
    w=load_workbook(SOURCE)
    w['PRODUK']['D7']=10000
    for r,code,area in [(8,'UJI-JT','W01'),(9,'UJI-JG','W02')]:
        for c,v in enumerate([code,'SIMULASI - BUKAN OUTLET ASLI',area,'PIC UJI','ALAMAT UJI','','AKTIF'],1):w['OUTLET'].cell(r,c,v)
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
    original=load_workbook(SOURCE,data_only=False)
    equal('invoice print area belongs to invoice',str(original['CETAK INVOICE'].print_area),"'CETAK INVOICE'!$A$7:$F$79")
    equal('dashboard print area belongs to dashboard',str(original['DASHBOARD'].print_area),"'DASHBOARD'!$A$1:$Q$48")
    equal('movement table includes PO and freight',original['MUTASI'].tables['tblMutasi'].ref,'A6:W2006')
    equal('VBA not falsely embedded in xlsx',original.vba_archive,None)
    equal('expected worksheet count',len(original.sheetnames),14)
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
            dest=td/(name+'.xlsx');w.save(dest);files.append(dest)
        # Test the real blank workbook too, without editing or adding fixtures.
        files.append(SOURCE)
        command=['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in files]]
        result=subprocess.run(command,capture_output=True,text=True,timeout=240)
        if result.returncode:raise RuntimeError(result.stdout+'\n'+result.stderr)
        equal('Recalculated expected files',len(list(out.glob('*.xlsx'))),len(files))
        loaded={p.stem:load_workbook(out/p.name,data_only=True) for p in files}
        for name,w in loaded.items():no_errors(name,w)
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
        equal('unbilled other region',w['DASHBOARD']['K21'].value,50000)
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
        equal('void invoice sales available for rebilling',w['DASHBOARD']['K21'].value,350000)
        w=loaded['POS_Gian_Fataya']
        equal('production workbook has no invented sale price',w['PRODUK']['D7'].value,None)
        equal('source material cost sum',w['BAHAN']['D14'].value,9410)
        equal('production workbook no sales',w['DASHBOARD']['A8'].value,0)
        equal('production workbook no invoice selected',w['CETAK INVOICE']['A14'].value,'PILIH ID INVOICE')
        equal('production workbook no transaction IDs',sum(bool(w['MUTASI'].cell(r,1).value) for r in range(7,2007)),0)
        # Keep a recalculated ready-to-open production artifact, without changing formulas via openpyxl.
        (ROOT/'output/POS_Gian_Fataya.xlsx').write_bytes((out/SOURCE.name).read_bytes())
    report={'passed':len(checks),'failed':0,'engine':'LibreOffice headless 26.8 alpha; real formula recalculation',
            'vba_runtime_tested':False,'excel_desktop_tested':False,'checks':checks}
    (ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps({'passed':len(checks),'failed':0},indent=2))


if __name__=='__main__':main()
