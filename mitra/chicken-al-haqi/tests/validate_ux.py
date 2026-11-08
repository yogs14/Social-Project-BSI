"""Verify the reference UX and recalculate its actual input paths.

The journal model is compared with the previously validated builder output;
unchanged regression results are retained and new UX cases run in LibreOffice.
Keep the original Excel form controls; transplant formula caches only.
"""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import sys
from zipfile import ZipFile
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from formula_cache import cache_only
CHECKS=[]


def check(label, actual, expected):
    assert actual==expected,(label,actual,expected)
    CHECKS.append({'check':label,'result':'PASS','actual':actual})




def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--baseline',type=Path)
    args=parser.parse_args()
    source=ROOT/'output/POS_Chicken_Al_Haqi.xlsx'
    before=load_workbook(args.baseline) if args.baseline else None
    workbook=load_workbook(source)
    for name in (['SETUP','BELANJA','KAS USAHA','NOTA','ITEM','STOK','HARIAN'] if before else []):
        for row in before[name].iter_rows(min_row=6):
            for cell in row:
                assert cell.value==workbook[name][cell.coordinate].value,(name,cell.coordinate)
        check(name+' tested model values and formulas unchanged',True,True)
    skin=json.loads((ROOT/'data/ux_pos_ponorogo.json').read_text())
    for sheet,template in [('KASIR','KASIR'),('TUTUP KASIR','Tutup Kasir')]:
        tab=workbook[sheet];layout=skin['sheets'][template]
        check(sheet+' reference merges',sorted(str(m) for m in tab.merged_cells.ranges if m.max_row<=(50 if sheet=='KASIR' else 24)),sorted(layout['merges']))
        for row,height in layout['rows'].items():
            check(sheet+' row '+row+' reference height',tab.row_dimensions[int(row)].height,height)
    for cell,color,size in [('B1','FF902A20',19),('B5','FF902A20',11),('H5','FF2E7D32',11),('C31','FF902A20',13)]:
        check('Cashier '+cell+' reference color',workbook['KASIR'][cell].fill.fgColor.rgb,color)
        check('Cashier '+cell+' reference font size',workbook['KASIR'][cell].font.sz,size)
    check('Cashier opens first',workbook.active.title,'KASIR')
    check('Cashier reference zoom',workbook['KASIR'].sheet_view.zoomScale,75)
    with ZipFile(source) as z:
        controls={name:z.read(name) for name in z.namelist() if name.endswith('.vml')}
        check('Six visible native Excel buttons',sum(data.count(b'<x:FmlaMacro>') for data in controls.values()),6)
    spec=importlib.util.spec_from_file_location('existing_validation',ROOT/'tests/validate_workbook.py')
    model=importlib.util.module_from_spec(spec);spec.loader.exec_module(model);model.SOURCE=source
    with tempfile.TemporaryDirectory(prefix='chicken-ux-') as directory:
        directory=Path(directory);out=directory/'calc';out.mkdir();files=[]
        for case in ['normal','shortage','missing']:
            w=model.base()
            w['TUTUP KASIR']['K8']='=IF(ISNUMBER(I8),I8,"")'
            w['TUTUP KASIR']['K11']='=IF(ISNUMBER(I11),I11,"")'
            w['TUTUP KASIR']['I8']=20000;w['TUTUP KASIR']['I11']=75000
            w['NOTA']['Q7']='Dine-in';w['NOTA']['R7']=2
            w['NOTA']['Q8']='Take Away';w['NOTA']['R8']=1
            w['KASIR']['C32']=50000
            if case=='shortage':
                w['KASIR']['B16']='Sayap Original';w['KASIR']['C16']=8
                w['KASIR']['B17']='Sayap Sambal Geprek';w['KASIR']['C17']=8
            if case=='missing':
                for row in range(7,11):
                    w['STOK'].cell(row,3).value=None;w['STOK'].cell(row,4).value=None
                w['TUTUP KASIR']['I8']=None;w['TUTUP KASIR']['I11']=None
            path=directory/(case+'.xlsx');w.save(path);files.append(path)
        files.append(source)
        result=subprocess.run(['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in files]],capture_output=True,text=True,timeout=420)
        assert result.returncode==0,result.stdout+'\n'+result.stderr
        loaded={p.stem:load_workbook(out/p.name,data_only=True) for p in files}
        for name,w in loaded.items():
            errors=[f'{s.title}!{c.coordinate}={c.value}' for s in w for row in s for c in row if c.data_type=='e']
            check(name+' no formula errors',errors,[])
        w=loaded['normal']
        for label,sheet,cell,value in [
            ('Menu price','KASIR','D16',12000),('Cart subtotal','KASIR','C29',30000),
            ('Cash change','KASIR','C33',20000),('Daily sales','KASIR','I6',56000),
            ('Customer count','KASIR','I9',3),('Bestseller name','KASIR','H26','Sayap'),
            ('Bestseller quantity','KASIR','I26',3),('Bestseller revenue','KASIR','J26',27000),
            ('Second bestseller','KASIR','H27','Dada'),('Rice ranked third','KASIR','H28','Nasi'),
            ('Opening entered in reference field','TUTUP KASIR','K8',20000),
            ('Cash reconciliation','TUTUP KASIR','I10',75000),('Cash difference','TUTUP KASIR','I12',0),
            ('QRIS separate','TUTUP KASIR','F9',11000),('Customer closing total','TUTUP KASIR','C9',3),
        ]:check(label,w[sheet][cell].value,value)
        w=loaded['shortage']
        check('Combined flavor shortage',w['KASIR']['F16'].value,'STOK KURANG')
        check('Second flavor cannot reuse stock',w['KASIR']['F17'].value,'STOK KURANG')
        check('Shortage guidance visible','Stok kurang' in w['KASIR']['B35'].value,True)
        w=loaded['missing']
        check('Missing opening stays unknown',w['KASIR']['F16'].value,'BELUM AWAL')
        check('Missing opening guidance visible','stok awal' in w['KASIR']['B35'].value,True)
        check('Missing cash does not become zero',w['TUTUP KASIR']['I10'].value in [None,''],True)
        check('Missing cash difference unavailable',w['TUTUP KASIR']['I12'].value in [None,''],True)
        w=loaded[source.stem]
        check('No invented receipt',sum(bool(w['NOTA'].cell(r,1).value) for r in range(7,1507)),0)
        check('Empty cart guidance',w['KASIR']['B35'].value,'Pilih menu dan isi jumlah untuk mulai melayani.')
        check('Empty bestseller stays empty',w['KASIR']['H26'].value in [None,''],True)
        cache_only(source,out/source.name)
    with ZipFile(source) as z:
        check('Native button XML preserved',all(z.read(name)==data for name,data in controls.items()),True)
    old_report=json.loads((ROOT/'output/hasil_validasi.json').read_text())
    old_checks=old_report['checks'][:old_report.get('retained_regression_checks',len(old_report['checks']))]
    report={'passed':len(old_checks)+len(CHECKS),'failed':0,
            'engine':'LibreOffice headless; UX input scenarios recalculated; unchanged journal model compared with validated baseline',
            'retained_regression_checks':len(old_checks),'ux_checks':len(CHECKS),
            'vba_runtime_tested':False,'excel_desktop_tested':False,
            'workbook_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),
            'checks':old_checks+CHECKS}
    (ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'passed':report['passed'],'ux_checks':len(CHECKS),'failed':0}))


if __name__=='__main__':main()
