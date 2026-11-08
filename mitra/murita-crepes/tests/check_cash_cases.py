"""Exercise actual closing formulas for QRIS, over/short cash and missing input."""
import datetime as dt,json,subprocess,tempfile
from pathlib import Path
from openpyxl import Workbook,load_workbook

def verify(source,report_path):
    source=load_workbook(source,data_only=False,read_only=True)
    formulas={a:source['TUTUP KASIR'][a].value for a in ['F8','F9','F10','F11','F12','F13','F14','F15','I9','I10','I12','H14']}
    source.close();w=Workbook();w.remove(w.active);j=w.create_sheet('JURNAL')
    for r,date,method,net in [(7,'2026-10-24','Tunai',14000),(8,'2026-10-25','QRIS',5000)]:
        j.cell(r,2,dt.datetime.fromisoformat(date));j.cell(r,5,method);j.cell(r,16,net)
    j.cell(4035,2,dt.datetime(2026,10,24));j.cell(4035,8,'Tunai');j.cell(4035,16,2400)
    for r,kind,amount in [(10049,'BIAYA USAHA',-1000),(10050,'PRIVE',-2000),(10051,'MODAL MASUK',3000),(10052,'PENDAPATAN LAIN',500)]:
        j.cell(r,2,dt.datetime(2026,10,24));j.cell(r,3,kind);j.cell(r,5,'Tunai');j.cell(r,9,amount)
    cases=[('QRIS',dt.datetime(2026,10,25),0,0),('LEBIH',dt.datetime(2026,10,24),10000,22200),('KURANG',dt.datetime(2026,10,24),10000,22000),('BELUM_FISIK',dt.datetime(2026,10,25),0,None)]
    for name,date,opening,physical in cases:
        s=w.create_sheet(name);s['C5']=date;s['I8']=opening;s['I11']=physical
        for cell,formula in formulas.items():s[cell]=formula
    checks=[]
    with tempfile.TemporaryDirectory(prefix='mufrita-cash-cases-') as td:
        td=Path(td);out=td/'calc';out.mkdir();p=td/'cash_cases.xlsx';w.save(p)
        result=subprocess.run(['soffice','-env:UserInstallation='+(td/'profile').as_uri(),'--headless','--convert-to','xlsx','--outdir',str(out),str(p)],capture_output=True,text=True,timeout=120)
        assert result.returncode==0,result.stdout+result.stderr
        calc=load_workbook(out/p.name,data_only=True)
        expected=[('QRIS','F9',5000),('QRIS','F8',0),('QRIS','I10',0),('QRIS','I12',0),('QRIS','H14','KAS COCOK - TIDAK ADA SELISIH'),('LEBIH','I10',22100),('LEBIH','I12',100),('LEBIH','H14','KAS LEBIH Rp 100'),('KURANG','I12',-100),('KURANG','H14','KAS KURANG Rp 100'),('BELUM_FISIK','I12',None),('BELUM_FISIK','H14','ISI UANG FISIK (0 BILA KOSONG)')]
        for name,cell,value in expected:
            actual=calc[name][cell].value
            assert actual==value,(name,cell,actual,value)
            checks.append({'check':f'Closing cash input / {name} {cell}','result':'PASS','actual':actual})
    report=json.loads(Path(report_path).read_text())
    report['checks']=[c for c in report['checks'] if not c['check'].startswith('Closing cash input / ')]+checks
    report['passed']=len(report['checks']);Path(report_path).write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str))
    return checks

if __name__=='__main__':
    import sys
    print('Additional cash checks:',len(verify(Path(sys.argv[1]),Path(sys.argv[2]))))
