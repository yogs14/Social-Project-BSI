"""Real formula recalculation + end-to-end PF1 CSV tests; no VBA runtime claims."""
import csv,datetime as dt,hashlib,json,math,subprocess,sys,tempfile
from pathlib import Path
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'src'))
from build_workbooks import FILES
import transfer_cli as exchange
P=json.loads((ROOT/'data/profile.json').read_text());MR={m['code']:i+7 for i,m in enumerate(P['materials'])}
DRINK=next(m['code'] for m in P['materials'] if 'Santri' in m['name'])
BASE=next(m['code'] for m in P['materials'] if m['name']=='Base Pizza L')
RAW=P['materials'][0]['code'];CHECKS=[]
def day(v):return dt.datetime.fromisoformat(v)
def check(name,actual,expected):
    if isinstance(expected,(int,float)) and not isinstance(expected,bool):assert isinstance(actual,(int,float)) and math.isclose(actual,expected,abs_tol=1e-5,rel_tol=1e-8),(name,actual,expected)
    else:assert actual==expected,(name,actual,expected)
    CHECKS.append({'check':name,'result':'PASS','actual':actual})
def write(s,row,values):
    for c,v in values.items():
        s.cell(row,c,v)
        if isinstance(v,str):s.cell(row,c).data_type='s'
def stock(w,code):return w['BAHAN'].cell(MR[code],12).value
def move(w,row,id,when,kind,code,qty,**kw):
    v={1:id,2:day(when),3:kind,4:code,5:qty,13:'AKTIF',14:'SIMULASI BUKAN DATA MITRA'}
    for k,c in {'amount':6,'method':7,'transfer':8,'origin':9,'dest':10,'request':11,'nota':12,'shipped':20,'qtyshipped':21,'production':22}.items():
        if k in kw:v[c]=kw[k]
    write(w['MUTASI'],row,v)
def reject(name,fn,absent=None):
    try:fn()
    except (ValueError,FileExistsError):
        check(name,True,True)
        if absent:check(name+' / no partial output',Path(absent).exists(),False)
    else:raise AssertionError(name+' should have failed')
def edit_csv(source,dest,fn):
    with source.open(encoding='utf-8-sig',newline='') as f:rows=list(csv.DictReader(f))
    fn(rows)
    with dest.open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=exchange.HEADERS);writer.writeheader();writer.writerows(rows)

def main():
    for role,file in FILES.items():
        w=load_workbook(ROOT/'output'/file)
        check(role+' workbook identity',w['CONFIG']['B6'].value,role)
        check(role+' PF1 schema',w['CONFIG']['B7'].value,'PF1')
        check(role+' chart retained',len(w['DASHBOARD']._charts),1)
        check(role+' no external Excel links',len(w._external_links),0)
    check('Karah has no replacement cashier','KASIR' in load_workbook(ROOT/'output'/FILES['KARAH']).sheetnames,False)
    check('49 source menus',len(P['menus']),49)
    with tempfile.TemporaryDirectory(prefix='pizza-tests-') as folder:
        td=Path(folder);out=td/'calc';out.mkdir()
        wh=load_workbook(ROOT/'output'/FILES['GUDANG'])
        for r,code,qty in [(7,DRINK,100),(8,RAW,2000)]:move(wh,r,'AWAL-'+str(r),'2026-10-27','AWAL',code,qty)
        for r,id,code,qty,dest,req in [(9,'T1',DRINK,20,'MEDOKAN','R1'),(10,'T2',DRINK,10,'KARAH','R2'),(11,'T3',RAW,100,'MEDOKAN','R3')]:move(wh,r,id+'-G','2026-10-27','KIRIM',code,qty,transfer=id,origin='GUDANG',dest=dest,request=req)
        for r,id,code,qty,dest in [(7,'R1',DRINK,25,'MEDOKAN'),(8,'R2',DRINK,10,'KARAH'),(9,'R3',RAW,100,'MEDOKAN')]:write(wh['PERMINTAAN'],r,{1:id,2:day('2026-10-27'),3:dest,4:code,5:qty,6:'SIMULASI'})
        wh['DASHBOARD']['B5']=day('2026-10-01');wh['DASHBOARD']['E5']=day('2026-10-31')
        wh_path=td/'warehouse-send.xlsx';wh.save(wh_path)
        reject('mixed recipients require peer filter',lambda:exchange.export_records(wh_path,td/'mixed.csv','SHIPMENT'))
        m_csv=td/'medokan.csv';k_csv=td/'karah.csv'
        check('two Medokan shipment rows',exchange.export_records(wh_path,m_csv,'SHIPMENT','MEDOKAN')['exported'],2)
        check('one Karah shipment row',exchange.export_records(wh_path,k_csv,'SHIPMENT','KARAH')['exported'],1)
        mp=td/'medokan-received.xlsx';kp=td/'karah-received.xlsx'
        check('Medokan receipt batch',exchange.import_shipments(ROOT/'output'/FILES['MEDOKAN'],m_csv,mp,'2026-10-28',{'T1':19})['inserted'],2)
        check('Karah receipt batch',exchange.import_shipments(ROOT/'output'/FILES['KARAH'],k_csv,kp,'2026-10-28')['inserted'],1)
        repeat=td/'medokan-repeat.xlsx'
        check('duplicate shipment import skips without doubling',exchange.import_shipments(mp,m_csv,repeat,'2026-10-29')['skipped'],2)
        repeated=load_workbook(repeat)
        check('actual received quantity preserved on repeat',repeated['MUTASI']['E7'].value,19)
        check('receipt date not moved by repeat import',repeated['MUTASI']['B7'].value,day('2026-10-28'))
        before=hashlib.sha256(mp.read_bytes()).hexdigest()
        bad=td/'bad-unit.csv';edit_csv(m_csv,bad,lambda rows:rows[1].__setitem__('unit','kg'))
        reject('wrong unit batch rejected',lambda:exchange.import_shipments(mp,bad,td/'bad-output.xlsx','2026-10-28'),td/'bad-output.xlsx')
        check('failed import preserves source byte-for-byte',hashlib.sha256(mp.read_bytes()).hexdigest(),before)
        reject('wrong destination rejected',lambda:exchange.import_shipments(mp,k_csv,td/'wrong-dest.xlsx','2026-10-28'),td/'wrong-dest.xlsx')
        conflict=td/'conflict.csv';edit_csv(m_csv,conflict,lambda rows:rows[0].__setitem__('quantity_shipped','21'))
        reject('same ID changed shipment rejected',lambda:exchange.import_shipments(mp,conflict,td/'conflict.xlsx','2026-10-28'),td/'conflict.xlsx')
        dup=td/'dup.csv';edit_csv(m_csv,dup,lambda rows:rows.append(rows[0].copy()))
        reject('duplicate IDs in one CSV rejected',lambda:exchange.import_shipments(mp,dup,td/'dup.xlsx','2026-10-28'),td/'dup.xlsx')
        reject('source file cannot be overwritten',lambda:exchange.import_shipments(mp,m_csv,mp,'2026-10-28'))
        reject('receipt before shipment rejected',lambda:exchange.import_shipments(ROOT/'output'/FILES['MEDOKAN'],m_csv,td/'early.xlsx','2026-10-26'),td/'early.xlsx')
        inject=td/'text.csv';edit_csv(m_csv,inject,lambda rows:rows[0].__setitem__('note','=1+1'))
        safe=td/'safe-text.xlsx';exchange.import_shipments(ROOT/'output'/FILES['MEDOKAN'],inject,safe,'2026-10-28')
        check('CSV note not evaluated as Excel formula',load_workbook(safe)['MUTASI']['N7'].data_type,'s')
        ack_csv=td/'ack.csv';exchange.export_records(mp,ack_csv,'RECEIPT','GUDANG')
        acknowledged=td/'warehouse-ack.xlsx';check('receipt import records confirmations',exchange.import_receipts(wh_path,ack_csv,acknowledged)['inserted'],2)
        ack_repeat=td/'warehouse-ack-repeat.xlsx';check('repeated confirmation skipped',exchange.import_receipts(acknowledged,ack_csv,ack_repeat)['skipped'],2)
        wrongack=td/'wrong-ack.csv';edit_csv(ack_csv,wrongack,lambda rows:rows[0].__setitem__('quantity_shipped','999'))
        reject('confirmation must match original shipment',lambda:exchange.import_receipts(wh_path,wrongack,td/'wrong-ack.xlsx'),td/'wrong-ack.xlsx')
        returned=load_workbook(mp);move(returned,9,'RETURN-G','2026-10-29','KIRIM',DRINK,1,transfer='RET1',origin='MEDOKAN',dest='GUDANG')
        rp=td/'medokan-return.xlsx';returned.save(rp);rcsv=td/'return.csv';exchange.export_records(rp,rcsv,'SHIPMENT','GUDANG')
        warehouse_return=td/'warehouse-return.xlsx';exchange.import_shipments(acknowledged,rcsv,warehouse_return,'2026-10-30')

        # Actual formula fixtures: real journals with independent synthetic values.
        med=load_workbook(mp);med['DASHBOARD']['B5']=day('2026-10-01');med['DASHBOARD']['E5']=day('2026-10-31')
        med['BAHAN'].cell(MR[DRINK],5,2000);med['BAHAN'].cell(MR[DRINK],6,1)
        med['MENU']['G51']='YA'
        for r in range(7,607):
            if med['BOM'].cell(r,1).value=='PF045':med['BOM'].cell(r,5,'YA')
        move(med,9,'MD-AWAL','2026-10-27','AWAL',DRINK,20)
        move(med,10,'MD-BELI','2026-10-28','BELI',DRINK,12,amount=24000,method='Tunai')
        move(med,11,'MD-PAKAI','2026-10-28','PAKAI',DRINK,2,nota='MD-N1')
        for r,id,payment,discount,received,status in [(7,'MD-N1','GoPay',0,10000,'LENGKAP'),(8,'MD-N2','Tunai',5000,50000,'BELUM LENGKAP')]:write(med['NOTA'],r,{1:id,2:day('2026-10-28'),3:21/24,4:'SIMULASI',5:payment,6:discount,7:received,8:'AKTIF',9:'DATA UJI',17:status})
        for r,id,nota,sku,qty,price,hpp,state in [(7,'I1','MD-N1','PF045',2,5000,2000,'LENGKAP'),(8,'I2','MD-N2','PF001',1,45000,None,'BELUM LENGKAP'),(9,'I3','MD-N2','PF045',1,5000,2000,'LENGKAP')]:
            vals={1:id,2:nota,3:sku,4:qty,5:price,7:state,8:'SIMULASI'}
            if hpp is not None:vals[6]=hpp
            write(med['ITEM'],r,vals)
        write(med['KEUANGAN'],7,{1:'MD-K1',2:day('2026-10-28'),3:'BIAYA USAHA',4:1000,5:'Tunai',6:'SIMULASI',7:'AKTIF'})
        med['TUTUP KASIR']['B5']=day('2026-10-28');med['TUTUP KASIR']['B7']=10000;med['TUTUP KASIR']['E7']=30000
        med['STRUK']['B5']='MD-N2';med['KASIR']['B13']='PF049';med['KASIR']['D13']=1
        med_path=td/'medokan-sales.xlsx';med.save(med_path)
        void=load_workbook(med_path);void['NOTA']['H7']='VOID';void_path=td/'medokan-void.xlsx';void.save(void_path)
        historical=load_workbook(med_path);historical['DASHBOARD']['E5']=day('2026-10-27');hist_path=td/'medokan-historical.xlsx';historical.save(hist_path)
        kar=load_workbook(kp);kar['DASHBOARD']['B5']=day('2026-10-01');kar['DASHBOARD']['E5']=day('2026-10-31')
        write(kar['REKAP MAJOO'],7,{1:'K-MAJOO-1',2:day('2026-10-28'),3:100000,4:10000,5:40000,6:30000,7:0,8:20000,9:0,10:0,11:0,12:'AKTIF',18:'SIMULASI'})
        move(kar,8,'K-BELI','2026-10-28','BELI',DRINK,5,amount=5000,method='Tunai')
        move(kar,9,'K-PAKAI','2026-10-28','PAKAI',DRINK,2)
        write(kar['KEUANGAN'],7,{1:'K-FEE',2:day('2026-10-28'),3:'BIAYA USAHA',4:3000,5:'Transfer',6:'SIMULASI',7:'AKTIF'})
        kar['TUTUP KASIR']['B5']=day('2026-10-28');kar['TUTUP KASIR']['B7']=10000;kar['TUTUP KASIR']['E7']=45000
        kar_path=td/'karah-summary.xlsx';kar.save(kar_path)
        invalid=load_workbook(kar_path);invalid['REKAP MAJOO']['H7']=19000;badkar=td/'karah-invalid.xlsx';invalid.save(badkar)
        prod=load_workbook(acknowledged)
        write(prod['BAHAN'],80,{1:'UJI-RAW',2:'BAHAN SIMULASI',3:'UJI',4:'gram',8:'YA'})
        move(prod,12,'P-AWAL','2026-10-28','AWAL','UJI-RAW',1000)
        move(prod,13,'P-USE','2026-10-28','PRODUKSI PAKAI','UJI-RAW',500,production='P-UJI')
        move(prod,14,'P-IN','2026-10-28','PRODUKSI MASUK',BASE,2,production='P-UJI')
        prodpath=td/'warehouse-production.xlsx';prod.save(prodpath)
        incomplete=load_workbook(prodpath);incomplete['MUTASI']['V14']='P-OTHER';ipath=td/'warehouse-bad-production.xlsx';incomplete.save(ipath)
        files=[med_path,void_path,hist_path,kar_path,badkar,acknowledged,ack_repeat,warehouse_return,prodpath,ipath]+[ROOT/'output'/file for file in FILES.values()]
        result=subprocess.run(['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in files]],capture_output=True,text=True,timeout=360)
        assert result.returncode==0,result.stdout+'\n'+result.stderr
        check('all selected formula fixtures recalculated',len(list(out.glob('*.xlsx'))),len(files))
        loaded={p.stem:load_workbook(out/p.name,data_only=True) for p in files}
        for name,wb in loaded.items():
            errors=[f'{s.title}!{c.coordinate}={c.value}' for s in wb for row in s for c in row if c.data_type=='e']
            check(name+' no formula errors',errors,[])
        m=loaded['medokan-sales']
        for label,actual,expected in [('49th menu usable at cashier',m['KASIR']['E13'].value,5000),('sale receipt valid',m['NOTA']['O7'].value,'OK'),('pending recipe does not erase cash sale',m['NOTA']['O8'].value,'OK'),('Medokan net sales',m['DASHBOARD']['A8'].value,55000),('GoPay available and counted',m['DASHBOARD']['H29'].value,10000),('cash sales after discount',m['DASHBOARD']['H26'].value,45000),('sale change uses net',m['NOTA']['L8'].value,5000),('Santri purchase cash-out',m['TUTUP KASIR']['B11'].value,24000),('Medokan drawer excludes GoPay',m['TUTUP KASIR']['B14'].value,30000),('Medokan drawer reconciles',m['TUTUP KASIR']['B15'].value,0),('known recipe HPP complete',m['MENU']['I51'].value,'LENGKAP'),('unconfirmed pizza HPP incomplete',m['MENU']['I7'].value,'BELUM LENGKAP'),('pending recipe notice',m['DASHBOARD']['D22'].value,1),('physical Santri stock distinct from pending note',stock(m,DRINK),49),('raw transfer received once',stock(m,RAW),100),('L size quantity recap',m['DASHBOARD']['M26'].value,1),('all drinks quantity recap',m['DASHBOARD']['M30'].value,3),('receipt first line quantity',m['STRUK']['D13'].value,1),('receipt second line quantity',m['STRUK']['D14'].value,1),('receipt final total',m['STRUK']['F28'].value,45000)]:check(label,actual,expected)
        v=loaded['medokan-void'];check('VOID reverses sale',v['DASHBOARD']['A8'].value,45000);check('VOID restores consumed stock',stock(v,DRINK),51)
        h=loaded['medokan-historical'];check('historical receipt cutoff',stock(h,DRINK),20);check('historical sales cutoff',h['DASHBOARD']['A8'].value,0)
        k=loaded['karah-summary']
        for label,actual,expected in [('Majoo report reconciles',k['REKAP MAJOO']['O7'].value,'OK'),('Karah net not double-discounted',k['DASHBOARD']['A8'].value,90000),('Karah cash excludes noncash',k['TUTUP KASIR']['B10'].value,40000),('Karah independent purchases tracked',stock(k,DRINK),13),('Karah cash balance',k['TUTUP KASIR']['B14'].value,45000),('Karah drawer reconciles',k['TUTUP KASIR']['B15'].value,0)]:check(label,actual,expected)
        k=loaded['karah-invalid'];check('Majoo payment mismatch detected',k['REKAP MAJOO']['O7'].value,'SELISIH METODE');check('invalid Majoo report excluded',k['DASHBOARD']['A8'].value,0)
        a=loaded['warehouse-ack'];check('ack does not add stock to sender',stock(a,DRINK),70);check('partial receipt discrepancy visible',a['KONFIRMASI']['K7'].value,'SELISIH');check('partial receipt delta',a['KONFIRMASI']['L7'].value,-1);check('matching receipt validates',a['KONFIRMASI']['K8'].value,'OK');check('transfer never becomes revenue',a['DASHBOARD']['A8'].value,0);check('request partial realization',a['PERMINTAAN']['G7'].value,20);check('request remaining',a['PERMINTAAN']['H7'].value,5);check('request status',a['PERMINTAAN']['I7'].value,'SEBAGIAN');check('Karah request fulfilled',a['PERMINTAAN']['I8'].value,'TERPENUHI')
        check('repeated ack still does not move stock',stock(loaded['warehouse-ack-repeat'],DRINK),70)
        check('new reverse transfer adds return once',stock(loaded['warehouse-return'],DRINK),71)
        z=loaded['warehouse-production'];check('paired production input validates',z['MUTASI']['Q13'].value,'OK');check('paired production output validates',z['MUTASI']['Q14'].value,'OK');check('production output stock',stock(z,BASE),2);check('production input stock',z['BAHAN']['L80'].value,500)
        z=loaded['warehouse-bad-production'];check('one-sided production detected',z['MUTASI']['Q13'].value,'LENGKAPI DUA SISI PRODUKSI');check('one-sided production excluded',z['MUTASI']['P13'].value,0)
        for role,file in FILES.items():
            z=loaded[Path(file).stem]
            check(role+' operational stock movements blank',sum(bool(z['MUTASI'].cell(r,1).value) for r in range(7,4007)),0)
            check(role+' operational sales blank',z['DASHBOARD']['A8'].value,0)
            check(role+' source ingredient price not invented',z['BAHAN']['E7'].value,None)
            (ROOT/'output'/file).write_bytes((out/file).read_bytes())
    report={'passed':len(CHECKS),'failed':0,'formula_engine':'LibreOffice headless 26.8 alpha; actual workbook formulas','csv_engine':'Python offline CSV round trip; real workbook files','vba_runtime_tested':False,'excel_desktop_tested':False,'checks':CHECKS}
    (ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str)+'\n')
    print(json.dumps({'passed':len(CHECKS),'failed':0}))
if __name__=='__main__':main()
