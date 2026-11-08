"""Use actual Excel formulas and synthetic journals; VBA is not executed here."""
import datetime as dt,json,math,subprocess,tempfile,hashlib,sys
from zipfile import ZipFile
from pathlib import Path
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1];SOURCE=ROOT/'output/POS_Chicken_Al_Haqi.xlsx';CHECKS=[]
def date(s):return dt.datetime.fromisoformat(s)
def check(label,actual,expected):
    if isinstance(expected,(int,float)) and not isinstance(expected,bool):assert isinstance(actual,(int,float)) and math.isclose(actual,expected,abs_tol=1e-6,rel_tol=1e-9),(label,actual,expected)
    else:assert actual==expected,(label,actual,expected)
    CHECKS.append({'check':label,'result':'PASS','actual':actual})
def write(s,r,values):
    for c,v in values.items():s.cell(r,c,v)
def base():
    w=load_workbook(SOURCE);w['TUTUP KASIR']['C5']=date('2026-10-24');w['TUTUP KASIR']['K8']=20000;w['TUTUP KASIR']['K11']=75000;w['KASIR']['C7']=date('2026-10-24')
    for r in range(7,11):w['SETUP'].cell(r,7,'YA');write(w['STOK'],r,{3:date('2026-10-24'),4:5})
    w['SETUP']['F11']=1500;w['SETUP']['G11']='YA'
    for r,cost in [(32,1000),(33,1200),(34,900)]:w['SETUP'].cell(r,3,cost);w['SETUP'].cell(r,4,'YA')
    for r,id,day,code,kg,pcs,total,method in [(7,'B1','2026-10-24','DADA',2,17,70000,'Tunai'),(8,'B2','2026-10-24','SAYAP',1,11,35000,'Transfer'),(9,'B3','2026-10-25','DADA',5,47,175000,'Tunai')]:write(w['BELANJA'],r,{1:id,2:date(day),3:code,4:kg,5:pcs,6:total,7:method,8:'AKTIF',9:'SIMULASI BUKAN TRANSAKSI MITRA'})
    for r,key,day,time,method,disc,received in [(7,'N1','2026-10-24',6.5/24,'Tunai',1000,50000),(8,'N2','2026-10-24',7/24,'QRIS',0,11000),(9,'N3','2026-10-25',9/24,'Tunai',0,10000)]:write(w['NOTA'],r,{1:key,2:date(day),3:time,4:'KASIR SIMULASI',5:method,6:disc,7:received,8:'AKTIF',9:'SIMULASI'})
    for r,nota,code,flavor,qty,price,hpp in [(7,'N1','SAYAP','Original',2,8000,7000),(8,'N1','DADA','Saus Keju',2,12000,8200),(9,'N1','NASI','Original',2,3000,1500),(10,'N2','SAYAP','Hot Spicy',1,11000,7900),(11,'N3','DADA','Original',1,9000,7000)]:write(w['ITEM'],r,{1:'I'+str(r),2:nota,3:code,4:flavor,5:qty,6:price,7:hpp,8:'DIKONFIRMASI'})
    for r,kind,amount in [(7,'BIAYA USAHA',10000),(8,'MODAL MASUK',100000),(9,'PRIVE',5000),(10,'KE USAHA LAIN',7000),(11,'DARI USAHA LAIN',2000)]:write(w['KAS USAHA'],r,{1:'K'+str(r),2:date('2026-10-24'),3:kind,4:amount,5:'Tunai',6:'SIMULASI; USAHA LAIN MISAL OBAT',7:'AKTIF'})
    write(w['STOK'],36,{1:'C1',2:date('2026-10-24'),3:'SAYAP',4:'RUSAK',5:1,6:'SIMULASI',7:'AKTIF'})
    write(w['STOK'],37,{1:'C2',2:date('2026-10-25'),3:'PAHA-ATAS',4:'TAMBAH',5:2,6:'SIMULASI',7:'AKTIF'})
    w['KASIR']['B16']='Dada Saus Keju';w['KASIR']['C16']=2;w['KASIR']['B17']='Nasi';w['KASIR']['C17']=2
    w['KASIR']['I54']='DADA';w['KASIR']['I55']=2;w['SETUP']['B59']=5
    return w
def fingerprint(path):
    # ZIP timestamps and the editor modified time do not change the model inputs.
    with ZipFile(path) as archive:
        return hashlib.sha256(b''.join(name.encode()+archive.read(name) for name in sorted(archive.namelist()) if name!='docProps/core.xml')).hexdigest()

def main():
    original=load_workbook(SOURCE)
    check('nine visible user tabs',sum(s.sheet_state=='visible' for s in original),9)
    check('one hidden item journal',sum(s.sheet_state=='hidden' for s in original),1)
    check('daily chart present',len(original['HARIAN']._charts),1)
    check('no external workbook links',len(original._external_links),0)
    check('VBA not falsely embedded',original.vba_archive,None)
    check('correction table starts separately',original['STOK'].tables['tblKoreksi'].ref,'A35:J235')
    with tempfile.TemporaryDirectory(prefix='chicken-tests-') as td:
        td=Path(td);out=td/'calc';out.mkdir();files=[]
        for name in ['normal','unknown_opening','incomplete_hpp','changed_master','void','nextday','before_opening','invalid','negative_stock','duplicate_purchase','flavor_shortage','twelve_items','thirteen_items','invalid_channel']:
            w=base()
            if name=='unknown_opening':
                for r in range(7,11):w['STOK'].cell(r,3).value=None;w['STOK'].cell(r,4).value=None
            if name=='incomplete_hpp':
                for r in range(7,12):w['SETUP'].cell(r,7,'TIDAK')
                for r in range(7,12):w['ITEM'].cell(r,7).value=None;w['ITEM'].cell(r,8,'BELUM DIKONFIRMASI')
            if name=='changed_master':w['SETUP']['C7']=9500;w['SETUP']['F7']=9999;w['SETUP']['B33']=4500;w['SETUP']['E10']=8
            if name=='void':w['NOTA']['H7']='VOID'
            if name=='nextday':w['TUTUP KASIR']['C5']=date('2026-10-25')
            if name=='before_opening':w['TUTUP KASIR']['C5']=date('2026-10-23')
            if name=='invalid':
                w['BELANJA']['E7']=1.5;w['ITEM']['D9']='Sambal Geprek';w['NOTA']['G8']=1000
                write(w['STOK'],38,{1:'BAD-C',2:date('2026-10-24'),3:'DADA',4:'RUSAK',5:999,6:'SIMULASI',7:'AKTIF'})
            if name=='negative_stock':w['STOK']['D7']=0;w['BELANJA']['H8']='VOID'
            if name=='duplicate_purchase':w['BELANJA']['A8']='B1'
            if name=='flavor_shortage':
                w['KASIR']['B16']='Sayap Original';w['KASIR']['C16']=8
                w['KASIR']['B17']='Sayap Sambal Geprek';w['KASIR']['C17']=8
            if name in ['twelve_items','thirteen_items']:
                w['STOK']['D7']=100
                count=12 if name=='twelve_items' else 13
                write(w['NOTA'],10,{1:'N-MAX',2:date('2026-10-24'),3:9/24,4:'SIMULASI',5:'Tunai',6:0,7:200000,8:'AKTIF',17:'Take Away',18:3})
                for j in range(count):write(w['ITEM'],12+j,{1:'MAX-'+str(j),2:'N-MAX',3:'SAYAP',4:'Original',5:1,6:8000,7:7000,8:'DIKONFIRMASI'})
            if name=='invalid_channel':w['NOTA']['Q7']='BUKAN CHANNEL'
            dest=td/(name+'.xlsx');w.save(dest);files.append(dest)
        files.append(SOURCE)
        pending=[]
        reuse=Path(sys.argv[2]) if len(sys.argv)==3 and sys.argv[1]=='--reuse-calculated' else None
        for file in files:
            prior=(reuse/file.name) if reuse else None
            calculated=(reuse/'calc'/file.name) if reuse else None
            if prior and prior.exists() and calculated.exists() and fingerprint(prior)==fingerprint(file):
                (out/file.name).write_bytes(calculated.read_bytes())
            else:pending.append(file)
        if pending:
            result=subprocess.run(['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in pending]],capture_output=True,text=True,timeout=900)
            assert result.returncode==0,result.stdout+'\n'+result.stderr
        if reuse:
            for file in files:
                (reuse/file.name).write_bytes(file.read_bytes())
                (reuse/'calc'/file.name).write_bytes((out/file.name).read_bytes())
        check('all selected workbook files recalculated',len(list(out.glob('*.xlsx'))),len(files))
        loaded={p.stem:load_workbook(out/p.name,data_only=True) for p in files}
        for name,w in loaded.items():
            errors=[f'{s.title}!{c.coordinate}={c.value}' for s in w for row in s for c in row if c.data_type=='e']
            check(name+' no formula errors',errors,[])
        w=loaded['normal']
        for label,actual,expected in [
          ('merged sauce prices correctly applied',w['SETUP']['B33'].value,3000),('other merged sauce price',w['SETUP']['B34'].value,3000),
          ('source per-pcs cost sums to 7000 only once',w['SETUP']['C55'].value,7000),
          ('two kg dada estimate is 18, not stock',w['BELANJA']['J7'].value,18),('actual 17 pcs differs from estimate',w['BELANJA']['K7'].value,-1),('actual pcs is effective stock input',w['BELANJA']['M7'].value,17),
          ('sayap per-kg estimate differs by one',w['BELANJA']['K8'].value,1),('purchase above three kg allowed',w['BELANJA']['L9'].value,'OK'),
          ('five kg planning split into two seasoning batches',w['SETUP']['B62'].value,2),
          ('cashier adds sauce per piece',w['KASIR']['D16'].value,12000),('two sauce portions charge twice',w['KASIR']['E16'].value,24000),('rice separate subtotal',w['KASIR']['E17'].value,6000),('cashier gross',w['KASIR']['C29'].value,30000),
          ('Saturday 06:30 accepted',w['NOTA']['O7'].value,'OK'),('early 07:00 accepted',w['NOTA']['O8'].value,'OK'),
          ('receipt gross includes rice and sauce',w['NOTA']['J7'].value,46000),('discount once per receipt',w['NOTA']['K7'].value,45000),('cash change based on netto',w['NOTA']['L7'].value,5000),
          ('daily net sales',w['HARIAN']['A8'].value,56000),('daily receipt count',w['HARIAN']['E8'].value,2),('chicken count excludes rice',w['HARIAN']['I8'].value,5),
          ('purchase cash and transfer total',w['HARIAN']['A13'].value,105000),('business expenses exclude owner/medicine transfers',w['HARIAN']['E13'].value,10000),('confirmed HPP comparison excludes double purchase cost',w['HARIAN']['I13'].value,14700),
          ('only net cash sales enter drawer',w['HARIAN']['B18'].value,45000),('only cash purchases reduce drawer',w['HARIAN']['B19'].value,70000),('owner and other-business cash movement included once',w['HARIAN']['B20'].value,80000),('cash balance reconciles',w['HARIAN']['B22'].value,75000),('physical cash difference zero',w['HARIAN']['B23'].value,0),('QRIS kept separate',w['HARIAN']['F20'].value,11000),
          ('dada stock actual pcs not kg plus pcs',w['STOK']['H10'].value,20),('sayap sale and waste once',w['STOK']['H7'].value,12),('nasi deliberately not tracked',w['STOK']['I11'].value,'TIDAK DILACAK'),('waste does not become revenue',w['STOK']['I36'].value,-1),('unconfirmed-stock alert clear with initial counts',w['HARIAN']['D36'].value,0),
        ]:check(label,actual,expected)
        check('flattened menu has seventeen options',sum(bool(original['SETUP'].cell(r,1).value) for r in range(69,89)),17)
        check('cashier maps menu to the existing dada stock',w['KASIR']['X16'].value,'DADA')
        check('cashier maps selected sauce',w['KASIR']['Y16'].value,'Saus Keju')
        check('live cashier revenue uses cashier date',w['KASIR']['I6'].value,56000)
        check('closing cash matches existing business reconciliation',w['TUTUP KASIR']['K14'].value,75000)
        check('closing physical difference matches dashboard',w['TUTUP KASIR']['K15'].value,0)
        check('closing QRIS is separate',w['TUTUP KASIR']['E12'].value,11000)
        w=loaded['flavor_shortage'];check('different flavors share stock requirement',w['KASIR']['J15'].value,16);check('shortage detected across flavors',w['KASIR']['F16'].value,'STOK KURANG');check('second flavor cannot reuse same stock',w['KASIR']['F17'].value,'STOK KURANG')
        w=loaded['twelve_items'];check('twelve-line receipt accepted',w['NOTA']['O10'].value,'OK');check('twelve-line receipt total',w['NOTA']['K10'].value,96000)
        w=loaded['thirteen_items'];check('thirteenth line rejected',w['NOTA']['O10'].value,'MAKS 12 BARIS')
        w=loaded['invalid_channel'];check('invalid channel rejected',w['NOTA']['O7'].value,'CEK CHANNEL')
        w=loaded['unknown_opening'];check('missing stock opening is labelled',w['STOK']['J7'].value,'BELUM AWAL');check('unknown stock stays blank not fictitious zero',w['STOK']['H7'].value in [None,''],True);check('four unknown chicken openings visible',w['HARIAN']['D36'].value,4);check('cash recording survives missing stock opening',w['HARIAN']['A8'].value,56000)
        w=loaded['incomplete_hpp'];check('unconfirmed HPP never presented as profit',w['HARIAN']['I13'].value,'HPP BELUM DIKONFIRMASI');check('unknown HPP does not erase actual cash sales',w['HARIAN']['A8'].value,56000)
        w=loaded['changed_master'];check('changed sauce surcharge applies only to next cart',w['KASIR']['D16'].value,13500);check('posted sale prices remain frozen',w['HARIAN']['A8'].value,56000);check('posted HPP remains frozen',w['HARIAN']['I13'].value,14700);check('actual pcs unaffected by changed yield estimate',w['STOK']['H10'].value,20)
        w=loaded['void'];check('VOID removes receipt net',w['HARIAN']['A8'].value,11000);check('VOID restores dada pcs',w['STOK']['H10'].value,22);check('VOID restores sayap pcs',w['STOK']['H7'].value,14)
        w=loaded['nextday'];check('report one selected day only',w['HARIAN']['A8'].value,9000);check('stock carries across days including next purchase',w['STOK']['H10'].value,66);check('next-day physical correction included',w['STOK']['H9'].value,7)
        w=loaded['before_opening'];check('sales before report cutoff excluded',w['HARIAN']['A8'].value,0);check('stock before opening not invented',w['STOK']['H10'].value in [None,''],True);check('stock before opening labelled',w['STOK']['I10'].value,'BELUM MULAI')
        w=loaded['invalid'];check('fractional purchased pieces rejected',w['BELANJA']['L7'].value,'ISI PCS AKTUAL');check('invalid purchase excluded from stock',w['BELANJA']['M7'].value,0);check('nasi cannot have sauce price',w['ITEM']['M9'].value,'NASI TANPA SAUS');check('insufficient payment detected',w['NOTA']['O8'].value,'CEK PEMBAYARAN');check('invalid receipts not counted as sales',w['HARIAN']['A8'].value,0);check('oversized waste creates negative alert',w['STOK']['J38'].value<0,True)
        w=loaded['negative_stock'];check('stock shortage visible in sale history',w['ITEM']['O7'].value<0,True);check('negative closing stock marked',w['STOK']['I7'].value,'STOK NEGATIF')
        w=loaded['duplicate_purchase'];check('duplicate purchase IDs rejected',w['BELANJA']['L7'].value,'ID GANDA');check('duplicate purchase not added twice',w['BELANJA']['M8'].value,0)
        w=loaded['POS_Chicken_Al_Haqi'];check('main workbook no fake receipts',sum(bool(w['NOTA'].cell(r,1).value) for r in range(7,1507)),0);check('main workbook no fake purchases',sum(bool(w['BELANJA'].cell(r,1).value) for r in range(7,1507)),0);check('main workbook initial stock unknown',w['STOK']['J7'].value,'BELUM AWAL');check('main HPP not confirmed by assumption',w['SETUP']['G7'].value,'TIDAK');check('no invented opening cash',w['HARIAN']['B22'].value in [None,''],True)
        SOURCE.write_bytes((out/SOURCE.name).read_bytes())
    report={'passed':len(CHECKS),'failed':0,'engine':'LibreOffice headless; actual workbook formula recalculation','vba_runtime_tested':False,'excel_desktop_tested':False,'checks':CHECKS}
    (ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str)+'\n')
    print(json.dumps({'passed':len(CHECKS),'failed':0}))
if __name__=='__main__':main()
