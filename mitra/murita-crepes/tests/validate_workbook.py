"""Recalculate actual Excel formulas with independent synthetic journal scenarios.

No VBA execution is claimed; all synthetic records stay outside delivered files.
"""
import datetime as dt,json,math,subprocess,tempfile
from pathlib import Path
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
SOURCE=ROOT/'output/POS_Murita_Crepes.xlsx'
PROFILE=json.loads((ROOT/'data/profile.json').read_text())
MR={m['code']:i+7 for i,m in enumerate(PROFILE['materials'])}
CHECKS=[]
def date(s):return dt.datetime.fromisoformat(s)
def check(label,actual,expected):
    if isinstance(expected,(float,int)) and not isinstance(expected,bool):
        assert isinstance(actual,(float,int)) and math.isclose(actual,expected,abs_tol=1e-5,rel_tol=1e-8),(label,actual,expected)
    else:assert actual==expected,(label,actual,expected)
    CHECKS.append({'check':label,'result':'PASS','actual':actual})
def write(s,r,vals):
    for c,v in enumerate(vals,1):
        if v is not None:s.cell(r,c,v)
def stock(w,code):return w['BAHAN'].cell(MR[code],12).value
def no_errors(name,w):
    e=[]
    for s in w:
        for row in s:
            for c in row:
                if c.data_type=='e':e.append(f'{s.title}!{c.coordinate}={c.value}')
    check(name+' / no formula errors',e,[])

DRY_WEIGHT=sum(m.get('recipe_weight',0) for m in PROFILE['materials'])
DRY_COST=sum(m.get('recipe_weight',0)*m['price']/m['pack'] for m in PROFILE['materials'][:7])
WET_PER=1000*(1+.065+.83)/30
BASE_COST=(DRY_COST/DRY_WEIGHT*1000+65*30)/30+5.63*10000/490+4*22000/360+.25*100
HPP1=BASE_COST+15*24
HPP2=BASE_COST+15*14.4

def fixture():
    w=load_workbook(SOURCE)
    w['BAHAN'].cell(MR['SKM'],5,10000)
    w['REFERENSI']['B16']='YA';w['REFERENSI']['B17']=5000
    w['DASHBOARD']['B5']=date('2026-10-01');w['DASHBOARD']['E5']=date('2026-10-31')
    write(w['PELANGGAN'],7,['UJI-CUST','SIMULASI PELANGGAN',None,'YA',date('2026-10-01'),'DATA UJI SINTETIS'])
    g=w['GERAK BAHAN'];r=7
    def move(kind,code,qty,day='2026-10-23',doc='UJI',nota=None,amount=None,method=None):
        nonlocal r
        write(g,r,[f'G-UJI-{r}',date(day),doc,kind,code,qty,amount,method,nota,'AKTIF','SIMULASI']);r+=1
    for m in PROFILE['materials']:
        if m['track']=='YA' and m['group']!='ANTARA':move('AWAL',m['code'],10000)
    for m in PROFILE['materials'][:7]:move('PRODUKSI PAKAI',m['code'],1000*m['recipe_weight']/DRY_WEIGHT,doc='UJI-PREMIX')
    move('PRODUKSI MASUK','PREMIX',1000,doc='UJI-PREMIX')
    for code,qty in [('PREMIX',1000),('BTR',65),('WTR',830)]:move('PRODUKSI PAKAI',code,qty,doc='UJI-WET')
    move('PRODUKSI MASUK','WET',1895,doc='UJI-WET')
    move('BELI','TOP01',100,day='2026-10-24',amount=2400,method='Tunai')
    for nr,id,day,method,discount,received,points,entries in [
        (7,'NOTA-UJI-1','2026-10-24','Tunai',1000,20000,2,[('MC001',2,HPP1),('MC002',1,HPP2)]),
        (8,'NOTA-UJI-2','2026-10-25','QRIS',0,5000,1,[('MC001',1,HPP1)])]:
        write(w['NOTA'],nr,[id,date(day),23/24,'UJI-CUST',method,discount,received,'AKTIF','SIMULASI'])
        w['NOTA'].cell(nr,14,points)
        for sku,qty,hpp in entries:
            ir=7+sum(bool(w['ITEM'].cell(x,1).value) for x in range(7,10))
            write(w['ITEM'],ir,[f'I-UJI-{ir}',id,sku,qty,5000,hpp,'LENGKAP','SIMULASI'])
        total=sum(e[1] for e in entries)
        usage={'WET':total*WET_PER,'SKM':total*5.63,'GAS':total*4,'PACK':total*.25,'TOP01':sum(q for sku,q,h in entries if sku=='MC001')*15,'TOP02':sum(q for sku,q,h in entries if sku=='MC002')*15}
        for code,qty in usage.items():
            if qty:move('JUAL PAKAI',code,qty,day,doc=id,nota=id)
    for r,kind,amount in [(7,'BIAYA USAHA',1000),(8,'PRIVE',2000),(9,'MODAL MASUK',3000),(10,'PENDAPATAN LAIN',500)]:write(w['KEUANGAN'],r,['K-UJI-'+str(r),date('2026-10-24'),kind,amount,'Tunai','SIMULASI','AKTIF'])
    write(w['POIN'],7,['P-UJI',date('2026-10-25'),'UJI-CUST','TUKAR',1,'SIMULASI HADIAH','AKTIF'])
    w['TUTUP KASIR']['B5']=date('2026-10-24');w['TUTUP KASIR']['B7']=10000;w['TUTUP KASIR']['E7']=22100
    w['STRUK']['B5']='NOTA-UJI-1';w['KASIR']['B13']='MC034';w['KASIR']['D13']=1
    return w

def main(report_path=None, fixture_hook=None, validation_hook=None, calculator=None):
    structure=load_workbook(SOURCE)
    check('20 sheets',len(structure.sheetnames),20)
    check('34 source menus',sum(bool(structure['MENU'].cell(r,1).value) for r in range(7,57)),34)
    check('receipt print area',str(structure['STRUK'].print_area),"'STRUK'!$A$7:$F$35")
    check('VBA separate, not falsely embedded',structure.vba_archive,None)
    with tempfile.TemporaryDirectory(prefix='murita-tests-') as td:
        td=Path(td);out=td/'calc';out.mkdir();files=[]
        for name in ['normal','incomplete','changed_masters','yield28','historical','void','invalid','missing_usage','too_many_items']:
            w=fixture()
            if name=='incomplete':
                w['BAHAN'].cell(MR['SKM'],5).value=None
                for r in range(7,10):w['ITEM'].cell(r,6).value=None;w['ITEM'].cell(r,7,'BELUM LENGKAP')
            if name=='changed_masters':
                w['MENU']['C7']=9000;w['MENU']['G7']='NONAKTIF';w['REFERENSI']['B17']=100
                for c,v in [('E7',10),('F7',12),('G7',14),('H7',2)]:w['KALIBRASI'][c]=v
            if name=='yield28':w['ADONAN']['B7']=28
            if name=='historical':w['DASHBOARD']['E5']=date('2026-10-24')
            if name=='void':w['NOTA']['H7']='VOID'
            if name=='invalid':
                w['NOTA']['F7']=999999
                write(w['GERAK BAHAN'],100,['BAD-Q',date('2026-10-26'),'UJI','RUSAK','TOP01',-1,None,None,None,'AKTIF','SIMULASI'])
                write(w['GERAK BAHAN'],101,['NEGATIVE',date('2026-10-26'),'UJI','RUSAK','TOP01',999999,None,None,None,'AKTIF','SIMULASI'])
                for r in [102,103]:write(w['GERAK BAHAN'],r,['DUP',date('2026-10-26'),'UJI','AWAL','TOP01',100,None,None,None,'AKTIF','SIMULASI'])
                w['ITEM']['E9']=0
                write(w['POIN'],8,['P-NEG',date('2026-10-26'),'UJI-CUST','TUKAR',10,'SIMULASI','AKTIF'])
            if name=='missing_usage':
                for r in range(7,70):
                    if w['GERAK BAHAN'].cell(r,9).value=='NOTA-UJI-1':w['GERAK BAHAN'].cell(r,10,'VOID')
            if name=='too_many_items':
                for r in range(10,21):write(w['ITEM'],r,['EXTRA-'+str(r),'NOTA-UJI-1','MC001',1,5000,HPP1,'LENGKAP','SIMULASI'])
            if fixture_hook:fixture_hook(name,w)
            dest=td/(name+'.xlsx');w.save(dest);files.append(dest)
        files.append(SOURCE)
        cmd=['soffice','--headless','--convert-to','xlsx','--outdir',str(out),*[str(p) for p in files]]
        result=(calculator or subprocess.run)(cmd,capture_output=True,text=True,timeout=900)
        assert result.returncode==0,result.stdout+'\n'+result.stderr
        check('all files recalculated',len(list(out.glob('*.xlsx'))),len(files))
        loaded={p.stem:load_workbook(out/p.name,data_only=True) for p in files}
        for name,w in loaded.items():no_errors(name,w)
        w=loaded['normal']
        for label,actual,expected in [
          ('dry weight sums all seven ingredients',w['PREMIKS']['B18'].value,2908.5),
          ('dry cost uses actual ingredient prices',w['PREMIKS']['B19'].value,56731.25),
          ('unit premix cost',w['PREMIKS']['B20'].value,DRY_COST/DRY_WEIGHT),
          ('wet target output',w['ADONAN']['B13'].value,1895),
          ('wet portion for 30 crepes',w['ADONAN']['B15'].value,WET_PER),
          ('SKM known cost included in HPP',w['MENU']['H7'].value,HPP1),
          ('recipe HPP complete after actual SKM input',w['MENU']['I7'].value,'LENGKAP'),
          ('34th menu accepted beyond legacy 20',w['KASIR']['E13'].value,10000),
          ('23:00 accepted',w['NOTA']['O7'].value,'OK'),
          ('cash change correct',w['NOTA']['L7'].value,6000),
          ('net sales after one receipt discount',w['DASHBOARD']['A8'].value,19000),
          ('receipt count independent of item count',w['DASHBOARD']['E8'].value,2),
          ('four items sold',w['DASHBOARD']['I8'].value,4),
          ('business expenses exclude prive',w['DASHBOARD']['A13'].value,1000),
          ('stock purchases reported separately',w['DASHBOARD']['E13'].value,2400),
          ('contribution avoids double counting purchases/prive',w['DASHBOARD']['I13'].value,19000-(3*HPP1+HPP2)-1000),
          ('raw materials consumed only at premix production',stock(w,'RAW01'),10000-1000*1000/DRY_WEIGHT),
          ('premix fully transferred into wet dough',stock(w,'PREMIX'),0),
          ('wet stock consumed once by sales',stock(w,'WET'),round(1895-4*WET_PER,6)),
          ('butter consumed only at wet production',stock(w,'BTR'),9935),
          ('topping stock accounts for purchase and four crepes',stock(w,'TOP01'),10055),
          ('second topping usage separate',stock(w,'TOP02'),9985),
          ('SKM quantity consumed despite different cost unit',stock(w,'SKM'),9977.48),
          ('CRM visits per receipt',w['PELANGGAN']['G7'].value,2),
          ('CRM net spend',w['PELANGGAN']['H7'].value,19000),
          ('CRM latest actual visit date',w['PELANGGAN']['I7'].value,date('2026-10-25')),
          ('loyalty earned per receipt',w['PELANGGAN']['J7'].value,3),
          ('loyalty redeemed',w['PELANGGAN']['K7'].value,-1),
          ('loyalty balance',w['PELANGGAN']['L7'].value,2),
          ('cash drawer excludes QRIS and gross cash tender',w['TUTUP KASIR']['B14'].value,22100),
          ('cash drawer reconciles',w['TUTUP KASIR']['B15'].value,0),
          ('receipt printable',w['STRUK']['A10'].value,'SIAP CETAK'),
          ('receipt first menu quantity',w['STRUK']['D13'].value,2),
          ('receipt second menu quantity',w['STRUK']['D14'].value,1),
          ('receipt total',w['STRUK']['F28'].value,14000),
        ]:check(label,actual,expected)
        w=loaded['incomplete'];check('missing SKM does not fabricate full margin',w['DASHBOARD']['I13'].value,'HPP BELUM LENGKAP');check('missing HPP still permits sales recording',w['DASHBOARD']['A8'].value,19000)
        w=loaded['changed_masters']
        for label,actual,expected in [('three sampled spoon weights averaged',w['KALIBRASI']['I7'].value,12),('spoon calibration replaces standard usage',w['KALIBRASI']['J7'].value,24),('calibration mode',w['KALIBRASI']['K7'].value,'TERKALIBRASI'),('changed prices do not alter sales snapshot',w['DASHBOARD']['A8'].value,19000),('changed recipe does not alter stock snapshot',stock(w,'TOP01'),10055),('changed loyalty rules do not alter past points',w['PELANGGAN']['J7'].value,3),('retiring menu does not invalidate history',w['NOTA']['O7'].value,'OK')]:check(label,actual,expected)
        w=loaded['yield28'];check('yield 28 adjusts wet grams per portion',w['ADONAN']['B15'].value,1895/28);check('yield change leaves posted consumption frozen',stock(w,'WET'),round(1895-4*WET_PER,6))
        w=loaded['historical'];check('historical sales cutoff',w['DASHBOARD']['A8'].value,14000);check('historical visits cutoff',w['PELANGGAN']['G7'].value,1);check('historical points cutoff',w['PELANGGAN']['L7'].value,2);check('historical wet stock cutoff',stock(w,'WET'),1705.5)
        w=loaded['void'];check('VOID removes net sales',w['DASHBOARD']['A8'].value,5000);check('VOID removes CRM visit',w['PELANGGAN']['G7'].value,1);check('VOID restores stock consumption',stock(w,'TOP01'),10085);check('VOID reverses earned points',w['PELANGGAN']['J7'].value,1);check('VOID receipt cannot print',w['STRUK']['A10'].value,'NOTA VOID / TIDAK VALID')
        w=loaded['invalid'];check('discount above gross caught',w['NOTA']['O7'].value,'CEK DISKON');check('negative movement quantity caught',w['GERAK BAHAN']['N100'].value,'CEK QTY');check('negative stock caught',w['GERAK BAHAN']['O101'].value<0,True);check('duplicate stock IDs caught',w['GERAK BAHAN']['N102'].value,'ID GANDA');check('invalid price caught',w['ITEM']['M9'].value,'CEK HARGA');check('negative points caught',w['POIN']['J8'].value<0,True)
        w=loaded['missing_usage'];check('missing stock usage caught before receipt accepted',w['NOTA']['O7'].value,'ISI PEMAKAIAN');check('incomplete receipt excluded from net sales',w['NOTA']['P7'].value,0)
        w=loaded['too_many_items'];check('overlong manual note detected',w['NOTA']['O7'].value,'MAKS 12 ITEM');check('overlong receipt not silently truncated',w['STRUK']['A10'].value,'NOTA VOID / TIDAK VALID')
        w=loaded[SOURCE.stem]
        check('operational workbook has no fake notes',sum(bool(w['NOTA'].cell(r,1).value) for r in range(7,1007)),0)
        check('operational workbook has no customer data',sum(bool(w['PELANGGAN'].cell(r,1).value) for r in range(7,307)),0)
        check('SKM purchase price deliberately blank',w['BAHAN'].cell(MR['SKM'],5).value,None)
        check('loyalty disabled without invented policy',w['REFERENSI']['B16'].value,'TIDAK')
        check('no invented points threshold',w['REFERENSI']['B17'].value,None)
        check('default source HPP incomplete honestly',w['MENU']['I7'].value,'BELUM LENGKAP')
        check('default premix validates',w['PREMIKS']['B21'].value,'OK')
        if validation_hook:validation_hook(loaded)
        SOURCE.write_bytes((out/SOURCE.name).read_bytes())
    report={'passed':len(CHECKS),'failed':0,'engine':'LibreOffice headless 26.8 alpha; real workbook formula recalculation','vba_runtime_tested':False,'excel_desktop_tested':False,'checks':CHECKS}
    (report_path or ROOT/'output/hasil_validasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str))
    print(json.dumps({'passed':len(CHECKS),'failed':0}))

if __name__=='__main__':main()
