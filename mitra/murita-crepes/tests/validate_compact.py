"""Run the same business fixtures against the consolidated physical workbook."""
import sys, re
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
import validate_workbook as original
from compact_layout import BLOCKS, ORDER, BookView
from build_compact import DEST, OUT
from openpyxl import load_workbook
from pos_ui import FIRST_CLOSE,LAST_CLOSE
from recalculate_parallel import calculate
from check_cash_cases import verify as verify_cash_cases

SR={m['code']:i+8 for i,m in enumerate(original.PROFILE['materials'])}

def fixture_hook(name,w):
    for code,limit in [('TOP01',11000),('TOP02',0)]:w['BAHAN'].cell(original.MR[code],9,limit)
    t=w.raw['TUTUP KASIR']
    if name=='historical':t['I11']=22000
    if name=='invalid':
        t['C5']=original.date('2026-10-27');t['I8']=0;t['I11']=0
        original.write(w.raw['JURNAL'],FIRST_CLOSE,['BAD-CLOSE',original.date('2026-10-24'),'abc','SIMULASI',0,0,0,0,0,0,0,0,0,0,'LENGKAP','SIMULASI',original.date('2026-10-24')])
    if name=='changed_masters':
        for off,rev,physical in [(0,1,22000),(1,2,22100)]:
            original.write(w.raw['JURNAL'],FIRST_CLOSE+off,[f'UJI-CLOSE-{rev}',original.date('2026-10-24'),rev,'SIMULASI',10000,14000,2400,500,22100,physical,physical-22100,14000,1,3,'LENGKAP','ALASAN SIMULASI',original.date('2026-10-24')])

def validate_pos(loaded):
    check=original.check;w=loaded['normal'];t=w.raw['TUTUP KASIR'];s=w.raw['STOK']
    for label,actual,expected in [
        ('POS closing notes only selected day',t['C8'].value,1),
        ('POS closing items only selected day',t['C10'].value,3),
        ('POS closing gross',t['C11'].value,15000),('POS closing single receipt discount',t['C12'].value,1000),
        ('POS closing netto',t['C13'].value,14000),('POS closing cash',t['F8'].value,14000),
        ('POS closing QRIS excluded from cash day',t['F9'].value,0),
        ('POS closing cash purchase',t['F12'].value,2400),('POS closing business expense',t['F13'].value,1000),
        ('POS closing prive',t['F14'].value,2000),('POS closing modal and other cash',t['F15'].value,3500),
        ('POS closing drawer',t['I10'].value,22100),('POS closing status zero difference',t['H14'].value,'KAS COCOK - TIDAK ADA SELISIH'),
        ('POS closing accepts 23:00',t['F16'].value,'23:00'),
        ('POS closing top menu',t['H19'].value,original.PROFILE['menus'][0]['name']),
        ('POS closing top menu quantity',t['I19'].value,2),('POS closing second menu quantity',t['I20'].value,1),
        ('POS closing no phantom third menu',t['H21'].value,None),
        ('POS stock raw consumption',s[f'F{SR["RAW01"]}'].value,1000*1000/original.DRY_WEIGHT),
        ('POS stock premix production received',s[f'E{SR["PREMIX"]}'].value,1000),
        ('POS stock premix used once',s[f'F{SR["PREMIX"]}'].value,1000),
        ('POS stock wet received',s[f'E{SR["WET"]}'].value,1895),
        ('POS stock wet consumed once',s[f'F{SR["WET"]}'].value,4*original.WET_PER),
        ('POS stock red zero balance',s[f'L{SR["PREMIX"]}'].value,'MERAH'),
        ('POS stock yellow threshold',s[f'L{SR["TOP01"]}'].value,'KUNING'),
        ('POS stock green threshold',s[f'L{SR["TOP02"]}'].value,'HIJAU'),
        ('POS stock no false green without threshold',s[f'L{SR["RAW01"]}'].value,'ISI BATAS'),
        ('POS stock gas not physically tracked',s[f'L{SR["GAS"]}'].value,'TIDAK DILACAK'),
        ('POS stock reorder rounds to pack',s[f'M{SR["TOP01"]}'].value,1000),
        ('POS stock estimated purchase cost',s['K5'].value,24000),
    ]:check(label,actual,expected)
    for code in SR:
        if next(m for m in original.PROFILE['materials'] if m['code']==code)['track']=='YA':
            check('POS stock agrees with ledger '+code,s[f'H{SR[code]}'].value,original.stock(w,code))
    t=loaded['invalid'].raw['TUTUP KASIR']
    check('POS close actual zero cash is valid',t['I12'].value,0)
    check('POS close zero cash status',t['H14'].value,'KAS COCOK - TIDAK ADA SELISIH')
    check('POS malformed archive revision does not produce formula error',loaded['invalid'].raw['JURNAL'].cell(FIRST_CLOSE,18).value,'CEK INPUT')
    t=loaded['historical'].raw['TUTUP KASIR'];check('POS cash shortage amount',t['I12'].value,-100);check('POS shortage message',t['H14'].value.startswith('KAS KURANG'),True)
    t=loaded['void'].raw['TUTUP KASIR'];check('POS closing excludes VOID',t['C13'].value,0);check('POS closing counts VOID',t['C17'].value,1)
    t=loaded['incomplete'].raw['TUTUP KASIR'];check('POS incomplete HPP blocks margin claim',t['C16'].value,'HPP BELUM LENGKAP')
    check('POS incomplete receipt blocks closing readiness',loaded['missing_usage'].raw['TUTUP KASIR']['C18'].value,1)
    t=loaded['changed_masters'].raw['TUTUP KASIR'];check('POS archive revision count',t['C25'].value,2)
    for i in range(2):check('POS archive snapshot validation '+str(i+1),loaded['changed_masters'].raw['JURNAL'].cell(FIRST_CLOSE+i,18).value,'OK')
    w=loaded[original.SOURCE.stem];t=w.raw['TUTUP KASIR']
    check('POS closing missing opening cash flagged',t['H14'].value,'ISI MODAL AWAL (0 BILA TIDAK ADA)')
    check('POS new stock distinguishes no opening inventory',w.raw['STOK']['L8'].value,'BELUM AWAL')
    check('POS output archive empty',sum(bool(w.raw['JURNAL'].cell(r,1).value) for r in range(FIRST_CLOSE,LAST_CLOSE+1)),0)

def load_aliases(*args,**kwargs):return BookView(load_workbook(*args,**kwargs))

def main():
    w=load_workbook(DEST)
    assert w.sheetnames==ORDER
    assert sum(s.sheet_state=='visible' for s in w)==8
    assert len([t for s in w for t in s.tables.values()])==12
    assert not w._external_links
    # All transformed formulas refer to surviving sheets, including chart refs,
    # names and dropdown definitions. No stale source tab may remain.
    refs=re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!")
    for s in w:
        for row in s:
            for c in row:
                if c.data_type=='f':
                    for quoted,plain in refs.findall(c.value):
                        assert (quoted or plain) in ORDER,(s.title,c.coordinate,c.value)
    raw_check=original.check
    def compact_check(label,actual,expected):
        if label=='20 sheets':label='9 sheets / 8 visible';expected=9
        if label=='receipt print area':expected="'KASIR'!$A$53:$F$81"
        return raw_check(label,actual,expected)
    original.SOURCE=DEST
    original.load_workbook=load_aliases
    original.check=compact_check
    original.main(report_path=OUT/'hasil_validasi.json',fixture_hook=fixture_hook,validation_hook=validate_pos,calculator=calculate)
    verify_cash_cases(DEST,OUT/'hasil_validasi.json')

if __name__=='__main__':main()
