"""Offline transfer exchange for the three .xlsx workbooks. No network calls.

Imports always produce a NEW file; a failed batch never modifies the source.
Use the supplied Excel VBA for .xlsm files so form controls are not stripped.
"""
from __future__ import annotations
import argparse,csv,datetime as dt,json,math,re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

HEADERS=['schema','kind','transfer_id','shipped_on','received_on','origin','destination','material','unit','quantity_shipped','quantity_received','request_id','note']
LOCATIONS={'MEDOKAN','KARAH','GUDANG'}
IDENT=re.compile(r'^[A-Za-z0-9_-]{1,120}$')
SIGNS={'AWAL':1,'BELI':1,'TERIMA':1,'PRODUKSI MASUK':1,'KOREKSI MASUK':1,'KIRIM':-1,'PAKAI':-1,'RUSAK':-1,'KOREKSI KELUAR':-1,'PRODUKSI PAKAI':-1}

def identifier(x,optional=False):
    x=str(x or '')
    if optional and not x:return ''
    if not IDENT.fullmatch(x):raise ValueError('ID tidak valid: '+x)
    return x

def day(x):
    if isinstance(x,dt.datetime):x=x.date()
    if isinstance(x,dt.date):return x
    try:return dt.date.fromisoformat(str(x))
    except (ValueError,TypeError):raise ValueError('Tanggal harus ISO yyyy-mm-dd: '+str(x))

def num(x):
    try:n=float(x)
    except (ValueError,TypeError):raise ValueError('Kuantitas harus angka: '+str(x))
    if not math.isfinite(n) or n<=0:raise ValueError('Kuantitas harus positif dan finite.')
    return n

def same(a,b):return math.isclose(float(a),float(b),abs_tol=1e-6,rel_tol=0)

def load(path):
    path=Path(path)
    if path.suffix.lower()!='.xlsx':raise ValueError('CLI hanya menerima .xlsx; gunakan VBA Excel untuk .xlsm.')
    w=load_workbook(path)
    if w['CONFIG']['B6'].value not in LOCATIONS or w['CONFIG']['B7'].value!='PF1':raise ValueError('Identitas/schema workbook tidak dikenal.')
    return w

def materials(w):
    m={}
    for row in w['BAHAN'].iter_rows(min_row=7,max_row=86,values_only=True):
        if row[0]:
            code=identifier(row[0])
            if code in m:raise ValueError('Kode bahan ganda: '+code)
            m[code]=(str(row[3]),row[7]=='YA')
    return m

def movements(w):
    for row in w['MUTASI'].iter_rows(min_row=7,max_row=4006,values_only=True):
        if row[0]:yield row
        elif any(v is not None and v!='' for v in row[:14]):raise ValueError('Ada baris MUTASI parsial tanpa ID.')

def transfer_map(w,kind):
    result={}
    for row in movements(w):
        if row[2] in ['KIRIM','TERIMA'] and row[7]:
            key=identifier(row[7])
            if key in result:raise ValueError('ID transfer ganda: '+key)
            result[key]=row
    return {k:r for k,r in result.items() if r[2]==kind}

def check_stock(w):
    known=materials(w);notes={}
    if 'NOTA' in w:
        notes={str(r[0]):r[7] for r in w['NOTA'].iter_rows(min_row=7,max_row=1006,values_only=True) if r[0]}
    rows=[];seen=set()
    for i,row in enumerate(movements(w)):
        key=identifier(row[0])
        if key in seen:raise ValueError('ID gerak ganda.')
        seen.add(key)
        if row[12]=='VOID':continue
        if row[12]!='AKTIF' or row[2] not in SIGNS:raise ValueError('Jenis/status MUTASI tidak valid.')
        material=identifier(row[3])
        if material not in known:raise ValueError('Bahan tidak dikenal.')
        when=day(row[1])
        if when>dt.date.today():raise ValueError('Gerak stok di masa depan.')
        qty=num(row[4])
        if row[11] and notes.get(str(row[11]))=='VOID':continue
        rows.append((when,i,material,SIGNS[row[2]]*qty))
    balance={}
    for _,_,code,qty in sorted(rows):
        balance[code]=balance.get(code,0)+qty
        if known[code][1] and balance[code]<-1e-6:raise ValueError('Stok historis negatif: '+code)

def validate_record(record,kind,known):
    if record['schema']!='PF1' or record['kind']!=kind:raise ValueError('Schema/jenis CSV berbeda.')
    identifier(record['transfer_id']);identifier(record['request_id'],optional=True)
    if record['origin'] not in LOCATIONS or record['destination'] not in LOCATIONS or record['origin']==record['destination']:raise ValueError('Asal/tujuan salah.')
    code=identifier(record['material'])
    if code not in known or record['unit']!=known[code][0]:raise ValueError('Bahan atau satuan berbeda.')
    if day(record['shipped_on'])>dt.date.today():raise ValueError('Tanggal kirim di masa depan.')
    num(record['quantity_shipped'])
    if kind=='SHIPMENT' and (record['received_on'] or record['quantity_received']):raise ValueError('CSV kirim tidak boleh membawa penerimaan fiktif.')
    if kind=='RECEIPT':
        if not day(record['shipped_on'])<=day(record['received_on'])<=dt.date.today():raise ValueError('Tanggal terima tidak valid.')
        num(record['quantity_received'])

def read_csv(path):
    with Path(path).open(encoding='utf-8-sig',newline='') as f:
        reader=csv.DictReader(f)
        if reader.fieldnames!=HEADERS:raise ValueError('Header CSV PF1 tidak sesuai.')
        rows=list(reader)
    if not rows:raise ValueError('CSV tidak berisi transaksi.')
    if any(None in r or any(v is None for v in r.values()) for r in rows):raise ValueError('Jumlah kolom CSV tidak sesuai.')
    keys=[r['transfer_id'] for r in rows]
    if len(set(keys))!=len(keys):raise ValueError('ID transfer ganda dalam CSV.')
    return rows

def put(sheet,row,values):
    for c,v in values.items():
        cell=sheet.cell(row,c,v)
        if isinstance(v,str):cell.data_type='s'  # CSV text is never executed as an Excel formula.

def free(sheet,last,input_cols):
    for r in range(7,last+1):
        if not sheet.cell(r,1).value:
            if any(sheet.cell(r,c).value is not None for c in input_cols):raise ValueError('Ada baris input parsial: '+sheet.title)
            return r
    raise ValueError('Tabel penuh: '+sheet.title)

def save_new(w,source,output):
    source,output=Path(source).resolve(),Path(output).resolve()
    if source==output or output.exists() or output.suffix.lower()!='.xlsx':raise ValueError('Pilih file keluaran .xlsx BARU, berbeda dari sumber.')
    output.parent.mkdir(parents=True,exist_ok=True)
    w.calculation=CalcProperties(calcId=0,fullCalcOnLoad=True,forceFullCalc=True)
    from io import BytesIO
    buf=BytesIO();w.save(buf)
    with output.open('xb') as f:f.write(buf.getvalue())

def export_records(source,output,kind,peer=None):
    w=load(source);known=materials(w);role=w['CONFIG']['B6'].value;check_stock(w)
    expected='KIRIM' if kind=='SHIPMENT' else 'TERIMA';rows=[]
    for key,row in transfer_map(w,expected).items():
        if row[12]!='AKTIF':continue
        if peer and row[9 if kind=='SHIPMENT' else 8]!=peer:continue
        record=dict(zip(HEADERS,['PF1',kind,key,day(row[1] if expected=='KIRIM' else row[19]).isoformat(),'' if expected=='KIRIM' else day(row[1]).isoformat(),str(row[8]),str(row[9]),str(row[3]),known[row[3]][0],format(num(row[4] if expected=='KIRIM' else row[20]),'.15g'),'' if expected=='KIRIM' else format(num(row[4]),'.15g'),str(row[10] or ''),str(row[13] or '')]))
        validate_record(record,kind,known)
        if role!=record['origin' if kind=='SHIPMENT' else 'destination']:raise ValueError('Lokasi file tidak cocok dengan transfer.')
        rows.append(record)
    if not rows:raise ValueError('Tidak ada transfer aktif untuk diekspor.')
    if len({r['destination' if kind=='SHIPMENT' else 'origin'] for r in rows})>1:raise ValueError('Pilih --peer agar CSV hanya untuk satu lokasi lawan.')
    output=Path(output)
    with output.open('x',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=HEADERS);writer.writeheader();writer.writerows(rows)
    return {'exported':len(rows),'file':str(output)}

def import_shipments(source,csv_path,output,received_on,received=None):
    w=load(source);known=materials(w);role=w['CONFIG']['B6'].value
    records=read_csv(csv_path);existing=transfer_map(w,'TERIMA');received=received or {};date=day(received_on)
    if date>dt.date.today():raise ValueError('Tanggal penerimaan di masa depan.')
    if set(received)-{r['transfer_id'] for r in records}:raise ValueError('Override qty memakai ID yang tidak ada di CSV.')
    inserted=skipped=0
    for x in records:
        validate_record(x,'SHIPMENT',known)
        if x['destination']!=role or day(x['shipped_on'])>date:raise ValueError('Tujuan atau tanggal terima salah.')
        key=x['transfer_id'];qty=num(received.get(key,x['quantity_shipped']));shipped=num(x['quantity_shipped'])
        if key in existing:
            r=existing[key]
            if not(r[12]=='AKTIF' and r[8]==x['origin'] and r[9]==role and r[3]==x['material'] and (key not in received or same(r[4],qty)) and day(r[19])==day(x['shipped_on']) and same(r[20],shipped) and str(r[10] or '')==x['request_id']):raise ValueError('ID penerimaan sudah ada dengan data berbeda: '+key)
            skipped+=1;continue
        if any(r[7]==key for r in movements(w)):raise ValueError('ID transfer sudah dipakai oleh jenis lain.')
        r=free(w['MUTASI'],4006,list(range(1,15))+[20,21,22])
        put(w['MUTASI'],r,{1:'RX-'+key,2:dt.datetime.combine(date,dt.time()),3:'TERIMA',4:x['material'],5:qty,8:key,9:x['origin'],10:role,11:x['request_id'],13:'AKTIF',14:x['note'],20:dt.datetime.combine(day(x['shipped_on']),dt.time()),21:shipped})
        inserted+=1
    check_stock(w);save_new(w,source,output)
    return {'inserted':inserted,'skipped':skipped,'file':str(output)}

def import_receipts(source,csv_path,output):
    w=load(source);known=materials(w);role=w['CONFIG']['B6'].value;ships=transfer_map(w,'KIRIM')
    existing={str(r[0]):r for r in w['KONFIRMASI'].iter_rows(min_row=7,max_row=2006,values_only=True) if r[0]}
    if len(existing)!=sum(bool(w['KONFIRMASI'].cell(r,1).value) for r in range(7,2007)):raise ValueError('ID konfirmasi ganda.')
    inserted=skipped=0
    for x in read_csv(csv_path):
        validate_record(x,'RECEIPT',known);key=x['transfer_id']
        if x['origin']!=role or key not in ships:raise ValueError('Konfirmasi bukan pengiriman lokasi ini.')
        ship=ships[key]
        if not(ship[12]=='AKTIF' and ship[8]==role and ship[9]==x['destination'] and ship[3]==x['material'] and same(ship[4],x['quantity_shipped']) and day(ship[1])==day(x['shipped_on']) and str(ship[10] or '')==x['request_id']):raise ValueError('Konfirmasi tidak cocok dengan data kirim.')
        values=[key,dt.datetime.combine(day(x['received_on']),dt.time()),role,x['destination'],x['material'],num(x['quantity_shipped']),num(x['quantity_received']),x['unit'],dt.datetime.combine(day(x['shipped_on']),dt.time()),x['note']]
        if key in existing:
            old=existing[key]
            if not all((day(a)==day(v) if i in [1,8] else same(a,v) if i in [5,6] else str(a or '')==str(v or '')) for i,(a,v) in enumerate(zip(old[:10],values))):raise ValueError('Konfirmasi ID sama memiliki data berbeda.')
            skipped+=1;continue
        r=free(w['KONFIRMASI'],2006,range(1,11));put(w['KONFIRMASI'],r,{i+1:v for i,v in enumerate(values)});inserted+=1
    save_new(w,source,output)
    return {'inserted':inserted,'skipped':skipped,'file':str(output)}

def main():
    p=argparse.ArgumentParser();p.add_argument('command',choices=['export-shipments','import-shipments','export-receipts','import-receipts']);p.add_argument('workbook',type=Path);p.add_argument('--csv',type=Path,required=True);p.add_argument('--output',type=Path);p.add_argument('--peer',choices=sorted(LOCATIONS));p.add_argument('--received-on');p.add_argument('--received',type=Path,help='JSON {transfer_id: actual_received_qty}')
    a=p.parse_args()
    if a.command.startswith('export'):result=export_records(a.workbook,a.csv,'SHIPMENT' if a.command=='export-shipments' else 'RECEIPT',a.peer)
    else:
        if not a.output:p.error('--output wajib dan harus file baru')
        if a.command=='import-shipments':
            if not a.received_on:p.error('--received-on wajib')
            result=import_shipments(a.workbook,a.csv,a.output,a.received_on,json.loads(a.received.read_text()) if a.received else None)
        else:result=import_receipts(a.workbook,a.csv,a.output)
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':main()
