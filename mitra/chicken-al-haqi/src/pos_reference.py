"""Cashier and cash reconciliation views adapted from the supplied sate POS.

Keep the existing actual-piece journals and HPP snapshots as the source of truth.
The flattened menu is only a cashier selector, not another inventory balance.
"""
from layout import rng, look

MENU_FIRST = 69
MENU_LAST = 88


def menu_range(column):
    return f"'SETUP'!${column}${MENU_FIRST}:${column}${MENU_LAST}"


def menu_lookup(value, column, default='""'):
    return f'IFERROR(INDEX({menu_range(column)},MATCH({value},{menu_range("A")},0)),{default})'


def add_menu_selector(b, profile):
    s, f = b.s['SETUP'], b.f
    s.write_row('A68', ['Menu kasir', 'Kode potongan', 'Rasa', 'Harga / porsi', 'Status'], f['head'])
    variants = [(m, flavor) for m in profile['items'] for flavor in profile['flavors']
                if m['code'] != 'NASI' or flavor['name'] == 'Original']
    wrap_label = b.w.add_format({'font_name':'Calibri','font_size':11,'font_color':'#34291F','bg_color':'#F5F2EC','text_wrap':True,'valign':'vcenter'})
    for r, (item, flavor) in enumerate(variants, MENU_FIRST):
        label = item['name'] if item['code'] == 'NASI' else item['name'] + ' ' + flavor['name']
        s.write_row(r - 1, 0, [label, item['code'], flavor['name']], f['calc'])
        s.write(f'A{r}', label, wrap_label)
        if len(label)>20:s.set_row(r-1,34)
        extra = f'INDEX($B$31:$B$34,MATCH(C{r},$A$31:$A$34,0))'
        b.formula(s, f'D{r}', f'{look(f"B{r}", "SETUP", "C", "0")}+{extra}', 'money')
        b.formula(s, f'E{r}', look(f'B{r}', 'SETUP', 'I'))
    s.merge_range('G68:L72', 'Daftar pilihan kasir mengikuti harga dasar dan tambahan rasa di atas. '
                  'Empat potongan × empat rasa, ditambah nasi tanpa saus. Stok setiap rasa tetap '
                  'memakai potongan yang sama; jangan membuat stok terpisah per rasa.', f['note'])
    b.w.define_name('MenuKasir', '=' + menu_range('A'))
    s.print_area('A1:L88')


def add_cashier(b, profile, today, stock_at):
    s = b.base('KASIR', 'Pilih menu dan jumlah, lalu simpan transaksi. Harga, kembalian, '
               'penjualan harian dan stok potongan mengikuti jurnal yang sama.')
    f = b.f
    for col, width in [('A', 3), ('B', 32), ('C', 21), ('D', 18), ('E', 19),
                       ('F', 21), ('G', 3), ('H', 25), ('I', 22), ('J', 13), ('K', 20), ('L', 3)]:
        s.set_column(f'{col}:{col}', width)
    s.merge_range('B5:F5', '1. INFO NOTA', f['head'])
    for r, label in [(6, 'No Nota'), (7, 'Tanggal'), (8, 'Jam'), (9, 'Nama Kasir'),
                     (10, 'Metode Bayar'), (11, 'Channel'), (12, 'Jumlah Pelanggan')]:
        s.write(f'B{r}', label, f['label'])
    s.write('C6', 'Dibuat saat simpan', f['calc'])
    s.write_datetime('C7', today, f['in_date'])
    s.write('C8', 9 / 24, f['in_time'])
    s.write_blank('C9', None, f['input'])
    s.write('C10', 'Tunai', f['input'])
    s.write('C11', 'Dine-in', f['input'])
    s.write('C12', 1, f['in_qty'])
    b.validate(s, 'C10', '=MetodeBayar')
    b.validate(s, 'C11', '=ChannelPenjualan')
    s.data_validation('C12', {'validate': 'integer', 'criteria': '>', 'value': 0})
    s.merge_range('H5:K5', 'PENJUALAN HARI INI', f['head'])
    for r, label, formula, fmt in [
        (6, 'Omzet berjalan', f'SUMIFS({rng("NOTA", "P")},{rng("NOTA", "B")},$C$7)', 'money'),
        (7, 'Jumlah nota', f'COUNTIFS({rng("NOTA", "B")},$C$7,{rng("NOTA", "H")},"AKTIF",{rng("NOTA", "O")},"OK")', 'qty'),
        (8, 'Item terjual', f'SUMIFS({rng("ITEM", "N")},{rng("ITEM", "I")},$C$7)', 'qty'),
        (9, 'Pelanggan tercatat', f'SUMIFS({rng("NOTA", "R")},{rng("NOTA", "B")},$C$7,{rng("NOTA", "H")},"AKTIF",{rng("NOTA", "O")},"OK")', 'qty'),
        (10, 'Rata-rata per nota', 'IF(I7=0,"",I6/I7)', 'money'),
    ]:
        s.write(f'H{r}', label, f['label'])
        b.formula(s, f'I{r}', formula, fmt, 0 if r != 10 else '')
    s.merge_range('B14:F14', '2. KERANJANG PEMBELI', f['head'])
    s.write_row('B15', ['Menu', 'Jumlah', 'Harga', 'Subtotal', 'Cek Stok'], f['head'])
    for r in range(16, 28):
        s.write_blank(f'B{r}', None, f['input'])
        s.write_blank(f'C{r}', None, f['in_qty'])
        b.formula(s, f'D{r}', f'IF(B{r}="","",{menu_lookup(f"B{r}", "D", "0")})', 'money')
        b.formula(s, f'E{r}', f'IF(B{r}="","",N(C{r})*N(D{r}))', 'money')
        b.formula(s, f'X{r}', menu_lookup(f'B{r}', 'B'))
        b.formula(s, f'Y{r}', menu_lookup(f'B{r}', 'C'))
        balance = stock_at(f'X{r}', '$C$7')
        needed = f'SUMIF($X$16:$X$27,X{r},$C$16:$C$27)'
        status = (f'IF(B{r}="","",IF(X{r}="","CEK MENU",IF({look(f"X{r}", "SETUP", "H")}<>"YA",'
                  f'"TIDAK DILACAK",IF({look(f"X{r}", "STOK", "J")}<>"OK",'
                  f'"BELUM AWAL",IF($C$7<{look(f"X{r}", "STOK", "C", "0")},'
                  f'"BELUM MULAI",IF({balance}<{needed},"STOK KURANG","CUKUP"))))))')
        b.formula(s, f'F{r}', status)
    b.validate(s, 'B16:B27', '=MenuKasir')
    s.data_validation('C16:C27', {'validate': 'integer', 'criteria': '>', 'value': 0})
    s.set_column('N:Y', None, None, {'hidden': True})
    for r, label in [(29, 'Subtotal'), (30, 'Diskon / Promo (Rp)'), (31, 'TOTAL BAYAR'),
                     (32, 'Uang Diterima'), (33, 'Kembalian')]:
        s.write(f'B{r}', label, f['label'])
    b.formula(s, 'C29', 'SUM(E16:E27)', 'money', 0)
    s.write('C30', 0, f['in_money'])
    b.formula(s, 'C31', 'MAX(0,C29-N(C30))', 'money', 0)
    s.write_blank('C32', None, f['in_money'])
    b.formula(s, 'C33', 'IF(OR(C32="",C10<>"Tunai"),"",N(C32)-C31)', 'money')
    s.merge_range('B35:F36', 'Stok memakai pcs/porsi aktual. Rasa memakai stok potongan yang sama. '
                  'Stok BELUM AWAL belum dapat diverifikasi; isi saldo fisik pada STOK.', f['warn'])
    s.merge_range('B38:F38', '3. SIMPAN TRANSAKSI', f['head'])
    s.merge_range('B39:F40', 'Harga dan HPP tersimpan pada nota. Gunakan tombol setelah modul VBA '
                  'dipasang, lalu Ctrl+S untuk menyimpan file.', f['note'])
    s.merge_range('H13:K13', 'STOK POTONGAN', f['head'])
    s.write_row('H14', ['Potongan', 'Saldo pcs', 'Dipesan', 'Status'], f['head'])
    for r, item in enumerate(profile['items'][:4], 15):
        code = '"' + item['code'] + '"'
        s.write(f'H{r}', item['name'], f['calc'])
        balance = stock_at(code, '$C$7')
        opened = f'AND({look(code,"STOK","J")}="OK",$C$7>={look(code,"STOK","C","0")})'
        b.formula(s, f'I{r}', f'IF({opened},{balance},"")', 'qty')
        b.formula(s, f'J{r}', f'SUMIF($X$16:$X$27,{code},$C$16:$C$27)', 'qty', 0)
        b.formula(s, f'K{r}', f'IF(NOT({opened}),"BELUM AWAL",IF(I{r}<J{r},"KURANG",IF(I{r}=0,"HABIS","TERSEDIA")))')
    s.merge_range('H21:K21', 'POTONGAN TERJUAL HARI INI', f['head'])
    s.write_row('H22', ['Potongan / Menu', 'Qty', 'Bruto', ''], f['head'])
    for r, item in enumerate(profile['items'], 23):
        code = '"' + item['code'] + '"'
        s.write(f'H{r}', item['name'], f['calc'])
        b.formula(s, f'I{r}', f'SUMIFS({rng("ITEM","N")},{rng("ITEM","C")},{code},{rng("ITEM","I")},$C$7)', 'qty', 0)
        s.merge_range(f'J{r}:K{r}', '', f['money'])
        b.formula(s, f'J{r}', f'SUMIFS({rng("ITEM","K")},{rng("ITEM","C")},{code},{rng("ITEM","I")},$C$7,{rng("ITEM","N")},">0")', 'money', 0)
    s.merge_range('H30:K30', 'NAVIGASI', f['head'])
    for r, tab, label in [(31, 'HARIAN', 'Dashboard penjualan'), (32, 'NOTA', 'Buku transaksi'),
                           (33, 'STOK', 'Stok potongan'), (34, 'TUTUP KASIR', 'Tutup kasir')]:
        s.write_url(f'H{r}', f"internal:'{tab}'!A1", f['nav'], label)
    s.merge_range('H37:K40', 'Tombol: simpan sebagai .xlsm, impor Modul_ChickenAlHaqi.bas, '
                  'lalu jalankan PasangTombol. Rumus tetap tersedia tanpa macro. '
                  'Belanja dan kas usaha ada di bawah form ini atau pada tab masing-masing.', f['note'])
    s.merge_range('H51:K51', 'BELANJA AYAM / NASI', f['head'])
    fields = [('H53', 'Tanggal', 'I53', 'date'), ('H54', 'Potongan / Nasi', 'I54', 'text'),
              ('H55', 'Kg (opsional)', 'I55', 'qty'), ('H56', 'Pcs/Porsi AKTUAL', 'I56', 'qty'),
              ('H57', 'Nominal total', 'I57', 'money'), ('H58', 'Metode', 'I58', 'text'),
              ('H65', 'Tanggal kas', 'I65', 'date'), ('H66', 'Jenis kas', 'I66', 'text'),
              ('H67', 'Nominal', 'I67', 'money'), ('H68', 'Metode', 'I68', 'text'),
              ('H69', 'Catatan', 'I69', 'text')]
    s.merge_range('H64:K64', 'KAS USAHA', f['head'])
    for cell, label, target, kind in fields:
        s.write(cell, label, f['label'])
        s.write_blank(target, None, f['input' if kind == 'text' else 'in_' + kind])
    for cell in ['I53', 'I65']:
        s.write_datetime(cell, today, f['in_date'])
    for cell, value in [('I58', 'Tunai'), ('I66', 'BIAYA USAHA'), ('I68', 'Tunai')]:
        s.write(cell, value, f['input'])
    for cell, source in [('I54', '=KodeMenu'), ('I58', '=MetodeBayar'),
                          ('I66', '=JenisKas'), ('I68', '=MetodeBayar')]:
        b.validate(s, cell, source)
    s.write('H59', 'Perkiraan pcs', f['label'])
    b.formula(s, 'I59', f'IF(OR(I54="",I55=""),"",N(I55)*{look("I54","SETUP","E","0")})', 'qty')
    s.write('H60', 'Acuan nominal', f['label'])
    b.formula(s, 'I60', 'IF(OR(I54="",I54="NASI",I55=""),"",N(I55)*SETUP!B40)', 'money')
    s.freeze_panes(15, 2)
    s.print_area('A1:L47')


def add_closing(b, today):
    s = b.base('TUTUP KASIR', 'Pilih tanggal, isi modal awal dan uang fisik khusus usaha ayam. '
               'Penjualan non-tunai ditampilkan terpisah.')
    f = b.f
    s.write('B5', 'Tanggal laporan', f['label'])
    s.write_datetime('C5', today, f['in_date'])
    s.merge_range('B7:F7', 'RINGKASAN PENJUALAN', f['head'])
    for r, label, formula, fmt in [
        (8, 'Jumlah nota', "'HARIAN'!E8", 'qty'),
        (9, 'Pcs ayam terjual', "'HARIAN'!I8", 'qty'),
        (10, 'Penjualan netto', "'HARIAN'!A8", 'money'),
        (11, 'Tunai', "'HARIAN'!B18", 'money'),
        (12, 'QRIS', "'HARIAN'!F20", 'money'),
        (13, 'Transfer', "'HARIAN'!F21", 'money'),
        (14, 'Lainnya', "'HARIAN'!F22", 'money'),
    ]:
        s.write(f'B{r}', label, f['label'])
        b.formula(s, f'E{r}', formula, fmt, 0)
    s.merge_range('H7:L7', 'COCOKKAN UANG LACI', f['head'])
    for r, label in [(8, 'Modal awal laci'), (9, 'Penjualan tunai'), (10, 'Belanja tunai'),
                     (11, 'Uang fisik di laci'), (12, 'Kas lain tunai (+/-)'),
                     (14, 'Kas seharusnya'), (15, 'Selisih kas')]:
        s.write(f'H{r}', label, f['label'])
    s.write_blank('K8', None, f['in_money'])
    s.write_blank('K11', None, f['in_money'])
    for cell, formula in [('K9', "'HARIAN'!B18"), ('K10', "'HARIAN'!B19"),
                           ('K12', "'HARIAN'!B20")]:
        b.formula(s, cell, formula, 'money', 0)
    b.formula(s, 'K14', 'IF(ISNUMBER(K8),K8+K9-K10+K12,"")', 'money')
    b.formula(s, 'K15', 'IF(AND(ISNUMBER(K11),ISNUMBER(K14)),K11-K14,"")', 'money')
    s.conditional_format('K15', {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': f['bad']})
    s.merge_range('B19:L22', 'Kas seharusnya = modal awal + penjualan tunai - belanja tunai '
                  '+ kas usaha tunai bertanda. Isi ulang modal dan uang fisik ketika tanggal diganti. '
                  'Form ini menghitung rekap; tidak mengunci periode atau menyimpan saldo per hari. '
                  'Simpan PDF/arsip setelah rekonsiliasi.', f['note'])
    s.print_area('A1:L23')
