"""Cashier and cash reconciliation views adapted from the supplied sate POS.

Keep the existing actual-piece journals and HPP snapshots as the source of truth.
The flattened menu is only a cashier selector, not another inventory balance.
"""
from layout import rng, look
from reference_ux import cashier_view, closing_view

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
    wrap_label = b.w.add_format({'font_name':'Calibri','font_size':11,'font_color':'#34291F','bg_color':'#F2F2F2','text_wrap':True,'valign':'vcenter'})
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
    s = cashier_view(b, profile, today, stock_at, menu_lookup, rng, look)
    f = b.f
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
    for cell,caption,macro in [('I62','SIMPAN BELANJA','SimpanBelanja'),('I71','SIMPAN KAS','SimpanKas')]:
        s.insert_button(cell,{'caption':caption,'macro':macro,'width':170,'height':27})


def add_closing(b, today):
    return closing_view(b,today)
