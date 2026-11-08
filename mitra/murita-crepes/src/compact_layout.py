"""Physical blocks for the compact workbook; shared by builder, VBA and tests."""
import re
from openpyxl.formula.tokenizer import Tokenizer

# (destination sheet, zero-based row offset). No business calculations change.
BLOCKS = {
    'DASHBOARD': ('RINGKASAN', 0), 'MULAI': ('RINGKASAN', 80),
    'KASIR': ('KASIR', 0), 'STRUK': ('KASIR', 46),
    'PRODUKSI': ('HARIAN', 0), 'FORM': ('HARIAN', 38),
    'TUTUP KASIR': ('TUTUP KASIR', 0),
    'BAHAN': ('KATALOG', 224),
    'MENU': ('KATALOG', 0), 'KALIBRASI': ('KATALOG', 62),
    'PREMIKS': ('KATALOG', 100), 'ADONAN': ('KATALOG', 134),
    'REFERENSI': ('KATALOG', 166),
    'PELANGGAN': ('PELANGGAN', 0),
    'NOTA': ('JURNAL', 0), 'ITEM': ('JURNAL', 1014),
    'GERAK BAHAN': ('JURNAL', 4028), 'KEUANGAN': ('JURNAL', 10042),
    'POIN': ('JURNAL', 11056),
    'ARSIP TUTUP': ('JURNAL', 12070),
    'RESEP': ('_RESEP', 0),
}
ORDER = ['RINGKASAN', 'KASIR', 'STOK', 'TUTUP KASIR', 'HARIAN', 'KATALOG', 'PELANGGAN', 'JURNAL', '_RESEP']
# Compatibility addresses for the original independent business fixtures only.
CLOSE_TEST_CELLS = {'B5':'C5','B7':'I8','E7':'I11','B14':'I10','B15':'I12'}
CELL = re.compile(r'^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$')

def shift_cell(address, offset):
    m = CELL.fullmatch(address)
    if not m:
        raise ValueError(f'Unsupported cell reference: {address}')
    return f'{m[1]}{m[2]}{m[3]}{int(m[4])+offset}'

def shift_range(address, offset):
    return ':'.join(shift_cell(a, offset) for a in address.split(':'))

def reference(value, context=None, local=False):
    """Map A1 references, preserving absolute flags and leaving names intact."""
    if '!' in value:
        sheet, address = value.rsplit('!', 1)
        sheet = sheet.strip("'").replace("''", "'")
    else:
        sheet, address = context, value
    if sheet not in BLOCKS or not all(CELL.fullmatch(a) for a in address.split(':')):
        return value
    target, offset = BLOCKS[sheet]
    mapped = shift_range(address, offset)
    if local and context is not None and target == BLOCKS[context][0]:
        return mapped
    return f"'{target}'!{mapped}"

def formula(value, context=None, local=False):
    if not value or not isinstance(value, str):
        return value
    prefixed = value.startswith('=')
    tokens = Tokenizer(value if prefixed else '='+value)
    for token in tokens.items:
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            token.value = reference(token.value, context, local)
    result = ''.join(t.value for t in tokens.items)
    return ('=' if prefixed else '') + result

class SheetView:
    """Only used by independent test fixtures to address logical source blocks."""
    def __init__(self, workbook, name):
        self.logical_name = name
        self.sheet, self.offset = BLOCKS[name]
        self.ws = workbook[self.sheet]
    def cell(self, row, column, value=None):
        return self.ws.cell(row+self.offset, column, value)
    def __getitem__(self, address):
        if self.logical_name=='TUTUP KASIR':address=CLOSE_TEST_CELLS.get(address,address)
        return self.ws[shift_range(address, self.offset)]
    def __setitem__(self, address, value):
        if self.logical_name=='TUTUP KASIR':address=CLOSE_TEST_CELLS.get(address,address)
        self.ws[shift_range(address, self.offset)] = value
    @property
    def print_area(self):
        return self.ws.print_area

class BookView:
    def __init__(self, workbook): self.raw = workbook
    def __getitem__(self, name):
        return SheetView(self.raw, name) if name in BLOCKS else self.raw[name]
    def __getattr__(self, name): return getattr(self.raw, name)
    def __iter__(self): return iter(self.raw)
