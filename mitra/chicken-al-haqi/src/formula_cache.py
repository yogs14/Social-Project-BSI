"""Refresh cached formula values while preserving native Excel UI objects."""
import re
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def cache_only(original, calculated):
    with ZipFile(original) as source, ZipFile(calculated) as calc:
        shared=[]
        if 'xl/sharedStrings.xml' in calc.namelist():
            shared=[''.join(si.itertext()) for si in ET.fromstring(calc.read('xl/sharedStrings.xml'))]
        parts={name:source.read(name) for name in source.namelist()}
        for name in parts:
            if not re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name):continue
            cached={c.attrib['r']:c for c in ET.fromstring(calc.read(name)).findall('.//s:c',NS)}
            raw=parts[name].decode()
            def patch(match):
                opening,coordinate,body=match.group(1),match.group(2),match.group(3)
                if '<f' not in body:return match.group(0)
                cell=cached.get(coordinate)
                assert cell is not None,coordinate
                value=cell.find('s:v',NS)
                kind=cell.get('t')
                text=value.text if value is not None and value.text is not None else ''
                if kind=='s':kind='str';text=shared[int(text)] if text else ''
                if kind=='inlineStr':kind='str';text=''.join(cell.find('s:is',NS).itertext())
                opening=re.sub(r'\s+t="[^"]*"','',opening)
                if kind in ['str','b','e']:opening=opening[:-1]+' t="'+kind+'">'
                body=re.sub(r'<v>.*?</v>|<v\s*/>','',body,flags=re.S)
                return opening+body+'<v>'+escape(text)+'</v></c>'
            parts[name]=re.sub(r'(<c\b[^>]*\br="([^"]+)"[^>]*>)(.*?)</c>',patch,raw,flags=re.S).encode()
    temp=original.with_suffix('.validated.xlsx')
    with ZipFile(temp,'w',ZIP_DEFLATED) as out:
        for name,data in parts.items():out.writestr(name,data)
    temp.replace(original)

