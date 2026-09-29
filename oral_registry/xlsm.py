"""Preserve the original XLSM package; only Records cells and style tables change."""
from copy import deepcopy
from io import BytesIO
from pathlib import PurePosixPath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)

def records_part(archive):
    workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    rid = next(s for s in workbook.find(f'{{{NS}}}sheets') if s.get('name') == 'Records').get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    target = next(r.get('Target') for r in rels if r.get('Id') == rid)
    return target.lstrip('/') if target.startswith('/') else str(PurePosixPath('xl') / target)

def replace_section(xml, tag, element):
    value = ET.tostring(element, encoding='unicode')
    pattern = rf'<{tag}\b[^>]*(?:/>|>.*?</{tag}>)'
    if re.search(pattern, xml, re.S):
        return re.sub(pattern, lambda _: value, xml, count=1, flags=re.S)
    # numFmts is optional and must precede fonts.
    if tag == 'numFmts':
        return xml.replace('<fonts', value + '<fonts', 1)
    raise ValueError(f'Unsupported XLSM style table: {tag}')

def write_xlsm(original, book, destination):
    generated = BytesIO()
    book.save(generated)
    with ZipFile(original) as src, ZipFile(generated) as new:
        if any(n.startswith('_xmlsignatures/') for n in src.namelist()):
            raise ValueError('ไฟล์มีลายเซ็นเอกสาร กรุณาใช้สำเนาที่ไม่มีลายเซ็นเอกสารก่อนบันทึก')
        old_style = src.read('xl/styles.xml').decode('utf-8')
        old = ET.fromstring(old_style)
        fresh = ET.fromstring(new.read('xl/styles.xml'))
        maps = {}
        for table in ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs'):
            old_table = old.find(f'{{{NS}}}{table}')
            new_table = fresh.find(f'{{{NS}}}{table}')
            if new_table is None:
                continue
            if old_table is None:
                old_table = ET.Element(f'{{{NS}}}{table}')
            mapping = {}
            for index, entry in enumerate(new_table):
                entry = deepcopy(entry)
                if table in ('cellStyleXfs', 'cellXfs'):
                    for attr, name in [('fontId','fonts'), ('fillId','fills'), ('borderId','borders'), ('numFmtId','numFmts'), ('xfId','cellStyleXfs')]:
                        if attr in entry.attrib:
                            number = int(entry.get(attr))
                            entry.set(attr, str(maps.get(name, {}).get(number, number)))
                if table == 'numFmts':
                    key = int(entry.get('numFmtId'))
                    match = next((e for e in old_table if e.get('formatCode') == entry.get('formatCode')), None)
                    if match is None:
                        entry.set('numFmtId', str(max([163] + [int(e.get('numFmtId')) for e in old_table]) + 1))
                        old_table.append(entry)
                        match = entry
                    mapping[key] = int(match.get('numFmtId'))
                else:
                    serialized = ET.tostring(entry)
                    match = next((i for i,e in enumerate(old_table) if ET.tostring(e) == serialized), None)
                    if match is None:
                        match = len(old_table)
                        old_table.append(entry)
                    mapping[index] = match
            maps[table] = mapping
            old_table.set('count', str(len(old_table)))
            old_style = replace_section(old_style, table, old_table)
        part = records_part(src)
        sheet = ET.fromstring(new.read(records_part(new)))
        data = sheet.find(f'{{{NS}}}sheetData')
        for cell in data.iter(f'{{{NS}}}c'):
            if 's' in cell.attrib:
                cell.set('s', str(maps['cellXfs'][int(cell.get('s'))]))
        xml = src.read(part).decode('utf-8')
        xml = replace_section(xml, 'sheetData', data)
        dimension = sheet.find(f'{{{NS}}}dimension')
        if dimension is not None:
            xml = replace_section(xml, 'dimension', dimension)
        replacements = {part: xml.encode('utf-8'), 'xl/styles.xml': old_style.encode('utf-8')}
        with ZipFile(destination, 'w') as output:
            for item in src.infolist():
                output.writestr(item, replacements.get(item.filename, src.read(item.filename)))
