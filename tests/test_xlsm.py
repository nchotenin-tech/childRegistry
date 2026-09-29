from zipfile import ZipFile, ZIP_DEFLATED
from oral_registry.storage import Registry
from oral_registry.core import assess
from oral_registry.xlsm import records_part
from test_registry import sample


def test_xlsm_save_edit_delete_preserves_package(registry, tmp_path):
    target = tmp_path / 'macro.xlsm'
    with ZipFile(registry.path) as src, ZipFile(target, 'w', ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = data.replace(b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', b'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
            out.writestr(item, data)
        out.writestr('xl/vbaProject.bin', b'synthetic macro preservation fixture')
        out.writestr('customUI/customUI.xml', b'<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>')
    original = target.read_bytes()
    with ZipFile(target) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
        allowed = {records_part(archive), 'xl/styles.xml'}
    r = Registry(target)
    data = assess(sample(r), r.book)
    data['cid'] = '9900000000998'
    _, backup = r.save(data)
    assert backup.suffix == '.xlsm'
    assert backup.read_bytes() == original
    data['name'] = 'Updated synthetic child'
    r.save(data, allow_update=True)
    assert next(d for _, d in r.records() if d['cid'] == data['cid'])['name'] == data['name']
    r.delete(data['cid'], data['round_id'])
    assert len(r.records()) == 1
    with ZipFile(target) as result:
        assert set(result.namelist()) == set(parts)
        for name, content in parts.items():
            if name not in allowed:
                assert result.read(name) == content, name
