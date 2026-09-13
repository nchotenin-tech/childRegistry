def test_blank_template_and_first_run(tmp_path):
    from oral_registry.runtime import initial_database
    from oral_registry.storage import Registry
    from pathlib import Path
    template=Registry('templateRegistry.xlsx')
    assert template.records()==[]
    for sheet in template.book:
        assert not any(c.comment or c.hyperlink for row in sheet for c in row)
    target=initial_database(tmp_path)
    assert Registry(target).records()==[]
    target.write_bytes(b'existing database must not be overwritten')
    assert initial_database(tmp_path).read_bytes()==b'existing database must not be overwritten'
