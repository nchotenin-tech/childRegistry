import shutil
from pathlib import Path
from datetime import date
import pytest
from oral_registry.storage import Registry
from oral_registry.core import TEETH, FRONT, validate_record, assess

TEMPLATE=Path(__file__).resolve().parents[1]/'templateRegistry.xlsx'

def create_registry(path, count=1, rounds=1):
    shutil.copy2(TEMPLATE,path)
    r=Registry(path)
    for i in range(count):
        for n in range(1,rounds+1):
            raw={}
            for key,spec in r.fields.items():
                raw[key]=str(spec.get('allowed_codes')).split(',')[0] if spec['type']=='enum' else 0 if spec['type'] in ('integer','decimal') else 'ข้อมูลทดสอบ'
            raw.update(cid=f'9900000{i:06d}',name=f'เด็กทดสอบ {i}',round_id=n,dob=date(2022,9,6),iv_date=date(2026,[1,5,9][n-1] if rounds>1 else 9,11),wt=17,ht=105,br_method=2,br_freq=2,ftp_freq=2,tm_morn=2,tm_bed=2,bottle=4,dev=1,dev_note='',sweet_freq=i%2)
            for t in TEETH:raw['t'+t]='A'
            for t in FRONT:raw['d'+t]=0
            r.save(assess(validate_record(raw,r.fields),r.book))
    return r

@pytest.fixture
def registry(tmp_path):
    return create_registry(tmp_path/'registry.xlsx')

@pytest.fixture
def cohort(tmp_path):
    return create_registry(tmp_path/'cohort.xlsx',3,3)
@pytest.fixture(autouse=True)
def output_directory():
    Path('tmp').mkdir(exist_ok=True)
