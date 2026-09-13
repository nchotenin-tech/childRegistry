from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

from pydantic import BaseModel, Field, field_validator, model_validator

TEETH = '55 54 53 52 51 61 62 63 64 65 85 84 83 82 81 71 72 73 74 75'.split()
FRONT = '52 51 61 62 82 81 71 72'.split()
COLORS = {'เสี่ยงต่ำ': 'DCFCE7', 'เสี่ยงสูง': 'FEF3C7', 'เสี่ยงสูงมาก': 'FECACA'}
TOOTH_CODES = {
    'white_discolor': ['H'], 'caries_hole': ['B', 'C'],
    'filled': ['C', 'D'], 'extracted': ['E'], 'absent': ['E', '9'],
}
TOOTH_LABELS = {
    'A': 'ปกติ', 'H': 'ขาวขุ่น/เปลี่ยนสีน้ำตาล',
    'N': 'เคลือบฟันแตก/เงาจากเนื้อฟัน', 'B': 'ผุเป็นรูถึงเนื้อฟัน',
    'C': 'อุดแล้วมีผุ', 'D': 'อุดไม่มีผุ/ครอบ SSC', 'E': 'หายเพราะฟันผุ',
    'F': 'เคลือบหลุมร่องฟัน', 'G': 'หลักสะพาน/ครอบพิเศษ/วีเนียร์',
    'T': 'บาดเจ็บ/แตกหัก', '9': 'หายอื่น/ยังไม่ขึ้น/ไม่ได้บันทึก',
}


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parts = re.split(r'[-/]', str(value).strip())
    if len(parts) != 3:
        raise ValueError('วันที่ต้องเป็น DD-MM-YYYY (พ.ศ.) หรือ YYYY-MM-DD (ค.ศ.)')
    if len(parts[0]) == 4:
        year, month, day = map(int, parts)
    else:
        day, month, year = map(int, parts)
    if year >= 2400:
        year -= 543
    return date(year, month, day)


class Identity(BaseModel):
    round_id: int = Field(ge=1)
    name: str = Field(min_length=1)
    cid: str
    dob: date
    iv_date: date
    wt: float = Field(gt=10, allow_inf_nan=False)
    ht: float = Field(gt=70, allow_inf_nan=False)

    @field_validator('dob', 'iv_date', mode='before')
    @classmethod
    def dates(cls, value):
        return parse_date(value)

    @field_validator('cid')
    @classmethod
    def identifier(cls, value):
        if not re.fullmatch(r'[0-9]{13}', value) or value.startswith('00'):
            raise ValueError('รหัสต้องมีตัวเลข 13 หลัก และไม่ขึ้นต้นด้วย 00')
        return value

    @model_validator(mode='after')
    def chronology(self):
        if self.iv_date < self.dob:
            raise ValueError('วันตรวจต้องไม่ก่อนวันเกิด')
        return self


def age_at(birth: date, exam: date):
    months = (exam.year - birth.year) * 12 + exam.month - birth.month - (exam.day < birth.day)
    return divmod(months, 12)


def alert_matches(value: Any, expression: Any) -> bool:
    if value is None or expression is None:
        return False
    for token in str(expression).split(','):
        token = token.strip().strip('"\'')
        match = re.fullmatch(r'(>=|<=|>|<)\s*(-?\d+(?:\.\d+)?)', token)
        if match:
            try:
                x, y = float(value), float(match[2])
            except (TypeError, ValueError):
                continue
            if {'>': x > y, '<': x < y, '>=': x >= y, '<=': x <= y}[match[1]]:
                return True
        elif str(value) == token:
            return True
        else:
            try:
                if float(value) == float(token):
                    return True
            except (ValueError, TypeError):
                pass
    return False


def validate_record(raw, fields):
    data, errors = {}, []
    for key, spec in fields.items():
        if key.startswith(('calc_', 'r_')) or key in ('age_y', 'age_m', 'riskscore', 'risklevel'):
            continue
        value = raw.get(key)
        if isinstance(value, str):
            value = value.strip()
        if value in (None, ''):
            if key != 'dev_note' or spec.get('NotBlank') == 1:
                errors.append(f"{spec['thai_label']}: กรุณากรอกข้อมูล")
            data[key] = None
            continue
        try:
            kind = spec['type']
            codes = str(spec.get('allowed_codes') or '').split(',')
            if kind == 'enum':
                value = str(value)
                if value not in codes:
                    raise ValueError('รหัสที่อนุญาต: ' + ', '.join(codes))
                value = int(value) if value.isdigit() else value
            elif kind in ('integer', 'decimal'):
                number = float(value)
                if not math.isfinite(number) or number < 0 or (kind == 'integer' and not number.is_integer()):
                    raise ValueError('ต้องเป็นจำนวนไม่ติดลบ' + ('เต็ม' if kind == 'integer' else ''))
                value = int(number) if kind == 'integer' else number
            elif kind == 'date':
                value = parse_date(value)
            else:
                value = str(value)
            data[key] = value
        except (ValueError, TypeError) as exc:
            errors.append(f"{spec['thai_label']}: {exc}")
    if errors:
        raise ValueError('\n'.join(errors))
    try:
        identity = Identity.model_validate(data)
    except ValueError as exc:
        raise ValueError(str(exc)) from exc
    if data.get('dev') == 2 and not data.get('dev_note'):
        raise ValueError('กรุณาระบุด้านพัฒนาการที่ล่าช้า')
    data.update(identity.model_dump())
    data['age_y'], data['age_m'] = age_at(identity.dob, identity.iv_date)
    return data


def assess(data, workbook, tooth_codes=None):
    tooth_codes = tooth_codes or TOOTH_CODES
    d = dict(data)
    statuses = [str(d['t' + t]) for t in TEETH]
    for category in ('white_discolor', 'caries_hole', 'filled', 'extracted'):
        d['calc_' + category] = sum(s in tooth_codes[category] for s in statuses)
    d['calc_teeth_in_mouth'] = sum(s not in tooth_codes['absent'] for s in statuses)
    if '9' in statuses:
        d['calc_teeth_in_mouth'] = None
    d['calc_caries_front8'] = sum(str(d['t' + t]) in tooth_codes['caries_hole'] for t in FRONT)
    plaque = any(d['d' + t] in (1, 2, 3) for t in FRONT)
    if all(d['d' + t] == 9 for t in FRONT) and d['age_y'] < 3:
        raise ValueError('Debris Index เป็น 9 ทุกซี่ จึงประเมินความสะอาด Front-8 ไม่ได้')
    d['calc_plaque_front8'] = 2 if plaque else 1
    yes = {
        'r_brush': d['br_method'] in (1, 2) and d['br_freq'] == 2,
        'r_ohplaque': plaque,
        'r_caries': any(s in ('H', 'N', 'B', 'C') for s in statuses),
        'r_brushflu': d['ftp_freq'] == 2 and d['tm_morn'] == 2 and d['tm_bed'] == 2,
        'r_milksweet': d['sweet_bottle'] == 1,
        'r_milksleep': d['feed_sleep12'] == 1,
        'r_sweet': d['sweet_freq'] == 1,
        'r_cg': d['cg_caries6m'] == 1,
        'r_vplaque': d['vplaque'] == 1,
        'r_usebottle': d['bottle'] in (2, 3),
        'r_siblings': d['sib_caries'] == 1,
    }
    s = workbook['ตารางRiskScore']
    young_rows = dict(zip(list(yes)[:8], range(6, 14)))
    older_order = ['r_brush', 'r_vplaque', 'r_caries', 'r_brushflu', 'r_usebottle', 'r_milksleep', 'r_sweet', 'r_siblings']
    older_rows = dict(zip(older_order, range(17, 25)))
    group_col = 3 if d['age_y'] < 3 else 4
    mapping = young_rows if group_col == 3 else older_rows
    total = 0
    for row in range(5, 16):
        key = s.cell(row, 1).value
        if key not in yes:
            raise ValueError('ตารางRiskScore มีชื่อปัจจัยที่ไม่รองรับ')
        active = str(s.cell(row, group_col).value or '').strip() == '/'
        source = mapping.get(key, young_rows.get(key, older_rows.get(key)))
        value = s.cell(source, 7 if yes[key] else 8).value
        if not isinstance(value, (int, float)) or value < 0:
            raise ValueError('คะแนนในตารางRiskScore ต้องเป็นตัวเลขไม่ติดลบ')
        d[key] = value
        if active:
            total += value
    thresholds = []
    for row in (27, 28):
        match = re.match(r'(\d+)\s*-\s*(\d+)\s*คะแนน', str(s.cell(row, 6).value))
        if not match:
            raise ValueError('อ่านเกณฑ์แบ่งคะแนนไม่ได้')
        thresholds.append(int(match[2]))
    d['riskscore'] = total
    d['risklevel'] = 'เสี่ยงต่ำ' if total <= thresholds[0] else 'เสี่ยงสูง' if total <= thresholds[1] else 'เสี่ยงสูงมาก'
    return d
