"""Read-only longitudinal analysis using Codebook AlertValue."""
from collections import defaultdict, Counter
from statistics import mean, median
from datetime import date
import math

from .core import alert_matches, parse_date, age_at

STATES = ['ดีขึ้น', 'แย่ลง', 'คงผลดี', 'ยังเสี่ยง', 'เปรียบเทียบไม่ได้']
PALETTE = ['#dcfce7', '#fecaca', '#dbeafe', '#fef3c7', '#eeeeee']
UNKNOWN9 = {'prob_any','food_imp','halitosis','gum_bleed','tooth_pain','gin_abs'}
LEVELS = {'เสี่ยงต่ำ': 0, 'เสี่ยงสูง': 1, 'เสี่ยงสูงมาก': 2}


def known(key, value):
    if value is None or str(value).strip() == '':
        return False
    if str(value).strip() in ('9', '9.0') and (key in UNKNOWN9 or key.startswith('t') and key[1:].isdigit() or key.startswith('d') and key[1:].isdigit()):
        return False
    return True


def field_state(key, before, after, spec):
    rule = 3 if key == 'r_ohplaque' else spec.get('AlertValue')
    if rule is None or str(rule).strip() == '' or not known(key, before) or not known(key, after):
        return STATES[4]
    codes = str(spec.get('allowed_codes') or '').split(',')
    if spec.get('type') == 'enum':
        if any(str(v) not in codes for v in (before, after)) or not any(alert_matches(c, rule) for c in codes):
            return STATES[4]
    if spec.get('type') in ('integer', 'decimal'):
        try:
            if any(not math.isfinite(float(v)) for v in (before, after)):
                return STATES[4]
        except (ValueError, TypeError):
            return STATES[4]
    a, b = alert_matches(before, rule), alert_matches(after, rule)
    return STATES[3] if a and b else STATES[0] if a else STATES[1] if b else STATES[2]


def exam_date(d):
    try:
        return parse_date(d.get('iv_date'))
    except (ValueError, TypeError):
        return None


def age_group(d):
    try:
        birth, exam = parse_date(d.get('dob')), parse_date(d.get('iv_date'))
        if exam < birth:
            return None
        return 0 if age_at(birth, exam)[0] < 3 else 1
    except (ValueError, TypeError):
        return None


def build_analysis(registry, mode='first', school='', rounds=None):
    people = defaultdict(list)
    for _, record in registry.records():
        people[str(record['cid'])].append(record)
    fields = {k: s for k,s in registry.fields.items() if s.get('AlertValue') is not None and str(s['AlertValue']).strip()}
    result = []
    for cid, history in people.items():
        history.sort(key=lambda d: (exam_date(d) or date.min, str(d.get('round_id'))))
        latest = history[-1]
        selected_before = None
        if rounds is not None:
            if rounds[0] == rounds[1]:
                raise ValueError('กรุณาเลือกรอบตรวจต่างกัน')
            first = [d for d in history if str(d.get('round_id')) == str(rounds[0])]
            second = [d for d in history if str(d.get('round_id')) == str(rounds[1])]
            if len(first) != 1 or len(second) != 1:
                continue
            selected_before, latest = first[0], second[0]
        if school and str(latest.get('sch')) != school:
            continue
        item = dict(cid=cid, name=latest.get('name'), history=history, before=None, after=latest, states={}, delta=None, risk_change='เปรียบเทียบไม่ได้', reason='ตรวจครั้งเดียว')
        result.append(item)
        if len(history) < 2:
            continue
        counts = Counter(d.get('round_id') for d in history)
        if any(v > 1 for v in counts.values()) or any(exam_date(d) is None for d in history):
            item['reason'] = 'รอบซ้ำหรือวันที่ตรวจไม่ถูกต้อง'
            continue
        before = selected_before if rounds is not None else history[0] if mode == 'first' else history[-2]
        if exam_date(before) >= exam_date(latest):
            item['reason'] = 'วันที่รอบปลายทางต้องหลังรอบต้นทาง'
            continue
        item['before'] = before
        comparable = age_group(before) is not None and age_group(before) == age_group(latest)
        item['reason'] = '' if comparable else 'เปลี่ยนเกณฑ์อายุหรืออายุไม่ถูกต้อง'
        for key, spec in fields.items():
            state = field_state(key, before.get(key), latest.get(key), spec)
            if key.startswith('r_') or key in ('riskscore', 'risklevel'):
                group = age_group(latest)
                active = {registry.book['ตารางRiskScore'].cell(r,1).value for r in range(5,16) if registry.book['ตารางRiskScore'].cell(r,3 if group == 0 else 4).value == '/'}
                if not comparable or key.startswith('r_') and key not in active:
                    state = STATES[4]
            item['states'][key] = state
        if comparable:
            try:
                delta = float(latest['riskscore']) - float(before['riskscore'])
                if math.isfinite(delta):
                    item['delta'] = delta
            except (ValueError, TypeError, KeyError):
                pass
            a,b = LEVELS.get(before.get('risklevel')), LEVELS.get(latest.get('risklevel'))
            if a is not None and b is not None:
                item['risk_change'] = STATES[0] if b<a else STATES[1] if b>a else STATES[2] if b==0 else STATES[3]
    return result, fields


def percentage(n, total):
    return f'{n}/{total} คน ({n / total * 100:.1f}%)' if total else '0/0 คน (ไม่มีข้อมูลเปรียบเทียบ)'


def show_dashboard(parent):
    from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QTableWidget, QTableWidgetItem, QAbstractItemView, QLineEdit, QSplitter, QHeaderView, QPushButton, QStyle
    from PySide6.QtGui import QColor
    from PySide6.QtCore import Qt
    from .storage import Registry
    registry = Registry(parent.registry.path)
    dialog = QDialog(parent)
    dialog.setWindowTitle('ติดตามผลรายบุคคลและ Dashboard')
    dialog.resize(1350, 900)
    layout = QVBoxLayout(dialog)
    heading = QHBoxLayout()
    badge = QLabel()
    badge.setPixmap(dialog.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon).pixmap(32,32))
    heading.addWidget(badge)
    title = QLabel('ติดตามพัฒนาการสุขภาพช่องปาก')
    title.setStyleSheet('font-size:22px; font-weight:bold; color:#173e65;')
    heading.addWidget(title,1)
    charts = QPushButton('กราฟเปรียบเทียบรอบตรวจ')
    charts.setIcon(dialog.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight))
    charts.setCursor(Qt.CursorShape.PointingHandCursor)
    def open_charts():
        from .charts import show_charts
        show_charts(dialog, registry)
    charts.clicked.connect(open_charts)
    heading.addWidget(charts)
    layout.addLayout(heading)
    filters = QHBoxLayout()
    school = QComboBox(); school.addItem('ทุกโรงเรียน', '')
    for value in sorted({str(d.get('sch') or '') for _, d in registry.records()}):
        if value: school.addItem(value,value)
    mode = QComboBox(); mode.addItem('ครั้งแรก → ล่าสุด','first'); mode.addItem('ครั้งก่อน → ล่าสุด','previous')
    search = QLineEdit(); search.setPlaceholderText('ค้นหาเด็ก: ชื่อ / CID')
    filters.addWidget(school); filters.addWidget(mode); filters.addWidget(search)
    layout.addLayout(filters)
    summary = QLabel(); summary.setWordWrap(True)
    summary.setStyleSheet('background:#e3edf8; color:#173e65; padding:14px; font-size:16px; font-weight:bold;')
    layout.addWidget(summary)
    cards = QHBoxLayout()
    card_labels = []
    for caption, icon, color in [('ดีขึ้น', QStyle.StandardPixmap.SP_ArrowUp, '#166534'), ('แย่ลง', QStyle.StandardPixmap.SP_ArrowDown, '#b91c1c'), ('คงผลดี', QStyle.StandardPixmap.SP_DialogApplyButton, '#1d4ed8'), ('ยังเสี่ยง', QStyle.StandardPixmap.SP_MessageBoxWarning, '#92400e')]:
        symbol=QLabel();symbol.setPixmap(dialog.style().standardIcon(icon).pixmap(28,28))
        value=QLabel(caption);value.setStyleSheet(f'color:{color};font-size:16px;font-weight:bold;padding:8px;')
        cards.addWidget(symbol);cards.addWidget(value,1);card_labels.append((caption,value))
    layout.addLayout(cards)
    note = QLabel('เรียงครั้งตรวจตามวันที่ • ตัวหารใช้คนที่มีข้อมูลเปรียบเทียบได้ในแต่ละรายการ • คลิกช่องสถิติเพื่อดูรายชื่อ แล้วคลิกเด็กเพื่อดูรายละเอียด\nคะแนนใช้ค่าที่บันทึกใน Records; ไม่นำคะแนนข้ามเกณฑ์อายุมารวม • สถานะ “ยังเสี่ยง” ไม่ได้หมายความว่าค่าตัวเลขเท่าเดิม • AlertValue ที่ไม่ตรงรหัสอนุญาตจะไม่เปรียบเทียบ')
    note.setWordWrap(True); layout.addWidget(note)
    matrix = QTableWidget(3,3)
    matrix.setHorizontalHeaderLabels(['ล่าสุด: ต่ำ','ล่าสุด: สูง','ล่าสุด: สูงมาก'])
    matrix.setVerticalHeaderLabels(['ก่อน: ต่ำ','ก่อน: สูง','ก่อน: สูงมาก'])
    matrix.setMaximumHeight(125)
    matrix.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    matrix.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    layout.addWidget(matrix)
    stats = QTableWidget(); stats.setObjectName('analysisStats')
    stats.setMaximumHeight(290)
    layout.addWidget(stats)
    insight = QLabel(); insight.setWordWrap(True); layout.addWidget(insight)
    split = QSplitter()
    people = QTableWidget(); people.setObjectName('analysisPeople')
    details = QTableWidget(); details.setObjectName('analysisDetails')
    split.addWidget(people); split.addWidget(details); split.setSizes([520,780]); layout.addWidget(split,1)
    for table in (stats,people,details):
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setWordWrap(True)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setStretchLastSection(True)
    context = {}
    def fill(table, headers, rows):
        table.setColumnCount(len(headers)); table.setHorizontalHeaderLabels(headers); table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                item=QTableWidgetItem('' if value is None else str(value));table.setItem(r,c,item)
                if value in STATES: item.setBackground(QColor(PALETTE[STATES.index(value)]))
        table.resizeColumnsToContents()
        for c in range(table.columnCount()): table.setColumnWidth(c,min(300,max(90,table.columnWidth(c))))
        table.resizeRowsToContents()
    def list_people(items):
        query = search.text().strip().casefold()
        context['visible'] = [p for p in items if query in f"{p['cid']} {p['name']}".casefold()]
        fill(people,['ชื่อเด็ก','CID','ครั้ง','ความเสี่ยง','คะแนน Δ'],[[p['name'],p['cid'],len(p['history']),p['risk_change'],p['delta']] for p in context['visible']])
        details.setRowCount(0)
    def refresh():
        data, fields = build_analysis(registry,mode.currentData(),school.currentData())
        context.update(data=data,fields=fields, drill=data)
        valid=[p for p in data if p['risk_change'] != STATES[4]]
        deltas=[p['delta'] for p in data if p['delta'] is not None]
        change=Counter(p['risk_change'] for p in valid)
        for caption,label in card_labels:
            label.setText(caption + '\n' + percentage(change[caption],len(valid)))
        summary.setText(f"เด็ก {len(data)} คน • ตรวจครั้งเดียว {sum(len(p['history'])==1 for p in data)} คน • มีหลายครั้ง {sum(len(p['history'])>1 for p in data)} คน • เทียบระดับความเสี่ยงได้ {len(valid)} คน\nดีขึ้น {percentage(change[STATES[0]],len(valid))} | แย่ลง {percentage(change[STATES[1]],len(valid))}\n" + (f'คะแนนล่าสุด − ก่อนหน้า: เฉลี่ย {mean(deltas):+.2f} • มัธยฐาน {median(deltas):+.2f} • n={len(deltas)} (ค่าติดลบ=คะแนนลด)' if deltas else 'ยังไม่มีคะแนนที่เปรียบเทียบได้'))
        rows=[]; context['keys']=['__risk__']+list(fields)
        for key in context['keys']:
            counts=Counter(p['risk_change'] if key=='__risk__' else p['states'].get(key,STATES[4]) for p in data)
            n=sum(counts[s] for s in STATES[:4])
            label='ระดับความเสี่ยงรวม' if key=='__risk__' else fields[key]['thai_label']
            if key=='r_ohplaque': label='พบคราบ Front-8 (ค่าเสี่ยง=3)'
            rows.append([label]+[percentage(counts[s],n) for s in STATES[:4]]+[str(counts[STATES[4]])+' คน',percentage(counts[STATES[0]],counts[STATES[0]]+counts[STATES[3]]),percentage(counts[STATES[1]],counts[STATES[1]]+counts[STATES[2]])])
        fill(stats,['รายการ']+STATES+['ดีขึ้นในกลุ่มเดิมเสี่ยง','แย่ลงในกลุ่มเดิมดี'],rows)
        for r in range(stats.rowCount()):
            for c in range(1,6): stats.item(r,c).setBackground(QColor(PALETTE[c-1]))
        behavioral=[k for k in fields if k in registry.columns and 25 <= registry.columns[k] <= 46]
        ranked=sorted([(sum(p['states'].get(k)==STATES[3] for p in data),k) for k in behavioral],reverse=True)
        if ranked and ranked[0][0]:
            count,key=ranked[0]; n=sum(p['states'].get(key) in STATES[:4] for p in data)
            insight.setText('พฤติกรรมที่ยังเสี่ยงทั้งสองครั้งมากที่สุด: '+fields[key]['thai_label']+' • '+percentage(count,n))
        else: insight.setText('ยังไม่มีพฤติกรรมเสี่ยงต่อเนื่องที่สรุปได้จากข้อมูลเปรียบเทียบ')
        for r in range(3):
            for c in range(3):
                count=sum(LEVELS.get(p['before']['risklevel'])==r and LEVELS.get(p['after']['risklevel'])==c for p in valid)
                item=QTableWidgetItem(f'{count} คน')
                item.setBackground(QColor(PALETTE[0 if c<r else 1 if c>r else 2 if c==0 else 3]))
                matrix.setItem(r,c,item)
        list_people(data)
    def matrix_drill(r,c):
        items=[p for p in context['data'] if p['risk_change']!=STATES[4] and LEVELS.get(p['before']['risklevel'])==r and LEVELS.get(p['after']['risklevel'])==c]
        context['drill']=items;list_people(items)
    matrix.cellClicked.connect(matrix_drill)
    def drill(row,col):
        key=context['keys'][row]
        items=context['data'] if col==0 else [p for p in context['data'] if (p['risk_change'] if key=='__risk__' else p['states'].get(key,STATES[4]))==STATES[0 if col==6 else 1]] if col>=6 else [p for p in context['data'] if (p['risk_change'] if key=='__risk__' else p['states'].get(key,STATES[4]))==STATES[col-1]]
        context['drill']=items;list_people(items)
    def individual(row,col):
        p=context['visible'][row]; b=p['before']; a=p['after']
        rows=[['ประวัติการตรวจ',' → '.join(f"รอบ {d['round_id']} ({exam_date(d)}) คะแนน {d.get('riskscore')}" for d in p['history']),'',''],['ข้อสังเกต',p['reason'],'','']]
        if b:
            rows.append(['วันที่เปรียบเทียบ',exam_date(b),exam_date(a),''])
            for key,spec in context['fields'].items():
                rows.append([spec['thai_label'],b.get(key),a.get(key),p['states'].get(key,STATES[4])])
        fill(details,['รายการ','ก่อน','ล่าสุด','ผล'],rows)
        details.setColumnWidth(0,220)
        details.setColumnWidth(1,130)
        details.setColumnWidth(2,130)
        details.setColumnWidth(3,130)
        details.resizeRowsToContents()
    school.currentIndexChanged.connect(refresh);mode.currentIndexChanged.connect(refresh)
    search.textChanged.connect(lambda: list_people(context.get('drill',[])))
    stats.cellClicked.connect(drill);people.cellClicked.connect(individual)
    refresh(); dialog.exec()
