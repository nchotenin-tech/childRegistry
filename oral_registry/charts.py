"""Paired, round-specific charts. Every percentage uses the same children at both visits."""
from collections import Counter
from statistics import mean, median

from .analysis import build_analysis, age_group, STATES, LEVELS, percentage
from .core import alert_matches


def paired_metrics(registry, first, last, school='', age=None):
    data, fields = build_analysis(registry, school=school, rounds=(first,last))
    if age is not None:
        data = [p for p in data if p['before'] is not None and age_group(p['before']) == age]
    metrics = {}
    for key, spec in fields.items():
        pairs = [p for p in data if p['states'].get(key) in STATES[:4]]
        rule = 3 if key == 'r_ohplaque' else spec['AlertValue']
        before = sum(alert_matches(p['before'].get(key),rule) for p in pairs)
        after = sum(alert_matches(p['after'].get(key),rule) for p in pairs)
        metrics[key] = dict(n=len(pairs), before=before, after=after, pairs=pairs,
                            excluded=len(data)-len(pairs), change=Counter(p['states'][key] for p in pairs))
    return data, metrics


def show_charts(parent, registry):
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QColor, QPainter, QPen, QFont
    from PySide6.QtWidgets import (QDialog,QVBoxLayout,QHBoxLayout,QLabel,QComboBox,QScrollArea,
        QWidget,QPushButton,QGroupBox,QLayout,QTableWidget,QTableWidgetItem,QAbstractItemView,QStyle)

    class Bars(QWidget):
        def __init__(self, rows, maximum=100):
            super().__init__()
            self.rows=rows;self.maximum=max(1,maximum)
            self.setProperty('chartRows', rows)
            self.setMinimumHeight(len(rows)*38+24)
            self.setMinimumWidth(420)
        def paintEvent(self,event):
            p=QPainter(self);p.setRenderHint(QPainter.RenderHint.Antialiasing)
            left=110; width=max(50,self.width()-left-180)
            p.setFont(QFont('Tahoma',10))
            for index,(caption,value,label,color) in enumerate(self.rows):
                y=index*38+6
                p.setPen(QColor('#334155'));p.drawText(0,y+18,caption)
                p.setPen(Qt.PenStyle.NoPen);p.setBrush(QColor('#e8edf3'));p.drawRoundedRect(left,y,width,25,5,5)
                p.setBrush(QColor(color))
                if value: p.drawRoundedRect(left,y,int(width*value/self.maximum),25,5,5)
                p.setPen(QColor('#173e65'));p.drawText(left+width+10,y+18,label)
            p.end()

    dialog=QDialog(parent);dialog.setObjectName('roundCharts');dialog.setWindowTitle('กราฟเปรียบเทียบรอบตรวจ')
    dialog.resize(1280,900);layout=QVBoxLayout(dialog)
    title=QLabel('กราฟเปรียบเทียบ • เด็กกลุ่มเดิมทั้งสองรอบ')
    title.setStyleSheet('font-size:23px;font-weight:bold;color:#173e65;padding:8px;');layout.addWidget(title)
    display=QComboBox();display.setObjectName('chartDisplay')
    display.addItem('คำตอบทุกรหัส (พร้อมสรุปความเสี่ยง)', 'answers')
    display.addItem('เฉพาะร้อยละที่เสี่ยงตาม AlertValue', 'risk')
    layout.addWidget(display)
    filters=QHBoxLayout()
    visits=sorted({str(d.get('round_id')) for _,d in registry.records()},key=lambda v:int(v) if v.isdigit() else 999999)
    first=QComboBox();first.setObjectName('chartFirst');last=QComboBox();last.setObjectName('chartLast')
    for v in visits: first.addItem('รอบ '+v,v);last.addItem('รอบ '+v,v)
    if len(visits)>1:last.setCurrentIndex(len(visits)-1)
    schools=QComboBox();schools.addItem('ทุกโรงเรียน','')
    for v in sorted({str(d.get('sch') or '') for _,d in registry.records()}):
        if v:schools.addItem(v,v)
    age=QComboBox();age.setObjectName('chartAge')
    for label,value in [('ทุกกลุ่มอายุ',None),('ต่ำกว่า 3 ปี ณ รอบต้นทาง',0),('ตั้งแต่ 3 ปี ณ รอบต้นทาง',1)]:age.addItem(label,value)
    category=QComboBox();category.setObjectName('chartCategory')
    mapping={}
    if 'fieldmapping' in registry.book:
        for row in registry.book['fieldmapping'].iter_rows(min_row=2,values_only=True):
            if len(row)>=3 and row[2]:mapping[row[2]]=str(row[0] or 'อื่น ๆ')
    category.addItem('ภาพรวม + ทุกหมวด','')
    groups=[]
    for key in registry.fields:
        group='เกณฑ์ความเสี่ยง' if key.startswith('r_') else mapping.get(key,'อื่น ๆ')
        if group not in groups:groups.append(group)
    for group in groups:category.addItem(group,group)
    for label,control in [('จาก',first),('เทียบกับ',last),('โรงเรียน',schools),('อายุ',age),('หมวด',category)]:
        filters.addWidget(QLabel(label));filters.addWidget(control)
    layout.addLayout(filters)
    summary=QLabel();summary.setWordWrap(True);summary.setObjectName('chartSummary')
    summary.setStyleSheet('background:#e3edf8;color:#173e65;padding:12px;font-size:16px;');layout.addWidget(summary)
    note=QLabel('สีฟ้า = รอบต้นทาง • สีม่วง = รอบปลายทาง • เลือกดูสัดส่วนคำตอบทุกรหัส หรือร้อยละที่เสี่ยงตาม AlertValue ได้\nกราฟอายุใช้กลุ่มต้นทาง; ผู้ข้ามเกณฑ์อายุไม่เทียบคะแนน/ปัจจัยคำนวณ แต่ยังเทียบคำตอบที่นิยามตรงกันได้ • ข้อมูลขาด/รหัสไม่ทราบไม่ถือว่าไม่เสี่ยง')
    note.setWordWrap(True);layout.addWidget(note)
    scroll=QScrollArea();scroll.setWidgetResizable(True);layout.addWidget(scroll,1)
    def names(pairs, label):
        pop=QDialog(dialog);pop.setWindowTitle(label);pop.resize(1050,600);v=QVBoxLayout(pop)
        table=QTableWidget(len(pairs),7);table.setHorizontalHeaderLabels(['ชื่อเด็ก','CID','รอบต้นทาง','รอบปลายทาง','ค่าก่อน','ค่าหลัง','เปลี่ยนแปลง'])
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        for r,(p,a,b,state) in enumerate(pairs):
            for c,val in enumerate((p['name'],p['cid'],p['before']['round_id'],p['after']['round_id'],a,b,state)):
                table.setItem(r,c,QTableWidgetItem(str(val)))
        table.resizeColumnsToContents();v.addWidget(table);pop.exec()
    def refresh():
        page=QWidget();cards=QVBoxLayout(page);cards.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize);old=scroll.takeWidget()
        if old:old.deleteLater()
        scroll.setWidget(page)
        if first.currentData() is None or first.currentData()==last.currentData():
            summary.setText('เลือกรอบตรวจต่างกันอย่างน้อย 2 รอบ');return
        data,metrics=paired_metrics(registry,first.currentData(),last.currentData(),schools.currentData(),age.currentData())
        pairs=[p for p in data if p['before'] is not None]
        valid=[p for p in pairs if p['risk_change']!=STATES[4]]
        deltas=[p['delta'] for p in pairs if p['delta'] is not None]
        cross=sum(age_group(p['before']) != age_group(p['after']) for p in pairs)
        summary.setText(f"{first.currentText()} → {last.currentText()} | มีข้อมูลทั้งสองรอบ {len(data)} คน • ลำดับวันถูกต้อง {len(pairs)} คน • ข้ามกลุ่มอายุ {cross} คน • เทียบระดับความเสี่ยงได้ {len(valid)} คน\n" + (f'คะแนนเปลี่ยนเฉลี่ย {mean(deltas):+.2f} • มัธยฐาน {median(deltas):+.2f} • n={len(deltas)} (คะแนนลดเป็นค่าลบ)' if deltas else 'ไม่มีคะแนนคู่ที่เปรียบเทียบได้'))
        if not data:
            cards.addWidget(QLabel('ไม่มีเด็กที่มีทั้งสองรอบตามตัวกรองนี้'));return
        chosen=category.currentData()
        if not chosen:
            box=QGroupBox('ระดับความเสี่ยงก่อน–หลัง • ฐานเด็กชุดเดียวกัน');v=QVBoxLayout(box)
            rows=[]
            for label in LEVELS:
                for side,color in [('before','#3688cf'),('after','#9861bd')]:
                    count=sum(p[side].get('risklevel')==label for p in valid)
                    rows.append((label+(' ก่อน' if side=='before' else ' หลัง'),100*count/len(valid) if valid else 0,percentage(count,len(valid)),color))
            v.addWidget(Bars(rows));cards.addWidget(box)
            box=QGroupBox('ทิศทางการเปลี่ยนระดับความเสี่ยง');v=QVBoxLayout(box)
            counts=Counter(p['risk_change'] for p in valid)
            v.addWidget(Bars([(s,100*counts[s]/len(valid) if valid else 0,percentage(counts[s],len(valid)),color) for s,color in zip(STATES[:4],['#2c9f69','#d75160','#3688cf','#d6a132'])]));cards.addWidget(box)
        shown=0
        for key,spec in registry.fields.items():
            group='เกณฑ์ความเสี่ยง' if key.startswith('r_') else mapping.get(key,'อื่น ๆ')
            if chosen and group!=chosen:continue
            if key not in metrics:
                if chosen:
                    label=QLabel(spec['thai_label']+' • ไม่มี AlertValue จึงไม่สรุปว่าเสี่ยง/ไม่เสี่ยง');label.setWordWrap(True);cards.addWidget(label)
                continue
            m=metrics[key]; n=m['n']
            label='พบคราบ Front-8 (เสี่ยง=3)' if key=='r_ohplaque' else spec['thai_label']
            box=QGroupBox(group+' • '+key);v=QVBoxLayout(box)
            heading=QLabel(label);heading.setWordWrap(True);heading.setStyleSheet('font-weight:bold;font-size:15px;color:#173e65;');v.addWidget(heading)
            if not n:
                v.addWidget(QLabel(f'ไม่มีคู่ที่เปรียบเทียบได้ • ตัดออก {m["excluded"]} คน (ข้อมูล/AlertValue/กลุ่มอายุไม่ตรง)'))
            else:
                if display.currentData() == 'answers' and spec.get('type') == 'enum':
                    rows=[]
                    for code in str(spec.get('allowed_codes') or '').split(','):
                        for side,caption,color in [('before',first.currentText(),'#3688cf'),('after',last.currentText(),'#9861bd')]:
                            count=sum(str(p[side].get(key))==code for p in m['pairs'])
                            rows.append((f'{caption} • รหัส {code}',100*count/n,percentage(count,n),color))
                    v.addWidget(Bars(rows))
                    definition=QLabel('กราฟแสดงสัดส่วนคำตอบทุกรหัสในคู่ที่เปรียบเทียบได้ • ค่าเสี่ยงตาม AlertValue: ' + str(3 if key=='r_ohplaque' else spec.get('AlertValue')))
                    definition.setWordWrap(True);v.addWidget(definition)
                else:
                    v.addWidget(Bars([(first.currentText(),100*m['before']/n,percentage(m['before'],n),'#3688cf'),(last.currentText(),100*m['after']/n,percentage(m['after'],n),'#9861bd')]))
                risk_note=QLabel(f"เข้าเกณฑ์เสี่ยง: {first.currentText()} {percentage(m['before'],n)} → {last.currentText()} {percentage(m['after'],n)}")
                risk_note.setWordWrap(True);v.addWidget(risk_note)
                if m['before']==0 and m['after']==0:
                    explanation=QLabel(f'มีข้อมูล {n} คน แต่ไม่มีคำตอบที่ตรง AlertValue ทั้งสองรอบ จึงเป็น 0% ไม่ใช่ข้อมูลหาย')
                    explanation.setWordWrap(True);explanation.setStyleSheet('color:#176b59; background:#e5f5ef; padding:6px;');v.addWidget(explanation)
                diff=100*(m['after']-m['before'])/n
                change=m['change']
                text=QLabel(f'ความชุกค่าเสี่ยงเปลี่ยน {diff:+.1f} จุดร้อยละ • ดีขึ้น {change[STATES[0]]} คน / แย่ลง {change[STATES[1]]} คน • ตัดออก {m["excluded"]} คน')
                text.setWordWrap(True);v.addWidget(text)
                button=QPushButton('ดูเด็กที่นำมาเปรียบเทียบ')
                button.setCursor(Qt.CursorShape.PointingHandCursor)
                button.setIcon(dialog.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
                button.clicked.connect(lambda checked=False,k=key,metric=m,l=label:names([(p,p['before'].get(k),p['after'].get(k),p['states'][k]) for p in metric['pairs']],l))
                v.addWidget(button)
            box.setMinimumHeight(220 if n else 95)
            cards.addWidget(box);shown+=1
        if not shown:cards.addWidget(QLabel('หมวดนี้ไม่มีรายการที่กำหนด AlertValue'))
        cards.addStretch()
        cards.activate()
        page.adjustSize()
    for control in (first,last,schools,age,category,display):control.currentIndexChanged.connect(refresh)
    refresh();dialog.exec()
