from oral_registry.analysis import field_state, build_analysis, STATES
from oral_registry.storage import Registry
from datetime import date


def test_alert_transitions():
    spec={'type':'enum','allowed_codes':'0,1,9','AlertValue':1}
    assert field_state('prob_any',1,0,spec)==STATES[0]
    assert field_state('prob_any',0,1,spec)==STATES[1]
    assert field_state('prob_any',0,0,spec)==STATES[2]
    assert field_state('prob_any',1,1,spec)==STATES[3]
    assert field_state('prob_any',9,0,spec)==STATES[4]
    assert field_state('prob_any',None,0,spec)==STATES[4]
    assert field_state('x',1,0,{'AlertValue':None})==STATES[4]


def test_age_switch():
    class Fake:
        fields={'risklevel':{'AlertValue':'เสี่ยงสูง'},'sweet_freq':{'type':'enum','allowed_codes':'0,1','AlertValue':1}}
        def records(self):
            return [(4,dict(cid='1',round_id=1,dob=date(2023,1,1),iv_date=date(2025,1,1),sweet_freq=1,riskscore=5,risklevel='เสี่ยงสูง')),(5,dict(cid='1',round_id=2,dob=date(2023,1,1),iv_date=date(2026,1,1),sweet_freq=0,riskscore=2,risklevel='เสี่ยงต่ำ'))]
    import openpyxl
    fake=Fake();fake.book=openpyxl.load_workbook('templateRegistry.xlsx')
    data,_=build_analysis(fake)
    assert data[0]['delta'] is None
    assert data[0]['risk_change']==STATES[4]
    assert data[0]['states']['sweet_freq']==STATES[0]


def test_dashboard(monkeypatch, cohort):
    monkeypatch.setenv('QT_QPA_PLATFORM','offscreen')
    from PySide6.QtWidgets import QApplication,QDialog,QTableWidget
    from oral_registry.app import Window,configure_app
    from oral_registry.analysis import show_dashboard
    app=QApplication.instance() or QApplication([]);configure_app(app)
    w=Window(cohort.path)
    def inspect(d):
        d.show();app.processEvents()
        stats=d.findChild(QTableWidget,'analysisStats');people=d.findChild(QTableWidget,'analysisPeople')
        assert stats.rowCount()>10
        assert people.rowCount()==len({str(d['cid']) for _,d in w.registry.records()})
        stats.cellClicked.emit(0,1)
        if people.rowCount(): people.cellClicked.emit(0,0)
        app.processEvents();d.grab().save('tmp/analysis-dashboard.png');d.close();return 0
    monkeypatch.setattr(QDialog,'exec',inspect)
    show_dashboard(w);w.dirty=False;w.close()


def test_selected_round_pairs(cohort):
    from oral_registry.charts import paired_metrics
    r=cohort
    data, metrics=paired_metrics(r,'1','3')
    expected={str(d['cid']) for _,d in r.records() if str(d['round_id'])=='3'} & {str(d['cid']) for _,d in r.records() if str(d['round_id'])=='1'}
    assert {p['cid'] for p in data}==expected
    assert all(str(p['before']['round_id'])=='1' and str(p['after']['round_id'])=='3' for p in data)
    m=metrics['sweet_freq']
    assert m['n'] == len(data)
    assert m['before']==m['change'][STATES[0]]+m['change'][STATES[3]]
    assert m['after']==m['change'][STATES[1]]+m['change'][STATES[3]]
    young,_=paired_metrics(r,'1','3',age=0)
    from oral_registry.analysis import age_group
    assert all(age_group(p['before'])==0 for p in young)
    import pytest
    with pytest.raises(ValueError): paired_metrics(r,'1','1')
    reverse,_=paired_metrics(r,'3','1')
    assert all(p['before'] is None for p in reverse)


def test_charts_ui(monkeypatch, cohort):
    monkeypatch.setenv('QT_QPA_PLATFORM','offscreen')
    from PySide6.QtWidgets import QApplication,QDialog,QComboBox,QLabel
    from oral_registry.app import Window,configure_app
    from oral_registry.charts import show_charts
    app=QApplication.instance() or QApplication([]);configure_app(app)
    w=Window(cohort.path)
    def inspect(d):
        d.show();app.processEvents()
        assert d.objectName()=='roundCharts'
        assert 'ทั้งสองรอบ' in d.findChild(QLabel,'chartSummary').text()
        d.grab().save('tmp/round-charts.png')
        category=d.findChild(QComboBox,'chartCategory')
        category.setCurrentIndex(category.findData('เกณฑ์ความเสี่ยง'))
        d.findChild(QComboBox,'chartAge').setCurrentIndex(1)
        app.processEvents()
        d.grab().save('tmp/age-risk-charts.png')
        d.close();return 0
    monkeypatch.setattr(QDialog,'exec',inspect)
    show_charts(w,w.registry);w.dirty=False;w.close()
