"""Usage-statistics tests (dashboard feature, post-v0.4.0).

Covers the local JSONL usage log: recording, tolerant loading, engine
bucketing, stats aggregation with a fixed clock, and the execute()
wrapper recording success/failure without changing execute's contract.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, OSError, ValueError): pass

import datetime
import tempfile
from pathlib import Path
from unittest.mock import patch

import app

CHECKS = []

def ok(name, cond):
    assert cond, name
    CHECKS.append(name); print(f'  [PASS] {name}')

def run():
    with tempfile.TemporaryDirectory() as td:
        work = Path(td) / '.work'
        with patch.object(app, 'WORK', work):
            ok('U-01 chua co log: load tra rong, khong loi', app.load_usage_events() == [])
            app.record_usage('lookup', 'success', 1.5, rows=120, engine_used='Ollama Cục bộ (Offline)', source_name='D:/data/ho_so.xlsx')
            app.record_usage('summary', 'success', 2.5, rows=40, engine_used='Smart Heuristic Engine', source_name='so_quy.xlsx')
            app.record_usage('lookup', 'failed', 0.5, engine_used='OpenAI API', source_name='x.xlsx')
            with open(work / 'usage-log.jsonl', 'a', encoding='utf-8') as fh:
                fh.write('dong nay khong phai json\n')
            events = app.load_usage_events()
            ok('U-02 ghi 3 su kien, dong hong bi bo qua khi doc', len(events) == 3)
            ok('U-03 chi luu ten file, khong luu duong dan', events[0]['source'] == 'ho_so.xlsx')
            ok('U-04 phan loai dong co: Ollama/Heuristic/Cloud',
               [e['engine'] for e in events] == ['Ollama', 'Heuristic', 'Cloud'])

            now = datetime.datetime.now()
            stats = app.compute_usage_stats(events, days=30, now=now)
            ok('U-05 tong hop: 3 luot, 2 thanh cong, ty le 66.7%',
               stats['total_period'] == 3 and stats['success'] == 2 and abs(stats['success_rate'] - 66.666) < 0.1)
            ok('U-06 hom nay = 3, tong dong = 160, thoi gian TB = 1.5s',
               stats['today'] == 3 and stats['total_rows'] == 160 and abs(stats['avg_duration'] - 1.5) < 0.01)
            ok('U-07 per_day du 30 ngay va tong khop 3', len(stats['per_day']) == 30 and sum(stats['per_day'].values()) == 3)
            ok('U-08 per_action/per_engine dung', stats['per_action'] == {'lookup': 2, 'summary': 1} and stats['per_engine'].get('Ollama') == 1)
            old = dict(events[0])
            old['ts'] = (now - datetime.timedelta(days=45)).isoformat(timespec='seconds')
            stats2 = app.compute_usage_stats(events + [old], days=30, now=now)
            ok('U-09 su kien ngoai 30 ngay bi loai khoi ky nhung van dem o total_all',
               stats2['total_period'] == 3 and stats2['total_all'] == 4)

            with patch.object(app, '_execute_impl', return_value={'stats': {}}) as impl:
                src = {'path': 'a.xlsx', 'data': [[1], [2]], 'headers': ['A']}
                out = app.execute(src, None, {'action': 'clean', 'engine_used': 'Ollama Cục bộ (Offline)'}, approval_token='tok')
            ok('U-10 execute wrapper tra nguyen ket qua cua impl', out == {'stats': {}})
            last = app.load_usage_events()[-1]
            ok('U-11 execute thanh cong duoc ghi: success, rows=2, batch=False',
               last['status'] == 'success' and last['rows'] == 2 and last['batch'] is False)

            def boom(*a, **k):
                raise RuntimeError('excel boom')
            with patch.object(app, '_execute_impl', boom):
                try:
                    app.execute(src, None, {'action': 'clean'}, approval_token='tok', _usage_batch=True)
                    ok('U-12 execute loi van nem lai ngoai le', False)
                except RuntimeError:
                    ok('U-12 execute loi van nem lai ngoai le', True)
            last = app.load_usage_events()[-1]
            ok('U-13 execute loi duoc ghi failed va batch=True', last['status'] == 'failed' and last['batch'] is True)

            cur, prev = app.compute_usage_pair(events, days=30, now=now)
            ok('U-14 cap ky truoc/hien tai tach dung (ky truoc rong)', cur['total_period'] == 3 and prev['total_period'] == 0)
            old2 = dict(events[0])
            old2['ts'] = (now - datetime.timedelta(days=40)).isoformat(timespec='seconds')
            cur2, prev2 = app.compute_usage_pair(events + [old2], days=30, now=now)
            ok('U-15 su kien 40 ngay truoc roi vao ky truoc', cur2['total_period'] == 3 and prev2['total_period'] == 1)

            ancient = dict(events[0])
            ancient['ts'] = (now - datetime.timedelta(days=400)).isoformat(timespec='seconds')
            app.record_usage('clean', 'success', 0.1, engine_used='Ollama')
            with open(work / 'usage-log.jsonl', 'a', encoding='utf-8') as fh:
                fh.write(__import__('json').dumps(ancient) + '\n')
            removed = app.trim_usage_log(max_age_days=365)
            ok('U-16 trim bo su kien qua 365 ngay', removed >= 1 and all(
                (now - datetime.datetime.fromisoformat(e['ts'])).days <= 365 for e in app.load_usage_events()))

            app.save_dashboard_prefs(90, 'week')
            prefs = app.load_dashboard_prefs()
            ok('U-17 prefs luu/nap dung', prefs == {'days': 90, 'mode': 'week'})

    print(f'\nAll {len(CHECKS)} usage-stats checks passed.')

if __name__ == '__main__':
    run()
