"""
Live API + frontend integration test - Argos web app to'liq audit
"""
import urllib.request
import urllib.error
import json
import sys
import time

BASE = 'https://argostest.onrender.com'
TIMEOUT = 30

PASS = 0
FAIL = 0
ERRORS = []

def check(name, condition, detail=''):
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f'  PASS  {name}')
    else:
        FAIL += 1
        ERRORS.append(f'{name}: {detail}')
        print(f'  FAIL  {name}  |  {detail}')

def get(path):
    try:
        req = urllib.request.Request(BASE + path)
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            return r.status, r.read(), r.headers
    except urllib.error.HTTPError as e:
        return e.code, e.read(), None
    except Exception as e:
        return 0, str(e).encode(), None

def post(path, body=None):
    data = json.dumps(body or {}).encode()
    req = urllib.request.Request(BASE + path, data=data, headers={'Content-Type': 'application/json'}, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            return r.status, json.loads(r.read()), r.headers
    except urllib.error.HTTPError as e:
        return e.code, {}, None
    except Exception as e:
        return 0, {'_error': str(e)}, None


print('=' * 60)
print('ARGOS WEB APP - TO\'LIQ AUDIT')
print('=' * 60)

# ---- 1. Static assets ----
print('\n[1] STATIC ASSETS')
status, body, _ = get('/')
check('index page loads', status == 200, f'status={status}')
check('index has Source Sans', b'Source+Sans' in body or b'Source Sans' in body, 'font not found')
check('index has vacancy card JS', b'createVacancyCardElement' in body, 'function missing')
check('index has sort buttons', b'btn-history-sort' in body, 'sort buttons missing')

status, _, _ = get('/static/css/style.css')
check('style.css loads', status == 200, f'status={status}')
status, _, _ = get('/static/js/jszip.min.js')
check('jszip loads', status == 200, f'status={status}')
status, _, _ = get('/static/img/emblem.svg')
check('emblem.svg loads', status == 200, f'status={status}')

# ---- 2. Core GET APIs ----
print('\n[2] GET APIS')
status, body, _ = get('/api/health')
data = json.loads(body)
check('health check works', status == 200 and data.get('healthy') is True, f'status={status}')

status, body, _ = get('/api/regions')
data = json.loads(body)
check('regions returns list', isinstance(data, list), f'type={type(data)}')
check('regions not empty', len(data) > 0, f'count={len(data)}')

status, body, _ = get('/api/test-types')
data = json.loads(body)
check('test-types returns list', isinstance(data, list), f'type={type(data)}')

status, body, _ = get('/api/organizations')
data = json.loads(body)
check('organizations returns list', isinstance(data, list), f'type={type(data)}')
check('organizations not empty', len(data) > 0, f'count={len(data)}')

status, body, _ = get('/api/districts?region=26')
data = json.loads(body)
check('districts endpoint works', status == 200, f'status={status}')

status, body, _ = get('/api/history')
data = json.loads(body)
check('history returns list', isinstance(data, list), f'type={type(data)}')

# ---- 3. POST /api/vacancies ----
print('\n[3] POST /api/vacancies')
status, data, _ = post('/api/vacancies', {'page': 0, 'pageSize': 5})
check('vacancies returns 200', status == 200, f'status={status}')
check('vacancies has count field', 'count' in data, f'keys={list(data.keys())}')
check('vacancies has vacancies field', 'vacancies' in data, f'keys={list(data.keys())}')
total = data.get('count', 0)
items = data.get('vacancies', [])
check('vacancies total > 0', total > 0, f'total={total}')
check('vacancies items not empty', len(items) > 0, f'len={len(items)}')

if items:
    item = items[0]
    required_fields = ['id', 'position_name', 'organization', 'salary_display',
                       'region_display', 'date_start_uz', 'date_stop_uz', 'test_type_name']
    for f in required_fields:
        check(f'card has field: {f}', f in item, f'keys={list(item.keys())}')

# Search filter
status, data, _ = post('/api/vacancies', {'page': 0, 'pageSize': 5, 'search': 'bosh'})
check('vacancies search works', status == 200 and 'vacancies' in data, f'status={status}')

# Region filter
status, data, _ = post('/api/vacancies', {'page': 0, 'pageSize': 5, 'region': '26'})
check('vacancies region filter works', status == 200 and 'vacancies' in data, f'status={status}')

# Page 2
status, data2, _ = post('/api/vacancies', {'page': 1, 'pageSize': 5})
check('vacancies page 2 works', status == 200 and 'vacancies' in data2, f'status={status}')

# ---- 4. Scan API ----
print('\n[4] SCAN API')
status, _, _ = post('/api/stop')
check('stop endpoint works', status == 200, f'status={status}')

# ---- 5. Export ----
print('\n[5] EXPORT')
if items:
    real_id = items[0].get('id')
    status, exp_data, _ = post(f'/api/export-single/{real_id}')
    check(f'export-single/{real_id} works', status == 200, f'status={status}')
    check('export returns filename', 'filename' in exp_data, f'keys={list(exp_data.keys())}')
    if 'filename' in exp_data:
        fname = exp_data['filename']
        quoted_fname = urllib.parse.quote(fname)
        status2, _, _ = get(f'/api/download/{quoted_fname}')
        check('download exported file works', status2 == 200, f'status={status2}')

# 404 export
status, _, _ = post('/api/export-single/9999999')
check('non-existent export returns 404', status == 404, f'status={status}')

# ---- 6. Security ----
print('\n[6] SECURITY')
status, _, _ = get('/api/download/../app.py')
check('path traversal blocked (..)', status in (400, 403, 404), f'status={status}')
status, _, _ = get('/api/download/%2F..%2Fapp.py')
check('path traversal blocked (encoded)', status in (400, 403, 404), f'status={status}')

# ---- 7. Error handling ----
print('\n[7] ERROR HANDLING')
status, data, _ = post('/api/vacancies', {'page': 'invalid'})
check('vacancies handles invalid page gracefully', status == 200, f'status={status}')

# ---- Summary ----
print()
print('=' * 60)
print(f'NATIJA: {PASS} PASS, {FAIL} FAIL')
if ERRORS:
    print('\nXATOLIKLAR:')
    for e in ERRORS:
        print(f'  - {e}')
print('=' * 60)
