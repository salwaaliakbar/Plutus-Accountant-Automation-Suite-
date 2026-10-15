import time

import pytest

pytest.importorskip('httpx')
from fastapi.testclient import TestClient

from api import server


@pytest.fixture
def client(monkeypatch, tmp_path):
    out, up = tmp_path / 'output', tmp_path / 'uploads'
    out.mkdir()
    up.mkdir()
    store = server.JobStore(out / 'jobs', out, up)
    monkeypatch.setattr(server, 'JOBS', store)
    monkeypatch.setattr(server, 'OUTPUT_DIR', out)
    monkeypatch.setattr(server, 'UPLOAD_DIR', up)
    monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)
    return TestClient(server.app)


def _wait(client, job_id, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        job = client.get(f'/api/progress/{job_id}').json()
        if job['status'] != 'processing':
            return job
        time.sleep(0.1)
    raise AssertionError('job did not finish')


def test_rejects_unsupported_statement_type(client, make_template):
    with open(make_template(), 'rb') as t:
        r = client.post('/api/upload', files={'file': ('a.txt', b'x'), 'template': ('t.xlsx', t.read())})
    assert r.status_code == 400


@pytest.mark.parametrize('kind', ['csv', 'xlsx', 'xls'])
def test_upload_process_download(client, make_csv, make_xlsx, make_xls, make_template, kind):
    path = {'csv': make_csv, 'xlsx': make_xlsx, 'xls': make_xls}[kind]()
    files = {'file': (path.name, path.read_bytes()),
             'template': ('t.xlsx', make_template().read_bytes())}
    job_id = client.post('/api/upload', files=files).json()['job_id']

    job = _wait(client, job_id)
    assert job['status'] == 'done', job['message']
    assert job['transaction_count'] == 5

    r = client.get(f'/api/download/{job_id}')
    assert r.status_code == 200 and r.content[:2] == b'PK'


def test_unreadable_statement_reports_an_error(client, make_template):
    files = {'file': ('s.csv', b'nothing useful here\n'),
             'template': ('t.xlsx', make_template().read_bytes())}
    job = _wait(client, client.post('/api/upload', files=files).json()['job_id'])
    assert job['status'] == 'error' and job['message']
