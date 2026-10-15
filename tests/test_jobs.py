import os
import time

from api.jobs import JobStore


def _store(tmp_path):
    out = tmp_path / 'output'
    up = tmp_path / 'uploads'
    out.mkdir()
    up.mkdir()
    return JobStore(out / 'jobs', out, up), out, up


def test_finished_job_survives_restart(tmp_path):
    store, out, up = _store(tmp_path)
    store.create('j1', status='processing', progress=0)
    store.update('j1', status='done', progress=100, output_name='x.xlsx')

    again = JobStore(out / 'jobs', out, up)         # simulated restart
    assert again.get('j1')['status'] == 'done'
    assert again.get('j1')['output_name'] == 'x.xlsx'


def test_running_job_is_failed_after_restart(tmp_path):
    store, out, up = _store(tmp_path)
    store.create('j2', status='processing', step='categorising', progress=40)

    again = JobStore(out / 'jobs', out, up)
    job = again.get('j2')
    assert job['status'] == 'error'
    assert 'restarted' in job['message']


def test_unknown_job_is_none(tmp_path):
    store, _, _ = _store(tmp_path)
    assert store.get('nope') is None


def test_cleanup_removes_only_old_files(tmp_path):
    store, out, up = _store(tmp_path)
    store.create('old', status='done')
    store.create('new', status='done')
    old_xlsx, new_xlsx = out / 'old.xlsx', out / 'new.xlsx'
    old_up = up / 'old_statement.csv'
    for f in (old_xlsx, new_xlsx, old_up):
        f.write_text('x')
    two_days_ago = time.time() - 48 * 3600
    for f in (old_xlsx, old_up, out / 'jobs' / 'old.json'):
        os.utime(f, (two_days_ago, two_days_ago))

    assert store.cleanup(max_age_hours=24) == 3
    assert not old_xlsx.exists() and not old_up.exists()
    assert new_xlsx.exists()
    assert store.get('old') is None
    assert store.get('new')['status'] == 'done'
