"""Job tracking that survives a server restart, plus clean-up of old files.

Each job is kept in memory for fast polling and mirrored to
output/jobs/<job_id>.json. On start-up the saved jobs are loaded again, so a
finished workbook can still be downloaded after a restart. A job that was
still running when the server stopped is marked as failed (its worker thread
is gone) so the page does not poll forever.

Files older than the retention period (default 24 h, set PLUTUS_RETENTION_HOURS
to change it) are deleted: processed workbooks, saved job records and any
leftover uploads.
"""

import json
import os
import threading
import time
from pathlib import Path

RETENTION_HOURS = float(os.environ.get('PLUTUS_RETENTION_HOURS', '24'))


class JobStore:
    def __init__(self, jobs_dir: Path, output_dir: Path, upload_dir: Path):
        self.jobs_dir = Path(jobs_dir)
        self.output_dir = Path(output_dir)
        self.upload_dir = Path(upload_dir)
        self.jobs_dir.mkdir(parents=True, exist_ok=True)
        self._jobs: dict[str, dict] = {}
        self._lock = threading.Lock()
        self._load()

    # ── persistence ──────────────────────────────────────────────────────────

    def _path(self, job_id: str) -> Path:
        return self.jobs_dir / f'{job_id}.json'

    def _save(self, job_id: str) -> None:
        job = self._jobs[job_id]
        tmp = self._path(job_id).with_suffix('.tmp')
        try:
            tmp.write_text(json.dumps(job), encoding='utf-8')
            tmp.replace(self._path(job_id))     # atomic: never a half-written file
        except OSError as exc:
            print(f'[jobs] could not save job {job_id}: {exc}')

    def _load(self) -> None:
        for f in self.jobs_dir.glob('*.json'):
            try:
                job = json.loads(f.read_text(encoding='utf-8'))
            except (OSError, ValueError):
                continue
            if job.get('status') == 'processing':
                job.update(status='error', step='error', progress=0,
                           message='The server restarted while this file was '
                                   'processing. Please upload it again.')
                self._jobs[f.stem] = job
                self._save(f.stem)
            else:
                self._jobs[f.stem] = job

    # ── public API ───────────────────────────────────────────────────────────

    def create(self, job_id: str, **fields) -> None:
        with self._lock:
            self._jobs[job_id] = dict(fields)
            self._save(job_id)

    def update(self, job_id: str, **fields) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return
            job.update(fields)
            job.setdefault('status', 'processing')
            self._save(job_id)

    def get(self, job_id: str):
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job is not None else None

    # ── clean-up ─────────────────────────────────────────────────────────────

    def cleanup(self, max_age_hours: float = None) -> int:
        """Delete workbooks, job records and uploads older than the retention
        period. Returns how many files were removed."""
        hours = RETENTION_HOURS if max_age_hours is None else max_age_hours
        cutoff = time.time() - hours * 3600
        removed = 0

        for directory, pattern in ((self.output_dir, '*.xlsx'),
                                   (self.jobs_dir, '*.json'),
                                   (self.upload_dir, '*')):
            for f in directory.glob(pattern):
                try:
                    if f.is_file() and f.stat().st_mtime < cutoff:
                        f.unlink()
                        removed += 1
                        if directory is self.jobs_dir:
                            with self._lock:
                                self._jobs.pop(f.stem, None)
                except OSError:
                    pass
        return removed
