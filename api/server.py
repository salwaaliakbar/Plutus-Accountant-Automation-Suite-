"""Plutus Milestone 1 – FastAPI backend.

Endpoints:
  POST /api/upload          – upload bank statement + template, starts background job
  GET  /api/progress/{id}   – poll job status
  GET  /api/download/{id}   – download processed Excel when status == 'done'
  GET  /api/health          – health check
"""

import sys
import threading
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

load_dotenv()

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from main import _infer_year  # noqa: E402
from api.jobs import JobStore  # noqa: E402

OUTPUT_DIR = ROOT / 'output'
UPLOAD_DIR = ROOT / 'uploads'

OUTPUT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

JOBS = JobStore(OUTPUT_DIR / 'jobs', OUTPUT_DIR, UPLOAD_DIR)
JOBS.cleanup()

app = FastAPI(title='Plutus Accountant API')

app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_methods=['*'],
    allow_headers=['*'],
)


# ── Background worker ─────────────────────────────────────────────────────────

def _run_pipeline(job_id: str, file_path: Path, original_name: str, template_path: Path) -> None:
    try:
        _update(job_id, step='parsing', progress=10, message='Parsing bank statement…')
        suffix = file_path.suffix.lower()
        if suffix == '.csv':
            from parsers.parse_csv import parse_csv
            transactions = parse_csv(file_path)
        elif suffix == '.pdf':
            from parsers.parse_pdf import parse_pdf
            year_hint = _infer_year(original_name)
            transactions = parse_pdf(file_path, year_hint=year_hint)
        elif suffix in ('.xlsx', '.xls'):
            from parsers.parse_xls import parse_xls
            transactions = parse_xls(file_path)
        else:
            raise ValueError(f'Unsupported file type: {suffix}')

        if not transactions:
            raise ValueError('No transactions found in the uploaded file.')
        from_scan = any(t.get('from_scan') for t in transactions)

        from core.excel_writer import prepare_new_transactions
        total = len(transactions)
        first, last = transactions[0]['date'], transactions[-1]['date']
        transactions, dropped = prepare_new_transactions(transactions, template_path)
        n = len(transactions)
        if n == 0:
            raise ValueError(
                f'All {total} transactions in this statement ({first:%d %b %Y} to '
                f'{last:%d %b %Y}) are already in the workbook, so there is nothing '
                'to add. If you expected new rows, check that this is the right '
                'statement and that the template is the file from before this period.')
        skipped = f' ({dropped} already in the workbook, skipped)' if dropped else ''

        _update(job_id, step='categorising', progress=40,
                message=f'Parsed {n} new transactions{skipped}. Categorising with AI…')
        from core.categorise import categorise_detailed
        from core.excel_writer import extract_client_examples, extract_analysis_vocabulary
        client_examples = extract_client_examples(template_path)
        categories, review = categorise_detailed(
            transactions, client_examples=client_examples,
            analysis_vocab=extract_analysis_vocabulary(template_path))

        _update(job_id, step='writing', progress=75, message='Categorised. Writing to Excel…')
        from core.excel_writer import write_workbook
        stem = Path(original_name).stem
        output_path = OUTPUT_DIR / f"{job_id}_{stem}_processed.xlsx"
        template_warnings: list = []
        if from_scan:
            template_warnings.append(
                'This PDF is a scan, so the transactions were read from the page images '
                'by AI. Check every date and amount against the statement before relying on them.')
        problems = write_workbook(transactions, categories, template_path, output_path,
                                  review=review, warnings=template_warnings)

        from collections import Counter
        cat_summary = dict(Counter(categories).most_common())     # every category, not a top few
        warn = ''
        if problems:
            warn = (f' WARNING: {len(problems)} balance(s) do not add up — check these '
                    f'rows against the statement: ' + '; '.join(problems[:3]))
        _update(job_id, step='done', progress=100, status='done',
                message=(f'Complete! {n} transactions processed{skipped}. '
                         f'{sum(review)} highlighted in yellow for checking.{warn}'),
                balance_problems=problems,
                template_warnings=template_warnings,
                transaction_count=n,
                review_count=sum(review),
                category_count=len(cat_summary),
                top_categories=cat_summary,
                output_path=str(output_path),
                output_name=output_path.name)

    except Exception as exc:
        JOBS.update(job_id, status='error', progress=0, message=str(exc), step='error')
    finally:
        try:
            file_path.unlink(missing_ok=True)
        except Exception:
            pass
        try:
            template_path.unlink(missing_ok=True)
        except Exception:
            pass


def _update(job_id: str, **kwargs) -> None:
    JOBS.update(job_id, **kwargs)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.post('/api/upload')
async def upload(
    file: UploadFile = File(..., description='Bank statement (CSV, PDF, XLSX, XLS)'),
    template: UploadFile = File(..., description='Excel template workbook (XLSX)'),
):
    name = file.filename or ''
    ext  = Path(name).suffix.lower()
    if ext not in ('.csv', '.pdf', '.xlsx', '.xls'):
        raise HTTPException(status_code=400,
                            detail='Bank statement must be CSV, PDF, XLSX or XLS.')

    tmpl_name = template.filename or ''
    if not tmpl_name.lower().endswith('.xlsx'):
        raise HTTPException(status_code=400,
                            detail='Template must be an XLSX file.')

    JOBS.cleanup()      # drop files past the retention period
    job_id        = str(uuid.uuid4())
    saved_path    = UPLOAD_DIR / f"{job_id}_statement{ext}"
    template_path = UPLOAD_DIR / f"{job_id}_template.xlsx"

    content = await file.read()
    with open(saved_path, 'wb') as fh:
        fh.write(content)

    tmpl_content = await template.read()
    with open(template_path, 'wb') as fh:
        fh.write(tmpl_content)

    JOBS.create(job_id, status='processing', step='queued', progress=0,
                message='Queued…', filename=name)

    t = threading.Thread(
        target=_run_pipeline,
        args=(job_id, saved_path, name, template_path),
        daemon=True,
    )
    t.start()

    return {'job_id': job_id, 'filename': name}


@app.get('/api/progress/{job_id}')
async def progress(job_id: str):
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail='Job not found.')
    return job


@app.get('/api/download/{job_id}')
async def download(job_id: str):
    job = JOBS.get(job_id)
    if job is None or job.get('status') != 'done':
        raise HTTPException(status_code=404, detail='File not ready yet.')
    output_path = job.get('output_path')
    if not output_path or not Path(output_path).exists():
        raise HTTPException(status_code=500, detail='Output file missing.')
    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=job['output_name'],
    )


@app.get('/api/health')
async def health():
    return {'status': 'ok'}
