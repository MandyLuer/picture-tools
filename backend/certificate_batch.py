"""Bounded in-memory certificate queue and sequential native screenshot jobs."""
import copy
import sys
import threading
import uuid

from fastapi import File, HTTPException, UploadFile
from pydantic import BaseModel, Field
from cryptography.exceptions import UnsupportedAlgorithm

from .certificates import MAX_CERTIFICATE, read_certificate, validate_kinds, screenshot_certificate


class CaptureBatch(BaseModel):
    certificates: list[str] = Field(min_length=1, max_length=50)
    kinds: list[str] = Field(min_length=1, max_length=24)


def register_certificate_routes(app):
    lock = app.state.lock
    certificates, jobs = {}, {}

    def active():
        return any(job['status'] in ('running', 'cancelling') for job in jobs.values())

    @app.post('/api/certificates/import')
    def import_certificates(files: list[UploadFile] = File(...)):
        if len(files) > 50:
            raise ValueError('每批最多添加 50 张证书')
        accepted, errors = [], []
        with lock:
            if active():
                raise HTTPException(409, '截图期间不能修改证书队列')
            for file in files:
                filename = file.filename or 'certificate.crt'
                try:
                    der, metadata = read_certificate(file.file.read(MAX_CERTIFICATE + 1), filename)
                    if any(item[1]['sha256'] == metadata['sha256'] for item in certificates.values()):
                        raise ValueError('队列中已存在相同证书')
                    if len(certificates) >= 50:
                        raise ValueError('证书队列最多 50 张，请先移除不需要的项目')
                    key = uuid.uuid4().hex
                    metadata = {**metadata, 'id': key}
                    certificates[key] = (der, metadata)
                    accepted.append(metadata)
                except (ValueError, UnsupportedAlgorithm) as exc:
                    errors.append({'filename': filename, 'error': str(exc)})
        return {'certificates': accepted, 'errors': errors}

    @app.get('/api/certificates')
    def listing():
        with lock:
            return [item[1] for item in certificates.values()]

    @app.delete('/api/certificates')
    def clear_certificates():
        with lock:
            if active():
                raise HTTPException(409, '证书截图任务仍在运行，请等待完成或停止后续处理后再开始新任务')
            certificates.clear()
        return {'count': 0}

    @app.delete('/api/certificates/{key}')
    def remove_certificate(key: str):
        with lock:
            if active():
                raise HTTPException(409, '截图期间不能修改证书队列')
            if key not in certificates:
                raise HTTPException(404, '证书不存在，请刷新队列')
            del certificates[key]
        return {'removed': key}

    def execute(job, inputs, generation):
        try:
            for row, (der, metadata) in zip(job['results'], inputs):
                with lock:
                    if job['cancel_requested']:
                        break
                    row['status'] = 'running'
                try:
                    stage = 'capture'
                    pictures = screenshot_certificate(der, metadata, job['kinds'])
                    stage = 'import'
                    with lock:
                        workspace = app.state.workspace
                        if workspace.root != generation:
                            raise ValueError('工作区已变更，本次证书截图未导入')
                        added = []
                        try:
                            for data, name in pictures:
                                added.append(workspace.add(data, name))
                        except Exception:
                            for key in added:
                                workspace.remove(key)
                            raise
                        row.update(status='success', count=len(added))
                        job['count'] += len(added)
                        job['selected'] = job['selected'] or (added[0] if added else '')
                except Exception as exc:
                    with lock:
                        row.update(status='failed', error=str(exc),
                                   failed_kind=getattr(exc, 'failed_kind', ''),
                                   failed_field=getattr(exc, 'failed_field', ''), failed_stage=stage)
                with lock:
                    job['completed'] += 1
        finally:
            with lock:
                for row in job['results']:
                    if row['status'] == 'pending':
                        row['status'] = 'cancelled'
                job['status'] = 'cancelled' if job['cancel_requested'] else 'completed'

    @app.post('/api/certificate-jobs')
    def start_job(body: CaptureBatch):
        kinds = validate_kinds(body.kinds)
        if sys.platform != 'win32':
            raise ValueError('原生证书截图需要在 Windows 本机运行 Python 后端')
        with lock:
            if active():
                raise HTTPException(409, '已有证书截图任务正在运行')
            keys = list(dict.fromkeys(body.certificates))
            if any(key not in certificates for key in keys):
                raise ValueError('证书队列已变化，请重新加载证书')
            inputs = [certificates[key] for key in keys]
            key = uuid.uuid4().hex
            job = {'id': key, 'status': 'running', 'kinds': kinds, 'total': len(keys),
                   'completed': 0, 'count': 0, 'selected': '', 'cancel_requested': False,
                   'results': [{'id': metadata['id'], 'filename': metadata['filename'],
                                'status': 'pending', 'count': 0, 'error': ''} for _, metadata in inputs]}
            # Keep only the latest job; snapshots contain no certificate bytes or paths.
            jobs.clear()
            jobs[key] = job
            snapshot = copy.deepcopy(job)
            threading.Thread(target=execute, args=(job, inputs, app.state.workspace.root), daemon=True).start()
            return snapshot

    @app.get('/api/certificate-jobs/{key}')
    def job_status(key: str):
        with lock:
            if key not in jobs:
                raise HTTPException(404, '截图任务不存在或已过期')
            return copy.deepcopy(jobs[key])

    @app.get('/api/certificate-jobs')
    def latest_job():
        with lock:
            return copy.deepcopy(next(iter(jobs.values()), None))

    @app.post('/api/certificate-jobs/{key}/cancel')
    def cancel_job(key: str):
        with lock:
            if key not in jobs:
                raise HTTPException(404, '截图任务不存在或已过期')
            job = jobs[key]
            if job['status'] in ('running', 'cancelling'):
                job.update(cancel_requested=True, status='cancelling')
            return copy.deepcopy(job)
