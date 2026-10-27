import datetime
import io
import threading
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from PIL import Image
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

from backend.app import create_app
from backend.certificates import validate_kinds, find_field_index, ScreenshotFailure


def cert_data(number):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, 'Batch CA')])
    now = datetime.datetime.now(datetime.timezone.utc)
    return (x509.CertificateBuilder().subject_name(subject).issuer_name(subject).public_key(key.public_key())
            .serial_number(number).not_valid_before(now).not_valid_after(now + datetime.timedelta(days=1))
            .sign(key, hashes.SHA256()).public_bytes(serialization.Encoding.PEM))


@pytest.fixture
def client():
    with TestClient(create_app()) as client:
        yield client


def add(client, *numbers):
    return client.post('/api/certificates/import', files=[('files', (f'{number}.crt', cert_data(number))) for number in numbers]).json()['certificates']


def wait_job(client, key):
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        job = client.get('/api/certificate-jobs/' + key).json()
        if job['status'] not in ('running', 'cancelling'):
            return job
        time.sleep(.01)
    pytest.fail('Job did not finish')


def picture():
    output = io.BytesIO()
    Image.new('RGB', (20, 20), 'green').save(output, 'PNG')
    return output.getvalue(), 'Batch CA_序列号.png'


def test_custom_fields_are_exact_and_localized():
    assert validate_kinds(['general', 'serial_number', 'field:证书策略', 'serial_number']) == ['general', 'serial_number', 'field:证书策略']
    assert find_field_index(['Serial number', 'Public Key'], 'serial_number') == 0
    assert find_field_index(['Serial number', 'Public Key'], 'field:public key') == 1
    assert find_field_index(['版本', '证书策略'], 'field:证书策略') == 1
    with pytest.raises(ValueError, match='未找到唯一字段'):
        find_field_index(['证书策略编号'], 'field:证书策略')
    for fields in [['field:'], ['field: x'], ['field:a\n'], ['field:' + 'x' * 81], ['general'] * 25]:
        with pytest.raises(ValueError):
            validate_kinds(fields)


def test_import_rejects_invalid_and_duplicate_without_losing_valid_certificates(client):
    data = cert_data(1)
    result = client.post('/api/certificates/import', files=[('files', ('good.crt', data)), ('files', ('bad.crt', b'broken')), ('files', ('duplicate.crt', data))]).json()
    assert len(result['certificates']) == 1
    assert len(result['errors']) == 2
    assert len(client.get('/api/certificates').json()) == 1
    key = result['certificates'][0]['id']
    assert client.delete('/api/certificates/' + key).status_code == 200
    assert client.get('/api/certificates').json() == []


def test_sm2_and_rsa_batch_import_returns_json_and_preserves_both(client):
    sm2 = (Path(__file__).parent / 'fixtures' / 'sm2.crt').read_bytes()
    response = client.post('/api/certificates/import', files=[('files', ('sm2.crt', sm2)), ('files', ('rsa.crt', cert_data(11)))])
    assert response.status_code == 200
    assert response.json()['errors'] == []
    assert [item['filename'] for item in response.json()['certificates']] == ['sm2.crt', 'rsa.crt']
    assert len(client.get('/api/certificates').json()) == 2


def test_clear_certificate_queue_preserves_generated_images_and_allows_reimport(client):
    data = cert_data(123)
    client.post('/api/certificates/import', files=[('files', ('previous.crt', data))])
    raw, name = picture()
    client.post('/api/import', files=[('files', (name, raw))])
    before = client.get('/api/images').json()
    assert client.delete('/api/certificates').json() == {'count':0}
    assert client.get('/api/certificates').json() == []
    assert client.get('/api/images').json() == before
    result = client.post('/api/certificates/import', files=[('files', ('previous.crt', data))]).json()
    assert len(result['certificates']) == 1 and result['errors'] == []


def test_unexpected_import_errors_return_json_without_exposing_internal_details(monkeypatch):
    def unexpected(*args):
        raise RuntimeError('private server implementation detail')
    monkeypatch.setattr('backend.certificate_batch.read_certificate', unexpected)
    with TestClient(create_app(), raise_server_exceptions=False) as client:
        response = client.post('/api/certificates/import', files=[('files', ('test.crt', b'cert'))])
        assert response.status_code == 500
        assert response.headers['content-type'].startswith('application/json')
        assert '服务端处理失败' in response.json()['detail']
        assert 'private server' not in response.text


def test_batch_continues_after_failure_and_rolls_back_only_failed_certificate(client, monkeypatch):
    entries = add(client, 1, 2, 3)
    calls = []
    def capture(der, metadata, kinds):
        calls.append(metadata['filename'])
        assert kinds == ['serial_number', 'field:证书策略']
        return [picture(), (b'invalid', 'bad.png')] if metadata['filename'] == '2.crt' else [picture()]
    monkeypatch.setattr('backend.certificate_batch.screenshot_certificate', capture)
    job = client.post('/api/certificate-jobs', json={'certificates': [item['id'] for item in entries], 'kinds': ['serial_number', 'field:证书策略']}).json()
    result = wait_job(client, job['id'])
    assert calls == ['1.crt', '2.crt', '3.crt']
    assert result['completed'] == 3 and result['count'] == 2
    assert [row['status'] for row in result['results']] == ['success', 'failed', 'success']
    images = client.get('/api/images').json()
    assert len(images) == 2 and len({image['name'] for image in images}) == 2


def test_cancel_waits_for_current_certificate_and_blocks_concurrent_jobs(client, monkeypatch):
    entries = add(client, 1, 2)
    started, finish = threading.Event(), threading.Event()
    def capture(*args):
        started.set()
        assert finish.wait(3)
        return [picture()]
    monkeypatch.setattr('backend.certificate_batch.screenshot_certificate', capture)
    payload = {'certificates': [item['id'] for item in entries], 'kinds': ['general']}
    job = client.post('/api/certificate-jobs', json=payload).json()
    try:
        assert started.wait(2)
        assert client.post('/api/certificate-jobs', json=payload).status_code == 409
        assert client.delete('/api/certificates/' + entries[0]['id']).status_code == 409
        assert client.delete('/api/certificates').status_code == 409
        assert client.post('/api/certificate-jobs/' + job['id'] + '/cancel').json()['status'] == 'cancelling'
    finally:
        finish.set()
    result = wait_job(client, job['id'])
    assert result['status'] == 'cancelled' and result['count'] == 1
    assert [row['status'] for row in result['results']] == ['success', 'cancelled']


def test_clear_workspace_during_capture_prevents_stale_import(client, monkeypatch):
    entries = add(client, 1)
    started, finish = threading.Event(), threading.Event()
    def capture(*args):
        started.set()
        assert finish.wait(3)
        return [picture()]
    monkeypatch.setattr('backend.certificate_batch.screenshot_certificate', capture)
    job = client.post('/api/certificate-jobs', json={'certificates': [entries[0]['id']], 'kinds': ['general']}).json()
    try:
        assert started.wait(2)
        assert client.delete('/api/images').status_code == 200
    finally:
        finish.set()
    result = wait_job(client, job['id'])
    assert result['results'][0]['status'] == 'failed' and result['count'] == 0
    assert client.get('/api/images').json() == []


def test_batch_auth_and_unknown_ids_are_rejected():
    with TestClient(create_app('secret')) as client:
        assert client.get('/api/certificates').status_code == 403
        assert client.post('/api/certificate-jobs', headers={'x-app-token': 'secret'}, json={'certificates': ['missing'], 'kinds': ['general']}).status_code == 400


def test_failed_field_and_reason_are_preserved_in_job_result(client, monkeypatch):
    entries = add(client, 1)
    def capture(*args):
        raise ScreenshotFailure('当前证书未包含此字段', 'field:证书策略')
    monkeypatch.setattr('backend.certificate_batch.screenshot_certificate', capture)
    job = client.post('/api/certificate-jobs', json={'certificates':[entries[0]['id']], 'kinds':['general','field:证书策略']}).json()
    result = wait_job(client, job['id'])
    row = result['results'][0]
    assert row['filename'] == '1.crt'
    assert row['failed_field'] == '证书策略'
    assert row['failed_kind'] == 'field:证书策略'
    assert row['failed_stage'] == 'capture'
    assert row['error'] == '当前证书未包含此字段'
