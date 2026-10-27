import datetime
import io
import json
import subprocess
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

pytest.importorskip('cryptography')
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
from PIL import Image

from backend.certificates import read_certificate, validate_kinds, screenshot_certificate
from backend.service import Workspace


@pytest.fixture(scope='module')
def certificate_data():
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, 'Screenshot Test CA')])
    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (x509.CertificateBuilder().subject_name(subject).issuer_name(subject).public_key(key.public_key())
            .serial_number(100).not_valid_before(now - datetime.timedelta(days=1))
            .not_valid_after(now + datetime.timedelta(days=1)).sign(key, hashes.SHA256()))
    return cert.public_bytes(serialization.Encoding.PEM)


def test_read_pem_and_der(certificate_data):
    der, metadata = read_certificate(certificate_data, 'test.crt')
    assert metadata['name'] == 'Screenshot Test CA'
    assert metadata['public_key'] == 'RSA (2048 bits)'
    assert len(metadata['sha256']) == 64
    assert read_certificate(der, 'test.cer')[1]['subject'] == metadata['subject']


def test_sm2_public_key_does_not_block_certificate_import():
    from cryptography.exceptions import UnsupportedAlgorithm
    raw = (Path(__file__).parent / 'fixtures' / 'sm2.crt').read_bytes()
    certificate = x509.load_pem_x509_certificate(raw)
    # This reproduces the unsupported curve from the user's server traceback.
    with pytest.raises(UnsupportedAlgorithm):
        certificate.public_key()
    der, metadata = read_certificate(raw, 'sm2.crt')
    assert metadata['name'] == 'Shitu SM2 Import Regression'
    assert metadata['public_key'].startswith('EC（')
    assert der == certificate.public_bytes(serialization.Encoding.DER)
    assert read_certificate(der, 'sm2.cer')[1]['sha256'] == metadata['sha256']


@pytest.mark.parametrize('data', [b'', b'broken', b'-----BEGIN PRIVATE KEY-----\nabc', b'x' * (2 * 1024 * 1024 + 1)], ids=['empty','broken','private-key','oversize'])
def test_invalid_certificate_is_rejected(data):
    with pytest.raises(ValueError):
        read_certificate(data, 'bad.crt')


@pytest.mark.parametrize('kinds', [[], ['unknown'], None, [['general']], 'general'])
def test_invalid_screenshot_choices(kinds):
    with pytest.raises(ValueError):
        validate_kinds(kinds)


def test_choices_keep_requested_order():
    assert validate_kinds(['public_key', 'general', 'public_key']) == ['public_key', 'general']


def test_worker_timeout_is_reported(monkeypatch):
    monkeypatch.setattr('backend.certificates.sys.platform', 'win32')
    monkeypatch.setattr(subprocess, 'CREATE_NO_WINDOW', 0, raising=False)
    def timeout(*args, **kwargs):
        raise subprocess.TimeoutExpired(args[0], 1)
    monkeypatch.setattr(subprocess, 'run', timeout)
    with pytest.raises(ValueError, match='超时'):
        screenshot_certificate(b'cert', {'name': 'Test'}, ['general'], timeout=1)


def test_worker_timeout_identifies_the_field_being_captured(monkeypatch):
    monkeypatch.setattr('backend.certificates.sys.platform', 'win32')
    monkeypatch.setattr(subprocess, 'CREATE_NO_WINDOW', 0, raising=False)
    def timeout(command, **kwargs):
        (Path(command[-1]).parent / 'progress.json').write_text(json.dumps({'kind':'field:证书策略'}), encoding='utf-8')
        raise subprocess.TimeoutExpired(command, 1)
    monkeypatch.setattr(subprocess, 'run', timeout)
    with pytest.raises(ValueError, match='超时') as error:
        screenshot_certificate(b'cert', {'name':'Test'}, ['general','field:证书策略'], timeout=1)
    assert error.value.failed_kind == 'field:证书策略'
    assert error.value.failed_field == '证书策略'


def test_worker_failure_imports_no_partial_images(monkeypatch):
    monkeypatch.setattr('backend.certificates.sys.platform', 'win32')
    monkeypatch.setattr(subprocess, 'CREATE_NO_WINDOW', 0, raising=False)
    def fail(command, **kwargs):
        root = Path(command[-1]).parent
        (root / 'result.json').write_text(json.dumps({'error': '未找到公钥字段', 'failed_kind': 'public_key'}), encoding='utf-8')
        return SimpleNamespace(returncode=1)
    monkeypatch.setattr(subprocess, 'run', fail)
    with pytest.raises(ValueError, match='公钥字段') as error:
        screenshot_certificate(b'cert', {'name': 'Test'}, ['public_key'])
    assert error.value.failed_kind == 'public_key'
    assert error.value.failed_field == '公钥'


def test_capture_bridge_retains_existing_images_on_import_failure(certificate_data, monkeypatch):
    pytest.importorskip('webview')
    from desktop import DesktopBridge
    workspace = Workspace()
    image = io.BytesIO()
    Image.new('RGB', (10, 10), 'red').save(image, 'PNG')
    existing = workspace.add(image.getvalue(), 'existing.png')
    bridge = DesktopBridge(SimpleNamespace(state=SimpleNamespace(workspace=workspace, lock=threading.RLock())))
    bridge._certificate = read_certificate(certificate_data, 'test.crt')
    monkeypatch.setattr('backend.certificates.screenshot_certificate', lambda *args: [(image.getvalue(), 'good.png'), (b'invalid', 'bad.png')])
    try:
        assert 'error' in bridge.capture_certificate(['general', 'issuer'])
        assert list(workspace.images) == [existing]
        assert not bridge._capture_lock.locked()
    finally:
        workspace.close()


def test_certificate_bundle_is_rejected(certificate_data):
    with pytest.raises(ValueError):
        read_certificate(certificate_data + certificate_data, 'chain.pem')


def test_custom_worker_results_have_safe_names_and_keep_requested_fields(monkeypatch):
    monkeypatch.setattr('backend.certificates.sys.platform', 'win32')
    monkeypatch.setattr(subprocess, 'CREATE_NO_WINDOW', 0, raising=False)
    def complete(command, **kwargs):
        root = Path(command[-1]).parent
        request = json.loads((root / 'request.json').read_text(encoding='utf-8'))
        assert request['kinds'] == ['serial_number', 'field:custom/field']
        records = []
        for number, kind in enumerate(request['kinds']):
            filename = f'{number}.png'
            Image.new('RGB', (10, 10), 'green').save(root / filename)
            records.append({'kind':kind, 'file':filename})
        (root / 'result.json').write_text(json.dumps({'screenshots':records}), encoding='utf-8')
        return SimpleNamespace(returncode=0)
    monkeypatch.setattr(subprocess, 'run', complete)
    pictures = screenshot_certificate(b'cert', {'name':'x' * 110}, ['serial_number', 'field:custom/field'])
    assert len(pictures) == 2 and all(len(name) < 180 for _, name in pictures)
    assert pictures[1][1].endswith('_custom field.png')
