"""Certificate metadata and isolated Windows screenshot jobs."""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from cryptography import x509
from cryptography.exceptions import UnsupportedAlgorithm
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec, rsa, ed25519, ed448
from cryptography.x509.oid import NameOID

from .document_context import filename_from_text

SCREENSHOTS = {'general': '常规信息', 'issuer': '颁发者', 'public_key': '公钥', 'path': '证书路径'}
FIELD_ALIASES = {
    'issuer': ('颁发者', '發行者', '发行者', 'Issuer'),
    'public_key': ('公钥', '公開金鑰', '公用金鑰', 'Public key'),
    'serial_number': ('序列号', '序號', 'Serial number'),
    'signature_algorithm': ('签名算法', '簽章演算法', 'Signature algorithm'),
    'subject': ('使用者', '主题', '主體', 'Subject'),
    'valid_from': ('从', '有效期从', '有效期開始', 'Valid from'),
    'valid_to': ('到', '有效期至', '有效期截止', 'Valid to'),
    'thumbprint': ('指纹', '指紋', 'Thumbprint'),
    'key_usage': ('密钥用法', '金鑰使用方法', 'Key usage'),
    'enhanced_key_usage': ('增强型密钥用法', '增強金鑰使用方法', 'Enhanced key usage'),
    'subject_alt_name': ('使用者可选名称', '主题备用名称', '主體別名', 'Subject alternative name'),
    'basic_constraints': ('基本约束', '基本限制', 'Basic constraints'),
}
SCREENSHOTS.update({key: aliases[0] for key, aliases in FIELD_ALIASES.items() if key not in SCREENSHOTS})
MAX_CERTIFICATE = 2 * 1024 * 1024


def describe_public_key(cert):
    try:
        key = cert.public_key()
    except UnsupportedAlgorithm:
        # Metadata decoding must not prevent Windows from viewing the original
        # certificate (notably EC certificates using the SM2 curve).
        oid = cert.public_key_algorithm_oid.dotted_string
        algorithm = {'1.2.840.10045.2.1': 'EC', '1.2.156.10197.1.301': 'SM2'}.get(oid, f'OID {oid}')
        return f'{algorithm}（当前解析库不支持此公钥参数，原生截图以 Windows 显示为准）'
    if isinstance(key, rsa.RSAPublicKey):
        return f'RSA ({key.key_size} bits)'
    if isinstance(key, ec.EllipticCurvePublicKey):
        return f'EC {key.curve.name} ({key.key_size} bits)'
    if isinstance(key, (ed25519.Ed25519PublicKey, ed448.Ed448PublicKey)):
        return type(key).__name__.replace('PublicKey', '')
    return type(key).__name__


def read_certificate(data, filename):
    if not data or len(data) > MAX_CERTIFICATE:
        raise ValueError('证书文件为空或超过 2 MB')
    try:
        if b'-----BEGIN' in data:
            if data.count(b'-----BEGIN CERTIFICATE-----') != 1 or b'PRIVATE KEY' in data:
                raise ValueError('请选择单张公钥证书')
            cert = x509.load_pem_x509_certificate(data)
        else:
            cert = x509.load_der_x509_certificate(data)
    except ValueError as exc:
        raise ValueError('无法解析证书，请选择 PEM 或 DER 格式的 .crt/.cer/.pem 公钥证书') from exc
    common_names = cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)
    common_name = common_names[0].value if common_names else Path(filename).stem
    public_key = describe_public_key(cert)
    return cert.public_bytes(serialization.Encoding.DER), {
        'filename': Path(filename).name, 'name': common_name,
        'subject': cert.subject.rfc4514_string(), 'issuer': cert.issuer.rfc4514_string(),
        'public_key': public_key, 'not_before': cert.not_valid_before_utc.isoformat(),
        'not_after': cert.not_valid_after_utc.isoformat(),
        'sha256': cert.fingerprint(hashes.SHA256()).hex(),
    }


def validate_kinds(kinds):
    if not isinstance(kinds, list) or not kinds or len(kinds) > 24:
        raise ValueError('请选择 1 至 24 项截图项目')
    if any(not isinstance(kind, str) or (kind not in SCREENSHOTS and not (
            kind.startswith('field:') and 1 <= len(kind[6:].strip()) <= 80
            and kind[6:] == kind[6:].strip() and not any(ord(c) < 32 for c in kind))) for kind in kinds):
        raise ValueError('请至少选择一项有效的截图项目')
    return list(dict.fromkeys(kinds))


def screenshot_label(kind):
    return SCREENSHOTS.get(kind, kind.removeprefix('field:'))


class ScreenshotFailure(ValueError):
    def __init__(self, message, kind=None):
        super().__init__(message)
        self.failed_kind = kind or ''
        self.failed_field = screenshot_label(kind) if kind else ''


def find_field_index(names, kind):
    """Only exact, case-insensitive native field names; never fuzzy UI clicks."""
    aliases = FIELD_ALIASES.get(kind, (kind.removeprefix('field:'),))
    normalized = {alias.strip().casefold() for alias in aliases}
    indices = [i for i, name in enumerate(names) if name.strip().casefold() in normalized]
    if len(indices) != 1:
        raise ValueError(f'未找到唯一字段“{screenshot_label(kind)}”，请检查字段名称、系统语言及证书是否包含此字段')
    return indices[0]


def screenshot_certificate(der, metadata, kinds, timeout=75):
    if sys.platform != 'win32':
        raise ValueError('原生证书截图仅支持 Windows 桌面客户端')
    kinds = validate_kinds(kinds)
    with tempfile.TemporaryDirectory(prefix='shitu-certificate-') as folder:
        root = Path(folder)
        (root / 'certificate.der').write_bytes(der)
        request = root / 'request.json'
        request.write_text(json.dumps({'kinds': kinds}), encoding='utf-8')
        command = ([sys.executable, '--certificate-worker', str(request)] if getattr(sys, 'frozen', False)
                   else [sys.executable, '-m', 'backend.certificate_worker', str(request)])
        try:
            completed = subprocess.run(command, timeout=timeout, stdout=subprocess.DEVNULL,
                                       stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired as exc:
            kind = None
            try:
                progress = json.loads((root / 'progress.json').read_text(encoding='utf-8'))
                if progress.get('kind') in kinds:
                    kind = progress['kind']
            except (OSError, ValueError):
                pass
            raise ScreenshotFailure('证书截图超时，已停止独立任务。请确认桌面已解锁后重试', kind) from exc
        result_file = root / 'result.json'
        if not result_file.exists():
            raise ValueError('截图进程未完成，请确认桌面可交互且依赖安装完整')
        result = json.loads(result_file.read_text(encoding='utf-8'))
        if completed.returncode or result.get('error'):
            raise ScreenshotFailure(result.get('error') or '证书截图失败', result.get('failed_kind'))
        records = result.get('screenshots', [])
        if [record['kind'] for record in records] != kinds:
            raise ValueError('截图项目不完整，本次结果未导入')
        stem = filename_from_text(metadata['name']) or '证书'
        images = []
        for record in records:
            path = (root / record['file']).resolve()
            if path.parent != root.resolve() or path.suffix != '.png':
                raise ValueError('截图结果路径无效')
            label = filename_from_text(screenshot_label(record['kind'])) or '字段'
            filename = f'{stem}_{label}'[:170].rstrip(' .') + '.png'
            images.append((path.read_bytes(), filename))
        return images
