"""AI provider profiles. Secret values never appear in browser-facing settings."""
import base64
import copy
import json
import os
from pathlib import Path
import sys
import threading
import uuid
from urllib.parse import urlsplit

from pydantic import BaseModel, Field

PRESETS = {
    'local': {'label':'本地模型 / Ollama / LM Studio', 'base_url':'http://127.0.0.1:11434/v1', 'model':'', 'send_image':True},
    'openai': {'label':'OpenAI API', 'base_url':'https://api.openai.com/v1', 'model':'', 'send_image':True},
    'deepseek': {'label':'DeepSeek API', 'base_url':'https://api.deepseek.com', 'model':'deepseek-flash', 'send_image':True},
    'custom': {'label':'其他兼容 API', 'base_url':'', 'model':'', 'send_image':True},
}


class AIProfile(BaseModel):
    provider: str
    base_url: str = Field(max_length=500)
    model: str = Field(min_length=1, max_length=150)
    api_key: str | None = Field(default=None, max_length=4096, repr=False)
    clear_key: bool = False
    remember_key: bool = False
    send_image: bool = True
    use_context: bool = True
    instruction: str = Field(default='用简洁中文描述图片的主要用途或内容，避免空泛名称。', max_length=1000)
    prefix: str = Field(default='', max_length=30)
    numbered: bool = True
    timeout: int = Field(default=90, ge=10, le=180)


def crypt_secret(value, decrypt=False):
    """Windows DPAPI protects remembered keys for the current Windows account."""
    if sys.platform != 'win32':
        raise ValueError('当前系统仅支持在本次运行中保留密钥，请关闭“记住密钥”')
    import ctypes
    from ctypes import wintypes as wt
    class Blob(ctypes.Structure):
        _fields_ = [('size', wt.DWORD), ('data', ctypes.c_void_p)]
    raw = base64.b64decode(value) if decrypt else value.encode('utf-8')
    buffer = ctypes.create_string_buffer(raw)
    source, target = Blob(len(raw), ctypes.cast(buffer, ctypes.c_void_p)), Blob()
    library = ctypes.WinDLL('crypt32', use_last_error=True)
    function = library.CryptUnprotectData if decrypt else library.CryptProtectData
    function.argtypes = [ctypes.POINTER(Blob), ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
                         ctypes.c_void_p, wt.DWORD, ctypes.POINTER(Blob)]
    function.restype = wt.BOOL
    if not function(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(target)):
        raise ValueError('无法读取或保存加密密钥，请重新输入 API Key')
    kernel = ctypes.WinDLL('kernel32')
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    try:
        result = ctypes.string_at(target.data, target.size)
        return result.decode('utf-8') if decrypt else base64.b64encode(result).decode('ascii')
    finally:
        kernel.LocalFree(target.data)


class AISettings:
    def __init__(self, path=None):
        if path is None:
            root = Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Shitu' if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent / '.local'
            path = root / 'ai-settings.json'
        self.path = Path(path)
        self.lock = threading.RLock()
        self.active = 'local'
        self.keys = {}
        self.encrypted = {}
        self.profiles = {key: {**AIProfile(provider=key, base_url=value['base_url'], model=value['model'] or '未配置').model_dump(exclude={'api_key','clear_key'}), 'model':value['model']} for key, value in PRESETS.items()}
        if self.path.is_file():
            try:
                saved = json.loads(self.path.read_text(encoding='utf-8'))
                if saved.get('active') in PRESETS:
                    self.active = saved['active']
                for key, profile in saved.get('profiles', {}).items():
                    if key not in PRESETS:
                        continue
                    secret = profile.pop('encrypted_key', '')
                    parsed = AIProfile(**{**profile, 'provider':key, 'api_key':None}).model_dump(exclude={'api_key','clear_key'})
                    self.profiles[key] = parsed
                    if secret:
                        try:
                            self.keys[key] = crypt_secret(secret, decrypt=True)
                            self.encrypted[key] = secret
                        except ValueError:
                            pass
            except (OSError, ValueError, TypeError):
                pass

    def public(self):
        with self.lock:
            return {'active':self.active, 'can_remember_key':sys.platform == 'win32', 'profiles':{
                key:{**profile, 'label':PRESETS[key]['label'], 'has_api_key':bool(self.keys.get(key))}
                for key, profile in self.profiles.items()}}

    def save(self, profile):
        if profile.provider not in PRESETS:
            raise ValueError('请选择有效的模型配置')
        parts = urlsplit(profile.base_url.strip().rstrip('/'))
        if parts.scheme not in ('http','https') or not parts.hostname or parts.username or parts.password or parts.query or parts.fragment:
            raise ValueError('接口地址需为 HTTP/HTTPS 基础地址，不能包含账号、密码、查询参数或片段')
        if parts.path.rstrip('/').endswith(('/chat/completions','/responses')):
            raise ValueError('请填写基础地址（如 https://api.openai.com/v1），不要填写完整的聊天接口路径')
        if not profile.send_image and not profile.use_context:
            raise ValueError('请至少启用图片识别或文档描述')
        if profile.remember_key and sys.platform != 'win32':
            raise ValueError('当前系统不支持加密保存密钥，请关闭“记住密钥”')
        with self.lock:
            key = profile.provider
            values = profile.model_dump(exclude={'api_key','clear_key'})
            values.update(base_url=profile.base_url.strip().rstrip('/'), model=profile.model.strip())
            if not values['model']:
                raise ValueError('请填写模型名称')
            changed_endpoint = self.profiles[key]['base_url'] != values['base_url']
            secret = '' if profile.clear_key or changed_endpoint else self.keys.get(key, '')
            if profile.api_key is not None:
                secret = profile.api_key.strip()
            encrypted = crypt_secret(secret) if secret and profile.remember_key else ''
            profiles = copy.deepcopy(self.profiles)
            profiles[key] = values
            remembered = dict(self.encrypted)
            remembered[key] = encrypted
            output = {'active':key, 'profiles':{k:{**v, **({'encrypted_key':remembered[k]} if remembered.get(k) else {})} for k,v in profiles.items() if v['model']}}
            self.path.parent.mkdir(parents=True, exist_ok=True)
            temporary = self.path.with_name(self.path.name+'.'+uuid.uuid4().hex+'.tmp')
            try:
                temporary.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding='utf-8')
                temporary.replace(self.path)
            finally:
                temporary.unlink(missing_ok=True)
            self.profiles, self.encrypted = profiles, remembered
            self.keys[key], self.active = secret, key
            return self.public()

    def snapshot(self, provider):
        with self.lock:
            if provider not in PRESETS:
                raise ValueError('请选择模型配置')
            profile = copy.deepcopy(self.profiles[provider])
            if not profile['base_url'] or not profile['model']:
                raise ValueError('请先配置并保存接口地址和模型名称')
            secret = self.keys.get(provider, '')
            if provider in ('openai','deepseek') and not secret:
                raise ValueError('请先为此配置填写 API Key；聊天订阅不能代替 API Key')
            return profile, secret
