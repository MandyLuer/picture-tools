import io
import json
from pathlib import Path
import sys
import threading
import time

import httpx
import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.ai_settings import AIProfile, AISettings
from backend.ai_naming import call_model


@pytest.fixture
def client(tmp_path):
    with TestClient(create_app(ai_settings_path=tmp_path/'ai.json')) as client:
        yield client


def configure(client, **changes):
    values={'provider':'local','base_url':'http://127.0.0.1:11434/v1','model':'vision-test','numbered':False,**changes}
    result=client.put('/api/ai/settings',json=values)
    assert result.status_code==200,result.text
    return result.json()


def images(client):
    raw=io.BytesIO();Image.new('RGB',(30,20),'red').save(raw,'PNG')
    return client.post('/api/import',files=[('files',('one.png',raw.getvalue())),('files',('two.png',raw.getvalue()))]).json()


def wait(client,key):
    deadline=time.monotonic()+3
    while time.monotonic()<deadline:
        result=client.get('/api/ai/jobs/'+key).json()
        if result['status'] not in ('running','cancelling'):return result
        time.sleep(.01)
    pytest.fail('AI job timed out')


def generate(client,items,**extra):
    result=client.post('/api/ai/jobs',json={'provider':'local','images':[item['id'] for item in items],**extra})
    assert result.status_code==200,result.text
    return wait(client,result.json()['id'])


def test_profile_switching_and_secret_redaction(tmp_path):
    path=tmp_path/'ai.json';store=AISettings(path)
    store.save(AIProfile(provider='openai',base_url='https://api.openai.com/v1',model='vision-model',api_key='secret-example'))
    store.save(AIProfile(provider='local',base_url='http://localhost:11434/v1',model='local-model'))
    assert store.public()['active']=='local'
    assert store.public()['profiles']['openai']['has_api_key']
    assert 'secret-example' not in json.dumps(store.public())
    assert 'secret-example' not in path.read_text(encoding='utf-8')
    store.save(AIProfile(provider='openai',base_url='https://another.example/v1',model='vision-model'))
    assert not store.public()['profiles']['openai']['has_api_key']


@pytest.mark.skipif(sys.platform!='win32',reason='Windows DPAPI')
def test_remembered_key_is_encrypted_and_restored(tmp_path):
    path=tmp_path/'settings.json';store=AISettings(path)
    store.save(AIProfile(provider='deepseek',base_url='https://api.deepseek.com',model='deepseek-flash',api_key='secret-dpapi-example',remember_key=True))
    assert 'secret-dpapi-example' not in path.read_text(encoding='utf-8')
    restored=AISettings(path)
    assert restored.snapshot('deepseek')[1]=='secret-dpapi-example'
    restored.save(AIProfile(provider='deepseek',base_url='https://api.deepseek.com',model='deepseek-flash',clear_key=True))
    assert not AISettings(path).public()['profiles']['deepseek']['has_api_key']


@pytest.mark.parametrize('base_url',['file:///etc/passwd','https://user:pass@example.com','https://example.com?key=secret','https://example.com/v1/chat/completions'])
def test_invalid_urls_are_rejected(client,base_url):
    assert client.put('/api/ai/settings',json={'provider':'custom','base_url':base_url,'model':'test'}).status_code==400


def test_suggestions_are_sanitized_unique_and_not_applied_until_confirmation(client,monkeypatch):
    configure(client)
    original=images(client)
    monkeypatch.setattr('backend.ai_naming.call_model',lambda *args:{'name':'网络/结构.png','reason':'根据图中连接关系','needs_review':False})
    job=generate(client,original)
    assert all(row['status']=='success' for row in job['results'])
    assert [row['name'] for row in job['results']]==['网络 结构','网络 结构_2']
    assert client.get('/api/images').json()==original
    result=client.post('/api/ai/apply',json={'job_id':job['id'],'changes':[{'id':row['id'],'name':row['name']} for row in job['results']]})
    assert result.status_code==200
    assert [item['name'] for item in result.json()['images']]==['网络 结构.png','网络 结构_2.png']
    assert all(item['revision']==0 for item in result.json()['images'])


def test_numbering_uses_original_workspace_position(client,monkeypatch):
    configure(client,numbered=True)
    original=images(client)
    monkeypatch.setattr('backend.ai_naming.call_model',lambda *args:{'name':'设备','reason':'测试','needs_review':False})
    assert generate(client,original[1:])['results'][0]['name']=='002_设备'


def test_apply_is_atomic_on_conflict_and_rejects_stale_edits(client,monkeypatch):
    configure(client);original=images(client)
    monkeypatch.setattr('backend.ai_naming.call_model',lambda *args:{'name':'设备','reason':'测试','needs_review':False})
    job=generate(client,original)
    conflict=client.post('/api/ai/apply',json={'job_id':job['id'],'changes':[{'id':item['id'],'name':'duplicate'} for item in original]})
    assert conflict.status_code==400 and client.get('/api/images').json()==original
    client.post('/api/images/'+original[0]['id']+'/edit',json={'operation':'right'})
    stale=client.post('/api/ai/apply',json={'job_id':job['id'],'changes':[{'id':original[0]['id'],'name':'新名称'}]})
    assert stale.status_code==400 and '已变化' in stale.json()['detail']


def test_cancel_preserves_successful_suggestions_and_blocks_second_job(client,monkeypatch):
    configure(client);original=images(client)
    entered,release=threading.Event(),threading.Event()
    def model(*args):
        entered.set();assert release.wait(3)
        return {'name':'设备','reason':'测试','needs_review':True}
    monkeypatch.setattr('backend.ai_naming.call_model',model)
    job=client.post('/api/ai/jobs',json={'provider':'local','images':[item['id'] for item in original]}).json()
    try:
        assert entered.wait(2)
        assert client.post('/api/ai/jobs',json={'provider':'local','images':[original[0]['id']]}).status_code==409
        assert client.post('/api/ai/jobs/'+job['id']+'/cancel').status_code==200
    finally:release.set()
    result=wait(client,job['id'])
    assert result['status']=='cancelled'
    assert [row['status'] for row in result['results']]==['success','cancelled']
    assert result['results'][0]['needs_review']
    assert client.get('/api/images').json()==original


def test_text_only_requires_context_and_does_not_send_image(client,monkeypatch):
    configure(client,send_image=False)
    original=images(client)
    called=[]
    monkeypatch.setattr('backend.ai_naming.call_model',lambda *args:called.append(args))
    job=generate(client,original)
    assert all(row['status']=='failed' and '没有文档描述' in row['error'] for row in job['results'])
    assert not called


def test_text_only_uses_word_descriptions_without_image_payload(client,monkeypatch):
    configure(client,send_image=False)
    original=images(client)
    client.app.state.workspace.get(original[0]['id'])['contexts']=[{'above':['系统部署说明'],'below':['图 1 网络拓扑图']}]
    seen=[]
    def model(profile,secret,text,picture):
        seen.append((text,picture))
        return {'name':'网络拓扑','reason':'来自题注','needs_review':False}
    monkeypatch.setattr('backend.ai_naming.call_model',model)
    result=generate(client,original[:1])
    assert result['results'][0]['status']=='success'
    assert '图 1 网络拓扑图' in seen[0][0] and seen[0][1] is None


def test_clear_workspace_invalidates_old_suggestions(client,monkeypatch):
    configure(client);original=images(client)
    monkeypatch.setattr('backend.ai_naming.call_model',lambda *args:{'name':'设备','reason':'测试','needs_review':False})
    job=generate(client,original)
    client.delete('/api/images')
    result=client.post('/api/ai/apply',json={'job_id':job['id'],'changes':[{'id':original[0]['id'],'name':'名称'}]})
    assert result.status_code==400 and '工作区已变化' in result.json()['detail']


def test_http_protocol_vision_payload_and_json_parsing(monkeypatch):
    received=[];original_client=httpx.Client
    def transport(request):
        received.append(request)
        return httpx.Response(200,json={'choices':[{'message':{'content':'```json\n{"name":"结构图","reason":"图片","needs_review":false}\n```'}}]})
    monkeypatch.setattr('backend.ai_naming.httpx.Client',lambda **kwargs:original_client(transport=httpx.MockTransport(transport),**kwargs))
    profile=AIProfile(provider='openai',base_url='https://api.openai.com/v1',model='vision-test').model_dump()
    result=call_model(profile,'test-key','测试描述','data:image/jpeg;base64,AAAA')
    assert result['name']=='结构图'
    payload=json.loads(received[0].content)
    assert received[0].url.path=='/v1/chat/completions'
    assert payload['messages'][1]['content'][1]['type']=='image_url'
    assert received[0].headers['authorization']=='Bearer test-key'


def test_http_auth_error_does_not_expose_key(monkeypatch):
    original_client=httpx.Client
    monkeypatch.setattr('backend.ai_naming.httpx.Client',lambda **kwargs:original_client(transport=httpx.MockTransport(lambda request:httpx.Response(401,text='secret-key')),**kwargs))
    profile=AIProfile(provider='local',base_url='http://localhost/v1',model='test').model_dump()
    with pytest.raises(ValueError,match='API Key 无效') as error:
        call_model(profile,'secret-key','hello')
    assert 'secret-key' not in str(error.value)


def test_api_auth_protects_settings_and_job_access(tmp_path):
    with TestClient(create_app('token',ai_settings_path=tmp_path/'ai.json')) as client:
        assert client.get('/api/ai/settings').status_code==403
        assert client.get('/api/ai/settings',headers={'x-app-token':'token'}).status_code==200
