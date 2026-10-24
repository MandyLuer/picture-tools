import io
import zipfile

import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.document_context import filename_from_text
from backend.service import Workspace, safe_name


def document(body):
    image = io.BytesIO()
    Image.new('RGB', (30, 20), 'green').save(image, 'PNG')
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as archive:
        archive.writestr('word/document.xml', f'''<w:document
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>{body}</w:body></w:document>''')
        archive.writestr('word/_rels/document.xml.rels', '<Relationships><Relationship Id="a" Target="media/a.png"/><Relationship Id="b" Target="media/b.png"/></Relationships>')
        archive.writestr('word/media/a.png', image.getvalue())
        archive.writestr('word/media/b.png', image.getvalue())
    return output.getvalue()


PIC = '<w:p><w:r><w:drawing><a:blip r:embed="a"/></w:drawing></w:r></w:p>'


def paragraph(text):
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


@pytest.fixture
def workspace():
    work = Workspace()
    yield work
    work.close()


def test_nearby_text_caption_priority_and_preserve_on_edit(workspace):
    workspace.extract(document(paragraph('这里介绍部署方案') + PIC + paragraph('图 3-1 系统部署架构') + paragraph('该图展示服务之间的关系')), '架构.docx')
    item = workspace.listing()[0]
    context = item['contexts'][0]
    assert context['document'] == '架构.docx'
    assert context['paragraph'] == 2
    assert context['above'] == ['这里介绍部署方案']
    assert context['below'] == ['图 3-1 系统部署架构', '该图展示服务之间的关系']
    assert context['candidates'][0]['name'] == '图 3-1 系统部署架构'
    assert context['candidates'][0]['is_caption']
    workspace.rename(item['id'], context['candidates'][0]['name'])
    edited = workspace.edit(item['id'], 'crop', [0, 0, 10, 10])
    assert workspace.get(edited)['contexts'] == item['contexts']
    assert edited == item['id']
    assert workspace.get(edited)['name'] == '图 3-1 系统部署架构.png'
    assert len(workspace.images) == 1


def test_inline_text_and_repeated_image_occurrences(workspace):
    inline = '<w:p><w:r><w:t>登录页面：</w:t><w:drawing><a:blip r:embed="a"/></w:drawing><w:t>图 1 用户登录</w:t></w:r></w:p>'
    workspace.extract(document(inline + paragraph('分隔文字') + PIC + paragraph('图 2 管理员入口')))
    assert len(workspace.images) == 1
    contexts = workspace.listing()[0]['contexts']
    assert len(contexts) == 2
    assert contexts[0]['inline_before'] == '登录页面：'
    assert contexts[0]['inline_after'] == '图 1 用户登录'
    assert contexts[1]['below'] == ['图 2 管理员入口']


def test_table_cells_and_other_images_are_boundaries(workspace):
    body = '<w:tbl><w:tr><w:tc>' + paragraph('本单元格标题') + PIC + '</w:tc><w:tc>' + paragraph('其他单元格不得关联') + '</w:tc></w:tr></w:tbl>'
    body += paragraph('正文描述不得关联表格图片') + PIC.replace('embed="a"', 'embed="b"') + PIC + paragraph('图 5 第二张图')
    workspace.extract(document(body))
    first, second = workspace.listing()
    assert first['contexts'][0]['above'] == ['本单元格标题']
    assert first['contexts'][0]['below'] == []
    assert second['contexts'][0]['below'] == []


@pytest.mark.parametrize('text', ['部署:示意/图片?', 'CON', '测试' * 100, '  name.  '])
def test_suggested_names_are_valid(text):
    name = filename_from_text(text)
    assert safe_name(name) == name
    assert len(name) <= 100


def test_import_api_returns_document_context():
    app = create_app()
    with TestClient(app) as client:
        response = client.post('/api/import', files={'files': ('报告.docx', document(PIC + paragraph('图 8 设备配置')))})
        assert response.status_code == 200
        item = response.json()[0]
        assert item['contexts'][0]['document'] == '报告.docx'
        assert item['contexts'][0]['candidates'][0]['name'] == '图 8 设备配置'


def test_alternative_text_belongs_to_its_own_drawing(workspace):
    body = '''<w:p><w:r>
      <w:drawing><a:docPr descr="第一张图描述"/><a:blip r:embed="a"/></w:drawing>
      <w:drawing><a:docPr descr="第二张图描述"/><a:blip r:embed="b"/></w:drawing>
    </w:r></w:p>'''
    workspace.extract(document(body))
    first, second = workspace.listing()
    assert first['contexts'][0]['alt_text'] == ['第一张图描述']
    assert second['contexts'][0]['alt_text'] == ['第二张图描述']
