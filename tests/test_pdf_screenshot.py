import io

import pymupdf
import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.pdf_screenshot import parse_page_spec, render_pdf_pages


def make_pdf(page_count=2, size=(72, 72), color=(1, 0, 0), encrypt=False):
    document = pymupdf.open()
    for _ in range(page_count):
        page = document.new_page(width=size[0], height=size[1])
        page.draw_rect(page.rect, color=color, fill=color)
    options = {}
    if encrypt:
        options = {'encryption': pymupdf.PDF_ENCRYPT_AES_256, 'owner_pw': 'owner', 'user_pw': 'secret'}
    data = document.tobytes(**options)
    document.close()
    return data


def test_parse_page_spec_selects_sorts_and_deduplicates():
    assert parse_page_spec('3,1-2,2', 5) == [1, 2, 3]
    assert parse_page_spec(' 1 - 2 ，4 ', 5) == [1, 2, 4]
    assert parse_page_spec('', 3) == [1, 2, 3]


@pytest.mark.parametrize('spec', ['0', '4-2', 'abc', '6', '1-'])
def test_parse_page_spec_rejects_invalid_input(spec):
    with pytest.raises(ValueError):
        parse_page_spec(spec, 5)


def test_parse_page_spec_rejects_ranges_beyond_page_limit():
    with pytest.raises(ValueError, match='200'):
        parse_page_spec('1-300', 1000)


def test_render_pdf_pages_uses_150_dpi_and_keeps_page_order():
    rendered = list(render_pdf_pages(make_pdf(3), '报告.pdf', '1-2', 150))
    assert [name for name, _, _ in rendered] == ['报告_p001.png', '报告_p002.png']
    first_name, first_context, first_data = rendered[0]
    assert first_context['document'] == '报告.pdf'
    assert first_context['page'] == 1
    with Image.open(io.BytesIO(first_data)) as image:
        assert image.size == (150, 150)
        assert image.getpixel((75, 75)) == (255, 0, 0)


def test_render_pdf_pages_honours_300_dpi():
    (_, _, data), = render_pdf_pages(make_pdf(1), 'a.pdf', '', 300)
    with Image.open(io.BytesIO(data)) as image:
        assert image.size == (300, 300)


def test_render_pdf_pages_rejects_encrypted_document():
    with pytest.raises(ValueError, match='密码'):
        list(render_pdf_pages(make_pdf(1, encrypt=True), 'secret.pdf', '', 150))


def test_render_pdf_pages_rejects_invalid_document():
    with pytest.raises(ValueError, match='损坏|有效'):
        list(render_pdf_pages(b'not a pdf', 'broken.pdf', '', 150))


def test_render_pdf_pages_rejects_unsupported_dpi():
    with pytest.raises(ValueError, match='150|300'):
        list(render_pdf_pages(make_pdf(1), 'a.pdf', '', 200))


def test_render_pdf_pages_rejects_more_than_200_pages():
    with pytest.raises(ValueError, match='200'):
        list(render_pdf_pages(make_pdf(201, size=(18, 18)), 'long.pdf', '', 150))


def test_render_pdf_pages_rejects_oversize_page():
    with pytest.raises(ValueError, match='尺寸过大'):
        list(render_pdf_pages(make_pdf(1, size=(2000, 2000)), 'poster.pdf', '', 300))


def test_pdf_screenshot_endpoint_imports_selected_pages():
    with TestClient(create_app()) as client:
        response = client.post(
            '/api/pdf/screenshot',
            files={'file': ('报告.pdf', make_pdf(3), 'application/pdf')},
            data={'pages': '2,3', 'dpi': '150'},
        )
        assert response.status_code == 200
        payload = response.json()
        assert payload['pages'] == 2
        assert [item['name'] for item in payload['images']] == ['报告_p002.png', '报告_p003.png']
        assert payload['images'][0]['contexts'][0]['document'] == '报告.pdf'
        assert payload['images'][0]['contexts'][0]['page'] == 2


def test_pdf_screenshot_endpoint_rolls_back_batch_on_failure(monkeypatch):
    with TestClient(create_app()) as client:
        workspace = client.app.state.workspace
        original = workspace.add
        calls = {'count': 0}

        def failing_add(data, name, contexts=None):
            calls['count'] += 1
            if calls['count'] == 2:
                raise ValueError('工作区不能超过 250 MB，请清空后重试')
            return original(data, name, contexts)

        monkeypatch.setattr(workspace, 'add', failing_add)
        response = client.post(
            '/api/pdf/screenshot',
            files={'file': ('report.pdf', make_pdf(3), 'application/pdf')},
            data={'pages': '', 'dpi': '150'},
        )
        assert response.status_code == 400
        assert '250 MB' in response.json()['detail']
        assert client.get('/api/images').json() == []


def test_pdf_screenshot_endpoint_reports_encrypted_document():
    with TestClient(create_app()) as client:
        response = client.post(
            '/api/pdf/screenshot',
            files={'file': ('secret.pdf', make_pdf(1, encrypt=True), 'application/pdf')},
            data={'pages': '', 'dpi': '150'},
        )
        assert response.status_code == 400
        assert '密码' in response.json()['detail']


def test_pdf_screenshot_endpoint_rejects_page_out_of_range():
    with TestClient(create_app()) as client:
        response = client.post(
            '/api/pdf/screenshot',
            files={'file': ('short.pdf', make_pdf(2), 'application/pdf')},
            data={'pages': '3', 'dpi': '150'},
        )
        assert response.status_code == 400
        assert '共 2 页' in response.json()['detail']
