import io
import sys
import threading
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote
import mimetypes
import logging

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import Response, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from PIL import Image

from .service import MAX_FILE, Workspace


class Rename(BaseModel):
    name: str


class Edit(BaseModel):
    operation: str
    box: list[int] | None = None
    quality: int = Field(default=85, ge=1, le=100)


def create_app(token=None):
    workspace = Workspace()
    lock = threading.RLock()
    export_picker_lock = threading.Lock()

    @asynccontextmanager
    async def lifespan(app):
        yield
        workspace.close()

    app = FastAPI(title="拾图 Shitu · Word 图片工作台", version="2.6.0", lifespan=lifespan)
    app.state.workspace = workspace
    app.state.lock = lock
    from .certificate_batch import register_certificate_routes
    register_certificate_routes(app)

    @app.middleware("http")
    async def protect(request: Request, call_next):
        if token and request.url.path.startswith("/api/") and request.headers.get("x-app-token") != token and request.query_params.get("token") != token:
            return JSONResponse({"detail": "访问未授权"}, status_code=403)
        if request.method not in ("GET", "HEAD"):
            origin = request.headers.get("origin")
            if origin and origin != str(request.base_url).rstrip("/") and origin not in ("http://localhost:5173", "http://127.0.0.1:5173"):
                return JSONResponse({"detail": "来源未授权"}, status_code=403)
        return await call_next(request)

    @app.exception_handler(ValueError)
    async def invalid(request, exc):
        return JSONResponse({"detail": str(exc)}, status_code=400)

    @app.exception_handler(KeyError)
    async def missing(request, exc):
        return JSONResponse({"detail": "图片或文档内容不存在"}, status_code=404)

    @app.exception_handler(OSError)
    async def file_error(request, exc):
        return JSONResponse({"detail": "文件暂时无法写入或删除，请稍后重试"}, status_code=409)

    @app.exception_handler(Exception)
    async def unexpected_error(request, exc):
        logging.getLogger('shitu.api').error('Unhandled error for %s', request.url.path,
                                            exc_info=(type(exc), exc, exc.__traceback__))
        return JSONResponse({'detail': '服务端处理失败，请稍后重试；具体原因可查看后端日志'}, status_code=500)

    @app.get("/api/images")
    def listing():
        with lock:
            return workspace.listing()

    @app.post("/api/import")
    def upload(files: list[UploadFile] = File(...)):
        with lock:
            before = set(workspace.images)
            try:
                for file in files:
                    data = file.file.read(MAX_FILE + 1)
                    if (file.filename or "").lower().endswith(".docx"):
                        workspace.extract(data, file.filename or "Word 文档")
                    else:
                        workspace.add(data, file.filename or "image")
            except Exception as exc:
                for key in set(workspace.images) - before:
                    workspace.images.pop(key)["path"].unlink(missing_ok=True)
                raise HTTPException(400, "导入失败，请检查文件格式、大小和文档内容") from exc
            return workspace.listing()

    @app.get("/api/images/{key}/file")
    def image(key: str, preview: bool = False, thumbnail: bool = False):
        with lock:
            item = workspace.get(key)
            if preview:
                output = io.BytesIO()
                with Image.open(item["path"]) as source:
                    rendered = source.convert("RGBA")
                    if thumbnail:
                        rendered.thumbnail((96, 96))
                    rendered.save(output, "PNG")
                output.seek(0)
                return StreamingResponse(output, media_type="image/png", headers={"Cache-Control": "no-store"})
            # Snapshot bytes under the lock. Lazy file streaming could otherwise
            # race with an in-place edit, deletion or workspace clear on Windows.
            return Response(item['path'].read_bytes(), media_type=mimetypes.guess_type(item['name'])[0] or 'application/octet-stream',
                            headers={"Cache-Control": "no-store", "Content-Disposition": "attachment; filename*=UTF-8''" + quote(item['name'])})

    @app.patch("/api/images/{key}")
    def rename(key: str, body: Rename):
        with lock:
            workspace.rename(key, body.name)
            return workspace.listing()

    @app.post("/api/images/{key}/edit")
    def edit(key: str, body: Edit):
        with lock:
            new_id = workspace.edit(key, body.operation, body.box, body.quality)
            return {"images": workspace.listing(), "selected": new_id}

    @app.get("/api/export")
    def export():
        with lock:
            return StreamingResponse(io.BytesIO(workspace.export_zip()), media_type="application/zip", headers={"Content-Disposition": 'attachment; filename="images.zip"'})

    @app.post('/api/export/folder')
    def export_to_system_folder(request: Request):
        if not request.client or request.client.host not in ('127.0.0.1', '::1'):
            raise HTTPException(403, '系统文件夹导出仅适用于本机访问，请使用 ZIP 压缩包')
        if not export_picker_lock.acquire(blocking=False):
            raise HTTPException(409, '已有文件夹导出正在进行，请先完成或取消系统选择框')
        try:
            with lock:
                if not workspace.images:
                    raise ValueError('工作区没有可导出的图片')
                generation = workspace.root
            from .folder_picker import choose_export_folder
            target = choose_export_folder()
            if not target:
                return {'cancelled': True}
            with lock:
                if workspace.root != generation:
                    raise ValueError('工作区已变化，请重新导出')
                workspace.export_folder(target)
                return {'count': len(workspace.images), 'folder': Path(target).name}
        finally:
            export_picker_lock.release()

    @app.delete("/api/images")
    def clear():
        with lock:
            workspace.clear()
            return []

    @app.delete("/api/images/{key}")
    def delete_image(key: str):
        with lock:
            workspace.remove(key)
            return workspace.listing()

    dist = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)) / "frontend" / "dist"
    if dist.exists():
        app.mount("/", StaticFiles(directory=dist, html=True), name="frontend")
    return app
