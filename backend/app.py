import io
import sys
import threading
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote
import mimetypes

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

    @asynccontextmanager
    async def lifespan(app):
        yield
        workspace.close()

    app = FastAPI(title="拾图 Shitu · Word 图片工作台", version="2.3.0", lifespan=lifespan)
    app.state.workspace = workspace
    app.state.lock = lock

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
