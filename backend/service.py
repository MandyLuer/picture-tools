"""Image workspace independent of HTTP and desktop UI."""
import io
import posixpath
import re
import shutil
import tempfile
import uuid
import zipfile
import logging
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageOps
from .document_context import image_contexts

MAX_FILE = 50 * 1024 * 1024
MAX_TOTAL = 250 * 1024 * 1024
EXTENSIONS = {"JPEG": ".jpg", "PNG": ".png", "GIF": ".gif", "BMP": ".bmp", "TIFF": ".tiff", "WEBP": ".webp"}


def safe_name(name):
    name = name.strip()
    if not name or len(name) > 180 or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")):
        raise ValueError("文件名无效")
    if name.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *[f"COM{i}" for i in range(1, 10)], *[f"LPT{i}" for i in range(1, 10)]}:
        raise ValueError("不能使用系统保留文件名")
    return name


class Workspace:
    def __init__(self):
        self.temp = tempfile.TemporaryDirectory(prefix="picture-tools-")
        self.root = Path(self.temp.name)
        self.images = {}
        self._retired = []

    def close(self):
        for temp in [self.temp, *self._retired]:
            try:
                temp.cleanup()
            except OSError:
                logging.getLogger(__name__).warning("Temporary workspace cleanup deferred")

    def clear(self):
        # Switch generations first so a locked old file cannot leave a partially
        # deleted active workspace. Retired storage is retried when closing.
        fresh = tempfile.TemporaryDirectory(prefix="picture-tools-")
        retired = self.temp
        self.temp, self.root, self.images = fresh, Path(fresh.name), {}
        self._retired.append(retired)
        for temp in self._retired[:]:
            try:
                temp.cleanup()
            except OSError:
                logging.getLogger(__name__).warning("Old workspace cleanup deferred until close")
            else:
                self._retired.remove(temp)

    def remove(self, key):
        item = self.get(key)
        item['path'].unlink(missing_ok=True)
        del self.images[key]

    def add(self, data, name, contexts=None):
        if len(data) > MAX_FILE:
            raise ValueError("单个文件不能超过 50 MB")
        with Image.open(io.BytesIO(data)) as image:
            image.verify()
            ext = EXTENSIONS.get(image.format)
        if not ext:
            raise ValueError("不支持的图片格式")
        if sum(p.stat().st_size for p in self.root.iterdir()) + len(data) > MAX_TOTAL:
            raise ValueError("工作区不能超过 250 MB，请清空后重试")
        stem = safe_name(Path(name).stem)
        filename = stem + ext
        count = 1
        while any(item["name"].casefold() == filename.casefold() for item in self.images.values()):
            filename = f"{stem}_{count}{ext}"
            count += 1
        key = uuid.uuid4().hex
        path = self.root / (key + ext)
        path.write_bytes(data)
        self.images[key] = {"id": key, "name": filename, "path": path, "contexts": deepcopy(contexts or []), "revision": 0}
        return key

    def get(self, key):
        if key not in self.images:
            raise KeyError("图片不存在")
        return self.images[key]

    def listing(self):
        result = []
        for item in self.images.values():
            with Image.open(item["path"]) as image:
                result.append({"id": item["id"], "name": item["name"], "width": image.width, "height": image.height, "size": item["path"].stat().st_size, "contexts": item["contexts"], "revision": item["revision"]})
        return result

    def extract(self, data, document_name="Word 文档"):
        if len(data) > MAX_FILE:
            raise ValueError("Word 文件不能超过 50 MB")
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if sum(i.file_size for i in archive.infolist()) > MAX_TOTAL:
                raise ValueError("Word 解压后过大")
            rels = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
            mapping = {r.get("Id"): posixpath.normpath("word/" + r.get("Target", "")) for r in rels if r.get("TargetMode") != "External"}
            root = ET.fromstring(archive.read("word/document.xml"))
            contexts = image_contexts(root, mapping, document_name)
            ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            paths = []
            for node in root.iter():
                ref = node.get(ns + "embed") or node.get(ns + "id")
                target = mapping.get(ref, "")
                if target.startswith("word/media/") and target not in paths:
                    paths.append(target)
            prepared = []
            for path in paths:
                raw = archive.read(path)
                try:
                    with Image.open(io.BytesIO(raw)) as image:
                        image.verify()
                        if image.format not in EXTENSIONS:
                            continue
                except (OSError, ValueError):
                    continue
                prepared.append((raw, f"{len(prepared)+1:03d}", contexts.get(path, [])))
            for raw, name, context in prepared:
                self.add(raw, name, context)
        return len(prepared)

    def rename(self, key, name):
        item = self.get(key)
        filename = safe_name(name) + Path(item["name"]).suffix
        if any(i["id"] != key and i["name"].casefold() == filename.casefold() for i in self.images.values()):
            raise ValueError("文件名已存在")
        item["name"] = filename

    def edit(self, key, operation, box=None, quality=85):
        item = self.get(key)
        with Image.open(item["path"]) as source:
            image = source.copy()
            fmt = source.format
        if operation == "left":
            image = image.transpose(Image.Transpose.ROTATE_90)
        elif operation == "right":
            image = image.transpose(Image.Transpose.ROTATE_270)
        elif operation == "horizontal":
            image = ImageOps.mirror(image)
        elif operation == "vertical":
            image = ImageOps.flip(image)
        elif operation == "crop":
            if not box or len(box) != 4 or not (0 <= box[0] < box[2] <= image.width and 0 <= box[1] < box[3] <= image.height):
                raise ValueError("裁剪区域超出图片范围")
            image = image.crop(tuple(box))
        elif operation != "compress":
            raise ValueError("不支持的操作")
        output = io.BytesIO()
        if fmt == "JPEG":
            image = image.convert("RGB")
        image.save(output, format=fmt, quality=quality if operation == 'compress' else 95, optimize=True)
        data = output.getvalue()
        if len(data) > MAX_FILE:
            raise ValueError("编辑后的图片超过 50 MB")
        if sum(p.stat().st_size for p in self.root.iterdir()) - item['path'].stat().st_size + len(data) > MAX_TOTAL:
            raise ValueError("编辑后的工作区超过 250 MB")
        pending = item['path'].with_suffix('.pending')
        try:
            pending.write_bytes(data)
            pending.replace(item['path'])
        finally:
            pending.unlink(missing_ok=True)
        item['revision'] += 1
        return key

    def export_zip(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for item in self.images.values():
                archive.write(item["path"], item["name"])
        return output.getvalue()

    def export_folder(self, folder):
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        if any((folder / i["name"]).exists() for i in self.images.values()):
            raise ValueError("目标文件夹存在同名文件，请选择空文件夹")
        for item in self.images.values():
            shutil.copyfile(item["path"], folder / item["name"])
