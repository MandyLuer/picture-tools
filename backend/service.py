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
from collections import OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageOps
from .document_context import image_contexts

MAX_FILE = 50 * 1024 * 1024
MAX_TOTAL = 250 * 1024 * 1024
MAX_PREVIEW_CACHE = 64 * 1024 * 1024
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
        self._preview_cache = OrderedDict()
        self._preview_bytes = 0

    def close(self):
        self._preview_cache.clear()
        self._preview_bytes = 0
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
        self._preview_cache.clear()
        self._preview_bytes = 0
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
        self.invalidate_preview(key)

    def invalidate_preview(self, key):
        for cache_key in list(self._preview_cache):
            if cache_key[0] == key:
                self._preview_bytes -= len(self._preview_cache.pop(cache_key))

    def cached_preview(self, key, revision, thumbnail):
        cache_key = (key, revision, thumbnail)
        data = self._preview_cache.get(cache_key)
        if data is not None:
            self._preview_cache.move_to_end(cache_key)
        return data

    def cache_preview(self, key, revision, thumbnail, data):
        if key not in self.images or self.images[key]['revision'] != revision or len(data) > MAX_PREVIEW_CACHE:
            return
        cache_key = (key, revision, thumbnail)
        previous = self._preview_cache.pop(cache_key, b'')
        self._preview_bytes -= len(previous)
        while self._preview_cache and self._preview_bytes + len(data) > MAX_PREVIEW_CACHE:
            _, retired = self._preview_cache.popitem(last=False)
            self._preview_bytes -= len(retired)
        self._preview_cache[cache_key] = data
        self._preview_bytes += len(data)

    @staticmethod
    def render_preview(data, image_format, animated, thumbnail):
        # Browser-native static PNG previews need no decoding or re-encoding.
        if not thumbnail and image_format == 'PNG' and not animated:
            return data
        with Image.open(io.BytesIO(data)) as source:
            if thumbnail:
                source.thumbnail((96, 96))
            rendered = source.convert('RGBA')
            output = io.BytesIO()
            rendered.save(output, 'PNG')
            return output.getvalue()

    def add(self, data, name, contexts=None):
        if len(data) > MAX_FILE:
            raise ValueError("单个文件不能超过 50 MB")
        with Image.open(io.BytesIO(data)) as image:
            width, height = image.size
            animated = getattr(image, 'is_animated', False)
            image.verify()
            ext = EXTENSIONS.get(image.format)
            image_format = image.format
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
        self.images[key] = {"id": key, "name": filename, "path": path, "contexts": deepcopy(contexts or []), "revision": 0,
                            'width': width, 'height': height, 'size': len(data), 'format': image_format, 'animated': animated}
        return key

    def get(self, key):
        if key not in self.images:
            raise KeyError("图片不存在")
        return self.images[key]

    def listing(self):
        fields = ('id', 'name', 'width', 'height', 'size', 'contexts', 'revision')
        return [{**{field: item[field] for field in fields}, **({'has_source': True} if item.get('source') else {})} for item in self.images.values()]

    def extract(self, data, document_name="Word 文档"):
        if len(data) > MAX_FILE:
            raise ValueError("Word 文件不能超过 50 MB")
        if Path(document_name).suffix.lower() == '.doc':
            from .document_conversion import convert_doc
            data = convert_doc(data)
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
        item.update(width=image.width, height=image.height, size=len(data), format=fmt, animated=False)
        self.invalidate_preview(key)
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
