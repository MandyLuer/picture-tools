"""Source files are registered only by a native folder selection, never by HTTP paths."""
import hashlib
from pathlib import Path

from PIL import Image


def digest(data):
    return hashlib.sha256(data).hexdigest()


def import_folder(workspace, folder):
    before = set(workspace.images)
    try:
        for path in sorted(Path(folder).iterdir()):
            if not path.is_file() or path.is_symlink():
                continue
            try:
                with Image.open(path) as image:
                    image.verify()
            except (OSError, ValueError):
                continue
            if path.stat().st_size > 50 * 1024 * 1024:
                raise ValueError('单张图片不能超过 50 MB')
            data = path.read_bytes()
            key = workspace.add(data, path.name)
            workspace.get(key)['source'] = {'path': path.absolute(), 'digest': digest(data)}
    except Exception:
        for key in set(workspace.images) - before:
            workspace.remove(key)
        raise
    count = len(workspace.images) - len(before)
    if not count:
        raise ValueError('所选文件夹中没有找到支持的图片')
    return count


def checked_source(item):
    source = item.get('source')
    if not source:
        raise ValueError('请通过“关联原文件夹”重新导入图片')
    path = source['path']
    if path.is_symlink() or not path.is_file() or digest(path.read_bytes()) != source['digest']:
        raise ValueError(f'原文件“{path.name}”已变化或不存在，未修改或删除，请重新导入')
    return path


def rename_source(workspace, key, name):
    item = workspace.get(key)
    source = checked_source(item)
    if item['revision']:
        raise ValueError('图片已有编辑，原位置重命名只修改名称；请导出编辑结果后再删除原图')
    old_name = item['name']
    workspace.rename(key, name)
    # Preserve the original extension (e.g. .jpeg rather than normalized .jpg).
    target = source.with_name(Path(item['name']).stem + source.suffix)
    try:
        if target.name != source.name:
            if target.exists():
                raise ValueError(f'原文件夹已存在“{target.name}”，未覆盖原文件')
            # Exclusive creation prevents overwrites, including a concurrent creator.
            data = source.read_bytes()
            with target.open('xb') as output:
                output.write(data)
            if digest(target.read_bytes()) != item['source']['digest']:
                raise ValueError('新文件校验失败，原图已保留')
            checked_source(item).unlink()
            item['source']['path'] = target
    except Exception:
        item['name'] = old_name
        raise
    return {'id': key, 'name': item['name'], 'revision': item['revision']}


def export_and_remove(workspace, folder):
    items = list(workspace.images.values())
    sources = [(item, checked_source(item)) for item in items if item.get('source')]
    if not sources:
        raise ValueError('没有关联原文件的图片，请先通过“关联原文件夹”导入')
    if len({str(path).casefold() for _, path in sources}) != len(sources):
        raise ValueError('同一原图被重复导入，请删除重复项后重试')
    destination = Path(folder)
    destination.mkdir(parents=True, exist_ok=True)
    if any((destination / item['name']).exists() for item in items):
        raise ValueError('目标文件夹存在同名文件，请选择空文件夹；原图已保留')
    for item in items:
        with (destination / item['name']).open('xb') as output:
            output.write(item['path'].read_bytes())
    # Verify every exported byte before removing any source.
    for item in items:
        if (Path(folder) / item['name']).read_bytes() != item['path'].read_bytes():
            raise ValueError('导出结果校验失败，全部原图已保留')
    removed, errors = 0, []
    for item, path in sources:
        try:
            checked_source(item).unlink()
            item.pop('source')
            removed += 1
        except (OSError, ValueError):
            errors.append(path.name)
    return {'count':len(items), 'removed':removed, 'errors':errors, 'folder':Path(folder).name}
