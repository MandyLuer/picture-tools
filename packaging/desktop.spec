# Build on each target OS.
from pathlib import Path
root = Path(SPECPATH).parent
if not (root / 'frontend' / 'dist' / 'index.html').exists():
    raise RuntimeError('Run npm run build in frontend first')
a = Analysis([str(root / 'desktop.py')], pathex=[str(root)], binaries=[],
             datas=[(str(root / 'frontend' / 'dist'), 'frontend/dist'),
                    (str(root / 'assets' / 'branding' / 'shitu.ico'), 'assets/branding')],
             hiddenimports=['uvicorn.logging', 'uvicorn.loops.auto', 'uvicorn.protocols.http.auto', 'uvicorn.protocols.websockets.auto', 'uvicorn.lifespan.on'],
             hookspath=[], runtime_hooks=[], excludes=['tkinter'])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='拾图 Shitu',
          icon=str(root / 'assets' / 'branding' / 'shitu.ico'),
          version=str(root / 'packaging' / 'windows-version.txt'),
          debug=False, strip=False, upx=False, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='Shitu')
