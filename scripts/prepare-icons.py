"""Convert the generated brand master into desktop and web icon formats."""
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent


def main():
    with Image.open(ROOT / "assets/branding/shitu.png") as source:
        image = source.convert("RGBA")
        image.save(ROOT / "assets/branding/shitu.ico", sizes=[(n, n) for n in (16, 24, 32, 48, 64, 128, 256)])
        public = ROOT / "frontend/public"
        public.mkdir(parents=True, exist_ok=True)
        image.resize((256, 256), Image.Resampling.LANCZOS).save(public / "shitu.png")
        image.save(public / "favicon.ico", sizes=[(n, n) for n in (16, 32, 48)])


if __name__ == "__main__":
    main()
