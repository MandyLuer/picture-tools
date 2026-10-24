# 拾图 Shitu 品牌资源

软件名称：拾图 Shitu；版本：2.3.0。图标以青绿色文档页和抽出的照片表达图片提取与整理。

- `shitu.png`：使用内置 imagegen 工具生成的透明背景品牌原图。
- `shitu.ico`：Windows 多尺寸图标，含 16、24、32、48、64、128、256 像素。
- `../../frontend/public/shitu.png` 和 `favicon.ico`：界面及浏览器图标。

运行 `python scripts/prepare-icons.py` 可从原图重新生成以上分发格式。桌面打包脚本会自动执行转换；PNG 缩放和 ICO 编码使用 Pillow。

生成提示词：

> Use case: logo-brand. Create one production Windows desktop app icon for Shitu, a Word document image extraction and renaming utility. Square 1024x1024 icon asset. A deep teal rounded-square tile centered and filling 88% of canvas, transparent outside its rounded corners. Within the tile a strong minimal white document sheet silhouette with a folded top corner, and a smaller mint green landscape photo card emerging diagonally toward the bottom right from the page, conveying extracting images from documents. The photo card has one simple mountain silhouette and one small circular sun, excellent separation and clarity. Precise geometric silhouette, restrained subtle gradient within the teal tile, polished contemporary productivity app identity, crisp vector-like edges, front facing flat icon, generous internal padding. Large simple shapes remain identifiable at 16 and 32 pixels. No lettering, no Chinese characters, no words, no border, no watermark, no mockup, no external shadow, no extra objects. Transparent background outside tile.
