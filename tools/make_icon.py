"""
Render assets/icon.svg to assets/icon.png (512 px, used as the window icon) and assets/icon.ico (16-256 px, used
by PyInstaller for the executable). Run on Windows, where the Segoe UI font used by the SVG labels exists:

    python tools/make_icon.py
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "windows")  # the offscreen platform has no fonts
from PIL import Image  # noqa: E402
from PyQt6.QtCore import QRectF  # noqa: E402
from PyQt6.QtGui import QGuiApplication, QImage, QPainter  # noqa: E402
from PyQt6.QtSvg import QSvgRenderer  # noqa: E402

ASSETS = Path(__file__).resolve().parent.parent / "assets"
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]


def render(renderer, size):
    image = QImage(size, size, QImage.Format.Format_ARGB32)
    image.fill(0)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    renderer.render(painter, QRectF(0, 0, size, size))
    painter.end()
    return image


def main():
    app = QGuiApplication(sys.argv)  # noqa: F841 (Qt needs a running application object)
    renderer = QSvgRenderer(str(ASSETS / "icon.svg"))
    if not renderer.isValid():
        sys.exit(f"{ASSETS / 'icon.svg'} is not a valid SVG")
    render(renderer, 512).save(str(ASSETS / "icon.png"))
    frames = []
    for size in ICO_SIZES:
        temp = ASSETS / f"_icon_{size}.png"
        render(renderer, size).save(str(temp))
        frames.append(Image.open(temp).convert("RGBA"))
        temp.unlink()
    frames[-1].save(ASSETS / "icon.ico", format="ICO", sizes=[(s, s) for s in ICO_SIZES], append_images=frames[:-1])
    print(f"written {ASSETS / 'icon.png'} and {ASSETS / 'icon.ico'}")


if __name__ == "__main__":
    main()
