"""Собирает страницу для Firebase или GitHub Pages: берёт gas/index.html и вписывает ссылку /exec.

Запуск из папки проекта:
    python3 -I tools/build.py https://script.google.com/macros/s/…/exec          → public/index.html (Firebase)
    python3 -I tools/build.py https://script.google.com/macros/s/…/exec docs     → docs/index.html (GitHub Pages)

Исходник страницы один — gas/index.html. Собранный файл руками не правим: он перезаписывается.
"""
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
MARKER = 'const API_URL = "";'

if len(sys.argv) < 2 or not re.fullmatch(r"https://script\.google\.com/macros/s/[\w-]+/exec", sys.argv[1]):
    sys.exit("Нужна ссылка веб-приложения вида https://script.google.com/macros/s/…/exec")
url = sys.argv[1]
out_dir = ROOT / (sys.argv[2] if len(sys.argv) > 2 else "public")

src = (ROOT / "gas" / "index.html").read_text(encoding="utf-8")
if MARKER not in src:
    sys.exit(f"В gas/index.html нет строки {MARKER} — верните её, по ней вписывается ссылка")

out_dir.mkdir(exist_ok=True)
out = out_dir / "index.html"
out.write_text(src.replace(MARKER, f'const API_URL = "{url}";', 1), encoding="utf-8")
print("готово:", out)
