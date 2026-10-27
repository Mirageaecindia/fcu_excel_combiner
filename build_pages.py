"""Copy the python core into docs/ so GitHub Pages (Pyodide) can load it.

Run after changing anything in core/ or the reference template:
    python build_pages.py
"""
import json
import shutil
from pathlib import Path

root = Path(__file__).resolve().parent
dst = root / "docs" / "py" / "core"
if dst.exists():
    shutil.rmtree(dst)
dst.mkdir(parents=True)
files = []
for f in sorted((root / "core").glob("*.py")):
    shutil.copy(f, dst / f.name)
    files.append(f"core/{f.name}")
files.append("web_api.py")
(root / "docs" / "py" / "manifest.json").write_text(json.dumps(files))
print("copied", len(files) - 1, "core modules ->", dst)
