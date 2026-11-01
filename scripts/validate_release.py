#!/usr/bin/env python3
"""发布门禁：语法、未定义名称、旧规则、运行冒烟和 manifest 完整性。"""
import argparse
import ast
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]


def validate(check_manifest=True):
    scripts = sorted((ROOT / "scripts").glob("*.py"))
    for path in scripts:
        source = path.read_text()
        compile(source, str(path), "exec")
        tree = ast.parse(source)
        for node in ast.walk(tree):
            if isinstance(node, ast.Name) and node.id in {"FONT", "SZ_L", "SZ_M", "SZ_S"}:
                raise ValueError(f"{path.name}: obsolete constant {node.id}")
    subprocess.run([sys.executable, "-m", "pyflakes", *map(str, scripts)], check=True)
    for name in ["references/ppt-format-spec.md", "skill/SKILL.md"]:
        source = (ROOT / name).read_text()
        for pattern in [r"28\s*[/·]\s*15\s*[/·]\s*12", r"(?:12|15|32|48)\s*pt", r"(?:只有|仅用).*三档"]:
            if re.search(pattern, source):
                raise ValueError(f"{name}: obsolete rule {pattern}")
    with tempfile.TemporaryDirectory() as tmp:
        deck = str(Path(tmp) / "smoke.pptx")
        out = str(Path(tmp) / "formatted.pptx")
        def run(script, *args):
            return subprocess.run([sys.executable, str(ROOT / "scripts" / script), *args],
                                  check=True, capture_output=True, text=True).stdout
        run("new_deck.py", deck, "--title", "发布验证 Release 2027",
            "--subtitle", "测试", "--date", "2026年11月", "--sections", "一、经营")
        result = run("ppt_format_check.py", deck, "--brief")
        if "合计 0 处待修" not in result:
            raise ValueError(result)
        run("apply_format.py", deck)
        run("apply_format.py", deck, "--apply", "-o", out)
        result = run("ppt_format_check.py", out, "--brief")
        if "合计 0 处待修" not in result:
            raise ValueError(result)
        # 层级内原样保留；非层级吸附到最近层级（等距取小）；超范围不猜
        sys.path.insert(0, str(ROOT / "scripts"))
        import apply_format
        for size in apply_format.SZ_LEVELS:
            assert apply_format.map_size(size) == (size, False, False), size
        assert apply_format.map_size(15.5) == (16.0, True, True)     # 吸附·放大
        assert apply_format.map_size(10.5) == (10.0, False, True)    # 吸附·缩小
        assert apply_format.map_size(15) == (14.0, False, True)      # 等距取小
        assert apply_format.map_size(44) == (38.5, False, True)      # 超大并入最大层级
        assert apply_format.map_size(15.5, snap=False) == (None, False, False)
        assert apply_format.map_size(3) == (None, False, False)      # 低于 5 pt 不猜
        assert apply_format.map_size(72) == (None, False, False)     # 高于 60 pt 不猜
    if check_manifest:
        manifest = json.loads((ROOT / "manifest.json").read_text())
        for entry in manifest["files"]:
            data = (ROOT / entry["path"]).read_bytes()
            assert len(data) == entry["size"], entry["path"]
            assert hashlib.sha256(data).hexdigest() == entry["sha256"], entry["path"]
    print("PASS: compile, pyflakes, legacy rules, smoke, manifest")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-manifest", action="store_true")
    args = parser.parse_args()
    validate(not args.skip_manifest)
