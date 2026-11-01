#!/bin/bash
# 在本机安装 / 更新 dtb-ppt-spec 技能到 ~/.workbuddy/skills/
# 用法：bash install.sh          # 安装并校验版本
#      bash install.sh --offline # 只装文件，不联网校验
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"
DST="$HOME/.workbuddy/skills/dtb-ppt-spec"
PY="$(command -v python3 || command -v python)"

mkdir -p "$DST/scripts" "$DST/references" "$DST/assets"

# 按 manifest 复制（不硬编码文件名，避免新增文件漏装；dest 映射如 skill/SKILL.md → SKILL.md）
"$PY" - "$SRC" "$DST" <<'PY'
import json, sys, shutil, pathlib
src, dst = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
man = json.loads((src / 'manifest.json').read_text(encoding='utf-8'))
for f in man['files']:
    sp = src / f['path']
    if not sp.exists():
        sys.exit(f"✗ 源文件缺失：{f['path']}（仓库不完整）")
    dp = dst / (f.get('dest') or f['path'])
    dp.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(sp, dp)
    print(f"  + {f.get('dest') or f['path']}")
PY

# 写入版本指纹（后续由 sync_rules.py 维护）
"$PY" - "$SRC" "$DST" <<'PY'
import json, sys, hashlib, pathlib, time
src, dst = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
man = json.loads((src / 'manifest.json').read_text(encoding='utf-8'))
st = {
    "version": man["version"],
    "released_at": man.get("released_at"),
    "synced_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    "source": "jiney87/dtb-ppt-spec@main",
    "seeded": "install.sh",
    "files": [{"path": f["path"], "dest": f.get("dest"),
               "sha256": hashlib.sha256((dst / (f.get("dest") or f["path"])).read_bytes()).hexdigest(),
               "size": f["size"]} for f in man["files"]],
}
(dst / '.rule_version.json').write_text(json.dumps(st, ensure_ascii=False, indent=2), encoding='utf-8')
print(f"  installed v{st['version']}  {len(st['files'])} files")
PY

echo "✓ 已安装到 ${DST}（含默认模板 assets/company-template.pptx）"
if [ "$1" != "--offline" ]; then
  echo "→ 校验远端版本…"
  "$PY" "$DST/scripts/sync_rules.py" --check || true
  echo "  若返回 unauthorized：先配置 Token（见 README「新设备安装」）再运行"
  echo "  python3 $DST/scripts/sync_rules.py --update"
fi
