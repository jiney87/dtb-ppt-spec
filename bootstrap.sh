#!/usr/bin/env bash
# 在一台「没有本项目目录」的新设备上，直接从 GitHub 安装 dtb-ppt-spec 技能。
# 只依赖 bash + python3 + 一个 Token；不需要 iCloud 项目目录同步到位。
#
# 用法（配好 Token 之后）：
#   bash bootstrap.sh
#   DTB_SKILL_DIR=~/.workbuddy/skills/dtb-ppt-spec bash bootstrap.sh
#
# 也可远程一行执行：
#   TOKEN=$(cat ~/.workbuddy/.secrets/github_token)
#   curl -s -H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github.raw" \
#     "https://api.github.com/repos/jiney87/dtb-ppt-spec/contents/bootstrap.sh?ref=main" | bash
set -euo pipefail

REPO="${DTB_SPEC_REPO:-jiney87/dtb-ppt-spec}"
REF="${DTB_SPEC_REF:-main}"
API="${GITHUB_API:-https://api.github.com}"
DST="${DTB_SKILL_DIR:-$HOME/.workbuddy/skills/dtb-ppt-spec}"

# --- 1. 取 Token ---------------------------------------------------------
TOKEN="${GITHUB_TOKEN:-${GH_TOKEN:-}}"
if [ -z "$TOKEN" ]; then
  for f in "$HOME/.workbuddy/.secrets/github_token" "$HOME/.workbuddy/.github_token"; do
    if [ -f "$f" ]; then TOKEN="$(tr -d '\n\r' < "$f")"; break; fi
  done
fi
if [ -z "$TOKEN" ]; then
  cat >&2 <<'EOF'
缺 Token。先在 GitHub 建一个 fine-grained PAT（Repository permissions → Contents: Read），
然后写入本机（Token 不要贴进对话或提交到仓库）：

  mkdir -p ~/.workbuddy/.secrets
  printf '%s' '你的token' > ~/.workbuddy/.secrets/github_token
  chmod 600 ~/.workbuddy/.secrets/github_token
EOF
  exit 2
fi

echo "→ 目标目录 $DST"
mkdir -p "$DST/scripts"

# --- 2. 先单独把同步器拉下来（它自己不在 manifest 的可校验范围里）---------
python3 - "$REPO" "$REF" "$API" "$DST" "$TOKEN" <<'PY'
import sys, os, json, base64, urllib.request

repo, ref, api, dst, tok = sys.argv[1:6]
url = f"{api}/repos/{repo}/contents/scripts/sync_rules.py?ref={ref}"
req = urllib.request.Request(url, headers={
    "Accept": "application/vnd.github+json",
    "Authorization": f"Bearer {tok}",
    "User-Agent": "dtb-ppt-spec-bootstrap/1.0",
    "X-GitHub-Api-Version": "2022-11-28",
})
try:
    with urllib.request.urlopen(req, timeout=20) as r:
        d = json.loads(r.read().decode())
except Exception as e:
    sys.exit(f"✗ 下载 sync_rules.py 失败: {e}\n  检查 Token 是否有该私有仓库的 Contents: Read 权限")

blob = base64.b64decode(d["content"])
p = os.path.join(dst, "scripts", "sync_rules.py")
with open(p, "wb") as fh:
    fh.write(blob)
print(f"✓ sync_rules.py  {len(blob)}B")
PY

# --- 3. 全量拉取规则本体（含 SKILL.md / references / scripts）------------
cd "$DST"
GITHUB_TOKEN="$TOKEN" DTB_SPEC_REPO="$REPO" DTB_SPEC_REF="$REF" GITHUB_API="$API" \
  python3 scripts/sync_rules.py --update

echo
echo "✓ 安装完成：$DST"
echo "→ 重开一个 WorkBuddy 会话，技能 dtb-ppt-spec 才会被加载"
