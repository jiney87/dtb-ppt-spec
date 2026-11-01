#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
规则版本同步器 —— 每次调用 SKILL 前先跑这个。

用法：
  python3 sync_rules.py --check     # 只查版本（快，1 次请求）
  python3 sync_rules.py --update    # 落后则下载更新（幂等，已最新则空转）
  python3 sync_rules.py --status    # 看本地版本与文件指纹

输出（供 Agent 解析）：
  RULE_CHECK=latest   version=1.0.0
  RULE_CHECK=outdated local=1.0.0 remote=1.0.1 files=2
  RULE_CHECK=updated  version=1.0.1 files=2 backup=...
  RULE_CHECK=offline  local=1.0.0 reason=...
  RULE_CHECK=unauthorized local=1.0.0 reason=...

退出码：0 可用（含 latest/updated/offline 降级）；2 鉴权失败；3 本地无规则可退。
"""
import os, sys, json, base64, hashlib, urllib.request, urllib.error, shutil, time

REPO = os.environ.get("DTB_SPEC_REPO", "jiney87/dtb-ppt-spec")
REF = os.environ.get("DTB_SPEC_REF", "main")
API = os.environ.get("GITHUB_API", "https://api.github.com")
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE = os.path.join(SKILL_DIR, ".rule_version.json")
TOKEN_FILES = [
    os.path.expanduser("~/.workbuddy/.secrets/github_token"),
    os.path.expanduser("~/.workbuddy/.github_token"),
]


def token():
    for k in ("GITHUB_TOKEN", "GH_TOKEN", "DTB_SPEC_TOKEN"):
        v = os.environ.get(k)
        if v:
            return v.strip()
    for f in TOKEN_FILES:
        if os.path.exists(f):
            try:
                v = open(f).read().strip()
                if v:
                    return v
            except Exception:
                pass
    return None


def get(path):
    url = f"{API}/repos/{REPO}/contents/{path}?ref={REF}"
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json",
        "User-Agent": "dtb-ppt-spec-sync/1.0",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    t = token()
    if t:
        req.add_header("Authorization", f"Bearer {t}")
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())


def get_blob(sha, timeout=180):
    """Git Blobs API：Contents API 对 >1MB 的文件不返回 content，必须走这里。

    用 Accept: raw 直接拿二进制（5.8MB 而非 base64 后的 7.7MB），并分块读入：
    本机到 api.github.com 约 20KB/s 且服务端不支持 Range 续传，整块 read() 很容易中途断开。
    """
    url = f"{API}/repos/{REPO}/git/blobs/{sha}"
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github.raw",
        "User-Agent": "dtb-ppt-spec-sync/1.0",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    t = token()
    if t:
        req.add_header("Authorization", f"Bearer {t}")
    buf = bytearray()
    with urllib.request.urlopen(req, timeout=timeout) as r:
        while True:
            chunk = r.read(262144)
            if not chunk:
                break
            buf += chunk
    return bytes(buf)


def fetch_bytes(path, expect_sha=None, attempts=3):
    """取文件原始 bytes。模板等二进制由这里原样落盘。

    expect_sha 给了就校验；大文件（模板 5.8MB）在慢网络下可能中途断，失败重试。
    """
    last = None
    for i in range(attempts):
        try:
            data = get(path)
            if isinstance(data, list):
                raise ValueError(f"{path} 是目录，不是文件")
            if data.get("content") and data.get("encoding", "base64") == "base64":
                raw = base64.b64decode(data["content"])
            else:
                # >1MB：Contents API 返回 content="" / encoding="none"
                raw = get_blob(data["sha"])
            if expect_sha and sha(raw) != expect_sha:
                raise ValueError(f"校验和不符（拿到 {len(raw)}B）")
            return raw, data.get("sha")
        except Exception as e:
            last = e
            if i + 1 < attempts:
                print(f"   … {path} 第 {i + 1} 次失败（{type(e).__name__}），重试",
                      file=sys.stderr)
                time.sleep(2)
    raise last


# 兼容旧名（--check 只读 manifest，不涉及大文件）
def fetch_text(path):
    return fetch_bytes(path, attempts=1)


def sha(b):
    return hashlib.sha256(b).hexdigest()


def load_state():
    if os.path.exists(STATE):
        try:
            return json.load(open(STATE, encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_state(st):
    json.dump(st, open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def local_root():
    return SKILL_DIR


def local_path(f):
    """manifest 条目的本地落点：dest 优先（如 skill/SKILL.md → SKILL.md）"""
    return os.path.join(SKILL_DIR, f.get("dest") or f["path"])


def rel_dest(f):
    return f.get("dest") or f["path"]


def is_dirty(st):
    """本地被手工改过？"""
    for f in st.get("files", []):
        p = local_path(f)
        if not os.path.exists(p):
            return True
        if sha(open(p, "rb").read()) != f["sha256"]:
            return True
    return False


def cmd_status():
    st = load_state()
    if not st:
        print("RULE_CHECK=none  reason=尚未同步过规则")
        return 3
    dirty = is_dirty(st)
    print(f"RULE_CHECK=local  version={st.get('version')}  synced_at={st.get('synced_at')}  "
          f"files={len(st.get('files', []))}  dirty={dirty}")
    for f in st.get("files", []):
        print(f"   {f['path']}  {f['sha256'][:12]}  {f.get('size', '?')}B")
    return 0


def cmd_check(verbose=True):
    st = load_state()
    local_v = st.get("version")
    try:
        raw, _ = fetch_text("manifest.json")
        man = json.loads(raw.decode())
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 404):
            print(f"RULE_CHECK=unauthorized local={local_v or 'none'} "
                  f"reason=HTTP{e.code}（私有仓库需配置 GITHUB_TOKEN，见 SKILL.md 安装节）")
            return 2
        print(f"RULE_CHECK=offline local={local_v or 'none'} reason=HTTP{e.code}")
        return 0 if local_v else 3
    except Exception as e:
        print(f"RULE_CHECK=offline local={local_v or 'none'} reason={type(e).__name__}:{e}")
        return 0 if local_v else 3

    remote_v = man.get("version")
    files = man.get("files", [])
    changed = [f for f in files
               if not os.path.exists(local_path(f))
               or sha(open(local_path(f), "rb").read()) != f["sha256"]]

    if remote_v == local_v and not changed:
        if verbose:
            print(f"RULE_CHECK=latest version={remote_v} released={man.get('released_at')}")
        return 0

    if verbose:
        print(f"RULE_CHECK=outdated local={local_v or 'none'} remote={remote_v} files={len(changed)}")
        print(f"   notes: {man.get('notes', '')}")
    return 1


def cmd_update():
    st = load_state()
    local_v = st.get("version")
    try:
        raw, _ = fetch_text("manifest.json")
        man = json.loads(raw.decode())
    except urllib.error.HTTPError as e:
        msg = "unauthorized" if e.code in (401, 403, 404) else "offline"
        print(f"RULE_CHECK={msg} local={local_v or 'none'} reason=HTTP{e.code}")
        return 2 if msg == "unauthorized" else (0 if local_v else 3)
    except Exception as e:
        print(f"RULE_CHECK=offline local={local_v or 'none'} reason={type(e).__name__}:{e}")
        return 0 if local_v else 3

    remote_v = man.get("version")
    files = man.get("files", [])
    need = [
        f for f in files
        if not os.path.exists(local_path(f))
        or sha(open(local_path(f), "rb").read()) != f["sha256"]
    ]
    if not need:
        save_state({**st, "version": remote_v, "checked_at": time.strftime("%Y-%m-%d %H:%M:%S")})
        print(f"RULE_CHECK=latest version={remote_v}")
        return 0

    # 备份
    backup = None
    if st.get("version"):
        backup = os.path.join(SKILL_DIR, ".backup", st["version"])
        os.makedirs(backup, exist_ok=True)
        for f in files:
            src = local_path(f)
            if os.path.exists(src):
                dst = os.path.join(backup, rel_dest(f))
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(src, dst)

    done = []
    for f in need:
        big = f.get("size", 0) > 1_000_000
        if big:
            print(f"   ↓ 下载 {f['path']}（{f['size']/1e6:.1f}MB，网络慢时约需数分钟）",
                  file=sys.stderr)
        try:
            blob, _ = fetch_bytes(f["path"], expect_sha=f["sha256"])
        except Exception as e:
            print(f"RULE_CHECK=error local={local_v or 'none'} reason=下载 {f['path']} 失败: {e}")
            return 2
        dst = local_path(f)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        open(dst, "wb").write(blob)
        done.append(f["path"])

    save_state({
        "version": remote_v,
        "released_at": man.get("released_at"),
        "synced_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "source": f"{REPO}@{REF}",
        "files": [{"path": f["path"], "dest": f.get("dest"), "sha256": f["sha256"],
                   "size": len(open(local_path(f), 'rb').read())}
                  for f in files],
    })
    print(f"RULE_CHECK=updated version={remote_v} files={len(done)} backup={backup or '-'}")
    for d in done:
        print(f"   ↻ {d}")
    return 0


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "--check"
    if arg == "--status":
        sys.exit(cmd_status())
    elif arg == "--update":
        sys.exit(cmd_update())
    else:
        sys.exit(cmd_check())
