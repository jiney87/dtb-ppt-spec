# dtb-ppt-spec

《DTB项目专项汇报》PPT 格式规范的**唯一真源**。规范正文、自检脚本、改格式脚本都放这里，
各设备上的 SKILL 只做本地缓存，靠 `manifest.json` 的版本号与 sha256 判断是否落后。

## 仓库结构

```
manifest.json                    版本清单（version / released_at / files[].sha256）
references/
  ppt-format-spec.md             规范正文（人读 + Agent 读）
  changelog.md                   版本变更记录
scripts/
  ppt_format_check.py            合规自检（只读）
  apply_format.py                格式调整（默认干跑）
```

## 版本规则

| 变更 | 版本位 | 例子 |
|---|---|---|
| 改错别字、补案例、润色措辞 | PATCH | 1.0.0 → 1.0.1 |
| 新增规则项、调整阈值/色值 | MINOR | 1.0.1 → 1.1.0 |
| 不兼容重写（如三档字号改四档） | MAJOR | 1.1.0 → 2.0.0 |

## 发布一次新版本

1. 改 `references/` 或 `scripts/` 下的文件
2. 更新 `references/changelog.md`
3. 重算 `manifest.json`：

```bash
python3 - <<'PY'
import json, hashlib, datetime, pathlib
root = pathlib.Path('.')
files = []
for p in ('references/ppt-format-spec.md', 'references/changelog.md',
          'scripts/ppt_format_check.py', 'scripts/apply_format.py'):
    b = pathlib.Path(p).read_bytes()
    files.append({'path': p, 'sha256': hashlib.sha256(b).hexdigest(), 'size': len(b)})
m = json.load(open('manifest.json', encoding='utf-8'))
m['version'] = input('new version: ')
m['released_at'] = datetime.date.today().isoformat()
m['files'] = files
json.dump(m, open('manifest.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
print(json.dumps(m, ensure_ascii=False, indent=2))
PY
```

4. 提交推送。客户端下次调用 `sync_rules.py --check` 即返回 `outdated`，`--update` 拉取。

## 新设备安装

每台设备只需两步（Token 只配一次）：

```bash
# 1) 配 Token（fine-grained PAT，只勾 Contents: Read；写入本机，勿提交）
mkdir -p ~/.workbuddy/.secrets
printf '%s' 'ghp_xxxxxxxx' > ~/.workbuddy/.secrets/github_token && chmod 600 ~/.workbuddy/.secrets/github_token

# 2) 安装技能并校验版本（在本目录执行）
bash install.sh
```

安装后技能位于 `~/.workbuddy/skills/dtb-ppt-spec/`，WorkBuddy 重启会话即可调用。

## 客户端同步

每次调用技能时由 Agent 执行，不需要人记：

```bash
python3 scripts/sync_rules.py --check    # 1 次请求，比对 manifest 版本 + 逐文件 sha256
python3 scripts/sync_rules.py --update   # 落后则下载，校验 sha256，备份旧版到 .backup/<ver>/
python3 scripts/sync_rules.py --status    # 本地版本、文件指纹、是否被手工改过
```

Token 来源优先级：`GITHUB_TOKEN` / `GH_TOKEN` 环境变量 → `~/.workbuddy/.secrets/github_token`
→ `~/.workbuddy/.github_token`。

## 注意

- 规则改动**只能**在仓库改，不要直接改某台设备上的缓存文件——会被下次同步覆盖。
- `scripts/apply_format.py` 默认干跑，`--apply` 才写文件，且永远输出到新文件。
- 规范以「先实测、后裁决」为原则：每条规则都要能对应到基准文件的实测数据。
