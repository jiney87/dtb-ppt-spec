# dtb-ppt-spec

PPT 默认伊利模板、咨询式组件规范与输出规则的**唯一真源**。

> **默认模板**：`assets/company-template.pptx`（即《公司模板.pptx》）。
> 之后所有 PPT 输出都从它起稿，用户未特别指定模板时不需要再问。

- 仓库：`jiney87/dtb-ppt-spec`（Private）
- 当前版本：见 `manifest.json` 的 `version`
- 规范正文：`references/ppt-format-spec.md`
- 各设备上的 SKILL 只是本地缓存，靠 `manifest.json` 的版本号与 sha256 判断是否落后

## 仓库结构

```
manifest.json                    版本清单（version / released_at / files[].sha256 / dest）
bootstrap.sh                     新设备一行自举安装（不需要项目目录）
install.sh                       从本地副本安装（项目 iCloud 已同步时用）
scripts/validate_release.py      发布门禁：静态检查、冒烟与 manifest 校验
assets/
  company-template.pptx          默认模板本体（1.5MB，二进制，受 sha256 管理）
skill/
  SKILL.md                       技能定义（含「强制首步：核对规则版本」）
references/
  ppt-format-spec.md             规范正文（人读 + Agent 读）
  changelog.md                   版本变更记录
scripts/
  sync_rules.py                  版本同步器（--check / --update / --status）
  new_deck.py                    从默认模板生成新稿骨架
  ppt_format_check.py            合规自检（只读）
  apply_format.py                格式调整（默认干跑）
```

`manifest.json` 里 `files[]` 只登记**运行时文件**（`skill/`、`references/`、`scripts/`、`assets/`）；
`README.md`、`bootstrap.sh`、`install.sh`、`publish_rules.py` 属于仓库自身工具，不进清单。
条目中的 `dest` 表示客户端落点：`skill/SKILL.md` → 本地根目录 `SKILL.md`。

## 版本规则

| 变更 | 版本位 | 例子 |
|---|---|---|
| 改错别字、补案例、润色措辞 | PATCH | 1.0.0 → 1.0.1 |
| 新增规则项、调整阈值/色值 | MINOR | 1.0.1 → 1.1.0 |
| 不兼容重写（如三档字号改四档） | MAJOR | 1.1.0 → 2.0.0 |

## 发布一次新版本

当前稳定版本为 **2.0.1**。本版本修复 2.0.0 不完整迁移造成的脚本和规范冲突。

1. 修改仓库源文件并更新 changelog。
2. 安装检查依赖：`python -m pip install python-pptx pyflakes`。
3. 运行 `python scripts/validate_release.py --skip-manifest` 校验内容。
4. 更新 manifest 版本、发布日期和每个运行文件的字节数、SHA-256；新增运行脚本须加入清单。
5. 运行 `python scripts/validate_release.py`，全部通过后将内容与 manifest 一起提交并推送 main。
6. 从远端读取新提交，重新验证文件指纹和脚本。使用 `python publish_rules.py` 自动校验后推送。GitHub Actions 配置见 `.github/workflows/validate.yml`，对 push、pull_request 和手动触发执行同一套检查。

Git 提交原子发布内容与清单。不要使用逐文件写入方式提前发布 manifest。
检查器只能覆盖已实现的规则，不能替代视觉检查。改版器遇到不明字号角色会保留并报告人工确认。

## 新设备安装

Token 只需配一次（fine-grained 勾 `Contents: Read`；classic 勾 `repo`）：

```bash
mkdir -p ~/.workbuddy/.secrets
printf '%s' '你的token' > ~/.workbuddy/.secrets/github_token
chmod 600 ~/.workbuddy/.secrets/github_token
```

装技能，两条路二选一：

```bash
# A. 本机有项目目录（iCloud 已同步）
cd 工具/dtb-ppt-spec && bash install.sh

# B. 本机没有项目目录：从 GitHub 自举
TOKEN=$(cat ~/.workbuddy/.secrets/github_token)
curl -s -H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github.raw" \
  "https://api.github.com/repos/jiney87/dtb-ppt-spec/contents/bootstrap.sh?ref=main" | bash
```

安装后技能位于 `~/.workbuddy/skills/dtb-ppt-spec/`，**重开一个 WorkBuddy 会话**才会加载。

## 客户端同步

每次调用技能时由 Agent 执行，不需要人记：

```bash
SK="$HOME/.workbuddy/skills/dtb-ppt-spec"
python3 "$SK/scripts/sync_rules.py" --check    # 1 次请求，比对版本 + 逐文件 sha256
python3 "$SK/scripts/sync_rules.py" --update   # 落后则下载，校验 sha256，备份旧版到 .backup/<ver>/
python3 "$SK/scripts/sync_rules.py" --status    # 本地版本、文件指纹、是否被手工改过
```

输出状态：`latest` / `outdated` / `updated` / `offline` / `unauthorized`。
Token 来源优先级：`GITHUB_TOKEN`、`GH_TOKEN`、`DTB_SPEC_TOKEN` 环境变量 →
`~/.workbuddy/.secrets/github_token` → `~/.workbuddy/.github_token`。

私有仓库只能走 `api.github.com`（本机 `raw.githubusercontent.com` 被拦）。同步器取文件时：
`/contents/` 正常返回 base64；**超过 1MB 的文件该接口不返回 content**，自动回退到
`/git/blobs/{sha}` 取原文——模板（5.8MB）就走这条路径。

## 注意

- 规则改动**只能**在仓库改，不要直接改某台设备上的缓存文件——会被下次同步覆盖。
- `scripts/apply_format.py` 默认干跑，`--apply` 才写文件，且永远输出到新文件。
- 规范以「先实测、后裁决」为原则：每条规则都要能对应到基准文件的实测数据。


## v2.0 核心变化

- 默认模板仍为伊利 `assets/company-template.pptx`，未指定模板时强制调用。
- 咨询式组件库与模板解耦，不再把咨询公司PPT当母版。
- 去除原咨询公司品牌身份，只保留专业规范、组件结构和颜色功能逻辑。
- 中文字体：微软雅黑；英文与数字：Verdana。
- 字号采用咨询式层级，不再使用旧版 28/15/12 三档限制。
- ChatGPT Library 组件资产：`/DTB/PPT组件库/consulting-components-sanitized.pptx`。
