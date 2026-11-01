---
name: dtb-ppt-spec
description: DTB 项目 PPT 格式规范与改版工具。当需要生成 DTB 专项汇报/年度汇报/季度复盘 PPT、检查某份 PPT 是否符合格式规范、批量调整已有 PPT 的字体字号/表头/页码/版式、或询问"DTTB PPT 该用什么字号字体颜色"时使用。触发词：PPT格式、PPT规范、PPT排版、PPT检查、PPT改版、统一字体、统一字号、版式规范、微软雅黑、表头样式、页码缺失。
agent_created: true
---

# DTB 项目 PPT 格式规范

规则的唯一真源在 GitHub 私有仓库 `jiney87/dtb-ppt-spec`。本技能目录下的
`references/ppt-format-spec.md` 只是**本地缓存**。因此每次执行任务前必须先核对版本。

## 0. 强制首步：核对规则版本

**在读取规范、生成或修改任何 PPT 之前，先运行**：

```bash
SK="$HOME/.workbuddy/skills/dtb-ppt-spec"   # 技能安装目录；若装在项目级则用 .workbuddy/skills/dtb-ppt-spec
python3 "$SK/scripts/sync_rules.py" --check
```

按输出分支处理：

| 输出 | 动作 |
|---|---|
| `RULE_CHECK=latest` | 直接用本地规范继续，无需其他操作 |
| `RULE_CHECK=outdated ...` | **必须先**运行 `python3 "$SK/scripts/sync_rules.py" --update`，成功后再读规范 |
| `RULE_CHECK=updated` | 同步已完成，读 `references/ppt-format-spec.md` 再继续 |
| `RULE_CHECK=offline` | 网络不可用。用本地缓存版本继续，但**必须在交付说明里写明**："本次依据本地缓存 vX.X.X，未能联网校验是否为最新" |
| `RULE_CHECK=unauthorized` | 缺 GitHub Token。见下方「安装与凭据」；不要重试，直接告知用户 |

同步后必须**重新读取** `references/ppt-format-spec.md`（内容可能已变），不能凭记忆沿用旧规则。

`--check` 只发 1 次请求，代价很低，不要为了省一步而跳过。

## 1. 任务路由

| 用户意图 | 做法 |
|---|---|
| 做一份新 PPT | 先读 `references/ppt-format-spec.md`，按其网格/字号/色板搭建；产出后用 `scripts/ppt_format_check.py` 自检，报告必须附合规结论 |
| 调整已有 PPT 的格式 | 先 `python3 "$SK/scripts/apply_format.py" <文件>` 干跑看影响面 → 向用户汇报条数与风险页 → 确认后加 `--apply` 生成 `<原名>_规范版.pptx`，不覆盖源文件 |
| 只问规范 | 只读 `references/`，不要动脚本，不要改任何 PPT |
| 审核别人做的 PPT | 只跑 `scripts/ppt_format_check.py`，输出问题清单，不做修改 |

**改稿铁律**：`apply_format.py` 默认干跑，`--apply` 才写文件；永远输出到新文件；不删页、不删图、
不改变文字内容；脚本报"不确定"的项目一律保留原状并在报告中列出，由人决定。

## 2. 规范速查（全文见 references/ppt-format-spec.md）

四条硬约束：**一套网格、三档字号、一种字体、一套色板。**

| 项 | 值 |
|---|---|
| 画布 | 16:9，33.87 × 19.05 cm（1280 × 720 px 渲染坐标） |
| 版心 | 左 79.1px / 上 134.5px / 宽 1126.2px / 高 519.7px；右界 1205.3 |
| 页标题 | L79.1 / T50.1 / W1126.2 / H52.3，**28pt 加粗 #0084D5**，单行 ≤26 字 |
| 正文区 | 导语带 106–130；内容带 134.5–654.2 |
| 页码 | L917.3 / T667.8 / W288.0 / H38.3，右对齐，**12pt #7F7F7F**，除封面外每页 1 个 |
| 栏间距 | 24px；2 栏 551.1 / 3 栏 359.4 / 4 栏 263.6 / 6 栏 167.7 |
| 字号 | **仅 28 / 16 / 12 pt 三档**（L 页标题 / M 栏目标题 / S 正文·表格·图注） |
| 字体 | **微软雅黑**，全篇唯一（含表格、图表、页码） |
| 主色 | #0084D5 品牌蓝（标题/表头/强调）；#C00000 风险红；#7F7F7F 注释灰 |
| 表头 | #0084D5 底 + 白字 12pt 加粗，行高 36px；数据行 28px，12pt，无竖线 |
| 版式 | 只允许 4 个：封面 / 间隔页 / 正文页1 / 图片页 |
| 禁用色 | #FF0000 #C9C8C7 #FEFFFE #4EA5FF #61BA46 #008BCE #0070C0 #00B0F0 #2F75B5 #B5C6EA |

**加粗规则**：只有页标题、表头、卡片标题、关键数字加粗；正文一律不加粗，强调用颜色。
**行距**：正文 1.5，表格与图注 1.0；段前 0 / 段后 6pt；正文左对齐，禁用两端对齐。

## 3. 脚本

以下均以 `SK` 为技能安装目录（默认 `~/.workbuddy/skills/dtb-ppt-spec`）：

```bash
SK="$HOME/.workbuddy/skills/dtb-ppt-spec"

# 版本核对（每次任务第一步）
python3 "$SK/scripts/sync_rules.py" --check
python3 "$SK/scripts/sync_rules.py" --update
python3 "$SK/scripts/sync_rules.py" --status

# 合规自检：只读，输出按类型分组的违规清单（附页码与原文）
python3 "$SK/scripts/ppt_format_check.py" <file.pptx> [--brief]

# 格式调整：默认干跑，--apply 才写文件
python3 "$SK/scripts/apply_format.py" <file.pptx> [--apply] [-o 目标.pptx] [--only font,size,color,title,table,pagenum] [--zebra]
```

脚本依赖 `python-pptx`（已随运行环境提供）；输出文件与源文件同目录时不会覆盖源文件。

`apply_format.py` 处理后典型效果：2,877 处偏差 → 1 处（仅剩版式越界这类需人工判断的项）。
字号被放大的页面脚本会单独列出"⚠ 需人工复核是否溢出"，交付时须转述。

## 4. 安装与凭据（多设备）

规则仓库是**私有**的，读取需要 GitHub Token（fine-grained 勾 `Contents: Read` 即可；
classic 则勾 `repo`）。每台设备配一次 Token，之后全自动。

```bash
# 1) 配 Token（仅首次；Token 不要贴进对话、不要提交到仓库）
mkdir -p ~/.workbuddy/.secrets
printf '%s' '你的token' > ~/.workbuddy/.secrets/github_token
chmod 600 ~/.workbuddy/.secrets/github_token

# 2a) 本机有项目目录：在 工具/dtb-ppt-spec/ 下执行
bash install.sh

# 2b) 本机没有项目目录：一行从 GitHub 自举安装
TOKEN=$(cat ~/.workbuddy/.secrets/github_token)
curl -s -H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github.raw" \
  "https://api.github.com/repos/jiney87/dtb-ppt-spec/contents/bootstrap.sh?ref=main" | bash

# 3) 校验（应返回 latest）
python3 "$SK/scripts/sync_rules.py" --check
```

也支持环境变量 `GITHUB_TOKEN` / `GH_TOKEN`，优先级高于文件。
**Token 只存本机，不得写进仓库、不得贴进对话。**

## 5. 更新规则（当用户要改规范时）

规则改动走 GitHub 仓库，再让各设备同步，**不要直接改本地缓存文件**：

1. 改本地仓库副本 `工具/dtb-ppt-spec/`（`references/`、`scripts/`、`skill/SKILL.md`）
2. 更新 `references/changelog.md`
3. 发布（自动重算 manifest 并推送，无需手改 sha256）：
   ```bash
   cd 工具/dtb-ppt-spec
   python3 publish_rules.py --bump 1.1.0 --note "新增XX规则" --push
   ```
4. 各设备下次 `--check` 即返回 `outdated`，`--update` 自动追平

版本号语义：修文字/补案例 → PATCH（1.0.0→1.0.1）；新增规则项或改阈值 → MINOR；
不兼容重写（如三档改四档）→ MAJOR。

## 6. 不要做的事

- 不要跳过版本核对直接按记忆改 PPT
- 不要覆盖用户源文件；不要批量删除页面/对象来"达标"
- 不要把 12pt 以下字号用于"塞得下"——放不下应改结构（拆表、提炼要点）
- 不要在表格里混用字号；不要在正文里用两端对齐
- 不要把规范文档全文贴进对话，需要时读文件、只引相关段落
- 不要手工改本地 `SKILL.md` 或 `references/`——它们由 `sync_rules.py` 覆盖，改动会丢
