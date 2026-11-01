---
name: dtb-ppt-spec
description: PPT 默认模板（公司模板.pptx）与格式规范的唯一入口。凡是要新建/生成 PPT、检查某份 PPT 是否符合规范、批量统一字体字号表头页码版式、或询问"PPT 该用什么模板/字号/颜色"时使用。触发词：做PPT、生成PPT、PPT模板、公司模板、PPT格式、PPT规范、PPT排版、PPT检查、PPT改版、统一字体、统一字号、版式规范、微软雅黑、表头样式、页码缺失。
agent_created: true
---

# PPT 输出规范（默认模板：公司模板.pptx）

规则的唯一真源在 GitHub 私有仓库 `jiney87/dtb-ppt-spec`；本技能目录只是**本地缓存**，
其中 `assets/company-template.pptx` 就是**默认模板本体**。因此每次任务前必须先核对版本。

## 0. 强制首步：核对规则版本

**在读取规范、生成或修改任何 PPT 之前，先运行**：

```bash
SK="$HOME/.workbuddy/skills/dtb-ppt-spec"   # 技能安装目录；项目级则用 .workbuddy/skills/dtb-ppt-spec
python3 "$SK/scripts/sync_rules.py" --check
```

| 输出 | 动作 |
|---|---|
| `RULE_CHECK=latest` | 直接用本地规范与模板继续 |
| `RULE_CHECK=outdated ...` | **必须先** `--update`（会同时更新模板与脚本），成功后再继续 |
| `RULE_CHECK=updated` | 同步完成，读 `references/ppt-format-spec.md` 再继续 |
| `RULE_CHECK=offline` | 用本地缓存继续，但**交付说明里必须写明**："本次依据本地缓存 vX.X.X，未联网校验" |
| `RULE_CHECK=unauthorized` | 缺 GitHub Token，见第 4 节；不要重试，直接告知用户 |

同步后必须**重新读取** `references/ppt-format-spec.md`，不能凭记忆沿用旧规则。
`--check` 只发 1 次请求，不要为省一步而跳过。

## 1. 默认模板：只能用这一个

**凡是输出 PPT，一律从 `assets/company-template.pptx` 起稿。用户没指定模板时，这就是默认，
不需要再问。** 不要新建空白演示文稿，不要从网上/其他项目拷模板。

| 编号 | 规则 |
|---|---|
| T1 | 起稿跑 `new_deck.py`（复制模板生成骨架），**不要 `Presentation()` 新建空白文件** |
| T2 | 只用模板自带 7 个版式中的 6 个；禁止新建版式、禁止改母版与版式几何 |
| T3 | 底图、右上 logo、左上蓝竖条、浅色"品质源于热爱"水印都是品牌资产，**不得删除/遮盖/改色** |
| T4 | 用户明确指定别的模板时才切换；切换后仍须满足字体/字号/色板三约束 |
| T5 | 检查文件时，版式名不在模板清单内 → 判「模板不符」，先换模板再谈格式 |

模板本身受版本管理：`manifest.json` 记录其 sha256，`--update` 会自动同步，全设备一致。

## 2. 规范速查（全文见 references/ppt-format-spec.md）

**一套网格、三档字号、一种字体、一套色板。**

| 项 | 值 |
|---|---|
| 画布 | 16:9，33.87 × 19.05 cm（1280 × 720 px 渲染坐标） |
| 版心（内容区） | 左 55.0 / 上 135.0 / 宽 1151.0 / 高 520.0；右界 1206.0，下界 655.0 |
| 页标题 | L66.0 / T50.0 / W1030.0 / H52.0，**28pt 加粗 #0084D5**，单行 ≤26 字 |
| 说明 | 标题比内容区右移 11px，是为避让左上蓝竖条——**别为"对齐"挪占位符** |
| 页码 | L917.0 / T668.0 / W288.0 / H38.0，右对齐，**12pt #7F7F7F**，封面与结语页不加 |
| 栏间距 | 24px；2 栏 563.5 / 3 栏 367.7 / 4 栏 269.8 / 6 栏 171.8 |
| 字号 | **仅 28 / 15 / 12 pt 三档**（页标题 / 栏目标题·导语·要点 / 正文·表格·图注·页码） |
| 特例页 | 封面主标题 48、间隔页标题 32、结语页"谢谢！"——**模板原生值，别改** |
| 字体 | **微软雅黑**，全篇唯一（含表格、图表、页码） |
| 主色 | #0084D5 品牌蓝；#C00000 风险红；#7F7F7F 注释灰；#DDEBF7 浅蓝底 |
| 表头 | #0084D5 底 + 白字 12pt 加粗，行高 36px；数据行 28px / 12pt / 无竖线 |
| 版式 | 6 个：封面 / 间隔页 / 正文页1 / 正文页2 / 图片页 / 1_Section Header（结语页）；`自定义版式` 禁用 |
| 禁用色 | #FF0000 #C9C8C7 #FEFFFE #4EA5FF #61BA46 #008BCE #0070C0 #00B0F0 #2F75B5 #B5C6EA |

**加粗**：只有页标题、表头、卡片标题、关键数字加粗；正文不加粗，强调用颜色。
**行距**：正文 1.5，表格与图注 1.0；段前 0 / 段后 6pt；正文左对齐，禁两端对齐。

## 3. 脚本

```bash
SK="$HOME/.workbuddy/skills/dtb-ppt-spec"

# 版本核对（每次任务第一步）
python3 "$SK/scripts/sync_rules.py" --check      # --update / --status

# ① 新建 PPT：从默认模板生成骨架（封面→目录→章节页+正文页→结语页）
python3 "$SK/scripts/new_deck.py" 输出.pptx --title "2027年DTB项目\n战略规划" \
        --subtitle "汇报版 V1.0" --date "2026年11月" \
        --sections "一、经营回顾|二、问题诊断|三、下季动作" --body 2

# ② 合规自检：只读，按类型分组输出违规清单（附页码与原文）
python3 "$SK/scripts/ppt_format_check.py" <file.pptx> [--brief]

# ③ 批量改格式：默认干跑，--apply 才写文件
python3 "$SK/scripts/apply_format.py" <file.pptx> [--apply] [-o 目标.pptx] \
        [--only font,size,color,title,table,pagenum] [--zebra]
```

- `new_deck.py` 生成即合规（脚本已套微软雅黑 + 28/15/12 + 品牌蓝 + 页码）；封面标题可用 `\n` 手动断行。
- `apply_format.py` 只做确定性替换，不确定的一律报告不猜；放大字号会单独提示"需人工复核是否溢出"。
- 依赖 `python-pptx`；输出永不覆盖源文件。

## 4. 安装与凭据（多设备）

私有仓库，读取需 GitHub Token（fine-grained 勾 `Contents: Read`；classic 勾 `repo`）。每台设备配一次。

```bash
mkdir -p ~/.workbuddy/.secrets
printf '%s' '你的token' > ~/.workbuddy/.secrets/github_token
chmod 600 ~/.workbuddy/.secrets/github_token

# 本机有项目目录：在 工具/dtb-ppt-spec/ 下
bash install.sh                 # 含模板 assets/

# 本机无项目目录：一行自举
TOKEN=$(cat ~/.workbuddy/.secrets/github_token)
curl -s -H "Authorization: Bearer $TOKEN" -H "Accept: application/vnd.github.raw" \
  "https://api.github.com/repos/jiney87/dtb-ppt-spec/contents/bootstrap.sh?ref=main" | bash

python3 "$SK/scripts/sync_rules.py" --check    # 应返回 latest
```

也支持环境变量 `GITHUB_TOKEN` / `GH_TOKEN`，优先级高于文件。
**Token 只存本机，不得写进仓库、不得贴进对话。**

## 5. 更新规则 / 换模板

规则与模板都走 GitHub 仓库，**不要直接改本地缓存文件**（会被覆盖）：

1. 改本地仓库副本 `工具/dtb-ppt-spec/`（`references/`、`scripts/`、`skill/SKILL.md`；换模板则替换 `assets/company-template.pptx`）
2. 更新 `references/changelog.md`
3. 发布（自动重算 manifest 与 sha256）：
   ```bash
   cd 工具/dtb-ppt-spec
   python3 publish_rules.py --bump 1.1.0 --note "新增XX规则" --push
   ```
4. 各设备下次 `--check` 返回 `outdated`，`--update` 自动追平（模板一并更新）

版本号语义：修文字/补案例 → PATCH；新增规则项、改阈值、换模板 → MINOR；不兼容重写 → MAJOR。

## 6. 不要做的事

- 不要跳过版本核对，也不要凭记忆用旧的字号/网格值
- 不要用 `Presentation()` 建空白稿，不要自选模板，不要碰母版与版式几何
- 不要删/盖底图、logo、左上蓝竖条、水印
- 不要覆盖用户源文件；不要删页删图来"达标"
- 不要在正文用 12pt 以下字号"塞得下"——放不下改结构（拆表、提炼要点）
- 不要在表格里混用字号；不要用两端对齐；不要用阴影
- 不要把规范全文贴进对话，按需读文件、只引相关段落
- 不要手工改本地 `SKILL.md` / `references/` / `assets/`——由 `sync_rules.py` 覆盖，改动会丢
