---
name: dtb-ppt-spec
description: PPT 默认伊利模板、咨询式组件库与输出规范的唯一入口。凡是要新建/生成 PPT、检查或改版 PPT、选择页面结构/图表/框架元素、统一字体字号色板版式、或询问"PPT 该用什么模板/组件/字号/颜色"时使用。用户未明确指定模板时必须使用伊利默认模板 assets/company-template.pptx；咨询式组件库只提供可复用元素和专业规范，不替代模板。
agent_created: true
---

# PPT 输出规范（默认模板：伊利公司模板；独立咨询式组件库）

规则的唯一真源在 GitHub 私有仓库 `jiney87/dtb-ppt-spec`；本技能目录只是**本地缓存**，
其中 `assets/company-template.pptx` 是**伊利默认模板本体**。咨询式组件库与模板完全解耦：模板决定品牌母版、版式底图与色系，组件库决定可选的表达结构、图表、框架和信息图元素。因此每次任务前必须先核对版本。

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

## 1. 模板与组件库：强制分离

**凡是输出 PPT，用户没有明确指定其他模板时，一律从 `assets/company-template.pptx`（伊利模板）起稿，不需要再问。** 不要新建空白演示文稿，不要自选其他模板。

**咨询式组件库不是模板。** 组件库用于选择和复用逻辑树、驱动树、矩阵、流程、时间线、Roadmap、组织图、数据图表、战略框架、信息图、图标等表达元素；允许拆解、组合、重排，但必须嵌入当前激活模板中，并继承当前模板的品牌色系。默认情况下即使用伊利色系。

| 编号 | 规则 |
|---|---|
| T1 | 起稿跑 `new_deck.py`（复制模板生成骨架），**不要 `Presentation()` 新建空白文件** |
| T2 | 只用模板自带 7 个版式中的 6 个；禁止新建版式、禁止改母版与版式几何 |
| T3 | 底图、右上 logo、左上蓝竖条、浅色"品质源于热爱"水印都是品牌资产，**不得删除/遮盖/改色** |
| T4 | 用户明确指定别的模板时才切换；组件库仍可调用，但颜色必须映射到用户指定模板的色系 |
| T5 | 检查文件时，版式名不在模板清单内 → 判「模板不符」，先换模板再谈格式 |
| C1 | 组件库与母版无绑定关系：**只取元素，不复制咨询公司品牌母版、logo、页脚、版权、公司名、专属命名** |
| C2 | 组件选择必须先判断页面意图，再选最简单且能说明问题的元素；禁止为了“像咨询公司”而堆复杂图形 |
| C3 | 允许跨页拆取组件重新组合；所有可编辑元素优先保持原生可编辑，不以截图替代 |
| C4 | 在 ChatGPT 环境优先从 Library `/DTB/PPT组件库/consulting-components-sanitized.pptx` 调用去品牌化组件库；本地/WorkBuddy 环境若仓库内存在 `assets/consulting-components-sanitized.pptx` 则优先使用该资产 |
| C5 | 若组件库资产不可访问，必须明确说明，只能依据组件规范重建，**不得声称已调用原组件** |

模板本身受版本管理：`manifest.json` 记录其 sha256，`--update` 会自动同步，全设备一致。

## 2. 规范速查（全文见 references/ppt-format-spec.md）

**一套伊利模板网格、一套咨询式字号层级、两套字体映射、一套伊利色板。**

| 项 | 值 |
|---|---|
| 画布 | 16:9，33.87 × 19.05 cm（1280 × 720 px 渲染坐标） |
| 版心（内容区） | 左 55.0 / 上 135.0 / 宽 1151.0 / 高 520.0；右界 1206.0，下界 655.0 |
| 页标题 | L66.0 / T50.0 / W1030.0 / H52.0，**24pt 加粗 #0084D5**，单行 ≤26 字 |
| 说明 | 标题比内容区右移 11px，是为避让左上蓝竖条——**别为"对齐"挪占位符** |
| 页码 | L917.0 / T668.0 / W288.0 / H38.0，右对齐，**6.5pt #7F7F7F**（**间隔页用 #FFFFFF**），封面与结语页不加 |
| 栏间距 | 24px；2 栏 563.5 / 3 栏 367.7 / 4 栏 269.8 / 6 栏 171.8 |
| 字号 | **文件标题18 / 文件副标题16 / 封面信息10；章节标题38.5 / 章节副标题38.5；引述高亮28 / 引述内容24；页面标题24 / 页面副标题24；标题文字、正文、项目符号、图表文字14；图表标签、来源9；页脚6.5 pt** |
| 字体 | **中文：微软雅黑；英文与数字：Verdana**。同一文本框混排时通过 OOXML 的 eastAsia / latin 字体分别映射 |
| 说明 | 字号层级来自咨询式组件规范；伊利模板仅提供品牌母版与布局骨架，按语义角色选择字号 |
| 主色 | #0084D5 品牌蓝；#C00000 风险红；#7F7F7F 注释灰；#DDEBF7 浅蓝底 |
| **深底页** | **间隔页整页是品牌蓝底图（实测 #048ACE）→ 文字、页码、装饰线一律 `#FFFFFF`**。品牌蓝 `#0084D5` 与注释灰 `#7F7F7F` 在蓝底上对比度都只有 **1.05**，等于隐形 |
| 表头 | #0084D5 底 + 白字 14pt 加粗，行高 36px；数据行 28px / 14pt / 无竖线 |
| 版式 | 6 个：封面 / 间隔页 / 正文页1 / 正文页2 / 图片页 / 1_Section Header（结语页）；`自定义版式` 禁用 |
| 禁用色 | #FF0000 #C9C8C7 #FEFFFE #4EA5FF #61BA46 #008BCE #0070C0 #00B0F0 #2F75B5 #B5C6EA |

**加粗**：只有页标题、表头、卡片标题、关键数字加粗；正文不加粗，强调用颜色。
**行距**：正文 1.5，表格与图注 1.0；段前 0 / 段后 6pt；正文左对齐，禁两端对齐。
**对比度**：文字色服从可读性，色板让位。判据 WCAG —— 大字（≥18pt 加粗或 ≥24pt）≥3.0，其余 ≥4.5；同色相且明度差 <30% 直接判违规。深底用白，浅底用深色（详见 spec §3.5）。

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

- `new_deck.py` 生成即合规（中文微软雅黑 + 英文/数字 Verdana + 咨询式字号层级 + 伊利品牌色 + 页码；间隔页文字与页码为白色）；封面标题可用 `\n` 手动断行。
- `apply_format.py` 只做确定性替换，不确定的一律报告不猜；放大字号会单独提示"需人工复核是否溢出"；间隔页文字、页码、品牌蓝装饰自动改白（记入「深底改白」）。
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
3. 发布前安装 python-pptx、pyflakes，执行 `python3 scripts/validate_release.py --skip-manifest`；更新 manifest 版本和所有运行文件的 size/sha256 后执行 `python3 scripts/validate_release.py`，通过后将内容与清单一起提交并推送（详见仓库 README）。
4. 各设备下次 `--check` 返回 `outdated`，`--update` 自动追平（模板一并更新）

版本号语义：修文字/补案例 → PATCH；新增规则项、改阈值、换模板 → MINOR；不兼容重写 → MAJOR。

## 6. 不要做的事

- 不要跳过版本核对，也不要凭记忆用旧的字号/网格值
- 不要用 `Presentation()` 建空白稿，不要自选模板，不要碰母版与版式几何
- 不要删/盖底图、logo、左上蓝竖条、水印
- 不要覆盖用户源文件；不要删页删图来"达标"
- 不要为了“塞得下”擅自缩小咨询式字号层级——放不下改结构（拆表、提炼要点）
- 不要把咨询组件库当母版套用；不得带入原咨询公司名称、logo、版权、品牌页脚、专属图示命名或品牌色身份
- 不要在表格里无规则混用字号；不要用两端对齐；不要用阴影
- **不要在深底页（间隔页）上使用品牌蓝或注释灰**——文字、页码、装饰线一律白色；跟底色太近等于没写字
- 不要在白色弧线区域放白字；不要为了"守色板"牺牲可读性
- 不要把规范全文贴进对话，按需读文件、只引相关段落
- 不要手工改本地 `SKILL.md` / `references/` / `assets/`——由 `sync_rules.py` 覆盖，改动会丢
