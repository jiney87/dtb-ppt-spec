# -*- coding: utf-8 -*-
"""
PPT 规范自检工具（v1.1.0，基准模板 = 公司模板.pptx）

用法：python3 ppt_format_check.py <file.pptx> [--brief]
输出：逐条违规清单 + 统计摘要。只读，不修改源文件。

v1.1 变更：网格改为模板占位符原值（标题 L66/T50、内容 55–1206 / 135–655）；
三档字号 28/15/12；新增「模板不符」总检（画布 / 主题蓝 / 版式名）。
"""
import sys
from collections import Counter, defaultdict

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# ---- 基准常量（全部取自 assets/company-template.pptx 原值）----
FONT = "微软雅黑"
SIZES = {28.0, 15.0, 12.0}                      # 三档
SPECIAL_SZ = {"封面": {48.0}, "间隔页": {32.0}}   # 特例页模板原生值
TITLE_SZ = 28.0
TITLE_POS = (66.0, 50.0)                        # px，允许 ±4
BODY_LAYOUTS = ("正文页1", "正文页2", "图片页")
LAYOUTS_OK = ("封面", "间隔页", "正文页1", "正文页2", "图片页", "1_Section Header")
LAYOUTS_KNOWN = LAYOUTS_OK + ("自定义版式",)     # 模板内存在但禁用
NO_PAGENUM = ("封面", "1_Section Header")        # 这两类不加页码
CANVAS = (12192000, 6858000)                    # 33.87 × 19.05 cm
THEME_DK1 = "0084D5"                            # 主题 dk1 = 品牌蓝
PXW = 1280.0

BANNED = {"FF0000", "C9C8C7", "FEFFFE", "4EA5FF", "61BA46",
          "008BCE", "0070C0", "00B0F0", "2F75B5", "B5C6EA"}
HDR_FILL = "0084D5"


def px(v, scale):
    return None if v is None else round(v * scale, 1)


def _rpr(r):
    return r._r.find(A + "rPr")


def run_font(r):
    rPr = _rpr(r)
    if rPr is None:
        return None, None
    lat = rPr.find(A + "latin")
    ea = rPr.find(A + "ea")
    return (lat.get("typeface") if lat is not None else None,
            ea.get("typeface") if ea is not None else None)


def run_color(r):
    rPr = _rpr(r)
    if rPr is None:
        return None
    sf = rPr.find(A + "solidFill")
    if sf is None:
        return None
    c = sf.find(A + "srgbClr")
    return c.get("val") if c is not None else None


def theme_dk1(prs):
    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
        el = root.find(f".//{A}clrScheme/{A}dk1")
        ch = list(el)[0]
        # sysClr 的 val 是 systemColor 名（如 windowText），真实色值在 lastClr
        if etree.QName(ch).localname == "sysClr":
            return (ch.get("lastClr") or "").upper()
        return (ch.get("val") or "").upper()
    except Exception:
        return None


def main(path, brief=False):
    prs = Presentation(path)
    scale = PXW / prs.slide_width
    issues = defaultdict(list)
    stat = Counter()

    # ---- 0. 模板指纹总检 ----
    if (prs.slide_width, prs.slide_height) != CANVAS:
        issues["T0 模板不符"].append(
            f"画布 {prs.slide_width}×{prs.slide_height}，应为 12192000×6858000（33.87×19.05cm）")
    dk = theme_dk1(prs)
    if dk and dk != THEME_DK1:
        issues["T0 模板不符"].append(f"主题 dk1={dk}，应为 #{THEME_DK1}（非公司模板）")

    titles, pagenums = [], Counter()
    unknown = {}

    for idx, s in enumerate(prs.slides, start=1):
        has_num = False
        layout_name = s.slide_layout.name
        allow = SIZES | SPECIAL_SZ.get(layout_name, set())

        if layout_name not in LAYOUTS_KNOWN:
            unknown[layout_name] = unknown.get(layout_name, 0) + 1

        for sh in s.shapes:
            if sh.is_placeholder and str(sh.placeholder_format.type).startswith("TITLE"):
                t = sh.text_frame.text.strip()
                x, y = px(sh.left, scale), px(sh.top, scale)
                titles.append((idx, x, y, t))
                is_body = layout_name in BODY_LAYOUTS
                if is_body and (abs((y or 0) - TITLE_POS[1]) > 4 or abs((x or 0) - TITLE_POS[0]) > 4):
                    issues["T1 标题位置"].append(f"p{idx} L{x} T{y} :: {t[:24]}")
                if is_body and not t:
                    issues["T2 空标题"].append(f"p{idx}")
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        if r.font.bold is not True:
                            issues["T3 标题未加粗"].append(f"p{idx} :: {t[:24]}")
                        if is_body and r.font.size and abs(r.font.size.pt - TITLE_SZ) > 0.01:
                            issues["T4 标题字号"].append(f"p{idx} {r.font.size.pt}pt :: {t[:24]}")

            # 字号 / 字体 / 颜色 / 禁用色（正文、表格单元格）
            containers = []
            if sh.has_text_frame:
                containers.append(list(sh.text_frame.paragraphs))
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        containers.append(list(cell.text_frame.paragraphs))
            for paras in containers:
                for p in paras:
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        if r.font.size:
                            stat["字号 " + str(r.font.size.pt)] += 1
                            if r.font.size.pt not in allow:
                                issues["F1 字号超三档"].append(
                                    f"p{idx} {r.font.size.pt}pt :: {r.text[:20]}")
                        lat, ea = run_font(r)
                        for f in (lat, ea):
                            if f and f != FONT:
                                issues["F2 字体不一致"].append(f"p{idx} {f} :: {r.text[:16]}")
                        c = run_color(r)
                        if c and c.upper() in BANNED:
                            issues["C1 禁用色"].append(f"p{idx} #{c} :: {r.text[:16]}")

            if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
                has_num = True
                pagenums[idx] += 1

            if sh.has_table and len(sh.table.columns) > 1:
                try:
                    cell = sh.table.rows[0].cells[0]
                    if cell.text.strip():
                        fc = cell.fill.fore_color.rgb
                        if str(fc).upper() != HDR_FILL:
                            issues["B1 表头底色"].append(f"p{idx} #{fc}")
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if r.text.strip() and (
                                        (r.font.size and r.font.size.pt != 12.0)
                                        or r.font.bold is not True):
                                    issues["B2 表头字重/字号"].append(
                                        f"p{idx} {r.font.size.pt if r.font.size else None}pt bold={r.font.bold}")
                except Exception:
                    pass

        if not has_num and layout_name not in NO_PAGENUM:
            issues["P1 缺页码"].append(f"p{idx}")
        elif pagenums[idx] > 1:
            issues["P2 页码重复"].append(f"p{idx} ×{pagenums[idx]}")

        if len(s.shapes) == 0 and layout_name not in ("1_Section Header",):
            issues["S1 空白页"].append(f"p{idx}")
        if layout_name == "自定义版式":
            issues["W1 版式越界"].append(f"p{idx} 使用了禁用的「自定义版式」（模板效果示意图）")
        elif layout_name not in LAYOUTS_OK:
            issues["W1 版式越界"].append(
                f"p{idx} 使用版式「{layout_name}」，允许：{' / '.join(LAYOUTS_OK)}")

    if unknown:
        desc = "、".join(f"{k}×{v}页" for k, v in sorted(unknown.items(), key=lambda x: -x[1]))
        issues["T0 模板不符"].append(f"版式名不在公司模板内：{desc}")

    print(f"文件：{path.split('/')[-1]}")
    print(f"页数 {len(prs.slides)}｜标题 {len(titles)}｜页码 {sum(pagenums.values())}｜"
          f"主题dk1 #{dk or '?'}")
    dist = dict(sorted(((k.split()[1], v) for k, v in stat.items()), key=lambda x: -float(x[0])))
    print(f"字号分布：{dist}")
    print("-" * 60)
    if not issues:
        print("✅ 全部通过")
    for code in sorted(issues):
        v = issues[code]
        print(f"✗ {code}：{len(v)} 处")
        if not brief:
            for d in v[:8]:
                print("    ", d)
            if len(v) > 8:
                print(f"     … 另 {len(v) - 8} 处")
    print("-" * 60)
    print(f"合计 {sum(len(v) for v in issues.values())} 处待修")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], "--brief" in sys.argv)
