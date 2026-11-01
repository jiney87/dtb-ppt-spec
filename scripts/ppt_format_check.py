# -*- coding: utf-8 -*-
"""
PPT 格式规范自检工具（DTB项目专项汇报 规范 v1.0）
用法：python3 ppt_format_check.py <file.pptx> [--brief]
输出：逐条违规清单 + 统计摘要。只读，不修改源文件。
"""
import sys, re
from collections import Counter, defaultdict
from pptx import Presentation
from pptx.util import Emu

FONT = "微软雅黑"
SIZES = {28.0, 16.0, 12.0}
TITLE_SZ = 28.0
TITLE_POS = (79.1, 50.1)      # px, 允许 ±4
PXW = 1280.0
BANNED = {"FF0000", "C9C8C7", "FEFFFE", "4EA5FF", "61BA46",
          "008BCE", "0070C0", "00B0F0", "2F75B5", "B5C6EA"}
HDR_FILL = "0084D5"
LAYOUTS_OK = ("封面", "间隔页", "正文页1", "图片页")


def px(v, scale):
    return None if v is None else round(v * scale, 1)


def run_font(r):
    """取 run 的中文/西文字体，缺省继承 None"""
    rPr = r._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    if rPr is None:
        return None, None
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    lat = rPr.find(ns + 'latin')
    ea = rPr.find(ns + 'ea')
    return (lat.get('typeface') if lat is not None else None,
            ea.get('typeface') if ea is not None else None)


def run_color(r):
    rPr = r._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    if rPr is None:
        return None
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    sf = rPr.find(ns + 'solidFill')
    if sf is None:
        return None
    c = sf.find(ns + 'srgbClr')
    return c.get('val') if c is not None else None


def main(path, brief=False):
    prs = Presentation(path)
    scale = PXW / prs.slide_width
    issues = defaultdict(list)   # code -> [detail]
    stat = Counter()

    titles = []
    pagenums = Counter()

    for idx, s in enumerate(prs.slides, start=1):
        has_num = False
        layout_name = s.slide_layout.name
        for sh in s.shapes:
            if sh.is_placeholder and str(sh.placeholder_format.type).startswith('TITLE'):
                t = sh.text_frame.text.strip()
                x, y = px(sh.left, scale), px(sh.top, scale)
                titles.append((idx, x, y, t))
                # 目录页 / 章节间页的标题为版式内合法居中位，不按正文页标题位校验
                is_body = layout_name in ('正文页1', '图片页', '1_Section Header', '正文页2')
                if is_body and (abs((y or 0) - TITLE_POS[1]) > 4 or abs((x or 0) - TITLE_POS[0]) > 4):
                    issues['T1 标题位置'].append(f"p{idx} L{x} T{y} :: {t[:24]}")
                if not t:
                    issues['T2 空标题'].append(f"p{idx}")
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text.strip() and r.font.bold is not True:
                            issues['T3 标题未加粗'].append(f"p{idx} :: {t[:24]}")
                        if is_body and r.text.strip() and r.font.size and abs(r.font.size.pt - TITLE_SZ) > 0.01:
                            issues['T4 标题字号'].append(f"p{idx} {r.font.size.pt}pt :: {t[:24]}")

            # 字号 / 字体 / 颜色 / 禁用色
            containers = []
            if sh.has_text_frame:
                containers.append([p for p in sh.text_frame.paragraphs])
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
                            stat['字号 ' + str(r.font.size.pt)] += 1
                            if r.font.size.pt not in SIZES:
                                issues['F1 字号超三档'].append(f"p{idx} {r.font.size.pt}pt :: {r.text[:20]}")
                        lat, ea = run_font(r)
                        for f in (lat, ea):
                            if f and f != FONT:
                                issues['F2 字体不一致'].append(f"p{idx} {f} :: {r.text[:16]}")
                        c = run_color(r)
                        if c and c.upper() in BANNED:
                            issues['C1 禁用色'].append(f"p{idx} #{c} :: {r.text[:16]}")

            # 页码
            if sh.is_placeholder and 'SLIDE_NUMBER' in str(sh.placeholder_format.type):
                has_num = True
                pagenums[idx] += 1

            # 表格表头
            if sh.has_table and len(sh.table.columns) > 1:
                try:
                    cell = sh.table.rows[0].cells[0]
                    if cell.text.strip():
                        fc = cell.fill.fore_color.rgb
                        if str(fc).upper() != HDR_FILL:
                            issues['B1 表头底色'].append(f"p{idx} #{fc}")
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if r.text.strip() and (r.font.size and r.font.size.pt != 12.0 or r.font.bold is not True):
                                    issues['B2 表头字重/字号'].append(f"p{idx} {r.font.size.pt if r.font.size else None}pt bold={r.font.bold}")
                except Exception:
                    pass

        if not has_num and layout_name != '封面':
            issues['P1 缺页码'].append(f"p{idx}")
        elif pagenums[idx] > 1:
            issues['P2 页码重复'].append(f"p{idx} ×{pagenums[idx]}")

        if len(s.shapes) == 0:
            issues['S1 空白页'].append(f"p{idx}")
        if layout_name not in LAYOUTS_OK:
            issues['W1 版式越界'].append(f"p{idx} 使用版式「{layout_name}」，规范只允许 {', '.join(LAYOUTS_OK)}")

    total_title = len(titles)
    print(f"文件：{path.split('/')[-1]}")
    print(f"页数 {len(prs.slides)}｜标题 {total_title}｜页码 {sum(pagenums.values())}")
    print(f"字号分布：{dict(sorted(((k.split()[1], v) for k, v in stat.items()), key=lambda x: -int(float(x[0]))))}")
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


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], '--brief' in sys.argv)
