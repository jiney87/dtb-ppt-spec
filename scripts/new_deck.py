#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从公司模板生成一份新 PPT 骨架（默认模板 = assets/company-template.pptx）。

用法：
  python3 new_deck.py 输出.pptx --title "2027年DTB项目战略规划" --subtitle "汇报版 V1"
  python3 new_deck.py 输出.pptx --title "季度复盘" --sections "一、经营回顾|二、问题诊断|三、下季动作" --body 2
  python3 new_deck.py 输出.pptx --no-toc --no-thanks

产出结构：封面 → 目录 →（章节页 + N 张正文页）× 章节数 → 结语页
所有文字默认套用规范：微软雅黑、标题 28pt 加粗 #0084D5、正文 12pt、页码 12pt #7F7F7F。

原则：**只复制模板，不新建空白演示文稿**；版式只用模板自带的 6 个。
"""
import argparse
import copy
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

FONT = "微软雅黑"
BLUE = RGBColor(0x00, 0x84, 0xD5)
BLACK = RGBColor(0x00, 0x00, 0x00)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)

RANK = {"ln": 1, "noFill": 2, "solidFill": 2, "gradFill": 2, "blipFill": 2, "pattFill": 2,
        "grpFill": 2, "effectLst": 3, "effectDag": 3, "highlight": 4, "uLnTx": 5, "uLn": 5,
        "uFillTx": 6, "uFill": 6, "latin": 7, "ea": 8, "cs": 9, "sym": 10,
        "hlinkClick": 11, "hlinkMouseOver": 12, "rtl": 13, "extLst": 14}


def order_rpr(rPr):
    """按 CT_TextCharacterProperties schema 顺序重排子元素（lxml 移动节点即重排）"""
    kids = list(rPr)
    kids.sort(key=lambda c: RANK.get(etree.QName(c).localname, 50))
    for c in kids:
        rPr.append(c)


def set_font(run, size=None, bold=None, color=None, font=FONT):
    f = run.font
    f.name = font
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if color is not None:
        f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(A + "latin")
    if latin is None:
        latin = etree.SubElement(rPr, A + "latin")
    latin.set("typeface", font)
    ea = rPr.find(A + "ea")
    if ea is None:
        ea = etree.SubElement(rPr, A + "ea")
    ea.set("typeface", font)
    order_rpr(rPr)


def write(tf, lines, size=12, bold=False, color=BLACK, align=PP_ALIGN.LEFT, space_after=6):
    """清空并写入若干段落"""
    tf.clear()
    tf.word_wrap = True
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = line
        set_font(run, size, bold, color)
    return tf


def style_slide_number(sp_el):
    """把版式带来的页码字段统一成 微软雅黑 12pt #7F7F7F"""
    for fld in sp_el.iter(A + "fld"):
        rPr = fld.find(A + "rPr")
        if rPr is None:
            rPr = etree.Element(A + "rPr")
            fld.insert(0, rPr)
        rPr.set("sz", "1200")
        rPr.set("b", "0")
        for c in list(rPr):
            rPr.remove(c)
        sf = etree.SubElement(rPr, A + "solidFill")
        etree.SubElement(sf, A + "srgbClr").set("val", "7F7F7F")
        etree.SubElement(rPr, A + "latin").set("typeface", FONT)
        etree.SubElement(rPr, A + "ea").set("typeface", FONT)
        order_rpr(rPr)


def add_page_number(slide, layout):
    """模板 slide 默认不带页码占位符：从版式深拷贝一个"""
    for sh in slide.shapes:
        if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
            return
    for sh in layout.shapes:
        if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
            el = copy.deepcopy(sh._element)   # 深拷贝出来是游离节点，直接挂到 slide 上
            slide.shapes._spTree.append(el)
            style_slide_number(el)
            return


def placeholders(slide):
    return [sh for sh in slide.shapes if sh.is_placeholder]


def title_of(slide):
    for sh in placeholders(slide):
        if str(sh.placeholder_format.type).startswith("TITLE"):
            return sh
    return None


def body_of(slide):
    for sh in placeholders(slide):
        t = str(sh.placeholder_format.type)
        if t.startswith(("BODY", "OBJECT")):
            return sh
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--title", default="【项目名】汇报")
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--sections", default="", help='章节用 | 分隔，如 "一、经营回顾|二、问题诊断"')
    ap.add_argument("--body", type=int, default=1, help="每章正文页数（默认 1）")
    ap.add_argument("--no-toc", action="store_true")
    ap.add_argument("--no-thanks", action="store_true")
    ap.add_argument("--template", default="")
    o = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    tpl = o.template or os.path.join(os.path.dirname(here), "assets", "company-template.pptx")
    if not os.path.exists(tpl):
        sys.exit(f"✗ 找不到模板：{tpl}\n  先跑 sync_rules.py --update 拉取 assets/company-template.pptx")
    if os.path.exists(o.out):
        sys.exit(f"✗ 目标已存在，未覆盖：{o.out}")

    prs = Presentation(tpl)
    L = {l.name: l for l in prs.slide_layouts}
    for need in ("封面", "间隔页", "正文页1", "1_Section Header"):
        if need not in L:
            sys.exit(f"✗ 模板缺版式「{need}」，不是公司模板")

    # 清空模板自带的 7 张示例页
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        prs.part.drop_rel(sld.get(R + "id"))
        lst.remove(sld)

    pages = []

    # 1) 封面
    s = prs.slides.add_slide(L["封面"])
    ph = [p for p in placeholders(s) if str(p.placeholder_format.type).startswith("BODY")]
    ph.sort(key=lambda p: (p.width or 0), reverse=True)   # 宽的是主标题
    if ph:
        write(ph[0].text_frame, o.title.split("\\n"), 48, True, BLUE, space_after=0)
    if len(ph) > 1 and o.subtitle:
        write(ph[1].text_frame, o.subtitle, 15, True, BLACK, space_after=0)
    if len(ph) > 2 and o.date:
        write(ph[2].text_frame, o.date, 12, False, GRAY, space_after=0)
    pages.append("封面")

    sections = [x.strip() for x in o.sections.split("|") if x.strip()]

    # 2) 目录
    if not o.no_toc:
        s = prs.slides.add_slide(L["正文页1"])
        write(title_of(s).text_frame, "目录", 28, True, BLUE, space_after=0)
        items = sections or ["第一部分"]
        b = body_of(s)
        if b is not None:
            tf = b.text_frame
            tf.clear()
            tf.word_wrap = True
            for i, sec in enumerate(items, 1):
                p_ = tf.paragraphs[0] if i == 1 else tf.add_paragraph()
                p_.space_after = Pt(18)
                r1 = p_.add_run(); r1.text = f"{i:02d}"
                set_font(r1, 15, True, BLUE)
                r2 = p_.add_run(); r2.text = "    " + sec
                set_font(r2, 15, False, BLACK)
        add_page_number(s, L["正文页1"])
        pages.append("目录")

    # 3) 章节页 + 正文页
    for sec in sections:
        s = prs.slides.add_slide(L["间隔页"])
        t = title_of(s)
        if t is not None:
            write(t.text_frame, sec, 32, True, BLUE, align=PP_ALIGN.CENTER, space_after=0)
        add_page_number(s, L["间隔页"])
        pages.append(f"间隔页·{sec}")
        for _ in range(max(1, o.body)):
            s = prs.slides.add_slide(L["正文页1"])
            write(title_of(s).text_frame, "【页标题】写结论，≤26字", 28, True, BLUE, space_after=0)
            add_page_number(s, L["正文页1"])
            pages.append("正文页")

    if not sections:
        s = prs.slides.add_slide(L["正文页1"])
        write(title_of(s).text_frame, "【页标题】写结论，≤26字", 28, True, BLUE, space_after=0)
        add_page_number(s, L["正文页1"])
        pages.append("正文页")

    # 4) 结语页
    if not o.no_thanks:
        prs.slides.add_slide(L["1_Section Header"])
        pages.append("结语页")

    os.makedirs(os.path.dirname(os.path.abspath(o.out)) or ".", exist_ok=True)
    prs.save(o.out)

    print(f"✓ 已生成 {o.out}")
    print(f"  模板 {os.path.basename(tpl)}｜{len(pages)} 页")
    print("  结构：" + " → ".join(pages))
    print("  下一步：替换「【页标题】写结论」提示文字；内容守版心 55–1206 / 135–655；"
          "完成后跑 ppt_format_check.py 自检")


if __name__ == "__main__":
    main()
