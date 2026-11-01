# -*- coding: utf-8 -*-
"""
按《PPT 输出规范 v2.0.1》（默认模板 公司模板.pptx）批量调整 PPTX。

用法：
  python3 apply_format.py 源文件.pptx                  # 干跑，只输出将做的修改
  python3 apply_format.py 源文件.pptx --apply           # 输出 <源名>_规范版.pptx
  python3 apply_format.py 源文件.pptx --apply -o 目标.pptx
  可选项：--only font,size,color,title,table,pagenum
          --zebra   表格隔行填充 #F4F4F4
          --force   允许覆盖已存在的输出文件

原则：不修改源文件；只做规范内的确定性替换；不确定的一律报告，不猜。
"""
import sys, os, argparse
from copy import deepcopy
from collections import defaultdict

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from lxml import etree

# ─────────────────────────── 规范常量（与 spec 2.0.1 对齐；几何取自公司模板原值）
EA_FONT = "微软雅黑"
LATIN_FONT = "Verdana"
SZ_TITLE, SZ_BODY, SZ_LABEL, SZ_FOOTER = 24.0, 14.0, 9.0, 6.5
BLUE, DBLUE, LBLUE = "0084D5", "005E99", "DDEBF7"
BLACK, GRAY, MGRAY, LGRAY = "000000", "7F7F7F", "BFBFBF", "F4F4F4"
TITLE_BOX = (66.0, 50.0, 1030.0, 52.0)
PAGENUM_BOX = (917.0, 668.0, 288.0, 38.0)
LAYOUT_BODY = ("正文页1", "正文页2", "图片页")

COLOR_MAP = {
    "FF0000": "C00000", "C9C8C7": "BFBFBF", "FEFFFE": "FFFFFF",
    "4EA5FF": "0084D5", "61BA46": "78BE20", "008BCE": "0084D5",
    "0070C0": "0084D5", "00B0F0": "0084D5", "2F75B5": "005E99",
    "B5C6EA": "DDEBF7",
}
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PX = 9525  # 1 px = 9525 EMU (1280px = 12192000EMU)


def to_emu(px):
    return Emu(int(round(px * PX)))


def map_size(pt):
    """保留合法层级；无法仅凭字号确定语义时留待人工处理。"""
    if pt is None:
        return None, False
    if pt in {6.5, 9, 10, 14, 16, 18, 24, 28, 38.5}:
        return pt, False
    return None, False


# ─────────────────────────── XML 工具
RPR_TAIL = ("sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")


def set_ea_font(rPr, ea_name=EA_FONT, latin_name=LATIN_FONT):
    """在 rPr 上补齐 latin/ea/cs，且保持 schema 顺序（放在 fill 之后、hlinkClick 之前）"""
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn("a:" + tag))
        if el is None:
            el = etree.SubElement(rPr, qn("a:" + tag))
        el.set("typeface", latin_name if tag == "latin" else ea_name)
        el.set("panose", "020B0503020204020204")
        el.set("charset", "-122")
    # 重排：把三个字体节点移到第一个尾部节点之前
    fonts = [rPr.find(qn("a:" + t)) for t in ("latin", "ea", "cs")]
    anchor = None
    for child in rPr:
        if etree.QName(child).localname in RPR_TAIL:
            anchor = child
            break
    for f in fonts:
        rPr.remove(f)
    if anchor is not None:
        for f in fonts:
            anchor.addprevious(f)
    else:
        for f in fonts:
            rPr.append(f)


RANK = {'ln':1,'noFill':2,'solidFill':2,'gradFill':2,'blipFill':2,'pattFill':2,'grpFill':2,
        'effectLst':3,'effectDag':3,'highlight':4,'uLnTx':5,'uLn':5,'uFillTx':6,'uFill':6,
        'latin':7,'ea':8,'cs':9,'sym':10,'hlinkClick':11,'hlinkMouseOver':12,'rtl':13,'extLst':14}


def set_solid(rPr, hexcolor):
    """替换 rPr 的纯色填充，并按 CT_TextCharacterProperties 顺序插入"""
    ns = A
    old = rPr.find(ns + "solidFill")
    if old is not None:
        rPr.remove(old)
    sf = etree.Element(ns + "solidFill")
    c = etree.SubElement(sf, ns + "srgbClr")
    c.set("val", hexcolor)
    anchor = None
    for child in rPr:
        if RANK.get(etree.QName(child).localname, 0) > 2:
            anchor = child
            break
    if anchor is not None:
        anchor.addprevious(sf)
    else:
        rPr.append(sf)


def run_rgb(r):
    rPr = r._r.find(qn("a:rPr"))
    if rPr is None:
        return None
    c = rPr.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
    return c.get("val") if c is not None else None


# ─────────────────────────── 主流程
def walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(sh.shapes)


class Applier:
    def __init__(self, prs, opts):
        self.prs = prs
        self.opts = opts
        self.log = defaultdict(list)
        self.up = defaultdict(int)   # 放大字号的页统计

    def run(self):
        for i, slide in enumerate(self.prs.slides, start=1):
            lay = slide.slide_layout.name
            is_body = lay in LAYOUT_BODY
            title_id = None
            title_shape = None
            try:
                title_shape = slide.shapes.title
                title_id = title_shape.shape_id if title_shape is not None else None
            except Exception:
                pass
            for sh in walk(slide.shapes):
                if sh.has_text_frame:
                    is_title = (title_id is not None and sh.shape_id == title_id)
                    self.fix_text(sh, i, is_title, is_body)
                if sh.has_table and "table" in self.opts.only:
                    self.fix_table(sh, i)
            if "title" in self.opts.only and is_body and title_shape is not None:
                self.fix_title_box(title_shape, i)
            if "pagenum" in self.opts.only and lay != "封面":
                self.ensure_pagenum(slide, i)

    # --- 文本
    def fix_text(self, sh, page, is_title, is_body):
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                rPr = r._r.get_or_add_rPr()

                # 字号
                if "size" in self.opts.only:
                    cur = r.font.size.pt if r.font.size else None
                    new, grew = map_size(cur)
                    if is_title:
                        new = 38.5 if sh.part.slide.slide_layout.name == "间隔页" else SZ_TITLE
                        grew = cur is not None and new > cur
                    if new is None:
                        if is_title:
                            new = 38.5 if sh.part.slide.slide_layout.name == "间隔页" else SZ_TITLE
                        else:
                            self.log["字号待确认"].append(f"p{page} {cur or '继承'}pt :: {r.text[:14]}")
                    if new is not None and cur != new:
                        r.font.size = Pt(new)
                        self.log["字号"].append(f"p{page} {cur or '继承'}→{new:g}pt :: {r.text[:14]}")
                        if grew:
                            self.up[page] += 1

                # 字体
                if "font" in self.opts.only:
                    set_ea_font(rPr)

                # 加粗：标题加粗，正文取消（表头另处理）
                if is_title and r.font.bold is not True:
                    r.font.bold = True

                # 颜色
                if "color" in self.opts.only:
                    cur = run_rgb(r)
                    if cur and cur.upper() in COLOR_MAP:
                        set_solid(rPr, COLOR_MAP[cur.upper()])
                        self.log["颜色"].append(f"p{page} #{cur}→#{COLOR_MAP[cur.upper()]} :: {r.text[:14]}")

            # 标题统一品牌蓝
            if is_title and "color" in self.opts.only:
                for r in p.runs:
                    if r.text.strip():
                        set_solid(r._r.get_or_add_rPr(), BLUE)

    def fix_title_box(self, sh, page):
        l, t, w, h = TITLE_BOX
        if (sh.left, sh.top, sh.width, sh.height) != (to_emu(l), to_emu(t), to_emu(w), to_emu(h)):
            sh.left, sh.top, sh.width, sh.height = to_emu(l), to_emu(t), to_emu(w), to_emu(h)
            self.log["标题位置"].append(f"p{page} 归位至 L{l}/T{t}")

    # --- 表格
    def fix_table(self, sh, page):
        tbl = sh.table
        ncol = len(tbl.columns)
        if ncol == 1:      # 单列表格实为内容框，不染色，只统一字号
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            if r.text.strip():
                                r.font.size = Pt(SZ_BODY)
                                set_ea_font(r._r.get_or_add_rPr())
            self.log["表格"].append(f"p{page} 1×{len(tbl.rows)} 单列表，仅统一字号")
            return
        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                row.height = to_emu(36)
            else:
                row.height = to_emu(28)
            for cell in row.cells:
                has_text = any(p.runs for p in cell.text_frame.paragraphs)
                if not has_text:
                    continue
                if ri == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(BLUE)
                elif self.opts.zebra and ri % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(LGRAY)
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        r.font.size = Pt(SZ_BODY)
                        r.font.name = EA_FONT
                        set_ea_font(r._r.get_or_add_rPr())
                        if ri == 0:
                            r.font.bold = True
                            set_solid(r._r.get_or_add_rPr(), "FFFFFF")
                            p.alignment = PP_ALIGN.CENTER
                        else:
                            if r.font.bold is True and len(r.text.strip()) > 4:
                                r.font.bold = False   # 数据行误加粗的清掉
                            if "color" in self.opts.only:
                                cur = run_rgb(r)
                                if cur and cur.upper() in COLOR_MAP:
                                    set_solid(r._r.get_or_add_rPr(), COLOR_MAP[cur.upper()])
                                    self.log["颜色"].append(
                                        f"p{page} #{cur}→#{COLOR_MAP[cur.upper()]} :: {r.text[:14]}")
        self.log["表格"].append(f"p{page} {len(tbl.rows)}×{len(tbl.columns)} 表头重设")

    # --- 页码
    def ensure_pagenum(self, slide, page):
        target = None
        for sh in slide.placeholders:
            if sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                target = sh
                break
        if target is None:
            src = None
            for sh in slide.slide_layout.placeholders:
                if sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                    src = sh
                    break
            if src is None:
                self.log["页码未补"].append(f"p{page} 版式无页码占位符")
                return
            el = deepcopy(src._element)
            ids = [int(e.get("id")) for e in slide.shapes._spTree.iter(qn("p:cNvPr")) if e.get("id")]
            nid = max(ids or [1]) + 1
            nv = el.find(qn("p:nvSpPr") + "/" + qn("p:cNvPr"))
            if nv is not None:
                nv.set("id", str(nid))
            slide.shapes._spTree.append(el)
            self.log["页码补齐"].append(f"p{page} 新增页码占位符")
            target = slide.shapes[-1]

        l, t, w, h = PAGENUM_BOX
        target.left, target.top, target.width, target.height = to_emu(l), to_emu(t), to_emu(w), to_emu(h)

        tf = target.text_frame
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.RIGHT
            for r in p.runs:
                r.font.size = Pt(SZ_FOOTER)
                r.font.bold = False
                set_ea_font(r._r.get_or_add_rPr())
                set_solid(r._r.get_or_add_rPr(), GRAY)
        # 占位符可能是 fld 而非 run
        for fld in tf._txBody.iter(qn("a:fld")):
            rPr = fld.find(qn("a:rPr"))
            if rPr is None:
                rPr = etree.SubElement(fld, qn("a:rPr"))
            rPr.set("sz", "650")
            rPr.set("b", "0")
            set_ea_font(rPr)
            set_solid(rPr, GRAY)

    # --- 报告
    def report(self):
        tot = sum(len(v) for v in self.log.values())
        print("=" * 62)
        for k in ("字号", "字体", "颜色", "标题位置", "表格", "页码补齐", "页码未补", "字号待确认"):
            if k in self.log:
                print(f"{k}：{len(self.log[k])} 处")
        print("=" * 62)
        if self.opts.verbose:
            for k, v in self.log.items():
                print(f"\n【{k}】")
                for line in v[:40]:
                    print("  ", line)
                if len(v) > 40:
                    print(f"   … 另 {len(v)-40} 处")
        if self.up:
            risk = sorted(self.up.items(), key=lambda x: -x[1])[:6]
            print("\n⚠ 字号被放大的页面（需人工复核是否溢出）：")
            for pg, n in risk:
                print(f"   p{pg}：{n} 处")
        print(f"\n合计 {tot} 处调整")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("-o", "--out")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--only", default="font,size,color,title,table,pagenum")
    ap.add_argument("--zebra", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    o = ap.parse_args()
    o.only = set(o.only.split(","))

    prs = Presentation(o.src)
    applier = Applier(prs, o)
    applier.run()
    applier.report()

    if not o.apply:
        print("\n[干跑] 未写入任何文件。确认无误后加 --apply 生成规范版。")
        return
    out = o.out or os.path.splitext(o.src)[0] + "_规范版.pptx"
    if os.path.abspath(out) == os.path.abspath(o.src):
        sys.exit("输出路径与源文件相同，已中止")
    if os.path.exists(out) and not o.force:
        sys.exit(f"输出已存在：{out}（加 --force 覆盖）")
    prs.save(out)
    print(f"\n已写出：{out}")


if __name__ == "__main__":
    main()
