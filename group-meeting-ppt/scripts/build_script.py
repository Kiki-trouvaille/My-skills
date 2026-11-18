#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
汇报讲稿生成器
==============
读取与 build_deck.py 同一份 deck.json，输出：
    1) .docx 逐页讲稿（四步框架：Why / How / Read / So What）
    2) .md   同内容纯文本，便于快速改稿
    3) 把讲稿写入 pptx 的演讲者备注（build_deck.py 已做，这里可单独补写）

用法:
    python build_script.py deck.json -o 汇报讲稿.docx
"""

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "微软雅黑"
PRIMARY = RGBColor(0x08, 0x4B, 0x82)
TEXT = RGBColor(0x1B, 0x2F, 0x41)
GRAY = RGBColor(0x60, 0x78, 0x8A)
ORANGE = RGBColor(0xD8, 0x78, 0x38)

STEPS = [
    ("why", "Why 动机引入", PRIMARY),
    ("how", "How 具体做法", PRIMARY),
    ("read", "Read 图表精读", PRIMARY),
    ("so_what", "So What 总结与边界", ORANGE),
]


def set_font(run, size=11, bold=False, color=TEXT, italic=False, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)


def para(doc, text="", size=11, bold=False, color=TEXT, italic=False,
         before=0, after=6, line=1.55, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.first_line_indent = Cm(indent)
    p.alignment = align
    if text:
        # 行内标记：**加粗** / !!标红!!
        for seg in re.split(r"(\*\*.+?\*\*|!!.+?!!)", text):
            if not seg:
                continue
            if seg.startswith("**") and seg.endswith("**") and len(seg) > 4:
                r = p.add_run(seg[2:-2]); set_font(r, size, True, color, italic)
            elif seg.startswith("!!") and seg.endswith("!!") and len(seg) > 4:
                r = p.add_run(seg[2:-2]); set_font(r, size, True, RGBColor(0xC0, 0x39, 0x2B), italic)
            else:
                r = p.add_run(seg); set_font(r, size, bold, color, italic)
    return p


def rule(doc, color="C7DCE9", size=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def shade(p, color="F6FAFD"):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    pPr.append(shd)


def add_page_footer(section, label):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label + "　·　第 ")
    set_font(r, 8.5, False, GRAY)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "17")
    col = OxmlElement("w:color"); col.set(qn("w:val"), "60788A")
    rPr.append(sz); rPr.append(col)
    run.append(rPr)
    t = OxmlElement("w:t"); t.text = "1"
    run.append(t)
    fld.append(run)
    p._p.append(fld)
    r2 = p.add_run(" 页")
    set_font(r2, 8.5, False, GRAY)


def cn_len(s):
    return len(re.sub(r"[\s\*!]", "", s or ""))


def build(spec, out_docx):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.4)
    sec.right_margin = Cm(2.4)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(11)
    try:
        style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    except Exception:
        pass

    week = spec.get("week", "")
    cover = spec.get("cover", {})
    slides = [s for s in spec.get("slides", []) if s.get("type") != "closing"]
    total_chars = 0
    for s in slides:
        sp = s.get("speech") or {}
        for k, _, _ in STEPS:
            total_chars += cn_len(sp.get(k))
    minutes = max(1, round(total_chars / 230))

    # 标题区
    p = para(doc, "%s组会 · 汇报讲稿" % week, 20, True, PRIMARY, after=2, align=WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.line_spacing = 1.2
    para(doc, cover.get("title", "").replace("\n", " "), 12.5, True, TEXT, after=2, line=1.35, align=WD_ALIGN_PARAGRAPH.LEFT)
    para(doc, "汇报人：%s　｜　%s　｜　共 %d 页　｜　预计 %d 分钟"
         % (spec.get("author", ""), spec.get("date", ""), len(slides) + 1, minutes),
         9.5, False, GRAY, after=0, align=WD_ALIGN_PARAGRAPH.LEFT)
    rule(doc, "084B82", 12)

    # 使用说明
    p = para(doc, "朗读提示：每页按 Why → How → Read → So What 四步推进；"
                  "方括号内是给自己看的路标，不必念出。斜体句为翻页过渡。",
             9, False, GRAY, after=10, line=1.4, italic=True)

    order = 0
    for sl in spec.get("slides", []):
        t = sl.get("type")
        if t == "closing":
            continue
        order += 1
        title = sl.get("title") or sl.get("module") or sl.get("part") or ""
        if t == "part":
            head = "【第 %d 页：%s %s】" % (order + 1, sl.get("part", ""), sl.get("module", ""))
        else:
            head = "【第 %d 页：%s】" % (order + 1, title)

        p = para(doc, "", before=12, after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
        shade(p, "EDF5FA")
        r = p.add_run(head)
        set_font(r, 12.5, True, PRIMARY)

        sp = sl.get("speech") or {}
        if sl.get("lead"):
            para(doc, "页面导语：" + sl["lead"], 9, False, GRAY, after=6, line=1.4, italic=True)

        wrote = False
        for key, label, color in STEPS:
            txt = sp.get(key)
            if not txt:
                continue
            wrote = True
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.3
            r = p.add_run("【%s】" % label)
            set_font(r, 9, True, color)
            para(doc, txt, 11, False, TEXT, after=6, line=1.6)

        if not wrote:
            para(doc, "（本页为过渡页，口头一句带过：接下来进入%s。）" % (sl.get("module") or title),
                 9, False, GRAY, italic=True, after=6)

        if sp.get("transition"):
            para(doc, "→ " + sp["transition"], 9.5, False, GRAY, italic=True, after=4, line=1.4)

    # 结尾
    rule(doc, "084B82", 8)
    para(doc, "结束语：本周汇报到此，请老师和同学批评指正。", 11, True, PRIMARY, after=4)
    para(doc, "全文约 %d 字，按每分钟 230 字估算约 %d 分钟。" % (total_chars, minutes),
         9, False, GRAY, after=0)

    add_page_footer(sec, "%s组会汇报讲稿" % week)
    os.makedirs(os.path.dirname(os.path.abspath(out_docx)), exist_ok=True)
    doc.save(out_docx)

    # 同步输出 md
    md_path = os.path.splitext(out_docx)[0] + ".md"
    lines = ["# %s组会 · 汇报讲稿" % week, "",
             "> %s ｜ 汇报人：%s ｜ %s ｜ 预计 %d 分钟"
             % (cover.get("title", "").replace("\n", " "), spec.get("author", ""),
                spec.get("date", ""), minutes), ""]
    order = 0
    for sl in spec.get("slides", []):
        if sl.get("type") == "closing":
            continue
        order += 1
        title = sl.get("title") or sl.get("module") or ""
        if sl.get("type") == "part":
            lines.append("## 【第 %d 页：%s %s】" % (order + 1, sl.get("part", ""), sl.get("module", "")))
        else:
            lines.append("## 【第 %d 页：%s】" % (order + 1, title))
        sp = sl.get("speech") or {}
        for key, label, _ in STEPS:
            if sp.get(key):
                lines.append("**%s**  " % label)
                lines.append(sp[key])
                lines.append("")
        if sp.get("transition"):
            lines.append("*→ %s*" % sp["transition"])
            lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("**结束语**：本周汇报到此，请老师和同学批评指正。")
    lines.append("")
    lines.append("*全文约 %d 字，按每分钟 230 字估算约 %d 分钟。*" % (total_chars, minutes))
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_docx, md_path, minutes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default=None)
    args = ap.parse_args()
    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    out = args.out or os.path.splitext(args.spec)[0] + "_讲稿.docx"
    d, m, mins = build(spec, out)
    print("[输出] %s\n[输出] %s\n[时长] 约 %d 分钟" % (d, m, mins))


if __name__ == "__main__":
    sys.exit(main())
