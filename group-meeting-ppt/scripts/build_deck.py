#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
组会汇报 PPT 渲染器
====================
把 deck.json（内容规格）渲染成固定版式的 .pptx。
所有几何、字号、配色都写死在本文件里，保证每一周的组会 PPT 视觉完全一致。

用法:
    python build_deck.py deck.json -o 输出.pptx
    python build_deck.py deck.json -o 输出.pptx --check      # 只做内容规则自检

设计基准（与原稿一致）:
    幻灯片 33.87 x 19.05 cm (13.333 x 7.5 in, 16:9)
    左边距 1.27cm，内容宽 30.9cm
"""

import argparse
import json
import math
import os
import re
import sys

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ----------------------------------------------------------------------------
# 设计令牌 (Design tokens)
# ----------------------------------------------------------------------------
FONT = "微软雅黑"
MONO = "Consolas"

PRIMARY = "084B82"   # 主色 深蓝
BLUE2 = "176FA5"     # 次级蓝
TEAL = "27877B"      # 青绿（强调 / 正向）
ORANGE = "D87838"    # 橙（注意 / 局限）
SKY = "49A6C7"       # 浅蓝
TEXT = "1B2F41"      # 正文深灰蓝
GRAY = "60788A"      # 注释灰
GRAY2 = "94A3B8"     # 更浅的灰
RED = "C0392B"       # 标红（风险 / 反例）
BAND = "EDF5FA"      # 导语带
CARD = "F6FAFD"      # 卡片底
LINE = "C7DCE9"      # 卡片描边
WHITE = "FFFFFF"

ACCENTS = [PRIMARY, BLUE2, TEAL, ORANGE, SKY]

W, H = 33.87, 19.05
M = 1.27
CW = 30.9
BODY_TOP = 4.45      # 有导语时内容区起点
BODY_TOP_NOLEAD = 2.85
BODY_BOTTOM = 17.30  # 页脚线之上
FOOTER_Y = 17.96

CHROME_BOTTOM = 18.71


# ----------------------------------------------------------------------------
# 基础工具
# ----------------------------------------------------------------------------
def rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _units(text):
    """估算文本宽度（以字号为 1 单位）：中日韩 1.0，其余 0.56。"""
    u = 0.0
    for ch in text:
        if ord(ch) > 127:
            u += 1.0
        elif ch in " .,:;'|!":
            u += 0.32
        elif ch.isupper():
            u += 0.62
        else:
            u += 0.55
    return u


def _estimate(paras, width_cm, scale):
    """估算段落组在给定宽度下的总高度（cm）。"""
    total = 0.0
    usable = max(1.0, width_cm)
    for p in paras:
        size = max(6.0, p.get("size", 14) * scale)
        ls = p.get("ls", 1.3)
        cpl = max(1.0, usable / (size * 0.03528))
        txt = p.get("text", "")
        lines = 0
        for seg in str(txt).split("\n"):
            lines += max(1, math.ceil(_units(seg) / cpl))
        total += lines * size * 0.03528 * ls * 1.02
        total += p.get("space_after", 0) * 0.03528
        total += p.get("space_before", 0) * 0.03528
    return total


TOKEN_RE = re.compile(r"(\*\*.+?\*\*|!!.+?!!)")


def _set_run_font(run, size, color, bold, font=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)


def _add_runs(p, text, size, color, bold=False, font=FONT, italic=False):
    """支持行内标记：**加粗**  !!标红加粗!!"""
    for seg in TOKEN_RE.split(str(text)):
        if not seg:
            continue
        run = p.add_run()
        if seg.startswith("**") and seg.endswith("**") and len(seg) > 4:
            run.text = seg[2:-2]
            _set_run_font(run, size, color, True, font, italic)
        elif seg.startswith("!!") and seg.endswith("!!") and len(seg) > 4:
            run.text = seg[2:-2]
            _set_run_font(run, size, RED, True, font, italic)
        else:
            run.text = seg
            _set_run_font(run, size, color, bold, font, italic)


def textbox(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP,
            align=PP_ALIGN.LEFT, pad=0.0, autofit=True, min_scale=0.66,
            wrap=True):
    box = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Cm(pad)
    tf.margin_right = Cm(pad)
    tf.margin_top = Cm(0)
    tf.margin_bottom = Cm(0)
    tf.vertical_anchor = anchor

    scale = 1.0
    if autofit:
        while scale > min_scale and _estimate(paras, w - 2 * pad, scale) > h:
            scale -= 0.035

    first = True
    for spec in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = spec.get("align", align)
        # 用精确磅值锁定行高，避免 PowerPoint 的倍数行距（≈1.32em）把版式撑开
        _sz = max(6.0, spec.get("size", 14) * scale)
        p.line_spacing = Pt(_sz * spec.get("ls", 1.3))
        if spec.get("space_before"):
            p.space_before = Pt(spec["space_before"] * scale)
        p.space_after = Pt(spec.get("space_after", 0) * scale)
        _add_runs(p, spec.get("text", ""), max(6.0, spec.get("size", 14) * scale),
                  spec.get("color", TEXT), spec.get("bold", False),
                  spec.get("font", FONT), spec.get("italic", False))
    return box


def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=0.75,
          adj=None, shadow=False):
    sp = slide.shapes.add_shape(kind, Cm(x), Cm(y), Cm(w), Cm(h))
    if adj is not None:
        try:
            sp.adjustments[0] = adj
        except Exception:
            pass
    if fill:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line)
        sp.line.width = Pt(line_w)
    else:
        sp.line.fill.background()
    if not shadow:
        try:
            sp.shadow.inherit = False
        except Exception:
            pass
    sp.text_frame.word_wrap = True
    return sp


def rect(slide, x, y, w, h, fill, line=None):
    return shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=fill, line=line)


def card(slide, x, y, w, h, accent=PRIMARY, fill=CARD):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, line=LINE, adj=0.035)
    if accent:
        rect(slide, x, y, 0.19, h, accent)


def badge(slide, x, y, w, h, label, fill=PRIMARY, color=WHITE, size=13.5,
          kind=MSO_SHAPE.RECTANGLE, adj=None):
    sp = shape(slide, kind, x, y, w, h, fill=fill, line=fill, adj=adj)
    tf = sp.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(label)
    _set_run_font(r, size, color, True)
    return sp


# ----------------------------------------------------------------------------
# 表格
# ----------------------------------------------------------------------------
def _cell_border(cell, color=LINE, width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.Element(qn(tag))
        ln.set("w", str(int(width_pt * 12700)))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", color)
        tcPr.insert(0, ln)


def add_table(slide, x, y, w, h, head, rows, col_widths=None, font_size=13.0,
              head_size=14.0, first_col_bold=True, min_row=0.95):
    ncol = len(head)
    nrow = len(rows) + 1
    gt = slide.shapes.add_table(nrow, ncol, Cm(x), Cm(y), Cm(w), Cm(h))
    table = gt.table
    table.first_row = False
    table.horz_banding = False
    table.vert_banding = False

    # 列宽
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            table.columns[i].width = Cm(w * cw / total)
    else:
        for i in range(ncol):
            table.columns[i].width = Cm(w / ncol)

    # 行高：按内容估算
    head_h = max(1.05, min_row)
    table.rows[0].height = Cm(head_h)
    body_h = []
    for r in rows:
        need = min_row
        for c, val in zip(table.columns, r):
            cpl = max(1.0, (c.width.cm - 0.5) / (font_size * 0.03528))
            lines = sum(max(1, math.ceil(_units(s) / cpl)) for s in str(val).split("\n"))
            need = max(need, lines * font_size * 0.03528 * 1.28 + 0.42)
        body_h.append(need)
    avail = h - head_h
    if sum(body_h) > avail:  # 压缩字号
        k = avail / sum(body_h)
        font_size = max(9.5, font_size * k * 0.97)
        head_size = max(10.5, head_size * k * 0.97)
        body_h = [b * k for b in body_h]
    elif body_h:
        # 把多余高度分给各行，让表格铺满内容区（单行不超过 2.8cm）
        extra = (avail - sum(body_h)) / len(body_h)
        body_h = [min(b + extra, 2.3) for b in body_h]
        left = avail - sum(body_h)
        if left > 0.4:  # 触顶则整表垂直居中
            gt.top = Cm(y + left / 2.0)
    for i, bh in enumerate(body_h):
        table.rows[i + 1].height = Cm(bh)
    try:
        gt.height = Cm(head_h + sum(body_h))
    except Exception:
        pass

    for ci, txt in enumerate(head):
        cell = table.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(PRIMARY)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Cm(0.28)
        cell.margin_top = cell.margin_bottom = Cm(0.12)
        _cell_border(cell, PRIMARY, 0.75)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = Pt(head_size * 1.18)
        _add_runs(p, txt, head_size, WHITE, True)

    for ri, row in enumerate(rows):
        bg = CARD if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(bg)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Cm(0.28)
            cell.margin_top = cell.margin_bottom = Cm(0.10)
            _cell_border(cell, LINE, 0.75)
            tf = cell.text_frame
            tf.word_wrap = True
            lines = str(val).split("\n")
            for li, line in enumerate(lines):
                p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = Pt(font_size * 1.26)
                bold = bool(first_col_bold and ci == 0)
                color = PRIMARY if (first_col_bold and ci == 0) else TEXT
                _add_runs(p, line, font_size, color, bold)
    return gt


# ----------------------------------------------------------------------------
# 图片
# ----------------------------------------------------------------------------
def fit_image(slide, path, bx, by, bw, bh, caption=None):
    from PIL import Image  # 仅用于取尺寸；缺失时退化
    iw, ih = Image.open(path).size
    scale = min(bw / iw, bh / ih)
    w_cm, h_cm = iw * scale, ih * scale
    x = bx + (bw - w_cm) / 2
    y = by + (bh - h_cm) / 2
    slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w_cm), Cm(h_cm))
    return x, y, w_cm, h_cm


# ----------------------------------------------------------------------------
# 版式
# ----------------------------------------------------------------------------
class Deck:
    def __init__(self, spec, base_dir="."):
        self.spec = spec
        self.base = base_dir
        self.prs = Presentation()
        self.prs.slide_width = Cm(W)
        self.prs.slide_height = Cm(H)
        self.blank = self.prs.slide_layouts[6]
        self.page = 0
        self.no = 0

    # -- 通用页面骨架 ------------------------------------------------------
    def new(self):
        return self.prs.slides.add_slide(self.blank)

    def chrome(self, s, section):
        rect(s, 0, 0, W, 0.32, PRIMARY)
        rect(s, M, CHROME_BOTTOM, CW, 0.11, PRIMARY)
        textbox(s, 1.32, FOOTER_Y - 0.06, 24.0, 0.62,
                [dict(text=section or "", size=10.5, color=GRAY)], autofit=False)
        textbox(s, 28.6, FOOTER_Y - 0.06, 3.6, 0.62,
                [dict(text="%02d" % self.page, size=10.5, color=PRIMARY, bold=True)],
                align=PP_ALIGN.RIGHT, autofit=False)

    def header(self, s, title, number=None):
        label = str(number if number is not None else self.no)
        badge(s, M, 0.90, 1.38, 1.11, label, size=13.5)
        textbox(s, 3.07, 0.84, 29.10, 1.40,
                [dict(text=title, size=27, color=PRIMARY, bold=True, ls=1.12)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.72)

    def lead(self, s, text):
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, M, 2.54, CW, 1.38, fill=BAND, line=BAND, adj=0.15)
        textbox(s, M + 0.50, 2.80, CW - 1.0, 0.86,
                [dict(text=text, size=16.5, color=TEXT, bold=True, ls=1.22)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.72)

    def note(self, s, text, kind="conclusion", top_limit=BODY_BOTTOM):
        color = ORANGE if kind in ("warning", "limit") else PRIMARY
        w = CW - 0.45
        # 估算高度
        est = _estimate([dict(text=text, size=15, ls=1.28)], w - 0.5, 1.0)
        h = max(1.0, min(2.2, est + 0.30))
        y = top_limit - h
        rect(s, 1.32, y, 0.19, h, color)
        textbox(s, 1.75, y - 0.04, w, h + 0.08,
                [dict(text=text, size=15, color=color, bold=True, ls=1.28)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.70)
        return y

    # -- 单页渲染 ----------------------------------------------------------
    def cover(self, spec):
        s = self.new()
        rect(s, 0, 0, W, H, WHITE)
        rect(s, 0, 0, 11.11, H, PRIMARY)

        textbox(s, 1.64, 2.96, 8.2, 3.9, [
            dict(text=spec.get("left_kicker", "研究生"), size=44, color=WHITE, bold=True,
                 align=PP_ALIGN.CENTER, ls=1.35),
            dict(text=spec.get("left_week", ""), size=44, color=WHITE, bold=True,
                 align=PP_ALIGN.CENTER, ls=1.35),
            dict(text=spec.get("left_type", ""), size=28.5, color=WHITE, bold=True,
                 align=PP_ALIGN.CENTER, ls=1.4, space_before=6),
        ], anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)

        textbox(s, 1.33, 13.32, 9.13, 1.79, [
            dict(text=t, size=15.75, color="D9EDF8", ls=1.45) for t in spec.get("tags", [])
        ], min_scale=0.7)

        textbox(s, 12.91, 2.60, 18.8, 5.45, [
            dict(text=t, size=27, color=PRIMARY, bold=True, ls=1.30)
            for t in str(spec.get("title", "")).split("\n")
        ], anchor=MSO_ANCHOR.BOTTOM, min_scale=0.62)

        rect(s, 12.91, 8.73, 18.26, 0.13, PRIMARY)
        textbox(s, 12.91, 9.45, 18.2, 0.95,
                [dict(text=spec.get("meta_label", "本周学习主线"), size=16.5, color=BLUE2, bold=True)],
                anchor=MSO_ANCHOR.MIDDLE, autofit=False)
        textbox(s, 12.91, 11.35, 12.0, 1.05,
                [dict(text=spec.get("author") or self.spec.get("author", ""),
                      size=24, color=TEXT, bold=True)],
                anchor=MSO_ANCHOR.MIDDLE, autofit=False)
        textbox(s, 12.91, 12.95, 18.6, 1.35,
                [dict(text=spec.get("topics", ""), size=18, color=GRAY, ls=1.35)],
                min_scale=0.75)
        _date = spec.get("date_text") or str(self.spec.get("date", "")).replace("-", "")
        textbox(s, 27.6, 17.05, 5.4, 0.85,
                [dict(text=_date, size=13.5, color=GRAY)],
                align=PP_ALIGN.RIGHT, autofit=False)
        return s

    def part(self, spec):
        s = self.new()
        rect(s, 0, 0, W, H, WHITE)
        rect(s, 0, 0, 11.11, H, PRIMARY)
        textbox(s, 1.46, 7.73, 8.2, 2.4,
                [dict(text=spec.get("left_week", self.spec.get("week", "") + "组会"),
                      size=44, color=WHITE, bold=True, ls=1.3)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.6)

        textbox(s, 12.91, 3.10, 18.6, 2.2,
                [dict(text=spec.get("part", "Part 1"), size=60, color=PRIMARY, bold=True, ls=1.1)],
                min_scale=0.7)
        textbox(s, 12.91, 5.75, 18.6, 2.2,
                [dict(text=spec.get("title", ""), size=54, color=PRIMARY, bold=True, ls=1.15)],
                min_scale=0.55)
        if spec.get("module"):
            textbox(s, 12.91, 8.55, 18.6, 2.4,
                    [dict(text=spec["module"], size=40, color=TEXT, bold=True, ls=1.25)],
                    min_scale=0.5)
        if spec.get("kicker"):
            textbox(s, 13.08, 14.30, 18.4, 1.05,
                    [dict(text=spec["kicker"], size=13.5, color=GRAY, ls=1.35)],
                    anchor=MSO_ANCHOR.BOTTOM, min_scale=0.75)
        rect(s, 13.08, 15.92, 18.26, 0.13, PRIMARY)
        return s

    def toc(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", "目录"))
        self.header(s, spec.get("title", "目录"), spec.get("no"))
        items = spec.get("items", [])[:6]
        y = 4.87
        step = 2.06 if len(items) >= 5 else min(2.06, (16.9 - 4.87) / max(1, len(items)))
        for i, it in enumerate(items, 1):
            badge(s, M, y, 1.38, 1.11, str(it.get("no", i)), size=12.75)
            textbox(s, 3.23, y, 8.73, 1.11,
                    [dict(text=it.get("title", ""), size=13.5, color=TEXT, bold=True, ls=1.2)],
                    anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)
            textbox(s, 12.44, y, 19.74, 1.11,
                    [dict(text=it.get("desc", ""), size=9.75, color=GRAY, ls=1.3)],
                    anchor=MSO_ANCHOR.MIDDLE, min_scale=0.72)
            if i < len(items):
                rect(s, 3.23, y + 1.43, 28.95, 0.03, LINE)
            y += step
        return s

    def points(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        fig = spec.get("figure")
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        items = spec.get("items", [])[:3]
        text_w = 12.83 if fig else 28.0
        avail = bottom - top
        n = max(1, len(items))
        row_h = avail / n
        for i, it in enumerate(items):
            y = top + i * row_h
            badge(s, 1.32, y + 0.10, 1.16, 1.16, it.get("no", "%02d" % (i + 1)),
                  size=12.0, kind=MSO_SHAPE.TEAR, adj=1.0)
            textbox(s, 2.91, y + 0.06, text_w, 0.95,
                    [dict(text=it.get("title", ""), size=17.25, color=PRIMARY, bold=True, ls=1.2)],
                    anchor=MSO_ANCHOR.MIDDLE, min_scale=0.72)
            textbox(s, 2.91, y + 1.02, text_w, max(0.9, row_h - 1.15),
                    [dict(text=t, size=14.25, color=TEXT, ls=1.38) for t in _as_list(it.get("body"))],
                    min_scale=0.70)

        if fig:
            fx, fw = 17.20, 15.0
            cap_h = 0.95 if fig.get("caption") else 0.0
            fit_image(s, self._path(fig["path"]), fx, top, fw, bottom - top - cap_h)
            if fig.get("caption"):
                textbox(s, fx, bottom - cap_h, fw, cap_h,
                        [dict(text=fig["caption"], size=11.25, color=GRAY, ls=1.3)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.75)
        return s

    def cards(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        items = spec.get("items", [])
        n = len(items)
        if n <= 3:
            cols, rows = max(1, n), 1
        else:
            cols, rows = 2, math.ceil(n / 2)
        gap = 0.62
        cw = (CW - gap * (cols - 1)) / cols
        ch_avail = (bottom - top - gap * (rows - 1)) / rows
        if rows == 1:
            need = max([_block_height(it, cw - 1.1, has_label=bool(it.get("no")))
                        for it in items] + [3.2])
            ch = min(ch_avail, max(5.2, need + 0.2))
        else:
            ch = ch_avail
        y0 = top + (0 if rows > 1 else (ch_avail - ch) / 2)

        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            x = M + c * (cw + gap)
            y = y0 + r * (ch + gap)
            color = it.get("color") or ACCENTS[i % len(ACCENTS)]
            card(s, x, y, cw, ch, accent=color)
            label = it.get("no")
            if label:
                textbox(s, x + 0.55, y + 0.32, cw - 1.1, 0.62,
                        [dict(text=str(label), size=12.75, color=SKY, bold=True)],
                        autofit=False)
            hy = y + (0.95 if label else 0.38)
            textbox(s, x + 0.55, hy, cw - 1.1, 0.92,
                    [dict(text=it.get("title", ""), size=18.75, color=color, bold=True, ls=1.18)],
                    anchor=MSO_ANCHOR.MIDDLE, min_scale=0.62)
            body_top = hy + 1.02
            body_h = max(0.8, y + ch - body_top - 0.25)
            paras = []
            for t in _as_list(it.get("body")):
                mono = t.startswith("`") and t.endswith("`")
                paras.append(dict(text=t.strip("`"), size=15.75 if not mono else 13.5,
                                  color=TEXT, ls=1.32, space_after=4,
                                  font=MONO if mono else FONT))
            textbox(s, x + 0.55, body_top, cw - 1.1, body_h, paras, min_scale=0.62)
        return s

    def table(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        side = spec.get("side") or []
        tw = CW if not side else CW * 0.53
        th = bottom - top
        add_table(s, M, top, tw, th, spec["head"], spec["rows"],
                  col_widths=spec.get("col_widths"), font_size=spec.get("font_size", 13.0))

        if side:
            x = M + tw + 0.65
            w = CW - tw - 0.65
            items = side[:3]
            gap = 0.42
            ih = (th - gap * (len(items) - 1)) / max(1, len(items))
            for i, it in enumerate(items):
                y = top + i * (ih + gap)
                card(s, x, y, w, ih, accent=ACCENTS[(i + 1) % len(ACCENTS)])
                textbox(s, x + 0.5, y + 0.22, w - 1.0, 0.8,
                        [dict(text=it.get("title", ""), size=16.5, color=PRIMARY, bold=True)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)
                textbox(s, x + 0.5, y + 1.02, w - 1.0, max(0.6, ih - 1.15),
                        [dict(text=t, size=13.5, color=TEXT, ls=1.32) for t in _as_list(it.get("body"))],
                        min_scale=0.66)
        return s

    def flow(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD

        steps = spec.get("steps", [])
        n = max(1, len(steps))
        gap = 0.58
        cw = (CW - gap * (n - 1)) / n
        ch = spec.get("step_height", 5.55)
        for i, st in enumerate(steps):
            x = M + i * (cw + gap)
            sp = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, top, cw, ch, fill=CARD, line=LINE, adj=0.04)
            textbox(s, x + 0.42, top + 0.35, cw - 0.84, 0.66,
                    [dict(text=st.get("no", "%02d" % (i + 1)), size=12.75, color=SKY, bold=True)],
                    autofit=False)
            textbox(s, x + 0.42, top + 1.05, cw - 0.84, 0.95,
                    [dict(text=st.get("title", ""), size=16.5, color=PRIMARY, bold=True, ls=1.2)],
                    min_scale=0.6)
            textbox(s, x + 0.42, top + 2.12, cw - 0.84, ch - 2.35,
                    [dict(text=t, size=13.5, color=TEXT, ls=1.3) for t in _as_list(st.get("body"))],
                    min_scale=0.6)

        y = top + ch + 0.55
        panels = spec.get("panels") or []
        note_y = BODY_BOTTOM
        if spec.get("note"):
            note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion"))
        avail = note_y - 0.25 - y
        if panels and avail > 1.8:
            pw = (CW - 0.62 * (len(panels) - 1)) / len(panels)
            for i, p in enumerate(panels):
                x = M + i * (pw + 0.62)
                color = p.get("color") or (TEAL if i % 2 == 0 else ORANGE)
                card(s, x, y, pw, avail, accent=color)
                bodies = [dict(text=t, size=15.75, color=TEXT, ls=1.35) for t in _as_list(p.get("body"))]
                body_h = _estimate(bodies, pw - 1.2, 1.0)
                by = y + max(0.35, (avail - 0.95 - body_h) / 2)
                textbox(s, x + 0.6, by, pw - 1.2, 0.85,
                        [dict(text=p.get("title", ""), size=18.75, color=color, bold=True)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.65)
                textbox(s, x + 0.6, by + 1.00, pw - 1.2, body_h + 0.3,
                        bodies, min_scale=0.62)
        return s

    def compare(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        panels = spec.get("panels", [])[:3]
        n = max(1, len(panels))
        gap = 0.62
        pw = (CW - gap * (n - 1)) / n
        ph = bottom - top
        for i, p in enumerate(panels):
            x = M + i * (pw + gap)
            color = p.get("color") or ACCENTS[i % len(ACCENTS)]
            card(s, x, top, pw, ph, accent=color)

            bodies = [dict(text=t, size=15.75, color=TEXT, ls=1.4) for t in _as_list(p.get("body"))]
            body_h = _estimate(bodies, pw - 1.1, 1.0)
            block = (0.95 if p.get("tag") else 0) + 1.05 + body_h
            y = top + max(0.4, (ph - block) / 2)

            if p.get("tag"):
                shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.55, y, 3.2, 0.68,
                      fill=color, line=color, adj=0.12)
                textbox(s, x + 0.55, y, 3.2, 0.68,
                        [dict(text=p["tag"], size=9.0, color=WHITE, bold=True)],
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, autofit=False)
                y += 0.95
            textbox(s, x + 0.55, y, pw - 1.1, 1.05,
                    [dict(text=p.get("title", ""), size=20, color=color, bold=True, ls=1.15)],
                    anchor=MSO_ANCHOR.MIDDLE, min_scale=0.62)
            textbox(s, x + 0.55, y + 1.10, pw - 1.1, body_h + 0.35,
                    bodies, min_scale=0.60)
        return s

    def figure(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        side = spec.get("side") or []
        cap_h = 1.0 if spec.get("caption") else 0.0
        if side:
            iw = CW * 0.52
            fit_image(s, self._path(spec["image"]), M, top, iw, bottom - top - cap_h)
            if spec.get("caption"):
                textbox(s, M, bottom - cap_h, iw, cap_h,
                        [dict(text=spec["caption"], size=11.25, color=GRAY, ls=1.3)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.75)
            x = M + iw + 0.75
            w = CW - iw - 0.75
            items = side[:3]
            gap = 0.40
            ih = (bottom - top - gap * (len(items) - 1)) / max(1, len(items))
            for i, it in enumerate(items):
                y = top + i * (ih + gap)
                badge(s, x, y + 0.05, 1.0, 1.0, it.get("no", "%02d" % (i + 1)),
                      size=10.5, kind=MSO_SHAPE.TEAR, adj=1.0)
                textbox(s, x + 1.35, y, w - 1.35, 0.85,
                        [dict(text=it.get("title", ""), size=16.5, color=PRIMARY, bold=True)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)
                textbox(s, x + 1.35, y + 0.92, w - 1.35, max(0.7, ih - 1.05),
                        [dict(text=t, size=13.5, color=TEXT, ls=1.35) for t in _as_list(it.get("body"))],
                        min_scale=0.66)
        else:
            fit_image(s, self._path(spec["image"]), M, top, CW, bottom - top - cap_h)
            if spec.get("caption"):
                textbox(s, M, bottom - cap_h, CW, cap_h,
                        [dict(text=spec["caption"], size=11.25, color=GRAY, ls=1.3)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.75)
        return s

    def stats(self, spec):
        s = self.new()
        self.chrome(s, spec.get("section", ""))
        self.header(s, spec.get("title", ""), spec.get("no"))
        has_lead = bool(spec.get("lead"))
        if has_lead:
            self.lead(s, spec["lead"])
        top = BODY_TOP if has_lead else BODY_TOP_NOLEAD
        note_y = self.note(s, spec["note"], spec.get("note_kind", "conclusion")) if spec.get("note") else BODY_BOTTOM
        bottom = note_y - 0.42

        items = spec.get("items", [])[:4]
        n = max(1, len(items))
        gap = 0.62
        cw = (CW - gap * (n - 1)) / n
        blocks = spec.get("blocks") or []
        avail_all = bottom - top

        # 先算块高，再把「数字卡 + 说明块」作为一个整体垂直居中
        if blocks:
            bw = (CW - gap * (len(blocks) - 1)) / len(blocks)
            need = max([_block_height(b, bw - 1.1, title_size=15.75, body_size=13.5)
                        for b in blocks] + [2.4])
            bh = min(avail_all * 0.36, max(2.6, need + 0.1))
        else:
            bh = 0.0

        ch = min(4.0, avail_all - bh - 0.55) if blocks else min(4.2, avail_all * 0.62)
        ch = max(2.6, ch)
        stack = ch + (0.55 + bh if blocks else 0)
        y_cards = top + max(0.0, (avail_all - stack) / 2)

        for i, it in enumerate(items):
            x = M + i * (cw + gap)
            card(s, x, y_cards, cw, ch, accent=ACCENTS[i % len(ACCENTS)])
            textbox(s, x + 0.3, y_cards + 0.18, cw - 0.6, ch * 0.55,
                    [dict(text=str(it.get("value", "")), size=40, color=PRIMARY, bold=True)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, min_scale=0.55)
            textbox(s, x + 0.3, y_cards + ch * 0.62, cw - 0.6, ch * 0.26,
                    [dict(text=it.get("label", ""), size=12, color=GRAY)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)

        if blocks:
            y = y_cards + ch + 0.55
            for i, b in enumerate(blocks):
                x = M + i * (bw + gap)
                card(s, x, y, bw, bh, accent=ACCENTS[(i + 1) % len(ACCENTS)])
                textbox(s, x + 0.55, y + 0.20, bw - 1.1, 0.8,
                        [dict(text=b.get("title", ""), size=15.75, color=PRIMARY, bold=True)],
                        anchor=MSO_ANCHOR.MIDDLE, min_scale=0.7)
                textbox(s, x + 0.55, y + 0.98, bw - 1.1, max(0.6, bh - 1.15),
                        [dict(text=t, size=13.5, color=TEXT, ls=1.34) for t in _as_list(b.get("body"))],
                        min_scale=0.66)
        return s

    def closing(self, spec):
        s = self.new()
        rect(s, 0, 0, W, H, WHITE)
        rect(s, 0, 0, 11.11, H, PRIMARY)
        textbox(s, 1.46, 7.73, 8.2, 2.4,
                [dict(text=spec.get("left_week", self.spec.get("week", "") + "组会"),
                      size=44, color=WHITE, bold=True, align=PP_ALIGN.CENTER, ls=1.3)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.6)
        textbox(s, 12.91, 5.60, 19.8, 4.2,
                [dict(text=spec.get("message", "本周汇报完毕，请老师批评指正！"),
                      size=25, color=PRIMARY, bold=True, ls=1.5)],
                anchor=MSO_ANCHOR.MIDDLE, min_scale=0.6)
        rect(s, 12.91, 10.60, 18.26, 0.13, PRIMARY)
        textbox(s, 12.91, 11.30, 12.0, 1.22,
                [dict(text=spec.get("author", self.spec.get("author", "")), size=24,
                      color=TEXT, bold=True)],
                anchor=MSO_ANCHOR.MIDDLE, autofit=False)
        textbox(s, 12.91, 12.60, 18.0, 0.9,
                [dict(text=spec.get("date_text", ""), size=13.5, color=GRAY)],
                autofit=False)
        return s

    # -- 输出 --------------------------------------------------------------
    def _fallback_note(self, idx):
        """给没有 speech 的页（封面 / 过渡 / 结束）补一条可念的备注。"""
        cover = self.spec.get("cover", {})
        slides = self.spec.get("slides", [])
        if idx == 0:
            title = str(cover.get("title", "")).replace("\n", "")
            n = len(slides)
            return "开场：本周汇报《%s》，共 %d 页，分模块过一遍学习内容与文献结论。" % (title, n)
        sl = slides[idx - 1] if idx - 1 < len(slides) else {}
        t = sl.get("type", "")
        if t == "part":
            return "过渡：进入 %s，主题是%s。" % (
                sl.get("part") or sl.get("kicker") or "下一部分",
                sl.get("module") or sl.get("title") or "")
        if t == "closing":
            return "收尾：本周内容汇报完毕，请老师批评指正。"
        if t == "toc":
            return "目录：先讲学习任务，再进文献精读，最后给结论与下一步。"
        return ""

    def _path(self, p):
        if os.path.isabs(p) or os.path.exists(p):
            return p
        cand = os.path.join(self.base, p)
        return cand if os.path.exists(cand) else p

    def build(self, out):
        spec = self.spec
        self.cover(spec.get("cover", {}))
        self.page = 1

        counters = {}
        for sl in spec.get("slides", []):
            self.page += 1
            t = sl.get("type", "points")
            if t not in ("part", "closing"):
                if sl.get("no"):
                    self.no = int(sl["no"]) if str(sl["no"]).isdigit() else self.no + 1
                else:
                    self.no += 1
            fn = getattr(self, t, None)
            if fn is None:
                raise SystemExit("未知版式: %s（可用: cover/part/toc/points/cards/table/flow/compare/figure/stats/closing）" % t)
            if t == "cover":
                raise SystemExit("cover 只能出现在 deck.cover 字段")
            slide = fn(sl)
            # 演讲者备注：四段式讲稿
            sp = sl.get("speech") or {}
            if sp:
                notes = []
                for k, label in (("why", "Why 动机"), ("how", "How 做法"),
                                 ("read", "Read 图表"), ("so_what", "So What 结论"),
                                 ("transition", "过渡")):
                    if sp.get(k):
                        notes.append("[%s] %s" % (label, sp[k]))
                slide.notes_slide.notes_text_frame.text = "\n".join(notes)
            counters[t] = counters.get(t, 0) + 1

        # 封面、过渡页、结束页即使没写 speech，也要有可念的备注，讲稿才不断档
        for i, sld in enumerate(self.prs.slides):
            if sld.has_notes_slide and (sld.notes_slide.notes_text_frame.text or "").strip():
                continue
            note = self._fallback_note(i)
            if note:
                sld.notes_slide.notes_text_frame.text = note

        out_dir = os.path.dirname(os.path.abspath(out))
        os.makedirs(out_dir, exist_ok=True)
        self.prs.save(out)
        return len(self.prs.slides._sldIdLst)


def _block_height(item, width, title_size=18.75, body_size=15.75, has_label=False):
    """估算一张卡片的自然高度（cm），用于避免大块留白。"""
    h = 0.42 + (0.95 if has_label else 0.0)
    h += _estimate([dict(text=item.get("title", ""), size=title_size, ls=1.18)], width, 1.0) + 0.22
    bodies = [t.strip("`") for t in _as_list(item.get("body"))]
    h += _estimate([dict(text=t, size=body_size, ls=1.32, space_after=4) for t in bodies], width, 1.0)
    h += 0.32
    return h


def _as_list(v):
    if v is None:
        return []
    if isinstance(v, (list, tuple)):
        return [str(x) for x in v]
    return [str(v)]


# ----------------------------------------------------------------------------
# 内容规则自检（对齐组会汇报规范：标题公式 / 总分总 / 边界句 / 讲稿四步）
# ----------------------------------------------------------------------------
# 纯禁词：出现即返工
FUZZY_HARD = ("我觉得", "我认为", "个人认为", "或许", "也许", "大概", "似乎", "估计是")
# 软化词：降为警告，要求补条件
FUZZY_SOFT = ("可能", "一般来说", "应该是")
# 边界/局限语：每页的总结句必须命中其一
BOUNDARY = ("边界", "局限", "注意", "适用条件", "不代表", "不等于", "只能",
            "仅在", "不是", "取决于", "不能", "前提是")
# 读图用语：讲稿 Read 段必须命中其一，否则不叫"精读"
OBSERVE = ("看出", "观察到", "可以看到", "可见", "显示", "说明")
# 坐标/方位用语：表格看行列，图看横纵轴
AXIS = ("横轴", "纵轴", "坐标", "横坐标", "纵坐标", "行", "列",
        "左图", "右图", "左上方", "右上方", "上方", "下方", "颜色", "峰值", "区域")
GENERIC_TITLE = ("方法介绍", "结果展示", "背景介绍", "研究背景", "结果", "讨论", "小结", "总结")
CHART_STRICT = ("figure", "table", "stats")   # 真图表/数据：Read 必填且必须讲怎么看
CHART_SOFT = ("flow", "compare")             # 流程与对照：Read 可略，写了就要有内容
LEAD_TYPES = ("points", "cards", "table", "flow", "compare", "stats", "figure")


def _iter_text(sl):
    """遍历一页里所有成句的文字，供禁词扫描。"""
    for k in ("title", "lead", "note", "caption"):
        if sl.get(k):
            yield k, str(sl[k])
    for key in ("items", "steps", "panels", "blocks"):
        for blk in (sl.get(key) or []):
            if not isinstance(blk, dict):
                continue
            for k in ("title", "label", "value", "tag", "body"):
                v = blk.get(k)
                for t in ([v] if isinstance(v, str) else _as_list(v)) if v else []:
                    if t:
                        yield key, str(t)
    for row in (sl.get("rows") or []):
        for cell in row:
            if cell:
                yield "rows", str(cell)
    for k, v in (sl.get("speech") or {}).items():
        if v:
            yield "speech." + k, str(v)


def _sentences(text):
    return [x for x in re.split(r"[。！？；\n]", str(text)) if x.strip()]


def check(spec):
    issues, warns = [], []
    cover = spec.get("cover", {})
    for k in ("title", "date_text"):
        if not cover.get(k):
            issues.append("封面缺少 %s" % k)
    if not (cover.get("author") or spec.get("author")):
        issues.append("封面缺少 author")

    for i, sl in enumerate(spec.get("slides", []), 2):
        t = sl.get("type")
        p = "第%d页" % i
        if t == "part":
            if not sl.get("part"):
                issues.append("%s part 缺少 part 编号（应为 Part 1 / Part 2）" % p)
            continue
        if t in ("closing", "toc"):
            continue

        # --- 标题：序号 + 动作/对象 + 目的，或 结论X：核心发现 ---
        title = (sl.get("title") or "").strip()
        if not title:
            issues.append("%s 缺少标题" % p)
        else:
            if len(title) < 6:
                issues.append("%s 标题过短，需写清「做了什么、得到什么」：%s" % (p, title))
            bare = title.rstrip("：: ")
            if bare in GENERIC_TITLE or any(bare == g for g in GENERIC_TITLE):
                issues.append("%s 标题太空泛（%s），必须写成公式A/公式B" % (p, title))
            if "：" not in title and ":" not in title:
                warns.append("%s 标题没有冒号，可能不符合公式A/公式B：%s" % (p, title))

        # --- 总分总：导语 + 分块 + 总结句 ---
        if t in LEAD_TYPES and not sl.get("lead"):
            issues.append("%s 缺少导语句（lead，一句话核心事实）" % p)
        if t in LEAD_TYPES and not sl.get("note"):
            issues.append("%s 缺少总结/边界句（note，必须带局限或注意）" % p)
        elif sl.get("note") and not any(w in str(sl["note"]) for w in BOUNDARY):
            issues.append("%s 的 note 没写边界（缺「结论边界/注意/适用条件/不代表」等限定）：%s"
                          % (p, str(sl["note"])[:30]))
        if len(sl.get("items", []) or []) > 3 and t in ("points", "cards"):
            issues.append("%s 分块超过 3 个，应拆分" % p)
        if t == "stats":
            for it in (sl.get("items") or []):
                if not re.search(r"\d", str(it.get("value", ""))):
                    issues.append("%s 的大数字「%s」不含数据" % (p, it.get("value")))

        # --- 图表：图注必须带 Fig. 编号，表格必须两列以上 ---
        if t == "figure":
            cap = str(sl.get("caption") or "")
            if not cap:
                issues.append("%s 图表缺少图注（caption）" % p)
            elif not re.search(r"(Fig\.?\s*\d|图\s*\d|Figure\s*\d)", cap, re.I):
                issues.append("%s 图注缺 Fig. 编号（示例：论文 Fig.1B：中央凹结构示意）" % p)
        if t == "table":
            if not sl.get("head") or not sl.get("rows"):
                issues.append("%s 表格缺 head/rows" % p)
            elif len(sl["head"]) > 5:
                issues.append("%s 表格列数 >5，建议精简" % p)

        for blk in (sl.get("items") or []) + (sl.get("steps") or []):
            if len(_as_list(blk.get("body"))) > 3:
                issues.append("%s 「%s」正文超过 3 句，需拆成 01/02/03"
                              % (p, blk.get("title", "")))

        # --- 讲稿四步 ---
        sp = sl.get("speech") or {}
        for k, label in (("why", "Why 动机引入"), ("how", "How 具体做法"),
                         ("so_what", "So What 总结与边界")):
            if not sp.get(k):
                issues.append("%s 讲稿缺少 %s" % (p, label))
        if t in CHART_STRICT:
            read = sp.get("read")
            if not read:
                issues.append("%s 有图表但讲稿缺少 Read 图表精读" % p)
            else:
                if not any(w in read for w in AXIS):
                    issues.append("%s 讲稿 Read 段没讲怎么看图（横纵轴/行列/颜色/区域）" % p)
                if not any(w in read for w in OBSERVE):
                    issues.append("%s 讲稿 Read 段缺「从图中可以看出/观察到」的结论句" % p)
        elif t in CHART_SOFT and not sp.get("read"):
            warns.append("%s 是流程/对照页，建议补一句 Read 说明看哪一侧" % p)
        if t == "compare":
            warns.append("%s 用了双栏对照；规范要求对比/并列一律用 table，"
                         "compare 只留给「三栏并列且不是对比」的场合" % p)
        if sp.get("so_what") and not any(w in str(sp["so_what"]) for w in BOUNDARY):
            issues.append("%s 讲稿 So What 段没提结论边界（如「这只是转录组，不代表蛋白水平」）" % p)

        # --- 语言：禁模糊词 + 句子要短 ---
        for field, txt in _iter_text(sl):
            for w in FUZZY_HARD:
                if w in txt:
                    issues.append("%s「%s」出现模糊词「%s」，改为「说明/表明/体现」" % (p, field, w))
            for w in FUZZY_SOFT:
                if w in txt:
                    warns.append("%s「%s」用了「%s」，建议补条件（在…条件下）" % (p, field, w))
        for blk in (sl.get("items") or []) + (sl.get("steps") or []):
            for sent in _sentences("".join(_as_list(blk.get("body")))):
                if len(sent) > 46:
                    warns.append("%s 句子偏长（%d 字），建议断句：%s…"
                                 % (p, len(sent), sent[:16]))
    return issues, warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)

    issues, warns = check(spec)
    if issues:
        print("[自检] 发现 %d 处不符合内容规则：" % len(issues))
        for x in issues:
            print("  - " + x)
    else:
        print("[自检] 通过")
    if warns:
        print("[提醒] %d 条建议（不阻断，但最好改）：" % len(warns))
        for x in warns:
            print("  · " + x)

    if args.check:
        return 0 if not issues else 1

    out = args.out or os.path.splitext(args.spec)[0] + ".pptx"
    deck = Deck(spec, base_dir=os.path.dirname(os.path.abspath(args.spec)))
    deck.build(out)
    print("[输出] " + out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
