#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
离线版式预览器
==============
不依赖 PowerPoint / LibreOffice，直接读 .pptx 的几何与文字，用 PIL 画成 PNG。
用途：生成后快速自检——元素是否越界、文字是否溢出、层级是否合理。

用法:
    python preview.py 输出.pptx -o preview/
"""

import argparse
import os
import sys

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

PPC = 40.0  # 每厘米像素
FONT_DIR = r"C:\Windows\Fonts"
F_REG = os.path.join(FONT_DIR, "msyh.ttc")
F_BOLD = os.path.join(FONT_DIR, "msyhbd.ttc")
F_MONO = os.path.join(FONT_DIR, "consola.ttf")

_cache = {}


def font(size_px, bold=False, mono=False):
    key = (round(size_px), bold, mono)
    if key in _cache:
        return _cache[key]
    path = F_MONO if mono else (F_BOLD if bold else F_REG)
    if not os.path.exists(path):
        path = F_REG
    try:
        f = ImageFont.truetype(path, max(6, int(round(size_px))))
    except Exception:
        f = ImageFont.load_default()
    _cache[key] = f
    return f


def emu2px(v):
    return float(Emu(v).cm) * PPC


def hexcolor(color_obj, default=None):
    try:
        if color_obj and color_obj.type is not None:
            return "#" + str(color_obj.rgb)
    except Exception:
        pass
    return default


def wrap(text, fnt, max_w):
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur)
            cur = ""
            continue
        trial = cur + ch
        if fnt.getlength(trial) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = ch
    lines.append(cur)
    return lines


def draw_text_frame(d, tf, x, y, w, h, default_bold=False, anchor=None):
    paras = [p for p in tf.paragraphs if (p.text or "").strip()]
    if not paras:
        return
    # 收集行
    rendered = []  # (lines, font, color, line_h, align, space_before, space_after, indent)
    total_h = 0
    for p in paras:
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        size_pt = None
        bold = default_bold
        color = "#1B2F41"
        for r in runs:
            if r.font.size:
                size_pt = r.font.size.pt
            if r.font.bold:
                bold = True
            c = hexcolor(r.font.color, None)
            if c:
                color = c
        size_pt = size_pt or 12
        fnt = font(size_pt * PPC * 0.03528, bold)
        ls = p.line_spacing
        if hasattr(ls, "pt"):          # 精确行距（磅）
            line_h = ls.pt * PPC * 0.03528
        elif isinstance(ls, float):    # 倍数行距
            line_h = size_pt * PPC * 0.03528 * ls
        else:
            line_h = size_pt * PPC * 0.03528 * 1.2
        align = {PP_ALIGN.CENTER: "center", PP_ALIGN.RIGHT: "right"}.get(p.alignment, "left")
        text = "".join(r.text for r in runs)
        lines = wrap(text, fnt, w)
        sb = p.space_before.pt * PPC * 0.03528 if p.space_before else 0
        sa = p.space_after.pt * PPC * 0.03528 if p.space_after else 0
        total_h += sb + line_h * len(lines) + sa
        rendered.append((lines, fnt, color, line_h, align, sb, sa))

    anchor = anchor or tf.vertical_anchor
    cy = y
    if anchor == MSO_ANCHOR.MIDDLE:
        cy = y + (h - total_h) / 2
    elif anchor == MSO_ANCHOR.BOTTOM:
        cy = y + h - total_h

    for lines, fnt, color, line_h, align, sb, sa in rendered:
        cy += sb
        for ln in lines:
            tw = fnt.getlength(ln)
            if align == "center":
                tx = x + (w - tw) / 2
            elif align == "right":
                tx = x + w - tw
            else:
                tx = x
            d.text((tx, cy), ln, font=fnt, fill=color)
            cy += line_h
        cy += sa


def shape_fill(sh):
    try:
        if sh.fill.type == 1:  # solid
            return "#" + str(sh.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def shape_line(sh):
    try:
        if sh.line.color and sh.line.color.type is not None:
            w = sh.line.width.pt if sh.line.width else 0.75
            return "#" + str(sh.line.color.rgb), max(0.6, w * PPC * 0.03528)
    except Exception:
        pass
    return None


def round_rect(d, box, r, fill, outline):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=1)


def render_slide(slide, W, H, prs_w, prs_h):
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    for sh in slide.shapes:
        x, y = emu2px(sh.left), emu2px(sh.top)
        w, h = emu2px(sh.width), emu2px(sh.height)
        st = str(sh.shape_type)
        try:
            if "PICTURE" in st:
                pic = Image.open(sh.image.blob and __import__("io").BytesIO(sh.image.blob))
                pic = pic.convert("RGB").resize((max(1, int(w)), max(1, int(h))))
                img.paste(pic, (int(x), int(y)))
                continue
        except Exception:
            pass

        if sh.has_table:
            tbl = sh.table
            cy = y
            for r_i, row in enumerate(tbl.rows):
                cx = x
                rh = emu2px(row.height)
                for c_i, cell in enumerate(row.cells):
                    cw = emu2px(tbl.columns[c_i].width)
                    fill = shape_fill(cell) or "#FFFFFF"
                    d.rectangle([cx, cy, cx + cw, cy + rh], fill=fill, outline="#C7DCE9")
                    pad = 0.12 * PPC
                    # 单元格锚点在 tcPr@anchor 上，不在 a:bodyPr，需优先取 cell 的
                    draw_text_frame(d, cell.text_frame, cx + pad, cy + pad * 0.4,
                                    cw - 2 * pad, rh - pad * 0.8,
                                    anchor=getattr(cell, "vertical_anchor", None) or cell.text_frame.vertical_anchor)
                    cx += cw
                cy += rh
            continue

        fill = shape_fill(sh)
        line = shape_line(sh)
        outline = line[0] if line else None
        name = sh.name or ""
        if "ROUNDED" in str(sh.shape_type) or "圆角" in name:
            adj = 0.04
            try:
                adj = sh.adjustments[0]
            except Exception:
                pass
            round_rect(d, [x, y, x + w, y + h], min(w, h) * adj, fill, outline)
        elif "TEAR" in str(sh.shape_type) or "泪滴" in name:
            d.ellipse([x, y, x + w, y + h], fill=fill, outline=outline)
        elif fill or outline:
            d.rectangle([x, y, x + w, y + h], fill=fill, outline=outline)

        if sh.has_text_frame and sh.text_frame.text.strip():
            draw_text_frame(d, sh.text_frame, x, y, w, h)
    return img


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--width", type=int, default=1355)
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    ratio = prs.slide_height / prs.slide_width
    W = args.width
    H = int(W * ratio)
    out_dir = args.out or (os.path.splitext(args.pptx)[0] + "_preview")
    os.makedirs(out_dir, exist_ok=True)

    for i, s in enumerate(prs.slides, 1):
        im = render_slide(s, W, H, prs.slide_width, prs.slide_height)
        im.save(os.path.join(out_dir, "slide%02d.png" % i))
    print("[预览] %s（%d 页）" % (out_dir, len(prs.slides._sldIdLst)))


if __name__ == "__main__":
    sys.exit(main())
