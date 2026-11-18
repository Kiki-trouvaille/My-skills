#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""交付前自动检查：PPTX 是否可打开、字体是否统一、文字是否溢出。

用法：
    python check.py 输出.pptx
    python check.py 输出.pptx --strict      # 警告也当失败
    python check.py 输出.pptx --no-external # 不调用外部校验器

检查分三层，任何一层发现问题都要回到 deck.json 修，不要改脚本里的设计令牌：
    1. 包结构：zip 完整性、每个 XML 可解析、页面尺寸正确
    2. 内容：页数、每页是否带演讲者备注、字体是否全为微软雅黑
    3. 版式：用与 preview.py 相同的字体度量估算每个文本框是否溢出容器
外部校验器（可选，存在才跑）：
    pptx 技能的 scripts/validators.py       → OOXML 结构校验
    ppt-master 的 scripts/pptx_delivery_check.py → 交付风险清单
"""

import argparse
import glob
import os
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation                       # noqa: E402
from pptx.util import Emu                           # noqa: E402

import preview                                      # noqa: E402  复用同一套字体度量

FONT_OK = {"微软雅黑"}
TOL_CM = 0.08          # 允许的估算误差
CANVAS = (33.87, 19.05)

errors, warns = [], []


def err(msg):
    errors.append(msg)


def warn(msg):
    warns.append(msg)


# ---------- 1. 包结构 ----------
def check_package(path):
    if not zipfile.is_zipfile(path):
        err("不是有效的 zip/pptx 容器")
        return
    bad = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    ET.fromstring(z.read(n))
                except ET.ParseError as e:
                    bad.append("%s (%s)" % (n, e))
        if "[Content_Types].xml" not in names:
            err("缺少 [Content_Types].xml")
    for b in bad[:5]:
        err("XML 无法解析：" + b)
    if len(bad) > 5:
        err("另有 %d 个 XML 部件解析失败" % (len(bad) - 5))


# ---------- 2/3. 内容与版式 ----------
def tf_items(slide):
    """产出 (标签, text_frame, 宽cm, 高cm)，含表格单元格。"""
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            yield (sh.name, sh.text_frame, Emu(sh.width).cm, Emu(sh.height).cm)
        if getattr(sh, "has_table", False) and sh.has_table:
            tbl = sh.table
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    if cell.text_frame.text.strip():
                        # 单元格左右内边距各 0.25cm
                        yield ("%s[%d,%d]" % (sh.name, ri + 1, ci + 1), cell.text_frame,
                               Emu(tbl.columns[ci].width).cm - 0.5, Emu(row.height).cm)


def measure(tf, w_cm):
    """估算文本所需高度（cm），与 preview.draw_text_frame 同一套度量。"""
    total = 0.0
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        size = None
        bold = False
        for r in runs:
            if r.font.size:
                size = r.font.size.pt
            bold = bold or bool(r.font.bold)
        size = size or 12
        fnt = preview.font(size * preview.PPC * 0.03528, bold)
        ls = p.line_spacing
        if hasattr(ls, "pt"):
            lh = ls.pt * 0.03528
        elif isinstance(ls, float):
            lh = size * 0.03528 * ls
        else:
            lh = size * 0.03528 * 1.2
        text = "".join(r.text for r in runs)
        lines = preview.wrap(text, fnt, w_cm * preview.PPC)
        sb = p.space_before.pt * 0.03528 if p.space_before else 0
        sa = p.space_after.pt * 0.03528 if p.space_after else 0
        total += sb + lh * len(lines) + sa
    return total


def check_deck(path):
    prs = Presentation(path)
    cw, ch = Emu(prs.slide_width).cm, Emu(prs.slide_height).cm
    if abs(cw - CANVAS[0]) > 0.05 or abs(ch - CANVAS[1]) > 0.05:
        err("画布尺寸 %.2f×%.2f cm，应为 %.2f×%.2f" % (cw, ch, CANVAS[0], CANVAS[1]))

    fonts = {}
    for i, slide in enumerate(prs.slides, 1):
        notes = ""
        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip()
        if not notes and i not in (1,):
            warn("第 %d 页没有演讲者备注（讲稿会缺这一页）" % i)

        for label, tf, w, h in tf_items(slide):
            need = measure(tf, max(0.8, w))
            if need > h + TOL_CM:
                warn("第 %d 页 %s 文字溢出：需要 %.2fcm，容器 %.2fcm" % (i, label, need, h))
            for p in tf.paragraphs:
                for r in p.runs:
                    name = r.font.name
                    if name:
                        fonts[name] = fonts.get(name, 0) + 1
    for name, n in sorted(fonts.items(), key=lambda kv: -kv[1]):
        if name not in FONT_OK:
            warn("出现非模板字体 %s（%d 处），可能显示不一致" % (name, n))
    return len(prs.slides)


# ---------- 4. 外部校验器 ----------
def _run(cmd, label):
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    except Exception as e:
        return None, "%s 无法运行：%s" % (label, e)
    return out.stdout + out.stderr, None


def check_external(path):
    skills = os.path.expanduser("~/.workbuddy/skills")

    # pptx 技能：OOXML 结构校验（XSD 缺失时退化为结构检查）
    v = os.path.join(skills, "pptx", "scripts", "validate.py")
    if os.path.isfile(v):
        out, e = _run([sys.executable, v, path], "pptx 校验")
        if e:
            warn(e)
        elif out:
            if "All validations PASSED" in out:
                pass
            elif "No such file or directory" in out and ".xsd" in out:
                warn("pptx 校验器缺少 XSD 骨架，只能做结构校验（属工具限制，非文件问题）")
            elif "FAILED" in out:
                err("pptx 结构校验未通过（详见 validate.py 输出）")

    # ppt-master：交付风险清单
    d = os.path.join(skills, "ppt-master", "scripts", "pptx_delivery_check.py")
    if os.path.isfile(d):
        out, e = _run([sys.executable, d, path], "ppt-master 交付检查")
        if out:
            if '"status": "failed"' in out:
                err("ppt-master 交付检查失败")
            elif "font_portability" in out:
                warn("字体可移植性：微软雅黑非 Office 通用字体，换电脑需确认已安装"
                     "（模板要求如此，属已知取舍）")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--strict", action="store_true", help="有警告也返回非零")
    ap.add_argument("--no-external", action="store_true")
    a = ap.parse_args()

    if not os.path.isfile(a.pptx):
        print("找不到文件：%s" % a.pptx)
        return 2

    check_package(a.pptx)
    try:
        n = check_deck(a.pptx)
    except Exception as e:
        err("无法解析演示文稿：%s" % e)
        n = 0
    if not a.no_external and not errors:
        check_external(a.pptx)

    print("检查 %s" % os.path.basename(a.pptx))
    print("  页数 %d ｜ 错误 %d ｜ 警告 %d" % (n, len(errors), len(warns)))
    for m in errors:
        print("  [错误] " + m)
    for m in warns:
        print("  [警告] " + m)
    bad = errors or (a.strict and warns)
    print("\n结论：%s" % ("不通过，先改 deck.json" if bad else "通过"))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
