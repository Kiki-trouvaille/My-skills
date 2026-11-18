---
name: group-meeting-ppt
description: 按固定的学术组会版式，把每周学习与文献内容生成组会汇报 PPT 及配套逐页讲稿（docx/md）。当用户说「组会汇报」「组会 PPT」「周报 PPT」「本周学习汇报」「文献精读汇报」「汇报稿」「讲稿」「按我以前的风格做 PPT」「和上周一样的版式」时使用。内置标题公式、每页总—分—总、对比用表格、流程用箭头、图表必配图注与结论边界、深蓝学术配色与微软雅黑、33.87×19.05cm 画布；版式确定性渲染，并产出可直接上台念的四步式讲稿。
version: "1.1.0"
agent_created: true
---

# 组会汇报 PPT（固定版式）

把「本周学了什么」变成一套**版式恒定**的组会 PPT 与讲稿。样式不由模型即兴发挥，而是由
`scripts/build_deck.py` 按设计令牌确定性渲染，因此每周产出完全一致。

## 输入

用户通常给：本周学习内容（文字/大纲）、文献 PDF 或图、周次与日期。缺什么补什么，不要反复追问：

| 缺项 | 默认处理 |
|---|---|
| 周次 | 从日期或上文推断，如「第十周」 |
| 日期 | 用当天日期，格式 `2026-11-20` / `20261120` |
| 汇报人 | 陈诗琦（除非用户另说） |
| 图片 | 无图就不要用 `figure` 版式，改用 `points` / `cards` |

## 执行流程

**Step 1 · 定骨架（先想清楚再动手）**

按「学习任务 → 文献精读 → 总结计划」三段式拆模块，每段 3–6 页。在对话里先给一行式目录：

```
Part 1 学习任务（3 页）｜Part 2 文献精读（5 页）｜Part 3 总结（1 页）
```

页数参考：单周 10–16 页；纯文献精读 14–18 页。**每页只讲一件事**，宁可拆页不要堆。

**Step 2 · 写 deck.json**

严格按 [`references/schema.md`](references/schema.md) 的字段写。每个内容页必须齐四件：

1. `title`：公式化标题（见 [`references/content.md`](references/content.md)）
2. `lead`：一句话核心事实（导语带）
3. 主体：`items` / `steps` / `rows` / `panels`，**分块不超过 3 个**
4. `note`：结论或边界（`note_kind: warning` 时显示为橙色「注意/局限」）
5. `speech`：`why / how / read / so_what / transition` 五段（见 [`references/speech.md`](references/speech.md)）

**Step 3 · 渲染**

```bash
python "${SKILL_DIR}/scripts/build_deck.py" deck.json -o "陈诗琦_第X周_主题.pptx"
python "${SKILL_DIR}/scripts/build_script.py" deck.json -o "陈诗琦_第X周_汇报讲稿.docx"
```

`build_deck.py` 会先跑内容自检，输出分两级：`-` 开头是**错误**（空泛标题、缺导语、缺边界句、
分块超 3、图注没编号、模糊词、讲稿缺段都算），`·` 开头是**提醒**（缺冒号、句子偏长、用了「可能」）。
**有错误必须先改 JSON 再交付**；提醒尽量也改掉。

**Step 4 · 自检版式（必做）**

```bash
python "${SKILL_DIR}/scripts/check.py" 输出.pptx
python "${SKILL_DIR}/scripts/preview.py" 输出.pptx -o preview/
```

`check.py` 先跑机器可判的部分：包结构能否打开、画布尺寸、每页有无备注、字体是否统一、
每个文本框估算是否溢出，并顺带调用本机已装的 `pptx`（OOXML 结构校验）与 `ppt-master`
（交付风险清单）两个技能做外部复检。**它报 错误 就改 JSON，报 警告 就逐条判断**。

再逐张看 `preview/slideNN.png`，只查三件事：
- 文字是否溢出卡片/表格（被截断或压边）
- 图表是否过小、图注是否贴近
- 页脚页码、章节名是否对应

发现问题改 JSON（减字、拆页、换版式），不要改脚本里的设计令牌。

**Step 5 · 交付**

用 `present_files` 同时给出 `.pptx` 与 `讲稿.docx`，并在回复里写一句：
共几页、几个模块、讲稿约几分钟、哪些页的图需要用户替换成自己的截图。
同时提醒一句：模板字体是微软雅黑，换电脑放映前确认已安装，否则自动回退成宋体。

## 硬性规则（违反即返工）

`build_deck.py` 的 `check()` 会按下面这些逐条拦，报错就先改 JSON，不要绕过。

- 每页标题必须写明「做了什么 / 得到什么」，不允许「方法介绍」「结果展示」「研究背景」这类空标题；
  标题一律带冒号（公式A：`序号 动作/对象：核心目的`；公式B：`序号 结论X：核心发现`）。
- 每页必须有 `note`（结论 / 注意 / 边界），且句子里要出现「边界 / 局限 / 注意 / 适用条件 /
  不代表 / 不等于 / 只能 / 仅在 / 不是 / 取决于 / 不能 / 前提是」之一。
- 对比信息一律表格；流程一律 `flow`（序号 + 箭头语义）；模块切换一律 `part` 过渡页。
- 图注必须带编号：`论文 Fig.1B：中央凹结构与侏儒通路示意`（脚本校验 `Fig.N` / `图 N`）。
- 有大数字的页，每个 `value` 必须含数字；表格页至少 2 行。
- 语言：短句、陈述句。**「我觉得 / 我认为 / 或许 / 也许 / 大概 / 似乎」出现即报错**；
  「可能 / 应该是」只提醒，改写成「在…条件下」；用「说明 / 表明 / 体现」。
- 单页正文分块 ≤ 3；每条正文 ≤ 3 句；单句 ≤ 46 字。
- 讲稿四步缺一不可：`why / how / so_what` 必填；`figure / table / stats` 页 `read` 必填，
  且必须讲到坐标或方位（横轴、纵轴、行、列、颜色、区域、峰值）+ 一句
  「从图中可以明显看出 / 观察到」；`flow / compare` 页 `read` 可略。
- `so_what` 必须落到边界（同上那组边界词），不许只写「因此结果可靠」。

## 版式速查

| type | 用途 | 关键字段 |
|---|---|---|
| `cover` | 封面（左侧深蓝竖块） | `cover.title/meta_label/topics` |
| `part` | 模块过渡（Part 1 / 2） | `part/title/module/kicker` |
| `toc` | 目录（最多 6 条） | `items[{no,title,desc}]` |
| `points` | 01/02/03 竖排（可配右图） | `items` + `figure` |
| `cards` | 2–4 张卡片并列 | `items[{no,title,body}]` |
| `table` | 对比表（可带右侧解读卡） | `head/rows/col_widths/side` |
| `flow` | 横向步骤（可带下方双栏） | `steps` + `panels` |
| `compare` | 仅限三栏并列且**不是对比**（对比一律用 `table`） | `panels[{tag,title,body,color}]` |
| `figure` | 大图 + 右侧解读 | `image/caption/side` |
| `stats` | 大数字 + 说明块 | `items[{value,label}]` + `blocks` |
| `closing` | 结束页 | `message` |

完整几何、配色、字号见 [`references/design.md`](references/design.md)。

## 依赖

```bash
pip install python-pptx python-docx pillow
```

`scripts/check.py` 的外部复检是可选增强：只有本机装有 `pptx` 技能（`scripts/validate.py`）
或 `ppt-master` 技能（`scripts/pptx_delivery_check.py`）时才会调用，缺了自动跳过，不影响主流程。

## 已知限制

- 本机没有 LibreOffice / PowerPoint COM 自动化，`preview.py` 是离屏近似渲染
  （同一套字体度量，但抗锯齿与真机略有差异）。定稿前建议在 PowerPoint 里翻一遍。
- 微软雅黑不是 Office 通用字体；换到未安装该字体的机器会回退为宋体，版式仍成立但字形变了。
- 文献图不会被生成，只能按比例缩放居中、加图注；缺图时改用 `points` / `cards` 版式。

## 参考文件

- [`references/design.md`](references/design.md) — 画布、网格、配色、字号、各版式线框
- [`references/content.md`](references/content.md) — 标题公式、总分总、措辞替换表
- [`references/speech.md`](references/speech.md) — 四步讲稿框架、范例、时长控制
- [`references/schema.md`](references/schema.md) — deck.json 全字段说明
- [`examples/第四周_示例.json`](examples/第四周_示例.json) — 可直接改用的完整样例

脚本：`build_deck.py`（出 pptx + 备注）、`build_script.py`（出讲稿）、`check.py`（交付前机器自检）、
`preview.py`（出 PNG 供目视检查）。
