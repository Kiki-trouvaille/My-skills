# deck.json 字段说明

一份 JSON 同时驱动 PPT 与讲稿，两者永远同步。UTF-8 编码，中文直接写。

## 顶层

```json
{
  "week": "第十周",
  "date": "2026-11-20",
  "author": "陈诗琦",
  "cover": { ... },
  "slides": [ ... ]
}
```

| 字段 | 必填 | 说明 |
|---|---|---|
| `week` | ✅ | 「第十周」，用于页脚与讲稿标题 |
| `date` | ✅ | `2026-11-20`，讲稿元信息 |
| `author` | ✅ | 汇报人 |
| `cover` | ✅ | 封面配置 |
| `slides` | ✅ | 页面数组，顺序即页序（不含封面） |

## cover

```json
"cover": {
  "left_kicker": "研究生",
  "left_week": "第十周组会",
  "left_type": "学习与分析汇报",
  "tags": ["空间转录组", "细胞邻域与配体受体", "文献精读"],
  "title": "从空间坐标到细胞邻域：\n组织结构的定量解读",
  "meta_label": "本周学习主线",
  "topics": "SAW 流程  /  邻域分析  /  文献精读",
  "author": "陈诗琦",
  "date_text": "20261120"
}
```

`title` 用 `\n` 手动断行，2–3 行为佳（第一行是主题，第二行是落点）。
`meta_label` 文献周写「本周共读文献」，学习周写「本周学习主线」。

## 通用字段（除 part / closing 外都要）

| 字段 | 说明 |
|---|---|
| `type` | 版式，见下表 |
| `no` | 序号方块里的数字。不写则自动连号；跨 Part 想重置就手写 |
| `section` | 页脚左侧章节名，如 `03  文献汇报｜结论` |
| `title` | 27pt 主标题，公式化（见 content.md） |
| `lead` | 导语一句话，进浅蓝带 |
| `note` | 底部总结句，必填 |
| `note_kind` | `conclusion`（默认，深蓝）/ `warning`（橙，注意与局限） |
| `speech` | 讲稿五段，见下 |

### speech

```json
"speech": {
  "why": "为了解决注释层级不清的问题，首先确认 GFF3 的字段结构。",
  "how": "具体来说，按第 9 列 Parent 分组，长度用 end − start + 1 计算。",
  "read": "请看左侧表格：第一列是字段，第三列是用途……从表中可以看出本数据没有 gene 行。",
  "so_what": "因此第一步不是写循环，而是确认层级。边界：该写法只适用于当前数据。",
  "transition": "字段清楚之后，下一步是把规则翻译成代码。"
}
```

`why / how / so_what` 必填；有图表时 `read` 必填；`transition` 可选但强烈建议。

## 各版式专有字段

### part — 模块过渡

```json
{"type": "part", "part": "Part 2", "title": "文献精读：",
 "module": "视网膜细胞类型演化", "kicker": "本模块 5 页：背景 → 方法 → 图 3/4/5 → 结论",
 "left_week": "第十周组会"}
```

### toc — 目录（≤6 条）

```json
{"type": "toc", "title": "目录", "section": "目录",
 "items": [{"no": "1", "title": "研究背景与核心问题", "desc": "六大类细胞；三种假说"}]}
```

### points — 01/02/03 竖排

```json
{"type": "points", "no": "1", "section": "03  文献汇报",
 "title": "研究背景：中央凹如何形成高精度视觉",
 "lead": "中央凹面积很小，却承担高精度视觉。",
 "items": [
   {"no": "01", "title": "细胞组成", "body": ["中央凹是否拥有外周没有的细胞类型？"]},
   {"no": "02", "title": "区域特化", "body": ["共享类型是否改变表达状态？"]}
 ],
 "figure": {"path": "fig/image1.png", "caption": "论文 Fig.1B：中央凹结构与侏儒通路示意"},
 "note": "视锥 → 侏儒双极细胞 → 侏儒神经节细胞是高精度通路。"}
```

`items` ≤3；`figure` 可选，有图时文字列自动收窄。`body` 支持字符串或字符串数组（每项一段）。

### cards — 卡片

```json
{"type": "cards", "items": [
  {"no": "01", "title": "数量核对", "body": ["蛋白数 17,064，与 mRNA 条数一致。"], "color": "084B82"}
]}
```

2 张并排、3 张并排、4 张 2×2 自动排布。`color` 可省略，默认按 深蓝/次级蓝/青绿/橙 轮转。
`body` 若以反引号包裹（`` `code` ``）会按等宽字体渲染。

### table — 表格

```json
{"type": "table", "head": ["比较", "Bulk", "scRNA-seq"],
 "rows": [["测序单位", "组织整体", "单个细胞"], ["能否看异质性", "不能", "能"]],
 "col_widths": [4, 5, 5],
 "side": [{"title": "核心差异", "body": ["一个平均值 vs 每个细胞一个身份"]}]}
```

`col_widths` 为相对权重，可省略（等宽）。`side` 可选，表格收窄到 53%，右侧放 2–3 张解读卡。
列数 ≤5；列数 ≥4 时字号会自动下调。

### flow — 步骤

```json
{"type": "flow",
 "steps": [{"no": "01", "title": "读入 FASTA", "body": ["多行序列拼成字典"]}],
 "step_height": 5.55,
 "panels": [{"title": "正链 +", "body": ["按坐标升序拼接。"], "color": "27877B"}]}
```

`steps` 3–5 个等宽排列；`panels` 可选，2 栏对照放在下方，给「两种情况」用。

### compare — 并列（非对比）三栏

对比信息（A vs B）请一律用 `table`，本版式只用于「三栏并列但不构成对照」的场合。

```json
{"type": "compare", "panels": [
  {"tag": "假说一", "title": "细胞类型先分化", "color": "176FA5",
   "body": ["先出现类型，再出现区域差异。", "证据：跨区域类型高度重合。"]}
]}
```

`tag` 可选（小胶囊标签）；`body` 2–4 句，整组在卡片内垂直居中。

### figure — 大图

```json
{"type": "figure", "image": "fig/umap.png",
 "caption": "本分析：Fovea 与 Periphery 的 UMAP 分布",
 "side": [{"no": "01", "title": "分布", "body": ["两类样本按区域分开。"]}]}
```

有 `side` 时图占左 52%，无 `side` 时占满整宽。图片按比例缩放，不会拉伸。

### stats — 大数字

```json
{"type": "stats",
 "items": [{"value": "68", "label": "成年中央凹细胞类型"}],
 "blocks": [{"title": "主数据", "body": ["10x scRNA-seq 建立四组图谱。"]}]}
```

`items` 3–4 个；`blocks` 可选，0–2 个。

### closing — 结束页

```json
{"type": "closing", "left_week": "第十周组会",
 "message": "本周学习与文献汇报完毕，请老师批评指正！",
 "author": "陈诗琦", "date_text": "2026 年 11 月 20 日"}
```

## 行内标记

正文任意位置可用：

- `**加粗**` → 加粗、继承原色（强调关键名词）
- `!!标红!!` → 加粗、`#C0392B`（仅用于风险/反例，一页最多一处）

```json
"body": ["**17,064** 条 mRNA 与课程一致，!!但 X 占 0.27%!!，需追查阅读框。"]
```

## 图片路径

相对 `deck.json` 所在目录，或写绝对路径。建议同目录建 `fig/` 放图。
渲染器按原始长宽比缩放并居中，不会裁切。

## 最小可用样例

见 [`../examples/第四周_示例.json`](../examples/第四周_示例.json)（10 页，覆盖 8 种版式）。
复制它改内容，比从零写快。
