---
name: visio-flowchart
description: "Generate editable Visio (.vsdx) flowcharts by driving the local Visio via COM automation. Convert natural-language process descriptions into a JSON spec, then render a native .vsdx with precise canvas cropping for direct Word insertion. Handles connector obstacle-avoidance, direction-aware connection-point distribution, and on-line labels. Triggers: Visio, 流程图, flowchart, vsdx, 业务流程, 判断分支, 泳道, process diagram."
license: MIT
metadata:
  version: "3.0"
  category: productivity
---

# Visio 流程图生成器

## 概述

本 skill 用 COM 自动化驱动**本机已安装的 Visio**，把流程描述生成为**原生可编辑的 `.vsdx`**。
引擎层（COM 操作、连线、裁剪）完全固化，agent 只需把用户的流程描述理解成一份结构化 JSON，引擎即可输出「画板精确贴边、可直接插入 Word」的 vsdx。

## 何时使用

| 场景 | 是否适用 |
|------|----------|
| 业务流程图、判断分支流程、算法流程 | ✅ 最佳场景 |
| 含开始/结束、输入输出、判断、处理的标准化流程图 | ✅ |
| 需要可编辑 .vsdx 交付（要在 Visio 里二次修改） | ✅ |
| 自由布局的架构图、神经网络结构图、LDM 那类空间自由的图 | ❌ 用 PPT / draw.io 更合适 |
| 用户的机器没装 Visio（或 agent 无法跑 shell） | ❌ COM 方案不可用 |

## 前置依赖

- **本机已安装 Visio**（COM 服务器 `Visio.Application`）
- **pywin32**：`pip install pywin32`
- 运行环境：**带桌面会话的 Windows**（SSH/无头环境 COM 会失败）

验证 COM 可用：`python -c "import win32com.client as w; v=w.Dispatch('Visio.Application'); print(v.Version); v.Quit()"`

## Quick Reference

| 项 | 值 |
|------|-------|
| **输出格式** | `.vsdx`（原生 Visio，可二次编辑） |
| **4 类形状** | `terminator`(椭圆/起止) `process`(矩形/处理) `decision`(菱形/判断) `data`(平行四边形/输入输出) |
| **形状→母版映射** | terminator→`Start/End`, process→`Process`, decision→`Decision`, data→`Data` |
| **颜色格式** | 6位 hex 字符串，如 `"2E75B6"`（引擎内部转 RGB；**勿用 #hex**） |
| **坐标单位** | 英寸（inch），原点左下，y 向上 |
| **画板裁剪** | 引擎自动按内容包围盒精确贴边，四周各留 0.1"，可直接插 Word |
| **虚线分组框** | `groups` 字段，画虚线框包住一组节点（区域划分/泳道） |
| **上下标渲染** | `subscript`/`superscript` 字段，渲染 ε_θ、z_T、σ² 等数学符号 |
| **PNG 导出** | `--png` 标志，生成 vsdx 时同步导出同名 PNG |
| **连线避障** | 无标签连线走自由路由，自动绕开中间面状元素 |
| **连接点均匀分布** | 同一边多条线连入时按方向均匀分布连接点，异向线分到边的两端不交点 |
| **标签在线上** | 带标签连线走直线 + 无填充文本框定位在路径上；自动扫描避让其他元素，标签底层仅有所属线 |
| **引擎入口** | `python generate.py <flow.json> -o <out.vsdx> --png` |

## 连线与标签机制（v3.0 制图质量规则）

引擎内置五条制图质量规则，agent 无需手动处理，但需理解其影响布局：

| 规则 | 引擎行为 | 对布局的要求 |
|------|----------|------------|
| **① 连线避障** | 无 `label` 的连线走自由路由（`ShapeRouteStyle=0` + `ConFixedCode=0` + 页面 `PageShapeSplit=1`），Visio 路由器自动绕开中间面状元素 | 无特殊要求 |
| **② 连边方向规则** | 每个端点按"线的另一端在哪"选边：对方在上方→top；对方在下方→bottom；平行水平位置→left/right 中最近边（\|dy\|/\|dx\|<0.5 视为平行，即偏离水平 27° 内） | 无特殊要求 |
| **③ 连接点方向感知均匀分布** | 同一条边多条线连入时，按对方端点坐标排序：来自上方的线连到边的上部、来自下方的连到下部，异向线分到边的两端不相交；同向线均匀分布 | 无特殊要求 |
| **④ 标签线走直线** | 有 `label` 的连线走直线（`ShapeRouteStyle=16`），路径=起终点直线可预测，标签能精确放线上 | 带标签的两节点间最好无其他形状遮挡（直线不绕障） |
| **⑤ 标签无填充+避让** | 标签是无填充无边框文本框，沿连线扫描（含法线偏移）找不与任何其他元素重叠的位置；标签底层只能有所属连线 | 判断框与汇聚点勿过近，否则标签无处可放（引擎会警告） |

> **布局黄金法则**：带标签的分支线（是/否、成功/失败）所连的两个节点之间，留出足够间距让标签有落脚处；判断框（菱形）下方汇聚点至少留 1.5" 以上纵向距离。若引擎打印 `[警告] 标签'X'全段被遮挡`，说明该处布局过密，需调大节点间距。

## 引擎文件

| 文件 | 作用 |
|------|------|
| [generate.py](generate.py) | 引擎入口：读 JSON → 布局 → COM 绘制 → 裁剪 → 输出 |
| [render.py](render.py) | 渲染函数库（cell 操作、形状母版、连线、裁剪、Visio 生命周期） |
| [layout.py](layout.py) | 自动布局兜底（缺 `pos` 时按层级算法算坐标） |
| [examples/aerial_qc.json](examples/aerial_qc.json) | 验证范例（航摄质检流程，含判断分支/汇聚/标签） |

---

## 工作流

### Step 1: 理解流程描述 → 结构化

把用户的流程描述理解成节点（nodes）和连线（edges）。
- 每个节点定：`id`、`type`（4类之一）、`text`、（可选）`pos`、`size`、`fill`、`text_color`、`font_size`、`bold`
- 每条线定：`from`、`to`、（可选）`label`、`route`、`dashed`

### Step 2: 布局（pos 优先，缺则自动）

- **简单线性流程**：可不写 `pos`，引擎的 `layout.py` 会按层级自动算坐标（自上而下，层间距 1.6"）
- **复杂图（并列子列、多路汇聚、并行分支）**：**必须由 agent 显式给每个节点 `pos`**（`[x, y]`），否则自动布局会乱

**布局原则**：
- 同一流程列的节点 x 相同，y 自上而下递减
- 判断框分出的「是/否」分支向左右散开
- 汇聚节点（多个分支汇入）放在分支的下方居中
- 节点间距：纵向 1.2~1.6"，横向 2.0~2.4"
- **判断框与下方汇聚点间距 ≥ 1.5"**：带标签的分支线（是/否）走直线，标签需落脚空间；过近会触发遮挡警告

### Step 3: 写 JSON

```json
{
  "title": "流程名（用作页名，可选）",
  "canvas_pad": 0.1,
  "nodes": [
    {"id": "start", "type": "terminator", "text": "开始", "fill": "70AD47", "text_color": "white", "pos": [4.0, 9.0]},
    {"id": "cond", "type": "decision", "text": "条件?", "fill": "FFD966", "pos": [4.0, 5.8], "size": [2.4, 1.0]},
    {"id": "yes", "type": "process", "text": "是分支", "pos": [2.0, 3.6]},
    {"id": "no", "type": "process", "text": "否分支", "pos": [6.0, 3.6]}
  ],
  "edges": [
    {"from": "start", "to": "cond"},
    {"from": "cond", "to": "yes", "label": "是"},
    {"from": "cond", "to": "no", "label": "否"}
  ]
}
```

### Step 4: 跑引擎

```bash
cd ~/.zcode/skills/visio-flowchart
python generate.py <你的flow.json> -o <输出.vsdx> --png
```

引擎会打印：放置形状数、连线数、分组框数、裁剪后的画板尺寸。
`--png` 可选，加上后同步导出同名 PNG（不写则只出 vsdx）。

---

## JSON Schema

### node 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `id` | ✅ | 唯一标识，edges 用它引用 |
| `type` | ✅ | `terminator`/`process`/`decision`/`data` |
| `text` | ✅ | 节点文字 |
| `pos` | ⚠️ | `[x, y]` 英寸。简单图可不填（自动布局）；复杂图必须填 |
| `size` | 可选 | `[w, h]` 英寸。decision 默认 2.4×1.0，其他按文字自适应 |
| `fill` | 可选 | 6位 hex，如 `"FFD966"`。缺省按 type 给默认色 |
| `text_color` | 可选 | `"white"` 表示白字（深色填充时用） |
| `font_size` | 可选 | pt，默认 9 |
| `bold` | 可选 | bool，decision 默认 true |
| `subscript` | 可选 | 下标区间列表 `[[start,end],...]`，字符索引基于已设好的 text。如 text=`"εθ"`、`subscript:[[1,2]]` → θ 为下标 |
| `superscript` | 可选 | 上标区间列表，格式同 subscript |

> **上下标技巧**：常见的上标（²³¹⁰⁰ 等）直接用 unicode 字符写入 text 更简单（无需 superscript 字段）；下标（i、t、θ、T 等）才用 `subscript` 字段经 COM 渲染。字符索引从 0 开始，基于最终写入的 text 字符串。

### edge 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `from` | ✅ | 起点节点 id |
| `to` | ✅ | 终点节点 id |
| `label` | 可选 | 连线标签（如 `"是"`/`"否"`）。带标签连线走直线，标签为**无填充无边框**文本框，落在路径上并自动避让其他元素——标签底层只能有所属连线。若引擎警告"全段被遮挡"，说明判断框与汇聚点过近，需调大节点间距 |
| `route` | 可选 | `"straight"` 强制直线 / `"orthogonal"` 右角 / `"free"` 自由绕障。默认：有标签走直线、无标签走绕障 |
| `dashed` | 可选 | bool，虚线 |

### group 字段（虚线分组框，可选）

顶层 `groups` 数组，每个元素画一个虚线框包住一组节点（用于区域划分/泳道）：

| 字段 | 必填 | 说明 |
|------|------|------|
| `pos` | ✅ | `[x0, y0, x1, y1]` 英寸，左下角(x0,y0)/右上角(x1,y1)。须手算包住目标节点 |
| `label` | 可选 | 框内标题文字 |
| `color` | 可选 | 虚线颜色 hex，默认 `"8FAADC"` |
| `weight` | 可选 | 线宽，默认 `"1.5pt"` |
| `font_size` | 可选 | 标题字号，默认 10 |

> 分组框在裁剪前绘制，其包围盒会纳入画板裁剪计算。分组框置于最底层，不遮挡节点。

---

## 坑清单（已固化为引擎逻辑，了解即可）

这些是开发过程中踩过、已在 render.py 里处理的坑。如果**引擎报错**或**修改引擎**，对照排查：

1. **COM 阻塞**：Visio 弹模态对话框会同步卡死 Python（曾 5 分钟超时）。引擎用 `AlertResponse=7` 自动拒绝所有弹窗。
2. **母版名**：中文版 Visio 的母版名是中文（如"开始/结束"不是"Terminator"）。引擎用 `Masters.ItemU(NameU)` 取通用名（`Start/End`/`Process`/`Decision`/`Data`）跨语言稳定。
3. **颜色格式**：Visio 的 FillForegnd 不认 `#RRGGBB`（报 `#NAME?`），必须 `RGB(r,g,b)`。引擎内部已转换，JSON 里写 hex 即可。
4. **连线边缘连接**：不能用 `GlueTo(PinX)` 连中心（多线挤一点、无法避障）。引擎用 `GlueToPos(shape,fx,fy)` 连边缘 + `ShapeRouteStyle`/`ConFixedCode=0` + 页面 `PageShapeSplit=1` 自由路由避障。
4b. **连边方向判定**：选边不能简单用 45° 对角线（`abs(dy)>=abs(dx)`）划分垂直/水平——斜上方/斜下方的连接会被误划成水平，连到错误的左右边。引擎用 `|dy|/|dx|<0.5`（偏离水平 27° 内）才判为平行水平，其余一律按上下方位走 top/bottom；两个端点各自独立按"对方在哪"选边。
5. **连接点无法 AddRow**：实测 `shape.AddRow` 给实例化形状加连接点会被 Visio 拒绝（"不允许此行类型的操作"，连接点 section 只在母版编辑层可改）。改用 `GlueToPos` 的 0~1 比例在边上任意位置连接，等价无限连接点且可均匀分布——见 `EdgeAllocator`。
6. **画板裁剪**：`AutoSizeDrawing()` 不彻底（按整页倍数留大片空白，无法插 Word）。引擎用包围盒计算+平移+设 PageWidth/Height，真正贴边。
7. **路径**：Visio COM 的 SaveAs 需要反斜杠 Windows 路径，引擎已用 `os.path.abspath` 规范化。
8. **进程泄漏**：必须 `try/finally` 调 `Quit()`，否则 Visio 进程残留。
9. **标签惰性路由**：动态连接器自带文字（`conn.Text`）的定位依赖 Visio 惰性路由，自动化生成/导出时位置不收敛，会脱离线条。引擎改用**独立无填充文本框**精确定位在连线起终点中点；带标签连线走直线使中点可预测。
10. **标签遮挡**：标签若压在分组框/其他节点上会歧义。引擎沿连线+法线扫描找无遮挡位置；紧凑布局（判断框紧贴汇聚点）找不到时打印警告，需调大间距。

---

## 当前版本范围（v3.0）

**核心能力**：
- 4 类形状（terminator/process/decision/data）、连线、分支标签
- 自动布局兜底（缺 `pos` 时按层级算坐标；pos 优先）
- 精确裁剪（画板贴边，可直接插 Word）
- 虚线分组框（`groups`）、上下标渲染（`subscript`/`superscript`）、PNG 导出（`--png`）

**制图质量机制（v3.0 新增）**：
- **连线避障**：无标签连线自由路由，自动绕开中间面状元素
- **连接点方向感知均匀分布**：同边多线按方向分到边的两端，异向线不相交
- **标签在线上**：带标签线走直线 + 无填充文本框，沿连线扫描避让其他元素，标签底层仅有所属线

**暂不支持**：多主题色板/样式预设、复杂图（并列子列/多汇聚）的自动布局优化（这类图请由 agent 显式给 pos）。

如需对已有 vsdx 单独导出 PNG（不走生成流程），用独立脚本：
```python
import win32com.client as wc
v = wc.Dispatch("Visio.Application"); v.Visible=False; v.AlertResponse=7
d = v.Documents.Open(r"<绝对路径.vsdx>"); v.ActivePage.Export(r"<绝对路径.png>"); d.Close(); v.Quit()
```

