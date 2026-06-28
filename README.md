# visio-flowchart

> 用 COM 自动化驱动本机 Visio，把流程描述（JSON）生成为**原生可编辑的 `.vsdx`**。
> 画板精确贴边、连接线避障、连接点方向感知均匀分布、标签无填充贴线。
> 生成的 vsdx 可直接插入 Word，也可在 Visio 里二次编辑。

---

## ⚡ 一键安装（把本仓库地址交给你的智能体）

把这个仓库地址复制给你的智能体（Claude Code / ZCode / Cursor 等），让它执行下面的安装。智能体读到这里即可自动完成：

> **告诉你的智能体**：「请安装 visio-flowchart skill：`https://github.com/Syx0406/visio-flowchart`」

**安装命令**（智能体执行 / 手动执行均可）：

```bash
# 1. clone 到 skills 目录（按你用的工具选一个）
git clone https://github.com/Syx0406/visio-flowchart.git ~/.zcode/skills/visio-flowchart      # ZCode
# 或
git clone https://github.com/Syx0406/visio-flowchart.git ~/.claude/skills/visio-flowchart     # Claude Code

# 2. 装依赖
pip install pywin32

# 3. 验证（应输出模具路径 + Visio 版本号）
cd ~/.zcode/skills/visio-flowchart    # 换成你 clone 的路径
python -c "import render; print('模具:', render.detect_stencil())"
python -c "import win32com.client as w; v=w.Dispatch('Visio.Application'); print('Visio:', v.Version); v.Quit()"

# 4. 跑示例验证
python generate.py examples/leap_year.json -o test.vsdx --png
```

安装成功后**新开一个会话**，即可用 `/visio-flowchart` 调用，或直接描述流程需求。

> 模具路径**自动探测**，兼容 Program Files / (x86)、Office16/15、中文2052/英文1033 等各版本，无需手动配置。

---

## 适用场景

- ✅ 业务流程图、判断分支流程、算法流程
- ✅ 需要可编辑 `.vsdx` 交付
- ❌ 自由布局的架构图 / 神经网络结构图（用 PPT / draw.io）
- ❌ 机器没装 Visio 或非 Windows（COM 方案不可用）

## 前提条件

| 条件 | 说明 |
|------|------|
| **Windows** | COM 自动化需要带桌面会话的 Windows（SSH/无头容器不行） |
| **已安装 Visio** | 需带 `BASFLO` 基本流程图模具（Visio 标配） |
| **Python 3.8+** | 推荐 3.10+ |
| **pywin32** | `pip install pywin32` |

## 作为独立工具使用（不用 skill 也能用）

把目录 clone 到任意位置，命令行调用：

```bash
cd visio-flowchart
python generate.py 流程.json -o 输出.vsdx --png
```

## 安装验证

```bash
python -c "import render; print('模具:', render.detect_stencil())"
python -c "import win32com.client as w; v=w.Dispatch('Visio.Application'); print('Visio:', v.Version); v.Quit()"
```

两条都能正常输出即可用。

## 快速开始

**1. 写流程 JSON**（最小示例）：

```json
{
  "title": "闰年判断",
  "canvas_pad": 0.1,
  "nodes": [
    {"id": "start", "type": "terminator", "text": "开始", "fill": "70AD47", "text_color": "white", "pos": [4.0, 9.0]},
    {"id": "input", "type": "data", "text": "输入年份", "fill": "4472C4", "text_color": "white", "pos": [4.0, 7.6]},
    {"id": "cond", "type": "decision", "text": "(year%4==0 && year%100!=0) ‖ (year%400==0)", "fill": "ED7D31", "text_color": "white", "pos": [4.0, 5.8], "size": [3.2, 1.2]},
    {"id": "yes", "type": "data", "text": "输出\"是闰年\"", "fill": "4472C4", "text_color": "white", "pos": [2.2, 3.6]},
    {"id": "no", "type": "data", "text": "输出\"不是闰年\"", "fill": "4472C4", "text_color": "white", "pos": [5.8, 3.6]},
    {"id": "end", "type": "terminator", "text": "结束", "fill": "70AD47", "text_color": "white", "pos": [4.0, 1.8]}
  ],
  "edges": [
    {"from": "start", "to": "input"},
    {"from": "input", "to": "cond"},
    {"from": "cond", "to": "yes", "label": "是"},
    {"from": "cond", "to": "no", "label": "否"},
    {"from": "yes", "to": "end"},
    {"from": "no", "to": "end"}
  ]
}
```

> 简单线性流程的节点可省略 `pos`，引擎自动按层级布局；复杂图（并列/汇聚）必须显式给 `pos`。

**2. 生成**：

```bash
python generate.py leap_year.json -o leap_year.vsdx --png
```

**3. 完成**：得到 `leap_year.vsdx`（可 Visio 编辑、可插 Word）+ `leap_year.png`。

## 命令行参数

```
python generate.py <flow.json> [-o 输出.vsdx] [--pad 0.1] [--png]
```

| 参数 | 说明 |
|------|------|
| `<flow.json>` | 流程描述 JSON（必填） |
| `-o, --output` | 输出 vsdx 路径（默认与 JSON 同名） |
| `--pad` | 画板四周留白英寸，默认 0.1（设 0 为绝对贴边） |
| `--png` | 同时导出同名 PNG |

## JSON Schema 摘要

- **nodes**：`id` / `type`(terminator\|process\|decision\|data) / `text` / `pos`([x,y]英寸) / `size` / `fill`(hex) / `text_color`("white"或省) / `subscript`([[start,end],...]) / `superscript`
- **edges**：`from` / `to` / `label` / `route`(straight\|orthogonal\|free) / `dashed`
- **groups**（可选）：`pos`([x0,y0,x1,y1]) / `label` / `color` / `dashed` —— 虚线分组框

完整字段说明见 [SKILL.md](SKILL.md) 的「JSON Schema」章节。

## 制图质量机制（自动生效，无需配置）

1. **连线避障**：无标签连线自动绕开中间面状元素
2. **连边方向规则**：上方要素连顶边、下方连底边、平行连最近侧边
3. **连接点方向感知均匀分布**：同边多线按方向分到边的两端，异向不相交
4. **标签线走直线**：带标签连线走直线，标签可精确放线上
5. **标签无填充避让**：标签无填充无边框，沿连线扫描避让其他元素，底层仅所属线

## 文件说明

| 文件 | 作用 |
|------|------|
| `generate.py` | 引擎入口（命令行） |
| `render.py` | 渲染函数库（COM 操作、避障、裁剪、模具自动探测） |
| `layout.py` | 自动布局兜底（缺 `pos` 时按层级算坐标） |
| `SKILL.md` | 完整文档（agent 调用指南，含全部字段/坑清单） |
| `examples/` | 范例 JSON（aerial_qc 复杂流程 / quality_eval 分组框+上下标） |

## 常见问题

- **找不到模具**：报 `未找到 BASFLO_M.vssx` → 确认装了 Visio；或 `open_visio(stencil_path=你的路径)`
- **Visio 弹窗卡死**：引擎已用 `AlertResponse=7` 自动拒绝弹窗；若仍卡，确认没用 SSH 远程跑
- **标签遮挡警告**：`[警告] 标签'X'全段被遮挡` → 判断框与汇聚点过近，调大节点间距
- **中文母版名**：引擎用 `Masters.ItemU(NameU)` 取通用名，中英文版 Visio 都兼容

## 许可证

MIT
