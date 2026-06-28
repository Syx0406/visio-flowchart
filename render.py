# -*- coding: utf-8 -*-
"""
render.py — Visio 渲染函数库
原子级 COM 操作（形状/连线/裁剪/避障/标签），供 generate.py 调用。
所有函数接收/返回 COM 对象，不关心 JSON 结构。
"""
import win32com.client as wc

# BASFLO 模具路径：自动探测，兼容不同 Office 安装位置和语言版本
DEFAULT_STENCIL = None   # None → 运行时由 detect_stencil() 自动定位


def detect_stencil():
    """
    自动探测基本流程图模具 BASFLO_M.vssx 的路径。
    策略：遍历常见 Office 安装根目录 × 各语言 LCID 目录，找到第一个存在的文件。
    兼容：Program Files / (x86)、Office16 / Office15、中文2052 / 英文1033 等。
    找不到返回 None（调用方会报错提示）。
    """
    import os
    candidates_root = [
        r"C:\Program Files\Microsoft Office\root\Office16",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16",
        r"C:\Program Files\Microsoft Office\Office16",
        r"C:\Program Files (x86)\Microsoft Office\Office16",
        r"C:\Program Files\Microsoft Office\root\Office15",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15",
    ]
    for root in candidates_root:
        content = os.path.join(root, "Visio Content")
        if not os.path.isdir(content):
            continue
        # 扫描所有 LCID 子目录（如 2052/1033/1028...）
        for lcid in os.listdir(content):
            lcid_dir = os.path.join(content, lcid)
            if not os.path.isdir(lcid_dir):
                continue
            for fn in ("BASFLO_M.vssx", "BASFLO_M.VSSX"):
                p = os.path.join(lcid_dir, fn)
                if os.path.exists(p):
                    return p
    return None

# 节点 type → 模具母版 NameU（通用名，跨语言稳定）
MASTER_BY_TYPE = {
    "terminator": "Start/End",   # 椭圆：开始/结束
    "process":    "Process",      # 矩形：处理
    "decision":   "Decision",     # 菱形：判断
    "data":       "Data",         # 平行四边形：输入/输出
}

# 默认填充色（无 fill 时用）
DEFAULT_FILL = {
    "terminator": "70AD47",   # 绿
    "process":    "9DC3E6",   # 浅蓝
    "decision":   "FFD966",   # 黄
    "data":       "9DC3E6",   # 浅蓝
}

# 连线路由：16=直线，1=右角正交
ROUTE_STRAIGHT = "16"
ROUTE_ORTHOGONAL = "1"

DEFAULT_LINE_COLOR = "595959"


# ============ 原子 cell 操作 ============

def _rgb(hexv):
    """hex(如 '2E75B6') → Visio 的 'RGB(r,g,b)' 公式。注意：不能用 #hex，会报 #NAME?"""
    r = int(hexv[0:2], 16)
    g = int(hexv[2:4], 16)
    b = int(hexv[4:6], 16)
    return f"RGB({r},{g},{b})"


def set_fill(shape, hex_rgb):
    shape.CellsU("FillForegnd").FormulaU = _rgb(hex_rgb)


def set_line(shape, weight="1pt", color=DEFAULT_LINE_COLOR, dashed=False):
    shape.CellsU("LinePattern").FormulaU = "2" if dashed else "1"
    shape.CellsU("LineWeight").FormulaU = weight
    shape.CellsU("LineColor").FormulaU = _rgb(color)


def no_line(shape):
    shape.CellsU("LinePattern").FormulaU = "0"


def set_font(shape, size=9, bold=False, white=False):
    shape.CellsU("Char.Size").FormulaU = f"{size}pt"
    if bold:
        shape.CellsU("Char.Style").FormulaU = "1"  # 1=bold
    if white:
        shape.CellsU("Char.Color").FormulaU = "RGB(255,255,255)"


# 上下标：visCharacterPos 值，1=上标, 2=下标
_POS_SUPER = 1
_POS_SUB = 2


def apply_script(shape, segments):
    """
    给 shape 的文字应用上标/下标。
    segments: [{"start":int, "end":int, "type":"sub"|"super"}, ...]
    start/end 是字符索引（基于已设置的 Text）。必须先设好 shape.Text 再调用。
    """
    for seg in segments:
        stype = _POS_SUB if seg["type"] == "sub" else _POS_SUPER
        try:
            chs = shape.Characters
            chs.Begin = seg["start"]
            chs.End = seg["end"]
            chs.CharProps(4, stype)   # 4 = visCharacterPos
        except Exception:
            pass


# ============ 形状/连线 ============

def draw_node(page, stencil, node):
    """
    按 node 字典放一个形状。
    node 字段：type, text, pos=[x,y], size=[w,h], fill, text_color(white/None),
               font_size, bold, route 仅连线用
    返回 COM shape。
    """
    ntype = node["type"]
    master_name = MASTER_BY_TYPE.get(ntype)
    if master_name is None:
        raise ValueError(f"未知节点 type: {ntype}（支持: {list(MASTER_BY_TYPE)})")

    x, y = node["pos"]
    master = stencil.Masters.ItemU(master_name)
    shape = page.Drop(master, x, y)

    # 尺寸（decision 默认给大一点）
    if "size" in node:
        w, h = node["size"]
        shape.CellsU("Width").FormulaU = f"{w} in"
        shape.CellsU("Height").FormulaU = f"{h} in"
    elif ntype == "decision":
        shape.CellsU("Width").FormulaU = "2.4 in"
        shape.CellsU("Height").FormulaU = "1.0 in"

    # 填充
    fill = node.get("fill", DEFAULT_FILL.get(ntype, "9DC3E6"))
    set_fill(shape, fill)

    # 边线
    set_line(shape, weight=node.get("line_weight", "1pt"))

    # 文字
    if node.get("text"):
        shape.Text = node["text"]
        set_font(
            shape,
            size=node.get("font_size", 9),
            bold=node.get("bold", ntype == "decision"),
            white=(node.get("text_color") == "white"),
        )
        # 上标/下标（text 必须先设好）
        if node.get("subscript") or node.get("superscript"):
            segs = []
            for s in node.get("subscript", []):
                segs.append({"start": s[0], "end": s[1], "type": "sub"})
            for s in node.get("superscript", []):
                segs.append({"start": s[0], "end": s[1], "type": "super"})
            apply_script(shape, segs)
    return shape


def draw_group(page, group):
    """
    画一个虚线分组框（无填充），置于最底层。
    group 字段：pos=[x0,y0,x1,y2]（左下x0,y0 / 右上x1,y1），label?(可选标题)
    返回 COM shape。
    """
    x0, y0, x1, y1 = group["pos"]
    g = page.DrawRectangle(x0, y0, x1, y1)
    g.CellsU("FillPattern").FormulaU = "0"        # 无填充
    g.CellsU("LinePattern").FormulaU = "2"        # 虚线
    g.CellsU("LineWeight").FormulaU = group.get("weight", "1.5pt")
    g.CellsU("LineColor").FormulaU = _rgb(group.get("color", "8FAADC"))
    if group.get("label"):
        g.Text = group["label"]
        set_font(g, size=group.get("font_size", 10), bold=True)
    g.SendToBack()
    return g


def _pick_edges(src_shape, dst_shape):
    """
    根据 src→dst 的相对方位，决定起边和终边（让连线顺着流向、减少穿越）。
    返回 (src_edge, dst_edge)，值为 'bottom'/'top'/'left'/'right'。
    """
    sx = src_shape.CellsU("PinX").ResultIU
    sy = src_shape.CellsU("PinY").ResultIU
    tx = dst_shape.CellsU("PinX").ResultIU
    ty = dst_shape.CellsU("PinY").ResultIU
    # 用形状宽高修正"上方/下方/平行"判定，使判定基于真实边界而非中心
    sw = abs(src_shape.CellsU("Width").ResultIU)
    sh = abs(src_shape.CellsU("Height").ResultIU)
    dw = abs(dst_shape.CellsU("Width").ResultIU)
    dh = abs(dst_shape.CellsU("Height").ResultIU)
    return _pick_edges_xy(sx, sy, tx, ty, (sw, sh), (dw, dh))


# 水平平行判定阈值：|dy| / |dx| < _HORIZ_RATIO 视为"水平平行位置"
_HORIZ_RATIO = 0.5   # 偏离水平方向 27° 以内算平行


def _pick_edges_xy(sx, sy, tx, ty, src_size=None, dst_size=None):
    """
    坐标版选边逻辑，供 EdgeAllocator 在 shape 未创建时用 node pos 预算。

    连边方向规则（确定 src 和 dst 各用哪条边）：
      对每个端点，看"线的另一端"相对自己的位置：
        ① 另一端在上方 → 用 top 边（上方来的线连上面的边）
        ③ 另一端在下方 → 用 bottom 边（下方来的线从下面的边连；规则③亦允许两侧）
        ② 另一端在平行水平位置 → 用 left/right 中距离最近的那条边
      "平行水平"判定：|dy|/|dx| < _HORIZ_RATIO（偏离水平 27° 以内）
    """
    dx_ = tx - sx
    dy_ = ty - sy
    # 两个端点各自选边
    def edge_for(self_xy, other_xy):
        ex_, ey_ = other_xy[0] - self_xy[0], other_xy[1] - self_xy[1]
        if abs(ex_) > 1e-6 and abs(ey_) / abs(ex_) < _HORIZ_RATIO:
            # 平行水平：左右最近边
            return "right" if ex_ > 0 else "left"
        # 垂直方向：y 大=上，y 小=下
        if ey_ > 0:
            return "top"       # 对方在上方(y大) → 顶边（规则①）
        elif ey_ < 0:
            return "bottom"    # 对方在下方(y小) → 底边（规则③）
        else:
            return "top"       # 完全重叠，默认顶边
    src_edge = edge_for((sx, sy), (tx, ty))
    dst_edge = edge_for((tx, ty), (sx, sy))
    return (src_edge, dst_edge)


# 边 → (主轴 fx, 主轴 fy) 的基础锚点（沿边的中心），垂直边主轴是 y，水平边主轴是 x
_EDGE_AXIS = {
    "top":    ("x", 1.0),   # 顶边，沿 x 分布，fy=1
    "bottom": ("x", 0.0),   # 底边，沿 x 分布，fy=0
    "left":   ("y", 0.0),   # 左边，沿 y 分布，fx=0
    "right":  ("y", 1.0),   # 右边，沿 y 分布，fx=1
}


class EdgeAllocator:
    """
    连接点分配器：同一条边上多条线时，连接点均匀分布且按方向排序。
    关键规则：同一边上，来自上方的线分到边的上部、来自下方的线分到边的下部，
    使异向线分到边的两端、不交点；同向线均匀分布在对应半区。
    解决"同边异向线在边附近交叉"的问题。
    用法：
      alloc = EdgeAllocator()
      alloc.preserve(shapes, edges)        # 预统计每边占用 + 方向排序
      pos = alloc.next_pos(shape_id, edge, other_xy, axis)  # 按方向序号分配
    """
    def __init__(self):
        # (shape_id, edge) -> 该边上所有连线的对方主轴坐标列表（待排序）
        self._pending = {}
        # (shape_id, edge) -> 排序后的坐标列表（决定分配序）
        self._ordered = {}
        # (shape_id, edge) -> 已分配计数
        self._used = {}

    def preserve(self, shapes, edges):
        """
        预统计：对每个 (节点, 边)，按"对方端点主轴坐标"排序，算出每条线
        （按 edges 顺序）在该边上的最终 frac，存入 _slot_frac。
        connect 按 edges 相同顺序调用 next_pos，依次取出，保证方向一致。
        """
        def get_xy(s):
            try:
                return (s.CellsU("PinX").ResultIU, s.CellsU("PinY").ResultIU)
            except Exception:
                return (s["pos"][0], s["pos"][1])

        # 收集每条边上的线及其对方坐标 + 在 edges 中的原始顺序
        edge_lines = {}   # (node, edge) -> [(other_coord, edge_index)]
        for ei, e in enumerate(edges):
            f, t = e["from"], e["to"]
            sf, df = shapes.get(f), shapes.get(t)
            if sf is None or df is None:
                continue
            fx, fy = get_xy(sf)
            tx, ty = get_xy(df)
            src_edge, dst_edge = _pick_edges_xy(fx, fy, tx, ty)
            # src 端
            s_axis = _EDGE_AXIS[src_edge][0]
            s_other = tx if s_axis == "x" else ty
            edge_lines.setdefault((f, src_edge), []).append((s_other, ei))
            # dst 端
            d_axis = _EDGE_AXIS[dst_edge][0]
            d_other = fx if d_axis == "x" else fy
            edge_lines.setdefault((t, dst_edge), []).append((d_other, ei))

        # 对每条边：按对方坐标排序，算出每条线的 frac，按 edge_index 存回
        # _slot_frac[(node,edge)][edge_index] = frac
        self._slot_frac = {}
        for key, items in edge_lines.items():
            total = len(items)
            # 按对方坐标升序排，相同坐标按 edge_index 稳定排序
            items_sorted = sorted(items, key=lambda x: (x[0], x[1]))
            for slot_i, (coord, ei) in enumerate(items_sorted):
                frac = (slot_i + 1) / (total + 1)
                self._slot_frac.setdefault(key, {})[ei] = frac
        self._used = {}

    def next_pos(self, shape_id, edge, other_xy=None, edge_index=None):
        """
        返回 (fx, fy)。需传 edge_index（该线在 edges 中的序号）以取出预计算的 frac。
        方向感知：preserve 已按对方坐标排序，故来自同方向的线 frac 相邻、异向分两端。
        """
        key = (shape_id, edge)
        if edge_index is not None and key in self._slot_frac:
            frac = self._slot_frac[key].get(edge_index, 0.5)
        else:
            # 退化：按调用顺序均匀
            used = self._used.get(key, 0)
            total = len(self._slot_frac.get(key, {})) or 1
            frac = (used + 1) / (total + 1)
            self._used[key] = used + 1
        edge_axis, base = _EDGE_AXIS[edge]
        if edge_axis == "x":
            return (frac, base)
        else:
            return (base, frac)


def connect(page, dyn_master, src_shape, dst_shape, label=None, route=None, dashed=False,
            label_pos="mid", src_id=None, dst_id=None, allocator=None, edge_index=None):
    """
    用 Dynamic connector 连 src→dst，连线自动绕开中间的面状元素。
    关键机制（解决"连线穿越其他形状"问题）：
      - GlueToPos 连到形状【边缘】而非中心，给路由器绕障空间
      - 智能选边：按相对方位自动定起/终边，顺流向
      - 同边多线时经 allocator 均匀分布连接点，避免带文字线段重叠
      - ConFixedCode=0 + ShapeRouteStyle=0：自由路由，Visio 路由器自动绕障
      - 依赖页面 PageShapeSplit=1（在 open_visio 中已设）
    route: None(默认避障) / "straight"(直线，不避障) / "orthogonal"(右角)
    label_pos: 标签位置 "start"/"mid"(默认)/"end"
    src_id/dst_id/allocator: 传入时启用连接点均匀分配
    返回 connector COM 对象。
    """
    conn = page.Drop(dyn_master, 0, 0)

    # 边缘连接点（0~1 比例）：底中/顶中/左中/右中
    edge_pos = {"bottom": (0.5, 0), "top": (0.5, 1), "left": (0, 0.5), "right": (1, 0.5)}

    # 路由策略：
    #   - 显式 route 参数优先
    #   - 否则：有标签(label)的线走直线(straight)，路径可预测→标签能精确放线上
    #          无标签的走避障(free)，自动绕开面状元素
    src_edge, dst_edge = _pick_edges(src_shape, dst_shape)
    if allocator is not None and src_id is not None and dst_id is not None:
        sx, sy = allocator.next_pos(src_id, src_edge, edge_index=edge_index)
        dx_, dy_ = allocator.next_pos(dst_id, dst_edge, edge_index=edge_index)
    else:
        sx, sy = edge_pos[src_edge]
        dx_, dy_ = edge_pos[dst_edge]
    conn.CellsU("BeginX").GlueToPos(src_shape, sx, sy)
    conn.CellsU("EndX").GlueToPos(dst_shape, dx_, dy_)

    if route == "orthogonal":
        conn.CellsU("ShapeRouteStyle").FormulaU = ROUTE_ORTHOGONAL
    elif route == "straight" or (route is None and label):
        # 有标签→直线：路径=起终点直线，中点可精确计算，标签必然在线上
        conn.CellsU("ShapeRouteStyle").FormulaU = ROUTE_STRAIGHT
    else:
        # 无标签→自由路由避障
        conn.CellsU("ShapeRouteStyle").FormulaU = "0"
        try:
            conn.CellsU("ConFixedCode").FormulaU = "0"
        except Exception:
            pass

    conn.CellsU("LineWeight").FormulaU = "1pt"
    conn.CellsU("LineColor").FormulaU = _rgb(DEFAULT_LINE_COLOR)
    conn.CellsU("EndArrow").FormulaU = "4"   # 箭头
    if dashed:
        conn.CellsU("LinePattern").FormulaU = "2"
    if label:
        # 标签做成独立【无填充无边框】文本框，落在连线上。
        # 要求标签底层只能有所属连线：_add_label_box 会沿连线扫描，
        # 选不与任何其他元素重叠的位置；找不到则警告需调整布局。
        _add_label_box(page, conn, label, label_pos)
    return conn


def _add_label_box(page, conn, text, label_pos="mid"):
    """
    在连线 conn 上放一个【无填充、无边框】的标签文本框。
    要求：标签底层只能是所属连接线，不能压在其他元素上。
    做法：沿连线扫描多个候选位置，选第一个不与任何其他元素重叠的点；
         找不到则落在中点并打印警告（提示需调整布局）。
    """
    bx = conn.CellsU("BeginX").ResultIU
    by = conn.CellsU("BeginY").ResultIU
    ex = conn.CellsU("EndX").ResultIU
    ey = conn.CellsU("EndY").ResultIU

    # 标签尺寸
    w = max(0.6, len(text) * 0.18)
    h = 0.32

    # 收集页面上除本连线外的所有元素包围盒（用于遮挡检测）
    obstacles = []
    conn_id = id(conn)
    for s in page.Shapes:
        if id(s) == conn_id:
            continue
        try:
            sx = s.CellsU("PinX").ResultIU
            sy = s.CellsU("PinY").ResultIU
            sw = s.CellsU("Width").ResultIU
            sh = s.CellsU("Height").ResultIU
            obstacles.append((sx - sw/2, sy - sh/2, sx + sw/2, sy + sh/2))
        except Exception:
            pass

    def overlaps_any(cx, cy):
        """标签中心(cx,cy)的框是否与任何障碍重叠。"""
        lx0, ly0 = cx - w/2, cy - h/2
        lx1, ly1 = cx + w/2, cy + h/2
        for ox0, oy0, ox1, oy1 in obstacles:
            if lx0 < ox1 and lx1 > ox0 and ly0 < oy1 and ly1 > oy0:
                return True
        return False

    # 沿连线扫描候选位置：优先 label_pos 指定的段，再扫整条线找空位
    candidates = []
    if label_pos == "start":
        candidates = [0.2, 0.15, 0.25, 0.3, 0.1, 0.35, 0.4, 0.45, 0.5]
    elif label_pos == "end":
        candidates = [0.8, 0.85, 0.75, 0.7, 0.9, 0.65, 0.6, 0.55, 0.5]
    else:
        candidates = [0.5, 0.4, 0.6, 0.3, 0.7, 0.2, 0.8, 0.15, 0.85]
    # 连线方向单位向量 + 法线方向（用于必要时沿法线偏移）
    import math as _m
    seg_dx, seg_dy = ex - bx, ey - by
    seg_len = _m.hypot(seg_dx, seg_dy) or 1.0
    ux, uy = seg_dx / seg_len, seg_dy / seg_len          # 沿线方向
    nx, ny = -uy, ux                                      # 法线方向（垂直于线）

    chosen = None   # (mx, my)
    # 第一轮：沿线段扫描，不偏移
    for t in candidates:
        cx = bx + seg_dx * t
        cy = by + seg_dy * t
        if not overlaps_any(cx, cy):
            chosen = (cx, cy)
            break

    # 第二轮：沿线段各点 + 法线方向小幅偏移（让标签滑出菱形等包围盒的尖角）
    if chosen is None:
        offsets = [0.0, 0.22, -0.22, 0.4, -0.4, 0.6, -0.6]
        for t in candidates:
            for off in offsets:
                cx = bx + seg_dx * t + nx * off
                cy = by + seg_dy * t + ny * off
                if not overlaps_any(cx, cy):
                    chosen = (cx, cy)
                    break
            if chosen:
                break

    if chosen is None:
        # 全部尝试都被遮挡 → 落中点并警告（需调整布局）
        chosen = (bx + seg_dx * 0.5, by + seg_dy * 0.5)
        print(f"[警告] 标签'{text}'所在连线全段被遮挡，需调整布局（节点坐标/连线走向）")

    mx, my = chosen
    lbl = page.DrawRectangle(mx - w/2, my - h/2, mx + w/2, my + h/2)
    lbl.Text = text
    lbl.CellsU("LinePattern").FormulaU = "0"          # 无边框
    lbl.CellsU("FillPattern").FormulaU = "0"          # 无填充（透明，透出连线）
    lbl.CellsU("Char.Size").FormulaU = "9pt"
    lbl.CellsU("Char.Style").FormulaU = "1"           # bold
    # 置顶：确保标签在最上层，其下只能是所属连线（已被遮挡检测保证无其他元素）
    try:
        lbl.BringToFront()
    except Exception:
        pass


# ============ 精确裁剪（替代不彻底的 AutoSizeDrawing）============

def crop_to_content(page, pad=0.1):
    """
    把画板精确裁剪到内容包围盒，四周留 pad 英寸。
    解决 AutoSizeDrawing 留大片空白、无法直接插 Word 的问题。
    """
    # 1. 算所有形状的包围盒
    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")
    for s in page.Shapes:
        try:
            pinx = s.CellsU("PinX").ResultIU
            piny = s.CellsU("PinY").ResultIU
            w = s.CellsU("Width").ResultIU
            h = s.CellsU("Height").ResultIU
            if w < 0.01 or h < 0.01:
                continue
            min_x = min(min_x, pinx - w / 2)
            min_y = min(min_y, piny - h / 2)
            max_x = max(max_x, pinx + w / 2)
            max_y = max(max_y, piny + h / 2)
        except Exception:
            pass
    if min_x == float("inf"):
        return  # 空页

    content_w = max_x - min_x
    content_h = max_y - min_y

    # 2. 平移所有形状，包围盒左下角对齐到 (pad, pad)
    dx = pad - min_x
    dy = pad - min_y
    for s in page.Shapes:
        try:
            old_x = s.CellsU("PinX").ResultIU
            old_y = s.CellsU("PinY").ResultIU
            s.CellsU("PinX").FormulaU = f"{old_x + dx} in"
            s.CellsU("PinY").FormulaU = f"{old_y + dy} in"
        except Exception:
            pass

    # 3. 设画板尺寸 + 页边距 0
    ps = page.PageSheet
    ps.CellsU("PageWidth").FormulaU = f"{content_w + 2 * pad} in"
    ps.CellsU("PageHeight").FormulaU = f"{content_h + 2 * pad} in"
    for cell in ("PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin"):
        ps.CellsU(cell).FormulaU = "0 in"

    return content_w, content_h


# ============ Visio 生命周期 ============

def open_visio(stencil_path=None):
    """
    启动 Visio，返回 (visio_app, doc, stencil, page, dyn_master)。
    固化关键坑：
      - AlertResponse=7：自动拒绝弹窗，防止 COM 同步阻塞（5分钟超时的元凶）
      - Visible=False：不弹窗，但仍在桌面会话运行（COM 必须）
    """
    if stencil_path is None:
        # 自动探测模具路径（兼容不同 Office 安装位置/语言）
        stencil_path = DEFAULT_STENCIL or detect_stencil()
        if stencil_path is None:
            raise FileNotFoundError(
                "未找到 BASFLO_M.vssx 模具。请确认本机已安装 Visio，"
                "或在调用 open_visio(stencil_path=...) 时显式传入模具路径。"
            )
    visio = wc.Dispatch("Visio.Application")
    visio.Visible = False
    try:
        visio.AlertResponse = 7   # 6=Yes 7=No —— 自动拒绝所有模态弹窗
    except Exception:
        pass
    doc = visio.Documents.Add("")
    stencil = visio.Documents.OpenEx(stencil_path, 64)  # 64=visOpenDocked
    page = visio.ActivePage
    # 页面级开启避障分割：让连接器路由器自动绕开中间的面状元素
    try:
        page.PageSheet.CellsU("PageShapeSplit").FormulaU = "1"
    except Exception:
        pass
    dyn_master = stencil.Masters.ItemU("Dynamic connector")
    return visio, doc, stencil, page, dyn_master


def close_visio(visio, doc):
    """安全关闭：先关文档再退 Visio，finally 兜底 Quit 防进程泄漏。"""
    if doc is not None:
        try:
            doc.Close()
        except Exception:
            pass
    try:
        visio.Quit()
    except Exception:
        pass
