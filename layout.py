# -*- coding: utf-8 -*-
"""
layout.py — 自动布局兜底
仅当节点缺少 pos 时启用（pos 优先）。用层级算法（Sugiyama 风格简化版）算坐标。

适用：线性 / 简单分支流程图。
不适用：并列子列、多路汇聚、复杂回流的图 —— 那种请由 agent 显式给 pos。

算法：
  1. 拓扑排序算每个节点的 depth（从入度为0的根起算最长路径）
  2. 同 depth 的节点归一层，y 由 depth 决定（自上而下）
  3. 层内节点按出现顺序分配 x，整体相对中线居中
"""

# 布局参数（英寸）
BASE_Y = 9.0        # 第 0 层（根）的 y 坐标
LAYER_GAP = 1.6     # 层间距
NODE_GAP = 2.4      # 同层节点横向间距
CENTER_X = 4.0      # 默认中线 x


def _compute_depth(nodes, edges):
    """
    返回 {node_id: depth}。depth = 从任意根到该节点的最长路径。
    根 = 没有入边的节点。若有环，取拓扑序中的位置（保守）。
    """
    # 建邻接 + 入度
    ids = [n["id"] for n in nodes]
    out_adj = {i: [] for i in ids}
    in_deg = {i: 0 for i in ids}
    id_set = set(ids)
    for e in edges:
        f, t = e["from"], e["to"]
        if f in id_set and t in id_set:
            out_adj[f].append(t)
            in_deg[t] += 1

    # Kahn 拓扑排序
    from collections import deque
    queue = deque([i for i in ids if in_deg[i] == 0])
    topo = []
    indeg = dict(in_deg)
    while queue:
        n = queue.popleft()
        topo.append(n)
        for m in out_adj[n]:
            indeg[m] -= 1
            if indeg[m] == 0:
                queue.append(m)
    # 环内节点（没进 topo）追加到末尾
    for i in ids:
        if i not in topo:
            topo.append(i)

    # 按拓扑序算最长路径 depth
    depth = {i: 0 for i in ids}
    for n in topo:
        for m in out_adj[n]:
            if depth[m] < depth[n] + 1:
                depth[m] = depth[n] + 1
    return depth


def fill_missing_pos(nodes, edges):
    """
    原地修改 nodes：给缺 pos 的节点补 [x, y]。
    已有 pos 的不动（pos 优先）。
    返回是否做过补全。
    """
    missing = [n for n in nodes if "pos" not in n]
    if not missing:
        return False

    depth = _compute_depth(nodes, edges)

    # 按层分组（只针对缺 pos 的节点；有 pos 的不影响布局但占位）
    layers = {}
    for n in missing:
        d = depth.get(n["id"], 0)
        layers.setdefault(d, []).append(n)

    max_layer = max(layers.keys()) if layers else 0
    for d, layer_nodes in layers.items():
        count = len(layer_nodes)
        # 层内均分 x，相对中线对称
        total_w = (count - 1) * NODE_GAP
        start_x = CENTER_X - total_w / 2
        for i, n in enumerate(layer_nodes):
            x = start_x + i * NODE_GAP
            y = BASE_Y - d * LAYER_GAP
            n["pos"] = [x, y]
    return True
