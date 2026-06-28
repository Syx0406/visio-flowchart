# -*- coding: utf-8 -*-
"""
generate.py — Visio 流程图生成引擎入口

用法：
  python generate.py <flow.json> [-o output.vsdx] [--pad 0.1]

引擎黑盒：读 JSON → 布局(pos优先,缺则自动) → COM 绘制 → 精确裁剪 → 输出 vsdx
agent 只需产出符合 schema 的 JSON，无需关心 Visio cell 操作。
"""
import os
import sys
import json
import argparse

import render
import layout


def generate(flow, output_path, pad=0.1, export_png=False):
    """
    flow: dict，schema:
      {
        "title": str(可选, 用作页名),
        "canvas_pad": float(可选, 覆盖默认 pad),
        "nodes": [{"id","type","text","pos"?,...}],
        "edges": [{"from","to","label"?,"route"?,"dashed"?}],
        "groups": [{"pos":[x0,y0,x1,y1],"label"?,"color"?}]   (可选, 虚线分组框)
      }
    output_path: 输出 .vsdx 路径
    export_png: True 时同时导出同名 .png
    """
    nodes = flow["nodes"]
    edges = flow.get("edges", [])
    groups = flow.get("groups", [])
    pad = flow.get("canvas_pad", pad)

    # ① 布局：补全缺 pos 的节点（pos 优先）
    did_layout = layout.fill_missing_pos(nodes, edges)
    if did_layout:
        print("[布局] 已自动补全缺 pos 的节点坐标")

    # ② 校验：所有节点必须有 pos 和合法 type
    valid_types = set(render.MASTER_BY_TYPE.keys())
    for n in nodes:
        if "pos" not in n:
            raise ValueError(f"节点 {n.get('id')} 无 pos 且布局失败")
        if n["type"] not in valid_types:
            raise ValueError(f"节点 {n['id']} type 非法: {n['type']}（支持: {valid_types}）")

    # ③ COM 绘制
    visio, doc, stencil, page, dyn = render.open_visio()
    try:
        if flow.get("title"):
            try:
                page.Name = flow["title"]
            except Exception:
                pass

        # 放形状
        shapes = {}
        for n in nodes:
            shapes[n["id"]] = render.draw_node(page, stencil, n)
        print(f"[绘制] 放置 {len(shapes)} 个形状")

        # 连线（先预统计连接点占用，再均匀分配，避免同边多线重叠）
        id_set = set(shapes.keys())
        valid_edges = []
        for e in edges:
            f, t = e["from"], e["to"]
            if f in id_set and t in id_set:
                valid_edges.append(e)
            else:
                print(f"[警告] 连线 {f}->{t} 引用了不存在的节点，跳过")
        # 预统计每边占用数，使 next_pos 能均匀分配
        allocator = render.EdgeAllocator()
        allocator.preserve(shapes, valid_edges)
        n_edges = 0
        for ei, e in enumerate(valid_edges):
            f, t = e["from"], e["to"]
            render.connect(
                page, dyn, shapes[f], shapes[t],
                label=e.get("label"),
                route=e.get("route"),
                dashed=e.get("dashed", False),
                label_pos=e.get("label_pos", "mid"),
                src_id=f, dst_id=t, allocator=allocator, edge_index=ei,
            )
            n_edges += 1
        print(f"[绘制] 连接 {n_edges} 条线")

        # 分组框（虚线框，置于最底层；须在裁剪前画，使包围盒纳入裁剪计算）
        for g in groups:
            render.draw_group(page, g)
        if groups:
            print(f"[绘制] 放置 {len(groups)} 个分组框")

        # ④ 精确裁剪（替代不彻底的 AutoSizeDrawing）
        result = render.crop_to_content(page, pad=pad)
        if result:
            cw, ch = result
            print(f"[裁剪] 内容 {cw:.2f} x {ch:.2f}\" -> 画板 {cw+2*pad:.2f} x {ch+2*pad:.2f}\" (四向空白各 {pad}\")")

        # ⑤ 输出（Visio COM 需要反斜杠路径）
        output_path = os.path.abspath(output_path)
        if os.path.exists(output_path):
            os.remove(output_path)
        doc.SaveAs(output_path)
        print(f"[完成] -> {output_path}")

        # ⑥ 可选 PNG 导出
        if export_png:
            png_path = os.path.splitext(output_path)[0] + ".png"
            page.Export(png_path)
            print(f"[完成] -> {png_path}")
    finally:
        render.close_visio(visio, doc)


def main():
    parser = argparse.ArgumentParser(description="Visio 流程图生成引擎")
    parser.add_argument("json_path", help="流程描述 JSON 路径")
    parser.add_argument("-o", "--output", default=None, help="输出 .vsdx 路径（默认与 JSON 同名）")
    parser.add_argument("--pad", type=float, default=0.1, help="画板四周留白（英寸），默认 0.1")
    parser.add_argument("--png", action="store_true", help="同时导出同名 .png")
    args = parser.parse_args()

    with open(args.json_path, "r", encoding="utf-8") as f:
        flow = json.load(f)

    output = args.output
    if output is None:
        output = os.path.splitext(args.json_path)[0] + ".vsdx"

    generate(flow, output, pad=args.pad, export_png=args.png)


if __name__ == "__main__":
    main()
