#!/usr/bin/env python3
"""将 pptx 文件解析为结构化 Markdown（HTML 表格 + Mermaid 流程图 + 图表数据）。

解析能力：
  - 文本框 → Markdown 文本（自动识别标题层级）
  - 表格   → HTML <table>（支持合并单元格）
  - 图表   → HTML 表格 + <caption> 标注图表类型（Bar/Pie/Line）
  - 流程图 → Mermaid graph LR（从矩形/菱形/圆角矩形/箭头自动构建）

用法：
  python pptx_to_md.py input.pptx [output.md]
"""
import argparse
import os
import re
from collections import OrderedDict

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

SHAPE_TYPE_MAP = {
    MSO_SHAPE_TYPE.AUTO_SHAPE: "shape",
    MSO_SHAPE_TYPE.PLACEHOLDER: "placeholder",
    MSO_SHAPE_TYPE.FREEFORM: "freeform",
    MSO_SHAPE_TYPE.CHART: "chart",
    MSO_SHAPE_TYPE.TABLE: "table",
    MSO_SHAPE_TYPE.GROUP: "group",
    MSO_SHAPE_TYPE.PICTURE: "picture",
    MSO_SHAPE_TYPE.LINE: "connector",
}


def is_connector(shape):
    """判断 shape 是否为连接线（cxnSp）。"""
    try:
        return shape._element.tag == qn('p:cxnSp')
    except Exception:
        return shape.shape_type == MSO_SHAPE_TYPE.LINE

CHART_TYPE_NAMES = {
    # xlChartType → 中文名
    51: "COLUMN_CLUSTERED", 54: "COLUMN_STACKED",
    57: "BAR_STACKED",
    4: "LINE", 63: "LINE_MARKERS",
    5: "PIE", 68: "PIE_EXPLODED",
    73: "DOUGHNUT",
    11: "AREA",
    20: "XY_SCATTER",
    -4151: "RADAR",
}

FRIENDLY_CHART = {
    "COLUMN_CLUSTERED": "Bar Chart",
    "COLUMN_STACKED": "Bar Chart",
    "BAR_CLUSTERED": "Bar Chart",
    "BAR_STACKED": "Bar Chart",
    "LINE": "Line Chart",
    "LINE_MARKERS": "Line Chart",
    "PIE": "Pie Chart",
    "PIE_EXPLODED": "Pie Chart",
    "DOUGHNUT": "Pie Chart",
    "AREA": "Area Chart",
    "XY_SCATTER": "Scatter Chart",
    "RADAR": "Radar Chart",
}

DIAMOND_TYPES = {4}  # MSO_AUTO_SHAPE_TYPE.DIAMOND = 4
ROUNDED_TYPES = {5}  # ROUNDED_RECTANGLE


def get_auto_shape_type(shape):
    """从 shape 元素获取 autoShapeType（仅对 auto_shape 有效）。"""
    try:
        sp_el = shape._element.find(qn('p:sp'))
        if sp_el is None:
            sp_el = shape._element
        prst = sp_el.find(qn('p:spPr'))
        if prst is None:
            prst = sp_el.find(qn('a:spPr'))
        if prst is not None:
            geom = prst.find(qn('a:prstGeom'))
            if geom is not None:
                return geom.get('prst')
    except Exception:
        pass
    return None


def get_text(shape):
    """提取 shape 的全部文本，保留换行。"""
    if not shape.has_text_frame:
        return ""
    parts = []
    for p in shape.text_frame.paragraphs:
        line = p.text.strip()
        if line:
            parts.append(line)
    return "\n".join(parts) if parts else ""


def parse_table(shape):
    """将 pptx table → HTML 表格，支持合并单元格。"""
    tbl = shape.table
    nrows = len(tbl.rows)
    ncols = len(tbl.columns)

    # 构建合并单元格映射
    merges = []
    for r in range(nrows):
        for c in range(ncols):
            cell = tbl.cell(r, c)
            span = getattr(cell, "span_width", 1) or 1
            vspan = getattr(cell, "span_height", 1) or 1
            merges.append(((r, c), span, vspan, cell.text.strip()))

    html = "<table>\n"
    for r in range(nrows):
        html += "  <tr>\n"
        seen = set()
        for c in range(ncols):
            key = (r, c)
            if key in seen:
                continue
            cell = tbl.cell(r, c)
            text = cell.text.strip()
            colspan = getattr(cell, "span_width", 1) or 1
            rowspan = getattr(cell, "span_height", 1) or 1
            for dr in range(rowspan):
                for dc in range(colspan):
                    seen.add((r + dr, c + dc))
            attrs = ""
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            if rowspan > 1:
                attrs += f' rowspan="{rowspan}"'
            tag = "th" if r == 0 else "td"
            html += f"    <{tag}{attrs}>\n      {text}\n    </{tag}>\n"
        html += "  </tr>\n"
    html += "</table>"
    return html


def _fmt_val(val):
    """格式化数值：整数去掉 .0，保留小数原样。"""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)


def parse_chart(shape):
    """将 pptx 图表 → HTML 表格 + caption 标注图表类型。"""
    chart = shape.chart
    ct = chart.chart_type
    ct_name = CHART_TYPE_NAMES.get(ct, f"UNKNOWN({ct})")
    friendly = FRIENDLY_CHART.get(ct_name, "Chart")

    # 提取系列数据
    series_data = []
    for s in chart.series:
        vals = [_fmt_val(v) if v is not None else "" for v in s.values]
        sname = ""
        try:
            sname = s.name if hasattr(s, 'name') else ""
        except Exception:
            pass
        series_data.append((sname or f"Series {len(series_data)}", vals))

    # 提取类别标签
    categories = []
    try:
        for pt in chart.plots[0].categories:
            categories.append(str(pt))
    except Exception:
        categories = [f"Cat{i+1}" for i in range(len(series_data[0][1]) if series_data else 0)]

    html = "<table>\n"
    html += f"  <caption>{friendly}: Chart Title</caption>\n"
    html += "  <tr>\n    <th>\n    </th>\n"
    for sname, _ in series_data:
        html += f"    <th>\n      {sname}\n    </th>\n"
    html += "  </tr>\n"
    n_series = len(series_data)
    n_cats = len(categories)
    for i, cat in enumerate(categories):
        html += "  <tr>\n"
        html += f"    <td>\n      {cat}\n    </td>\n"
        for sname, vals in series_data:
            val = vals[i] if i < len(vals) else ""
            html += f"    <td>\n      {val}\n    </td>\n"
        html += "  </tr>\n"
    html += "</table>"
    return html


def detect_flowchart(shapes):
    """检测一页是否为流程图：至少 2 个矩形类 shape + 连接线/箭头。"""
    flow_shapes = []
    connectors = []
    for s in shapes:
        if is_connector(s):
            connectors.append(s)
            continue
        if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            ast = get_auto_shape_type(s)
            # 箭头形状也是连接线（RIGHT_ARROW 等），文本为空
            if ast and "arrow" in ast.lower() and not get_text(s):
                connectors.append(s)
                continue
            text = get_text(s)
            if text:
                flow_shapes.append(s)
    is_flow = len(flow_shapes) >= 2 and len(connectors) >= 1
    return is_flow, flow_shapes, connectors


def _get_shape_bbox_emu(shape):
    """获取 shape 的边界框 (x1, y1, x2, y2) 单位 EMU。"""
    el = shape._element
    xfrm = el.find('.//' + qn('a:xfrm'))
    if xfrm is None:
        return None
    off = xfrm.find(qn('a:off'))
    ext = xfrm.find(qn('a:ext'))
    if off is None or ext is None:
        return None
    x = int(off.get('x', 0))
    y = int(off.get('y', 0))
    cx = int(ext.get('cx', 0))
    cy = int(ext.get('cy', 0))
    flip_h = xfrm.get('flipH') == '1'
    flip_v = xfrm.get('flipV') == '1'
    return x, y, cx, cy, flip_h, flip_v


def _connector_endpoints(connector):
    """获取连接线的视觉起点和终点 (x, y) EMU。"""
    bbox = _get_shape_bbox_emu(connector)
    if not bbox:
        return None, None
    x, y, cx, cy, flip_h, flip_v = bbox
    # xfrm 的 off 是左上角；对于 flipH，视觉起点在右侧
    if flip_h:
        start = (x + cx, y)
        end = (x, y + cy)
    else:
        start = (x, y)
        end = (x + cx, y + cy)
    return start, end


def _node_center(shape):
    """获取节点中心点 (x, y) EMU。"""
    bbox = _get_shape_bbox_emu(shape)
    if not bbox:
        return (0, 0)
    x, y, cx, cy, _, _ = bbox
    return (x + cx // 2, y + cy // 2)


def _nearest_node(point, nodes, exclude_id=None):
    """在节点列表中找离 point 最近的节点（排除 exclude_id）。"""
    best = None
    best_dist = float('inf')
    for s in nodes:
        if exclude_id and s.shape_id == exclude_id:
            continue
        cx, cy = _node_center(s)
        dx = point[0] - cx
        dy = point[1] - cy
        dist = dx * dx + dy * dy
        if dist < best_dist:
            best_dist = dist
            best = s
    return best


def shape_to_mermaid_id(text):
    """将 shape 文本简化为 Mermaid 节点 ID。"""
    tid = re.sub(r'[\s\-/\\]+', '', text)
    tid = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fff]', '', tid)
    if not tid:
        tid = "N"
    return tid


def shapes_to_mermaid(flow_shapes, connectors):
    """将流程图 shapes → Mermaid 语法（基于空间位置匹配连接关系）。"""
    node_map = OrderedDict()
    for i, s in enumerate(flow_shapes):
        text = get_text(s)
        tid = shape_to_mermaid_id(text) or f"node{i}"
        node_map[s.shape_id] = (tid, text, s)

    # 根据空间位置推断连接线两端连接的是哪些节点
    edges = []
    for conn in connectors:
        start_pt, end_pt = _connector_endpoints(conn)
        if not start_pt or not end_pt:
            continue
        src_node = _nearest_node(start_pt, flow_shapes)
        dst_node = _nearest_node(end_pt, flow_shapes)
        if src_node and dst_node and src_node.shape_id != dst_node.shape_id:
            s_tid = node_map[src_node.shape_id][0]
            e_tid = node_map[dst_node.shape_id][0]
            edges.append((s_tid, e_tid))

    lines = ["graph LR"]
    for sid, (tid, text, _) in node_map.items():
        ast = get_auto_shape_type(node_map[sid][2])
        label = text.replace('"', '\\"')
        if ast and "diamond" in ast.lower():
            lines.append(f"    {tid}{{\"{label}\"}}")
        else:
            lines.append(f"    {tid}[\"{label}\"]")
    for s_tid, e_tid in edges:
        lines.append(f"    {s_tid} --> {e_tid}")
    return "\n".join(lines)


def detect_heading_level(text, font_size):
    """根据字号判断 Markdown 标题层级。"""
    if font_size >= 40:
        return 1  # #
    if font_size >= 26:
        return 2  # ##
    if font_size >= 20:
        return 3  # ###
    return 0  # 普通文本


def get_max_font_size(shape):
    """获取 shape 中最大字号。"""
    if not shape.has_text_frame:
        return 0
    sizes = []
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if run.font.size:
                sizes.append(run.font.size.pt)
    return max(sizes) if sizes else 0


def parse_slide(prs_slide, slide_idx):
    """解析单页 pptx → Markdown 文本块列表。"""
    blocks = []
    shapes = list(prs_slide.shapes)

    # 先检测是否为流程图
    is_flow, flow_shapes, connectors = detect_flowchart(shapes)

    if is_flow and len(flow_shapes) >= 2:
        # 尝试构建 mermaid
        mermaid = shapes_to_mermaid(flow_shapes, connectors)
        if mermaid:
            blocks.append(("flowchart", mermaid))
            return blocks

    # 逐 shape 处理
    for shape in shapes:
        stype = shape.shape_type

        if stype == MSO_SHAPE_TYPE.TABLE:
            blocks.append(("table", parse_table(shape)))

        elif stype == MSO_SHAPE_TYPE.CHART:
            blocks.append(("chart", parse_chart(shape)))

        elif is_connector(shape):
            continue  # 单独的连接线跳过

        elif shape.has_text_frame:
            text = get_text(shape)
            if not text:
                continue
            max_sz = get_max_font_size(shape)
            level = detect_heading_level(text, max_sz)
            if level > 0:
                blocks.append(("heading", (level, text)))
            else:
                # 判断是否为列表
                if text.startswith("• ") or text.startswith("- ") or text.startswith("* "):
                    blocks.append(("list", text))
                else:
                    blocks.append(("text", text))

        elif stype == MSO_SHAPE_TYPE.AUTO_SHAPE:
            # 形状中的短文本（流程图节点等）
            text = get_text(shape)
            if text:
                blocks.append(("text", text))

    return blocks


def blocks_to_markdown(blocks, slide_title):
    """将解析后的 blocks 列表转为 Markdown 段落。"""
    md = []
    if slide_title:
        md.append(f"# {slide_title}\n")

    for btype, content in blocks:
        if btype == "heading":
            level, text = content
            # 跳过与页面标题重复的 heading
            if slide_title and text == slide_title:
                continue
            md.append(f"\n{'#' * level} {text}\n")
        elif btype == "text":
            md.append(f"\n{content}\n")
        elif btype == "list":
            md.append(f"\n{content}\n")
        elif btype == "table":
            md.append(f"\n{content}\n")
        elif btype == "chart":
            md.append(f"\n{content}\n")
        elif btype == "flowchart":
            md.append(f"\n```mermaid\n{content}\n```\n")
    return "\n".join(md)


def pptx_to_markdown(input_path, output_path=None):
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".md"

    prs = Presentation(input_path)
    all_md = []

    for i, slide in enumerate(prs.slides):
        # 获取标题（优先用 title placeholder）
        title = ""
        for shape in slide.shapes:
            if shape.is_placeholder and shape.placeholder_format.idx == 0:
                title = get_text(shape)
                break
        if not title:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = get_text(shape)
                    if text:
                        max_sz = get_max_font_size(shape)
                        if max_sz >= 26:
                            title = text
                            break

        blocks = parse_slide(slide, i)
        if not blocks:
            continue

        # 跳过纯「谢谢」等结束页
        if len(blocks) == 1:
            btype = blocks[0][0]
            t = blocks[0][1].strip() if btype == "text" else ""
            if btype == "heading":
                t = blocks[0][1][1].strip()
            if t in ("谢谢", "Thank you", "Thanks"):
                all_md.append(f"\n{t}")
                continue

        slide_md = blocks_to_markdown(blocks, title)
        all_md.append(slide_md)

    full_md = "\n\n".join(all_md)
    # 清理多余空行
    full_md = re.sub(r'\n{3,}', '\n\n', full_md)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(full_md)

    print(f"已生成：{output_path}")
    print(f"共 {len(prs.slides)} 页 → {len(all_md)} 个 Markdown 段落")


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="pptx → Markdown 解析器")
    p.add_argument("input", help="输入 pptx 文件路径")
    p.add_argument("output", nargs="?", default=None,
                   help="输出 md 文件路径（默认同名 .md）")
    args = p.parse_args()
    pptx_to_markdown(args.input, args.output)
