#!/usr/bin/env python3
# scripts/chart_composer.py
"""
PPT Master - 图表模块布局组合工具

将多个图表模块SVG自动布局组合成一页完整的PPT用SVG。

Usage:
    python3 scripts/chart_composer.py --modules modules/ --data data.json -o output.svg
    python3 scripts/chart_composer.py --modules modules/ --layout grid_2x2 -o output.svg
    python3 scripts/chart_composer.py --modules modules/ --recommendation rec.json -o output.svg
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

SVG_NS = "http://www.w3.org/2000/svg"

CANVAS_W = 1280
CANVAS_H = 720
MARGIN = 60
TITLE_H = 120
CHART_AREA_Y = TITLE_H + 10
CHART_AREA_H = CANVAS_H - CHART_AREA_Y - MARGIN
CHART_AREA_W = CANVAS_W - MARGIN * 2
GAP = 20


# ─── 布局定义 ───


def _make_slots(cols: int, rows: int) -> List[Dict[str, float]]:
    """生成网格布局插槽"""
    slot_w = (CHART_AREA_W - GAP * (cols - 1)) / cols
    slot_h = (CHART_AREA_H - GAP * (rows - 1)) / rows
    slots = []
    for r in range(rows):
        for c in range(cols):
            slots.append({
                "x": MARGIN + c * (slot_w + GAP),
                "y": CHART_AREA_Y + r * (slot_h + GAP),
                "w": slot_w,
                "h": slot_h,
            })
    return slots


LAYOUTS: Dict[str, List[Dict[str, float]]] = {
    "single": [{"x": MARGIN, "y": CHART_AREA_Y, "w": CHART_AREA_W, "h": CHART_AREA_H}],
    "two_column": _make_slots(2, 1),
    "three_column": _make_slots(3, 1),
    "grid_2x2": _make_slots(2, 2),
    "grid_top2_bottom3": (
        _make_slots(2, 1)[:2]
        + [
            {"x": MARGIN + i * ((CHART_AREA_W - GAP * 2) / 3 + GAP),
             "y": CHART_AREA_Y + (CHART_AREA_H - GAP) / 2 + GAP,
             "w": (CHART_AREA_W - GAP * 2) / 3,
             "h": (CHART_AREA_H - GAP) / 2}
            for i in range(3)
        ]
    ),
    "grid_2x3": _make_slots(3, 2),
}

# 修正 grid_top2_bottom3 的上半部分高度
LAYOUTS["grid_top2_bottom3"][0]["h"] = (CHART_AREA_H - GAP) / 2
LAYOUTS["grid_top2_bottom3"][1]["h"] = (CHART_AREA_H - GAP) / 2


def get_layout_slots(layout_name: str) -> List[Dict[str, float]]:
    """获取指定布局的插槽列表"""
    return LAYOUTS.get(layout_name, LAYOUTS["single"])


def auto_select_layout(count: int) -> str:
    """根据图表数量自动选择布局"""
    layout_map = {1: "single", 2: "two_column", 3: "three_column",
                  4: "grid_2x2", 5: "grid_top2_bottom3", 6: "grid_2x3"}
    return layout_map.get(count, "grid_2x3")


# ─── 缩放计算 ───


def calculate_scale(src_w: float, src_h: float,
                    slot_w: float, slot_h: float) -> Tuple[float, float, float]:
    """计算等比缩放和居中偏移，返回 (scale, offset_x, offset_y)"""
    if src_w <= 0 or src_h <= 0:
        return (1.0, 0.0, 0.0)
    scale_x = slot_w / src_w
    scale_y = slot_h / src_h
    scale = min(scale_x, scale_y)
    scaled_w = src_w * scale
    scaled_h = src_h * scale
    offset_x = (slot_w - scaled_w) / 2
    offset_y = (slot_h - scaled_h) / 2
    return (scale, offset_x, offset_y)


# ─── SVG模块解析 ───


def _parse_module_svg(svg_path: Path) -> Tuple[str, float, float, List[str], str]:
    """解析模块SVG，返回 (inner_content, vb_w, vb_h, defs_list, viewbox_origin)"""
    content = svg_path.read_text(encoding="utf-8")
    root = ET.fromstring(content)

    vb = root.get("viewBox", "0 0 400 300")
    parts = vb.split()
    vb_x, vb_y = float(parts[0]), float(parts[1])
    vb_w, vb_h = float(parts[2]), float(parts[3])

    defs_elements = []
    inner_parts = []

    for child in root:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        child_str = ET.tostring(child, encoding="unicode")
        if tag == "defs":
            for d in child:
                defs_elements.append(ET.tostring(d, encoding="unicode"))
        else:
            inner_parts.append(child_str)

    inner_content = "\n".join(inner_parts)
    viewbox_origin = f"{vb_x} {vb_y}"
    return (inner_content, vb_w, vb_h, defs_elements, viewbox_origin)


# ─── 页面组合 ───


def compose_page(module_paths: List[Path],
                 title: str = "",
                 layout: str = "auto",
                 color_scheme: Optional[str] = None) -> str:
    """将多个模块SVG组合成一页完整SVG"""
    n = len(module_paths)
    if layout == "auto":
        layout = auto_select_layout(n)

    slots = get_layout_slots(layout)

    bg_color = "#FFFFFF"
    text_color = "#2C3E50"
    text_muted = "#7F8C8D"
    if color_scheme:
        try:
            from config import DESIGN_COLORS
            scheme = DESIGN_COLORS.get(color_scheme, {})
            bg_color = scheme.get("background", bg_color)
            text_color = scheme.get("text_dark", text_color)
            text_muted = scheme.get("text_muted", text_muted)
        except ImportError:
            pass

    all_defs: List[str] = []
    module_groups: List[str] = []

    for i, mod_path in enumerate(module_paths):
        if i >= len(slots):
            break

        slot = slots[i]
        inner, vb_w, vb_h, defs, vb_origin = _parse_module_svg(mod_path)

        prefix = f"m{i}_"
        for d_str in defs:
            renamed = re.sub(r'id="([^"]+)"', lambda m: f'id="{prefix}{m.group(1)}"', d_str)
            all_defs.append(renamed)
        inner = re.sub(r'url\(#([^)]+)\)', lambda m: f'url(#{prefix}{m.group(1)})', inner)

        scale, off_x, off_y = calculate_scale(vb_w, vb_h, slot["w"], slot["h"])

        vb_parts = vb_origin.split()
        vb_ox, vb_oy = float(vb_parts[0]), float(vb_parts[1])

        tx = slot["x"] + off_x - vb_ox * scale
        ty = slot["y"] + off_y - vb_oy * scale

        module_groups.append(
            f'<g transform="translate({tx:.1f},{ty:.1f}) scale({scale:.4f})">\n'
            f'{inner}\n'
            f'</g>'
        )

    parts = [
        f'<svg xmlns="{SVG_NS}" viewBox="0 0 {CANVAS_W} {CANVAS_H}"'
        f' width="{CANVAS_W}" height="{CANVAS_H}">',
    ]

    if all_defs:
        parts.append("<defs>")
        parts.extend(all_defs)
        parts.append("</defs>")

    parts.append(f'<rect width="{CANVAS_W}" height="{CANVAS_H}" fill="{bg_color}"/>')

    if title:
        parts.append(
            f'<text x="{MARGIN}" y="70"'
            f' font-family="-apple-system, BlinkMacSystemFont, \'Segoe UI\', sans-serif"'
            f' font-size="28" font-weight="bold" fill="{text_color}">'
            f'<tspan>{title}</tspan></text>'
        )

    parts.extend(module_groups)
    parts.append("</svg>")

    return "\n".join(parts)


# ─── CLI ───


def main() -> int:
    parser = argparse.ArgumentParser(
        description="PPT Master - 图表模块布局组合工具",
    )
    parser.add_argument("--modules", type=Path, required=True,
                        help="图表模块目录（含 modules_index.json）")
    parser.add_argument("--data", type=Path, default=None,
                        help="数据文件（JSON/CSV/Excel），触发自动推荐")
    parser.add_argument("--recommendation", type=Path, default=None,
                        help="chart_recommender输出的推荐结果JSON")
    parser.add_argument("--layout", type=str, default="auto",
                        choices=list(LAYOUTS.keys()) + ["auto"],
                        help="布局策略（默认: auto）")
    parser.add_argument("--title", type=str, default="",
                        help="页面标题")
    parser.add_argument("--color-scheme", type=str, default=None,
                        help="配色方案（consulting/general/tech/academic/government）")
    parser.add_argument("-o", "--output", type=Path, required=True,
                        help="输出SVG文件路径")
    args = parser.parse_args()

    modules_dir = args.modules.expanduser().resolve()
    if not modules_dir.exists():
        print(f"[ERROR] 模块目录不存在: {modules_dir}")
        return 1

    module_paths: List[Path] = []

    index_file = modules_dir / "modules_index.json"
    if index_file.exists():
        index = json.loads(index_file.read_text(encoding="utf-8"))
        for mod in index.get("modules", []):
            svg_file = modules_dir / mod["file"]
            if svg_file.exists():
                module_paths.append(svg_file)
    else:
        module_paths = sorted(modules_dir.glob("*.svg"))

    if not module_paths:
        print("[ERROR] 模块目录中没有SVG文件")
        return 1

    if args.recommendation:
        rec_path = args.recommendation.expanduser().resolve()
        if rec_path.exists():
            rec = json.loads(rec_path.read_text(encoding="utf-8"))
            layout_info = rec.get("suggested_layout", {})
            if args.layout == "auto" and "layout_type" in layout_info:
                args.layout = layout_info["layout_type"]

    if args.data and not args.recommendation:
        from chart_recommender import load_json_data, load_csv_data, load_excel_data
        from chart_recommender import analyze_dataset, recommend_chart, suggest_layout

        data_path = args.data.expanduser().resolve()
        suffix = data_path.suffix.lower()
        if suffix == ".json":
            datasets = load_json_data(data_path)
        elif suffix == ".csv":
            datasets = load_csv_data(data_path)
        elif suffix in (".xlsx", ".xls"):
            datasets = load_excel_data(data_path)
        else:
            datasets = []

        if datasets:
            recs = []
            for ds in datasets:
                features = analyze_dataset(ds)
                recs.append(recommend_chart(features))
            layout_info = suggest_layout(recs)
            if args.layout == "auto":
                args.layout = layout_info["layout_type"]
            print(f"[INFO] 自动推荐布局: {args.layout}")

    output_svg = compose_page(
        module_paths=module_paths,
        title=args.title,
        layout=args.layout,
        color_scheme=args.color_scheme,
    )

    out_path = args.output.expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(output_svg, encoding="utf-8")

    print(f"[OK] 组合了 {min(len(module_paths), len(get_layout_slots(args.layout)))} 个图表模块")
    print(f"[OK] 布局: {args.layout}")
    print(f"[OK] 输出: {out_path}")
    print(f"\n下一步:")
    print(f"  python3 scripts/finalize_svg.py <project_dir>")
    print(f"  python3 scripts/svg_to_pptx.py <project_dir> -s final")
    return 0


if __name__ == "__main__":
    sys.exit(main())
