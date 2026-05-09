#!/usr/bin/env python3
# scripts/chart_extractor.py
"""
PPT Master - 图表模块拆分提取工具

从一页完整的SVG中识别并提取每个独立的图表区域为模块。

Usage:
    python3 scripts/chart_extractor.py slide.svg -o modules/
    python3 scripts/chart_extractor.py template.pptx -o modules/
    python3 scripts/chart_extractor.py slide.svg -o modules/ --types bar_chart,donut_chart
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

SVG_NS = "http://www.w3.org/2000/svg"
NS = {"svg": SVG_NS}


# ─── 页面级元素过滤 ───


def _parse_viewbox(root: ET.Element) -> Dict[str, int]:
    """从SVG根元素解析画布尺寸"""
    vb = root.get("viewBox", "")
    parts = vb.split()
    if len(parts) == 4:
        return {"width": int(float(parts[2])), "height": int(float(parts[3]))}
    w = int(float(root.get("width", "1280")))
    h = int(float(root.get("height", "720")))
    return {"width": w, "height": h}


def _get_float(el: ET.Element, attr: str, default: float = 0.0) -> float:
    """安全获取元素的浮点属性"""
    val = el.get(attr)
    if val is None:
        return default
    try:
        return float(val)
    except ValueError:
        return default


def _is_page_level_element(el: ET.Element, canvas: Dict[str, int]) -> bool:
    """判断元素是否为页面级元素（背景/标题/页脚），应从图表提取中排除"""
    tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag

    # 背景矩形：覆盖画布95%以上面积
    if tag == "rect":
        w = _get_float(el, "width")
        h = _get_float(el, "height")
        if w >= canvas["width"] * 0.95 and h >= canvas["height"] * 0.95:
            return True

    # 页面标题：位于顶部区域且字号较大
    if tag == "text":
        y = _get_float(el, "y")
        font_size = _get_float(el, "font-size", 16)
        if y < 120 and font_size >= 24:
            return True
        # 页脚：位于底部区域且字号较小
        if y > canvas["height"] - 40 and font_size <= 14:
            return True

    return False


# ─── 元素统计辅助 ───


def _collect_elements(group: ET.Element) -> Dict[str, List[ET.Element]]:
    """递归收集分组内所有SVG图形元素，按标签分类"""
    result: Dict[str, List[ET.Element]] = {}
    for el in group.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag in ("rect", "circle", "ellipse", "line", "path",
                   "polyline", "polygon", "text", "image", "g"):
            result.setdefault(tag, []).append(el)
    return result


def _has_arc_command(path_d: str) -> bool:
    """检查SVG path d属性是否包含弧线命令(A/a)"""
    return bool(re.search(r"[Aa]\s", path_d))


def _paths_converge_to_center(paths: List[ET.Element]) -> bool:
    """检查多个path是否汇聚到同一中心点（饼图/环形图特征）"""
    if len(paths) < 2:
        return False
    endpoints = []
    for p in paths:
        d = p.get("d", "")
        if "Z" in d or "z" in d:
            parts = d.replace(",", " ").split()
            for i, tok in enumerate(parts):
                if tok in ("L", "l") and i + 2 < len(parts):
                    try:
                        endpoints.append((float(parts[i + 1]), float(parts[i + 2])))
                    except ValueError:
                        pass
    if len(endpoints) < 2:
        return False
    cx = sum(e[0] for e in endpoints) / len(endpoints)
    cy = sum(e[1] for e in endpoints) / len(endpoints)
    threshold = 30.0
    return all(
        abs(e[0] - cx) < threshold and abs(e[1] - cy) < threshold
        for e in endpoints
    )


# ─── 图表类型分类 ───


def _classify_chart_type(group: ET.Element) -> Dict[str, Any]:
    """根据元素特征签名判断图表类型，返回 {type, confidence}"""
    elements = _collect_elements(group)
    rects = elements.get("rect", [])
    circles = elements.get("circle", [])
    paths = elements.get("path", [])
    polylines = elements.get("polyline", [])
    polygons = elements.get("polygon", [])
    lines = elements.get("line", [])
    texts = elements.get("text", [])

    # ── 柱状图: 多个等宽rect ──
    if len(rects) >= 3:
        widths = [_get_float(r, "width") for r in rects]
        data_widths = [w for w in widths if 10 < w < 200]
        if len(data_widths) >= 3:
            from collections import Counter
            wc = Counter(int(w) for w in data_widths)
            most_common_w, count = wc.most_common(1)[0]
            if count >= 3:
                return {"type": "bar_chart", "confidence": round(min(count / len(rects), 1.0), 2)}

    # ── 环形图: path含弧线 + 中心circle ──
    arc_paths = [p for p in paths if _has_arc_command(p.get("d", ""))]
    if len(arc_paths) >= 2 and len(circles) >= 1:
        large_circles = [c for c in circles if _get_float(c, "r") >= 50]
        if large_circles:
            return {"type": "donut_chart", "confidence": 0.90}

    # ── 饼图: path含弧线, 汇聚中心, 无大circle ──
    if len(arc_paths) >= 2:
        large_circles = [c for c in circles if _get_float(c, "r") >= 50]
        if not large_circles and _paths_converge_to_center(arc_paths):
            return {"type": "pie_chart", "confidence": 0.85}

    # ── 折线图: polyline + 多个小circle数据点 ──
    if len(polylines) >= 1 and len(circles) >= 3:
        small_circles = [c for c in circles if _get_float(c, "r") <= 8]
        if len(small_circles) >= 3:
            return {"type": "line_chart", "confidence": 0.88}

    # ── 折线图变体: 连续path(无弧线) + 多个小circle ──
    non_arc_paths = [p for p in paths if not _has_arc_command(p.get("d", ""))]
    if len(non_arc_paths) >= 1 and len(circles) >= 3:
        small_circles = [c for c in circles if _get_float(c, "r") <= 8]
        if len(small_circles) >= 3:
            return {"type": "line_chart", "confidence": 0.75}

    # ── 雷达图: polygon + 多条从中心辐射的line ──
    if len(polygons) >= 1 and len(lines) >= 3:
        return {"type": "radar_chart", "confidence": 0.80}

    # ── KPI卡片: 多个大号数字text + 小号标签text ──
    if len(texts) >= 4:
        large_texts = [t for t in texts if _get_float(t, "font-size", 16) >= 36]
        small_texts = [t for t in texts if _get_float(t, "font-size", 16) <= 18]
        if len(large_texts) >= 2 and len(small_texts) >= 2:
            return {"type": "kpi_cards", "confidence": 0.78}

    # ── 漏斗图: 多个宽度递减的rect ──
    if len(rects) >= 3:
        sorted_rects = sorted(rects, key=lambda r: _get_float(r, "y"))
        widths_ordered = [_get_float(r, "width") for r in sorted_rects]
        if len(widths_ordered) >= 3 and all(
            widths_ordered[i] >= widths_ordered[i + 1]
            for i in range(len(widths_ordered) - 1)
        ):
            return {"type": "funnel_chart", "confidence": 0.75}

    # ── 表格: 网格状line + 对齐的text ──
    if len(lines) >= 4 and len(texts) >= 4:
        h_lines = [l for l in lines if abs(_get_float(l, "y1") - _get_float(l, "y2")) < 2]
        v_lines = [l for l in lines if abs(_get_float(l, "x1") - _get_float(l, "x2")) < 2]
        if len(h_lines) >= 2 and len(v_lines) >= 2:
            return {"type": "table", "confidence": 0.72}

    return {"type": "unknown", "confidence": 0.0}


# ─── 包围盒计算 ───


def _element_bounds(el: ET.Element) -> Optional[Tuple[float, float, float, float]]:
    """计算单个元素的包围盒 (x, y, w, h)，返回None表示无法计算"""
    tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag

    if tag == "rect":
        return (_get_float(el, "x"), _get_float(el, "y"),
                _get_float(el, "width"), _get_float(el, "height"))
    if tag == "circle":
        cx, cy, r = _get_float(el, "cx"), _get_float(el, "cy"), _get_float(el, "r")
        return (cx - r, cy - r, r * 2, r * 2)
    if tag == "ellipse":
        cx = _get_float(el, "cx")
        cy = _get_float(el, "cy")
        rx, ry = _get_float(el, "rx"), _get_float(el, "ry")
        return (cx - rx, cy - ry, rx * 2, ry * 2)
    if tag == "line":
        x1, y1 = _get_float(el, "x1"), _get_float(el, "y1")
        x2, y2 = _get_float(el, "x2"), _get_float(el, "y2")
        return (min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
    if tag == "text":
        x, y = _get_float(el, "x"), _get_float(el, "y")
        fs = _get_float(el, "font-size", 16)
        text_content = "".join(el.itertext())
        est_w = len(text_content) * fs * 0.6
        return (x, y - fs, est_w, fs * 1.2)
    if tag == "polyline":
        points_str = el.get("points", "")
        coords = re.findall(r"[\d.]+", points_str)
        if len(coords) >= 4:
            xs = [float(coords[i]) for i in range(0, len(coords), 2)]
            ys = [float(coords[i]) for i in range(1, len(coords), 2)]
            return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
    if tag == "path":
        nums = re.findall(r"-?[\d.]+", el.get("d", ""))
        if len(nums) >= 4:
            floats = [float(n) for n in nums]
            xs = floats[0::2]
            ys = floats[1::2]
            if xs and ys:
                return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
    return None


def _group_bounds(group: ET.Element) -> Optional[Tuple[float, float, float, float]]:
    """计算一个<g>分组的合并包围盒"""
    min_x, min_y = float("inf"), float("inf")
    max_x, max_y = float("-inf"), float("-inf")
    found = False

    for el in group.iter():
        if el is group:
            continue
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag == "g":
            continue
        b = _element_bounds(el)
        if b is None:
            continue
        x, y, w, h = b
        found = True
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_x = max(max_x, x + w)
        max_y = max(max_y, y + h)

    if not found:
        return None
    return (min_x, min_y, max_x - min_x, max_y - min_y)


# ─── defs 提取 ───


def _extract_referenced_defs(group: ET.Element, defs_el: Optional[ET.Element]) -> List[ET.Element]:
    """提取分组中引用的defs（渐变、滤镜等）"""
    if defs_el is None:
        return []

    referenced_ids = set()
    group_str = ET.tostring(group, encoding="unicode")
    for match in re.finditer(r'url\(#([^)]+)\)', group_str):
        referenced_ids.add(match.group(1))

    result = []
    for child in defs_el:
        child_id = child.get("id", "")
        if child_id in referenced_ids:
            result.append(child)
    return result


# ─── 核心提取逻辑 ───


def extract_modules(svg_path: Path) -> List[Dict[str, Any]]:
    """从SVG文件中提取所有图表模块"""
    tree = ET.parse(str(svg_path))
    root = tree.getroot()
    canvas = _parse_viewbox(root)

    defs_el = root.find(f"{{{SVG_NS}}}defs")
    if defs_el is None:
        defs_el = root.find("defs")

    candidates: List[Tuple[ET.Element, str]] = []

    for g in root.iter():
        tag = g.tag.split("}")[-1] if "}" in g.tag else g.tag
        if tag == "g" and g.get("id"):
            candidates.append((g, g.get("id", "")))

    if not candidates:
        for child in root:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "g":
                candidates.append((child, ""))

    modules = []
    idx = 1
    for group, group_id in candidates:
        child_count = sum(1 for _ in group)
        if child_count == 0:
            continue

        classification = _classify_chart_type(group)
        if classification["type"] == "unknown" and classification["confidence"] == 0.0:
            all_page_level = True
            for child in group:
                if not _is_page_level_element(child, canvas):
                    all_page_level = False
                    break
            if all_page_level:
                continue

        bounds = _group_bounds(group)
        if bounds is None:
            continue

        x, y, w, h = bounds
        if w < 50 or h < 30:
            continue

        ref_defs = _extract_referenced_defs(group, defs_el)
        module_svg = _build_module_svg(group, ref_defs, bounds)

        module_id = f"module_{idx:03d}"
        modules.append({
            "id": module_id,
            "type": classification["type"],
            "confidence": classification["confidence"],
            "bounds": {"x": round(x, 1), "y": round(y, 1),
                       "w": round(w, 1), "h": round(h, 1)},
            "svg_content": module_svg,
            "source_group_id": group_id,
        })
        idx += 1

    return modules


def _build_module_svg(group: ET.Element, ref_defs: List[ET.Element],
                      bounds: Tuple[float, float, float, float]) -> str:
    """构建独立模块SVG字符串"""
    x, y, w, h = bounds
    padding = 10
    vb_x = max(0, x - padding)
    vb_y = max(0, y - padding)
    vb_w = w + padding * 2
    vb_h = h + padding * 2

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg"'
        f' viewBox="{vb_x:.1f} {vb_y:.1f} {vb_w:.1f} {vb_h:.1f}"'
        f' width="{vb_w:.0f}" height="{vb_h:.0f}">'
    ]

    if ref_defs:
        parts.append("<defs>")
        for d in ref_defs:
            parts.append(ET.tostring(d, encoding="unicode"))
        parts.append("</defs>")

    parts.append(ET.tostring(group, encoding="unicode"))
    parts.append("</svg>")
    return "\n".join(parts)


def export_modules(modules: List[Dict[str, Any]], source_path: Path,
                   output_dir: Path) -> Path:
    """将提取的模块导出到目录"""
    output_dir.mkdir(parents=True, exist_ok=True)

    index_modules = []
    for mod in modules:
        filename = f"{mod['id']}_{mod['type']}.svg"
        svg_path = output_dir / filename
        svg_path.write_text(mod["svg_content"], encoding="utf-8")

        index_modules.append({
            "id": mod["id"],
            "file": filename,
            "type": mod["type"],
            "confidence": mod["confidence"],
            "original_bounds": mod["bounds"],
        })

    index = {
        "source": source_path.name,
        "modules": index_modules,
    }
    index_path = output_dir / "modules_index.json"
    index_path.write_text(
        json.dumps(index, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return index_path


# ─── PPTX 支持 ───


def extract_from_pptx(pptx_path: Path, output_dir: Path,
                      types_filter: Optional[List[str]] = None) -> Path:
    """从PPTX文件提取图表模块（先导出SVG再解析）"""
    import tempfile

    svg_dir = Path(tempfile.mkdtemp(prefix="chart_extract_"))
    try:
        from pptx_template_import import export_pptx_slides_to_svg_with_fallback
        svg_files, mode = export_pptx_slides_to_svg_with_fallback(pptx_path, svg_dir)
        print(f"[INFO] PPTX导出模式: {mode}, {len(svg_files)}页SVG")
    except ImportError:
        print("[ERROR] 需要 pptx_template_import 模块支持PPTX导入")
        sys.exit(1)

    all_modules = []
    for svg_file in sorted(svg_dir.glob("*.svg")):
        modules = extract_modules(svg_file)
        if types_filter:
            modules = [m for m in modules if m["type"] in types_filter]
        all_modules.extend(modules)

    for i, mod in enumerate(all_modules, 1):
        mod["id"] = f"module_{i:03d}"

    return export_modules(all_modules, pptx_path, output_dir)


# ─── CLI ───


def main() -> int:
    parser = argparse.ArgumentParser(
        description="PPT Master - 图表模块拆分提取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s slide.svg -o modules/
  %(prog)s template.pptx -o modules/
  %(prog)s slide.svg -o modules/ --types bar_chart,donut_chart
        """,
    )
    parser.add_argument("input_file", type=Path, help="输入文件（SVG或PPTX）")
    parser.add_argument("-o", "--output", type=Path, default=Path("modules"),
                        help="输出目录（默认: modules/）")
    parser.add_argument("--types", type=str, default=None,
                        help="只提取指定类型（逗号分隔，如 bar_chart,donut_chart）")
    args = parser.parse_args()

    input_path = args.input_file.expanduser().resolve()
    if not input_path.exists():
        print(f"[ERROR] 文件不存在: {input_path}")
        return 1

    types_filter = args.types.split(",") if args.types else None
    output_dir = args.output.expanduser().resolve()
    suffix = input_path.suffix.lower()

    if suffix == ".svg":
        modules = extract_modules(input_path)
        if types_filter:
            modules = [m for m in modules if m["type"] in types_filter]
        index_path = export_modules(modules, input_path, output_dir)
    elif suffix in (".pptx", ".pptm"):
        index_path = extract_from_pptx(input_path, output_dir, types_filter)
    else:
        print(f"[ERROR] 不支持的文件类型: {suffix}（支持 .svg, .pptx）")
        return 1

    index = json.loads(index_path.read_text(encoding="utf-8"))
    n = len(index["modules"])
    print(f"[OK] 提取了 {n} 个图表模块")
    for mod in index["modules"]:
        print(f"  {mod['id']}: {mod['type']} (置信度 {mod['confidence']:.0%})")
    print(f"[OK] 输出目录: {output_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
