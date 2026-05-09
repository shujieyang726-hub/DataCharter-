# tests/test_chart_extractor.py
"""chart_extractor 单元测试"""
import sys
import json
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))


def _make_svg(inner: str, w: int = 1280, h: int = 720) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}"'
        f' width="{w}" height="{h}">{inner}</svg>'
    )


def test_filter_background_rect():
    """全画布背景rect应被过滤"""
    from chart_extractor import _is_page_level_element

    svg = _make_svg("")
    root = ET.fromstring(svg)
    canvas = {"width": 1280, "height": 720}

    bg = ET.fromstring('<rect width="1280" height="720" fill="#FFFFFF"/>')
    assert _is_page_level_element(bg, canvas) is True


def test_filter_page_title():
    """y<120且font-size>=24的text应被识别为页面标题"""
    from chart_extractor import _is_page_level_element

    canvas = {"width": 1280, "height": 720}
    title = ET.fromstring(
        '<text x="60" y="80" font-size="32" font-weight="bold">'
        "<tspan>标题</tspan></text>"
    )
    assert _is_page_level_element(title, canvas) is True


def test_filter_page_footer():
    """y>680且font-size<=14的text应被识别为页脚"""
    from chart_extractor import _is_page_level_element

    canvas = {"width": 1280, "height": 720}
    footer = ET.fromstring(
        '<text x="60" y="695" font-size="14" fill="#95A5A6">'
        "<tspan>数据来源</tspan></text>"
    )
    assert _is_page_level_element(footer, canvas) is True


def test_keep_chart_element():
    """图表区域内的rect不应被过滤"""
    from chart_extractor import _is_page_level_element

    canvas = {"width": 1280, "height": 720}
    bar = ET.fromstring('<rect x="220" y="180" width="50" height="370" fill="#2196F3"/>')
    assert _is_page_level_element(bar, canvas) is False


def test_classify_bar_chart():
    """包含多个等宽rect的分组应被识别为bar_chart"""
    from chart_extractor import _classify_chart_type

    group = ET.fromstring(
        '<g xmlns="http://www.w3.org/2000/svg">'
        '<rect x="220" y="180" width="50" height="370" fill="#2196F3"/>'
        '<rect x="400" y="266" width="50" height="284" fill="#2196F3"/>'
        '<rect x="580" y="294" width="50" height="256" fill="#2196F3"/>'
        '<rect x="760" y="358" width="50" height="192" fill="#4CAF50"/>'
        "</g>"
    )
    result = _classify_chart_type(group)
    assert result["type"] == "bar_chart"
    assert result["confidence"] >= 0.7


def test_classify_donut_chart():
    """包含弧线path+中心circle的分组应被识别为donut_chart"""
    from chart_extractor import _classify_chart_type

    group = ET.fromstring(
        '<g xmlns="http://www.w3.org/2000/svg">'
        '<path d="M 0,-180 A 180,180 0 0,1 156,88 L 86,49 A 100,100 0 0,0 0,-100 Z" fill="#2196F3"/>'
        '<path d="M 156,88 A 180,180 0 0,1 -123,130 L -68,72 A 100,100 0 0,0 86,49 Z" fill="#4CAF50"/>'
        '<circle cx="0" cy="0" r="100" fill="#FFFFFF"/>'
        "</g>"
    )
    result = _classify_chart_type(group)
    assert result["type"] == "donut_chart"
    assert result["confidence"] >= 0.7


def test_classify_line_chart():
    """包含polyline+多个小circle的分组应被识别为line_chart"""
    from chart_extractor import _classify_chart_type

    group = ET.fromstring(
        '<g xmlns="http://www.w3.org/2000/svg">'
        '<polyline points="225,422 310,398 395,370 480,342" fill="none" stroke="#2196F3"/>'
        '<circle cx="225" cy="422" r="5" fill="#2196F3"/>'
        '<circle cx="310" cy="398" r="5" fill="#2196F3"/>'
        '<circle cx="395" cy="370" r="5" fill="#2196F3"/>'
        '<circle cx="480" cy="342" r="5" fill="#2196F3"/>'
        "</g>"
    )
    result = _classify_chart_type(group)
    assert result["type"] == "line_chart"
    assert result["confidence"] >= 0.7


def test_classify_pie_chart():
    """包含多个弧线path但无中心circle的分组应被识别为pie_chart"""
    from chart_extractor import _classify_chart_type

    group = ET.fromstring(
        '<g xmlns="http://www.w3.org/2000/svg">'
        '<path d="M 0,-180 A 180,180 0 0,1 156,88 L 0,0 Z" fill="#2196F3"/>'
        '<path d="M 156,88 A 180,180 0 0,1 -123,130 L 0,0 Z" fill="#4CAF50"/>'
        '<path d="M -123,130 A 180,180 0 0,1 0,-180 L 0,0 Z" fill="#FF9800"/>'
        "</g>"
    )
    result = _classify_chart_type(group)
    assert result["type"] == "pie_chart"
    assert result["confidence"] >= 0.7


def test_classify_unknown():
    """无法匹配的分组应返回unknown"""
    from chart_extractor import _classify_chart_type

    group = ET.fromstring(
        '<g xmlns="http://www.w3.org/2000/svg">'
        '<text x="100" y="100" font-size="16">Hello</text>'
        "</g>"
    )
    result = _classify_chart_type(group)
    assert result["type"] == "unknown"


import tempfile
import os


def test_extract_modules_from_bar_chart_svg():
    """从包含柱状图的完整SVG中提取出模块"""
    from chart_extractor import extract_modules

    svg_content = _make_svg(
        '<rect width="1280" height="720" fill="#FFFFFF"/>'
        '<text x="60" y="80" font-size="32" font-weight="bold">标题</text>'
        '<text x="60" y="695" font-size="14" fill="#95A5A6">数据来源</text>'
        '<g id="chartArea">'
        '  <rect x="220" y="180" width="50" height="370" fill="#2196F3"/>'
        '  <rect x="400" y="266" width="50" height="284" fill="#2196F3"/>'
        '  <rect x="580" y="294" width="50" height="256" fill="#2196F3"/>'
        '  <rect x="760" y="358" width="50" height="192" fill="#4CAF50"/>'
        "</g>"
    )
    with tempfile.NamedTemporaryFile(mode="w", suffix=".svg", delete=False, encoding="utf-8") as f:
        f.write(svg_content)
        svg_path = f.name

    try:
        modules = extract_modules(Path(svg_path))
        assert len(modules) >= 1
        chart_module = modules[0]
        assert chart_module["type"] == "bar_chart"
        assert "bounds" in chart_module
        assert "svg_content" in chart_module
    finally:
        os.unlink(svg_path)


def test_export_modules_to_directory():
    """导出模块到目录，生成SVG文件和modules_index.json"""
    from chart_extractor import extract_modules, export_modules

    svg_content = _make_svg(
        '<rect width="1280" height="720" fill="#FFFFFF"/>'
        '<g id="chartArea">'
        '  <rect x="220" y="180" width="50" height="370" fill="#2196F3"/>'
        '  <rect x="400" y="266" width="50" height="284" fill="#2196F3"/>'
        '  <rect x="580" y="294" width="50" height="256" fill="#2196F3"/>'
        "</g>"
    )
    with tempfile.NamedTemporaryFile(mode="w", suffix=".svg", delete=False, encoding="utf-8") as f:
        f.write(svg_content)
        svg_path = Path(f.name)

    with tempfile.TemporaryDirectory() as out_dir:
        out_path = Path(out_dir)
        modules = extract_modules(svg_path)
        export_modules(modules, svg_path, out_path)

        index_file = out_path / "modules_index.json"
        assert index_file.exists()
        index = json.loads(index_file.read_text(encoding="utf-8"))
        assert len(index["modules"]) >= 1
        assert index["modules"][0]["type"] == "bar_chart"

        svg_files = list(out_path.glob("module_*.svg"))
        assert len(svg_files) >= 1

    os.unlink(svg_path)
