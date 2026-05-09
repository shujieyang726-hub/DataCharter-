# tests/test_chart_composer.py
"""chart_composer 单元测试"""
import sys
import json
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))


def test_get_layout_slots():
    """布局策略应返回正确数量的插槽"""
    from chart_composer import get_layout_slots

    assert len(get_layout_slots("single")) == 1
    assert len(get_layout_slots("two_column")) == 2
    assert len(get_layout_slots("three_column")) == 3
    assert len(get_layout_slots("grid_2x2")) == 4


def test_auto_select_layout():
    """自动布局选择应匹配图表数量"""
    from chart_composer import auto_select_layout

    assert auto_select_layout(1) == "single"
    assert auto_select_layout(2) == "two_column"
    assert auto_select_layout(3) == "three_column"
    assert auto_select_layout(4) == "grid_2x2"


def test_calculate_scale():
    """缩放计算应保持宽高比"""
    from chart_composer import calculate_scale

    scale, offset_x, offset_y = calculate_scale(
        src_w=1000, src_h=500,
        slot_w=500, slot_h=400,
    )
    assert scale <= 0.5 + 0.01
    assert offset_x >= 0
    assert offset_y >= 0


def test_compose_svg_output():
    """组合输出应为合法SVG"""
    from chart_composer import compose_page

    module_svg_1 = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300" width="400" height="300">'
        '<rect x="10" y="10" width="50" height="200" fill="#2196F3"/>'
        '<rect x="80" y="50" width="50" height="160" fill="#2196F3"/>'
        '</svg>'
    )
    module_svg_2 = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 300" width="300" height="300">'
        '<circle cx="150" cy="150" r="100" fill="#4CAF50"/>'
        '</svg>'
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        m1 = Path(tmpdir) / "m1.svg"
        m2 = Path(tmpdir) / "m2.svg"
        m1.write_text(module_svg_1, encoding="utf-8")
        m2.write_text(module_svg_2, encoding="utf-8")

        result = compose_page(
            module_paths=[m1, m2],
            title="测试页面",
            layout="two_column",
        )

    assert 'viewBox="0 0 1280 720"' in result
    assert "测试页面" in result
    root = ET.fromstring(result)
    assert root.tag.endswith("svg")


def test_full_pipeline_extract_recommend_compose():
    """端到端: 提取 → 推荐 → 组合"""
    from chart_extractor import extract_modules, export_modules
    from chart_recommender import analyze_dataset, recommend_chart, suggest_layout
    from chart_composer import compose_page

    svg_content = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720" width="1280" height="720">'
        '<rect width="1280" height="720" fill="#FFFFFF"/>'
        '<text x="60" y="80" font-size="32" font-weight="bold">标题</text>'
        '<g id="chartArea">'
        '  <rect x="220" y="180" width="50" height="370" fill="#2196F3"/>'
        '  <rect x="400" y="266" width="50" height="284" fill="#2196F3"/>'
        '  <rect x="580" y="294" width="50" height="256" fill="#2196F3"/>'
        '</g>'
        '</svg>'
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        svg_path = tmpdir / "input.svg"
        svg_path.write_text(svg_content, encoding="utf-8")

        # Step 1: 提取
        modules = extract_modules(svg_path)
        modules_dir = tmpdir / "modules"
        export_modules(modules, svg_path, modules_dir)

        # Step 2: 推荐
        dataset = {"title": "测试", "categories": ["A", "B", "C"], "values": [100, 80, 60]}
        features = analyze_dataset(dataset)
        rec = recommend_chart(features)
        assert rec["primary"]["type"] == "bar_chart"

        # Step 3: 组合
        module_paths = sorted(modules_dir.glob("module_*.svg"))
        assert len(module_paths) >= 1

        result = compose_page(
            module_paths=module_paths,
            title="集成测试",
            layout="single",
        )

        assert 'viewBox="0 0 1280 720"' in result
        assert "集成测试" in result

        root = ET.fromstring(result)
        assert root.tag.endswith("svg")
