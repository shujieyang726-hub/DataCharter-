# tests/test_chart_recommender.py
"""chart_recommender 单元测试"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))


def test_analyze_comparison_data():
    """类别对比数据应识别为comparison意图"""
    from chart_recommender import analyze_dataset

    dataset = {
        "title": "各区域销售额",
        "categories": ["华东", "华南", "华北", "西南"],
        "values": [185, 142, 128, 96],
    }
    features = analyze_dataset(dataset)
    assert features["data_intent"] == "comparison"
    assert features["category_count"] == 4
    assert features["series_count"] == 1
    assert features["is_time_series"] is False
    assert features["is_percentage"] is False


def test_analyze_percentage_data():
    """百分比数据应识别为composition意图"""
    from chart_recommender import analyze_dataset

    dataset = {
        "title": "产品线占比",
        "labels": ["企业服务", "消费电子", "云计算", "物联网", "其他"],
        "values": [35, 28, 20, 12, 5],
        "unit": "%",
    }
    features = analyze_dataset(dataset)
    assert features["is_percentage"] is True
    assert features["data_intent"] == "composition"
    assert features["category_count"] == 5


def test_analyze_time_series():
    """时间序列数据应识别为trend意图"""
    from chart_recommender import analyze_dataset

    dataset = {
        "title": "月度营收",
        "x_axis": ["1月", "2月", "3月", "4月", "5月", "6月"],
        "series": [
            {"name": "营收", "values": [32, 38, 45, 52, 58, 62]},
        ],
    }
    features = analyze_dataset(dataset)
    assert features["is_time_series"] is True
    assert features["data_intent"] == "trend"
    assert features["series_count"] == 1


def test_analyze_multi_series_time():
    """多系列时间序列"""
    from chart_recommender import analyze_dataset

    dataset = {
        "title": "营收与利润",
        "x_axis": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            {"name": "营收", "values": [100, 120, 130, 150]},
            {"name": "利润", "values": [20, 25, 28, 35]},
        ],
    }
    features = analyze_dataset(dataset)
    assert features["series_count"] == 2
    assert features["is_time_series"] is True


def test_analyze_negative_values():
    """包含负数值"""
    from chart_recommender import analyze_dataset

    dataset = {
        "title": "利润变化",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "values": [50, -20, 30, -10],
    }
    features = analyze_dataset(dataset)
    assert features["has_negative"] is True


def test_recommend_donut_for_percentage():
    """百分比数据应推荐donut_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "composition",
        "category_count": 5,
        "series_count": 1,
        "is_time_series": False,
        "is_percentage": True,
        "has_negative": False,
        "value_range": (5, 35),
        "label_length": 3.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "donut_chart"
    assert len(result["alternatives"]) >= 1


def test_recommend_line_for_time_series():
    """单系列时间序列应推荐line_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "trend",
        "category_count": 12,
        "series_count": 1,
        "is_time_series": True,
        "is_percentage": False,
        "has_negative": False,
        "value_range": (32, 92),
        "label_length": 2.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "line_chart"


def test_recommend_dual_axis_for_two_series_time():
    """双系列时间序列应推荐dual_axis_line_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "trend",
        "category_count": 12,
        "series_count": 2,
        "is_time_series": True,
        "is_percentage": False,
        "has_negative": False,
        "value_range": (12, 92),
        "label_length": 2.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "dual_axis_line_chart"


def test_recommend_bar_for_comparison():
    """短标签类别对比应推荐bar_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "comparison",
        "category_count": 6,
        "series_count": 1,
        "is_time_series": False,
        "is_percentage": False,
        "has_negative": False,
        "value_range": (52, 185),
        "label_length": 2.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "bar_chart"


def test_recommend_horizontal_bar_for_long_labels():
    """长标签应推荐horizontal_bar_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "comparison",
        "category_count": 5,
        "series_count": 1,
        "is_time_series": False,
        "is_percentage": False,
        "has_negative": False,
        "value_range": (10, 100),
        "label_length": 8.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "horizontal_bar_chart"


def test_recommend_waterfall_for_negative():
    """包含负数应推荐waterfall_chart"""
    from chart_recommender import recommend_chart

    features = {
        "data_intent": "comparison",
        "category_count": 5,
        "series_count": 1,
        "is_time_series": False,
        "is_percentage": False,
        "has_negative": True,
        "value_range": (-20, 50),
        "label_length": 2.0,
    }
    result = recommend_chart(features)
    assert result["primary"]["type"] == "waterfall_chart"
