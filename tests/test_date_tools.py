"""日期工具测试"""
from datetime import date


def test_get_week_range_this_week():
    """get_week_range 返回本周的起止日期，周一是起始"""
    from src.tools.date_tools import get_week_range

    result = get_week_range(week_offset=0)

    assert result["success"] is True
    assert "start" in result
    assert "end" in result
    assert "week_label" in result
    # 起始必须是周一
    d = date.fromisoformat(result["start"])
    assert d.weekday() == 0


def test_get_week_range_last_week():
    """get_week_range week_offset=-1 返回上周"""
    from src.tools.date_tools import get_week_range

    this_week = get_week_range(week_offset=0)
    last_week = get_week_range(week_offset=-1)

    assert last_week["start"] < this_week["start"]


def test_format_date_chinese():
    """format_date 中文格式输出"""
    from src.tools.date_tools import format_date

    result = format_date("2026-06-16")

    assert result["success"] is True
    assert "2026" in result["formatted"]
    assert "06" in result["formatted"]
    assert "16" in result["formatted"]
