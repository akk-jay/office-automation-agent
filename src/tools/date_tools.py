"""日期与日程工具"""

from datetime import date, datetime, timedelta


def get_week_range(week_offset: int = 0) -> dict:
    """获取指定周的起止日期（周一到周日）。

    Args:
        week_offset: 周偏移量，0=本周，-1=上周，1=下周

    Returns:
        dict: {"success": bool, "start": str, "end": str, "week_label": str, "error": str}
    """
    try:
        today = date.today()
        # 计算本周一
        monday = today - timedelta(days=today.weekday())
        # 应用偏移
        target_monday = monday + timedelta(weeks=week_offset)
        target_sunday = target_monday + timedelta(days=6)

        return {
            "success": True,
            "start": target_monday.isoformat(),
            "end": target_sunday.isoformat(),
            "week_label": f"第{target_monday.isocalendar()[1]}周 ({target_monday} ~ {target_sunday})",
            "error": None,
        }
    except Exception as e:
        return {
            "success": False,
            "start": None,
            "end": None,
            "week_label": None,
            "error": str(e),
        }


def format_date(date_str: str, output_format: str = "%Y年%m月%d日") -> dict:
    """将日期字符串转换为指定格式。

    Args:
        date_str: 日期字符串，如 "2026-06-16"
        output_format: 输出格式，默认 "2026年06月16日"

    Returns:
        dict: {"success": bool, "formatted": str, "error": str}
    """
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return {
            "success": True,
            "formatted": dt.strftime(output_format),
            "error": None,
        }
    except Exception as e:
        return {"success": False, "formatted": None, "error": str(e)}


# ============================================================
# LangChain Tool 包装版本（供 Agent 使用）
# ============================================================

from langchain_core.tools import tool


@tool
def tool_get_week_range(week_offset: int = 0) -> str:
    """获取指定周的起止日期。

    当用户提到"本周"、"上周"、"下周"等时间范围时使用此工具。
    适用场景："本周的销售数据"、"上周的报告"、"下周的日程"等。

    Args:
        week_offset: 0=本周, -1=上周, -2=上上周, 1=下周
    """
    result = get_week_range(week_offset)
    if result["success"]:
        return f"{result['week_label']}（{result['start']} 至 {result['end']}）"
    return f"获取日期失败: {result['error']}"
