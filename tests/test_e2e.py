"""端到端集成测试"""
import sys
sys.path.insert(0, ".")
from src.agent.graph_pipeline import run_pipeline


def test_full_scenario():
    """完整场景: 整理销售数据 → 分析 → 生成周报 → 发送邮件"""
    result = run_pipeline(
        "整理本周销售数据，分析销售趋势，生成周报发送给经理 zhangjingli@company.com",
        verbose=False,
    )
    assert result is not None
    assert result.get("user_intent"), "意图解析不应为空"
    assert result.get("analysis_result"), "分析结果不应为空"
    assert result.get("report_path"), "报告路径不应为空"
    assert result.get("email_status"), "邮件状态不应为空"

    print("端到端测试通过！")
    print(f"  意图: {result['user_intent']}")
    print(f"  报告: {result['report_path']}")
    print(f"  邮件: {result['email_status']}")
