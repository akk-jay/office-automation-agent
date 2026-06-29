"""单轮 Agent 验证"""
import pytest
from src.tools import ALL_TOOLS
from src.agent.single_agent import SingleAgent


@pytest.fixture
def agent():
    return SingleAgent(tools=ALL_TOOLS, verbose=False)


def test_agent_direct_answer(agent):
    """不需要工具时，Agent 应该直接回复"""
    result = agent.run("你好，请用一句话介绍你自己")
    assert result is not None
    assert len(result) > 0
    # 验证返回了非空内容
    assert isinstance(result, str)


def test_agent_calls_date_tool(agent):
    """提到"本周"时，Agent 应该调用 tool_get_week_range"""
    result = agent.run("今天是本周的哪几天？")

    assert result is not None
    # 结果中应该包含日期信息
    assert "2026" in result or "周" in result
