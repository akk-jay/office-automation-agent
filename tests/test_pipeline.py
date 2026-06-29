"""LangGraph 管线验证"""
import pytest


def test_pipeline_builds_successfully():
    """管线可以成功编译（build 不报错）"""
    from src.agent.graph_pipeline import build_pipeline

    pipeline = build_pipeline()
    assert pipeline is not None
    # 编译后的图应该有 invoke 方法
    assert hasattr(pipeline, "invoke")
    assert hasattr(pipeline, "stream")


def test_pipeline_runs_simple_instruction():
    """简单指令能走完整个管线（不因数据文件缺失而崩溃）"""
    from src.agent.graph_pipeline import run_pipeline

    result = run_pipeline(
        "帮我看看这周的销售情况，生成周报",
        verbose=False,
    )
    assert result is not None
    # 即使数据文件不存在，管线也应该优雅处理
    # error 字段可能存在（数据文件未找到），但管线本身不应该崩溃
