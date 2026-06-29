"""邮件工具测试"""
import pytest


def test_send_email_dry_run():
    """dry_run 模式不真正发送，返回成功"""
    from src.tools.email_tools import send_email

    result = send_email(
        to="test@example.com",
        subject="测试邮件",
        body="这是一封测试邮件",
        dry_run=True,
    )

    assert result["success"] is True
    assert result["mode"] == "dry_run"
    assert result["to"] == "test@example.com"


def test_send_email_with_attachment_dry_run():
    """dry_run 模式下附件不存在也不应该报错（只是模拟）"""
    from src.tools.email_tools import send_email

    result = send_email(
        to="test@example.com",
        subject="带附件的测试",
        body="正文内容",
        attachment="不存在的文件.pdf",
        dry_run=True,
    )

    assert result["success"] is True
    assert result["mode"] == "dry_run"
