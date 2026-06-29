"""邮件发送工具 — 基于 smtplib，默认 dry-run 模式"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path


def send_email(
    to: str,
    subject: str,
    body: str,
    attachment: str = None,
    smtp_server: str = "smtp.qq.com",
    smtp_port: int = 587,
    sender_email: str = None,
    sender_password: str = None,
    dry_run: bool = True,
) -> dict:
    """发送邮件（支持附件）。

    Args:
        to: 收件人邮箱
        subject: 邮件主题
        body: 邮件正文（纯文本）
        attachment: 附件文件路径（可选）
        smtp_server: SMTP 服务器地址
        smtp_port: SMTP 端口
        sender_email: 发件人邮箱
        sender_password: SMTP 授权码
        dry_run: True 时只打印邮件内容不实际发送

    Returns:
        dict: {"success": bool, "mode": str, "to": str, "subject": str, "error": str}
    """
    try:
        if dry_run:
            # Dry-run 模式：只打印内容
            print(f"\n{'='*50}")
            print(f"[DRY RUN] 模拟发送邮件")
            print(f"  发件人: {sender_email or '未配置'}")
            print(f"  收件人: {to}")
            print(f"  主题: {subject}")
            print(f"  正文预览: {body[:200]}...")
            if attachment:
                print(f"  附件: {attachment}")
            print(f"{'='*50}\n")
            return {
                "success": True,
                "mode": "dry_run",
                "to": to,
                "subject": subject,
                "error": None,
            }

        # 真实发送
        if not sender_email or not sender_password:
            return {
                "success": False,
                "mode": "real",
                "to": to,
                "subject": subject,
                "error": "缺少发件人邮箱或授权码",
            }

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        if attachment:
            attach_path = Path(attachment)
            if not attach_path.exists():
                return {
                    "success": False,
                    "mode": "real",
                    "to": to,
                    "subject": subject,
                    "error": f"附件不存在: {attachment}",
                }

            with open(attach_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{attach_path.name}"',
                )
                msg.attach(part)

        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to, msg.as_string())

        return {
            "success": True,
            "mode": "real",
            "to": to,
            "subject": subject,
            "error": None,
        }
    except Exception as e:
        return {
            "success": False,
            "mode": "dry_run" if dry_run else "real",
            "to": to,
            "subject": subject,
            "error": str(e),
        }


# ============================================================
# LangChain Tool 包装版本（供 Agent 使用）
# ============================================================

from langchain_core.tools import tool


@tool
def tool_send_email(
    to: str,
    subject: str,
    body: str,
    attachment: str = None,
    dry_run: bool = True,
) -> str:
    """发送邮件，支持附件。

    当用户需要发送邮件时使用此工具。
    适用场景："把周报发给经理"、"邮件通知团队成员"、"发送报表给领导"等。
    注意：默认以 dry_run 模式运行（不真正发送），调试安全。

    Args:
        to: 收件人邮箱地址
        subject: 邮件主题
        body: 邮件正文内容
        attachment: 附件文件路径（可选）
        dry_run: 是否仅模拟发送，默认 True
    """
    result = send_email(
        to=to, subject=subject, body=body, attachment=attachment, dry_run=dry_run
    )
    if result["success"]:
        if result["mode"] == "dry_run":
            return f"[模拟模式] 邮件已准备: 收件人={to}, 主题={subject}"
        return f"邮件已发送: 收件人={to}, 主题={subject}"
    return f"邮件发送失败: {result['error']}"
