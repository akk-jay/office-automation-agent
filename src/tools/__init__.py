"""办公自动化工具集 — 统一导出所有工具"""

# 底层函数（供直接调用）
from src.tools.excel_tools import read_excel, write_excel, summarize_excel
from src.tools.file_tools import list_files, classify_files
from src.tools.email_tools import send_email
from src.tools.date_tools import get_week_range, format_date

# LangChain @tool 包装版本（供 Agent 调用）
from src.tools.excel_tools import (
    tool_read_excel,
    tool_write_excel,
    tool_summarize_excel,
)
from src.tools.file_tools import tool_list_files, tool_classify_files
from src.tools.email_tools import tool_send_email
from src.tools.date_tools import tool_get_week_range

# 所有可供 Agent 调用的工具列表
ALL_TOOLS = [
    tool_read_excel,
    tool_write_excel,
    tool_summarize_excel,
    tool_list_files,
    tool_classify_files,
    tool_send_email,
    tool_get_week_range,
]

# 工具名 → 工具对象的映射（方便按名字查找执行）
TOOL_MAP = {t.name: t for t in ALL_TOOLS}
