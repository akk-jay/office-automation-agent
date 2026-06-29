"""验证 @tool 包装是否生效"""


def test_all_tools_have_schema():
    """每个工具都应有 name、description、args_schema"""
    from src.tools import ALL_TOOLS

    for t in ALL_TOOLS:
        assert t.name, f"工具 {t} 缺少 name"
        assert t.description, f"工具 {t.name} 缺少 description"
        assert t.args_schema, f"工具 {t.name} 缺少 args_schema"


def test_tool_map_contains_all_tools():
    """TOOL_MAP 应该包含所有工具"""
    from src.tools import ALL_TOOLS, TOOL_MAP

    for t in ALL_TOOLS:
        assert t.name in TOOL_MAP, f"{t.name} 不在 TOOL_MAP 中"
    assert len(TOOL_MAP) == len(ALL_TOOLS)


def test_tool_read_excel_can_be_invoked(tmp_path):
    """工具应该能通过 .invoke() 调用"""
    import openpyxl
    from src.tools import TOOL_MAP

    # 创建临时 Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["姓名", "分数"])
    ws.append(["张三", 95])
    filepath = tmp_path / "test.xlsx"
    wb.save(filepath)

    # 通过工具调用
    result = TOOL_MAP["tool_read_excel"].invoke({"file_path": str(filepath)})

    assert "张三" in result
    assert "95" in result


def test_tool_count():
    """确认总共有 7 个工具"""
    from src.tools import ALL_TOOLS

    assert len(ALL_TOOLS) == 7, f"期望 7 个工具，实际 {len(ALL_TOOLS)} 个"
