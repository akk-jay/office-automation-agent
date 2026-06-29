"""文件操作工具测试"""
import pytest
from pathlib import Path


@pytest.fixture
def sample_dir(tmp_path):
    """创建模拟目录结构"""
    (tmp_path / "report.docx").write_text("报告内容")
    (tmp_path / "data.xlsx").write_text("表格数据")
    (tmp_path / "photo.png").write_text("图片数据")
    (tmp_path / "subdir").mkdir()
    (tmp_path / "subdir" / "notes.txt").write_text("笔记")
    return tmp_path


def test_list_files_no_filter(sample_dir):
    """list_files 不带过滤条件时列出所有文件（不含子目录）"""
    from src.tools.file_tools import list_files

    result = list_files(str(sample_dir))

    assert result["success"] is True
    assert result["file_count"] == 3
    filenames = [f["name"] for f in result["files"]]
    assert "report.docx" in filenames
    assert "data.xlsx" in filenames
    assert "photo.png" in filenames


def test_list_files_with_extension_filter(sample_dir):
    """list_files 按扩展名过滤"""
    from src.tools.file_tools import list_files

    result = list_files(str(sample_dir), pattern="*.docx")

    assert result["success"] is True
    assert result["file_count"] == 1
    assert result["files"][0]["name"] == "report.docx"


def test_classify_files_moves_to_subdirs(sample_dir):
    """classify_files 按扩展名分类移动到子目录"""
    from src.tools.file_tools import classify_files

    result = classify_files(str(sample_dir))

    assert result["success"] is True
    assert result["moved_count"] == 3
    # 确认分类目录已创建
    assert (sample_dir / "docx文件").exists()
    assert (sample_dir / "xlsx文件").exists()
    assert (sample_dir / "图片文件").exists()
    # 确认文件已移动
    assert (sample_dir / "docx文件" / "report.docx").exists()
    assert not (sample_dir / "report.docx").exists()


def test_list_files_directory_not_found():
    """list_files 对不存在的目录应返回失败"""
    from src.tools.file_tools import list_files

    result = list_files("不存在的目录")
    assert result["success"] is False
    assert "不存在" in result["error"]
