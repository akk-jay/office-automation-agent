"""文件系统操作工具 — 基于 os/shutil"""

import shutil
import fnmatch
from pathlib import Path
from datetime import datetime


def list_files(directory: str = ".", pattern: str = "*") -> dict:
    """列出指定目录下的文件（不含子目录）。

    Args:
        directory: 目录路径，默认当前目录
        pattern: 文件名过滤模式，如 "*.xlsx"、"report*"

    Returns:
        dict: {"success": bool, "files": list[dict], "file_count": int, "error": str}
    """
    try:
        dir_path = Path(directory)
        if not dir_path.exists():
            return {
                "success": False,
                "files": [],
                "file_count": 0,
                "error": f"目录不存在: {directory}",
            }

        files = []
        for entry in dir_path.iterdir():
            if entry.is_file() and fnmatch.fnmatch(entry.name, pattern):
                stat = entry.stat()
                files.append({
                    "name": entry.name,
                    "path": str(entry.absolute()),
                    "size": stat.st_size,
                    "size_human": _format_size(stat.st_size),
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "extension": entry.suffix.lower(),
                })

        return {
            "success": True,
            "files": sorted(files, key=lambda f: f["name"]),
            "file_count": len(files),
            "error": None,
        }
    except Exception as e:
        return {"success": False, "files": [], "file_count": 0, "error": str(e)}


def classify_files(directory: str = ".") -> dict:
    """将目录下的文件按扩展名分类，移动到对应的子目录中。

    Args:
        directory: 要整理的目录路径

    Returns:
        dict: {"success": bool, "moved_count": int, "details": list[dict], "error": str}
    """
    try:
        dir_path = Path(directory)
        if not dir_path.exists():
            return {
                "success": False,
                "moved_count": 0,
                "details": [],
                "error": f"目录不存在: {directory}",
            }

        # 扩展名 → 分类名映射
        ext_map = {
            ".xlsx": "xlsx文件",
            ".xls": "xls文件",
            ".docx": "docx文件",
            ".doc": "doc文件",
            ".pdf": "pdf文件",
            ".png": "图片文件",
            ".jpg": "图片文件",
            ".jpeg": "图片文件",
            ".gif": "图片文件",
            ".txt": "文本文件",
            ".csv": "csv文件",
            ".pptx": "pptx文件",
            ".zip": "压缩文件",
            ".rar": "压缩文件",
        }

        moved = []
        for entry in dir_path.iterdir():
            if not entry.is_file():
                continue

            ext = entry.suffix.lower()
            category = ext_map.get(ext, f"{ext[1:] if ext else '未知'}文件")

            target_dir = dir_path / category
            target_dir.mkdir(exist_ok=True)

            target_path = target_dir / entry.name
            shutil.move(str(entry), str(target_path))
            moved.append({"file": entry.name, "moved_to": str(target_dir)})

        return {
            "success": True,
            "moved_count": len(moved),
            "details": moved,
            "error": None,
        }
    except Exception as e:
        return {"success": False, "moved_count": 0, "details": [], "error": str(e)}


def _format_size(size_bytes: int) -> str:
    """字节数转人类可读格式"""
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


# ============================================================
# LangChain Tool 包装版本（供 Agent 使用）
# ============================================================

from langchain_core.tools import tool


@tool
def tool_list_files(directory: str = ".", pattern: str = "*") -> str:
    """列出指定目录下的文件。

    当用户需要查看某个目录中有哪些文件时使用。
    适用场景："桌面上有什么文件"、"项目目录下有哪些 Excel"、"帮我找一下文档"等。

    Args:
        directory: 目录路径，默认当前目录
        pattern: 文件名过滤模式，如 "*.xlsx"、"*.docx"
    """
    result = list_files(directory, pattern)
    if result["success"]:
        if result["file_count"] == 0:
            return f"目录 {directory} 中没有匹配 '{pattern}' 的文件"
        lines = [f"目录: {directory}，匹配 '{pattern}' 共 {result['file_count']} 个文件:"]
        for f in result["files"]:
            lines.append(f"  {f['name']} ({f['size_human']}, {f['modified']})")
        return "\n".join(lines)
    return f"列出文件失败: {result['error']}"


@tool
def tool_classify_files(directory: str = ".") -> str:
    """将目录下的文件按扩展名自动分类，移动到对应的子目录中。

    当用户说"整理桌面"、"分类文件"、"把文件归档"时使用此工具。

    Args:
        directory: 要整理的目录路径
    """
    result = classify_files(directory)
    if result["success"]:
        if result["moved_count"] == 0:
            return f"目录 {directory} 中没有需要整理的文件"
        lines = [f"已整理 {result['moved_count']} 个文件:"]
        for detail in result["details"]:
            lines.append(f"  {detail['file']} → {detail['moved_to']}")
        return "\n".join(lines)
    return f"整理失败: {result['error']}"
