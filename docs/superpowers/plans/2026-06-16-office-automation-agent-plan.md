# 办公自动化 Agent 实施计划

> **Goal:** 从零构建一个基于 DeepSeek + LangChain + LangGraph 的办公自动化 Agent，通过自然语言驱动 Excel 读写、数据分析、周报生成和邮件发送。

> **Architecture:** 7 步自底向上：环境搭建 → 底层工具函数 → LangChain Tool 封装 → 单轮 Agent → LangGraph 多步编排 → CLI 交互 → 场景集成。每一步依赖前一步，可独立验证。

> **Tech Stack:** Python 3.11, DeepSeek Chat API, LangChain Core, LangGraph, openpyxl, rich, smtplib, python-dotenv

---

## 文件结构总览

```
办公自动化系统/
├── .env
├── requirements.txt
├── src/
│   ├── __init__.py
│   ├── tools/
│   │   ├── __init__.py
│   │   ├── excel_tools.py
│   │   ├── file_tools.py
│   │   ├── email_tools.py
│   │   └── date_tools.py
│   ├── agent/
│   │   ├── __init__.py
│   │   ├── single_agent.py
│   │   └── graph_pipeline.py
│   └── cli/
│       ├── __init__.py
│       └── main.py
├── data/
│   └── (模拟数据)
├── output/
└── tests/
    ├── __init__.py
    ├── test_excel_tools.py
    ├── test_file_tools.py
    ├── test_date_tools.py
    └── test_email_tools.py
```

---

### Task 0: 环境搭建

**说明：** 这是所有后续任务的基础。每一步你必须亲手操作。

- [ ] **Step 1: 检查 Python 版本**

打开终端（Git Bash 或 PowerShell），运行：

```bash
python --version
```

期望输出：`Python 3.11.x`（3.10 也行，3.12 可能有包兼容问题）。

**为什么是 3.11：** LangChain 和 LangGraph 目前对 3.11 的 CI 覆盖最完整，3.12/3.13 部分依赖（如 openpyxl 某些老版本）可能报错。安装 3.11 可以避免浪费时间排查。

如果版本不对，去 https://www.python.org/downloads/ 下载 Python 3.11。

- [ ] **Step 2: 创建虚拟环境**

```bash
cd "C:/Users/29542/Desktop/办公自动化系统"
python -m venv venv
```

**什么是 venv：** 虚拟环境在项目目录下建一个独立的 Python 副本，pip install 的包只对这个项目生效，不会搞乱系统的 Python。

- [ ] **Step 3: 激活虚拟环境**

Git Bash 下：
```bash
source venv/Scripts/activate
```

PowerShell 下：
```powershell
venv\Scripts\Activate.ps1
```

激活成功后，终端提示符前面会出现 `(venv)` 标识。

- [ ] **Step 4: 创建 requirements.txt**

创建文件 `C:\Users\29542\Desktop\办公自动化系统\requirements.txt`：

```
langchain-core>=0.3.0
langchain-openai>=0.2.0
langgraph>=0.2.0
openpyxl>=3.1.0
python-dotenv>=1.0.0
rich>=13.0.0
```

**每个包的作用：**
- `langchain-core` — 提供 Tool、Message、LLM 接口等基础抽象
- `langchain-openai` — 让 LangChain 能对接 OpenAI 兼容接口（DeepSeek 兼容此格式）
- `langgraph` — 构建 StateGraph，实现多步任务编排
- `openpyxl` — 读写 .xlsx 文件
- `python-dotenv` — 从 .env 文件加载环境变量
- `rich` — CLI 美化输出

- [ ] **Step 5: 安装依赖**

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

**为什么用清华镜像：** 国内从 PyPI 官方源下载很慢，清华镜像速度快 10 倍以上。

- [ ] **Step 6: 获取 DeepSeek API Key**

1. 打开 https://platform.deepseek.com
2. 注册/登录（支持微信扫码）
3. 进入"API Keys"页面，点击"创建 API Key"
4. 复制 Key，妥善保存

**费用说明：** DeepSeek 约 ¥1/百万 token。开发调试阶段，整个项目的 API 调用费用不会超过 ¥5。充值 ¥10 足够。

- [ ] **Step 7: 创建 .env 文件**

创建 `C:\Users\29542\Desktop\办公自动化系统\.env`：

```
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
```

**为什么用 .env：** 密钥和代码分离，避免不小心把 Key 传到 GitHub 上。`.env` 文件要加入 `.gitignore`。

- [ ] **Step 8: 创建项目目录结构**

```bash
mkdir -p src/tools src/agent src/cli data output tests
touch src/__init__.py src/tools/__init__.py src/agent/__init__.py src/cli/__init__.py tests/__init__.py
```

- [ ] **Step 9: 验证 DeepSeek 连通性**

创建并运行测试脚本 `tests/test_connection.py`：

```python
import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI

load_dotenv()

llm = ChatOpenAI(
    model="deepseek-chat",
    base_url=os.getenv("DEEPSEEK_BASE_URL"),
    api_key=os.getenv("DEEPSEEK_API_KEY"),
)

response = llm.invoke("你好，请用一句话介绍你自己")
print(response.content)
```

运行：
```bash
python tests/test_connection.py
```

期望输出：DeepSeek 返回一段中文自我介绍。

**如果报错：**
- `ModuleNotFoundError: No module named 'langchain_openai'` → 检查 venv 是否激活
- `AuthenticationError` → 检查 API Key 是否正确
- `ConnectionError` → 检查网络是否能访问 api.deepseek.com

---

### Task 1: 底层工具函数 — Excel 读写

**文件：** 创建 `src/tools/excel_tools.py`、`tests/test_excel_tools.py`

**目标：** 用 openpyxl 实现三个 Excel 操作函数。这一步跟 AI 无关，纯 Python。

- [ ] **Step 1: 编写测试**

创建 `tests/test_excel_tools.py`：

```python
import pytest
import openpyxl
from pathlib import Path

# 创建测试用的临时 Excel
@pytest.fixture
def sample_excel(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["日期", "产品", "销量", "单价", "销售额"])
    ws.append(["2026-06-16", "智能门锁 Pro", 12, 2999, 35988])
    ws.append(["2026-06-17", "智能门锁 Pro", 8, 2999, 23992])
    filepath = tmp_path / "test.xlsx"
    wb.save(filepath)
    return filepath


def test_read_excel_returns_all_rows(sample_excel):
    """read_excel 应该返回 Excel 中的所有数据行（不含表头）"""
    from src.tools.excel_tools import read_excel
    
    result = read_excel(str(sample_excel))
    
    assert len(result["data"]) == 2
    assert result["data"][0]["日期"] == "2026-06-16"
    assert result["data"][0]["销量"] == 12


def test_summarize_excel_returns_stats(sample_excel):
    """summarize_excel 应该返回行数、列名等基本统计"""
    from src.tools.excel_tools import summarize_excel
    
    result = summarize_excel(str(sample_excel))
    
    assert result["row_count"] == 2
    assert "日期" in result["columns"]
    assert result["sheet_name"] == "Sheet1"


def test_write_excel_creates_file(tmp_path):
    """write_excel 应该能创建新 Excel 文件"""
    from src.tools.excel_tools import write_excel
    
    outpath = tmp_path / "output.xlsx"
    data = [
        {"姓名": "张三", "分数": 95},
        {"姓名": "李四", "分数": 87},
    ]
    result = write_excel(data, str(outpath))
    
    assert result["success"] is True
    assert Path(outpath).exists()
    
    # 验证内容
    wb = openpyxl.load_workbook(outpath)
    ws = wb.active
    assert ws["A1"].value == "姓名"
    assert ws["A2"].value == "张三"
    assert ws["B2"].value == 95
```

**为什么先写测试？** TDD（测试驱动开发）的思路：先定义"正确的结果是什么"，再写代码去满足它。这样你不会写出一段"跑得通但不知道对不对"的代码。

- [ ] **Step 2: 运行测试，确认失败**

```bash
python -m pytest tests/test_excel_tools.py -v
```

期望：3 个测试都报 `ModuleNotFoundError` 或 `ImportError`，因为 `excel_tools.py` 还不存在。

- [ ] **Step 3: 实现 excel_tools.py**

创建 `src/tools/excel_tools.py`：

```python
"""Excel 读写工具 — 基于 openpyxl"""

import openpyxl
from pathlib import Path


def read_excel(file_path: str, sheet_name: str = None) -> dict:
    """读取 Excel 文件并返回所有数据行。

    Args:
        file_path: Excel 文件的路径
        sheet_name: 要读取的工作表名称，默认读取第一个工作表

    Returns:
        dict: {"success": bool, "data": list[dict], "row_count": int, "columns": list, "sheet_name": str, "error": str}
    """
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {"success": False, "data": [], "row_count": 0, "columns": [], "error": "文件为空"}

        headers = [str(h) if h else f"col_{i}" for i, h in enumerate(rows[0])]
        data = []
        for row in rows[1:]:
            record = {}
            for i, value in enumerate(row):
                if i < len(headers):
                    record[headers[i]] = value
            data.append(record)

        return {
            "success": True,
            "data": data,
            "row_count": len(data),
            "columns": headers,
            "sheet_name": ws.title,
            "error": None,
        }
    except FileNotFoundError:
        return {"success": False, "data": [], "row_count": 0, "columns": [], "error": f"文件不存在: {file_path}"}
    except Exception as e:
        return {"success": False, "data": [], "row_count": 0, "columns": [], "error": str(e)}


def write_excel(data: list, file_path: str, sheet_name: str = "Sheet1") -> dict:
    """将数据列表写入 Excel 文件。

    Args:
        data: 数据列表，每条记录是一个 dict，key 为列名
        file_path: 输出文件路径
        sheet_name: 工作表名称

    Returns:
        dict: {"success": bool, "file_path": str, "row_count": int, "error": str}
    """
    try:
        if not data:
            return {"success": False, "file_path": file_path, "row_count": 0, "error": "没有数据可写入"}

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        # 写表头
        headers = list(data[0].keys())
        ws.append(headers)

        # 写数据
        for record in data:
            row = [record.get(h) for h in headers]
            ws.append(row)

        Path(file_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(file_path)

        return {
            "success": True,
            "file_path": str(Path(file_path).absolute()),
            "row_count": len(data),
            "error": None,
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "row_count": 0, "error": str(e)}


def summarize_excel(file_path: str) -> dict:
    """获取 Excel 文件的概览信息：行数、列名、基本数值统计。

    Args:
        file_path: Excel 文件路径

    Returns:
        dict: {"success": bool, "row_count": int, "columns": list, "sheet_name": str,
               "numeric_stats": dict, "error": str}
    """
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {"success": False, "row_count": 0, "columns": [], "error": "文件为空"}

        headers = [str(h) if h else f"col_{i}" for i, h in enumerate(rows[0])]
        data_rows = rows[1:]

        # 统计数值列
        numeric_stats = {}
        for col_idx, header in enumerate(headers):
            values = [row[col_idx] for row in data_rows if col_idx < len(row) and isinstance(row[col_idx], (int, float))]
            if values:
                numeric_stats[header] = {
                    "count": len(values),
                    "sum": sum(values),
                    "avg": round(sum(values) / len(values), 2),
                    "min": min(values),
                    "max": max(values),
                }

        return {
            "success": True,
            "row_count": len(data_rows),
            "columns": headers,
            "sheet_name": ws.title,
            "numeric_stats": numeric_stats,
            "error": None,
        }
    except FileNotFoundError:
        return {"success": False, "row_count": 0, "columns": [], "error": f"文件不存在: {file_path}"}
    except Exception as e:
        return {"success": False, "row_count": 0, "columns": [], "error": str(e)}
```

**关键设计决策：**
- 所有返回值统一为 `dict` 格式，包含 `success` 字段。后面 LLM 可以一眼判断操作成功还是失败。
- `data_only=True`：读取时只取计算后的值，不取公式。
- 异常全捕获，不要让一个工具的错误导致整个 Agent 崩溃。

- [ ] **Step 4: 运行测试验证**

```bash
python -m pytest tests/test_excel_tools.py -v
```

期望：3 个测试全部 PASS。

---

### Task 2: 底层工具函数 — 文件操作

**文件：** 创建 `src/tools/file_tools.py`、`tests/test_file_tools.py`

- [ ] **Step 1: 编写测试**

创建 `tests/test_file_tools.py`：

```python
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
    """list_files 不带过滤条件时列出所有文件"""
    from src.tools.file_tools import list_files

    result = list_files(str(sample_dir))

    assert result["success"] is True
    assert result["file_count"] == 3  # 不含子目录中的文件
    filenames = [f["name"] for f in result["files"]]
    assert "report.docx" in filenames
    assert "data.xlsx" in filenames


def test_list_files_with_extension_filter(sample_dir):
    """list_files 按扩展名过滤"""
    from src.tools.file_tools import list_files

    result = list_files(str(sample_dir), pattern="*.docx")

    assert result["file_count"] == 1
    assert result["files"][0]["name"] == "report.docx"


def test_classify_files_moves_to_subdirs(sample_dir):
    """classify_files 按扩展名分类移动到子目录"""
    from src.tools.file_tools import classify_files

    result = classify_files(str(sample_dir))

    assert result["success"] is True
    # 确认分类目录已创建
    assert (sample_dir / "docx文件").exists()
    assert (sample_dir / "xlsx文件").exists()
    assert (sample_dir / "png文件").exists()
    # 确认文件已移动
    assert (sample_dir / "docx文件" / "report.docx").exists()
    assert not (sample_dir / "report.docx").exists()
```

- [ ] **Step 2: 运行测试确认失败**

```bash
python -m pytest tests/test_file_tools.py -v
```

- [ ] **Step 3: 实现 file_tools.py**

创建 `src/tools/file_tools.py`：

```python
"""文件系统操作工具 — 基于 os/shutil"""

import os
import shutil
import fnmatch
from pathlib import Path
from datetime import datetime


def list_files(directory: str = ".", pattern: str = "*") -> dict:
    """列出指定目录下的文件（不含子目录）。

    Args:
        directory: 目录路径，默认当前目录
        pattern: 文件名过滤模式，如 "*.xlsx"、"report*"，默认列出所有文件

    Returns:
        dict: {"success": bool, "files": list[dict], "file_count": int, "error": str}
    """
    try:
        dir_path = Path(directory)
        if not dir_path.exists():
            return {"success": False, "files": [], "file_count": 0, "error": f"目录不存在: {directory}"}

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

    Excel 文件 → xlsx文件/，Word 文件 → docx文件/，图片 → 图片文件/，等等。

    Args:
        directory: 要整理的目录路径

    Returns:
        dict: {"success": bool, "moved_count": int, "details": list[dict], "error": str}
    """
    try:
        dir_path = Path(directory)
        if not dir_path.exists():
            return {"success": False, "moved_count": 0, "details": [], "error": f"目录不存在: {directory}"}

        # 扩展名 → 分类名 映射
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
```

- [ ] **Step 4: 运行测试验证**

```bash
python -m pytest tests/test_file_tools.py -v
```

---

### Task 3: 底层工具函数 — 日期工具 + 邮件工具

**文件：** 创建 `src/tools/date_tools.py`、`src/tools/email_tools.py` 及对应测试

- [ ] **Step 1: 日期工具测试**

创建 `tests/test_date_tools.py`：

```python
from datetime import date


def test_get_week_range_this_week():
    """get_week_range 返回本周的起止日期"""
    from src.tools.date_tools import get_week_range

    result = get_week_range(week_offset=0)

    assert result["success"] is True
    assert "start" in result
    assert "end" in result
    assert "week_label" in result
    # 周一起始
    d = date.fromisoformat(result["start"])
    assert d.weekday() == 0  # Monday


def test_get_week_range_last_week():
    """get_week_range week_offset=-1 返回上周"""
    from src.tools.date_tools import get_week_range

    this_week = get_week_range(week_offset=0)
    last_week = get_week_range(week_offset=-1)

    assert last_week["start"] < this_week["start"]
```

- [ ] **Step 2: 实现 date_tools.py**

创建 `src/tools/date_tools.py`：

```python
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
        return {"success": False, "start": None, "end": None, "week_label": None, "error": str(e)}


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
```

- [ ] **Step 3: 实现 email_tools.py（dry-run 模式）**

创建 `src/tools/email_tools.py`：

```python
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
        body: 邮件正文（纯文本或 HTML）
        attachment: 附件文件路径（可选）
        smtp_server: SMTP 服务器地址
        smtp_port: SMTP 端口
        sender_email: 发件人邮箱
        sender_password: SMTP 授权码
        dry_run: True 时只打印邮件内容不实际发送（调试用）

    Returns:
        dict: {"success": bool, "mode": str, "to": str, "subject": str, "error": str}
    """
    try:
        if dry_run:
            # Dry-run 模式：只打印内容
            print(f"\n{'='*50}")
            print(f"📧 [DRY RUN] 模拟发送邮件")
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
            return {"success": False, "mode": "real", "to": to, "subject": subject, "error": "缺少发件人邮箱或授权码"}

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        if attachment:
            attach_path = Path(attachment)
            if not attach_path.exists():
                return {"success": False, "mode": "real", "to": to, "subject": subject, "error": f"附件不存在: {attachment}"}

            with open(attach_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{attach_path.name}"')
                msg.attach(part)

        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to, msg.as_string())

        return {"success": True, "mode": "real", "to": to, "subject": subject, "error": None}
    except Exception as e:
        return {"success": False, "mode": "dry_run" if dry_run else "real", "to": to, "subject": subject, "error": str(e)}
```

- [ ] **Step 4: 运行所有工具函数测试**

```bash
python -m pytest tests/ -v
```

期望：全部 PASS。

---

### Task 4: LangChain Tool 封装

**文件：** 修改 `src/tools/__init__.py`，各工具模块加 `@tool` 装饰

**目标：** 把 Step 1-3 的普通函数用 `@tool` 装饰器包装，LLM 能看到每个工具的名称、描述、参数。

- [ ] **Step 1: 理解 @tool 做了什么**

`@tool` 装饰器做的事情：读取函数的类型标注和 docstring → 转成 JSON Schema → 当 LLM 收到用户消息时，工具列表随请求一起发送 → LLM 根据 schema 判断"该不该调用这个工具"。

```python
# 不加 @tool — 这就是个普通函数
def read_excel(file_path: str) -> dict:
    """读取 Excel"""
    ...

# 加上 @tool — LLM 能看到它的名字、描述、参数类型
from langchain_core.tools import tool

@tool
def read_excel(file_path: str) -> str:
    """读取 Excel 文件并返回所有数据行。当用户需要查看、分析或汇总 Excel 数据时使用此工具。"""
    ...
```

两个关键变化：
1. 返回值从 `dict` 变成 `str`（LLM 更容易理解文本）
2. docstring 要写清"什么时候用这个工具"（这是给 LLM 看的说明书）

- [ ] **Step 2: 改造 excel_tools.py**

在 `src/tools/excel_tools.py` 文件末尾追加 LangChain 包装版本：

```python
# ============================================================
# LangChain Tool 包装版本（供 Agent 使用）
# ============================================================

from langchain_core.tools import tool


@tool
def tool_read_excel(file_path: str, sheet_name: str = "Sheet1") -> str:
    """读取 Excel 文件并返回所有数据行。

    当用户需要查看、分析或汇总 Excel 表格数据时使用此工具。
    适用场景：用户提到"看下销售数据"、"汇总报表"、"这个 Excel 里有什么"等。

    Args:
        file_path: Excel 文件的完整路径
        sheet_name: 要读取的工作表名称，默认 Sheet1
    """
    result = read_excel(file_path, sheet_name)
    if result["success"]:
        lines = [f"工作表: {result['sheet_name']}，共 {result['row_count']} 行"]
        lines.append(f"列名: {', '.join(result['columns'])}")
        lines.append("---数据预览（前 20 行）---")
        for i, row in enumerate(result["data"][:20]):
            lines.append(str(row))
        if result["row_count"] > 20:
            lines.append(f"... 还有 {result['row_count'] - 20} 行未显示")
        return "\n".join(lines)
    return f"读取失败: {result['error']}"


@tool
def tool_write_excel(data_json: str, file_path: str, sheet_name: str = "Sheet1") -> str:
    """将数据写入 Excel 文件。

    当用户需要保存、导出、或生成 Excel 报表时使用此工具。
    适用场景："生成周报"、"导出数据"、"保存为 Excel"等。

    Args:
        data_json: JSON 格式的数据，如 '[{"姓名":"张三","分数":95}, {"姓名":"李四","分数":87}]'
        file_path: 输出文件的完整路径
        sheet_name: 工作表名称，默认 Sheet1
    """
    import json
    try:
        data = json.loads(data_json)
    except json.JSONDecodeError:
        return "写入失败: 数据格式不正确，需要 JSON 格式"

    result = write_excel(data, file_path, sheet_name)
    if result["success"]:
        return f"已成功写入 {result['row_count']} 行数据到 {result['file_path']}"
    return f"写入失败: {result['error']}"


@tool
def tool_summarize_excel(file_path: str) -> str:
    """获取 Excel 文件的概览信息（行数、列名、数值统计）。

    当用户想要快速了解 Excel 文件的结构和内容概况时使用。
    适用场景："这个表有多少行"、"销售额大概是多少"、"有哪些列"等。

    Args:
        file_path: Excel 文件的完整路径
    """
    result = summarize_excel(file_path)
    if result["success"]:
        lines = [
            f"文件概览: {result['sheet_name']}",
            f"数据行数: {result['row_count']}",
            f"列名: {', '.join(result['columns'])}",
        ]
        if result.get("numeric_stats"):
            lines.append("数值列统计:")
            for col, stats in result["numeric_stats"].items():
                lines.append(f"  {col}: 总计={stats['sum']}, 平均={stats['avg']}, 最大={stats['max']}, 最小={stats['min']}")
        return "\n".join(lines)
    return f"概览失败: {result['error']}"
```

- [ ] **Step 3: 改造 file_tools.py**

在 `src/tools/file_tools.py` 末尾追加：

```python
# ============================================================
# LangChain Tool 包装版本
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
```

- [ ] **Step 4: 改造 email_tools.py**

在 `src/tools/email_tools.py` 末尾追加：

```python
# ============================================================
# LangChain Tool 包装版本
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
```

- [ ] **Step 5: 改造 date_tools.py**

在 `src/tools/date_tools.py` 末尾追加：

```python
# ============================================================
# LangChain Tool 包装版本
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
```

- [ ] **Step 6: 更新 __init__.py 汇集所有工具**

修改 `src/tools/__init__.py`：

```python
"""办公自动化工具集 — 统一导出所有 LangChain Tool"""

from src.tools.excel_tools import (
    read_excel,
    write_excel,
    summarize_excel,
    tool_read_excel,
    tool_write_excel,
    tool_summarize_excel,
)
from src.tools.file_tools import (
    list_files,
    classify_files,
    tool_list_files,
    tool_classify_files,
)
from src.tools.email_tools import (
    send_email,
    tool_send_email,
)
from src.tools.date_tools import (
    get_week_range,
    format_date,
    tool_get_week_range,
)

# 所有可供 Agent 调用的工具
ALL_TOOLS = [
    tool_read_excel,
    tool_write_excel,
    tool_summarize_excel,
    tool_list_files,
    tool_classify_files,
    tool_send_email,
    tool_get_week_range,
]

# 工具名 → 工具对象的映射（方便按名字查找）
TOOL_MAP = {t.name: t for t in ALL_TOOLS}
```

- [ ] **Step 7: 验证 @tool 包装是否生效**

创建 `tests/test_tool_wrappers.py`：

```python
def test_all_tools_have_schema():
    """检查所有工具是否都有有效的 name 和 description"""
    from src.tools import ALL_TOOLS
    
    for tool in ALL_TOOLS:
        assert tool.name, f"工具 {tool} 缺少 name"
        assert tool.description, f"工具 {tool.name} 缺少 description"
        assert tool.args_schema, f"工具 {tool.name} 缺少 args_schema"
        print(f"✓ {tool.name}: {tool.description[:50]}...")


def test_tool_map_contains_all_tools():
    """TOOL_MAP 应该包含所有工具"""
    from src.tools import ALL_TOOLS, TOOL_MAP
    
    for tool in ALL_TOOLS:
        assert tool.name in TOOL_MAP, f"{tool.name} 不在 TOOL_MAP 中"
```

运行：
```bash
python -m pytest tests/test_tool_wrappers.py -v
```

---

### Task 5: 单轮 Agent

**文件：** 创建 `src/agent/__init__.py`、`src/agent/single_agent.py`

**目标：** DeepSeek + bind_tools，实现最简单的 "用户一句话 → Agent 选工具 → 执行 → 返回" 循环。

- [ ] **Step 1: 理解核心流程**

```
HumanMessage("读一下 sales.xlsx")
        │
        ▼
   llm_with_tools.invoke([msg])
        │
        ▼
   AIMessage(tool_calls=[{name: "tool_read_excel", args: {file_path: "sales.xlsx"}}])
        │
        ▼
   执行 tool_read_excel("sales.xlsx") → "共 120 行..."
        │
        ▼
   ToolMessage(content="共 120 行...", tool_call_id="...")
        │
        ▼
   llm_with_tools.invoke([HumanMessage, AIMessage, ToolMessage])
        │
        ▼
   AIMessage(content="该表包含 120 条销售记录，列包括...")
```

- [ ] **Step 2: 实现单轮 Agent**

创建 `src/agent/single_agent.py`：

```python
"""单轮 Agent — 一次对话中自动选择并执行工具"""

import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage

load_dotenv()


class SingleAgent:
    """最简 Agent：LLM + 工具列表，单轮对话自动调用工具。"""

    def __init__(self, tools: list, verbose: bool = True):
        self.tools = tools
        self.tool_map = {t.name: t for t in tools}
        self.verbose = verbose

        # 连接 DeepSeek
        self.llm = ChatOpenAI(
            model="deepseek-chat",
            base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1"),
            api_key=os.getenv("DEEPSEEK_API_KEY"),
            temperature=0,  # 温度设为 0，减少随机性，工具调用更稳定
        )

        # 绑定工具 — 这是关键：bind_tools 让 LLM 知道有哪些工具可用
        self.llm_with_tools = self.llm.bind_tools(self.tools)

    def run(self, user_input: str) -> str:
        """处理用户输入，自动调用工具并返回最终结果。

        Args:
            user_input: 用户的自然语言指令

        Returns:
            Agent 的最终回复文本
        """
        messages = [HumanMessage(content=user_input)]

        if self.verbose:
            print(f"\n{'─'*40}")
            print(f"🧑 用户: {user_input}")

        # 第一轮：LLM 决定调哪个工具
        response = self.llm_with_tools.invoke(messages)

        # 检查是否有工具调用
        if not response.tool_calls:
            # LLM 觉得不需要工具，直接回答
            if self.verbose:
                print(f"🤖 Agent: {response.content}")
            return response.content

        # 有工具调用 → 逐个执行
        if self.verbose:
            print(f"🔧 LLM 决定调用 {len(response.tool_calls)} 个工具:")

        messages.append(response)

        for tool_call in response.tool_calls:
            tool_name = tool_call["name"]
            tool_args = tool_call["args"]
            tool_id = tool_call["id"]

            if self.verbose:
                print(f"  → {tool_name}({tool_args})")

            # 执行工具
            tool_func = self.tool_map[tool_name]
            tool_result = tool_func.invoke(tool_args)

            if self.verbose:
                print(f"  ← 结果: {tool_result[:100]}...")

            # 将工具结果以 ToolMessage 形式追加到对话
            messages.append(ToolMessage(content=tool_result, tool_call_id=tool_id))

        # 第二轮：LLM 根据工具返回的结果，生成最终回复
        final_response = self.llm_with_tools.invoke(messages)

        if self.verbose:
            print(f"🤖 Agent: {final_response.content}")

        return final_response.content


# 便捷函数
def create_agent(tools: list) -> SingleAgent:
    """创建单轮 Agent 实例"""
    return SingleAgent(tools=tools)
```

- [ ] **Step 3: 编写验证脚本**

创建 `tests/test_single_agent.py`：

```python
"""验证单轮 Agent 能否正确调用工具"""
import pytest
from src.tools import ALL_TOOLS
from src.agent.single_agent import SingleAgent


@pytest.fixture
def agent():
    return SingleAgent(tools=ALL_TOOLS, verbose=False)


def test_agent_direct_answer(agent):
    """不需要工具时直接回答"""
    result = agent.run("你好，请介绍一下你自己")
    assert result is not None
    assert len(result) > 0


def test_agent_calls_date_tool(agent):
    """提到"本周"时应调用 date 工具"""
    result = agent.run("今天是本周的哪几天？")
    assert result is not None
    # 因为 verbose=False，我们无法直接看到 tool_calls，
    # 但结果应该包含日期信息
    assert "2026" in result or "周" in result
```

运行：
```bash
python -m pytest tests/test_single_agent.py -v
```

---

### Task 6: LangGraph 多步编排

**文件：** 创建 `src/agent/graph_pipeline.py`

**目标：** 用 StateGraph 搭建 5 节点管线，让复杂指令按顺序自动拆解执行。

- [ ] **Step 1: 理解 StateGraph 核心概念**

StateGraph 管三件事：
1. **State（状态）** — 一个 dict，在节点间流转，每个节点可以读/写它
2. **Nodes（节点）** — 处理函数，接收 State，返回 State 的部分更新
3. **Edges（边）** — 节点间的连线，普通边固定连接，条件边根据 State 决定走哪条

- [ ] **Step 2: 实现 graph_pipeline.py**

创建 `src/agent/graph_pipeline.py`：

```python
"""LangGraph 多步任务编排 — 将复杂指令拆解为有序步骤"""

import os
from typing import TypedDict, Annotated, Literal
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage
from langgraph.graph import StateGraph, END
from langgraph.graph.message import add_messages

from src.tools import ALL_TOOLS, TOOL_MAP

load_dotenv()


# ============================================================
# 1. 状态定义 — 在节点间流转的数据结构
# ============================================================

class AgentState(TypedDict):
    """Agent 的全局状态，每个节点都可以读写这些字段"""
    messages: Annotated[list, add_messages]  # 对话历史（Annotated + add_messages = 自动追加而非覆盖）
    user_intent: str       # 用户意图描述
    sub_tasks: list        # 拆解后的子任务列表
    week_start: str        # 周报的起始日期
    week_end: str          # 周报的结束日期
    data_summary: str      # 数据概要
    analysis_result: str   # LLM 分析结论
    report_path: str       # 生成的周报文件路径
    email_status: str      # 邮件发送状态
    error: str             # 错误信息（如有）
    next_action: str       # 下一步动作（条件路由用）


# ============================================================
# 2. LLM 实例
# ============================================================

llm = ChatOpenAI(
    model="deepseek-chat",
    base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1"),
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    temperature=0,
)

llm_with_tools = llm.bind_tools(ALL_TOOLS)


# ============================================================
# 3. 节点函数 — 每个节点做一件事
# ============================================================

def node_intent_parser(state: AgentState) -> dict:
    """节点1: 意图解析 — 理解用户要做什么"""
    last_msg = state["messages"][-1] if state["messages"] else None
    user_input = last_msg.content if last_msg else ""

    prompt = f"""分析以下用户指令，提取关键信息并以 JSON 格式返回。

用户指令: "{user_input}"

返回格式（只返回 JSON，不要其他文字）:
{{
    "intent": "简要描述用户的意图",
    "sub_tasks": ["子任务1", "子任务2", "子任务3"],
    "needs_date_range": true/false,
    "needs_excel_read": true/false,
    "needs_analysis": true/false,
    "needs_report": true/false,
    "needs_email": true/false,
    "email_recipient": "收件人邮箱或null"
}}
"""
    response = llm.invoke([HumanMessage(content=prompt)])
    
    # 解析 JSON（容错处理）
    import json
    try:
        intent_data = json.loads(response.content)
    except json.JSONDecodeError:
        # 尝试从响应中提取 JSON
        import re
        match = re.search(r'\{.*\}', response.content, re.DOTALL)
        if match:
            intent_data = json.loads(match.group())
        else:
            return {"error": f"无法解析意图: {response.content}", "next_action": "error"}

    return {
        "user_intent": intent_data.get("intent", ""),
        "sub_tasks": intent_data.get("sub_tasks", []),
        "next_action": "continue",
        "error": None,
    }


def node_data_fetcher(state: AgentState) -> dict:
    """节点2: 数据获取 — 调用工具获取日期范围和 Excel 数据"""
    results = []

    # 先获取日期范围
    date_tool = TOOL_MAP["tool_get_week_range"]
    date_result = date_tool.invoke({"week_offset": 0})
    results.append(f"[日期范围] {date_result}")

    # 尝试读取数据文件
    excel_tool = TOOL_MAP["tool_summarize_excel"]
    summary_result = excel_tool.invoke({"file_path": "data/sales_2026W25.xlsx"})
    results.append(f"[数据概要] {summary_result}")

    return {
        "data_summary": "\n".join(results),
        "next_action": "continue",
    }


def node_analyzer(state: AgentState) -> dict:
    """节点3: 分析处理 — LLM 分析数据，生成洞察"""
    prompt = f"""你是一位数据分析师。根据以下数据概要，写一段简短的销售分析总结（200字以内）。

数据概要:
{state.get("data_summary", "无数据")}

请分析: 销售趋势、亮点、需要注意的问题。
"""
    response = llm.invoke([HumanMessage(content=prompt)])
    return {"analysis_result": response.content, "next_action": "continue"}


def node_report_generator(state: AgentState) -> dict:
    """节点4: 结果输出 — 生成周报 Excel"""
    import json
    
    analysis = state.get("analysis_result", "无分析结果")
    
    # 构造周报数据
    report_data = [
        {"项目": "周报", "内容": "销售数据周报"},
        {"项目": "分析结论", "内容": analysis},
        {"项目": "数据概要", "内容": state.get("data_summary", "")},
    ]

    write_tool = TOOL_MAP["tool_write_excel"]
    result = write_tool.invoke({
        "data_json": json.dumps(report_data, ensure_ascii=False),
        "file_path": "output/周报_最新.xlsx",
        "sheet_name": "周报",
    })

    return {"report_path": "output/周报_最新.xlsx", "next_action": "continue"}


def node_sender(state: AgentState) -> dict:
    """节点5: 消息推送 — 发送邮件"""
    analysis = state.get("analysis_result", "")
    report_path = state.get("report_path", "")

    # 从用户原始消息中提取收件人
    email_tool = TOOL_MAP["tool_send_email"]
    result = email_tool.invoke({
        "to": "manager@company.com",
        "subject": "销售数据周报",
        "body": f"您好，以下是本周销售数据周报。\n\n分析结论:\n{analysis}\n\n周报附件: {report_path}",
        "attachment": report_path,
        "dry_run": True,
    })

    return {"email_status": result, "next_action": "done"}


# ============================================================
# 4. 条件路由 — 决定下一步走哪个节点
# ============================================================

def router(state: AgentState) -> Literal["data_fetcher", "end"]:
    """从意图解析到数据获取的路由判断"""
    if state.get("error"):
        return "end"
    return "data_fetcher"


def router_after_fetch(state: AgentState) -> Literal["analyzer", "report_generator", "end"]:
    """数据获取后的路由"""
    if state.get("error"):
        return "end"
    return "analyzer"


# ============================================================
# 5. 构建图
# ============================================================

def build_pipeline() -> StateGraph:
    """构建 LangGraph 任务管线"""
    workflow = StateGraph(AgentState)

    # 添加节点
    workflow.add_node("intent_parser", node_intent_parser)
    workflow.add_node("data_fetcher", node_data_fetcher)
    workflow.add_node("analyzer", node_analyzer)
    workflow.add_node("report_generator", node_report_generator)
    workflow.add_node("sender", node_sender)

    # 设置入口
    workflow.set_entry_point("intent_parser")

    # 连线: 意图解析 → (条件) → 数据获取 → 分析 → 报告生成 → 发送 → 结束
    workflow.add_conditional_edges("intent_parser", router, {
        "data_fetcher": "data_fetcher",
        "end": END,
    })
    workflow.add_conditional_edges("data_fetcher", router_after_fetch, {
        "analyzer": "analyzer",
        "report_generator": "report_generator",
        "end": END,
    })
    workflow.add_edge("analyzer", "report_generator")
    workflow.add_edge("report_generator", "sender")
    workflow.add_edge("sender", END)

    return workflow.compile()


# ============================================================
# 6. 便捷函数
# ============================================================

def run_pipeline(user_input: str, verbose: bool = True) -> dict:
    """运行完整的 LangGraph 管线"""
    pipeline = build_pipeline()
    
    initial_state = {
        "messages": [HumanMessage(content=user_input)],
    }

    if verbose:
        print(f"\n{'='*50}")
        print(f"🚀 启动 Agent 管线")
        print(f"📝 用户指令: {user_input}")
        print(f"{'='*50}")

    # 使用 stream 模式，每执行一个节点就输出一次
    for step_output in pipeline.stream(initial_state):
        node_name = list(step_output.keys())[0]
        node_state = step_output[node_name]

        if verbose:
            if node_state.get("user_intent"):
                print(f"\n🎯 [意图解析] {node_state['user_intent']}")
            if node_state.get("data_summary"):
                print(f"\n📊 [数据获取]\n{node_state['data_summary']}")
            if node_state.get("analysis_result"):
                print(f"\n📈 [分析处理]\n{node_state['analysis_result']}")
            if node_state.get("report_path"):
                print(f"\n📄 [周报生成] {node_state['report_path']}")
            if node_state.get("email_status"):
                print(f"\n📧 [邮件发送] {node_state['email_status']}")
            if node_state.get("error"):
                print(f"\n❌ [错误] {node_state['error']}")

    # 获取最终状态
    final_state = pipeline.invoke(initial_state)

    if verbose:
        print(f"\n{'='*50}")
        print(f"✅ 管线执行完成")
        print(f"{'='*50}\n")

    return final_state
```

- [ ] **Step 3: 验证管线能跑通**

创建 `tests/test_pipeline.py`：

```python
def test_pipeline_builds_successfully():
    """管线可以成功编译"""
    from src.agent.graph_pipeline import build_pipeline
    pipeline = build_pipeline()
    assert pipeline is not None


def test_pipeline_runs_simple_instruction():
    """简单指令可以走完管线"""
    from src.agent.graph_pipeline import run_pipeline
    
    result = run_pipeline(
        "帮我看看这周的销售情况，生成周报", 
        verbose=False
    )
    assert result is not None
    assert result.get("error") is None
```

先故意跑一次验证 build 成功（数据文件还没创建，数据获取节点会报错但管线本身应该不崩）。

---

### Task 7: CLI 交互层

**文件：** 创建 `src/cli/main.py`

- [ ] **Step 1: 实现 CLI**

创建 `src/cli/main.py`：

```python
"""办公自动化 Agent — 命令行交互界面"""

import sys
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.markdown import Markdown

from src.agent.graph_pipeline import run_pipeline

console = Console()

HELP_TEXT = """
**可用命令:**
- 直接输入自然语言指令，Agent 自动执行
- `/help` — 显示此帮助
- `/tools` — 列出所有可用工具
- `/history` — 显示对话历史（本次会话）
- `/quit` 或 `/exit` — 退出
"""


def show_banner():
    console.print(Panel.fit(
        "[bold blue]办公自动化 Agent v1.0[/bold blue]\n"
        "[dim]基于 DeepSeek + LangGraph | 输入 /help 查看帮助[/dim]",
        border_style="blue",
    ))


def show_tools():
    from src.tools import ALL_TOOLS
    table = Table(title="可用工具", style="cyan")
    table.add_column("工具名", style="green")
    table.add_column("描述", style="white")
    for tool in ALL_TOOLS:
        table.add_row(tool.name, tool.description[:80])
    console.print(table)


def main():
    show_banner()
    history = []

    while True:
        try:
            user_input = console.input("\n[bold green]🧑 你:[/bold green] ")
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]再见！[/dim]")
            break

        user_input = user_input.strip()
        if not user_input:
            continue

        # 处理内置命令
        if user_input in ("/quit", "/exit"):
            console.print("[dim]再见！[/dim]")
            break
        if user_input == "/help":
            console.print(Markdown(HELP_TEXT))
            continue
        if user_input == "/tools":
            show_tools()
            continue
        if user_input == "/history":
            if not history:
                console.print("[dim]暂无对话历史[/dim]")
            for i, (q, a) in enumerate(history, 1):
                console.print(f"\n[dim]--- 第{i}轮 ---[/dim]")
                console.print(f"[green]🧑 {q}[/green]")
                console.print(f"[blue]🤖 {a[:200]}...[/blue]")
            continue

        # 调用 Agent 管线
        try:
            result = run_pipeline(user_input, verbose=True)
            history.append((user_input, result.get("analysis_result", "")))
        except Exception as e:
            console.print(f"[red]❌ 执行出错: {e}[/red]")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 更新 src/cli/__init__.py**

```python
from src.cli.main import main
```

- [ ] **Step 3: 创建启动脚本**

创建项目根目录的 `run.py`：

```python
"""办公自动化 Agent 启动入口"""
from src.cli.main import main

if __name__ == "__main__":
    main()
```

验证：
```bash
python run.py
```

期望看到 banner 和交互提示符 `🧑 你:`。输入 `/tools` 应列出 7 个工具。

---

### Task 8: 模拟数据 + 端到端集成

- [ ] **Step 1: 生成模拟销售数据**

创建 `scripts/generate_mock_data.py`：

```python
"""生成模拟销售数据"""
import random
from datetime import date, timedelta
from src.tools.excel_tools import write_excel

# 模拟配置
PRODUCTS = [
    {"name": "智能门锁 Pro", "price": 2999},
    {"name": "智能门锁 Lite", "price": 1599},
    {"name": "智能摄像头 360", "price": 499},
    {"name": "门窗传感器", "price": 199},
    {"name": "智能网关", "price": 899},
]

REGIONS = ["华东", "华南", "华北", "西南", "华中"]
SALES_REPS = ["张三", "李四", "王五", "赵六"]

# 生成本周数据（6/16-6/20，共5天）
today = date.today()
monday = today - timedelta(days=today.weekday())

data = []
for day_offset in range(5):
    current_date = monday + timedelta(days=day_offset)
    for _ in range(random.randint(20, 35)):  # 每天 20-35 条记录
        product = random.choice(PRODUCTS)
        quantity = random.randint(1, 20)
        data.append({
            "日期": current_date.isoformat(),
            "产品名称": product["name"],
            "单价": product["price"],
            "销量": quantity,
            "销售额": product["price"] * quantity,
            "销售员": random.choice(SALES_REPS),
            "区域": random.choice(REGIONS),
        })

result = write_excel(data, "data/sales_2026W25.xlsx", "销售明细")
print(f"已生成 {result['row_count']} 条模拟销售数据 → data/sales_2026W25.xlsx")
```

运行：
```bash
python scripts/generate_mock_data.py
```

- [ ] **Step 2: 端到端测试**

创建 `tests/test_e2e.py`：

```python
"""端到端集成测试"""
from src.agent.graph_pipeline import run_pipeline


def test_full_scenario():
    """完整场景: 整理销售数据 → 分析 → 生成周报 → 发送邮件"""
    result = run_pipeline(
        "整理本周销售数据，分析销售趋势，生成周报发送给经理 zhangjingli@company.com",
        verbose=False,
    )

    # 验证管线执行完成
    assert result is not None
    assert result.get("user_intent"), "意图解析不应为空"
    assert result.get("analysis_result"), "分析结果不应为空"
    assert result.get("report_path"), "报告路径不应为空"
    assert result.get("email_status"), "邮件状态不应为空"

    print("\n✅ 端到端测试通过！")
    print(f"  意图: {result['user_intent']}")
    print(f"  报告: {result['report_path']}")
    print(f"  邮件: {result['email_status']}")
```

- [ ] **Step 3: 运行完整验证**

```bash
# 1. 生成模拟数据
python scripts/generate_mock_data.py

# 2. 运行端到端测试
python -m pytest tests/test_e2e.py -v

# 3. 手动交互测试
python run.py
# 输入: 整理本周销售数据，分析趋势，生成周报发给经理
```

---

## 完成检查清单

全部完成后，确认以下事项：

- [ ] 7 个工具函数均被 `@tool` 包装，在 ALL_TOOLS 列表中
- [ ] 单轮 Agent 能正确调用工具并返回结果
- [ ] LangGraph 管线 5 个节点全部连通
- [ ] CLI 交互正常（`/help`、`/tools`、输入指令）
- [ ] 端到端场景跑通（模拟数据 → 分析 → 周报 → 邮件）
- [ ] 异常不会导致崩溃（如文件不存在时优雅报错）
