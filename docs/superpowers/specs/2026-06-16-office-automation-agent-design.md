# 办公自动化 Agent — 设计方案

> 日期: 2026-06-16 | 状态: 已确认

## 项目概述

基于 DeepSeek 大模型的办公自动化 Agent。用户通过自然语言下达指令（如"整理本周销售数据并生成周报发送给经理"），Agent 自动完成数据读取、分析、文档生成与邮件发送的全流程闭环。

## 技术选型

| 组件 | 选型 | 原因 |
|------|------|------|
| 大模型 | DeepSeek Chat | 兼容 OpenAI SDK，Function Calling 稳定，¥1/百万 token |
| LLM 框架 | LangChain Core | 工具定义、消息管理、模型抽象 |
| 任务编排 | LangGraph StateGraph | 多步任务拆解、条件分支、异常处理 |
| Excel | openpyxl | 轻量纯 Python，无 C 依赖 |
| 邮件 | smtplib + email | Python 标准库，无需额外安装 |
| 文件操作 | os + shutil | 标准库 |
| CLI 界面 | rich | 彩色输出、表格、Markdown 渲染 |
| Web 界面（未来） | Gradio | 一行代码出聊天 UI，支持文件上传 |
| 配置管理 | python-dotenv | API Key 与代码分离 |

## 架构分层

```
┌─────────────────────────────────────────┐
│  Step 5: CLI 交互层 (rich)               │
│  用户输入 → stream → 结果展示             │
├─────────────────────────────────────────┤
│  Step 4: LangGraph 任务编排               │
│  StateGraph: 意图解析→数据获取→分析→输出→推送 │
├─────────────────────────────────────────┤
│  Step 3: 单轮 Agent                      │
│  DeepSeek + bind_tools → ReAct 循环      │
├─────────────────────────────────────────┤
│  Step 2: LangChain Tool 封装              │
│  @tool 装饰器 → JSON Schema              │
├─────────────────────────────────────────┤
│  Step 1: 底层工具函数                      │
│  openpyxl / smtplib / os.shutil          │
└─────────────────────────────────────────┘
```

## 实施路线（7 步）

### Step 0: 环境搭建
- Python 3.11 + venv
- 安装依赖：langchain-core, langchain-openai, langgraph, openpyxl, python-dotenv, rich
- DeepSeek API Key 配置（platform.deepseek.com）
- 验证：运行一次 ChatOpenAI 调用，确认连通

### Step 1: 底层工具函数
纯 Python 函数，不涉及 LLM。四个模块共 9 个函数：

| 模块 | 函数 | 功能 |
|------|------|------|
| excel_tools.py | read_excel() | 读取 Excel 返回行数据 |
| | write_excel() | 写入数据到 Excel |
| | summarize_excel() | 概览：行数、列名、基本统计 |
| file_tools.py | list_files() | 列出目录文件 |
| | classify_files() | 按扩展名分类移动到子目录 |
| email_tools.py | send_email() | SMTP 发送邮件（dry-run 模式） |
| date_tools.py | get_week_range() | 计算本周/上周起止日期 |
| | format_date() | 日期格式化 |

设计原则：每个函数只做一件事，返回结构化为 dict，参数简单明确。

### Step 2: LangChain Tool 封装
- 用 `@tool` 装饰器包装 Step 1 函数
- `@tool` 自动从类型标注 + docstring 生成 JSON Schema
- LLM 通过 JSON Schema 理解每个工具的用途和参数
- 汇总到 `ALL_TOOLS` 列表

### Step 3: 单轮 Agent
- `ChatOpenAI(model="deepseek-chat", base_url="https://api.deepseek.com/v1")`
- `.bind_tools(ALL_TOOLS)` 绑定工具
- ReAct 循环：用户消息 → LLM 判断 → 执行工具 → 返回结果
- 消息类型：HumanMessage / AIMessage / ToolMessage

### Step 4: LangGraph 多步编排
StateGraph 5 节点管线：
1. **意图解析** — LLM 理解用户要什么，拆成子任务
2. **数据获取** — 调用 Excel/文件工具拿数据
3. **分析处理** — LLM 分析数据，生成洞察
4. **结果输出** — 生成周报文档
5. **消息推送** — 发送邮件

关键特性：AgentState TypedDict 状态定义、conditional_edges 条件分支、异常回退处理。

### Step 5: CLI 交互层
- rich 库美化输出（Console、Panel、status、table）
- 主循环：读取输入 → 调用 LangGraph → 展示结果
- 流式输出：`app.stream()` 实时显示每步进度
- 内置命令：`/help`、`/quit`、`/history`

### Step 6: 模拟场景集成
- 模拟数据：sales_2026W25.xlsx（100-150 条销售记录）
- 端到端验证："整理本周销售数据 → 分析趋势 → 生成周报 → 发送经理"
- 验证清单：工具调用正确、状态流转完整、周报格式正确、异常处理到位

### Step 7: Web 界面（未来扩展）
- Gradio ChatInterface，纯 Python，无需前端代码
- 支持文件上传（用户直接拖 Excel）

## 项目目录结构

```
办公自动化系统/
├── .env                    # API Key
├── requirements.txt
├── src/
│   ├── __init__.py
│   ├── tools/              # Step 1-2
│   │   ├── __init__.py
│   │   ├── excel_tools.py
│   │   ├── file_tools.py
│   │   ├── email_tools.py
│   │   └── date_tools.py
│   ├── agent/              # Step 3-4
│   │   ├── __init__.py
│   │   ├── single_agent.py
│   │   └── graph_pipeline.py
│   └── cli/                # Step 5
│       ├── __init__.py
│       └── main.py
├── data/                   # Step 6 模拟数据
│   ├── sales_2026W25.xlsx
│   └── config.yaml
├── output/                 # 生成物
└── docs/
    └── superpowers/
        └── specs/
            └── 2026-06-16-office-automation-agent-design.md
```
