"""LangGraph 多步任务编排 — 将复杂指令拆解为有序步骤

这是简历里写的：
"基于 LangGraph 构建任务拆解图（StateGraph），将复杂指令拆解为
'意图解析→数据获取→分析处理→结果输出→消息推送'等有序步骤节点，
支持条件分支判断与异常回退处理"

核心概念：
  - StateGraph = 状态机 + 有向图
  - State = 在节点间流转的数据字典
  - Node = 一个处理函数，读 State、返回 State 的更新
  - Edge = 节点间的连线，条件边根据 State 决定走哪条
"""

import os
import json
import re
from typing import TypedDict, Annotated, Literal
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage
from langgraph.graph import StateGraph, END
from langgraph.graph.message import add_messages

from src.tools import ALL_TOOLS, TOOL_MAP

load_dotenv()


# ============================================================
# 1. State 定义 — 在节点间流转的数据
# ============================================================

class AgentState(TypedDict):
    """Agent 的全局状态。

    每个字段都可以被任何节点读取和写入。
    add_messages 注解表示 messages 是追加而非覆盖。
    """
    messages: Annotated[list, add_messages]  # 对话历史
    user_intent: str       # 意图描述
    sub_tasks: list        # 拆解后的子任务列表
    data_summary: str      # 数据获取结果
    analysis_result: str   # LLM 分析结论
    report_path: str       # 生成的周报路径
    email_status: str      # 邮件发送状态
    error: str             # 错误信息


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
# 3. 节点函数 — 每个节点只做一件事
# ============================================================

def node_intent_parser(state: AgentState) -> dict:
    """节点1: 意图解析

    用 LLM 分析用户指令，提取：
    - intent: 用户想做什么
    - sub_tasks: 拆成哪些子任务
    - 需要哪些操作（读Excel / 分析 / 写报告 / 发邮件）
    """
    last_msg = state["messages"][-1]
    user_input = last_msg.content if hasattr(last_msg, 'content') else str(last_msg)

    prompt = f"""分析以下用户指令，提取关键信息并以 JSON 格式返回。

用户指令: "{user_input}"

返回格式（只返回 JSON，不要其他文字）:
{{
    "intent": "简要描述用户的意图",
    "sub_tasks": ["子任务1", "子任务2"],
    "needs_data_fetch": true,
    "needs_analysis": true,
    "needs_report": true,
    "needs_email": true,
    "email_recipient": "收件人邮箱或null"
}}
"""
    response = llm.invoke([HumanMessage(content=prompt)])

    # 解析 JSON（容错：有时 LLM 会在 JSON 外面包 markdown 代码块）
    try:
        intent_data = json.loads(response.content)
    except json.JSONDecodeError:
        match = re.search(r'\{.*\}', response.content, re.DOTALL)
        if match:
            intent_data = json.loads(match.group())
        else:
            return {"error": f"意图解析失败: {response.content[:200]}"}

    return {
        "user_intent": intent_data.get("intent", ""),
        "sub_tasks": intent_data.get("sub_tasks", []),
        "error": None,
    }


def node_data_fetcher(state: AgentState) -> dict:
    """节点2: 数据获取

    调用工具获取：
    1. 本周日期范围 (tool_get_week_range)
    2. 销售数据文件概览 (tool_summarize_excel)
    """
    results = []

    # 获取日期范围
    try:
        date_result = TOOL_MAP["tool_get_week_range"].invoke({"week_offset": 0})
        results.append(f"[日期范围] {date_result}")
    except Exception as e:
        results.append(f"[日期范围] 获取失败: {e}")

    # 获取数据文件概览
    try:
        summary = TOOL_MAP["tool_summarize_excel"].invoke(
            {"file_path": "data/sales_2026W25.xlsx"}
        )
        results.append(f"[数据概览] {summary}")
    except Exception as e:
        results.append(f"[数据概览] 获取失败: {e}")

    return {"data_summary": "\n".join(results), "error": None}


def node_analyzer(state: AgentState) -> dict:
    """节点3: 分析处理

    LLM 根据数据获取结果，生成分析洞察。
    这是整个流程中最依赖 LLM 的环节——让模型理解数据并得出结论。
    """
    prompt = f"""你是一位数据分析师。根据以下数据概要，写一段简短的销售分析总结（200字以内）。

数据概要:
{state.get("data_summary", "无数据")}

请分析:
- 整体销售表现
- 值得关注的亮�或问题
- 建议
"""
    response = llm.invoke([HumanMessage(content=prompt)])
    return {"analysis_result": response.content, "error": None}


def node_report_generator(state: AgentState) -> dict:
    """节点4: 结果输出

    将分析结果生成周报 Excel 文件。
    """
    analysis = state.get("analysis_result", "无分析结果")
    data_summary = state.get("data_summary", "")

    # 构造周报内容
    report_data = [
        {"项目": "销售数据周报", "内容": ""},
        {"项目": "分析结论", "内容": analysis},
        {"项目": "数据概要", "内容": data_summary[:500]},
    ]

    try:
        result = TOOL_MAP["tool_write_excel"].invoke({
            "data_json": json.dumps(report_data, ensure_ascii=False),
            "file_path": "output/周报_最新.xlsx",
            "sheet_name": "周报",
        })
        return {"report_path": "output/周报_最新.xlsx", "error": None}
    except Exception as e:
        return {"report_path": None, "error": f"周报生成失败: {e}"}


def node_sender(state: AgentState) -> dict:
    """节点5: 消息推送

    将周报通过邮件发送。
    """
    analysis = state.get("analysis_result", "")
    report_path = state.get("report_path", "")

    try:
        result = TOOL_MAP["tool_send_email"].invoke({
            "to": "manager@company.com",
            "subject": "销售数据周报",
            "body": f"您好，以下是本周销售数据周报。\n\n分析结论:\n{analysis}",
            "attachment": report_path,
            "dry_run": True,
        })
        return {"email_status": result, "error": None}
    except Exception as e:
        return {"email_status": None, "error": f"邮件发送失败: {e}"}


# ============================================================
# 4. 条件路由 — 根据 State 决定下一步
# ============================================================

def route_after_intent(state: AgentState) -> Literal["data_fetcher", "__end__"]:
    """意图解析后：有错误 → 结束，否则 → 数据获取"""
    if state.get("error"):
        return "__end__"
    return "data_fetcher"


def route_after_fetch(state: AgentState) -> Literal["analyzer", "__end__"]:
    """数据获取后：有错误 → 结束，否则 → 分析"""
    if state.get("error"):
        return "__end__"
    return "analyzer"


def route_after_analysis(state: AgentState) -> Literal["report_generator", "__end__"]:
    """分析后：有错误 → 结束，否则 → 生成报告"""
    if state.get("error"):
        return "__end__"
    return "report_generator"


def route_after_report(state: AgentState) -> Literal["sender", "__end__"]:
    """报告生成后：有错误 → 结束，否则 → 发送"""
    if state.get("error"):
        return "__end__"
    return "sender"


# ============================================================
# 5. 构建图
# ============================================================

def build_pipeline():
    """构建 LangGraph 任务管线。

    管线结构:
        START
          │
        [intent_parser] ──(error)── END
          │ (ok)
        [data_fetcher] ──(error)── END
          │ (ok)
        [analyzer] ──(error)── END
          │ (ok)
        [report_generator] ──(error)── END
          │ (ok)
        [sender]
          │
         END

    每个节点出错时优雅降级，不会让整个管线崩溃。
    """
    workflow = StateGraph(AgentState)

    # 注册节点
    workflow.add_node("intent_parser", node_intent_parser)
    workflow.add_node("data_fetcher", node_data_fetcher)
    workflow.add_node("analyzer", node_analyzer)
    workflow.add_node("report_generator", node_report_generator)
    workflow.add_node("sender", node_sender)

    # 入口
    workflow.set_entry_point("intent_parser")

    # 条件边 — 每个节点之后都检查是否有错误
    workflow.add_conditional_edges("intent_parser", route_after_intent, {
        "data_fetcher": "data_fetcher",
        "__end__": END,
    })
    workflow.add_conditional_edges("data_fetcher", route_after_fetch, {
        "analyzer": "analyzer",
        "__end__": END,
    })
    workflow.add_conditional_edges("analyzer", route_after_analysis, {
        "report_generator": "report_generator",
        "__end__": END,
    })
    workflow.add_conditional_edges("report_generator", route_after_report, {
        "sender": "sender",
        "__end__": END,
    })

    # 最后一步 → 结束
    workflow.add_edge("sender", END)

    return workflow.compile()


# ============================================================
# 6. 便捷函数
# ============================================================

def run_pipeline(user_input: str, verbose: bool = True) -> dict:
    """运行完整的 LangGraph 管线。

    Args:
        user_input: 用户自然语言指令
        verbose: 是否打印每个节点的输出

    Returns:
        最终的 AgentState
    """
    pipeline = build_pipeline()

    initial_state = {"messages": [HumanMessage(content=user_input)]}

    if verbose:
        print(f"\n{'='*50}")
        print(f"Agent 管线启动")
        print(f"用户指令: {user_input}")
        print(f"{'='*50}")

    # stream 模式：每执行一个节点就输出一次
    for step in pipeline.stream(initial_state):
        node_name = list(step.keys())[0]
        node_output = step[node_name]

        if verbose:
            if node_name == "intent_parser" and node_output.get("user_intent"):
                print(f"\n[意图解析] {node_output['user_intent']}")
                print(f"  子任务: {node_output.get('sub_tasks', [])}")
            elif node_name == "data_fetcher" and node_output.get("data_summary"):
                print(f"\n[数据获取]")
                print(f"  {node_output['data_summary'][:300]}")
            elif node_name == "analyzer" and node_output.get("analysis_result"):
                print(f"\n[分析处理]")
                print(f"  {node_output['analysis_result'][:300]}")
            elif node_name == "report_generator" and node_output.get("report_path"):
                print(f"\n[周报生成] {node_output['report_path']}")
            elif node_name == "sender" and node_output.get("email_status"):
                print(f"\n[邮件发送] {node_output['email_status'][:200]}")
            if node_output.get("error"):
                print(f"\n[错误] {node_output['error']}")

    # 获取最终完整状态
    final_state = pipeline.invoke(initial_state)

    if verbose:
        print(f"\n{'='*50}")
        print("管线执行完成")
        print(f"{'='*50}\n")

    return final_state
