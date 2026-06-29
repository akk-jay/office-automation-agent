"""单轮 Agent — 一次对话中自动选择并执行工具

这是整个项目最核心的一步：让 LLM "看见"工具列表，根据用户意图自动选择工具并执行。

数据流：
  用户输入 → HumanMessage
       ↓
  LLM + 工具列表 → AIMessage (含 tool_calls 或直接回复)
       ↓
  如果有 tool_calls → 执行工具 → ToolMessage 追加到对话
       ↓
  LLM 再次推理 → 最终回复
"""

import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage

load_dotenv()


class SingleAgent:
    """最简 Agent：LLM + 工具列表，自动完成"理解意图 → 选工具 → 执行 → 回复"。

    这个类的核心就两行：
        self.llm = ChatOpenAI(...)            # 连接模型
        self.llm_with_tools = llm.bind_tools(tools)  # 把工具列表"注入"模型
    """

    def __init__(self, tools: list, verbose: bool = True):
        """
        Args:
            tools: @tool 装饰过的工具列表 (来自 src.tools.ALL_TOOLS)
            verbose: 是否打印中间过程
        """
        self.tools = tools
        self.tool_map = {t.name: t for t in tools}
        self.verbose = verbose

        # 连接 DeepSeek
        # 为什么用 ChatOpenAI 而不是专门的 DeepSeek 类？
        # 因为 DeepSeek API 完全兼容 OpenAI SDK 格式，可以直接复用
        self.llm = ChatOpenAI(
            model="deepseek-chat",
            base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1"),
            api_key=os.getenv("DEEPSEEK_API_KEY"),
            temperature=0,  # temperature=0 让模型输出更确定性，工具调用更稳定
        )

        # bind_tools 做了什么？
        # 1. 把每个 @tool 的 name/description/参数schema 转成 OpenAI Function Calling 格式
        # 2. 每次调用 LLM 时，工具列表随请求一起发送
        # 3. LLM 可以根据工具描述决定"该不该调工具、调哪个、传什么参数"
        self.llm_with_tools = self.llm.bind_tools(self.tools)

    def run(self, user_input: str) -> str:
        """处理用户输入，自动调用工具并返回最终结果。

        这就是简历里说的：
        "LLM 根据用户意图自动选择并调用对应工具，实现自然语言到具体操作的自动映射"

        Args:
            user_input: 用户的自然语言指令

        Returns:
            Agent 的最终回复文本
        """
        messages = [HumanMessage(content=user_input)]

        if self.verbose:
            print(f"\n{'─'*40}")
            print(f" 用户: {user_input}")

        # 第一轮：LLM 分析意图，决定要不要调工具
        response = self.llm_with_tools.invoke(messages)

        # 检查 response.tool_calls
        # - 如果为空 → LLM 觉得不需要工具，直接文字回复
        # - 如果有内容 → LLM 要求调用工具，我们需要执行并返回结果
        if not response.tool_calls:
            if self.verbose:
                print(f" Agent (直接回复): {response.content}")
            return response.content

        # LLM 决定调用工具
        if self.verbose:
            print(f" 调用 {len(response.tool_calls)} 个工具:")

        messages.append(response)

        for tool_call in response.tool_calls:
            tool_name = tool_call["name"]
            tool_args = tool_call["args"]
            tool_id = tool_call["id"]

            if self.verbose:
                print(f"    → {tool_name}")
                # 简化参数显示
                args_short = {k: str(v)[:60] for k, v in tool_args.items()}
                print(f"      参数: {args_short}")

            # 执行工具 — 这是我们 Step 1-3 写的纯 Python 函数
            tool_func = self.tool_map[tool_name]
            tool_result = tool_func.invoke(tool_args)

            if self.verbose:
                print(f"    ← 结果: {tool_result[:150]}...")

            # 把工具执行结果以 ToolMessage 形式追加到对话
            # 这很关键：ToolMessage 关联到对应的 tool_call_id，
            # LLM 才能知道"这个结果是哪个工具调用返回的"
            messages.append(ToolMessage(content=tool_result, tool_call_id=tool_id))

        # 第二轮：LLM 根据工具返回的结果，生成最终回复
        final_response = self.llm_with_tools.invoke(messages)

        if self.verbose:
            print(f" Agent (最终回复): {final_response.content}")

        return final_response.content
