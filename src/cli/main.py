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
    for t in ALL_TOOLS:
        table.add_row(t.name, t.description[:80])
    console.print(table)


def main():
    show_banner()
    while True:
        try:
            user_input = console.input("\n[bold green]你:[/bold green] ")
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]再见！[/dim]")
            break

        user_input = user_input.strip()
        if not user_input:
            continue
        if user_input in ("/quit", "/exit"):
            console.print("[dim]再见！[/dim]")
            break
        if user_input == "/help":
            console.print(Markdown(HELP_TEXT))
            continue
        if user_input == "/tools":
            show_tools()
            continue

        try:
            run_pipeline(user_input, verbose=True)
        except Exception as e:
            console.print(f"[red]执行出错: {e}[/red]")


if __name__ == "__main__":
    main()
