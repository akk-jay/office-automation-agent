"""生成 Demo GIF — 模拟终端录屏效果，展示管线完整执行流程"""
from PIL import Image, ImageDraw, ImageFont
import os
import sys
import textwrap

# --- Config ---
W, H = 800, 600
BG = (15, 17, 26)        # dark terminal bg
FG = (200, 210, 220)      # main text
DIM = (100, 110, 120)     # dim text
ACCENT = (56, 189, 248)   # cyan accent
GREEN = (74, 222, 128)    # success green
YELLOW = (250, 204, 21)   # yellow
PURPLE = (168, 139, 250)  # purple
ROSE = (251, 113, 133)    # red/rose
ORANGE = (251, 146, 60)   # orange

FONT_PATH_WIN = "C:\\Windows\\Fonts\\msyh.ttc"
FONT_PATH_FALLBACK = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

if os.path.exists(FONT_PATH_WIN):
    FONT = ImageFont.truetype(FONT_PATH_WIN, 14)
    FONT_SM = ImageFont.truetype(FONT_PATH_WIN, 12)
    FONT_BOLD = ImageFont.truetype(FONT_PATH_WIN, 16)
    FONT_TITLE = ImageFont.truetype(FONT_PATH_WIN, 18)
else:
    try:
        FONT = ImageFont.truetype(FONT_PATH_FALLBACK, 13)
        FONT_SM = ImageFont.truetype(FONT_PATH_FALLBACK, 11)
        FONT_BOLD = ImageFont.truetype(FONT_PATH_FALLBACK, 14)
        FONT_TITLE = ImageFont.truetype(FONT_PATH_FALLBACK, 16)
    except:
        FONT = ImageFont.load_default()
        FONT_SM = FONT
        FONT_BOLD = FONT
        FONT_TITLE = FONT


def create_frame(lines, current_step=None, header=True):
    """Create a single terminal frame"""
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    y = 16

    # Title bar
    draw.rectangle([0, 0, W, 32], fill=(30, 35, 50))
    draw.text((12, 8), "●  ●  ●", fill=(80, 80, 90), font=FONT_SM)
    draw.text((W // 2 - 60, 8), "办公自动化 Agent", fill=(150, 155, 165), font=FONT_SM)

    y = 44

    if header:
        # Banner
        draw.text((16, y), "办公自动化 Agent v1.0", fill=ACCENT, font=FONT_BOLD)
        y += 16
        draw.text((16, y), "基于 DeepSeek + LangGraph | 输入 /help 查看帮助", fill=DIM, font=FONT_SM)
        y += 24
        draw.text((16, y), "你: 整理本周销售数据并生成周报发送给经理", fill=GREEN, font=FONT_BOLD)
        y += 24

        # Separator
        draw.line([(16, y), (W - 16, y)], fill=(40, 45, 55))
        y += 12
        draw.text((16, y), "Agent 管线启动", fill=DIM, font=FONT_SM)
        y += 4
        draw.line([(16, y), (W - 16, y)], fill=(40, 45, 55))
        y += 16

    # Render each line
    for line_info in lines:
        if isinstance(line_info, dict):
            color = line_info.get("color", FG)
            text = line_info.get("text", "")
            indent = line_info.get("indent", 0)
            font = line_info.get("font", FONT)
        else:
            color = FG
            text = line_info
            indent = 0
            font = FONT

        if text == "---":
            draw.line([(16 + indent, y), (W - 16, y)], fill=(40, 45, 55))
            y += 8
            continue

        draw.text((16 + indent, y), text, fill=color, font=font)
        y += 20

    # Highlight current step
    if current_step:
        draw.rectangle([8, H - 28, W - 8, H - 8], fill=(25, 30, 40))
        draw.text((16, H - 22), current_step, fill=ACCENT, font=FONT_SM)

    return img


# --- Build Frame Sequence ---

frames = []
delay = 2000  # ms per frame

# Frame 1: User input
f1 = create_frame([
    {"text": "用户输入:", "color": DIM, "font": FONT_SM},
], current_step="等待输入...")
frames.append(f1)

# Frame 2: Input received + pipeline starts
f2 = create_frame([
    {"text": "[步骤 1/5] 意图解析中...", "color": YELLOW, "font": FONT_SM},
    {"text": "---"},
    {"text": "[意图解析] 整理本周销售数据并生成周报发送给经理", "color": ACCENT, "font": FONT},
    {"text": "  子任务:", "color": DIM, "font": FONT},
    {"text": "    1. 读取销售数据文件", "color": FG, "font": FONT},
    {"text": "    2. 数据分析与汇总", "color": FG, "font": FONT},
    {"text": "    3. 生成周报 Excel", "color": FG, "font": FONT},
    {"text": "    4. 邮件发送给经理", "color": FG, "font": FONT},
], current_step="▶ 节点 1/5: intent_parser → 意图解析完成")
frames.append(f2)

# Frame 3: Data fetching
f3 = create_frame([
    {"text": "[意图解析] 整理本周销售数据并生成周报发送给经理", "color": DIM, "font": FONT},
    {"text": "  子任务: 读取销售数据 → 分析数据 → 生成报告 → 发送邮件", "color": DIM, "font": FONT},
    {"text": "---"},
    {"text": "[步骤 2/5] 数据获取中...", "color": YELLOW, "font": FONT_SM},
    {"text": "---"},
    {"text": "[数据获取]", "color": ACCENT, "font": FONT},
    {"text": "  📅 本周: 第27周 (2026-06-29 ~ 2026-07-05)", "color": FG, "font": FONT},
    {"text": "  📊 文件: data/sales_2026W25.xlsx", "color": FG, "font": FONT},
    {"text": "  📈 订单数: 129", "color": FG, "font": FONT},
    {"text": "  💰 总销售额: ¥1,549,237", "color": GREEN, "font": FONT_BOLD},
    {"text": "  📊 平均客单价: ¥12,009", "color": FG, "font": FONT},
], current_step="▶ 节点 2/5: data_fetcher → 数据获取完成")
frames.append(f3)

# Frame 4: Analysis
f4 = create_frame([
    {"text": "[数据获取] 第27周 (6/29-7/5) · 129 单 · ¥1,549,237", "color": DIM, "font": FONT},
    {"text": "---"},
    {"text": "[步骤 3/5] LLM 分析处理中...", "color": YELLOW, "font": FONT_SM},
    {"text": "---"},
    {"text": "[分析处理]", "color": ACCENT, "font": FONT},
    {"text": "  本周(第27周)销售分析：", "color": FG, "font": FONT},
    {"text": "  • 总销售额约154.9万元，共1363件商品", "color": FG, "font": FONT},
    {"text": "  • 平均客单价约1,136元，存在高价产品(最高2,999元)", "color": FG, "font": FONT},
    {"text": "  • 最高单品销量20件，高端&低价产品均有市场", "color": FG, "font": FONT},
    {"text": "  • 建议：对低价产品做促销活动 → 提高销量", "color": GREEN, "font": FONT},
], current_step="▶ 节点 3/5: analyzer → LLM 分析完成")
frames.append(f4)

# Frame 5: Report generation
f5 = create_frame([
    {"text": "[分析处理] 销售额154.9万 · 平均单价1,136元 · 建议促销", "color": DIM, "font": FONT},
    {"text": "---"},
    {"text": "[步骤 4/5] 周报生成中...", "color": YELLOW, "font": FONT_SM},
    {"text": "---"},
    {"text": "[周报生成]", "color": ACCENT, "font": FONT},
    {"text": "  📄 文件: output/周报_最新.xlsx", "color": GREEN, "font": FONT_BOLD},
    {"text": "  📋 工作表: 周报", "color": FG, "font": FONT},
    {"text": "  📝 内容: 销售数据周报 + 分析结论 + 数据概要", "color": FG, "font": FONT},
    {"text": "  ✅ 报告生成成功!", "color": GREEN, "font": FONT_BOLD},
], current_step="▶ 节点 4/5: report_generator → 周报生成完成")
frames.append(f5)

# Frame 6: Email send
f6 = create_frame([
    {"text": "[周报生成] output/周报_最新.xlsx ✅", "color": DIM, "font": FONT},
    {"text": "---"},
    {"text": "[步骤 5/5] 邮件发送中...", "color": YELLOW, "font": FONT_SM},
    {"text": "---"},
    {"text": "[邮件发送]", "color": ACCENT, "font": FONT},
    {"text": "  [DRY RUN 模拟模式]", "color": ORANGE, "font": FONT_BOLD},
    {"text": "  📧 发件人: (未配置)", "color": DIM, "font": FONT},
    {"text": "  📧 收件人: manager@company.com", "color": FG, "font": FONT},
    {"text": "  📧 主题: 销售数据周报", "color": FG, "font": FONT},
    {"text": "  📎 附件: output/周报_最新.xlsx", "color": FG, "font": FONT},
    {"text": "  ✅ 邮件准备就绪 (dry_run)", "color": GREEN, "font": FONT_BOLD},
], current_step="▶ 节点 5/5: sender → 邮件发送完成 (dry_run)")
frames.append(f6)

# Frame 7: Completion
f7 = create_frame([
    {"text": "  ✅ 意图解析 → ✅ 数据获取 → ✅ 分析处理 →", "color": GREEN, "font": FONT},
    {"text": "  ✅ 报告生成 → ✅ 邮件发送", "color": GREEN, "font": FONT},
    {"text": "---"},
    {"text": "", "color": FG, "font": FONT},
    {"text": "  ▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔", "color": DIM, "font": FONT},
    {"text": "  管线执行完成! 🎉", "color": GREEN, "font": FONT_BOLD},
    {"text": "  ▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁▁", "color": DIM, "font": FONT},
    {"text": "", "color": FG, "font": FONT},
    {"text": "  一句话 → 自动完成 Excel分析+报告+邮件", "color": ACCENT, "font": FONT_BOLD},
], current_step="✅ 全部完成 — 5个节点 / 0个错误", header=False)
frames.append(f7)

# --- Combine into GIF ---
output_path = "docs/demo.gif"
os.makedirs("docs", exist_ok=True)

# Duplicate last frame for a pause effect
frames.append(frames[-1])
frames.append(frames[-1])

frames[0].save(
    output_path,
    save_all=True,
    append_images=frames[1:],
    duration=[1200, 2500, 2500, 3000, 2000, 2200, 2500, 1500, 2000],
    loop=0,
    optimize=False,
    quality=85,
)

print(f"Demo GIF saved to {output_path}")
print(f"  Frames: {len(frames)}")
print(f"  Size: {os.path.getsize(output_path) / 1024:.1f} KB")
