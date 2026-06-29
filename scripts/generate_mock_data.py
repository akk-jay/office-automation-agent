"""生成模拟销售数据"""
import sys
sys.path.insert(0, ".")
import random
from datetime import date, timedelta
from src.tools.excel_tools import write_excel

PRODUCTS = [
    {"name": "智能门锁 Pro", "price": 2999},
    {"name": "智能门锁 Lite", "price": 1599},
    {"name": "智能摄像头 360", "price": 499},
    {"name": "门窗传感器", "price": 199},
    {"name": "智能网关", "price": 899},
]

REGIONS = ["华东", "华南", "华北", "西南", "华中"]
SALES_REPS = ["张三", "李四", "王五", "赵六"]

today = date.today()
monday = today - timedelta(days=today.weekday())

data = []
for day_offset in range(5):
    current_date = monday + timedelta(days=day_offset)
    for _ in range(random.randint(20, 35)):
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
