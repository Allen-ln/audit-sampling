# -*- coding: utf-8 -*-
"""生成模拟应收账款明细（全部虚构，不含真实客户信息）。

金额按审计总体的典型形态分布：大量小额、少量大额（长尾），
并掺入 3 笔红字冲销（负数），用于演示零负金额项目的单独处理。
运行：python sample_data/generate_sample.py
"""
from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

random.seed(20251201)
OUT = Path(__file__).resolve().parent

CUSTOMERS = [
    "杭州晨曦贸易有限公司", "宁波远舟物流有限公司", "苏州恒泰机械制造有限公司",
    "上海知微信息技术有限公司", "嘉兴南湖包装材料有限公司", "绍兴越诚纺织有限公司",
    "合肥中衡检测技术有限公司", "南京金鼎工程咨询有限公司", "温州港盛进出口有限公司",
    "湖州长兴电子科技有限公司", "台州宏信模具有限公司", "金华万邦供应链有限公司",
    "义乌小商品城贸易有限公司", "杭州云栖数据服务有限公司", "常州华通电气有限公司",
    "无锡嘉信精密仪器有限公司", "南通海纳新材料有限公司", "扬州广陵建设工程有限公司",
    "镇江润州商贸有限公司", "徐州彭城重工有限公司",
]

SUMMARIES = ["销售商品-货款", "提供服务-技术服务费", "销售商品-配件", "项目结算款",
             "销售商品-批量订单", "技术服务-年度维护", "销售商品-样品", "工程结算"]


def _amount() -> float:
    r = random.random()
    if r < 0.60:
        return round(random.uniform(1_000, 60_000), 2)
    if r < 0.90:
        return round(random.uniform(60_000, 250_000), 2)
    if r < 0.99:
        return round(random.uniform(250_000, 900_000), 2)
    return round(random.uniform(900_000, 2_500_000), 2)


def build() -> Path:
    rows = []
    for i in range(1, 801):
        d = datetime(2025, 12, 1) + timedelta(days=random.randint(0, 30))
        rows.append({
            "凭证号": f"记-{i:04d}",
            "记账日期": d.strftime("%Y-%m-%d"),
            "摘要": random.choice(SUMMARIES),
            "对方单位": random.choice(CUSTOMERS),
            "科目": "1122 应收账款",
            "借方金额": f"{_amount():,.2f}",
        })

    # 3 笔红字冲销（负数），用于演示零负金额项目的单独处理
    for k, amt in enumerate([-18_600.00, -245_000.00, -7_320.50], start=1):
        rows.append({
            "凭证号": f"记-冲{k:02d}",
            "记账日期": "2025-12-28",
            "摘要": "红字冲销-退货",
            "对方单位": random.choice(CUSTOMERS),
            "科目": "1122 应收账款",
            "借方金额": f"{amt:,.2f}",
        })

    df = pd.DataFrame(rows)
    path = OUT / "应收账款明细.xlsx"
    df.to_excel(path, index=False)
    total = sum(float(str(x).replace(",", "")) for x in df["借方金额"])
    print(f"已生成：{path.name}（{len(df)} 笔，合计 {total:,.2f} 元）")
    return path


if __name__ == "__main__":
    build()
