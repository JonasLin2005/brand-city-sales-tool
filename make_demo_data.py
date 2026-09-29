#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Write sample_data/contracts_demo.xlsx: a synthetic contract book (brand,
branch, funding date) in the same shape as the real one, for trying the
pipeline end to end. Every figure is random; nothing comes from real records.
"""

import os
import random
from datetime import datetime, timedelta

import openpyxl

COMPANY = "示例汽车服务有限公司"
BRANDS = {"比亚迪": 1.0, "理想": 0.5, "AITO 问界": 0.35}
BRANCHES = ["广州第一", "佛山", "杭州第一", "合肥", "四川", "长沙"]
START, DAYS = datetime(2025, 1, 1), 546          # Jan 2025 - Jun 2026


def main():
    rng = random.Random(410)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "模板"
    ws.append(["品牌", "分公司", "放款日期"])
    for brand, weight in BRANDS.items():
        for branch in BRANCHES:
            n = int(rng.randint(150, 400) * weight)
            for _ in range(n):
                when = START + timedelta(days=rng.randrange(DAYS),
                                         hours=rng.randint(9, 17))
                ws.append([brand, f"{COMPANY}{branch}分公司", when])
    # a few rows the builder must skip
    ws.append(["比亚迪", f"{COMPANY}佛山分公司", None])
    ws.append(["理想", f"{COMPANY}合肥分公司", datetime(2024, 12, 20)])
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "sample_data", "contracts_demo.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print(f"wrote {ws.max_row - 1} rows to {out}")


if __name__ == "__main__":
    main()
