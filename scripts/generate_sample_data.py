"""
サンプルデータ生成スクリプト
3店舗 × 指定年の日別売上CSVを生成する
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

STORES = [
    {"name": "渋谷店", "slug": "shibuya", "base_sales": 150000, "base_count": 80},
    {"name": "新宿店", "slug": "shinjuku", "base_sales": 200000, "base_count": 110},
    {"name": "池袋店", "slug": "ikebukuro", "base_sales": 120000, "base_count": 65},
]

MONTH_FACTOR = {1: 1.1, 2: 0.95, 3: 1.0, 4: 1.05, 5: 1.1,
                6: 0.95, 7: 1.0, 8: 0.9, 9: 1.0, 10: 1.05,
                11: 1.1, 12: 1.3}

# 年ごとの成長係数（2025年は前年比+5%）
YEAR_GROWTH = {2024: 1.0, 2025: 1.05}

YEARS = [2024, 2025]

output_dir = Path(__file__).parent.parent / "data" / "sample"
output_dir.mkdir(exist_ok=True)

for year in YEARS:
    random.seed(42 + year)
    growth = YEAR_GROWTH.get(year, 1.0)
    start = date(year, 1, 1)
    end = date(year, 12, 31)

    for store in STORES:
        rows = []
        current = start
        while current <= end:
            weekday = current.weekday()
            day_factor = 1.3 if weekday >= 5 else (0.85 if weekday == 0 else 1.0)
            month_factor = MONTH_FACTOR[current.month]
            noise = random.uniform(0.85, 1.15)

            sales = int(store["base_sales"] * day_factor * month_factor * noise * growth)
            count = int(store["base_count"] * day_factor * month_factor * noise * growth)
            rows.append({
                "日付": current.isoformat(),
                "店舗名": store["name"],
                "売上金額": sales,
                "客数": count,
            })
            current += timedelta(days=1)

        filename = f"store_{store['slug']}_{year}.csv"
        filepath = output_dir / filename
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["日付", "店舗名", "売上金額", "客数"])
            writer.writeheader()
            writer.writerows(rows)
        print(f"生成: {filepath} ({len(rows)}行)")

print("サンプルデータ生成完了")
