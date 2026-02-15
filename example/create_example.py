#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例Excel文件
"""

import pandas as pd
from datetime import date, timedelta

# 创建示例数据
# 使用相对今天的日期来生成示例
today = date.today()

data = {
    '姓名': [
        '张三',      # 25天后过期 - 紧急
        '李四',      # 30天后过期 - 紧急边界
        '王五',      # 45天后过期 - 预警
        '赵六',      # 75天后过期 - 预警
        '钱七',      # 95天后过期 - 正常
        '孙八',      # 120天后过期 - 正常
        '周九',      # 5天后过期 - 紧急
        '吴十',      # 已过期 - 紧急
    ],
    '证件类型': [
        '驾驶证',
        '身份证',
        '工作证',
        '安全员证',
        '特种设备操作证',
        '健康证',
        '焊工证',
        '建造师证',
    ],
    '有效期': [
        today + timedelta(days=25),
        today + timedelta(days=30),
        today + timedelta(days=45),
        today + timedelta(days=75),
        today + timedelta(days=95),
        today + timedelta(days=120),
        today + timedelta(days=5),
        today - timedelta(days=10),  # 已过期
    ],
    '联系方式': [
        '138****1234',
        '139****5678',
        '136****9012',
        '137****3456',
        '135****7890',
        '158****1122',
        '159****3344',
        '186****5566',
    ]
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存为Excel
output_file = 'input_example.xlsx'
df.to_excel(output_file, index=False)

print(f"Example file created: {output_file}")
print(f"Reference date: {today}")
print("\nData preview:")
print(df)
