#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证件管理程序 - 宽表格版本
用于处理"一行一人，多证书列"格式的Excel文件
"""

import sys
import io
from datetime import date, datetime
from pathlib import Path

# 设置stdout编码
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd


# 证书列配置：列名索引 -> 证书类型名称
CERTIFICATE_COLUMNS = {
    11: ('IADC/IWCF', 9, 10, 11),  # (证书类型, 证编号列, 发证日期列, 到期时间列)
    14: ('HSE证(H2S)', 12, 13, 14),
    17: ('井控证', 15, 16, 17),
    20: ('司机证', 18, 19, 20),
    23: ('局部指吊证', 21, 22, 23),
    32: ('起重证', 30, 31, 32),
    35: ('单位证', 33, 34, 35),
}

NAME_COLUMN = 2  # 姓名列索引


def parse_date(value):
    """解析日期"""
    if pd.isna(value):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def calculate_days_remaining(expiry_date):
    """计算剩余天数"""
    today = date.today()
    delta = expiry_date - today
    return delta.days


def classify_urgency(days_remaining):
    """分类紧急程度"""
    if days_remaining < 30:
        return '🔴 紧急', '【30天内】', 1
    elif days_remaining < 90:
        return '🟡 预警', '【90天内】', 2
    else:
        return '🟢 正常', '【正常】', 3


def process_wide_format(file_path):
    """
    处理宽格式的Excel文件

    Args:
        file_path: Excel文件路径

    Returns:
        DataFrame: 处理后的结果
    """
    # 读取Excel，不使用header
    df = pd.read_excel(file_path, header=None)

    results = []

    # 从第3行开始（索引3）是数据行
    for idx in range(3, len(df)):
        row = df.iloc[idx]
        name = row[NAME_COLUMN]

        if pd.isna(name):
            continue

        # 检查每个证书类型
        for expiry_col, cert_info in CERTIFICATE_COLUMNS.items():
            cert_type, num_col, issue_col, exp_col = cert_info

            # 获取到期日期
            expiry_value = row[exp_col]
            expiry_date = parse_date(expiry_value)

            if expiry_date is None:
                continue

            # 获取证书编号
            cert_num = row[num_col] if num_col < len(row) else ''
            if pd.isna(cert_num):
                cert_num = ''

            # 计算剩余天数和分类
            days_remaining = calculate_days_remaining(expiry_date)
            priority, status, _ = classify_urgency(days_remaining)

            results.append({
                '优先级': priority,
                '状态标识': status,
                '姓名': str(name).strip(),
                '证件类型': cert_type,
                '证书编号': str(cert_num).strip(),
                '有效期': expiry_date.strftime('%Y-%m-%d'),
                '剩余天数': days_remaining,
            })

    if not results:
        print("⚠️ 警告：没有解析到任何有效的证书数据")
        sys.exit(1)

    result_df = pd.DataFrame(results)

    # 按紧急程度排序
    result_df['排序值'] = result_df['剩余天数'].apply(
        lambda x: 1 if x < 30 else (2 if x < 90 else 3)
    )
    result_df = result_df.sort_values(['排序值', '剩余天数'])
    result_df = result_df.drop('排序值', axis=1).reset_index(drop=True)

    return result_df


def generate_report(df, output_path):
    """生成CSV报告"""
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"✅ 报告已生成: {output_path}")

        # 统计信息
        urgent_count = len(df[df['优先级'] == '🔴 紧急'])
        warning_count = len(df[df['优先级'] == '🟡 预警'])
        normal_count = len(df[df['优先级'] == '🟢 正常'])

        print("\n📈 统计信息：")
        print(f"   🔴 紧急（<30天）: {urgent_count} 个证书")
        print(f"   🟡 预警（30-90天）: {warning_count} 个证书")
        print(f"   🟢 正常（≥90天）: {normal_count} 个证书")

    except Exception as e:
        print(f"❌ 生成报告失败: {e}")
        sys.exit(1)


def main():
    """主函数"""
    if len(sys.argv) < 3:
        print("用法: python certificate_checker_wide.py <输入文件.xlsx> <输出文件.csv>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    # 检查输入文件
    if not Path(input_file).exists():
        print(f"❌ 错误：输入文件不存在: {input_file}")
        sys.exit(1)

    print("=" * 50)
    print("🔍 证件管理程序（宽表格版）")
    print("=" * 50)

    # 处理数据
    print("\n⚙️ 处理数据...")
    result_df = process_wide_format(input_file)

    # 生成报告
    print("\n📝 生成报告...")
    generate_report(result_df, output_file)

    print("\n" + "=" * 50)
    print("✅ 完成！")
    print("=" * 50)


if __name__ == '__main__':
    main()
