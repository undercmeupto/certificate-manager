#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证件管理程序 - Certificate Manager
自动检测证件有效期并生成预警报告
"""

import sys
import io

# 设置stdout编码为UTF-8，确保在Windows控制台也能正确显示中文和emoji
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd
import argparse
import sys
from datetime import date, datetime
from pathlib import Path


# 支持的列名映射
COLUMN_MAPPING = {
    'name': ['姓名', '名字', 'Name', 'name', '员工姓名', '人员姓名'],
    'certificate_type': ['证件类型', '证书类型', 'Certificate Type', 'certificate_type', '证件名称', '证书名称'],
    'expiry_date': ['有效期', '到期日', '有效期至', 'expiry_date', 'expire_date', '有效日期', '到期日期', '有效期到', '证件有效期', '证书有效期'],
    'contact': ['联系方式', '电话', '手机', 'Phone', 'Tel', '联系电话', '手机号', 'Contact']
}


def detect_columns(df):
    """
    自动检测Excel文件中的列名

    Args:
        df: DataFrame对象

    Returns:
        dict: 检测到的列名映射 {'name': '姓名', 'certificate_type': '证件类型', ...}
    """
    detected = {}
    columns = df.columns.tolist()

    for key, possible_names in COLUMN_MAPPING.items():
        for col in columns:
            if col in possible_names:
                detected[key] = col
                break

    # 检查必需的列
    required = ['name', 'certificate_type', 'expiry_date']
    missing = [k for k in required if k not in detected]

    if missing:
        print("❌ 错误：无法识别以下必需列：")
        for m in missing:
            print(f"   - {m} (支持的列名: {', '.join(COLUMN_MAPPING[m])})")
        print(f"\n📋 当前文件中的列: {', '.join(columns)}")
        sys.exit(1)

    # 联系方式是可选的
    if 'contact' not in detected:
        detected['contact'] = None

    return detected


def parse_expiry_date(value):
    """
    解析各种格式的日期

    Args:
        value: 日期值（可能是字符串、datetime等）

    Returns:
        date对象或None
    """
    if pd.isna(value):
        return None

    # 如果已经是datetime对象
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()

    # 尝试解析字符串
    if isinstance(value, str):
        # 尝试常见格式
        formats = [
            '%Y-%m-%d',
            '%Y/%m/%d',
            '%Y.%m.%d',
            '%d-%m-%Y',
            '%d/%m/%Y',
            '%d.%m.%Y',
            '%m-%d-%Y',
            '%m/%d/%Y',
            '%Y年%m月%d日',
        ]
        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue

        # 尝试pandas的to_datetime
        try:
            return pd.to_datetime(value).date()
        except:
            pass

    return None


def read_excel_file(file_path):
    """
    读取Excel文件

    Args:
        file_path: Excel文件路径

    Returns:
        tuple: (DataFrame, 检测到的列名映射)
    """
    try:
        df = pd.read_excel(file_path)
        print(f"✅ 成功读取文件: {file_path}")
        print(f"📊 共 {len(df)} 条记录")
        return df
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        sys.exit(1)


def calculate_days_remaining(expiry_date):
    """
    计算剩余天数

    Args:
        expiry_date: 有效期日期对象

    Returns:
        int: 剩余天数（负数表示已过期）
    """
    today = date.today()
    delta = expiry_date - today
    return delta.days


def classify_urgency(days_remaining):
    """
    分类紧急程度

    Args:
        days_remaining: 剩余天数

    Returns:
        tuple: (优先级标签, 状态标识, 排序值)
    """
    if days_remaining < 30:
        return '🔴 紧急', '【30天内】', 1
    elif days_remaining < 90:
        return '🟡 预警', '【90天内】', 2
    else:
        return '🟢 正常', '【正常】', 3


def process_certificates(df, columns):
    """
    处理证件数据，计算剩余天数并分类

    Args:
        df: 原始DataFrame
        columns: 检测到的列名映射

    Returns:
        DataFrame: 处理后的数据
    """
    results = []

    for _, row in df.iterrows():
        expiry_value = row[columns['expiry_date']]
        expiry_date = parse_expiry_date(expiry_value)

        if expiry_date is None:
            # 无法解析日期，跳过或标记
            continue

        days_remaining = calculate_days_remaining(expiry_date)
        priority, status, _ = classify_urgency(days_remaining)

        result = {
            '优先级': priority,
            '状态标识': status,
            '姓名': row[columns['name']],
            '证件类型': row[columns['certificate_type']],
            '有效期': expiry_date.strftime('%Y-%m-%d'),
            '剩余天数': days_remaining,
        }

        # 添加联系方式（如果存在）
        if columns.get('contact'):
            result['联系方式'] = row[columns['contact']]
        else:
            result['联系方式'] = ''

        results.append(result)

    if not results:
        print("⚠️ 警告：没有解析到任何有效的日期数据")
        sys.exit(1)

    result_df = pd.DataFrame(results)

    # 按紧急程度排序（紧急程度高的在前）
    result_df['排序值'] = result_df['剩余天数'].apply(
        lambda x: 1 if x < 30 else (2 if x < 90 else 3)
    )
    result_df = result_df.sort_values(['排序值', '剩余天数'])
    result_df = result_df.drop('排序值', axis=1)

    return result_df


def generate_report(df, output_path):
    """
    生成CSV报告

    Args:
        df: 处理后的DataFrame
        output_path: 输出文件路径
    """
    try:
        # 使用UTF-8-sig编码以确保Excel能正确打开中文
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"✅ 报告已生成: {output_path}")

        # 统计信息
        urgent_count = len(df[df['优先级'] == '🔴 紧急'])
        warning_count = len(df[df['优先级'] == '🟡 预警'])
        normal_count = len(df[df['优先级'] == '🟢 正常'])

        print("\n📈 统计信息：")
        print(f"   🔴 紧急（<30天）: {urgent_count} 人")
        print(f"   🟡 预警（30-90天）: {warning_count} 人")
        print(f"   🟢 正常（≥90天）: {normal_count} 人")

    except Exception as e:
        print(f"❌ 生成报告失败: {e}")
        sys.exit(1)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='证件管理程序 - 自动检测证件有效期并生成预警报告',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例用法:
  python certificate_checker.py input.xlsx output.csv
  python certificate_checker.py data.xlsx report.csv

支持的列名:
  姓名: 姓名, 名字, Name, name, 员工姓名, 人员姓名
  证件类型: 证件类型, 证书类型, Certificate Type, 证件名称
  有效期: 有效期, 到期日, 有效期至, expiry_date, 有效日期, 证件有效期
  联系方式: 联系方式, 电话, 手机, Phone (可选)
        '''
    )

    parser.add_argument('input_file', help='输入Excel文件路径')
    parser.add_argument('output_file', help='输出CSV文件路径')

    args = parser.parse_args()

    # 检查输入文件是否存在
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"❌ 错误：输入文件不存在: {args.input_file}")
        sys.exit(1)

    print("=" * 50)
    print("🔍 证件管理程序")
    print("=" * 50)

    # 读取Excel文件
    df = read_excel_file(args.input_file)

    # 检测列名
    print("\n🔎 检测列名...")
    columns = detect_columns(df)
    print(f"   姓名: {columns['name']}")
    print(f"   证件类型: {columns['certificate_type']}")
    print(f"   有效期: {columns['expiry_date']}")
    if columns.get('contact'):
        print(f"   联系方式: {columns['contact']}")

    # 处理数据
    print("\n⚙️ 处理数据...")
    result_df = process_certificates(df, columns)

    # 生成报告
    print("\n📝 生成报告...")
    generate_report(result_df, args.output_file)

    print("\n" + "=" * 50)
    print("✅ 完成！")
    print("=" * 50)


if __name__ == '__main__':
    main()
