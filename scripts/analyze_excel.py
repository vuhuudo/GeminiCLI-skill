import pandas as pd
import sys
import os
from datetime import datetime


def analyze_excel(input_path, output_report_path):
    if not os.path.exists(input_path):
        print(f"Error: File '{input_path}' not found.")
        sys.exit(1)

    try:
        xls = pd.ExcelFile(input_path)
    except Exception as e:
        print(f"Error: Could not open '{input_path}': {e}")
        sys.exit(1)

    summary = [
        "=" * 40,
        f"BÁO CÁO PHÂN TÍCH EXCEL",
        f"File    : {os.path.basename(input_path)}",
        f"Thời điểm: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        f"Số sheet : {len(xls.sheet_names)}",
        "=" * 40,
    ]

    for sheet_name in xls.sheet_names:
        summary.append(f"\n### Sheet: {sheet_name} ###")
        try:
            df = xls.parse(sheet_name)
        except Exception as e:
            summary.append(f"  Lỗi khi đọc sheet: {e}")
            continue

        # Ensure column names are strings for safe formatting
        df.columns = [str(col) for col in df.columns]

        summary.append(f"Số cột  : {len(df.columns)}")
        summary.append(f"Số hàng : {len(df)}")
        summary.append(f"Cột     : {', '.join(df.columns)}")
        summary.append("-" * 40)

        numeric_df = df.select_dtypes(include=['number'])
        if not numeric_df.empty:
            summary.append("Phân tích số liệu:")
            summary.append(numeric_df.describe().to_string())

        # List columns with missing values
        missing = df.isnull().sum()
        missing = missing[missing > 0]
        if not missing.empty:
            summary.append("\nCột có giá trị thiếu:")
            for col, count in missing.items():
                summary.append(f"  - {col}: {count} giá trị thiếu")

        summary.append("-" * 40)

    try:
        with open(output_report_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(summary))
    except OSError as e:
        print(f"Error: Could not write report to '{output_report_path}': {e}")
        sys.exit(1)

    print(f"Report generated at: {output_report_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python analyze_excel.py <input_path> <output_report_path>")
        sys.exit(1)
    analyze_excel(sys.argv[1], sys.argv[2])
