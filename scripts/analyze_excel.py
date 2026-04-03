import pandas as pd
import sys
import os

def analyze_excel(input_path, output_report_path):
    if not os.path.exists(input_path):
        print(f"Error: File {input_path} not found.")
        return

    # Read all sheets
    xls = pd.ExcelFile(input_path)
    summary = []

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        summary.append(f"Sheet: {sheet_name}")
        summary.append(f"Columns: {', '.join(df.columns)}")
        summary.append(f"Rows: {len(df)}")
        summary.append("-" * 20)

        # Basic analysis (example: mean/sum of numeric columns)
        numeric_df = df.select_dtypes(include=['number'])
        if not numeric_df.empty:
            summary.append("Numerical Analysis:")
            summary.append(numeric_df.describe().to_string())
            summary.append("-" * 20)

    # Write summary to file (or docx if requested, but for now text)
    with open(output_report_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(summary))
    
    print(f"Report generated at: {output_report_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python analyze_excel.py <input_path> <output_report_path>")
    else:
        analyze_excel(sys.argv[1], sys.argv[2])
