import pandas as pd
import os
import sys

def analyze_gl(file_path):
    print(f"Analyzing {file_path}...")

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    # Clean Debit and Credit columns
    # Convert to numeric, coercing errors to NaN
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce').fillna(0)

    # Ensure Amount is numeric
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)

    # --- Statistics ---

    # Row count
    row_count = len(df)

    # Date Ranges
    # Ensure datetime format
    df['EffectiveDate'] = pd.to_datetime(df['EffectiveDate'], errors='coerce')
    df['EntryDate'] = pd.to_datetime(df['EntryDate'], errors='coerce')

    eff_date_min = df['EffectiveDate'].min()
    eff_date_max = df['EffectiveDate'].max()
    entry_date_min = df['EntryDate'].min()
    entry_date_max = df['EntryDate'].max()

    # Financial Totals
    total_debit = df['Debit'].sum()
    total_credit = df['Credit'].sum()

    # Amount Stats
    amount_sum = df['Amount'].sum()
    amount_mean = df['Amount'].mean()
    amount_min = df['Amount'].min()
    amount_max = df['Amount'].max()

    # Group By Stats
    account_type_counts = df['AccountType'].value_counts()
    source_counts = df['Source'].value_counts()

    # --- Output Report ---

    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    report_path = os.path.join(output_dir, 'report.md')

    with open(report_path, 'w') as f:
        f.write("# GL Data Analysis Report\n\n")
        f.write(f"**File Analyzed:** {file_path}\n\n")

        f.write("## Basic Statistics\n")
        f.write(f"- **Total Rows:** {row_count:,}\n")
        f.write(f"- **Total Debits:** {total_debit:,.2f}\n")
        f.write(f"- **Total Credits:** {total_credit:,.2f}\n")
        f.write(f"- **Net Amount (Sum):** {amount_sum:,.2f}\n")
        f.write(f"- **Average Amount:** {amount_mean:,.2f}\n")
        f.write(f"- **Min Amount:** {amount_min:,.2f}\n")
        f.write(f"- **Max Amount:** {amount_max:,.2f}\n\n")

        f.write("## Date Ranges\n")
        f.write("### Effective Date\n")
        f.write(f"- **Min:** {eff_date_min}\n")
        f.write(f"- **Max:** {eff_date_max}\n")
        f.write("### Entry Date\n")
        f.write(f"- **Min:** {entry_date_min}\n")
        f.write(f"- **Max:** {entry_date_max}\n\n")

        f.write("## Account Type Counts\n")
        f.write("| Account Type | Count |\n")
        f.write("|---|---:|\n")
        for type_name, count in account_type_counts.items():
            f.write(f"| {type_name} | {count:,} |\n")
        f.write("\n")

        f.write("## Source Counts\n")
        f.write("| Source | Count |\n")
        f.write("|---|---:|\n")
        for source_name, count in source_counts.items():
            f.write(f"| {source_name} | {count:,} |\n")
        f.write("\n")

    print(f"Analysis complete. Report saved to {report_path}")

if __name__ == "__main__":
    analyze_gl('je_samples (1).xlsx')
